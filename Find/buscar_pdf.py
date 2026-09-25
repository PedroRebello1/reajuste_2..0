#!/usr/bin/env python3
"""
Busca um nome em todos os PDFs das pastas de 2022 e 2023 e salva
para cada página em que o nome aparece, o reajuste na pasta output/.

Desempenho: os PDFs são divididos em blocos de páginas processados em paralelo,
um processo por núcleo da CPU.

Requisito: pip install pymupdf
"""

from __future__ import annotations

import math
import os
import re
import sys
import time
import unicodedata
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:
    import fitz  # versões antigas do PyMuPDF

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
PADRAO_PASTA_ANO = re.compile(r"^\d{4}$")  # detecta 2022, 2023, 2024...

# Paralelismo: número de processos (padrão = núcleos da CPU) e tamanho mínimo de bloco.
MAX_PROCESSOS = os.cpu_count() or 1
MIN_PAGINAS_POR_BLOCO = 200

# Junta palavras hifenizadas no fim da linha e ignora texto fora da página.
FLAGS_TEXTO = getattr(fitz, "TEXT_DEHYPHENATE", 0) | getattr(fitz, "TEXT_MEDIABOX_CLIP", 0)
CARACTERES_INVALIDOS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_ACENTOS = re.compile(r"[\u0300-\u036f]")
_ESPACOS = re.compile(r"\s+")


@dataclass
class Ocorrencia:
    ano: str
    arquivo: Path
    pagina: int         # índice 0-based da página onde o nome foi encontrado
    quantidade: int     # quantas vezes o nome aparece nessa página
    total_paginas: int
    destino: Path | None = None  # arquivo de saída (None se não há página seguinte)


@dataclass
class ResultadoBloco:
    arquivo: Path
    paginas_lidas: int
    ocorrencias: list[Ocorrencia] = field(default_factory=list)
    erros: list[str] = field(default_factory=list)


# Utilidades

def normalizar(texto: str) -> str:
    """Remove acentos, expande ligaduras (ﬁ -> fi), ignora maiúsculas e unifica espaços."""
    texto = _ACENTOS.sub("", unicodedata.normalize("NFKD", texto))
    return _ESPACOS.sub(" ", texto.casefold())


def formatar_duracao(segundos: float) -> str:
    if segundos < 60:
        return f"{segundos:.1f} s"
    minutos, seg = divmod(round(segundos), 60)
    horas, minutos = divmod(minutos, 60)
    return f"{horas}h {minutos:02d}min {seg:02d}s" if horas else f"{minutos}min {seg:02d}s"


def listar_pastas_ano() -> list[Path]:
    return sorted(p for p in BASE_DIR.iterdir() if p.is_dir() and PADRAO_PASTA_ANO.match(p.name))


def listar_pdfs(pasta: Path) -> list[Path]:
    pdfs = (p for p in pasta.iterdir() if p.is_file() and p.suffix.lower() == ".pdf")
    return sorted(pdfs, key=lambda p: p.name.lower())


def abrir_pdf(caminho: Path) -> fitz.Document | None:
    try:
        doc = fitz.open(caminho)
    except Exception as e:
        print(f"    [ERRO] Não foi possível abrir '{caminho.name}': {e}")
        return None
    if doc.needs_pass and not doc.authenticate(""):
        print(f"    [AVISO] '{caminho.name}' é protegido por senha. Ignorado.")
        doc.close()
        return None
    return doc


# Busca (executada em paralelo nos processos filhos)

def ler_bloco(caminho: Path, ano: str, inicio: int, fim: int, padrao_str: str) -> ResultadoBloco:
    """Lê as páginas [inicio, fim) de um PDF e conta as ocorrências do nome."""
    padrao = re.compile(padrao_str)
    res = ResultadoBloco(caminho, fim - inicio)
    try:
        doc = fitz.open(caminho)
        if doc.needs_pass:
            doc.authenticate("")
    except Exception as e:
        res.erros.append(f"'{caminho.name}': {e}")
        return res

    with doc:
        total = len(doc)
        for i in range(inicio, fim):
            try:
                texto = doc[i].get_text("text", flags=FLAGS_TEXTO)
            except Exception as e:
                res.erros.append(f"'{caminho.name}', página {i + 1}: {e}")
                continue
            if not texto.strip():
                continue  # página só de imagem (aviso de reajuste): nada a buscar
            qtd = len(padrao.findall(normalizar(texto)))
            if qtd:
                res.ocorrencias.append(Ocorrencia(ano, caminho, i, qtd, total))
    return res


def mapear_pdfs(pastas: list[Path]) -> list[tuple[Path, str, int]]:
    """Lista (arquivo, ano, total de páginas) de todos os PDFs legíveis."""
    arquivos = []
    for pasta in pastas:
        pdfs = listar_pdfs(pasta)
        print(f"[{pasta.name}] {len(pdfs)} PDF(s)")
        for pdf in pdfs:
            doc = abrir_pdf(pdf)
            if doc is None:
                continue
            with doc:
                total = len(doc)
            print(f"  - {pdf.name} ({total:,} páginas)".replace(",", "."))
            if total:
                arquivos.append((pdf, pasta.name, total))
    return arquivos


def buscar_em_paralelo(arquivos: list[tuple[Path, str, int]], padrao_str: str) -> list[Ocorrencia]:
    total_geral = sum(t for _, _, t in arquivos)
    # Blocos pequenos o bastante para equilibrar a carga entre os núcleos.
    tamanho = max(MIN_PAGINAS_POR_BLOCO, math.ceil(total_geral / (MAX_PROCESSOS * 4)))
    blocos = [
        (pdf, ano, ini, min(ini + tamanho, total))
        for pdf, ano, total in arquivos
        for ini in range(0, total, tamanho)
    ]
    processos = max(1, min(MAX_PROCESSOS, len(blocos)))
    print(f"\nProcessando {total_geral:,} páginas em {processos} processo(s)...".replace(",", "."))

    ocorrencias: list[Ocorrencia] = []
    erros: list[str] = []
    lidas = 0

    executor = ProcessPoolExecutor(max_workers=processos)
    try:
        futuros = [executor.submit(ler_bloco, pdf, ano, ini, fim, padrao_str)
                   for pdf, ano, ini, fim in blocos]
        for futuro in as_completed(futuros):
            try:
                r = futuro.result()
            except Exception as e:
                erros.append(f"falha em um processo: {e}")
                continue
            lidas += r.paginas_lidas
            ocorrencias.extend(r.ocorrencias)
            erros.extend(r.erros)
            print(f"  progresso: {lidas / total_geral:6.1%} ({lidas}/{total_geral} páginas)",
                  end="\r", flush=True)
    except KeyboardInterrupt:
        executor.shutdown(wait=False, cancel_futures=True)
        raise
    executor.shutdown()
    print()

    for erro in erros:
        print(f"  [ERRO] {erro}")

    # Os blocos terminam fora de ordem; reordena para numeração determinística.
    ocorrencias.sort(key=lambda o: (o.ano, o.arquivo.name.lower(), o.pagina))
    return ocorrencias


# Nomes de saída e gravação

def definir_destinos(ocorrencias: list[Ocorrencia], nome_arquivo: str) -> None:
    """Define o arquivo de saída de cada ocorrência, numerando quando há mais de uma por ano."""
    por_ano: dict[str, list[Ocorrencia]] = {}
    for oc in ocorrencias:
        if oc.pagina + 1 < oc.total_paginas:  # só se existir página seguinte
            por_ano.setdefault(oc.ano, []).append(oc)

    for ano, lista in por_ano.items():
        for n, oc in enumerate(lista, start=1):
            sufixo = f"_{n:02d}" if len(lista) > 1 else ""
            oc.destino = OUTPUT_DIR / f"Reajuste_{nome_arquivo}_{ano}{sufixo}.pdf"


def salvar_paginas(ocorrencias: list[Ocorrencia]) -> int:
    OUTPUT_DIR.mkdir(exist_ok=True)
    salvos = 0
    abertos: dict[Path, fitz.Document | None] = {}

    try:
        for oc in ocorrencias:
            if oc.destino is None:
                continue
            if oc.arquivo not in abertos:
                abertos[oc.arquivo] = abrir_pdf(oc.arquivo)
            origem = abertos[oc.arquivo]
            if origem is None:
                continue

            proxima = oc.pagina + 1
            try:
                with fitz.open() as novo:
                    novo.insert_pdf(origem, from_page=proxima, to_page=proxima)
                    novo.save(oc.destino, garbage=3, deflate=True)
                salvos += 1
            except Exception as e:
                print(f"  [ERRO] Falha ao salvar '{oc.destino.name}': {e}")
    finally:
        for doc in abertos.values():
            if doc is not None:
                doc.close()
    return salvos


# Programa principal

def processar(nome: str, pastas: list[Path]) -> int:
    termo = normalizar(nome).strip()
    # (?<!\w) e (?!\w): casa só o nome inteiro ("Ana" não casa com "Mariana").
    padrao_str = r"(?<!\w)" + re.escape(termo) + r"(?!\w)"
    nome_arquivo = CARACTERES_INVALIDOS.sub("", nome.upper())

    print(f"\nBuscando '{nome}' nas pastas: {', '.join(p.name for p in pastas)}\n")
    arquivos = mapear_pdfs(pastas)
    if not arquivos:
        print("\nNenhum PDF legível encontrado.")
        return 1

    ocorrencias = buscar_em_paralelo(arquivos, padrao_str)
    if not ocorrencias:
        print(f"\nO nome '{nome}' não foi encontrado em nenhum PDF.")
        return 0

    definir_destinos(ocorrencias, nome_arquivo)
    for oc in ocorrencias:
        if oc.destino is None:
            print(f"  [AVISO] '{oc.arquivo.name}', página {oc.pagina + 1}: é a última página, "
                  "não há página seguinte para salvar.")
    salvar_paginas(ocorrencias)

    qtd_arquivos = len({oc.arquivo for oc in ocorrencias})
    print(f"\nTotal: {len(ocorrencias)} página(s) com o nome em {qtd_arquivos} arquivo(s).")
    return 0


def main() -> int:
    pastas = listar_pastas_ano()
    if not pastas:
        print(f"Nenhuma pasta de ano (ex.: 2022, 2023) encontrada em: {BASE_DIR}")
        return 1

    nome = _ESPACOS.sub(" ", input("Digite o nome a buscar: ")).strip()
    if not nome:
        print("Erro: o nome não pode ser vazio.")
        return 1

    t0 = time.perf_counter()
    try:
        return processar(nome, pastas)
    finally:
        print(f"Duração: {formatar_duracao(time.perf_counter() - t0)}")


if __name__ == "__main__":  # obrigatório para o multiprocessamento no Windows
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nInterrompido pelo usuário.")
        sys.exit(130)