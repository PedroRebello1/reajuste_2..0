"""Corrige defeitos de layout herdados da conversão HTML -> DOCX dos modelos.

Os modelos vieram de uma conversão de HTML e trazem dois problemas que aparecem
no aviso gerado:

1. Faixa preta acima do banner: a tabela do cabeçalho tem uma borda superior
   `w:color="auto"`, que o Word resolve como preto, com 4,5pt de espessura.
2. "Aba" azul sobrando à direita: a tabela interna do banner (logo | texto) é
   declarada mais larga que a célula que a contém, então transborda.

Uso:
    python corrigir_modelo.py modelos/2026/participativo.docx [outro.docx ...]

Sem argumentos, corrige todos os .docx de modelos/.
"""

import shutil
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

# Bordas "auto" a partir desta espessura (em oitavos de ponto) são tratadas
# como faixa preta indesejada. 8 = 1pt.
SZ_MIN_BORDA_INDESEJADA = 8


def largura_grid(tbl):
    grid = tbl.find(qn('w:tblGrid'))
    if grid is None:
        return None
    return sum(int(gc.get(qn('w:w'))) for gc in grid)


def remover_borda_preta(body):
    """Zera bordas de tabela grossas com cor automática (renderizadas em preto)."""
    corrigidas = []
    for tbl in body.iter(qn('w:tbl')):
        pr = tbl.find(qn('w:tblPr'))
        if pr is None:
            continue
        bordas = pr.find(qn('w:tblBorders'))
        if bordas is None:
            continue
        for lado in bordas:
            val = lado.get(qn('w:val'))
            cor = (lado.get(qn('w:color')) or "").upper()
            sz = int(lado.get(qn('w:sz')) or 0)
            if val in (None, "nil", "none") or cor not in ("AUTO", "000000"):
                continue
            if sz < SZ_MIN_BORDA_INDESEJADA:
                continue
            nome = lado.tag.split('}')[1]
            corrigidas.append(f"{nome} {sz / 8:.2f}pt cor={cor}")
            # Mantém o elemento, mas neutraliza: val="nil" também sobrepõe
            # qualquer borda herdada do estilo da tabela.
            lado.set(qn('w:val'), "nil")
            for attr in ('w:sz', 'w:color', 'w:space'):
                if lado.get(qn(attr)) is not None:
                    del lado.attrib[qn(attr)]
    return corrigidas


def largura_texto_pagina(body):
    """Largura útil da página (twips), descontando as margens."""
    sect = body.find(qn('w:sectPr'))
    if sect is None:
        return None
    pg = sect.find(qn('w:pgSz'))
    mar = sect.find(qn('w:pgMar'))
    if pg is None or mar is None:
        return None
    return (int(pg.get(qn('w:w')))
            - int(mar.get(qn('w:left')))
            - int(mar.get(qn('w:right'))))


def largura_disponivel(tbl, body):
    """Espaço em que a tabela pode ser desenhada: a célula que a contém,
    ou a largura útil da página quando ela está no corpo do documento."""
    pai = tbl.getparent()
    if pai.tag == qn('w:tc'):
        return largura_grid(pai.getparent().getparent())
    return largura_texto_pagina(body)


def ajustar_tabela_larga(body):
    """Reduz tabelas declaradas mais largas do que o espaço que as contém.

    A conversão de HTML deixou `w:tblW` maior que o próprio `w:tblGrid` em
    vários níveis aninhados; o Word usa `tblW` para desenhar a tabela, então
    ela transborda para fora do bloco pai.
    """
    ajustadas = []
    for tbl in body.iter(qn('w:tbl')):
        largura = largura_grid(tbl)
        if not largura:
            continue

        disponivel = largura_disponivel(tbl, body)
        if not disponivel:
            continue

        # 1) Grid maior que o espaço disponível: reduz proporcionalmente.
        if largura > disponivel:
            fator = disponivel / largura
            for gc in tbl.find(qn('w:tblGrid')):
                gc.set(qn('w:w'), str(round(int(gc.get(qn('w:w'))) * fator)))
            for linha in tbl.findall(qn('w:tr')):
                for tc in linha.findall(qn('w:tc')):
                    tcpr = tc.find(qn('w:tcPr'))
                    tcw = tcpr.find(qn('w:tcW')) if tcpr is not None else None
                    if tcw is not None and tcw.get(qn('w:type')) == "dxa":
                        valor = int(tcw.get(qn('w:w')) or 0)
                        if valor:
                            tcw.set(qn('w:w'), str(round(valor * fator)))
            ajustadas.append(
                f"grid {largura} -> {disponivel} twips "
                f"(excedia {(largura - disponivel) / 20:.1f}pt)"
            )
            largura = disponivel

        # 2) tblW acima do grid ou do espaço disponível: limita ao menor.
        pr = tbl.find(qn('w:tblPr'))
        w = pr.find(qn('w:tblW')) if pr is not None else None
        if w is None or w.get(qn('w:type')) != "dxa":
            continue
        declarada = int(w.get(qn('w:w')) or 0)
        limite = min(largura, disponivel)
        if declarada > limite:
            w.set(qn('w:w'), str(limite))
            ajustadas.append(
                f"tblW {declarada} -> {limite} twips "
                f"(excedia {(declarada - limite) / 20:.1f}pt)"
            )
    return ajustadas


def corrigir(caminho: Path):
    doc = Document(str(caminho))
    body = doc.element.body

    bordas = remover_borda_preta(body)
    larguras = ajustar_tabela_larga(body)

    if not bordas and not larguras:
        print(f"{caminho.name}: nada a corrigir")
        return False

    backup = caminho.with_suffix(".docx.bak")
    if not backup.exists():
        shutil.copy(caminho, backup)
    doc.save(str(caminho))

    print(f"{caminho.name}:")
    for b in bordas:
        print(f"  borda preta removida: {b}")
    for a in larguras:
        print(f"  tabela reduzida: {a}")
    print(f"  backup: {backup.name}")
    return True


def main(argv):
    if argv:
        alvos = [Path(a) for a in argv]
    else:
        alvos = sorted(Path("modelos").rglob("*.docx"))

    for alvo in alvos:
        if alvo.name.startswith("~$"):
            continue
        if not alvo.exists():
            print(f"{alvo}: não encontrado")
            continue
        corrigir(alvo)


if __name__ == "__main__":
    main(sys.argv[1:])
