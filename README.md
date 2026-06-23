# Automatizacao de Reajustes

Automacao em Python para capturar dados na tela (via `pyautogui`), identificar beneficiario/protocolo e gerar documentos de reajuste com base nos modelos e tabelas CSV do projeto.

---
**Essa versão foi otimizada e pré-configurada para rodar utilizando apenas uma tela (Monitor esquerdo), no computador** `SEDEWKS825900`**, com os aplicativos abertos em tela dividida da seguinte forma:**

`Edge em "cassi.lightning.force.com"` (50%) | `SOC Produção` (50%)

---

## Requisitos

- Windows 11.
- Visual Studio Code ou outra IDE de sua preferência
- Python 3.10+ instalado e no `PATH`.
- Dependencias do `requirements.txt`.
- Arquivos de dados presentes na raiz:
	- `valores.csv`
	- `ans_vida.csv`
- Modelos `.docx` em `modelos/2024`, `modelos/2025` e `modelos/2026`
- Opcional para gerar PDF automaticamente: Word instalado e com conta logada

## Instalacao

1. Baixe este repositorio (zip) e extraia em uma pasta local.
2. Abra a pasta no VS Code.
3. No terminal da pasta do projeto, instale as dependencias:

```bash
pip install -r requirements.txt
```

## Ajuste de coordenadas (importante)

As automacoes dependem das coordenadas definidas nos arquivos Python (`COORD`).
Se o clique estiver em local errado, ajuste os valores para o seu monitor/resolucao.

Arquivos com coordenadas:

- `searcher.py`
- `autofill_cassi.py`

## Como executar

Script principal atual:

```bash
python searcher.py
```

Fluxo de execucao:

1. O script abre uma janela para digitar protocolos.
2. Voce pode separar protocolos por `,`, `;`, `.`, ou espaco. **(recomenda-se usar apenas 1 protocolo por vez)**
3. Para cada protocolo, o script busca no sistema, pede confirmacoes por janelas e executa o preenchimento.
4. Os arquivos gerados sao salvos em `prontos/`.

## Estrutura resumida

- `searcher.py`: orquestra o fluxo de busca de protocolos e chamadas do preenchimento.
- `autofill_cassi.py`: captura campos na tela e gera os documentos com base nos modelos/CSV.
- `modelos/`: templates `.docx` por ano e tipo de plano.
- `prontos/`: saida dos arquivos gerados.

## Solucao de problemas

- Erro de importacao de pacote:
	execute novamente `pip install -r requirements.txt`.
- Nenhum documento gerado:
	confira se os CSVs e modelos existem e se o plano foi mapeado.
- PDF nao gerado:
	verifique se o Word esta instalado e a conta conectada (com pacote 365).
- Captura de texto incorreta:
	revise coordenadas e foco da janela do sistema antes de rodar.

## Observacoes

- Durante a execucao, nao mova janelas nem altere resolucao/zoom do sistema.
- Os scripts foram pensados para execucao assistida (com confirmacoes manuais indicadas pelos pop-ups).
