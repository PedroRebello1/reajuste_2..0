# TODO: Make this shit work again
# TODO: Try selecting fields with mouse movement
# TODO: Integração ao SOC via request - Databricks Gold



# Versão manual
## Para acessar a versão executável acesse: github.com/PedroRebello1/reajuste_2.0 - Ainda indisponível
# Automatizacao de Reajustes

Automacao em Python para capturar dados na tela (via `pyautogui`), identificar beneficiario/protocolo e gerar documentos de reajuste com base nos modelos e tabelas CSV do projeto.

## Status da versao executavel

Repositorio da versao executavel:
`github.com/PedroRebello1/reajuste_2.0`

No momento, esta versao esta indisponivel.

## Requisitos

- Windows (fluxo atual usa coordenadas fixas de tela).
- Python 3.10+ instalado e no `PATH`.
- Dependencias do `requirements.txt`.
- Arquivos de dados presentes na raiz:
	- `valores.csv`
	- `ans_vida.csv`
- Modelos `.docx` em `modelos/2024` e `modelos/2025`.
- Opcional para gerar PDF automaticamente: LibreOffice instalado em
	`C:\Program Files\LibreOffice\program\soffice.exe`.

## Instalacao

1. Baixe este repositorio (zip) e extraia em uma pasta local.
2. Abra a pasta no VS Code.
3. No terminal da pasta do projeto, instale as dependencias:

```bash
pip install -r requirements.txt
```

## Como executar

Script principal atual:

```bash
python searcher.py
```

Fluxo de execucao:

1. O script abre uma janela para digitar protocolos.
2. Voce pode separar protocolos por `,`, `;`, `.`, ou espaco.
3. Para cada protocolo, o script busca no sistema, pede confirmacoes por janelas e executa o preenchimento.
4. Os arquivos gerados sao salvos em `prontos/`.

Tambem e possivel executar o mesmo fluxo com:

```bash
python test.py
```

## Ajuste de coordenadas (importante)

As automacoes dependem das coordenadas definidas nos arquivos Python (`COORD`).
Se o clique estiver em local errado, ajuste os valores para o seu monitor/resolucao.

Arquivos com coordenadas:

- `searcher.py`
- `autofill_cassi.py`

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
	verifique se o LibreOffice esta instalado no caminho esperado.
- Captura de texto incorreta:
	revise coordenadas e foco da janela do sistema antes de rodar.

## Observacoes

- Durante a execucao, nao mova janelas nem altere resolucao/zoom do sistema.
- Os scripts foram pensados para execucao assistida (com confirmacoes manuais em popups).