import os
import re
import time
import unicodedata
import pyautogui
import pyperclip
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate

try:
    import win32com.client as win32
    import pythoncom
except ImportError:
    win32 = None
    pythoncom = None

CONVERSAO_PDF_DISPONIVEL = True

# --- CARREGAMENTO DOS CSVS ---
PASTA_PRONTOS = "./prontos/"
os.makedirs(PASTA_PRONTOS, exist_ok=True)

try:
    df_valores = pd.read_csv("valores.csv", sep=';')
    df_valores['data_obj'] = pd.to_datetime(df_valores['data'], format='%d/%m/%Y')
except Exception as e:
    print(f"Erro ao carregar valores.csv: {e}")
    df_valores = None

try:
    df_ans = pd.read_csv("ans_vida.csv", sep=';')
except Exception as e:
    print(f"Erro ao carregar ans_vida.csv: {e}")
    df_ans = None

# --- COORDENADAS E DICIONÁRIOS ---
COORD = {
    "plano_sys2": (1183, 594),
    "nome": (1388, 712),
    "num_plano": (1155, 632),
    "jump": (1425, 804),
    "jump2": (1425, 250),
    "data_adesao": (1022, 267)
}

MODELOS_PLANO = {
    "vida": {
        "2024": "./modelos/2024/vida.docx",
        "2025": "./modelos/2025/vida.docx",
        "2026": "./modelos/2026/vida.docx"
    },
    "essencial": {
        "2024": "./modelos/2024/essencial.docx",
        "2025": "./modelos/2025/essencial.docx",
        "2026": "./modelos/2026/essencial.docx"
    },
    "familia1": {
        "2024": "./modelos/2024/familia1.docx",
        "2025": "./modelos/2025/familia1.docx",
        "2026": "./modelos/2026/familia1.docx"
    },
    "familia2ae": {
        "2024": "./modelos/2024/familia2ae.docx",
        "2025": "./modelos/2025/familia2ae.docx",
        "2026": "./modelos/2026/familia2ae.docx"
    },
    "familia2de": {
        "2024": "./modelos/2024/familia2de.docx",
        "2025": "./modelos/2025/familia2de.docx",
        "2026": "./modelos/2026/familia2de.docx"
    },
    "gdi": {
        "2024": "./modelos/2024/gdi.docx",
        "2025": "./modelos/2025/gdi.docx",
        "2026": "./modelos/2026/gdi.docx"
    },
    # FAMÍLIA INTEGRAL / FAMÍLIA PARTICIPATIVO: planos criados em 2026, então só
    # existem avisos de 2026 em diante. São o mesmo plano (mesmo registro ANS e
    # mesmo reajuste em todas as UFs) e só se diferenciam pelo modelo.
    "integral": {
        "2026": "./modelos/2026/integral.docx"
    },
    "participativo": {
        "2026": "./modelos/2026/participativo.docx"
    }
}

# Mês (sempre no dia 1) em que o reajuste de cada plano cai a cada ano.
# Planos família I/II não entram aqui: reajustam no mês de aniversário da adesão.
MES_REAJUSTE_POR_PLANO = {
    "essencial": 6,
    "vida": 11,
    "gdi": 4,
    "integral": 8,
    "participativo": 8
}

# Até que mês os reajustes de um ano ainda em curso já foram liberados.
# Reajuste com mês posterior ao limite ainda não aconteceu, então não gera aviso.
MES_LIMITE_REAJUSTE_2026 = 10
MES_LIMITE_POR_ANO = {2026: MES_LIMITE_REAJUSTE_2026}

vida_2026_active = False
essencial_2026_active = True

def plano_2026_ativo(tipo_plano):
    if tipo_plano == "vida":
        return vida_2026_active
    if tipo_plano == "essencial":
        return essencial_2026_active
    return True

def mes_do_reajuste(tipo_plano, mes_adesao):
    """Mês em que o reajuste do plano cai a cada ano.

    Planos com mês fixo vêm de MES_REAJUSTE_POR_PLANO; família I/II reajustam
    no mês de aniversário da adesão.
    """
    return MES_REAJUSTE_POR_PLANO.get(tipo_plano, mes_adesao)

def calcular_anos_para_gerar(ano_adesao, mes_adesao, tipo_plano=None):
    anos = []
    mes_reajuste = mes_do_reajuste(tipo_plano, mes_adesao)

    # Os anos possíveis são os que têm modelo cadastrado para o plano.
    for ano in sorted(MODELOS_PLANO.get(tipo_plano, {})):
        ano_int = int(ano)

        if ano_int == 2026 and not plano_2026_ativo(tipo_plano):
            continue

        # O reajuste sempre cai no dia 1, então comparar (ano, mês) equivale a
        # comparar as datas cheias: quem aderiu em qualquer dia anterior ao
        # reajuste — mesmo 1 dia antes — recebe o aviso; quem aderiu no dia do
        # reajuste ou depois, não.
        if (ano_adesao, mes_adesao) >= (ano_int, mes_reajuste):
            continue

        # Reajuste de ano ainda em curso que não chegou a acontecer.
        limite = MES_LIMITE_POR_ANO.get(ano_int)
        if limite is not None and mes_reajuste > limite:
            continue

        anos.append(ano)

    return anos

# --- FUNÇÕES AUXILIARES ---

def copiar_texto(coord):
    try:
        pyautogui.click(coord[0], coord[1])
        for _ in range(6):
            pyautogui.hotkey('ctrl', 'shift', 'left')
        pyautogui.hotkey('ctrl', 'c')
        return pyperclip.paste().strip()
    except Exception as e:
        print(f"Erro ao copiar: {e}")
        return ""

def copiar_texto_mouse(coord):
    try:
        if coord in {COORD["jump"], COORD["jump2"]}:
            pyautogui.click(coord[0], coord[1])
            pyautogui.hotkey('ctrl', 'c')
            return pyperclip.paste().strip()

        x_inicio, y_inicio = coord
        x_fim = 850

        pyautogui.click(x_inicio, y_inicio)
        pyautogui.dragTo(x_fim, y_inicio, duration=0.3, button='left')
        pyautogui.hotkey('ctrl', 'c')
        return pyperclip.paste().strip()
    except Exception as e:
        print(f"Erro ao copiar com mouse: {e}")
        return ""

def mes_por_extenso(data_obj):
    meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    return meses.get(data_obj.month, "Mês Inválido")

def remover_acentos(txt):
    return unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('utf-8').lower()

# Apelidos/abreviações regionais que aparecem no nome do plano Vida
# (ex.: "CASSI VIDA BH") mapeados para a cidade correspondente no ans_vida.csv.
# Chave e valor em minúsculas/sem acento. Adicione novos conforme aparecerem.
APELIDOS_ANS = {
    "bh": "belo horizonte",
    "bsb": "brasilia",
    "poa": "porto alegre",
    "floripa": "florianopolis",
    "cwb": "curitiba",
    "rp": "ribeirao preto",
}

def buscar_ans(texto_plano):
    if df_ans is None:
        return "UFANS", "NUMEROANS"
    
    texto_clean = remover_acentos(texto_plano)

    # Expande apelidos regionais (ex.: "bh" -> "belo horizonte") para que o
    # nome abreviado do plano Vida case com a cidade do ans_vida.csv.
    for apelido, cidade in APELIDOS_ANS.items():
        if re.search(rf'\b{apelido}\b', texto_clean):
            texto_clean += " " + cidade

    for _, row in df_ans.iterrows():
        uf_clean = remover_acentos(str(row['UF']))
        
        match = re.search(r'\((.*?)\)', uf_clean)
        if match and match.group(1) in texto_clean:
            return str(row['UF']), str(row['Registro ANS'])
        
        estado = uf_clean[:2]
        if re.search(rf'\b{estado}\b', texto_clean):
            return str(row['UF']), str(row['Registro ANS'])
            
    return "UFANS", "NUMEROANS"

def preencher_e_salvar_documento(caminho_modelo, caminho_saida_docx, substituicoes):
    try:
        doc = DocxTemplate(caminho_modelo)
        doc.render(substituicoes)
        doc.save(caminho_saida_docx)
    except Exception as e:
        print(f"Erro ao processar template {caminho_modelo}: {e}")

def converter_para_pdf(caminho_docx):
    if os.name != "nt":
        print("Erro: Conversão para PDF via Word disponível apenas no Windows.")
        return

    if win32 is None or pythoncom is None:
        print("Erro: pacote 'pywin32' não encontrado. Instale com: pip install pywin32")
        return

    word = None
    doc = None

    try:
        pythoncom.CoInitialize()
        caminho_docx = os.path.abspath(caminho_docx)
        caminho_pdf = os.path.splitext(caminho_docx)[0] + ".pdf"

        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.Options.SaveNormalPrompt = False
        word.AutomationSecurity = 3

        doc = word.Documents.Open(
            caminho_docx,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        doc.ExportAsFixedFormat(caminho_pdf, 17, OpenAfterExport=False)
        doc.Close(False)
        doc = None

        if os.path.exists(caminho_pdf):
            os.remove(caminho_docx)
            print(f"-> Salvo como PDF: {caminho_pdf}")
    except Exception as e:
        print(f"-> Erro na conversão para PDF via Word: {e} (mantendo docx)")
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass

        if word is not None:
            try:
                word.Quit(SaveChanges=0)
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

# =====================================================================
# --- EXECUÇÃO ---
# =====================================================================

def fill(protocolo=None):
    print("\nCapturando dados...")
    time.sleep(0.2)

    texto_plano = copiar_texto_mouse(COORD["plano_sys2"])
    nome_beneficiario = copiar_texto_mouse(COORD["nome"])
    numero_plano = copiar_texto_mouse(COORD["num_plano"])

    pyautogui.click(COORD["jump"][0], COORD["jump"][1])
    pyautogui.click(COORD["jump"][0], COORD["jump"][1])
    data_adesao_str = copiar_texto_mouse(COORD["data_adesao"])

    pyautogui.click(COORD["jump2"][0], COORD["jump2"][1])
    pyautogui.click(COORD["jump2"][0], COORD["jump2"][1])

    print("\n--- Dados capturados ---")
    print(f"Plano: {texto_plano}")
    print(f"Nome do beneficiário: {nome_beneficiario}")
    print(f"Número do plano: {numero_plano}")
    print(f"Data de adesão: {data_adesao_str}")

    try:
        data_adesao_obj = datetime.strptime(data_adesao_str, "%d/%m/%Y")
        ano_adesao = data_adesao_obj.year
        mes_adesao = data_adesao_obj.month
    except ValueError:
        print(f"Erro: A data '{data_adesao_str}' não está no formato correto.")
        data_adesao_obj = None

    if data_adesao_obj:
        texto_lower = texto_plano.lower()
        tipo_plano = None
        chave_csv = None

        if "vida" in texto_lower:
            tipo_plano = "vida"
            chave_csv = "vida"
        elif "essencial" in texto_lower:
            tipo_plano = "essencial"
            chave_csv = texto_plano
        elif "dependentes indiretos" in texto_lower:
            tipo_plano = "gdi"
            chave_csv = "gdi"
        # FAMÍLIA PARTICIPATIVO / FAMÍLIA INTEGRAL vêm com a UF no final
        # ("... – SP"), que é descartada: o plano é o mesmo em todas as UFs.
        # Precisam vir antes de "família i", que casa com "FAMÍLIA INTEGRAL".
        elif "participativo" in texto_lower:
            tipo_plano = "participativo"
            chave_csv = "participativo"
        elif "integral" in texto_lower:
            tipo_plano = "integral"
            chave_csv = "integral"
        elif "família ii" in texto_lower or "família 2" in texto_lower or "familia 2" in texto_lower or "familia ii" in texto_lower or "fam2" in texto_lower:
            data_limite_fam2 = datetime(2004, 1, 1)
            tipo_plano = "familia2ae" if data_adesao_obj < data_limite_fam2 else "familia2de"
            chave_csv = "fam2"
        elif "família i" in texto_lower or "família 1" in texto_lower or "familia 1" in texto_lower or "fam1" in texto_lower:
            tipo_plano = "familia1"
            chave_csv = "fam1"
        else:
            print(f"O participante {nome_beneficiario}, do protocolo {protocolo}, possui um plano não mapeado: {texto_plano}")

        if tipo_plano and df_valores is not None:
            anos_para_gerar = calcular_anos_para_gerar(ano_adesao, mes_adesao, tipo_plano)
            if not anos_para_gerar:
                print("Status: Nenhum documento para gerar com as regras atuais.")
                return

            for ano_alvo in anos_para_gerar:
                if tipo_plano == "essencial":
                    filtro = df_valores['plano'].apply(lambda x: str(x).upper().strip() in chave_csv.upper().strip())
                else:
                    filtro = df_valores['plano'].apply(lambda x: str(x).upper().strip() == chave_csv.upper().strip())

                df_filtrado = df_valores[filtro].copy()
                df_filtrado = df_filtrado[df_filtrado['data_obj'].dt.year == int(ano_alvo)]

                if tipo_plano in ["familia1", "familia2ae", "familia2de"]:
                    df_filtrado = df_filtrado[df_filtrado['data_obj'].dt.month == mes_adesao]

                valor_encontrado = str(df_filtrado.iloc[0]['aumento']) if not df_filtrado.empty else "[VALOR NÃO ENCONTRADO]"

                substituicoes = {
                    "NOMEBENEFICIARIO": nome_beneficiario,
                    "NUMEROPLANO": numero_plano
                }

                if tipo_plano == "vida":
                    uf, ans = buscar_ans(texto_plano)
                    if uf == "UFANS" or ans == "NUMEROANS":
                        print(f"-> ATENÇÃO: UF/ANS não identificada no plano '{texto_plano}'. "
                              f"Verifique o documento e/ou cadastre o apelido em APELIDOS_ANS.")
                    substituicoes["UFANS"] = uf
                    substituicoes["NUMEROANS"] = ans
                elif tipo_plano == "essencial":
                    substituicoes["MESADESAO"] = 'JUNHO'
                    substituicoes["NOMEPLANO"] = texto_plano
                    substituicoes["PORCENTAGEMAUMENTO"] = valor_encontrado
                elif tipo_plano == "familia1":
                    substituicoes["MESADESAO"] = mes_por_extenso(data_adesao_obj)
                    substituicoes["PORCENTAGEMAUMENTO"] = valor_encontrado
                elif tipo_plano in ["familia2ae", "familia2de"]:
                    substituicoes["MESADESAO"] = mes_por_extenso(data_adesao_obj)
                    substituicoes["PORCENTAGEMAUMENTO"] = valor_encontrado
                    substituicoes["NOMEPLANO"] = texto_plano
                elif tipo_plano in ["integral", "participativo"]:
                    # Percentual e mês já estão fixos no modelo: só NOMEBENEFICIARIO
                    # e NUMEROPLANO precisam ser preenchidos.
                    pass
                elif tipo_plano == "gdi":
                    pass

                caminho_modelo = MODELOS_PLANO.get(tipo_plano, {}).get(ano_alvo)
                if caminho_modelo and os.path.exists(caminho_modelo):
                    caminho_saida = f"{PASTA_PRONTOS}reajuste_{nome_beneficiario}_{ano_alvo}.docx"

                    preencher_e_salvar_documento(caminho_modelo, caminho_saida, substituicoes)
                    print(f"\n-> Documento gerado: {caminho_saida}")

                    converter_para_pdf(caminho_saida)
                else:
                    print(f"-> Arquivo de modelo não encontrado: {caminho_modelo}")

if __name__ == "__main__":
    fill()