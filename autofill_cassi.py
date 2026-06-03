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
    "plano_sys2": (2654, 590),
    "nome": (2830, 710),
    "num_plano": (2600, 630),
    "jump": (2867, 804),
    "jump2": (2867, 250),
    "data_adesao": (2467, 267)
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
    }
}

MES_LIMITE_REAJUSTE_2026 = 7
MES_REAJUSTE_POR_PLANO = {
    "essencial": 6,
    "vida": 11
}
vida_2026_active = False
essencial_2026_active = False

def plano_2026_ativo(tipo_plano):
    if tipo_plano == "vida":
        return vida_2026_active
    if tipo_plano == "essencial":
        return essencial_2026_active
    return True

def calcular_anos_para_gerar(ano_adesao, mes_adesao, tipo_plano=None):
    if tipo_plano == "gdi":
        anos = []
        if ano_adesao < 2024:
            return ["2024", "2025", "2026"]
        elif ano_adesao == 2024:
            return ["2024", "2025", "2026"] if mes_adesao < 4 else ["2025", "2026"]
        elif ano_adesao == 2025:
            return ["2025", "2026"] if mes_adesao < 4 else ["2026"]
        elif ano_adesao == 2026 and mes_adesao < 4:
            return ["2026"]
        return anos

    anos = []
    ativo_2026 = plano_2026_ativo(tipo_plano)

    mes_reajuste = MES_REAJUSTE_POR_PLANO.get(tipo_plano)
    if mes_reajuste and mes_adesao <= (mes_reajuste - 1):
        ano_mesmo_ano = str(ano_adesao)
        if ano_mesmo_ano in {"2024", "2025"} or (ano_mesmo_ano == "2026" and ativo_2026):
            anos.append(ano_mesmo_ano)

    if ano_adesao < 2024:
        anos.extend(["2024", "2025"])
    elif ano_adesao == 2024:
        if "2025" not in anos:
            anos.append("2025")

    if ativo_2026 and ano_adesao <= 2025 and mes_adesao <= MES_LIMITE_REAJUSTE_2026:
        if "2026" not in anos:
            anos.append("2026")

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
        x_fim = 2350

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

def buscar_ans(texto_plano):
    if df_ans is None:
        return "UFANS", "NUMEROANS"
    
    texto_clean = remover_acentos(texto_plano)
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

def fill():
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
        elif "família ii" in texto_lower or "família 2" in texto_lower or "familia 2" in texto_lower or "familia ii" in texto_lower or "fam2" in texto_lower:
            data_limite_fam2 = datetime(2004, 1, 1)
            tipo_plano = "familia2ae" if data_adesao_obj < data_limite_fam2 else "familia2de"
            chave_csv = "fam2"
        elif "família i" in texto_lower or "família 1" in texto_lower or "familia 1" in texto_lower or "fam1" in texto_lower:
            tipo_plano = "familia1"
            chave_csv = "fam1"
        else:
            print("Status: Plano não mapeado nas regras atuais.")

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
                    substituicoes["UFANS"] = uf
                    substituicoes["NUMEROANS"] = ans
                elif tipo_plano == "essencial":
                    substituicoes["MESADESAO"] = mes_por_extenso(data_adesao_obj)
                    substituicoes["NOMEPLANO"] = texto_plano
                    substituicoes["PORCENTAGEMAUMENTO"] = valor_encontrado
                elif tipo_plano == "familia1":
                    substituicoes["MESADESAO"] = mes_por_extenso(data_adesao_obj)
                    substituicoes["PORCENTAGEMAUMENTO"] = valor_encontrado
                elif tipo_plano in ["familia2ae", "familia2de"]:
                    substituicoes["MESADESAO"] = mes_por_extenso(data_adesao_obj)
                    substituicoes["PORCENTAGEMAUMENTO"] = valor_encontrado
                    substituicoes["NOMEPLANO"] = texto_plano
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