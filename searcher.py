import importlib
import time
import tkinter as tk
from tkinter import messagebox, simpledialog
import re
import pyautogui
import pyperclip


COORD = {
    "searchbar": (352, 147),
    "protocolo": (387, 342),
    "searchbar2": (789, 251),
}


def parse_protocolos(raw_value: str) -> list[str]:
    return [item for item in re.split(r"[,.;\s]+", raw_value.strip()) if item]


def show_ok_message(text: str) -> None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo("Confirmação", text, parent=root)
    root.destroy()


def ask_protocolos() -> str:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    value = simpledialog.askstring(
        "Protocolos",
        "Digite os números separados por vírgula:",
        parent=root,
    )
    root.destroy()
    return value or ""


def search_protocolo(numero: str) -> None:
    pyautogui.click(*COORD["searchbar"])
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "a")
    pyperclip.copy(numero)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.click(*COORD["protocolo"])


def search_beneficiario(nome: str) -> None:
    delay = 0
    pyautogui.click(*COORD["searchbar2"])
    pyautogui.click(*COORD["searchbar2"])
    time.sleep(delay)
    pyautogui.press("f2")
    time.sleep(delay)
    pyautogui.hotkey("shift", "\\")
    pyperclip.copy(nome)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(delay)
    pyautogui.press("enter")


def run_fill() -> None:
    module = importlib.import_module("autofill_cassi")
    fill_function = getattr(module, "fill", None)
    if not callable(fill_function):
        raise AttributeError(
            "Erro importando fill()"
        )
    fill_function()


def main() -> None:
    time.sleep(2)  # Tempo para o usuário se preparar
    raw_input = ask_protocolos()
    protocolos = parse_protocolos(raw_input)

    if not protocolos:
        print("Nenhum número válido foi informado.")
        return

    for numero in protocolos:
        search_protocolo(numero)

        show_ok_message("Copie o nome do beneficiário e clique em OK.")
        nome_beneficiario = pyperclip.paste().strip()
        if not nome_beneficiario:
            print(f"Nome vazio para o protocolo {numero}. Pulando.")
            continue

        search_beneficiario(nome_beneficiario)
        show_ok_message("Selecione o beneficiário e clique em OK.")

        try:
            run_fill()
            print(f"Protocolo {numero} processado com sucesso.")
        except Exception as exc:
            print(f"Erro ao executar fill() para o protocolo {numero}: {exc}")


if __name__ == "__main__":
    main()