import customtkinter as ctk
from tkinter import filedialog
import PyPDF2
import os
import re
import unicodedata
from datetime import datetime

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

"""
cd C:\_RPA\APPaPDF

C:/Users/99andsouza/AppData/Local/Programs/Python/Python310/python.exe -m cx_Freeze .\RenomearVo01.py `
    --target-dir dist_renomear_vo `
    --base-name Win32GUI `
    --packages pandas,pyautogui,customtkinter,openpyxl,pyxlsb,keyboard,pyperclip,pdfplumber,PyPDF2 `
    --includes tkinter,customtkinter,pandas,pyautogui,keyboard,pyperclip,pdfplumber,PyPDF2
"""


class PDFRenamerCustomerThink:

    def __init__(self):

        self.app = ctk.CTk()
        self.app.title("CustomerThink | PDF Renamer RH - Igarapé Digital")
        self.app.geometry("900x650")

        self.files = []
        self.current_index = 0
        self.pdf_path = ""
        self.cpf = ""

        self.build_interface()

    # ----------------------------
    # Interface
    # ----------------------------

    def build_interface(self):

        header = ctk.CTkFrame(self.app, height=60, fg_color="#1F3A5F")
        header.pack(fill="x")

        title = ctk.CTkLabel(
            header,
            text="PDF CPF / Nome Renamer",
            font=("Segoe UI", 22, "bold"),
            text_color="white"
        )
        title.pack(pady=15)

        body = ctk.CTkFrame(self.app)
        body.pack(fill="both", expand=True, padx=20, pady=20)

        btn_pdf = ctk.CTkButton(
            body,
            text="Selecionar Pasta de PDFs",
            command=self.select_folder,
            width=220,
            height=40
        )
        btn_pdf.pack(pady=8)

        btn_rename = ctk.CTkButton(
            body,
            text="Renomear e Próximo",
            command=self.rename_selected,
            width=220,
            height=40
        )
        btn_rename.pack(pady=8)

        self.textbox = ctk.CTkTextbox(body, height=400)
        self.textbox.pack(fill="both", expand=True, pady=15)

        self.log = ctk.CTkTextbox(body, height=120)
        self.log.pack(fill="x", pady=5)

    # ----------------------------
    # Log
    # ----------------------------

    def write_log(self, text):

        now = datetime.now().strftime("%H:%M:%S")

        self.log.insert("end", f"[{now}] {text}\n")
        self.log.see("end")

    # ----------------------------
    # Normalizar nome
    # ----------------------------

    def normalize_name(self, name):

        name = unicodedata.normalize("NFKD", name)
        name = "".join(c for c in name if not unicodedata.combining(c))

        name = name.upper()

        name = re.sub(r"[^\w\s]", "", name)

        name = re.sub(r"\s+", "_", name)

        return name

    # ----------------------------
    # Selecionar pasta
    # ----------------------------

    def select_folder(self):

        folder = filedialog.askdirectory()

        if not folder:
            return

        self.files = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith(".pdf")
        ]

        self.current_index = 0

        self.write_log(f"{len(self.files)} PDFs encontrados")

        self.load_pdf()

    # ----------------------------
    # Carregar PDF
    # ----------------------------

    def load_pdf(self):

        if self.current_index >= len(self.files):

            self.write_log("Todos os PDFs foram processados")
            return

        self.pdf_path = self.files[self.current_index]

        self.write_log(f"Abrindo PDF {self.current_index+1}/{len(self.files)}")
        self.write_log(self.pdf_path)

        text = ""

        try:

            with open(self.pdf_path, "rb") as f:

                pdf = PyPDF2.PdfReader(f)

                for page in pdf.pages:
                    text += page.extract_text() or ""

        except Exception as e:

            self.write_log(f"Erro leitura PDF: {e}")

        self.textbox.delete("1.0", "end")
        self.textbox.insert("1.0", text)

        cpf_match = re.search(r"\d{3}\.\d{3}\.\d{3}-\d{2}", text)

        if cpf_match:

            self.cpf = re.sub(r"\D", "", cpf_match.group())

            self.write_log(f"CPF encontrado: {self.cpf}")

        else:

            self.cpf = "SEMCPF"
            self.write_log("CPF não encontrado")

    # ----------------------------
    # Renomear e avançar
    # ----------------------------

    def rename_selected(self):

        if not self.pdf_path:

            self.write_log("Nenhum PDF carregado")
            return

        try:

            ranges = self.textbox.tag_ranges("sel")

            if not ranges:
                self.write_log("Selecione o nome no texto")
                return

            selected_text = self.textbox.get(ranges[0], ranges[1])

        except:

            self.write_log("Erro ao capturar seleção")
            return

        nome = selected_text.strip()

        nome = nome.replace("Nome:", "")
        nome = nome.replace("CPF:", "")

        nome = self.normalize_name(nome)

        folder = os.path.dirname(self.pdf_path)

        new_name = f"{self.cpf}_{nome}.pdf"

        new_path = os.path.join(folder, new_name)

        contador = 1

        while os.path.exists(new_path):

            new_name = f"{self.cpf}_{nome}_{contador}.pdf"
            new_path = os.path.join(folder, new_name)

            contador += 1

        try:

            os.rename(self.pdf_path, new_path)

            self.write_log(f"Renomeado -> {new_name}")

        except Exception as e:

            self.write_log(f"Erro renomear: {e}")

        # avançar para próximo
        self.current_index += 1

        self.load_pdf()

    # ----------------------------

    def run(self):

        self.app.mainloop()


if __name__ == "__main__":

    sistema = PDFRenamerCustomerThink()

    sistema.run()
