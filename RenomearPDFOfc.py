import customtkinter as ctk
from tkinter import filedialog, messagebox
import PyPDF2
import os
import re
import unicodedata
from datetime import datetime


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class PDFRenamerCustomerThink:
    def __init__(self):
        self.app = ctk.CTk()
        self.app.title("CustomerThink | PDF Renamer RH")
        self.app.geometry("1100x760")
        self.app.minsize(980, 700)

        self.files = []
        self.current_index = 0
        self.pdf_path = ""
        self.cpf = ""

        self.build_interface()

    # --------------------------------------------------
    # Interface
    # --------------------------------------------------
    def build_interface(self):
        header = ctk.CTkFrame(self.app, height=70, fg_color="#1F3A5F", corner_radius=0)
        header.pack(fill="x")

        title = ctk.CTkLabel(
            header,
            text="PDF CPF / Nome Renamer",
            font=("Segoe UI", 24, "bold"),
            text_color="white"
        )
        title.pack(pady=(14, 4))

        subtitle = ctk.CTkLabel(
            header,
            text="Selecione o nome no texto e pressione Enter para renomear e seguir para o próximo PDF",
            font=("Segoe UI", 12),
            text_color="white"
        )
        subtitle.pack(pady=(0, 10))

        main = ctk.CTkFrame(self.app)
        main.pack(fill="both", expand=True, padx=18, pady=18)

        topbar = ctk.CTkFrame(main)
        topbar.pack(fill="x", pady=(0, 12))

        self.btn_folder = ctk.CTkButton(
            topbar,
            text="Selecionar Pasta de PDFs",
            command=self.select_folder,
            width=220,
            height=40
        )
        self.btn_folder.pack(side="left", padx=(10, 8), pady=10)

        self.btn_prev = ctk.CTkButton(
            topbar,
            text="Voltar Anterior",
            command=self.previous_pdf,
            width=140,
            height=40
        )
        self.btn_prev.pack(side="left", padx=8, pady=10)

        self.btn_skip = ctk.CTkButton(
            topbar,
            text="Pular PDF",
            command=self.skip_pdf,
            width=120,
            height=40
        )
        self.btn_skip.pack(side="left", padx=8, pady=10)

        self.btn_rename = ctk.CTkButton(
            topbar,
            text="Renomear e Próximo",
            command=self.rename_selected,
            width=180,
            height=40
        )
        self.btn_rename.pack(side="left", padx=8, pady=10)

        self.lbl_status = ctk.CTkLabel(
            topbar,
            text="0/0",
            font=("Segoe UI", 14, "bold")
        )
        self.lbl_status.pack(side="right", padx=12, pady=10)

        info = ctk.CTkFrame(main)
        info.pack(fill="x", pady=(0, 12))

        self.lbl_file = ctk.CTkLabel(
            info,
            text="Arquivo atual: -",
            anchor="w",
            justify="left",
            font=("Segoe UI", 13, "bold")
        )
        self.lbl_file.pack(fill="x", padx=12, pady=(10, 6))

        self.lbl_cpf = ctk.CTkLabel(
            info,
            text="CPF encontrado: -",
            anchor="w",
            justify="left",
            font=("Segoe UI", 13)
        )
        self.lbl_cpf.pack(fill="x", padx=12, pady=4)

        self.lbl_sugestao = ctk.CTkLabel(
            info,
            text="Nome sugerido: -",
            anchor="w",
            justify="left",
            font=("Segoe UI", 13)
        )
        self.lbl_sugestao.pack(fill="x", padx=12, pady=(4, 10))

        text_frame = ctk.CTkFrame(main)
        text_frame.pack(fill="both", expand=True, pady=(0, 12))

        self.textbox = ctk.CTkTextbox(text_frame, font=("Consolas", 13))
        self.textbox.pack(fill="both", expand=True, padx=10, pady=10)

        bottom = ctk.CTkFrame(main)
        bottom.pack(fill="x")

        help_text = (
            "Atalhos: Enter = renomear e próximo | Shift+Enter = pular | "
            "Ao selecionar um trecho no texto, o sistema usa a seleção como nome do arquivo."
        )
        self.lbl_help = ctk.CTkLabel(
            bottom,
            text=help_text,
            anchor="w",
            justify="left",
            font=("Segoe UI", 12)
        )
        self.lbl_help.pack(fill="x", padx=12, pady=(10, 6))

        self.log = ctk.CTkTextbox(bottom, height=150, font=("Consolas", 12))
        self.log.pack(fill="x", padx=12, pady=(0, 12))

        self.textbox.bind("<Return>", self.enter_renomear)
        self.textbox.bind("<Shift-Return>", self.shift_enter_skip)
        self.textbox.bind("<ButtonRelease-1>", self.on_text_selection)
        self.textbox.bind("<KeyRelease>", self.on_text_selection)

        self.update_navigation_buttons()

    # --------------------------------------------------
    # Util
    # --------------------------------------------------
    def write_log(self, text):
        now = datetime.now().strftime("%H:%M:%S")
        self.log.insert("end", f"[{now}] {text}\n")
        self.log.see("end")

    def normalize_name(self, name):
        name = unicodedata.normalize("NFKD", name)
        name = "".join(c for c in name if not unicodedata.combining(c))
        name = name.upper()
        name = re.sub(r"[^\w\s]", "", name)
        name = re.sub(r"\s+", "_", name)
        name = re.sub(r"_+", "_", name)
        return name.strip("_")

    def clean_selected_name(self, text):
        text = text.strip()

        substituicoes = [
            "Nome:",
            "NOME:",
            "Nome",
            "NOME",
            "CPF:",
            "CPF",
            "Beneficiário:",
            "BENEFICIÁRIO:",
            "BENEFICIARIO:",
            "BENEFICIÁRIO",
            "BENEFICIARIO"
        ]

        for item in substituicoes:
            text = text.replace(item, "")

        text = text.strip(" :-\n\t")
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    def get_selected_text(self):
        try:
            ranges = self.textbox.tag_ranges("sel")
            if not ranges:
                return ""
            return self.textbox.get(ranges[0], ranges[1]).strip()
        except Exception:
            return ""

    def update_status_label(self):
        total = len(self.files)
        atual = self.current_index + 1 if total > 0 and self.current_index < total else total
        self.lbl_status.configure(text=f"{atual}/{total}")

    def update_file_info(self):
        if self.pdf_path:
            self.lbl_file.configure(text=f"Arquivo atual: {self.pdf_path}")
        else:
            self.lbl_file.configure(text="Arquivo atual: -")

        if self.cpf:
            self.lbl_cpf.configure(text=f"CPF encontrado: {self.cpf}")
        else:
            self.lbl_cpf.configure(text="CPF encontrado: -")

    def update_navigation_buttons(self):
        tem_arquivos = len(self.files) > 0
        anterior_habilitado = tem_arquivos and self.current_index > 0
        atual_habilitado = tem_arquivos and self.current_index < len(self.files)

        self.btn_prev.configure(state="normal" if anterior_habilitado else "disabled")
        self.btn_skip.configure(state="normal" if atual_habilitado else "disabled")
        self.btn_rename.configure(state="normal" if atual_habilitado else "disabled")

    def update_suggestion_label(self):
        selecionado = self.get_selected_text()
        if not selecionado:
            self.lbl_sugestao.configure(text="Nome sugerido: -")
            return

        nome_limpo = self.clean_selected_name(selecionado)
        nome_final = self.normalize_name(nome_limpo)

        if not nome_final:
            self.lbl_sugestao.configure(text="Nome sugerido: -")
            return

        sugestao = f"{self.cpf}_{nome_final}.pdf" if self.cpf else f"{nome_final}.pdf"
        self.lbl_sugestao.configure(text=f"Nome sugerido: {sugestao}")

    # --------------------------------------------------
    # Eventos do texto
    # --------------------------------------------------
    def on_text_selection(self, event=None):
        self.update_suggestion_label()

    def enter_renomear(self, event):
        self.rename_selected()
        return "break"

    def shift_enter_skip(self, event):
        self.skip_pdf()
        return "break"

    # --------------------------------------------------
    # Selecionar pasta
    # --------------------------------------------------
    def select_folder(self):
        folder = filedialog.askdirectory(title="Selecione a pasta com os PDFs")

        if not folder:
            return

        self.files = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith(".pdf")
        ]
        self.files.sort(key=lambda x: os.path.basename(x).lower())

        self.current_index = 0
        self.pdf_path = ""
        self.cpf = ""

        self.write_log(f"{len(self.files)} PDFs encontrados em: {folder}")

        if not self.files:
            self.textbox.delete("1.0", "end")
            self.lbl_file.configure(text="Arquivo atual: -")
            self.lbl_cpf.configure(text="CPF encontrado: -")
            self.lbl_sugestao.configure(text="Nome sugerido: -")
            self.update_status_label()
            self.update_navigation_buttons()
            messagebox.showwarning("Aviso", "Nenhum PDF encontrado na pasta selecionada.")
            return

        self.load_pdf()

    # --------------------------------------------------
    # Leitura PDF
    # --------------------------------------------------
    def extract_text_from_pdf(self, pdf_path):
        text = ""

        with open(pdf_path, "rb") as f:
            pdf = PyPDF2.PdfReader(f)

            for page in pdf.pages:
                try:
                    text += page.extract_text() or ""
                    text += "\n"
                except Exception:
                    pass

        return text

    def find_cpf_in_text(self, text):
        candidatos = re.findall(r"(?<!\d)(?:\d[\.\-\s]*){10}\d(?!\d)", text)

        for candidato in candidatos:
            cpf_limpo = re.sub(r"\D", "", candidato)
            if len(cpf_limpo) == 11:
                return cpf_limpo

        return "SEMCPF"

    def load_pdf(self):
        if self.current_index >= len(self.files):
            self.textbox.delete("1.0", "end")
            self.pdf_path = ""
            self.cpf = ""
            self.update_status_label()
            self.update_file_info()
            self.update_suggestion_label()
            self.update_navigation_buttons()
            self.write_log("Todos os PDFs foram processados")
            messagebox.showinfo("Concluído", "Todos os PDFs foram processados.")
            return

        self.pdf_path = self.files[self.current_index]

        self.write_log(f"Abrindo PDF {self.current_index + 1}/{len(self.files)}")
        self.write_log(self.pdf_path)

        try:
            text = self.extract_text_from_pdf(self.pdf_path)
        except Exception as e:
            text = ""
            self.write_log(f"Erro leitura PDF: {e}")

        self.textbox.delete("1.0", "end")
        self.textbox.insert("1.0", text)

        self.cpf = self.find_cpf_in_text(text)

        if self.cpf != "SEMCPF":
            self.write_log(f"CPF encontrado: {self.cpf}")
        else:
            self.write_log("CPF não encontrado")

        self.update_status_label()
        self.update_file_info()
        self.update_suggestion_label()
        self.update_navigation_buttons()

        self.textbox.focus_set()

    # --------------------------------------------------
    # Navegação
    # --------------------------------------------------
    def previous_pdf(self):
        if not self.files:
            self.write_log("Nenhum PDF carregado")
            return

        if self.current_index <= 0:
            self.write_log("Você já está no primeiro PDF")
            return

        self.current_index -= 1
        self.load_pdf()

    def skip_pdf(self):
        if not self.pdf_path:
            self.write_log("Nenhum PDF carregado")
            return

        self.write_log(f"PDF pulado: {os.path.basename(self.pdf_path)}")
        self.current_index += 1
        self.load_pdf()

    # --------------------------------------------------
    # Renomear
    # --------------------------------------------------
    def build_new_name(self, nome_base):
        folder = os.path.dirname(self.pdf_path)
        nome_arquivo = f"{self.cpf}_{nome_base}.pdf"
        novo_caminho = os.path.join(folder, nome_arquivo)

        contador = 1
        while os.path.exists(novo_caminho) and os.path.abspath(novo_caminho) != os.path.abspath(self.pdf_path):
            nome_arquivo = f"{self.cpf}_{nome_base}_{contador}.pdf"
            novo_caminho = os.path.join(folder, nome_arquivo)
            contador += 1

        return nome_arquivo, novo_caminho

    def rename_selected(self):
        if not self.pdf_path:
            self.write_log("Nenhum PDF carregado")
            return

        selected_text = self.get_selected_text()

        if not selected_text:
            self.write_log("Selecione o nome no texto antes de pressionar ENTER")
            return

        nome_limpo = self.clean_selected_name(selected_text)

        if not nome_limpo:
            self.write_log("Não foi possível extrair um nome válido da seleção")
            return

        nome_normalizado = self.normalize_name(nome_limpo)

        if not nome_normalizado:
            self.write_log("Nome inválido após normalização")
            return

        novo_nome, novo_path = self.build_new_name(nome_normalizado)

        try:
            os.rename(self.pdf_path, novo_path)
            antigo_nome = os.path.basename(self.pdf_path)

            self.write_log(f"Renomeado: {antigo_nome} -> {novo_nome}")

            self.files[self.current_index] = novo_path
            self.current_index += 1
            self.load_pdf()

        except Exception as e:
            self.write_log(f"Erro renomear: {e}")
            messagebox.showerror("Erro", f"Erro ao renomear o arquivo:\n{e}")

    # --------------------------------------------------
    # Run
    # --------------------------------------------------
    def run(self):
        self.app.mainloop()


if __name__ == "__main__":
    sistema = PDFRenamerCustomerThink()
    sistema.run()
