import os
import sqlite3
import csv
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Separador interno usado pelo Anki para campos múltiplos
SEP = '\x1f'

class AnkiExtractorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Extrator de Frases Anki")
        self.geometry("500x350")

        # Variáveis para armazenar caminhos
        self.file_path_var = ctk.StringVar()
        self.csv_path_var  = ctk.StringVar()

        # Configuração de aparência
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Widgets
        self._create_widgets()

    def _create_widgets(self):
        # Seção Anki (.anki2)
        ctk.CTkLabel(self, text="Arquivo Anki (.anki2):").pack(pady=(20, 5))
        ctk.CTkEntry(self, textvariable=self.file_path_var, width=400, state="readonly").pack()
        ctk.CTkButton(self, text="Selecionar Arquivo", command=self.select_file).pack(pady=5)
        ctk.CTkButton(self, text="Extrair Campos Completos", command=self.extract_phrases).pack(pady=5)
        ctk.CTkButton(self, text="Extrair Apenas Frases", command=self.extract_only_phrases).pack(pady=5)

        # Seção CSV para limpeza após '['
        ctk.CTkLabel(self, text="Arquivo CSV para limpar:").pack(pady=(20, 5))
        ctk.CTkEntry(self, textvariable=self.csv_path_var, width=400, state="readonly").pack()
        ctk.CTkButton(self, text="Selecionar CSV", command=self.select_csv).pack(pady=5)
        ctk.CTkButton(self, text="Limpar CSV", command=self.clean_csv).pack(pady=5)

    def select_file(self):
        """Abre um diálogo para escolher o arquivo collection.anki2"""
        file_path = filedialog.askopenfilename(
            initialdir=os.getcwd(),
            title="Selecione o arquivo collection.anki2",
            filetypes=[("Coleção Anki", "*.anki2"), ("Todos os arquivos", "*")]
        )
        if file_path:
            self.file_path_var.set(file_path)

    def select_csv(self):
        """Abre um diálogo para escolher o CSV a ser limpo"""
        csv_path = filedialog.askopenfilename(
            initialdir=os.getcwd(),
            title="Selecione o arquivo CSV",
            filetypes=[("CSV", "*.csv"), ("Todos os arquivos", "*")]
        )
        if csv_path:
            self.csv_path_var.set(csv_path)

    def extract_phrases(self):
        """Extrai id, campo1 e campo2 de cada nota"""
        path = self.file_path_var.get()
        if not path:
            messagebox.showwarning("Aviso", "Por favor, selecione primeiro o arquivo .anki2.")
            return
        try:
            conn = sqlite3.connect(path)
            cur = conn.cursor()
            output_csv = os.path.join(os.path.dirname(path), 'frases_extraidas.csv')
            with open(output_csv, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['id', 'campo1', 'campo2'])
                for note_id, flds in cur.execute('SELECT id, flds FROM notes'):
                    parts = flds.split(SEP)
                    campo1 = parts[0] if len(parts) > 0 else ''
                    campo2 = parts[1] if len(parts) > 1 else ''
                    writer.writerow([note_id, campo1, campo2])
            conn.close()
            messagebox.showinfo("Sucesso", f"Campos extraídos em:\n{output_csv}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao extrair campos:\n{e}")

    def extract_only_phrases(self):
        """Extrai apenas o primeiro campo (frase) de cada nota"""
        path = self.file_path_var.get()
        if not path:
            messagebox.showwarning("Aviso", "Por favor, selecione primeiro o arquivo .anki2.")
            return
        try:
            conn = sqlite3.connect(path)
            cur = conn.cursor()
            output_csv = os.path.join(os.path.dirname(path), 'apenas_frases.csv')
            with open(output_csv, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['id', 'frase'])
                for note_id, flds in cur.execute('SELECT id, flds FROM notes'):
                    frase = flds.split(SEP)[0]
                    writer.writerow([note_id, frase])
            conn.close()
            messagebox.showinfo("Sucesso", f"Apenas frases extraídas em:\n{output_csv}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao extrair apenas frases:\n{e}")

    def clean_csv(self):
        """Limpa o CSV removendo tudo que vier depois da primeira '[' em cada célula"""
        path = self.csv_path_var.get()
        if not path:
            messagebox.showwarning("Aviso", "Por favor, selecione primeiro o arquivo CSV.")
            return
        try:
            output_csv = os.path.join(os.path.dirname(path), 'frases_limpo.csv')
            with open(path, newline='', encoding='utf-8') as fin, \
                 open(output_csv, 'w', newline='', encoding='utf-8') as fout:
                reader = csv.reader(fin)
                writer = csv.writer(fout)
                header = next(reader, None)
                if header:
                    writer.writerow(header)
                for row in reader:
                    new_row = [cell.split('[', 1)[0] for cell in row]
                    writer.writerow(new_row)
            messagebox.showinfo("Sucesso", f"CSV limpo gerado em:\n{output_csv}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao limpar CSV:\n{e}")

if __name__ == '__main__':
    app = AnkiExtractorApp()
    app.mainloop()
