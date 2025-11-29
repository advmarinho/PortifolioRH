# -*- coding: utf-8 -*-
"""
Preenchedor ADP (GUI) v6.5 — F8 EXCLUSIVO + leitura linha 1
Autor: Anderson + ChatGPT

Novidades v6.5:
- Lê tudo a partir da linha 1 (não pula cabeçalhos).
- Aba "Transpor" é convertida de formulário vertical (coluna A=campo, B=valor) para um registro 1 linha.
- Rótulos viram colunas preservando numeração (ex.: '1_Nome', '11_Funcao').
- Mantém F8 exclusivo, reset memória, suporte .xlsx/.xlsm/.xlsb.
- Novo modo opcional: leitura horizontal de várias linhas (avança para próxima linha ao fim das colunas).

python -m cx_Freeze preenchedor_adp_gui_v65.py `
    --target-dir dist_preenchedor `
    --base-name win32gui `
    --packages pandas,pyautogui,customtkinter,openpyxl,pyxlsb,pynput `
    --includes pynput,pynput.keyboard


"""

import re
import time
import pandas as pd
import pyautogui as pg
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path

try:
    from pynput import keyboard as pk
    HAS_PYNPUT = True
except Exception:
    HAS_PYNPUT = False

pg.PAUSE = 0.0
pg.FAILSAFE = True

APP_TITLE = "Preenchedor ADP (GUI) v6.5 — F8 exclusivo | F6 pula | Ctrl+K ignora"

SKIP_PREFIXES = ("#", "[skip]", "skip_")
SKIP_SUFFIXES = ("_skip",)

# ---------- Normalização ----------
def to_str(v):
    if v is None:
        return ""
    s = str(v)
    return "" if s.lower() == "nan" else s

def is_template_skip(colname: str) -> bool:
    n = colname.strip().lower()
    for p in SKIP_PREFIXES:
        if n.startswith(p):
            return True
    for s in SKIP_SUFFIXES:
        if n.endswith(s):
            return True
    return False

def _clean_label_keep_number(label: str) -> str:
    """Transforma '11 Função (lookup)' em '11_Funcao' preservando número inicial e '#'."""
    if label is None:
        return ""

    s = str(label).strip()
    starts_with_hash = s.startswith("#")

    # Captura número inicial + resto
    m = re.match(r"^\s*#?\s*(\d+)\s*[-.:)]*\s*(.*)", s)
    if m:
        num, rest = m.groups()
        s = f"{num}_{rest.strip()}"
    else:
        s = s

    # Remove sufixos entre parênteses
    s = re.sub(r"\s*\(.*?\)\s*$", "", s)

    # Normaliza espaços
    s = re.sub(r"\s+", " ", s).strip()

    if starts_with_hash and not s.startswith("#"):
        s = "#" + s

    return s

def sheet_looks_like_vertical_form(df: pd.DataFrame) -> bool:
    """Heurística: formulário vertical típico (coluna A=rótulo, B=valor)."""
    if df is None or df.shape[1] < 2:
        return False
    notnull_A = df.iloc[:, 0].notna().sum()
    notnull_B = df.iloc[:, 1].notna().sum()
    return notnull_A >= 5 and notnull_B >= 1

def vertical_form_to_one_row(df: pd.DataFrame) -> pd.DataFrame:
    """Converte formulário vertical em um DataFrame 1 linha com colunas numeradas."""
    labels = df.iloc[:, 0].astype(object).where(pd.notna(df.iloc[:, 0]), None).tolist()
    values = df.iloc[:, 1].tolist()

    record = {}
    used = set()
    for raw_label, val in zip(labels, values):
        if raw_label is None:
            continue
        col = _clean_label_keep_number(raw_label)
        if not col:
            continue
        base = col
        suffix = 2
        while col in used:
            col = f"{base}({suffix})"
            suffix += 1
        used.add(col)
        record[col] = val

    return pd.DataFrame([record]) if record else df

# ---------- App ----------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1100x680")
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.df = None
        self.registro = {}
        self.campos = []
        self.i = 0

        # controle de múltiplas linhas horizontais
        self.row_start = 0
        self.row_end = 0
        self.row_current = 0

        self._debounce_secs = 0.30
        self._last_enviar_at = 0.0
        self._sending = False

        self.hk_listener = None
        self.global_on = False
        self._local_f8_bound = False

        # Top
        self.frame_top = ctk.CTkFrame(self)
        self.frame_top.pack(fill="x", padx=12, pady=(12, 6))

        self.entry_arquivo = ctk.CTkEntry(
            self.frame_top,
            placeholder_text="Arquivo Excel (.xlsx, .xlsm, .xlsb)"
        )
        self.entry_arquivo.pack(side="left", fill="x", expand=True, padx=(8, 6), pady=8)

        self.btn_arquivo = ctk.CTkButton(
            self.frame_top,
            text="Escolher...",
            command=self.escolher_arquivo,
            width=120
        )
        self.btn_arquivo.pack(side="left", padx=6, pady=8)

        self.cmb_sheet = ctk.CTkComboBox(self.frame_top, values=[], width=240)
        self.cmb_sheet.set("Transpor")
        self.cmb_sheet.pack(side="left", padx=6, pady=8)

        self.entry_linha = ctk.CTkEntry(
            self.frame_top,
            width=80,
            placeholder_text="Linha (1..)"
        )
        self.entry_linha.insert(0, "1")
        self.entry_linha.pack(side="left", padx=6, pady=8)

        self.btn_carregar = ctk.CTkButton(
            self.frame_top,
            text="Carregar",
            command=self.carregar_dados,
            width=120
        )
        self.btn_carregar.pack(side="left", padx=6, pady=8)

        # Mid
        self.frame_mid = ctk.CTkFrame(self)
        self.frame_mid.pack(fill="both", expand=True, padx=12, pady=6)

        self.txt_preview = ctk.CTkTextbox(self.frame_mid, width=660, height=420)
        self.txt_preview.pack(side="left", fill="both", expand=True, padx=(8, 6), pady=8)

        self.frame_ordem = ctk.CTkFrame(self.frame_mid)
        self.frame_ordem.pack(side="left", fill="y", padx=(6, 8), pady=8)

        self.lbl_ordem = ctk.CTkLabel(
            self.frame_ordem,
            text="Ordem de campos (marcados com [SKIP] serão ignorados):"
        )
        self.lbl_ordem.pack(pady=(8, 4))

        self.listbox = ctk.CTkTextbox(self.frame_ordem, width=360, height=380)
        self.listbox.pack(padx=6, pady=4)

        # Bottom
        self.frame_bot = ctk.CTkFrame(self)
        self.frame_bot.pack(fill="x", padx=12, pady=(6, 12))

        self.chk_tab = ctk.CTkCheckBox(
            self.frame_bot,
            text="Enviar TAB após digitar (opcional)",
            onvalue=True,
            offvalue=False
        )
        self.chk_tab.deselect()
        self.chk_tab.pack(side="left", padx=8, pady=8)

        # NOVO: modo leitura horizontal multi-linhas
        self.chk_multilinha = ctk.CTkCheckBox(
            self.frame_bot,
            text="Ler várias linhas (horizontal)",
            onvalue=True,
            offvalue=False
        )
        self.chk_multilinha.deselect()
        self.chk_multilinha.pack(side="left", padx=8, pady=8)

        self.btn_iniciar = ctk.CTkButton(
            self.frame_bot,
            text="Iniciar",
            command=self.iniciar,
            fg_color="#198754",
            hover_color="#157347"
        )
        self.btn_iniciar.pack(side="left", padx=8, pady=8)

        self.btn_hotkey = ctk.CTkButton(
            self.frame_bot,
            text="Ativar HOTKEY Global (F8/F6)",
            command=self.toggle_global,
            fg_color="#0d6efd",
            hover_color="#0b5ed7"
        )
        self.btn_hotkey.pack(side="left", padx=8, pady=8)

        # se não tiver pynput, desabilita global
        if not HAS_PYNPUT:
            self.btn_hotkey.configure(
                state="disabled",
                text="HOTKEY Global indisponível (pynput ausente)"
            )

        self.btn_pular = ctk.CTkButton(
            self.frame_bot,
            text="Pular (F6)",
            command=lambda: self.on_pular(marcar=False)
        )
        self.btn_pular.pack(side="left", padx=8, pady=8)

        self.btn_toggle_skip = ctk.CTkButton(
            self.frame_bot,
            text="Ignorar/Restaurar (Ctrl+K)",
            command=self.toggle_ignore
        )
        self.btn_toggle_skip.pack(side="left", padx=8, pady=8)

        self.btn_voltar = ctk.CTkButton(
            self.frame_bot,
            text="Voltar (F9)",
            command=self.on_voltar
        )
        self.btn_voltar.pack(side="left", padx=8, pady=8)

        self.btn_reset = ctk.CTkButton(
            self.frame_bot,
            text="Resetar Memória",
            command=self.resetar_memoria,
            fg_color="#dc3545",
            hover_color="#bb2d3b"
        )
        self.btn_reset.pack(side="left", padx=8, pady=8)

        self.lbl_status = ctk.CTkLabel(self.frame_bot, text="Aguardando arquivo...")
        self.lbl_status.pack(side="left", padx=12, pady=8)

        self._bind_local_keys()
        self.update_preview("Selecione o Excel. Lê desde a linha 1. Aba 'Transpor' será convertida.")

    # ----- Bindings -----
    def _bind_local_keys(self):
        self.bind("<F8>", self._on_f8_local)
        self._local_f8_bound = True
        self.bind("<F6>", lambda e: self.on_pular(marcar=False))
        self.bind("<Control-k>", lambda e: self.toggle_ignore())
        self.bind("<F9>", lambda e: self.on_voltar())

    def _unbind_local_f8(self):
        if self._local_f8_bound:
            try:
                self.unbind("<F8>")
            except Exception:
                pass
            self._local_f8_bound = False

    def _has_modifiers(self, event) -> bool:
        st = getattr(event, "state", 0)
        return bool(st & 0x0001 or st & 0x0004 or st & 0x0008)

    def _on_f8_local(self, event):
        if self._has_modifiers(event):
            return
        self.on_enviar(avancar=True)

    # ----- Hotkey global -----
    def toggle_global(self):
        if not HAS_PYNPUT:
            messagebox.showerror("Hotkey Global", "Instale 'pynput'.")
            return
        if self.df is None:
            messagebox.showwarning("Iniciar", "Carregue o Excel primeiro.")
            return
        if not self.global_on:
            try:
                self._unbind_local_f8()
                mapping = {
                    "<f8>": lambda: self.after(0, lambda: self.on_enviar(avancar=True)),
                    "<f6>": lambda: self.after(0, lambda: self.on_pular(marcar=False)),
                    "<ctrl>+k": lambda: self.after(0, self.toggle_ignore),
                    "<f9>": lambda: self.after(0, self.on_voltar),
                }
                self.hk_listener = pk.GlobalHotKeys(mapping)
                self.hk_listener.start()
                self.global_on = True
                self.btn_hotkey.configure(
                    text="Desativar HOTKEY Global",
                    fg_color="#6c757d",
                    hover_color="#5c636a"
                )
                self.lbl_status.configure(
                    text="Ativo: F8 envia, F6 pula, Ctrl+K ignora, F9 volta."
                )
            except Exception as e:
                messagebox.showerror("Hotkey Global", f"Falha:\n{e}")
        else:
            try:
                if self.hk_listener:
                    self.hk_listener.stop()
            except Exception:
                pass
            self.hk_listener = None
            self.global_on = False
            self.bind("<F8>", self._on_f8_local)
            self._local_f8_bound = True
            self.btn_hotkey.configure(
                text="Ativar HOTKEY Global (F8/F6)",
                fg_color="#0d6efd",
                hover_color="#0b5ed7"
            )
            self.lbl_status.configure(text="Hotkey global desativada. F8 local reativado.")

    # ----- Fluxo -----
    def resetar_memoria(self):
        self.df = None
        self.registro = {}
        self.campos = []
        self.i = 0
        self.row_start = 0
        self.row_end = 0
        self.row_current = 0
        self.update_preview("Memória resetada. Carregue novamente.")
        self.lbl_status.configure(text="Memória limpa.")

    def _carregar_registro_da_linha(self, row_idx: int):
        if self.df is None:
            return
        self.registro = self.df.iloc[row_idx].to_dict()

    def iniciar(self):
        if self.df is None:
            messagebox.showwarning("Carregar", "Carregue o Excel primeiro.")
            return
        # volta para a linha inicial sempre que iniciar
        self.row_current = self.row_start
        self._carregar_registro_da_linha(self.row_current)
        self.i = 0
        self._pular_skips_automaticamente(+1)
        self.lbl_status.configure(text="Iniciado. F8 exclusivo.")
        self.update_preview()

    def on_enviar(self, avancar: bool):
        if self._sending:
            return
        now = time.monotonic()
        if (now - self._last_enviar_at) < self._debounce_secs:
            return
        self._last_enviar_at = now
        self._sending = True
        try:
            if self.df is None:
                return

            # se chegou ao fim das colunas, em modo multi-linha, tenta ir para próxima linha
            if self.i >= len(self.campos):
                if bool(self.chk_multilinha.get()):
                    if self.row_current < self.row_end:
                        self.row_current += 1
                        self._carregar_registro_da_linha(self.row_current)
                        self.i = 0
                        self._pular_skips_automaticamente(+1)
                    else:
                        # acabou todas as linhas
                        self.lbl_status.configure(text="Fim das linhas.")
                        self.update_preview("Fim das linhas.")
                        return
                else:
                    # comportamento antigo: simplesmente para
                    return

            if self.i >= len(self.campos):
                return

            campo = self.campos[self.i]
            if campo["skip"]:
                self._pular_skips_automaticamente(+1)
                self.update_preview("Campo ignorado.")
                return

            col = campo["nome"]
            val = to_str(self.registro.get(col, ""))
            if val:
                pg.write(val, interval=0.02)
            if bool(self.chk_tab.get()):
                pg.press("tab")
            if avancar:
                self.i += 1
                self._pular_skips_automaticamente(+1)
            self.update_preview()
        finally:
            time.sleep(0.02)
            self._sending = False

    def on_pular(self, marcar: bool = False):
        if self.df is None or self.i >= len(self.campos):
            return
        if marcar:
            self.campos[self.i]["skip"] = True
        self.i += 1
        self._pular_skips_automaticamente(+1)
        self.update_preview("Campo pulado.")

    def toggle_ignore(self):
        if self.df is None or self.i >= len(self.campos):
            return
        self.campos[self.i]["skip"] = not self.campos[self.i]["skip"]
        estado = "IGNORADO" if self.campos[self.i]["skip"] else "ATIVO"
        self.update_preview(f"Campo marcado como {estado}.")

    def on_voltar(self):
        if self.df is None:
            return
        if self.i > 0:
            self.i -= 1
            self._pular_skips_automaticamente(-1)
            self.update_preview("Voltou um campo.")
        # por simplicidade, não volta de uma linha para a anterior (ponto que dá para evoluir depois)

    # ----- utilitários -----
    def _pular_skips_automaticamente(self, sentido: int):
        n = len(self.campos)
        while 0 <= self.i < n and self.campos[self.i]["skip"]:
            self.i += sentido

    def escolher_arquivo(self):
        p = filedialog.askopenfilename(
            title="Escolha o Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xlsb")]
        )
        if not p:
            return
        self.entry_arquivo.delete(0, "end")
        self.entry_arquivo.insert(0, p)
        try:
            ext = Path(p).suffix.lower()
            if ext == ".xlsb":
                xls = pd.ExcelFile(p, engine="pyxlsb")
            else:
                xls = pd.ExcelFile(p)
            self.cmb_sheet.configure(values=xls.sheet_names)
            self.cmb_sheet.set(xls.sheet_names[0])
        except Exception as e:
            messagebox.showerror("Erro", f"Não consegui ler abas:\n{e}")

    def _read_excel_any(self, path: str, sheet: str) -> pd.DataFrame:
        ext = Path(path).suffix.lower()
        if ext in [".xlsx", ".xlsm"]:
            return pd.read_excel(
                path,
                sheet_name=sheet,
                dtype=object,
                engine="openpyxl",
                header=None
            )
        elif ext == ".xlsb":
            return pd.read_excel(
                path,
                sheet_name=sheet,
                dtype=object,
                engine="pyxlsb",
                header=None
            )
        else:
            raise ValueError(f"Extensão não suportada: {ext}")

    def carregar_dados(self):
        self.resetar_memoria()
        path = self.entry_arquivo.get().strip()
        if not path:
            return
        sheet = self.cmb_sheet.get().strip()
        if not sheet:
            return

        try:
            raw = self._read_excel_any(path, sheet)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler Excel:\n{e}")
            return

        if raw is None or raw.empty:
            messagebox.showerror("Erro", "Aba vazia.")
            return

        df = vertical_form_to_one_row(raw) if sheet_looks_like_vertical_form(raw) else raw

        row_idx = int(self.entry_linha.get().strip()) - 1
        if row_idx >= len(df):
            messagebox.showerror("Erro", f"Linha {row_idx+1} não existe na planilha.")
            return

        self.df = df
        self.row_start = row_idx
        self.row_current = row_idx
        self.row_end = len(df) - 1  # lê até a última linha existente

        self._carregar_registro_da_linha(self.row_current)

        self.campos = [{"nome": str(c), "skip": is_template_skip(str(c))} for c in df.columns]
        self.i = 0
        self._pular_skips_automaticamente(+1)
        self._render_lista()
        self.lbl_status.configure(
            text=f"Arquivo: {Path(path).name} | Aba: {sheet} | Linhas {self.row_start+1}..{self.row_end+1}"
        )
        self.update_preview("Arquivo carregado. Linha 1 lida.")

    def _render_lista(self):
        self.listbox.configure(state="normal")
        self.listbox.delete("1.0", "end")
        for idx, c in enumerate(self.campos, start=1):
            nome = c["nome"]
            tag = " [SKIP]" if c["skip"] else ""
            self.listbox.insert("end", f"{idx:03d}  {nome}{tag}\n")
        self.listbox.configure(state="disabled")

    def update_preview(self, msg=None):
        self.txt_preview.configure(state="normal")
        self.txt_preview.delete("1.0", "end")
        if msg:
            self.txt_preview.insert("end", msg + "\n\n")
        if self.df is not None and 0 <= self.i < len(self.campos):
            c = self.campos[self.i]
            val = to_str(self.registro.get(c["nome"], ""))
            skip_txt = " [SKIP]" if c["skip"] else ""
            info_linha = f"Linha atual: {self.row_current+1}/{self.row_end+1}"
            self.txt_preview.insert(
                "end",
                f"{info_linha}\nCampo atual ({self.i+1}/{len(self.campos)}): {c['nome']}{skip_txt}\nValor: {val}\n"
            )
        self.txt_preview.configure(state="disabled")
        self._render_lista()

if __name__ == "__main__":
    app = App()
    app.mainloop()
