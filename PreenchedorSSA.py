import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import pyautogui
import pyperclip
import keyboard
import threading
import time
import re
import json
import os

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class PreenchedorCDCColaF8:
    def __init__(self):
        self.app = ctk.CTk()
        self.app.title("Sonova | Preenchedor CDC por Cola + F8")
        self.app.geometry("1560x980")
        self.app.minsize(1380, 860)

        self.COR_HEADER = "#121C4E"
        self.COR_AZUL_1 = "#163A70"
        self.COR_AZUL_2 = "#245A9A"
        self.COR_AZUL_3 = "#3E7CC3"
        self.COR_AZUL_4 = "#DCE6F2"
        self.COR_AZUL_5 = "#EEF4FA"
        self.COR_SUCESSO = "#2F7D32"
        self.COR_ALERTA = "#C62828"
        self.COR_NEUTRA = "#6B7280"
        self.COR_INFO = "#406A9A"
        self.COR_CARD = "#FFFFFF"
        self.COR_FUNDO = "#F4F7FB"
        self.COR_BORDA = "#D7E0EA"
        self.COR_TEXTO = "#1F2937"
        self.COR_TEXTO_SUAVE = "#5F6B7A"
        self.COR_AMBAR = "#A97C00"

        self.df = pd.DataFrame(columns=["CDC", "DADO", "TIPO", "STATUS", "OBS"])
        self.capture_mode = None
        self.last_f8_time = 0.0

        self.pos_cdc = None
        self.pos_valor = None
        self.pos_percentual = None

        self.running = False
        self.paused = False
        self.stop_requested = False
        self.worker_thread = None

        self.batch_size_default = 20
        self.current_index_resume = None

        pyautogui.PAUSE = 0.10
        pyautogui.FAILSAFE = True

        self.criar_interface()
        self.registrar_hotkey_f8()
        self.app.protocol("WM_DELETE_WINDOW", self.fechar_app)

    # =========================================================
    # UI BASE
    # =========================================================
    # python -m cx_Freeze PreencedorSSA.py `
    #     --target-dir dist_preenchedor_ssa `
    #     --base-name Win32GUI `
    #     --packages pandas,pyautogui,customtkinter,openpyxl,pyxlsb,keyboard,pyperclip `
    #     --includes tkinter,customtkinter,pandas,pyautogui,keyboard,pyperclip
    # =========================================================
    def criar_card(self, parent):
        return ctk.CTkFrame(
            parent,
            fg_color=self.COR_CARD,
            corner_radius=12,
            border_width=1,
            border_color=self.COR_BORDA
        )

    def criar_botao(self, parent, texto, comando, cor, hover, width=140):
        return ctk.CTkButton(
            parent,
            text=texto,
            command=comando,
            fg_color=cor,
            hover_color=hover,
            text_color="white",
            width=width,
            height=38,
            corner_radius=8,
            font=("Segoe UI", 12, "bold")
        )

    def criar_entry(self, parent):
        return ctk.CTkEntry(
            parent,
            height=38,
            fg_color="white",
            border_color="#AEBFD3",
            text_color=self.COR_TEXTO,
            font=("Segoe UI", 12)
        )

    def criar_combo(self, parent, values, command=None):
        return ctk.CTkComboBox(
            parent,
            values=values,
            state="readonly",
            command=command,
            height=38,
            fg_color="white",
            border_color="#AEBFD3",
            button_color="#DCE6F2",
            button_hover_color="#C8D8EA",
            text_color=self.COR_TEXTO,
            dropdown_fg_color="white",
            dropdown_text_color=self.COR_TEXTO,
            font=("Segoe UI", 12)
        )

    def criar_interface(self):
        self.app.configure(fg_color=self.COR_FUNDO)

        header = ctk.CTkFrame(
            self.app,
            fg_color=self.COR_HEADER,
            corner_radius=0,
            height=78
        )
        header.pack(fill="x")

        ctk.CTkLabel(
            header,
            text="Preenchedor CDC por Cola + F8",
            font=("Segoe UI", 28, "bold"),
            text_color="white"
        ).pack(side="left", padx=20, pady=18)

        ctk.CTkLabel(
            header,
            text="Sonova | Cole CDC e Valor ou Percentual, capture o campo com F8 e preencha automaticamente",
            font=("Segoe UI", 13),
            text_color="#D9E7F5"
        ).pack(side="left", padx=10, pady=22)

        body = ctk.CTkFrame(self.app, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=12, pady=12)

        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(1, weight=1)

        self.criar_bloco_entrada(body)
        self.criar_bloco_config(body)
        self.criar_bloco_preview(body)
        self.criar_bloco_execucao(body)

        self.atualizar_labels_posicoes()
        self.atualizar_resumo()

    # =========================================================
    # BLOCO ENTRADA
    # =========================================================
    def criar_bloco_entrada(self, parent):
        frame = self.criar_card(parent)
        frame.grid(row=0, column=0, padx=(0, 6), pady=(0, 8), sticky="nsew")

        ctk.CTkLabel(
            frame,
            text="1. Entrada Rápida",
            font=("Segoe UI", 18, "bold"),
            text_color=self.COR_TEXTO
        ).pack(anchor="w", padx=16, pady=(14, 8))

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.pack(fill="x", padx=16, pady=(0, 8))

        self.criar_botao(
            botoes, "Ler Dados Colados", self.ler_dados_colados,
            self.COR_HEADER, "#1B2C63", 165
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            botoes, "Modelo Valor", self.inserir_modelo_valor,
            self.COR_AZUL_2, "#1E4F8A", 130
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            botoes, "Modelo Percentual", self.inserir_modelo_percentual,
            self.COR_AZUL_3, "#2D68A8", 160
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            botoes, "Limpar Texto", self.limpar_texto,
            self.COR_NEUTRA, "#5D6671", 120
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            botoes, "Limpar Base", self.limpar_base,
            "#A83E4A", "#933540", 120
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            botoes, "Validar Base", self.validar_base_visual,
            self.COR_INFO, "#355D87", 125
        ).pack(side="left")

        ctk.CTkLabel(
            frame,
            text="Aceita TAB, ponto e vírgula ou múltiplos espaços. Pode colar com ou sem cabeçalho.",
            font=("Segoe UI", 12),
            text_color=self.COR_TEXTO_SUAVE
        ).pack(anchor="w", padx=16, pady=(0, 6))

        texto_frame = ctk.CTkFrame(frame, fg_color="transparent")
        texto_frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.txt_entrada = ctk.CTkTextbox(
            texto_frame,
            height=260,
            font=("Consolas", 12),
            fg_color="#FBFCFE",
            text_color=self.COR_TEXTO,
            border_width=1,
            border_color=self.COR_BORDA
        )
        self.txt_entrada.pack(side="left", fill="both", expand=True)

        self.scroll_txt = ctk.CTkScrollbar(texto_frame, command=self.txt_entrada.yview)
        self.scroll_txt.pack(side="right", fill="y")
        self.txt_entrada.configure(yscrollcommand=self.scroll_txt.set)

    # =========================================================
    # BLOCO CONFIG
    # =========================================================
    def criar_bloco_config(self, parent):
        frame = self.criar_card(parent)
        frame.grid(row=0, column=1, padx=(6, 0), pady=(0, 8), sticky="nsew")

        ctk.CTkLabel(
            frame,
            text="2. Configuração",
            font=("Segoe UI", 18, "bold"),
            text_color=self.COR_TEXTO
        ).pack(anchor="w", padx=16, pady=(14, 8))

        grid = ctk.CTkFrame(frame, fg_color="transparent")
        grid.pack(fill="x", padx=16, pady=(0, 10))
        for i in range(5):
            grid.grid_columnconfigure(i, weight=1)

        labels_1 = ["Usar Campo", "Modo", "Delay entre ações", "Delay inicial", "Lote"]
        for i, txt in enumerate(labels_1):
            ctk.CTkLabel(
                grid, text=txt, font=("Segoe UI", 12, "bold"), text_color=self.COR_TEXTO
            ).grid(row=0, column=i, sticky="w", padx=6, pady=(0, 4))

        self.cmb_usar = self.criar_combo(
            grid, ["VALOR", "PERCENTUAL"],
            command=lambda _=None: self.atualizar_labels_posicoes()
        )
        self.cmb_usar.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 8))
        self.cmb_usar.set("VALOR")

        self.cmb_modo = self.criar_combo(
            grid, ["TAB", "POSICOES"],
            command=lambda _=None: self.atualizar_labels_posicoes()
        )
        self.cmb_modo.grid(row=1, column=1, sticky="ew", padx=6, pady=(0, 8))
        self.cmb_modo.set("TAB")

        self.entry_delay = self.criar_entry(grid)
        self.entry_delay.grid(row=1, column=2, sticky="ew", padx=6, pady=(0, 8))
        self.entry_delay.insert(0, "0,15")

        self.entry_delay_inicial = self.criar_entry(grid)
        self.entry_delay_inicial.grid(row=1, column=3, sticky="ew", padx=6, pady=(0, 8))
        self.entry_delay_inicial.insert(0, "3")

        self.entry_batch = self.criar_entry(grid)
        self.entry_batch.grid(row=1, column=4, sticky="ew", padx=6, pady=(0, 8))
        self.entry_batch.insert(0, "20")

        labels_2 = ["Qtde TAB após CDC", "Qtde TAB após Valor", "Quantidade a executar", "Ação final", "Regra sem CDC"]
        for i, txt in enumerate(labels_2):
            ctk.CTkLabel(
                grid, text=txt, font=("Segoe UI", 12, "bold"), text_color=self.COR_TEXTO
            ).grid(row=2, column=i, sticky="w", padx=6, pady=(8, 4))

        self.entry_tabs_cdc = self.criar_entry(grid)
        self.entry_tabs_cdc.grid(row=3, column=0, sticky="ew", padx=6, pady=(0, 8))
        self.entry_tabs_cdc.insert(0, "3")

        self.entry_tabs_final = self.criar_entry(grid)
        self.entry_tabs_final.grid(row=3, column=1, sticky="ew", padx=6, pady=(0, 8))
        self.entry_tabs_final.insert(0, "2")

        self.entry_qtd = self.criar_entry(grid)
        self.entry_qtd.grid(row=3, column=2, sticky="ew", padx=6, pady=(0, 8))
        self.entry_qtd.insert(0, "TODOS")

        acao_frame = ctk.CTkFrame(grid, fg_color="transparent")
        acao_frame.grid(row=3, column=3, sticky="w", padx=6, pady=(0, 8))

        self.chk_enter_var = ctk.BooleanVar(value=True)
        self.chk_enter = ctk.CTkCheckBox(
            acao_frame,
            text="Pressionar Enter",
            variable=self.chk_enter_var,
            text_color=self.COR_TEXTO,
            font=("Segoe UI", 12, "bold"),
            checkbox_width=22,
            checkbox_height=22,
            border_width=2
        )
        self.chk_enter.pack(anchor="w")

        self.cmb_sem_cdc = self.criar_combo(grid, ["IGNORAR", "PARAR", "ERRO"])
        self.cmb_sem_cdc.grid(row=3, column=4, sticky="ew", padx=6, pady=(0, 8))
        self.cmb_sem_cdc.set("IGNORAR")

        linha2 = ctk.CTkFrame(frame, fg_color="transparent")
        linha2.pack(fill="x", padx=16, pady=(0, 10))

        self.criar_botao(
            linha2, "Salvar Layout", self.salvar_layout_json,
            self.COR_AZUL_2, "#1E4F8A", 130
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            linha2, "Carregar Layout", self.carregar_layout_json,
            self.COR_AZUL_3, "#2D68A8", 145
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            linha2, "Exportar Log", self.exportar_log_txt,
            self.COR_INFO, "#355D87", 130
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            linha2, "Ignorar Selecionada", self.marcar_linha_ignorada,
            self.COR_NEUTRA, "#59616D", 165
        ).pack(side="left")

        box = ctk.CTkFrame(
            frame,
            fg_color=self.COR_AZUL_5,
            corner_radius=10,
            border_width=1,
            border_color=self.COR_BORDA
        )
        box.pack(fill="x", padx=16, pady=(0, 10))
        box.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkLabel(
            box,
            text="Captura por F8",
            font=("Segoe UI", 14, "bold"),
            text_color=self.COR_TEXTO
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 6))

        self.lbl_instrucao_f8 = ctk.CTkLabel(
            box,
            text="Clique no botão do campo e depois pressione F8 sobre o sistema alvo.",
            font=("Segoe UI", 12),
            text_color=self.COR_TEXTO_SUAVE,
            justify="left"
        )
        self.lbl_instrucao_f8.grid(row=1, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 8))

        self.lbl_pos_cdc = ctk.CTkLabel(box, text="CDC: não capturado", anchor="w", text_color=self.COR_TEXTO)
        self.lbl_pos_cdc.grid(row=2, column=0, sticky="ew", padx=12, pady=4)

        self.lbl_pos_valor = ctk.CTkLabel(box, text="Valor: não capturado", anchor="w", text_color=self.COR_TEXTO)
        self.lbl_pos_valor.grid(row=2, column=1, sticky="ew", padx=12, pady=4)

        self.lbl_pos_percentual = ctk.CTkLabel(box, text="Percentual: não capturado", anchor="w", text_color=self.COR_TEXTO)
        self.lbl_pos_percentual.grid(row=2, column=2, sticky="ew", padx=12, pady=4)

        self.criar_botao(
            box, "Armar F8 para CDC",
            lambda: self.armar_captura("CDC"),
            self.COR_HEADER, "#1B2C63", 170
        ).grid(row=3, column=0, sticky="ew", padx=12, pady=(8, 12))

        self.criar_botao(
            box, "Armar F8 para Valor",
            lambda: self.armar_captura("VALOR"),
            self.COR_AZUL_2, "#1E4F8A", 170
        ).grid(row=3, column=1, sticky="ew", padx=12, pady=(8, 12))

        self.criar_botao(
            box, "Armar F8 para Percentual",
            lambda: self.armar_captura("PERCENTUAL"),
            self.COR_AZUL_3, "#2D68A8", 190
        ).grid(row=3, column=2, sticky="ew", padx=12, pady=(8, 12))

    # =========================================================
    # BLOCO PREVIEW
    # =========================================================
    def criar_bloco_preview(self, parent):
        frame = self.criar_card(parent)
        frame.grid(row=1, column=0, padx=(0, 6), pady=(0, 0), sticky="nsew")

        ctk.CTkLabel(
            frame,
            text="3. Preview da Base",
            font=("Segoe UI", 18, "bold"),
            text_color=self.COR_TEXTO
        ).pack(anchor="w", padx=16, pady=(14, 8))

        resumo = ctk.CTkFrame(frame, fg_color="transparent")
        resumo.pack(fill="x", padx=16, pady=(0, 8))

        self.lbl_resumo = ctk.CTkLabel(
            resumo,
            text="Total: 0 | Pendentes: 0 | OK: 0 | Erro: 0 | Ignorados: 0",
            font=("Segoe UI", 12, "bold"),
            text_color=self.COR_TEXTO_SUAVE
        )
        self.lbl_resumo.pack(anchor="w")

        tree_frame = ctk.CTkFrame(frame, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.tree = ttk.Treeview(
            tree_frame,
            columns=("CDC", "DADO", "TIPO", "STATUS", "OBS"),
            show="headings"
        )

        self.tree.heading("CDC", text="CDC")
        self.tree.heading("DADO", text="Dado")
        self.tree.heading("TIPO", text="Tipo")
        self.tree.heading("STATUS", text="Status")
        self.tree.heading("OBS", text="Observação")

        self.tree.column("CDC", width=130, anchor="center")
        self.tree.column("DADO", width=150, anchor="e")
        self.tree.column("TIPO", width=100, anchor="center")
        self.tree.column("STATUS", width=120, anchor="center")
        self.tree.column("OBS", width=320, anchor="w")

        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            background="white",
            foreground="black",
            rowheight=30,
            fieldbackground="white",
            font=("Segoe UI", 11)
        )
        style.configure(
            "Treeview.Heading",
            background=self.COR_AZUL_4,
            foreground="black",
            font=("Segoe UI", 11, "bold")
        )
        style.map("Treeview", background=[("selected", "#D9E7F5")], foreground=[("selected", "black")])

        self.tree.grid(row=0, column=0, sticky="nsew")

        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")

        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

    # =========================================================
    # BLOCO EXECUÇÃO
    # =========================================================
    def criar_bloco_execucao(self, parent):
        frame = self.criar_card(parent)
        frame.grid(row=1, column=1, padx=(6, 0), pady=(0, 0), sticky="nsew")

        ctk.CTkLabel(
            frame,
            text="4. Execução",
            font=("Segoe UI", 18, "bold"),
            text_color=self.COR_TEXTO
        ).pack(anchor="w", padx=16, pady=(14, 8))

        txt = (
            "Fluxo operacional:\n"
            "1. Cole os dados.\n"
            "2. Leia e valide a base.\n"
            "3. Escolha Valor ou Percentual.\n"
            "4. Arme F8 para os campos.\n"
            "5. Vá ao sistema e pressione F8.\n"
            "6. Teste 1 linha.\n"
            "7. Execute tudo em lotes controlados."
        )
        ctk.CTkLabel(
            frame,
            text=txt,
            justify="left",
            font=("Segoe UI", 12),
            text_color=self.COR_TEXTO_SUAVE
        ).pack(anchor="w", padx=16, pady=(0, 10))

        botoes1 = ctk.CTkFrame(frame, fg_color="transparent")
        botoes1.pack(fill="x", padx=16, pady=(0, 8))

        self.criar_botao(
            botoes1, "Teste 1 Linha",
            lambda: self.iniciar_execucao(teste=True),
            self.COR_AZUL_2, "#1E4F8A", 130
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            botoes1, "Próxima Linha",
            self.executar_proxima_linha_manual,
            self.COR_AZUL_3, "#2D68A8", 130
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            botoes1, "Executar Tudo",
            lambda: self.iniciar_execucao(teste=False),
            self.COR_SUCESSO, "#286D2B", 135
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            botoes1, "Voltar Linha",
            self.voltar_linha,
            self.COR_NEUTRA, "#5D6671", 125
        ).pack(side="left", padx=(0, 8))

        botoes2 = ctk.CTkFrame(frame, fg_color="transparent")
        botoes2.pack(fill="x", padx=16, pady=(0, 10))

        self.criar_botao(
            botoes2, "Play",
            self.retomar_execucao,
            self.COR_SUCESSO, "#286D2B", 100
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            botoes2, "Pause",
            self.pausar_execucao,
            self.COR_AMBAR, "#916B00", 100
        ).pack(side="left", padx=(0, 8))

        self.criar_botao(
            botoes2, "Stop",
            self.parar_execucao,
            self.COR_ALERTA, "#AE2323", 100
        ).pack(side="left")

        self.lbl_status = ctk.CTkLabel(
            frame,
            text="Status: aguardando dados.",
            font=("Segoe UI", 12, "bold"),
            text_color=self.COR_HEADER
        )
        self.lbl_status.pack(anchor="w", padx=16, pady=(0, 8))

        self.lbl_linha = ctk.CTkLabel(
            frame,
            text="Linha atual: 0",
            font=("Segoe UI", 12),
            text_color=self.COR_TEXTO_SUAVE
        )
        self.lbl_linha.pack(anchor="w", padx=16, pady=(0, 4))

        self.lbl_progresso = ctk.CTkLabel(
            frame,
            text="Progresso: 0/0",
            font=("Segoe UI", 12, "bold"),
            text_color=self.COR_INFO
        )
        self.lbl_progresso.pack(anchor="w", padx=16, pady=(0, 8))

        ctk.CTkLabel(
            frame,
            text="Log operacional",
            font=("Segoe UI", 13, "bold"),
            text_color=self.COR_TEXTO
        ).pack(anchor="w", padx=16, pady=(0, 6))

        log_frame = ctk.CTkFrame(frame, fg_color="transparent")
        log_frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.txt_log = ctk.CTkTextbox(
            log_frame,
            height=440,
            font=("Consolas", 11),
            fg_color="#FBFCFE",
            text_color=self.COR_TEXTO,
            border_width=1,
            border_color=self.COR_BORDA
        )
        self.txt_log.pack(side="left", fill="both", expand=True)

        self.scroll_log = ctk.CTkScrollbar(log_frame, command=self.txt_log.yview)
        self.scroll_log.pack(side="right", fill="y")
        self.txt_log.configure(yscrollcommand=self.scroll_log.set)

    # =========================================================
    # LOG E STATUS
    # =========================================================
    def log(self, msg):
        timestamp = time.strftime("%H:%M:%S")
        linha = f"[{timestamp}] {msg}\n"
        if hasattr(self, "txt_log"):
            self.txt_log.insert("end", linha)
            self.txt_log.see("end")
            self.app.update_idletasks()
        else:
            print(linha)

    def set_status(self, msg, cor=None):
        self.lbl_status.configure(text=f"Status: {msg}", text_color=cor or self.COR_HEADER)
        self.app.update_idletasks()

    # =========================================================
    # HOTKEY F8
    # =========================================================
    def registrar_hotkey_f8(self):
        try:
            keyboard.add_hotkey("F8", self.on_f8_pressed, suppress=False)
            self.log("Hotkey F8 registrada com sucesso.")
        except Exception as e:
            self.log(f"Falha ao registrar F8: {e}")

    def on_f8_pressed(self):
        agora = time.time()
        if agora - self.last_f8_time < 0.35:
            return
        self.last_f8_time = agora

        if self.capture_mode is None:
            return

        try:
            pos = pyautogui.position()

            if self.capture_mode == "CDC":
                self.pos_cdc = pos
            elif self.capture_mode == "VALOR":
                self.pos_valor = pos
            elif self.capture_mode == "PERCENTUAL":
                self.pos_percentual = pos

            self.log(f"Posição capturada para {self.capture_mode}: x={pos.x}, y={pos.y}")
            self.set_status(f"posição {self.capture_mode} capturada", self.COR_SUCESSO)
            self.capture_mode = None
            self.lbl_instrucao_f8.configure(text="Captura concluída. Pode armar outro campo se precisar.")
            self.atualizar_labels_posicoes()
        except Exception as e:
            self.log(f"Erro ao capturar posição com F8: {e}")

    def armar_captura(self, campo):
        self.capture_mode = campo
        self.lbl_instrucao_f8.configure(
            text=f"Captura armada para {campo}. Vá até o sistema alvo e pressione F8 sobre o campo."
        )
        self.log(f"Captura armada para {campo}.")
        self.set_status(f"captura armada para {campo}", self.COR_INFO)

    def atualizar_labels_posicoes(self):
        self.lbl_pos_cdc.configure(text=f"CDC: {self.formatar_posicao(self.pos_cdc)}")
        self.lbl_pos_valor.configure(text=f"Valor: {self.formatar_posicao(self.pos_valor)}")
        self.lbl_pos_percentual.configure(text=f"Percentual: {self.formatar_posicao(self.pos_percentual)}")

        usar = self.cmb_usar.get()
        modo = self.cmb_modo.get()

        if usar == "VALOR":
            self.lbl_pos_percentual.configure(text="Percentual: ignorado neste modo")
        else:
            self.lbl_pos_valor.configure(text="Valor: ignorado neste modo")

        if modo == "TAB":
            self.log("Modo TAB ativo: o sistema usa TABs configuráveis após CDC e após Valor.")
        else:
            self.log("Modo POSICOES ativo: capture CDC e também o campo de Valor ou Percentual.")

    def formatar_posicao(self, pos):
        if pos is None:
            return "não capturado"
        return f"x={pos.x} | y={pos.y}"

    # =========================================================
    # ENTRADA / PARSER
    # =========================================================
    def inserir_modelo_valor(self):
        self.txt_entrada.delete("1.0", "end")
        self.txt_entrada.insert(
            "1.0",
            "CDC\tVALOR\n300000\t843,60\n301000\t3532,00\n303000\t2000,00"
        )
        self.cmb_usar.set("VALOR")

    def inserir_modelo_percentual(self):
        self.txt_entrada.delete("1.0", "end")
        self.txt_entrada.insert(
            "1.0",
            "CDC\tPERCENTUAL\n300000\t25\n301000\t35\n303000\t40"
        )
        self.cmb_usar.set("PERCENTUAL")

    def limpar_texto(self):
        self.txt_entrada.delete("1.0", "end")

    def limpar_base(self):
        self.df = pd.DataFrame(columns=["CDC", "DADO", "TIPO", "STATUS", "OBS"])
        self.renderizar_base()
        self.atualizar_resumo()
        self.atualizar_linha_atual()
        self.log("Base limpa.")
        self.set_status("base limpa")

    def detectar_cabecalho(self, primeira_linha):
        p1 = str(primeira_linha[0]).strip().lower()
        p2 = str(primeira_linha[1]).strip().lower()
        palavras = ["cdc", "valor", "percentual", "percent", "%", "dado"]
        return any(p in p1 for p in palavras) or any(p in p2 for p in palavras)

    def separar_linha(self, linha):
        linha = linha.strip()
        if not linha:
            return None

        if "\t" in linha:
            partes = [p.strip() for p in linha.split("\t") if p.strip() != ""]
            if len(partes) >= 2:
                return partes[0], partes[1]

        if ";" in linha:
            partes = [p.strip() for p in linha.split(";") if p.strip() != ""]
            if len(partes) >= 2:
                return partes[0], partes[1]

        partes = re.split(r"\s{2,}", linha)
        partes = [p.strip() for p in partes if p.strip()]
        if len(partes) >= 2:
            return partes[0], partes[1]

        return None

    def ler_dados_colados(self):
        texto = self.txt_entrada.get("1.0", "end").strip()
        if not texto:
            messagebox.showwarning("Atenção", "Cole os dados antes de ler.")
            return

        linhas = [l for l in texto.splitlines() if l.strip()]
        if not linhas:
            messagebox.showwarning("Atenção", "Não há linhas válidas para leitura.")
            return

        registros = []
        primeira_parse = self.separar_linha(linhas[0])

        inicio = 0
        if primeira_parse and self.detectar_cabecalho(primeira_parse):
            inicio = 1

        for i in range(inicio, len(linhas)):
            resultado = self.separar_linha(linhas[i])
            if not resultado:
                self.log(f"Linha ignorada por formato inválido: {linhas[i]}")
                continue

            cdc, dado = resultado
            registros.append({
                "CDC": str(cdc).strip(),
                "DADO": str(dado).strip(),
                "TIPO": self.cmb_usar.get(),
                "STATUS": "PENDENTE",
                "OBS": ""
            })

        if not registros:
            messagebox.showwarning("Atenção", "Nenhum dado válido foi identificado.")
            return

        self.df = pd.DataFrame(registros)
        self.pre_validar_dataframe()
        self.renderizar_base()
        self.atualizar_resumo()
        self.atualizar_linha_atual()
        self.log(f"{len(self.df)} linha(s) carregada(s) por cola.")
        self.set_status(f"{len(self.df)} linha(s) carregada(s)", self.COR_SUCESSO)

    # =========================================================
    # PRÉ-VALIDAÇÃO
    # =========================================================
    def pre_validar_dataframe(self):
        if self.df.empty:
            return

        for idx, row in self.df.iterrows():
            cdc = str(row["CDC"]).strip()
            dado = str(row["DADO"]).strip()

            if not cdc:
                self.df.at[idx, "STATUS"] = "IGNORADO"
                self.df.at[idx, "OBS"] = "CDC vazio"
                continue

            if not dado:
                self.df.at[idx, "STATUS"] = "IGNORADO"
                self.df.at[idx, "OBS"] = "Dado vazio"
                continue

            self.df.at[idx, "STATUS"] = "PENDENTE"
            self.df.at[idx, "OBS"] = ""

    def validar_base_visual(self):
        if self.df.empty:
            messagebox.showwarning("Validação", "Não há base carregada.")
            return
        self.pre_validar_dataframe()
        self.renderizar_base()
        self.atualizar_resumo()
        self.atualizar_linha_atual()
        self.log("Base validada visualmente.")
        self.set_status("base validada", self.COR_INFO)

    # =========================================================
    # PREVIEW / RESUMO
    # =========================================================
    def renderizar_base(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if self.df.empty:
            return

        for idx, row in self.df.iterrows():
            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(row["CDC"], row["DADO"], row["TIPO"], row["STATUS"], row["OBS"])
            )

    def atualizar_status_linha(self, idx, status, obs=None):
        if idx < 0 or idx >= len(self.df):
            return

        self.df.at[idx, "STATUS"] = status
        if obs is not None:
            self.df.at[idx, "OBS"] = obs

        self.tree.item(
            str(idx),
            values=(
                self.df.at[idx, "CDC"],
                self.df.at[idx, "DADO"],
                self.df.at[idx, "TIPO"],
                self.df.at[idx, "STATUS"],
                self.df.at[idx, "OBS"]
            )
        )
        self.atualizar_resumo()

    def atualizar_resumo(self):
        if self.df.empty:
            self.lbl_resumo.configure(text="Total: 0 | Pendentes: 0 | OK: 0 | Erro: 0 | Ignorados: 0")
            self.lbl_progresso.configure(text="Progresso: 0/0")
            return

        total = len(self.df)
        pend = int((self.df["STATUS"] == "PENDENTE").sum())
        ok = int((self.df["STATUS"] == "OK").sum())
        erro = int((self.df["STATUS"] == "ERRO").sum())
        ign = int((self.df["STATUS"] == "IGNORADO").sum())

        self.lbl_resumo.configure(
            text=f"Total: {total} | Pendentes: {pend} | OK: {ok} | Erro: {erro} | Ignorados: {ign}"
        )
        self.lbl_progresso.configure(text=f"Progresso: {ok + ign}/{total}")

    def obter_proxima_linha_pendente(self):
        if self.df.empty:
            return None
        pendentes = self.df.index[self.df["STATUS"].isin(["PENDENTE", "ERRO", "PAUSADO"])]
        if len(pendentes) == 0:
            return None
        return int(pendentes[0])

    def atualizar_linha_atual(self):
        idx = self.obter_proxima_linha_pendente()
        if idx is None:
            self.lbl_linha.configure(text="Linha atual: concluído")
        else:
            self.lbl_linha.configure(text=f"Linha atual: {idx + 1} de {len(self.df)}")

    # =========================================================
    # PARÂMETROS
    # =========================================================
    def obter_delay(self):
        try:
            return float(str(self.entry_delay.get()).replace(",", "."))
        except Exception:
            return 0.15

    def obter_delay_inicial(self):
        try:
            return int(float(str(self.entry_delay_inicial.get()).replace(",", ".")))
        except Exception:
            return 3

    def obter_qtd(self):
        txt = str(self.entry_qtd.get()).strip().upper()
        if txt == "" or txt == "TODOS":
            return len(self.df)
        try:
            return max(1, min(int(txt), len(self.df)))
        except Exception:
            return len(self.df)

    def obter_tabs_cdc(self):
        try:
            return max(0, int(float(str(self.entry_tabs_cdc.get()).replace(",", "."))))
        except Exception:
            return 3

    def obter_tabs_final(self):
        try:
            return max(0, int(float(str(self.entry_tabs_final.get()).replace(",", "."))))
        except Exception:
            return 2

    def obter_batch(self):
        try:
            return max(1, int(float(str(self.entry_batch.get()).replace(",", "."))))
        except Exception:
            return self.batch_size_default

    # =========================================================
    # EXECUÇÃO
    # =========================================================
    def validar_execucao(self):
        if self.df.empty:
            messagebox.showwarning("Validação", "Leia os dados antes de executar.")
            return False

        if self.pos_cdc is None:
            messagebox.showwarning("Validação", "Capture a posição do campo CDC.")
            return False

        usar = self.cmb_usar.get()
        modo = self.cmb_modo.get()

        if modo == "POSICOES":
            if usar == "VALOR" and self.pos_valor is None:
                messagebox.showwarning("Validação", "Capture a posição do campo Valor.")
                return False
            if usar == "PERCENTUAL" and self.pos_percentual is None:
                messagebox.showwarning("Validação", "Capture a posição do campo Percentual.")
                return False

        return True

    def escrever_texto(self, texto):
        pyperclip.copy(str(texto))
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.05)
        pyautogui.hotkey("ctrl", "v")

    def executar_tabs(self, quantidade, delay):
        for _ in range(quantidade):
            pyautogui.press("tab")
            time.sleep(delay)

    def tratar_linha_sem_cdc(self, idx):
        regra = self.cmb_sem_cdc.get()
        if regra == "IGNORAR":
            self.atualizar_status_linha(idx, "IGNORADO", "CDC vazio - ignorado por regra")
            self.log(f"Linha {idx + 1}: CDC vazio, ignorada.")
            return "CONTINUAR"
        if regra == "PARAR":
            self.atualizar_status_linha(idx, "ERRO", "CDC vazio - execução parada por regra")
            self.log(f"Linha {idx + 1}: CDC vazio, execução interrompida.")
            return "PARAR"
        self.atualizar_status_linha(idx, "ERRO", "CDC vazio")
        self.log(f"Linha {idx + 1}: CDC vazio, marcada como erro.")
        return "CONTINUAR"

    def executar_linha(self, idx):
        if idx is None or idx >= len(self.df):
            return False

        try:
            cdc = str(self.df.at[idx, "CDC"]).strip()
            dado = str(self.df.at[idx, "DADO"]).strip()
            usar = self.cmb_usar.get()
            modo = self.cmb_modo.get()
            delay = self.obter_delay()
            tabs_cdc = self.obter_tabs_cdc()
            tabs_final = self.obter_tabs_final()

            if not cdc:
                acao = self.tratar_linha_sem_cdc(idx)
                self.atualizar_linha_atual()
                return acao != "PARAR"

            if not dado:
                self.atualizar_status_linha(idx, "IGNORADO", "Dado vazio")
                self.log(f"Linha {idx + 1}: dado vazio, ignorada.")
                self.atualizar_linha_atual()
                return True

            self.atualizar_status_linha(idx, "PROCESSANDO", "")
            self.log(
                f"Linha {idx + 1}: CDC={cdc} | {usar}={dado} | modo={modo} | "
                f"tabs_cdc={tabs_cdc} | tabs_final={tabs_final}"
            )

            pyautogui.click(self.pos_cdc.x, self.pos_cdc.y)
            time.sleep(delay)
            self.escrever_texto(cdc)
            time.sleep(delay)

            if modo == "TAB":
                self.executar_tabs(tabs_cdc, delay)
                self.escrever_texto(dado)
                time.sleep(delay)

                if tabs_final > 0:
                    self.executar_tabs(tabs_final, delay)

            else:
                if usar == "VALOR":
                    pyautogui.click(self.pos_valor.x, self.pos_valor.y)
                else:
                    pyautogui.click(self.pos_percentual.x, self.pos_percentual.y)

                time.sleep(delay)
                self.escrever_texto(dado)
                time.sleep(delay)

                if tabs_final > 0:
                    self.executar_tabs(tabs_final, delay)

            if self.chk_enter_var.get():
                pyautogui.press("enter")
                time.sleep(delay)

            self.atualizar_status_linha(idx, "OK", "Preenchido com sucesso")
            self.log(f"Linha {idx + 1}: preenchida com sucesso.")
            self.atualizar_linha_atual()
            return True

        except Exception as e:
            self.atualizar_status_linha(idx, "ERRO", str(e))
            self.log(f"Linha {idx + 1}: erro ao preencher - {e}")
            self.atualizar_linha_atual()
            return False

    def executar_proxima_linha_manual(self):
        if not self.validar_execucao():
            return

        idx = self.obter_proxima_linha_pendente()
        if idx is None:
            messagebox.showinfo("Execução", "Não há linhas pendentes.")
            return

        ok = self.executar_linha(idx)
        if ok:
            self.set_status(f"linha {idx + 1} processada", self.COR_SUCESSO)
        else:
            self.set_status(f"erro na linha {idx + 1}", self.COR_ALERTA)

    def iniciar_execucao(self, teste=False):
        if not self.validar_execucao():
            return

        if self.running:
            messagebox.showwarning("Execução", "Já existe uma execução em andamento.")
            return

        self.running = True
        self.paused = False
        self.stop_requested = False
        self.worker_thread = threading.Thread(target=self.executar_lote, args=(teste,), daemon=True)
        self.worker_thread.start()

    def executar_lote(self, teste=False):
        try:
            delay_inicial = self.obter_delay_inicial()
            limite_total = 1 if teste else self.obter_qtd()
            batch = 1 if teste else self.obter_batch()

            self.set_status(f"início em {delay_inicial} segundo(s)", self.COR_INFO)
            self.log(f"Execução preparada. teste={teste} | limite={limite_total} | lote={batch}")

            for i in range(delay_inicial, 0, -1):
                if self.stop_requested:
                    self.running = False
                    return
                self.log(f"Iniciando em {i}...")
                time.sleep(1)

            executadas = 0

            while self.running and not self.stop_requested and executadas < limite_total:
                if self.paused:
                    self.set_status("execução pausada", self.COR_AMBAR)
                    time.sleep(0.3)
                    continue

                lote_atual = 0
                self.log(f"Iniciando lote de até {batch} linha(s).")

                while lote_atual < batch and executadas < limite_total and self.running and not self.stop_requested:
                    if self.paused:
                        break

                    idx = self.obter_proxima_linha_pendente()
                    if idx is None:
                        break

                    self.current_index_resume = idx
                    ok = self.executar_linha(idx)
                    executadas += 1
                    lote_atual += 1

                    if teste and not ok:
                        break

                if self.stop_requested:
                    break

                if self.paused:
                    continue

                if self.obter_proxima_linha_pendente() is None:
                    break

                if not teste and lote_atual > 0:
                    self.log("Lote concluído. Pequena pausa de segurança antes do próximo bloco.")
                    time.sleep(0.8)

            self.running = False
            self.atualizar_linha_atual()

            if self.stop_requested:
                self.set_status("execução interrompida por stop", self.COR_ALERTA)
                self.log("Execução encerrada por stop.")
            elif self.paused:
                self.set_status("execução pausada", self.COR_AMBAR)
                self.log("Execução pausada.")
            elif self.obter_proxima_linha_pendente() is None:
                self.set_status("execução finalizada, sem pendências", self.COR_SUCESSO)
                self.log("Execução finalizada com sucesso.")
            else:
                self.set_status("execução finalizada parcialmente", self.COR_INFO)
                self.log("Execução finalizada parcialmente.")

        except Exception as e:
            self.running = False
            self.set_status("erro durante execução", self.COR_ALERTA)
            self.log(f"Erro geral na execução: {e}")
            messagebox.showerror("Erro", f"Ocorreu um erro durante a execução.\n\nDetalhe: {e}")

    def pausar_execucao(self):
        if not self.running:
            self.set_status("nada em execução para pausar", self.COR_AMBAR)
            return
        self.paused = True
        self.log("Pausa solicitada pelo usuário.")
        self.set_status("pausa solicitada", self.COR_AMBAR)

    def retomar_execucao(self):
        if self.running and self.paused:
            self.paused = False
            self.log("Execução retomada.")
            self.set_status("execução retomada", self.COR_SUCESSO)
            return

        if self.running and not self.paused:
            self.set_status("execução já está ativa", self.COR_INFO)
            return

        idx = self.obter_proxima_linha_pendente()
        if idx is None:
            self.set_status("não há pendências para retomar", self.COR_INFO)
            return

        self.log("Retomando execução a partir das pendências.")
        self.iniciar_execucao(teste=False)

    def parar_execucao(self):
        self.stop_requested = True
        self.running = False
        self.paused = False
        self.set_status("stop solicitado", self.COR_ALERTA)
        self.log("Stop solicitado pelo usuário.")

    def voltar_linha(self):
        if self.df.empty:
            return

        linhas_ok = self.df.index[self.df["STATUS"] == "OK"].tolist()
        if not linhas_ok:
            messagebox.showinfo("Voltar", "Não há linha concluída para retornar.")
            return

        idx = linhas_ok[-1]
        self.atualizar_status_linha(idx, "PENDENTE", "Retornada manualmente")
        self.atualizar_linha_atual()
        self.log(f"Linha {idx + 1} retornada para PENDENTE.")
        self.set_status(f"linha {idx + 1} voltou para pendente", self.COR_INFO)

    # =========================================================
    # UTILITÁRIOS
    # =========================================================
    def marcar_linha_ignorada(self):
        selecionado = self.tree.selection()
        if not selecionado:
            messagebox.showwarning("Ignorar", "Selecione uma linha na grade.")
            return

        idx = int(selecionado[0])
        self.atualizar_status_linha(idx, "IGNORADO", "Ignorada manualmente")
        self.atualizar_linha_atual()
        self.log(f"Linha {idx + 1} marcada como ignorada manualmente.")
        self.set_status(f"linha {idx + 1} ignorada", self.COR_INFO)

    def exportar_log_txt(self):
        try:
            conteudo = self.txt_log.get("1.0", "end").strip()
            if not conteudo:
                messagebox.showwarning("Exportar", "O log está vazio.")
                return

            caminho = filedialog.asksaveasfilename(
                title="Salvar log",
                defaultextension=".txt",
                filetypes=[("Arquivo texto", "*.txt")],
                initialfile="log_preenchedor_cdc.txt"
            )
            if not caminho:
                return

            with open(caminho, "w", encoding="utf-8") as f:
                f.write(conteudo)

            self.log(f"Log exportado para: {caminho}")
            self.set_status("log exportado", self.COR_SUCESSO)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível exportar o log.\n\nDetalhe: {e}")

    def salvar_layout_json(self):
        try:
            dados = {
                "usar_campo": self.cmb_usar.get(),
                "modo": self.cmb_modo.get(),
                "delay": self.entry_delay.get(),
                "delay_inicial": self.entry_delay_inicial.get(),
                "batch": self.entry_batch.get(),
                "tabs_cdc": self.entry_tabs_cdc.get(),
                "tabs_final": self.entry_tabs_final.get(),
                "quantidade": self.entry_qtd.get(),
                "enter_final": self.chk_enter_var.get(),
                "regra_sem_cdc": self.cmb_sem_cdc.get(),
                "pos_cdc": {"x": self.pos_cdc.x, "y": self.pos_cdc.y} if self.pos_cdc else None,
                "pos_valor": {"x": self.pos_valor.x, "y": self.pos_valor.y} if self.pos_valor else None,
                "pos_percentual": {"x": self.pos_percentual.x, "y": self.pos_percentual.y} if self.pos_percentual else None,
            }

            caminho = filedialog.asksaveasfilename(
                title="Salvar layout",
                defaultextension=".json",
                filetypes=[("JSON", "*.json")],
                initialfile="layout_preenchedor_cdc.json"
            )
            if not caminho:
                return

            with open(caminho, "w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False, indent=4)

            self.log(f"Layout salvo em: {caminho}")
            self.set_status("layout salvo", self.COR_SUCESSO)

        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o layout.\n\nDetalhe: {e}")

    def carregar_layout_json(self):
        try:
            caminho = filedialog.askopenfilename(
                title="Carregar layout",
                filetypes=[("JSON", "*.json")]
            )
            if not caminho:
                return

            with open(caminho, "r", encoding="utf-8") as f:
                dados = json.load(f)

            self.cmb_usar.set(dados.get("usar_campo", "VALOR"))
            self.cmb_modo.set(dados.get("modo", "TAB"))
            self.entry_delay.delete(0, "end")
            self.entry_delay.insert(0, dados.get("delay", "0,15"))

            self.entry_delay_inicial.delete(0, "end")
            self.entry_delay_inicial.insert(0, dados.get("delay_inicial", "3"))

            self.entry_batch.delete(0, "end")
            self.entry_batch.insert(0, dados.get("batch", "20"))

            self.entry_tabs_cdc.delete(0, "end")
            self.entry_tabs_cdc.insert(0, dados.get("tabs_cdc", "3"))

            self.entry_tabs_final.delete(0, "end")
            self.entry_tabs_final.insert(0, dados.get("tabs_final", "2"))

            self.entry_qtd.delete(0, "end")
            self.entry_qtd.insert(0, dados.get("quantidade", "TODOS"))

            self.chk_enter_var.set(bool(dados.get("enter_final", True)))
            self.cmb_sem_cdc.set(dados.get("regra_sem_cdc", "IGNORAR"))

            pos_cdc = dados.get("pos_cdc")
            pos_valor = dados.get("pos_valor")
            pos_percentual = dados.get("pos_percentual")

            if pos_cdc:
                self.pos_cdc = pyautogui.Point(pos_cdc["x"], pos_cdc["y"])
            if pos_valor:
                self.pos_valor = pyautogui.Point(pos_valor["x"], pos_valor["y"])
            if pos_percentual:
                self.pos_percentual = pyautogui.Point(pos_percentual["x"], pos_percentual["y"])

            self.atualizar_labels_posicoes()
            self.log(f"Layout carregado de: {caminho}")
            self.set_status("layout carregado", self.COR_SUCESSO)

        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível carregar o layout.\n\nDetalhe: {e}")

    # =========================================================
    # FECHAMENTO
    # =========================================================
    def fechar_app(self):
        try:
            keyboard.unhook_all_hotkeys()
        except Exception:
            pass
        self.app.destroy()

    def run(self):
        self.app.mainloop()


if __name__ == "__main__":
    app = PreenchedorCDCColaF8()
    app.run()
