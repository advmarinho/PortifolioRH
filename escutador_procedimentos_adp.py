import os
import re
import time
import queue
import datetime
import webbrowser
from pathlib import Path

import customtkinter as ctk
import pandas as pd
import pyautogui
import pyperclip
from pynput import keyboard

try:
    import win32gui
except ImportError:
    win32gui = None


class JanelaNomePlaybook(ctk.CTkToplevel):
    def __init__(self, master, callback_salvar):
        super().__init__(master)

        self.callback_salvar = callback_salvar

        self.title("Nome do Playbook")
        self.geometry("620x260")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self.grab_set()

        self.grid_columnconfigure(0, weight=1)

        self.frame_topo = ctk.CTkFrame(self, fg_color="#121C4E", corner_radius=0)
        self.frame_topo.grid(row=0, column=0, sticky="ew")

        self.lbl_titulo = ctk.CTkLabel(
            self.frame_topo,
            text="Nome do Playbook",
            text_color="white",
            font=("Arial", 20, "bold")
        )
        self.lbl_titulo.pack(anchor="w", padx=18, pady=(14, 4))

        self.lbl_subtitulo = ctk.CTkLabel(
            self.frame_topo,
            text="Informe o nome do procedimento que será capturado.",
            text_color="white",
            font=("Arial", 12)
        )
        self.lbl_subtitulo.pack(anchor="w", padx=18, pady=(0, 12))

        self.lbl_orientacao = ctk.CTkLabel(
            self,
            text="Exemplo: ADP - Admissão Preliminar ou ADP - Holerite Provisório",
            font=("Arial", 12),
            text_color="#333333"
        )
        self.lbl_orientacao.grid(row=1, column=0, sticky="w", padx=18, pady=(16, 4))

        self.entry_nome = ctk.CTkEntry(
            self,
            placeholder_text="Digite o nome do playbook",
            font=("Arial", 14),
            height=38
        )
        self.entry_nome.grid(row=2, column=0, sticky="ew", padx=18, pady=(4, 16))
        self.entry_nome.focus_set()

        self.frame_botoes = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_botoes.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 18))
        self.frame_botoes.grid_columnconfigure(0, weight=1)

        self.btn_padrao = ctk.CTkButton(
            self.frame_botoes,
            text="Usar padrão",
            command=self.usar_padrao,
            fg_color="#646464",
            hover_color="#333333",
            width=120
        )
        self.btn_padrao.grid(row=0, column=1, padx=6)

        self.btn_salvar = ctk.CTkButton(
            self.frame_botoes,
            text="Salvar nome",
            command=self.salvar,
            fg_color="#0083CA",
            hover_color="#003C64",
            width=130
        )
        self.btn_salvar.grid(row=0, column=2, padx=6)

        self.bind("<Return>", lambda event: self.salvar())
        self.bind("<Escape>", lambda event: self.usar_padrao())

    def salvar(self):
        nome = self.entry_nome.get().strip()

        if not nome:
            nome = "Playbook ADP Capturado"

        self.callback_salvar(nome)
        self.destroy()

    def usar_padrao(self):
        self.callback_salvar("Playbook ADP Capturado")
        self.destroy()


class JanelaDescricao(ctk.CTkToplevel):
    def __init__(self, master, callback_salvar, dados_base):
        super().__init__(master)

        self.callback_salvar = callback_salvar
        self.dados_base = dados_base

        self.title("Descrever etapa capturada")
        self.geometry("780x560")
        self.resizable(True, True)
        self.attributes("-topmost", True)
        self.grab_set()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        self.frame_topo = ctk.CTkFrame(self, fg_color="#121C4E", corner_radius=0)
        self.frame_topo.grid(row=0, column=0, sticky="ew")

        self.lbl_titulo = ctk.CTkLabel(
            self.frame_topo,
            text="Nova etapa capturada",
            text_color="white",
            font=("Arial", 22, "bold")
        )
        self.lbl_titulo.pack(anchor="w", padx=20, pady=(16, 4))

        self.lbl_subtitulo = ctk.CTkLabel(
            self.frame_topo,
            text="Revise o texto selecionado e descreva a ação do procedimento.",
            text_color="white",
            font=("Arial", 13)
        )
        self.lbl_subtitulo.pack(anchor="w", padx=20, pady=(0, 14))

        self.frame_info = ctk.CTkFrame(self, fg_color="#F2F2F2")
        self.frame_info.grid(row=1, column=0, sticky="ew", padx=16, pady=(14, 8))
        self.frame_info.grid_columnconfigure(1, weight=1)

        self.lbl_passo = ctk.CTkLabel(
            self.frame_info,
            text=f"Passo: {dados_base.get('Passo', '')}",
            font=("Arial", 13, "bold"),
            text_color="#003C64"
        )
        self.lbl_passo.grid(row=0, column=0, padx=12, pady=8, sticky="w")

        self.lbl_janela = ctk.CTkLabel(
            self.frame_info,
            text=f"Janela ativa: {dados_base.get('JanelaAtiva', '')}",
            font=("Arial", 12),
            text_color="#333333",
            anchor="w"
        )
        self.lbl_janela.grid(row=0, column=1, padx=12, pady=8, sticky="ew")

        self.lbl_texto_sel = ctk.CTkLabel(
            self,
            text="Texto selecionado capturado:",
            font=("Arial", 14, "bold"),
            text_color="#003C64"
        )
        self.lbl_texto_sel.grid(row=2, column=0, sticky="w", padx=18, pady=(8, 4))

        self.txt_selecionado = ctk.CTkTextbox(
            self,
            height=90,
            font=("Arial", 13)
        )
        self.txt_selecionado.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 10))
        self.txt_selecionado.insert("1.0", dados_base.get("TextoSelecionado", ""))

        self.lbl_descricao = ctk.CTkLabel(
            self,
            text="Descrição da ação/procedimento:",
            font=("Arial", 14, "bold"),
            text_color="#003C64"
        )
        self.lbl_descricao.grid(row=4, column=0, sticky="w", padx=18, pady=(4, 4))

        self.txt_descricao = ctk.CTkTextbox(
            self,
            height=180,
            font=("Arial", 13)
        )
        self.txt_descricao.grid(row=5, column=0, sticky="nsew", padx=18, pady=(0, 12))

        texto_sugerido = self.gerar_sugestao_descricao(dados_base.get("TextoSelecionado", ""))
        self.txt_descricao.insert("1.0", texto_sugerido)
        self.txt_descricao.focus_set()

        self.frame_botoes = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_botoes.grid(row=6, column=0, sticky="ew", padx=18, pady=(0, 18))
        self.frame_botoes.grid_columnconfigure(0, weight=1)

        self.btn_cancelar = ctk.CTkButton(
            self.frame_botoes,
            text="Descartar",
            command=self.descartar,
            fg_color="#646464",
            hover_color="#333333",
            width=120
        )
        self.btn_cancelar.grid(row=0, column=1, padx=6)

        self.btn_salvar = ctk.CTkButton(
            self.frame_botoes,
            text="Salvar etapa",
            command=self.salvar,
            fg_color="#0083CA",
            hover_color="#003C64",
            width=140
        )
        self.btn_salvar.grid(row=0, column=2, padx=6)

        self.bind("<Control-Return>", lambda event: self.salvar())
        self.bind("<Escape>", lambda event: self.descartar())

    def gerar_sugestao_descricao(self, texto):
        texto = (texto or "").strip()

        if texto:
            return f"Acessar, selecionar ou validar a opção: {texto}"

        return "Descrever aqui a ação realizada nesta etapa."

    def salvar(self):
        texto_selecionado_final = self.txt_selecionado.get("1.0", "end").strip()
        descricao = self.txt_descricao.get("1.0", "end").strip()

        self.dados_base["TextoSelecionado"] = texto_selecionado_final
        self.dados_base["DescricaoProcedimento"] = descricao
        self.dados_base["Status"] = "Salvo"

        self.callback_salvar(self.dados_base)
        self.destroy()

    def descartar(self):
        try:
            caminho_print = self.dados_base.get("Print", "")
            if caminho_print and os.path.exists(caminho_print):
                os.remove(caminho_print)
        except Exception:
            pass

        self.destroy()


class EscutadorPlaybookADP(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Escutador de Playbook ADP - Captura Manual F8")
        self.geometry("1240x780")
        self.minsize(1100, 700)

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.pasta_base = Path(os.getcwd()) / "playbook_adp_capturado"
        self.pasta_prints = self.pasta_base / "prints"
        self.pasta_base.mkdir(parents=True, exist_ok=True)
        self.pasta_prints.mkdir(parents=True, exist_ok=True)

        self.registros = []
        self.contador = 0
        self.rodando = False
        self.listener_teclado = None
        self.fila_interface = queue.Queue()

        self.capturar_texto_selecionado = True
        self.abrir_descricao_apos_f8 = True

        self.nome_playbook = ""
        self.nome_arquivo_base = "playbook_adp_capturado"

        self._montar_layout()
        self._processar_fila_interface()

    def _montar_layout(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        self.frame_topo = ctk.CTkFrame(self, fg_color="#121C4E", corner_radius=0)
        self.frame_topo.grid(row=0, column=0, sticky="ew")

        self.lbl_titulo = ctk.CTkLabel(
            self.frame_topo,
            text="Escutador de Playbook ADP",
            text_color="white",
            font=("Arial", 24, "bold")
        )
        self.lbl_titulo.pack(anchor="w", padx=22, pady=(16, 4))

        self.lbl_subtitulo = ctk.CTkLabel(
            self.frame_topo,
            text="Use F8 para capturar etapa, texto selecionado, print, janela ativa e descrição do procedimento.",
            text_color="white",
            font=("Arial", 13)
        )
        self.lbl_subtitulo.pack(anchor="w", padx=22, pady=(0, 14))

        self.frame_identificacao = ctk.CTkFrame(self, fg_color="#F2F2F2")
        self.frame_identificacao.grid(row=1, column=0, sticky="ew", padx=16, pady=(12, 0))
        self.frame_identificacao.grid_columnconfigure(1, weight=1)

        self.lbl_nome_playbook_titulo = ctk.CTkLabel(
            self.frame_identificacao,
            text="Playbook:",
            text_color="#003C64",
            font=("Arial", 13, "bold")
        )
        self.lbl_nome_playbook_titulo.grid(row=0, column=0, padx=(14, 6), pady=10, sticky="w")

        self.lbl_nome_playbook = ctk.CTkLabel(
            self.frame_identificacao,
            text="Ainda não definido. Clique em Iniciar para informar o nome.",
            text_color="#333333",
            font=("Arial", 13),
            anchor="w"
        )
        self.lbl_nome_playbook.grid(row=0, column=1, padx=6, pady=10, sticky="ew")

        self.btn_alterar_nome = ctk.CTkButton(
            self.frame_identificacao,
            text="Alterar nome",
            command=self.pedir_nome_playbook,
            fg_color="#005A64",
            hover_color="#003C64",
            width=120
        )
        self.btn_alterar_nome.grid(row=0, column=2, padx=14, pady=8)

        self.frame_controles = ctk.CTkFrame(self)
        self.frame_controles.grid(row=2, column=0, sticky="ew", padx=16, pady=12)

        self.frame_controles.grid_columnconfigure(10, weight=1)

        self.btn_iniciar = ctk.CTkButton(
            self.frame_controles,
            text="Iniciar",
            command=self.iniciar,
            fg_color="#0083CA",
            hover_color="#003C64",
            width=110
        )
        self.btn_iniciar.grid(row=0, column=0, padx=6, pady=12)

        self.btn_parar = ctk.CTkButton(
            self.frame_controles,
            text="Parar",
            command=self.parar,
            fg_color="#7D0041",
            hover_color="#5A0030",
            width=100
        )
        self.btn_parar.grid(row=0, column=1, padx=6, pady=12)

        self.btn_capturar = ctk.CTkButton(
            self.frame_controles,
            text="Capturar agora F8",
            command=self.capturar_etapa,
            fg_color="#003C64",
            hover_color="#121C4E",
            width=150
        )
        self.btn_capturar.grid(row=0, column=2, padx=6, pady=12)

        self.btn_exportar_excel = ctk.CTkButton(
            self.frame_controles,
            text="Exportar Excel",
            command=self.exportar_excel,
            fg_color="#005A64",
            hover_color="#003C64",
            width=130
        )
        self.btn_exportar_excel.grid(row=0, column=3, padx=6, pady=12)

        self.btn_exportar_html = ctk.CTkButton(
            self.frame_controles,
            text="Gerar HTML",
            command=self.gerar_html,
            fg_color="#0083CA",
            hover_color="#003C64",
            width=120
        )
        self.btn_exportar_html.grid(row=0, column=4, padx=6, pady=12)

        self.btn_abrir_pasta = ctk.CTkButton(
            self.frame_controles,
            text="Abrir pasta",
            command=self.abrir_pasta_saida,
            fg_color="#646464",
            hover_color="#333333",
            width=110
        )
        self.btn_abrir_pasta.grid(row=0, column=5, padx=6, pady=12)

        self.btn_limpar = ctk.CTkButton(
            self.frame_controles,
            text="Limpar",
            command=self.limpar,
            fg_color="#8C321E",
            hover_color="#5A1F12",
            width=100
        )
        self.btn_limpar.grid(row=0, column=6, padx=6, pady=12)

        self.lbl_status = ctk.CTkLabel(
            self.frame_controles,
            text="Status: parado",
            text_color="#7D0041",
            font=("Arial", 13, "bold")
        )
        self.lbl_status.grid(row=0, column=7, padx=16, pady=12, sticky="w")

        self.frame_opcoes = ctk.CTkFrame(self)
        self.frame_opcoes.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 12))

        self.chk_texto = ctk.CTkCheckBox(
            self.frame_opcoes,
            text="Tentar capturar texto selecionado com Ctrl+C",
            command=self.alternar_texto
        )
        self.chk_texto.select()
        self.chk_texto.grid(row=0, column=0, padx=14, pady=10, sticky="w")

        self.chk_descricao = ctk.CTkCheckBox(
            self.frame_opcoes,
            text="Abrir caixa de descrição após F8",
            command=self.alternar_descricao
        )
        self.chk_descricao.select()
        self.chk_descricao.grid(row=0, column=1, padx=14, pady=10, sticky="w")

        self.lbl_atalhos = ctk.CTkLabel(
            self.frame_opcoes,
            text="Atalhos: F8 captura etapa | F10 exporta Excel | F12 gera HTML | ESC para",
            text_color="#333333",
            font=("Arial", 12)
        )
        self.lbl_atalhos.grid(row=0, column=2, padx=14, pady=10, sticky="w")

        self.frame_corpo = ctk.CTkFrame(self)
        self.frame_corpo.grid(row=4, column=0, sticky="nsew", padx=16, pady=(0, 12))
        self.frame_corpo.grid_columnconfigure(0, weight=2)
        self.frame_corpo.grid_columnconfigure(1, weight=1)
        self.frame_corpo.grid_rowconfigure(1, weight=1)

        self.lbl_log = ctk.CTkLabel(
            self.frame_corpo,
            text="Registros capturados",
            font=("Arial", 16, "bold"),
            text_color="#003C64"
        )
        self.lbl_log.grid(row=0, column=0, sticky="w", padx=14, pady=(12, 4))

        self.lbl_dica = ctk.CTkLabel(
            self.frame_corpo,
            text="Dica operacional",
            font=("Arial", 16, "bold"),
            text_color="#003C64"
        )
        self.lbl_dica.grid(row=0, column=1, sticky="w", padx=14, pady=(12, 4))

        self.txt_log = ctk.CTkTextbox(
            self.frame_corpo,
            font=("Consolas", 12),
            wrap="none"
        )
        self.txt_log.grid(row=1, column=0, sticky="nsew", padx=14, pady=(4, 14))

        self.txt_dica = ctk.CTkTextbox(
            self.frame_corpo,
            font=("Arial", 13),
            wrap="word"
        )
        self.txt_dica.grid(row=1, column=1, sticky="nsew", padx=(0, 14), pady=(4, 14))

        self.txt_dica.insert(
            "1.0",
            "Modelo recomendado para ADP:\n\n"
            "1. Clique em Iniciar e informe o nome do Playbook.\n"
            "2. Abra a tela que deseja documentar.\n"
            "3. Selecione com o mouse o texto principal da tela ou do menu.\n"
            "4. Pressione F8.\n"
            "5. Revise a descrição sugerida e salve a etapa.\n"
            "6. Continue para a próxima tela.\n\n"
            "Exemplo:\n"
            "Nome do Playbook: ADP - Admissão Preliminar\n"
            "Texto selecionado: Funcionalidades\n"
            "Descrição: Acessar Folha > eSocial > Funcionalidades para iniciar o processo.\n\n"
            "Arquivos gerados:\n"
            "adp_admissao_preliminar.xlsx\n"
            "adp_admissao_preliminar.html\n\n"
            "Observação:\n"
            "Revise os prints antes de compartilhar, pois podem conter dados pessoais."
        )
        self.txt_dica.configure(state="disabled")

        self.frame_rodape = ctk.CTkFrame(self, fg_color="#F2F2F2", corner_radius=0)
        self.frame_rodape.grid(row=5, column=0, sticky="ew")

        self.lbl_rodape = ctk.CTkLabel(
            self.frame_rodape,
            text="Anderson Marinho | Igarapé Digital",
            text_color="#646464",
            font=("Arial", 11)
        )
        self.lbl_rodape.pack(anchor="e", padx=16, pady=6)

    def _processar_fila_interface(self):
        try:
            while True:
                acao, payload = self.fila_interface.get_nowait()

                if acao == "capturar":
                    self.capturar_etapa()

                elif acao == "exportar_excel":
                    self.exportar_excel()

                elif acao == "gerar_html":
                    self.gerar_html()

                elif acao == "parar":
                    self.parar()

                elif acao == "log":
                    self.log(payload)

        except queue.Empty:
            pass

        self.after(100, self._processar_fila_interface)

    def alternar_texto(self):
        self.capturar_texto_selecionado = bool(self.chk_texto.get())

    def alternar_descricao(self):
        self.abrir_descricao_apos_f8 = bool(self.chk_descricao.get())

    def pedir_nome_playbook(self):
        JanelaNomePlaybook(self, self.definir_nome_playbook)

    def definir_nome_playbook(self, nome):
        self.nome_playbook = nome.strip() or "Playbook ADP Capturado"
        self.nome_arquivo_base = self.normalizar_nome_arquivo(self.nome_playbook)

        self.lbl_nome_playbook.configure(
            text=f"{self.nome_playbook} | Arquivo base: {self.nome_arquivo_base}"
        )

        self.log(f"Nome do Playbook definido: {self.nome_playbook}")
        self.log(f"Arquivo base definido: {self.nome_arquivo_base}")

        if not self.rodando:
            self.iniciar_captura_efetiva()

    def iniciar(self):
        if self.rodando:
            self.log("A captura já está ativa.")
            return

        if not self.nome_playbook:
            self.pedir_nome_playbook()
            return

        self.iniciar_captura_efetiva()

    def iniciar_captura_efetiva(self):
        if self.rodando:
            return

        self.rodando = True
        self.lbl_status.configure(text="Status: ativo", text_color="#0083CA")

        self.listener_teclado = keyboard.Listener(on_press=self._ao_pressionar_tecla)
        self.listener_teclado.start()

        self.log("Captura iniciada.")
        self.log(f"Playbook: {self.nome_playbook or 'Playbook ADP Capturado'}")
        self.log("Use F8 para capturar uma etapa.")
        self.log("Use F10 para exportar Excel.")
        self.log("Use F12 para gerar HTML.")
        self.log("Use ESC para parar.")

    def parar(self):
        if not self.rodando:
            self.lbl_status.configure(text="Status: parado", text_color="#7D0041")
            return

        self.rodando = False
        self.lbl_status.configure(text="Status: parado", text_color="#7D0041")

        try:
            if self.listener_teclado:
                self.listener_teclado.stop()
        except Exception:
            pass

        self.log("Captura parada.")

    def _ao_pressionar_tecla(self, key):
        if not self.rodando:
            return False

        try:
            if key == keyboard.Key.f8:
                self.fila_interface.put(("capturar", None))

            elif key == keyboard.Key.f10:
                self.fila_interface.put(("exportar_excel", None))

            elif key == keyboard.Key.f12:
                self.fila_interface.put(("gerar_html", None))

            elif key == keyboard.Key.esc:
                self.fila_interface.put(("parar", None))
                return False

        except Exception as erro:
            self.fila_interface.put(("log", f"Erro no atalho de teclado: {erro}"))

    def obter_janela_ativa(self):
        if win32gui is None:
            return "Não identificado"

        try:
            janela = win32gui.GetForegroundWindow()
            titulo = win32gui.GetWindowText(janela)
            return titulo if titulo else "Sem título"
        except Exception:
            return "Não identificado"

    def capturar_texto_clipboard(self):
        if not self.capturar_texto_selecionado:
            return ""

        texto_anterior = ""

        try:
            texto_anterior = pyperclip.paste()
        except Exception:
            texto_anterior = ""

        texto_capturado = ""

        try:
            pyperclip.copy("")
            time.sleep(0.05)

            pyautogui.hotkey("ctrl", "c")
            time.sleep(0.25)

            texto_capturado = pyperclip.paste()
            texto_capturado = texto_capturado.strip() if texto_capturado else ""

        except Exception:
            texto_capturado = ""

        try:
            if texto_anterior:
                pyperclip.copy(texto_anterior)
        except Exception:
            pass

        return texto_capturado

    def capturar_print(self, numero_passo):
        nome = f"{self.nome_arquivo_base}_passo_{numero_passo:03d}.png"
        caminho = self.pasta_prints / nome

        try:
            imagem = pyautogui.screenshot()
            imagem.save(caminho)
            return str(caminho)
        except Exception as erro:
            self.log(f"Erro ao capturar print: {erro}")
            return ""

    def capturar_etapa(self):
        if not self.nome_playbook:
            self.pedir_nome_playbook()
            return

        self.contador += 1

        data_hora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        janela = self.obter_janela_ativa()

        texto_selecionado = self.capturar_texto_clipboard()
        caminho_print = self.capturar_print(self.contador)

        dados = {
            "Playbook": self.nome_playbook,
            "Passo": self.contador,
            "DataHora": data_hora,
            "JanelaAtiva": janela,
            "TextoSelecionado": texto_selecionado,
            "DescricaoProcedimento": "",
            "Print": caminho_print,
            "Status": "Pendente descrição"
        }

        if self.abrir_descricao_apos_f8:
            JanelaDescricao(self, self.salvar_registro, dados)
        else:
            dados["DescricaoProcedimento"] = texto_selecionado or "Etapa capturada sem descrição."
            dados["Status"] = "Salvo"
            self.salvar_registro(dados)

    def salvar_registro(self, dados):
        self.registros.append(dados)

        passo = dados.get("Passo", "")
        texto = dados.get("TextoSelecionado", "")
        desc = dados.get("DescricaoProcedimento", "")
        janela = dados.get("JanelaAtiva", "")

        self.log(
            f"Passo {passo:03d} | Texto: {texto or 'Sem texto selecionado'} | "
            f"Descrição: {desc[:90]} | Janela: {janela}"
        )

    def exportar_excel(self):
        if not self.registros:
            self.log("Não há registros para exportar.")
            return

        try:
            arquivo = self.pasta_base / f"{self.nome_arquivo_base}.xlsx"

            df = pd.DataFrame(self.registros)

            colunas = [
                "Playbook",
                "Passo",
                "DataHora",
                "JanelaAtiva",
                "TextoSelecionado",
                "DescricaoProcedimento",
                "Print",
                "Status"
            ]

            for col in colunas:
                if col not in df.columns:
                    df[col] = ""

            df = df[colunas]

            resumo = pd.DataFrame({
                "Campo": [
                    "Nome do Playbook",
                    "Nome do arquivo base",
                    "Total de passos",
                    "Data da exportação",
                    "Pasta base",
                    "Pasta dos prints",
                    "Orientação"
                ],
                "Valor": [
                    self.nome_playbook or "Playbook ADP Capturado",
                    self.nome_arquivo_base,
                    len(df),
                    datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    str(self.pasta_base),
                    str(self.pasta_prints),
                    "Revisar prints antes de compartilhar, pois podem conter dados pessoais."
                ]
            })

            with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Playbook")
                resumo.to_excel(writer, index=False, sheet_name="Resumo")

                workbook = writer.book
                ws = workbook["Playbook"]

                larguras = {
                    "A": 28,
                    "B": 10,
                    "C": 22,
                    "D": 45,
                    "E": 35,
                    "F": 75,
                    "G": 65,
                    "H": 20
                }

                for coluna, largura in larguras.items():
                    ws.column_dimensions[coluna].width = largura

                for row in ws.iter_rows():
                    for cell in row:
                        cell.alignment = cell.alignment.copy(wrap_text=True, vertical="top")

            self.log(f"Excel exportado: {arquivo}")

        except Exception as erro:
            self.log(f"Erro ao exportar Excel: {erro}")

    def gerar_html(self):
        if not self.registros:
            self.log("Não há registros para gerar HTML.")
            return

        try:
            arquivo_html = self.pasta_base / f"{self.nome_arquivo_base}.html"

            titulo_html = self._escapar_html(self.nome_playbook or "Playbook ADP Capturado")

            linhas_html = []

            for item in self.registros:
                passo = item.get("Passo", "")
                data_hora = self._escapar_html(item.get("DataHora", ""))
                janela = self._escapar_html(item.get("JanelaAtiva", ""))
                texto = self._escapar_html(item.get("TextoSelecionado", ""))
                descricao = self._escapar_html(item.get("DescricaoProcedimento", ""))
                caminho_print = item.get("Print", "")

                img_rel = ""

                if caminho_print:
                    try:
                        img_rel = os.path.relpath(caminho_print, self.pasta_base)
                        img_rel = img_rel.replace("\\", "/")
                    except Exception:
                        img_rel = caminho_print

                bloco_img = ""

                if img_rel:
                    bloco_img = f'<img src="{img_rel}" alt="Print do passo {passo}">'

                bloco = f"""
                <section class="passo">
                    <div class="passo-header">
                        <div class="numero">Passo {passo}</div>
                        <div class="data">{data_hora}</div>
                    </div>

                    <div class="conteudo">
                        <p><strong>Janela ativa:</strong> {janela}</p>
                        <p><strong>Texto selecionado:</strong> {texto if texto else "Não informado"}</p>
                        <p><strong>Procedimento:</strong> {descricao if descricao else "Não informado"}</p>
                    </div>

                    <div class="print">
                        {bloco_img}
                    </div>
                </section>
                """

                linhas_html.append(bloco)

            html = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>{titulo_html}</title>
    <style>
        body {{
            margin: 0;
            font-family: Arial, sans-serif;
            background: #F7F7F7;
            color: #333333;
        }}

        header {{
            background: #121C4E;
            color: white;
            padding: 28px 42px;
        }}

        header h1 {{
            margin: 0;
            font-size: 30px;
        }}

        header p {{
            margin: 8px 0 0 0;
            font-size: 14px;
        }}

        main {{
            max-width: 1120px;
            margin: 28px auto;
            padding: 0 20px;
        }}

        .resumo {{
            background: white;
            border-left: 6px solid #0083CA;
            padding: 18px 22px;
            margin-bottom: 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }}

        .resumo p {{
            margin: 6px 0;
            font-size: 15px;
        }}

        .passo {{
            background: white;
            margin-bottom: 26px;
            border: 1px solid #DDDDDD;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }}

        .passo-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: #0083CA;
            color: white;
            padding: 12px 18px;
        }}

        .numero {{
            font-weight: bold;
            font-size: 18px;
        }}

        .data {{
            font-size: 13px;
        }}

        .conteudo {{
            padding: 18px;
            font-size: 15px;
            line-height: 1.5;
        }}

        .conteudo p {{
            margin: 8px 0;
        }}

        .print {{
            padding: 0 18px 18px 18px;
        }}

        .print img {{
            max-width: 100%;
            border: 1px solid #CCCCCC;
            box-shadow: 0 1px 6px rgba(0,0,0,0.12);
        }}

        footer {{
            text-align: right;
            color: #646464;
            font-size: 12px;
            padding: 20px 42px 30px 42px;
        }}

        @media print {{
            body {{
                background: white;
            }}

            .passo {{
                page-break-inside: avoid;
            }}

            header {{
                background: #121C4E;
                color: white;
            }}
        }}
    </style>
</head>
<body>
    <header>
        <h1>{titulo_html}</h1>
        <p>Procedimento gerado a partir de capturas manuais por F8.</p>
    </header>

    <main>
        <div class="resumo">
            <p><strong>Nome do Playbook:</strong> {titulo_html}</p>
            <p><strong>Arquivo base:</strong> {self._escapar_html(self.nome_arquivo_base)}</p>
            <p><strong>Total de passos:</strong> {len(self.registros)}</p>
            <p><strong>Gerado em:</strong> {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}</p>
            <p><strong>Observação:</strong> revise os prints antes de compartilhar, pois podem conter dados pessoais.</p>
        </div>

        {''.join(linhas_html)}
    </main>

    <footer>
        Anderson Marinho | Igarapé Digital
    </footer>
</body>
</html>
"""

            with open(arquivo_html, "w", encoding="utf-8") as f:
                f.write(html)

            self.log(f"HTML gerado: {arquivo_html}")

            try:
                webbrowser.open(str(arquivo_html))
            except Exception:
                pass

        except Exception as erro:
            self.log(f"Erro ao gerar HTML: {erro}")

    def abrir_pasta_saida(self):
        try:
            os.startfile(self.pasta_base)
        except Exception as erro:
            self.log(f"Erro ao abrir pasta: {erro}")

    def limpar(self):
        self.registros.clear()
        self.contador = 0
        self.txt_log.delete("1.0", "end")
        self.log("Registros limpos. O nome do Playbook foi mantido.")

    def log(self, texto):
        data = datetime.datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{data}] {texto}\n")
        self.txt_log.see("end")

    def normalizar_nome_arquivo(self, texto):
        texto = str(texto or "").strip().lower()

        substituicoes = {
            "á": "a", "à": "a", "ã": "a", "â": "a", "ä": "a",
            "é": "e", "ê": "e", "è": "e", "ë": "e",
            "í": "i", "ì": "i", "î": "i", "ï": "i",
            "ó": "o", "ô": "o", "õ": "o", "ò": "o", "ö": "o",
            "ú": "u", "ù": "u", "û": "u", "ü": "u",
            "ç": "c",
            "ñ": "n"
        }

        for antigo, novo in substituicoes.items():
            texto = texto.replace(antigo, novo)

        texto = re.sub(r"[^a-z0-9]+", "_", texto)
        texto = re.sub(r"_+", "_", texto)
        texto = texto.strip("_")

        if not texto:
            texto = "playbook_adp_capturado"

        return texto

    def _escapar_html(self, texto):
        texto = str(texto or "")
        return (
            texto.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&#039;")
        )

    def on_closing(self):
        self.parar()
        self.destroy()


if __name__ == "__main__":
    app = EscutadorPlaybookADP()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()
