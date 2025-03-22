import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import re
import os
import csv
import subprocess
import platform
import sys
import time
import threading
import PyPDF2  # Certifique-se de ter a versão atualizada
import pikepdf
import win32com.client as win32

def habilitar_cores_terminal():
    """
    Executa o comando para habilitar cores no terminal do Windows.
    O comando é executado em segundo plano.
    """
    if platform.system() == "Windows":
        comando = r'reg add HKEY_CURRENT_USER\Console /v VirtualTerminalLevel /t REG_DWORD /d 0x00000001 /f'
        try:
            subprocess.Popen(comando, shell=True)
        except Exception as e:
            print(f"Erro ao habilitar cores no terminal: {e}")

def gerar_banner():
    """
    Tenta gerar um banner com pyfiglet. Se não conseguir, retorna um banner padrão.
    """
    try:
        import pyfiglet
        banner = pyfiglet.figlet_format("PDF Protect Tool", font="slant")
    except ImportError:
        banner = "===== PDF PROTECT TOOL ====="
    return banner

def animate_hourglass(stop_event):
    """
    Exibe uma animação de contador de tempo em verde no terminal enquanto o sistema estiver rodando.
    Usa um spinner simples para simular a ampulheta e mostra o tempo decorrido.
    """
    spinner_frames = ["│", "/", "–", "\\"]
    start_time = time.time()
    frame_index = 0
    while not stop_event.is_set():
        elapsed = int(time.time() - start_time)
        # Monta a linha com o tempo e o frame do spinner em verde (ANSI 92)
        sys.stdout.write(f"\r\033[92mTempo em execução: {elapsed:3d} s {spinner_frames[frame_index]}\033[0m")
        sys.stdout.flush()
        time.sleep(0.2)
        frame_index = (frame_index + 1) % len(spinner_frames)
    # Limpa a linha quando a animação for interrompida
    sys.stdout.write("\r" + " " * 50 + "\r")
    sys.stdout.flush()

def print_entrada():
    # Cores com ANSI (caso o terminal suporte)
    azul = "\033[94m"
    amarelo = "\033[93m"
    reset = "\033[0m"
    
    banner = gerar_banner()
    mensagem = (
        f"{azul}{banner}{reset}\n"
        f"{amarelo}Bem-vindo ao Sistema de Proteção de PDF e Geração de Rascunho de E-mail{reset}\n\n"
        f"{amarelo}Este sistema realiza as seguintes ações:{reset}\n"
        "  1. Seleciona um arquivo PDF através de uma janela de diálogo.\n"
        "  2. Extrai o CPF presente no PDF no formato 999.999.999-99.\n"
        "  3. Protege o PDF com uma senha (o CPF sem formatação) de forma automática.\n"
        "  4. Cria um rascunho de e-mail no Outlook com o PDF protegido anexado.\n\n"
        f"{amarelo}Se preferir definir uma senha personalizada, clique Não na próxima pergunta.{reset}\n"
        "\nBy Anderson Marinho versão 1.0 - 03-2025\n"
    )
    print(mensagem)
    mostrar_total_casos()  # Exibe o total de casos já executados

def mostrar_total_casos():
    """
    Lê o arquivo de log (logInforme.csv) e conta o número de execuções (linhas, desconsiderando o cabeçalho).
    Exibe o total no terminal.
    """
    log_file = "logInforme.csv"
    total = 0
    if os.path.exists(log_file):
        with open(log_file, "r", newline="", encoding="utf-8") as csvfile:
            reader = csv.DictReader(csvfile)
            for _ in reader:
                total += 1
    print(f"Número total de casos executados: {total}\n")

def get_pdf_path():
    """Abre uma janela para o usuário selecionar o arquivo PDF."""
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not file_path:
        messagebox.showerror("Erro", "Nenhum arquivo foi selecionado!")
    root.destroy()
    return file_path

def extract_cpf_from_pdf(pdf_path):
    """
    Percorre as páginas do PDF procurando o primeiro CPF no formato 999.999.999-99.
    Retorna o CPF encontrado somente com dígitos.
    """
    cpf_pattern = r'\d{3}\.\d{3}\.\d{3}-\d{2}'
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text = page.extract_text() or ""
            match = re.search(cpf_pattern, text)
            if match:
                cpf_found = match.group()
                return cpf_found.replace('.', '').replace('-', '')
    raise ValueError("CPF não encontrado no PDF.")

def protect_pdf_with_password(pdf_path, password):
    """
    Cria uma cópia do PDF protegida com senha utilizando pikepdf.
    O novo arquivo terá o sufixo '_protegido.pdf' na mesma pasta.
    """
    base, ext = os.path.splitext(pdf_path)
    output_pdf_path = f"{base}_protegido{ext}"
    if os.path.exists(output_pdf_path):
        os.remove(output_pdf_path)
    with pikepdf.open(pdf_path) as pdf:
        pdf.save(
            output_pdf_path,
            encryption=pikepdf.Encryption(owner=password, user=password, R=4)
        )
    return output_pdf_path

def save_draft_with_attachment(pdf_path):
    """
    Cria um rascunho de e-mail no Outlook com o PDF anexado.
    """
    pdf_path_normalizado = os.path.normpath(os.path.abspath(pdf_path))
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = "Informe de Rendimentos 2024/2025"
    mail.Body = (
        "Segue em anexo o informe de rendimentos protegido.\n"
        "A senha para abrir o PDF é o seu CPF somente números, sem ponto ou traço.\n"
        "\n\n"
        "Atenciosamente,\n"
    )
    mail.Attachments.Add(pdf_path_normalizado)
    mail.Save()
    mail.Display()

def update_log(nome, cpf):
    """
    Registra a execução no arquivo de log (logInforme.csv), adicionando uma nova linha para cada operação.
    """
    log_file = "logInforme.csv"  # Usado para leitura e gravação
    fieldnames = ["Nome", "CPF", "Quantidade Executada"]
    write_header = not os.path.exists(log_file)
    with open(log_file, "a", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        if write_header:
            writer.writeheader()
        # Cada execução é registrada com quantidade = 1
        writer.writerow({"Nome": nome, "CPF": cpf, "Quantidade Executada": "1"})

def proteger_documento_personalizado():
    """
    Permite ao usuário selecionar um PDF e digitar uma senha para protegê-lo manualmente.
    O arquivo é salvo com o sufixo '_protegidoSenha' e é anexado ao rascunho de e-mail.
    """
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()
    pdf_path = filedialog.askopenfilename(
        title="Selecione o PDF para proteção personalizada",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_path:
        messagebox.showerror("Erro", "Nenhum arquivo foi selecionado!")
        root.destroy()
        return
    senha = simpledialog.askstring("Senha", "Digite a senha para proteger o documento:", show="*")
    if not senha:
        messagebox.showerror("Erro", "Nenhuma senha foi digitada!")
        root.destroy()
        return
    root.destroy()
    try:
        base, ext = os.path.splitext(pdf_path)
        output_pdf_path = f"{base}_protegidoSenha{ext}"
        if os.path.exists(output_pdf_path):
            os.remove(output_pdf_path)
        with pikepdf.open(pdf_path) as pdf:
            pdf.save(
                output_pdf_path,
                encryption=pikepdf.Encryption(owner=senha, user=senha, R=4)
            )
        messagebox.showinfo("Sucesso", f"Documento protegido com sucesso!\nSalvo como: {output_pdf_path}")
        print(f"Documento protegido personalizado criado: {output_pdf_path}")
        # Anexa o documento protegido ao rascunho de e-mail
        save_draft_with_attachment(output_pdf_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        print(f"Erro ao proteger documento personalizado: {e}")

def perguntar_continuar():
    """
    Pergunta se o usuário deseja processar outro arquivo.
    Como padrão, o retorno é 'Sim' (caso o usuário não altere, a operação continuará).
    """
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()
    # Usando askyesno, onde o padrão é Sim se o usuário pressionar Enter
    resposta = messagebox.askyesno("Continuar", "Deseja processar outro arquivo?")
    root.destroy()
    return resposta

def main():
    """
    Fluxo padrão: Seleciona PDF, extrai CPF, protege com senha (usando o CPF) e cria rascunho de e-mail.
    """
    try:
        pdf_path = get_pdf_path()
        if not pdf_path:
            return False
        nome_arquivo = os.path.basename(pdf_path)
        cpf = extract_cpf_from_pdf(pdf_path)
        print(f"CPF encontrado: {cpf}")
        protected_pdf_path = protect_pdf_with_password(pdf_path, cpf)
        print(f"PDF protegido criado: {protected_pdf_path}")
        save_draft_with_attachment(protected_pdf_path)
        print("Rascunho do e-mail salvo e exibido no Outlook.")
        update_log(nome_arquivo, cpf)
        root = tk.Tk()
        root.attributes("-topmost", True)
        root.withdraw()
        messagebox.showinfo("Sucesso", "Processo concluído com sucesso!")
        root.destroy()
        return True
    except Exception as e:
        root = tk.Tk()
        root.attributes("-topmost", True)
        root.withdraw()
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        root.destroy()
        print(f"Erro: {e}")
        return False

if __name__ == "__main__":
    # Habilita as cores do terminal (executa em segundo plano, se Windows)
    habilitar_cores_terminal()

    # Inicia a animação de contador de tempo em segundo plano
    stop_animation = threading.Event()
    animation_thread = threading.Thread(target=animate_hourglass, args=(stop_animation,), daemon=True)
    animation_thread.start()

    print_entrada()
    # A pergunta de proteção foi invertida:
    # Se o usuário deseja processar com proteção automática, ele clica Sim.
    # Se quiser definir uma senha personalizada, clique Não.
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()
    opcao = messagebox.askquestion("Opção de Proteção", "Deseja processar o documento com proteção automática?")
    root.destroy()
    if opcao == "yes":
        if not main():
            print("Encerrando o sistema. Até logo!")
    else:
        proteger_documento_personalizado()

    while perguntar_continuar():
        root = tk.Tk()
        root.attributes("-topmost", True)
        root.withdraw()
        opcao = messagebox.askquestion("Opção de Proteção", "Deseja processar o documento com proteção automática?")
        root.destroy()
        if opcao == "yes":
            if not main():
                print("Encerrando o sistema. Até logo!")
                break
        else:
            proteger_documento_personalizado()

    print("Encerrando o sistema. Até logo!")
    # Finaliza a animação e aguarda a thread encerrar
    stop_animation.set()
    animation_thread.join()
