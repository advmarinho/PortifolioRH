from __future__ import annotations

import csv
import json
import os
import re
import time
import traceback
import threading
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

from customtkinter import CTkEntry, CTkLabel, CTkButton, CTkTextbox

try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except Exception:
    ctk = None
    CTK_AVAILABLE = False

import pandas as pd
from tkinter import filedialog, messagebox


APP_NAME = "CustomerThinker | PDF Protect + Outlook"
VERSAO = "3.4"
SUBPASTA_PROTEGIDOS = "protegidos"
LOG_NAME = "log_envio_informes.csv"
FOOTER_UI = "Anderson Marinho | Igarapé Digital"
RH_CHAMADO_URL = ""

SUBJECT_TEMPLATE = "Informe de Rendimentos {ano_base} - Matrícula {matricula}"
BODY_TEMPLATE = (
    "Olá {nome},\n\n"
    "Segue em anexo o seu Informe de Rendimentos (ano-base {ano_base}).\n"
    "O arquivo está protegido.\n"
    "Senha: seu CPF (somente números, sem ponto e sem traço).\n\n"
    "Caso identifique divergência de valores/dados ou tenha dificuldade de acesso, "
    "abra um chamado para o RH por este canal.\n"
    "{url_chamado}\n\n"
    "Atenciosamente,\n"
    "Recursos Humanos\n"
)

TEMPLATES_EMAIL = {
    "Informe de Rendimento": {
        "assunto": SUBJECT_TEMPLATE,
        "corpo": BODY_TEMPLATE,
    },
    "Férias": {
        "assunto": "Aviso e Recibo de Férias - Matrícula {matricula}",
        "corpo": (
            "Olá {nome},\n\n"
            "Segue em anexo o seu documento de férias.\n"
            "O arquivo está protegido.\n"
            "Senha: seu CPF (somente números, sem ponto e sem traço).\n\n"
            "Em caso de dúvidas, entre em contato com o RH.\n\n"
            "Atenciosamente,\n"
            "Recursos Humanos\n"
        ),
    },
    "Rescisão": {
        "assunto": "Documentos Rescisórios - Matrícula {matricula}",
        "corpo": (
            "Olá {nome},\n\n"
            "Segue em anexo a documentação rescisória correspondente.\n"
            "O arquivo está protegido.\n"
            "Senha: seu CPF (somente números, sem ponto e sem traço).\n\n"
            "Em caso de dúvidas, entre em contato com o RH.\n\n"
            "Atenciosamente,\n"
            "Recursos Humanos\n"
        ),
    },
    "Documento Diverso": {
        "assunto": "Documento RH - Matrícula {matricula}",
        "corpo": (
            "Olá {nome},\n\n"
            "Segue em anexo o documento correspondente.\n"
            "Caso tenha dúvidas, entre em contato com o RH.\n\n"
            "Atenciosamente,\n"
            "Recursos Humanos\n"
        ),
    },
}


@dataclass
class Colaborador:
    matricula: str
    cpf: str
    nome: str
    email: str


@dataclass
class IdentificacaoPDF:
    cpf_encontrado: str
    nome_encontrado: str
    origem_cpf: str
    cpf_nome_arquivo: str
    divergencia_nome_arquivo: bool
    total_cpfs_validos: int
    texto_extraido: str


@dataclass
class ResultadoProcesso:
    pdf_original: str
    cpf_pdf: str
    nome_pdf: str
    cpf_nome_arquivo: str
    origem_cpf: str
    encontrado_base: bool
    matricula: str
    email: str
    pdf_protegido: str
    status: str
    detalhe: str


def normalizar_texto_simples(valor: str) -> str:
    valor = str(valor or "").strip().upper()
    valor = unicodedata.normalize("NFKD", valor)
    valor = "".join(ch for ch in valor if not unicodedata.combining(ch))
    valor = re.sub(r"[^A-Z0-9 ]", " ", valor)
    valor = re.sub(r"\s+", " ", valor).strip()
    return valor


def normalizar_cpf(valor) -> str:
    if valor is None:
        return ""
    s = re.sub(r"\D", "", str(valor).strip())
    if not s:
        return ""
    if len(s) < 11:
        s = s.zfill(11)
    return s


def formatar_cpf(cpf: str) -> str:
    cpf = normalizar_cpf(cpf)
    if len(cpf) != 11:
        return cpf
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"


def limpar_nome_extraido(nome: str) -> str:
    nome = " ".join(str(nome or "").split()).strip(" -:")
    cortes = [
        " Natureza do Rendimento",
        " CPF",
        " CNPJ",
        " Valores",
        " Ano Calendário",
    ]
    for marcador in cortes:
        pos = nome.upper().find(marcador.upper())
        if pos > 0:
            nome = nome[:pos].strip(" -:")
    return nome


def validar_cpf(cpf: str) -> bool:
    cpf = normalizar_cpf(cpf)

    if len(cpf) != 11:
        return False

    if cpf == cpf[0] * 11:
        return False

    soma1 = sum(int(cpf[i]) * (10 - i) for i in range(9))
    dig1 = (soma1 * 10) % 11
    dig1 = 0 if dig1 == 10 else dig1
    if dig1 != int(cpf[9]):
        return False

    soma2 = sum(int(cpf[i]) * (11 - i) for i in range(10))
    dig2 = (soma2 * 10) % 11
    dig2 = 0 if dig2 == 10 else dig2
    return dig2 == int(cpf[10])


def garantir_pasta(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def caminho_windows_estendido(p: Path) -> str:
    s = str(p)
    if os.name != "nt":
        return s
    s = str(p.resolve())
    if s.startswith('\\\\?\\\\'):
        return s
    if s.startswith('\\\\'):
        return '\\\\?\\\\UNC\\\\' + s.lstrip('\\\\')
    return '\\\\?\\\\' + s


def extrair_ano_base_do_nome_arquivo(nome_arquivo: str) -> str:
    m = re.search(r"(90\d{2})", nome_arquivo)
    if m:
        return m.group(1)
    ano_atual = int(time.strftime("%Y"))
    return str(ano_atual - 1)


def validar_email_basico(email: str) -> bool:
    if not email:
        return False
    e = email.strip()
    if " " in e or e.count("@") != 1:
        return False
    local, dom = e.split("@")
    if not local or not dom or "." not in dom:
        return False
    if not re.fullmatch(r"[A-Za-z0-9._%+\-]+", local):
        return False
    if not re.fullmatch(r"[A-Za-z0-9.\-]+", dom):
        return False
    tld = dom.rsplit(".", 1)[-1]
    if len(tld) < 2 or not re.fullmatch(r"[A-Za-z]{2,}", tld):
        return False
    return True


def limpar_nome_anexo_removendo_cpf(nome_arquivo: str, cpf11: str) -> str:
    base = nome_arquivo
    if cpf11:
        base = base.replace(cpf11, "")
        base = base.replace(formatar_cpf(cpf11), "")
    base = re.sub(r"__+", "_", base)
    base = re.sub(r"_\.", ".", base)
    base = re.sub(r"_-", "_", base)
    base = re.sub(r"_+", "_", base)
    base = re.sub(r"\s+", " ", base)
    base = base.strip("_").strip()
    return base or nome_arquivo


def extrair_nome_do_arquivo_sem_cpf(nome_arquivo: str, cpf11: str = "") -> str:
    nome = Path(str(nome_arquivo or "")).stem
    if cpf11:
        nome = nome.replace(cpf11, "")
        nome = nome.replace(formatar_cpf(cpf11), "")
    nome = re.sub(r"(?i)_?protegido(?:_\d+)?$", "", nome)
    nome = re.sub(r"^[\s_\-\.]+", "", nome)
    nome = re.sub(r"[\s_\-\.]+$", "", nome)
    nome = re.sub(r"[_\-]+", " ", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome


def reduzir_nome_para_caminho(pasta: Path, nome_arquivo: str, margem_segura: int = 230) -> str:
    nome_arquivo = str(nome_arquivo or "").strip()
    if not nome_arquivo:
        return "arquivo.pdf"

    candidato = pasta / nome_arquivo
    if len(str(candidato)) <= margem_segura:
        return nome_arquivo

    stem = Path(nome_arquivo).stem
    suffix = Path(nome_arquivo).suffix or ".pdf"
    excesso = len(str(candidato)) - margem_segura
    novo_tamanho = max(20, len(stem) - excesso)
    stem = stem[:novo_tamanho].rstrip(" _.-")
    nome_reduzido = f"{stem}{suffix}"

    while len(str(pasta / nome_reduzido)) > margem_segura and len(stem) > 20:
        stem = stem[:-1].rstrip(" _.-")
        nome_reduzido = f"{stem}{suffix}"

    return nome_reduzido


def mover_arquivo_windows_seguro(origem: Path, destino: Path) -> None:
    garantir_pasta(destino.parent)
    origem_ext = caminho_windows_estendido(origem)
    destino_ext = caminho_windows_estendido(destino)
    os.replace(origem_ext, destino_ext)


def normalizar_nome_arquivo_manual(nome: str) -> str:
    nome = unicodedata.normalize("NFKD", str(nome or ""))
    nome = "".join(c for c in nome if not unicodedata.combining(c))
    nome = nome.upper()
    nome = re.sub(r"[^\w\s]", "", nome)
    nome = re.sub(r"\s+", "_", nome)
    nome = re.sub(r"_+", "_", nome)
    return nome.strip("_")


def limpar_texto_selecionado_nome(texto: str) -> str:
    texto = str(texto or "").strip()
    for item in [
        "Nome:", "NOME:", "Nome", "NOME", "CPF:", "CPF",
        "Beneficiário:", "BENEFICIÁRIO:", "BENEFICIARIO:",
        "BENEFICIÁRIO", "BENEFICIARIO"
    ]:
        texto = texto.replace(item, "")
    texto = texto.strip(" :-\n\t")
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def encontrar_cpf_simples_no_texto(texto: str) -> str:
    padroes = [
        r"\d{3}\.\d{3}\.\d{3}-\d{2}",
        r"\b\d{11}\b",
    ]
    for padrao in padroes:
        match = re.search(padrao, texto or "")
        if match:
            cpf = normalizar_cpf(match.group())
            if cpf:
                return cpf
    return ""


def _extrair_texto_pypdf2(pdf_path: Path) -> str:
    try:
        try:
            from PyPDF2 import PdfReader
        except Exception:
            from pypdf import PdfReader
    except Exception:
        return ""

    textos = []
    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)
        for pagina in reader.pages:
            try:
                txt = pagina.extract_text() or ""
            except Exception:
                txt = ""
            textos.append(txt)
    return "\n".join(textos)


def _extrair_texto_pdfplumber(pdf_path: Path) -> str:
    try:
        import pdfplumber
    except Exception:
        return ""

    textos = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for pagina in pdf.pages:
                try:
                    txt = pagina.extract_text() or ""
                except Exception:
                    txt = ""
                textos.append(txt)
    except Exception:
        return ""
    return "\n".join(textos)


def extrair_texto_pdf(pdf_path: Path) -> str:
    texto1 = _extrair_texto_pypdf2(pdf_path)
    if texto1 and len(re.sub(r"\s+", "", texto1)) >= 20:
        return texto1

    texto2 = _extrair_texto_pdfplumber(pdf_path)
    if texto2 and len(re.sub(r"\s+", "", texto2)) > len(re.sub(r"\s+", "", texto1)):
        return texto2

    return texto1 or texto2 or ""


def extrair_cpf_do_nome_arquivo(nome_arquivo: str) -> str:
    candidatos = re.findall(r"\b\d{11}\b|\b\d{3}\.?\d{3}\.?\d{3}\-?\d{2}\b", nome_arquivo)
    for c in candidatos:
        cpf = normalizar_cpf(c)
        if validar_cpf(cpf):
            return cpf
    return ""


def extrair_identidade_secao_beneficiario(texto: str) -> Tuple[str, str]:
    if not texto:
        return "", ""

    match_secao = re.search(
        r"2\.\s*PESSOA\s+F[IÍ]SICA\s+BENEFICI[ÁA]RIA\s+DOS\s+RENDIMENTOS",
        texto,
        flags=re.IGNORECASE,
    )

    chunks = []
    if match_secao:
        ini = max(0, match_secao.start() - 250)
        meio = match_secao.end()
        fim = min(len(texto), match_secao.end() + 1200)
        chunks.append(texto[meio:fim])
        chunks.append(texto[ini:fim])
    else:
        chunks.append(texto)

    padroes = [
        (r"CPF\s*:?\s*([\d\.\-]{11,14}).{0,180}?NOME(?:\s+COMPLETO)?\s*:?\s*([A-ZÀ-Ú][A-ZÀ-Ú'\-\s]{5,})", "cpf_nome"),
        (r"NOME(?:\s+COMPLETO)?\s*:?\s*([A-ZÀ-Ú][A-ZÀ-Ú'\-\s]{5,}).{0,180}?CPF\s*:?\s*([\d\.\-]{11,14})", "nome_cpf"),
        (r"([\d\.\-]{11,14})\s*CPF\s*:?\s*([A-ZÀ-Ú][A-ZÀ-Ú'\-\s]{5,})\s*NOME", "cpf_nome_fechando"),
    ]

    for chunk in chunks:
        for padrao, modo in padroes:
            m = re.search(padrao, chunk, flags=re.IGNORECASE | re.DOTALL)
            if not m:
                continue

            if modo == "cpf_nome":
                cpf, nome = m.group(1), m.group(2)
            elif modo == "nome_cpf":
                nome, cpf = m.group(1), m.group(2)
            else:
                cpf, nome = m.group(1), m.group(2)

            cpf = normalizar_cpf(cpf)
            nome = limpar_nome_extraido(nome)
            if validar_cpf(cpf):
                return cpf, nome

    for chunk in chunks:
        linhas = [l.strip() for l in chunk.splitlines() if l.strip()]
        cpf, nome = "", ""
        for i, linha in enumerate(linhas):
            up = normalizar_texto_simples(linha)

            if up == "CPF" and i + 1 < len(linhas):
                candidato = normalizar_cpf(linhas[i + 1])
                if validar_cpf(candidato):
                    cpf = candidato

            if up in {"NOME", "NOME COMPLETO"} and i + 1 < len(linhas):
                prox = linhas[i + 1]
                prox_up = normalizar_texto_simples(prox)
                if prox_up not in {"CPF", "NOME", "NOME COMPLETO"} and len(prox_up) > 4:
                    nome = limpar_nome_extraido(prox)

        if cpf:
            return cpf, nome

    return "", ""


def encontrar_cpfs_no_texto(texto: str) -> List[Dict[str, object]]:
    resultados: List[Dict[str, object]] = []
    if not texto:
        return resultados

    padrao = re.compile(r"\b\d{3}\.?\d{3}\.?\d{3}\-?\d{2}\b|\b\d{11}\b")

    for match in padrao.finditer(texto):
        bruto = match.group(0)
        cpf = normalizar_cpf(bruto)

        if not validar_cpf(cpf):
            continue

        ini = max(0, match.start() - 160)
        fim = min(len(texto), match.end() + 160)
        contexto = texto[ini:fim]

        resultados.append(
            {
                "cpf": cpf,
                "cpf_bruto": bruto,
                "posicao": match.start(),
                "contexto": contexto,
            }
        )

    vistos = set()
    unicos: List[Dict[str, object]] = []
    for item in resultados:
        cpf = str(item["cpf"])
        if cpf not in vistos:
            vistos.add(cpf)
            unicos.append(item)

    return unicos


def pontuar_contexto_cpf(contexto: str) -> int:
    contexto_up = normalizar_texto_simples(contexto)
    score = 0

    termos_fortes = {
        "CPF": 20,
        "NOME": 10,
        "BENEFICIARIA": 12,
        "BENEFICIARIO": 12,
        "PESSOA FISICA": 12,
        "RENDIMENTOS": 8,
        "COLABORADOR": 8,
        "FUNCIONARIO": 8,
        "FUNCIONARIOA": 8,
        "TITULAR": 8,
    }

    termos_negativos = {
        "RESPONSAVEL": -15,
        "CONTADOR": -15,
        "REPRESENTANTE": -12,
        "PROCURADOR": -10,
        "FONTE PAGADORA": -6,
        "CNPJ": -6,
    }

    for termo, peso in termos_fortes.items():
        if termo in contexto_up:
            score += peso

    for termo, peso in termos_negativos.items():
        if termo in contexto_up:
            score += peso

    if re.search(r"NOME\s*:?\s*.*?CPF", contexto, flags=re.IGNORECASE | re.DOTALL):
        score += 30
    if re.search(r"CPF\s*:?\s*\d", contexto, flags=re.IGNORECASE):
        score += 20
    if re.search(r"2\.?\s*PESSOA\s+F[IÍ]SICA\s+BENEFICI[ÁA]RIA", contexto, flags=re.IGNORECASE):
        score += 30

    return score


def escolher_cpf_mais_provavel(texto: str) -> Optional[Dict[str, object]]:
    candidatos = encontrar_cpfs_no_texto(texto)
    if not candidatos:
        return None

    for item in candidatos:
        item["score"] = pontuar_contexto_cpf(str(item["contexto"]))

    candidatos.sort(key=lambda x: (int(x["score"]), -int(x["posicao"])), reverse=True)
    return candidatos[0]


def extrair_nome_proximo_ao_cpf(texto: str, cpf: str) -> str:
    if not texto or not cpf:
        return ""

    cpf_fmt = formatar_cpf(cpf)
    padroes = [
        rf"Nome\s*:?\s*([A-ZÀ-Úa-zà-ú'\-\s]+?)\s+CPF\s*:?\s*{re.escape(cpf_fmt)}",
        rf"Nome\s*:?\s*([A-ZÀ-Úa-zà-ú'\-\s]+?)\s+CPF\s*:?\s*{re.escape(cpf)}",
        rf"Nome\s*:?\s*([A-ZÀ-Úa-zà-ú'\-\s]+?)\s+CPF\s*:?\s*\d{{3}}\.?\d{{3}}\.?\d{{3}}\-?\d{{2}}",
        rf"NOME COMPLETO\s*:?\s*([A-ZÀ-Úa-zà-ú'\-\s]+?)\s+CPF",
    ]

    for padrao in padroes:
        m = re.search(padrao, texto, flags=re.IGNORECASE | re.DOTALL)
        if m:
            nome = " ".join(m.group(1).split()).strip(" -:")
            if nome:
                return limpar_nome_extraido(nome)

    indice = texto.find(cpf_fmt)
    if indice < 0:
        indice = texto.find(cpf)

    if indice >= 0:
        trecho_ini = max(0, indice - 200)
        trecho = texto[trecho_ini:indice + 50]
        linhas = [l.strip() for l in trecho.splitlines() if l.strip()]
        for linha in reversed(linhas):
            linha_up = normalizar_texto_simples(linha)
            if "NOME" in linha_up and "CPF" not in linha_up:
                candidato = re.sub(r"^.*?NOME\s*:?", "", linha, flags=re.IGNORECASE).strip(" -:")
                candidato = " ".join(candidato.split())
                if candidato:
                    return candidato

    return ""


def identificar_pdf(pdf_path: Path) -> IdentificacaoPDF:
    texto = extrair_texto_pdf(pdf_path)
    cpf_nome_arquivo = extrair_cpf_do_nome_arquivo(pdf_path.name)

    cpf_secao, nome_secao = extrair_identidade_secao_beneficiario(texto)
    if cpf_secao:
        divergencia = bool(cpf_nome_arquivo and cpf_secao and cpf_nome_arquivo != cpf_secao)
        return IdentificacaoPDF(
            cpf_encontrado=cpf_secao,
            nome_encontrado=nome_secao,
            origem_cpf="secao_beneficiario",
            cpf_nome_arquivo=cpf_nome_arquivo,
            divergencia_nome_arquivo=divergencia,
            total_cpfs_validos=len(encontrar_cpfs_no_texto(texto)),
            texto_extraido=texto,
        )

    melhor = escolher_cpf_mais_provavel(texto)
    if melhor:
        cpf_texto = str(melhor["cpf"])
        nome_texto = extrair_nome_proximo_ao_cpf(texto, cpf_texto)
        origem = "conteudo_pdf"
        divergencia = bool(cpf_nome_arquivo and cpf_texto and cpf_nome_arquivo != cpf_texto)
        return IdentificacaoPDF(
            cpf_encontrado=cpf_texto,
            nome_encontrado=nome_texto,
            origem_cpf=origem,
            cpf_nome_arquivo=cpf_nome_arquivo,
            divergencia_nome_arquivo=divergencia,
            total_cpfs_validos=len(encontrar_cpfs_no_texto(texto)),
            texto_extraido=texto,
        )

    if cpf_nome_arquivo:
        return IdentificacaoPDF(
            cpf_encontrado=cpf_nome_arquivo,
            nome_encontrado="",
            origem_cpf="nome_arquivo_fallback",
            cpf_nome_arquivo=cpf_nome_arquivo,
            divergencia_nome_arquivo=False,
            total_cpfs_validos=0,
            texto_extraido=texto,
        )

    return IdentificacaoPDF(
        cpf_encontrado="",
        nome_encontrado="",
        origem_cpf="nao_encontrado",
        cpf_nome_arquivo=cpf_nome_arquivo,
        divergencia_nome_arquivo=False,
        total_cpfs_validos=0,
        texto_extraido=texto,
    )


def _normalizar_cabecalho(cab: str) -> str:
    return normalizar_texto_simples(cab).replace(" ", "")


def _mapear_colunas(df: pd.DataFrame) -> Dict[str, str]:
    aliases = {
        "cpf": ["CPF"],
        "matricula": ["Matrícula", "Matricula", "Matr", "Registro"],
        "nome": ["Nome", "Nome Completo", "Colaborador", "Funcionario", "Funcionário"],
        "email": ["Email Alternativo", "Email", "E-mail", "Email Pessoal", "E-mail Alternativo"],
    }

    colunas_norm = {_normalizar_cabecalho(c): c for c in df.columns}
    resolvidas: Dict[str, str] = {}

    for chave, possiveis in aliases.items():
        encontrado = None
        for nome in possiveis:
            alvo = _normalizar_cabecalho(nome)
            if alvo in colunas_norm:
                encontrado = colunas_norm[alvo]
                break
        if not encontrado:
            raise ValueError(
                "Não foi possível localizar as colunas obrigatórias na base. "
                f"Coluna ausente: {chave}. Colunas encontradas: {', '.join(map(str, df.columns))}"
            )
        resolvidas[chave] = encontrado

    return resolvidas


def ler_base_colaboradores_arquivo(arquivo_path: Path) -> Dict[str, Colaborador]:
    ext = arquivo_path.suffix.lower()

    if ext == ".csv":
        df = pd.read_csv(
            arquivo_path,
            dtype=str,
            sep=None,
            engine="python",
            encoding="utf-8",
            keep_default_na=False,
        )
    elif ext in [".xls", ".xlsx", ".xlsm"]:
        df = pd.read_excel(arquivo_path, dtype=str).fillna("")
    elif ext == ".xlsb":
        df = pd.read_excel(arquivo_path, dtype=str, engine="pyxlsb").fillna("")
    else:
        raise ValueError("Formato não suportado. Use CSV/XLS/XLSX/XLSM/XLSB.")

    return _converter_dataframe_em_base(df)


def ler_base_colaboradores_texto(texto: str) -> Dict[str, Colaborador]:
    linhas = [linha.rstrip() for linha in texto.splitlines() if linha.strip()]
    if not linhas:
        raise ValueError("Nenhum dado foi colado na área de texto.")

    delimitador = "\t" if "\t" in linhas[0] else ";"
    cabecalhos = [c.strip() for c in linhas[0].split(delimitador)]
    dados = []

    for linha in linhas[1:]:
        partes = [p.strip() for p in linha.split(delimitador)]
        if len(partes) < len(cabecalhos):
            partes += [""] * (len(cabecalhos) - len(partes))
        dados.append(partes[: len(cabecalhos)])

    df = pd.DataFrame(dados, columns=cabecalhos)
    return _converter_dataframe_em_base(df)


def _converter_dataframe_em_base(df: pd.DataFrame) -> Dict[str, Colaborador]:
    if df.empty:
        raise ValueError("A base está vazia.")

    mapa = _mapear_colunas(df)
    base: Dict[str, Colaborador] = {}

    for _, row in df.fillna("").iterrows():
        cpf = normalizar_cpf(row.get(mapa["cpf"], ""))
        if not cpf:
            continue

        base[cpf] = Colaborador(
            matricula=str(row.get(mapa["matricula"], "")).strip(),
            cpf=cpf,
            nome=str(row.get(mapa["nome"], "")).strip(),
            email=str(row.get(mapa["email"], "")).strip(),
        )

    if not base:
        raise ValueError("Nenhum CPF válido foi identificado na base.")

    return base


def obter_base(base_path: Optional[Path], base_texto: str) -> Dict[str, Colaborador]:
    if base_texto and base_texto.strip():
        return ler_base_colaboradores_texto(base_texto)
    if base_path:
        return ler_base_colaboradores_arquivo(base_path)
    raise ValueError("Informe uma base por arquivo ou cole os dados diretamente na tela.")


def proteger_pdf_com_senha(pdf_path: Path, senha: str, saida_dir: Path) -> Path:
    try:
        import pikepdf
    except Exception as e:
        raise RuntimeError("pikepdf não está instalado. Instale com: pip install pikepdf") from e

    garantir_pasta(saida_dir)

    nome_base = extrair_nome_do_arquivo_sem_cpf(pdf_path.name, extrair_cpf_do_nome_arquivo(pdf_path.name)) or pdf_path.stem
    nome_base = normalizar_nome_arquivo_manual(nome_base) or normalizar_nome_arquivo_manual(pdf_path.stem) or "DOCUMENTO"

    out_name = f"{senha}_{nome_base}_protegido.pdf"
    out_name = reduzir_nome_para_caminho(saida_dir, out_name)
    out_path = saida_dir / out_name

    contador = 1
    while out_path.exists():
        out_name = reduzir_nome_para_caminho(saida_dir, f"{senha}_{nome_base}_protegido_{contador}.pdf")
        out_path = saida_dir / out_name
        contador += 1

    pdf_in_path = caminho_windows_estendido(pdf_path)
    pdf_out_path = caminho_windows_estendido(out_path)

    with pikepdf.open(pdf_in_path) as pdf:
        pdf.save(pdf_out_path, encryption=pikepdf.Encryption(owner=senha, user=senha, R=4))

    return out_path


def outlook_criar_rascunho_sem_exibir(
    para: str,
    assunto: str,
    corpo: str,
    anexo_path: Path,
    display_name: str,
) -> None:
    try:
        import win32com.client as win32
    except Exception as e:
        raise RuntimeError(
            "win32com.client não está disponível. Instale pywin32 para integração com Outlook."
        ) from e

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = para
    mail.Subject = assunto
    mail.Body = corpo
    mail.Attachments.Add(str(anexo_path.resolve()), 1, 1, display_name)
    mail.Save()


def escrever_log_csv(log_path: Path, resultados: List[ResultadoProcesso]) -> None:
    garantir_pasta(log_path.parent)
    existe = log_path.exists()

    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        if not existe:
            writer.writerow(
                [
                    "timestamp",
                    "pdf_original",
                    "cpf_pdf",
                    "nome_pdf",
                    "cpf_nome_arquivo",
                    "origem_cpf",
                    "encontrado_base",
                    "matricula",
                    "email",
                    "pdf_protegido",
                    "status",
                    "detalhe",
                ]
            )

        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        for r in resultados:
            writer.writerow(
                [
                    ts,
                    r.pdf_original,
                    r.cpf_pdf,
                    r.nome_pdf,
                    r.cpf_nome_arquivo,
                    r.origem_cpf,
                    "SIM" if r.encontrado_base else "NAO",
                    r.matricula,
                    r.email,
                    r.pdf_protegido,
                    r.status,
                    r.detalhe,
                ]
            )


def resumir_resultados(resultados: List[ResultadoProcesso], modo: str) -> str:
    total = len(resultados)
    ok = sum(1 for r in resultados if r.status == "OK")
    pulado = sum(1 for r in resultados if r.status == "PULADO")
    falha = sum(1 for r in resultados if r.status == "FALHA")

    linhas = [
        f"{APP_NAME} v{VERSAO}",
        f"Modo: {modo}",
        f"Total PDFs: {total}",
        f"OK: {ok}",
        f"PULADOS: {pulado}",
        f"FALHAS: {falha}",
    ]
    return "\n".join(linhas)


def listar_pdfs(origem: Path) -> List[Path]:
    if origem.is_file():
        if origem.suffix.lower() != ".pdf":
            raise ValueError("O arquivo selecionado não é um PDF.")
        return [origem]

    if origem.is_dir():
        pdfs = sorted([p for p in origem.glob("*.pdf") if p.is_file()])
        if not pdfs:
            raise ValueError("Nenhum PDF encontrado na pasta selecionada.")
        return pdfs

    raise ValueError("Origem inválida para PDFs.")


def localizar_colaborador(base: Dict[str, Colaborador], cpf: str) -> Optional[Colaborador]:
    cpf = normalizar_cpf(cpf)
    if not cpf:
        return None
    return base.get(cpf)


def processar_arquivos(
    origem_pdf: Path,
    base_path: Optional[Path] = None,
    base_texto: str = "",
    criar_rascunho: bool = True,
    on_progress: Optional[Callable[[int, int], None]] = None,
    on_log: Optional[Callable[[str], None]] = None,
    subject_template: Optional[str] = None,
    body_template: Optional[str] = None,
) -> Tuple[List[ResultadoProcesso], Path]:
    pdfs = listar_pdfs(origem_pdf)
    base: Dict[str, Colaborador] = {}

    if criar_rascunho:
        base = obter_base(base_path, base_texto)

    pasta_saida = origem_pdf.parent if origem_pdf.is_file() else origem_pdf
    saida_dir = pasta_saida / SUBPASTA_PROTEGIDOS
    garantir_pasta(saida_dir)

    assunto_template_final = subject_template or SUBJECT_TEMPLATE
    corpo_template_final = body_template or BODY_TEMPLATE

    resultados: List[ResultadoProcesso] = []
    total = len(pdfs)

    for i, pdf in enumerate(pdfs, start=1):
        if on_progress:
            on_progress(i, total)

        try:
            identificacao = identificar_pdf(pdf)
        except Exception as e:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf="",
                nome_pdf="",
                cpf_nome_arquivo="",
                origem_cpf="erro_extracao",
                encontrado_base=False,
                matricula="",
                email="",
                pdf_protegido="",
                status="FALHA",
                detalhe=f"Erro ao ler PDF: {e}",
            )
            resultados.append(r)
            if on_log:
                on_log(f"FALHA: {pdf.name} | {r.detalhe}")
            continue

        cpf_pdf = normalizar_cpf(identificacao.cpf_encontrado)
        nome_pdf = identificacao.nome_encontrado
        detalhe_base = []

        if identificacao.divergencia_nome_arquivo:
            detalhe_base.append(
                f"CPF do conteúdo diverge do nome do arquivo ({formatar_cpf(identificacao.cpf_nome_arquivo)})."
            )

        if not cpf_pdf:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf="",
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=False,
                matricula="",
                email="",
                pdf_protegido="",
                status="FALHA",
                detalhe="CPF não encontrado no conteúdo do PDF nem no nome do arquivo.",
            )
            resultados.append(r)
            if on_log:
                on_log(f"FALHA: {pdf.name} | {r.detalhe}")
            continue

        if not criar_rascunho:
            detalhe = f"CPF identificado por {identificacao.origem_cpf}."
            if nome_pdf:
                detalhe += f" Nome lido: {nome_pdf}."
            if detalhe_base:
                detalhe += " " + " ".join(detalhe_base)

            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=cpf_pdf,
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=False,
                matricula="",
                email="",
                pdf_protegido="",
                status="OK",
                detalhe=detalhe,
            )
            resultados.append(r)
            if on_log:
                on_log(
                    f"OK: {pdf.name} | CPF={formatar_cpf(cpf_pdf)} | origem={identificacao.origem_cpf}"
                )
            continue

        colab = localizar_colaborador(base, cpf_pdf)
        if not colab:
            detalhe = "CPF encontrado no PDF, mas não está na base informada."
            if detalhe_base:
                detalhe += " " + " ".join(detalhe_base)
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=cpf_pdf,
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=False,
                matricula="",
                email="",
                pdf_protegido="",
                status="PULADO",
                detalhe=detalhe,
            )
            resultados.append(r)
            if on_log:
                on_log(f"PULADO: {pdf.name} | CPF={formatar_cpf(cpf_pdf)} não localizado na base")
            continue

        if not colab.email:
            detalhe = "Email vazio na base para o CPF localizado."
            if detalhe_base:
                detalhe += " " + " ".join(detalhe_base)
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=cpf_pdf,
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=True,
                matricula=colab.matricula,
                email="",
                pdf_protegido="",
                status="FALHA",
                detalhe=detalhe,
            )
            resultados.append(r)
            if on_log:
                on_log(f"FALHA: {pdf.name} | e-mail vazio na base")
            continue

        if not validar_email_basico(colab.email):
            detalhe = f"E-mail inválido na base: {colab.email}"
            if detalhe_base:
                detalhe += " " + " ".join(detalhe_base)
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=cpf_pdf,
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=True,
                matricula=colab.matricula,
                email=colab.email,
                pdf_protegido="",
                status="FALHA",
                detalhe=detalhe,
            )
            resultados.append(r)
            if on_log:
                on_log(f"FALHA: {pdf.name} | e-mail inválido | {colab.email}")
            continue

        try:
            ano_base = extrair_ano_base_do_nome_arquivo(pdf.name)
            assunto = assunto_template_final.format(
                ano_base=ano_base,
                matricula=(colab.matricula or "N/A"),
                nome=(colab.nome or nome_pdf or "Colaborador(a)"),
                url_chamado=RH_CHAMADO_URL,
            )
            corpo = corpo_template_final.format(
                nome=(colab.nome or nome_pdf or "Colaborador(a)"),
                ano_base=ano_base,
                matricula=(colab.matricula or "N/A"),
                url_chamado=RH_CHAMADO_URL,
            )

            protegido = proteger_pdf_com_senha(pdf, cpf_pdf, saida_dir)
            display_name = limpar_nome_anexo_removendo_cpf(pdf.name, cpf_pdf)

            outlook_criar_rascunho_sem_exibir(
                para=colab.email,
                assunto=assunto,
                corpo=corpo,
                anexo_path=protegido,
                display_name=display_name,
            )

            detalhe = "Rascunho criado com PDF protegido."
            if detalhe_base:
                detalhe += " " + " ".join(detalhe_base)

            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=cpf_pdf,
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=True,
                matricula=colab.matricula,
                email=colab.email,
                pdf_protegido=protegido.name,
                status="OK",
                detalhe=detalhe,
            )
            resultados.append(r)
            if on_log:
                on_log(
                    f"OK: {pdf.name} | CPF={formatar_cpf(cpf_pdf)} | email={colab.email} | rascunho salvo"
                )

        except Exception as e:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=cpf_pdf,
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=True,
                matricula=colab.matricula,
                email=colab.email,
                pdf_protegido="",
                status="FALHA",
                detalhe=f"Erro ao proteger PDF ou criar rascunho: {e}",
            )
            resultados.append(r)
            if on_log:
                on_log(f"FALHA: {pdf.name} | {r.detalhe}")

    log_path = pasta_saida / LOG_NAME
    escrever_log_csv(log_path, resultados)
    return resultados, log_path


if CTK_AVAILABLE:

    class App(ctk.CTk):
        def __init__(self):
            super().__init__()

            ctk.set_appearance_mode("System")
            ctk.set_default_color_theme("blue")

            self.title(f"{APP_NAME} | v{VERSAO}")
            self.geometry("1380x980")
            self.minsize(1280, 900)

            self.origem_pdf: Optional[Path] = None
            self.base_path: Optional[Path] = None
            self.templates_email = json.loads(json.dumps(TEMPLATES_EMAIL))
            self.tipo_documento_atual = "Informe de Rendimento"

            self.ren_pdfs: List[Path] = []
            self.ren_index: int = 0
            self.ren_pdf_atual: Optional[Path] = None
            self.ren_cpf_atual: str = ""

            self._montar_ui()
            self._carregar_template_na_tela()

        def _montar_ui(self):
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(4, weight=1)

            frame_top = ctk.CTkFrame(self)
            frame_top.grid(row=0, column=0, padx=16, pady=(14, 8), sticky="ew")
            frame_top.grid_columnconfigure(0, weight=1)

            lbl_titulo = ctk.CTkLabel(
                frame_top,
                text=f"{APP_NAME} | v{VERSAO}",
                font=ctk.CTkFont(size=22, weight="bold"),
            )
            lbl_titulo.grid(row=0, column=0, padx=14, pady=(12, 2), sticky="w")

            lbl_sub = ctk.CTkLabel(
                frame_top,
                text=(
                    "Leitura universal de PDF para encontrar CPF no conteúdo, com suporte a PDF único "
                    "ou pasta inteira, proteção por senha e criação de rascunho no Outlook."
                ),
                font=ctk.CTkFont(size=13),
                justify="left",
            )
            lbl_sub.grid(row=1, column=0, padx=14, pady=(0, 12), sticky="w")

            frame_origem = ctk.CTkFrame(self)
            frame_origem.grid(row=1, column=0, padx=16, pady=8, sticky="ew")
            frame_origem.grid_columnconfigure(1, weight=1)

            ctk.CTkLabel(frame_origem, text="Origem PDF").grid(row=0, column=0, padx=12, pady=10, sticky="w")

            self.ent_origem = ctk.CTkEntry(frame_origem)
            self.ent_origem.grid(row=0, column=1, padx=12, pady=10, sticky="ew")
            self._set_entry_readonly(self.ent_origem, "")

            ctk.CTkButton(frame_origem, text="Selecionar Pasta", width=140, command=self._selecionar_pasta).grid(
                row=0, column=2, padx=6, pady=10
            )
            ctk.CTkButton(frame_origem, text="Selecionar PDF", width=140, command=self._selecionar_pdf).grid(
                row=0, column=3, padx=(6, 12), pady=10
            )

            frame_base = ctk.CTkFrame(self)
            frame_base.grid(row=2, column=0, padx=16, pady=8, sticky="ew")
            frame_base.grid_columnconfigure(1, weight=1)
            frame_base.grid_rowconfigure(2, weight=1)

            ctk.CTkLabel(frame_base, text="Base por arquivo").grid(row=0, column=0, padx=12, pady=10, sticky="w")

            self.ent_base = ctk.CTkEntry(frame_base)
            self.ent_base.grid(row=0, column=1, padx=12, pady=10, sticky="ew")
            self._set_entry_readonly(self.ent_base, "")

            ctk.CTkButton(frame_base, text="Selecionar Base", width=140, command=self._selecionar_base).grid(
                row=0, column=2, padx=(6, 12), pady=10
            )

            ctk.CTkLabel(
                frame_base,
                text="Ou cole a base abaixo. Cabeçalhos aceitos: CPF, Matrícula/Matricula, Nome, Email/Email Alternativo",
            ).grid(row=1, column=0, columnspan=3, padx=12, pady=(0, 6), sticky="w")

            self.txt_base = ctk.CTkTextbox(frame_base, height=130)
            self.txt_base.grid(row=2, column=0, columnspan=3, padx=12, pady=(0, 12), sticky="ew")
            self.txt_base.insert(
                "1.0",
                "CPF;Matrícula;Nome;Email Alternativo\n"
                "12345678900;123456;Nome e Sobrenome;email@exemplo.com\n"
            )

            frame_msg = ctk.CTkFrame(self)
            frame_msg.grid(row=3, column=0, padx=16, pady=8, sticky="ew")
            frame_msg.grid_columnconfigure(1, weight=1)
            frame_msg.grid_rowconfigure(2, weight=1)

            ctk.CTkLabel(frame_msg, text="Tipo do documento").grid(row=0, column=0, padx=12, pady=10, sticky="w")
            self.cmb_tipo_documento = ctk.CTkComboBox(
                frame_msg,
                values=list(self.templates_email.keys()),
                command=self._ao_mudar_tipo_documento,
                width=260
            )
            self.cmb_tipo_documento.grid(row=0, column=1, padx=12, pady=10, sticky="w")
            self.cmb_tipo_documento.set(self.tipo_documento_atual)

            ctk.CTkButton(frame_msg, text="Restaurar Modelo", width=150, command=self._restaurar_modelo_padrao).grid(
                row=0, column=2, padx=6, pady=10
            )
            ctk.CTkButton(frame_msg, text="Salvar Modelo Atual", width=170, command=self._salvar_modelo_atual).grid(
                row=0, column=3, padx=(6, 12), pady=10
            )

            ctk.CTkLabel(frame_msg, text="Assunto").grid(row=1, column=0, padx=12, pady=6, sticky="w")
            self.ent_assunto = ctk.CTkEntry(frame_msg)
            self.ent_assunto.grid(row=1, column=1, columnspan=3, padx=12, pady=6, sticky="ew")

            ctk.CTkLabel(frame_msg, text="Corpo do e-mail").grid(row=2, column=0, padx=12, pady=(6, 12), sticky="nw")
            self.txt_corpo = ctk.CTkTextbox(frame_msg, height=170)
            self.txt_corpo.grid(row=2, column=1, columnspan=3, padx=12, pady=(6, 12), sticky="ew")

            frame_exec = ctk.CTkFrame(self)
            frame_exec.grid(row=4, column=0, padx=16, pady=8, sticky="nsew")
            frame_exec.grid_columnconfigure(0, weight=1)
            frame_exec.grid_columnconfigure(1, weight=1)
            frame_exec.grid_rowconfigure(2, weight=1)

            botoes = ctk.CTkFrame(frame_exec)
            botoes.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 8), sticky="ew")
            botoes.grid_columnconfigure(6, weight=1)

            self.btn_localizar = ctk.CTkButton(
                botoes,
                text="Encontrar CPF(s)",
                width=180,
                command=self._iniciar_localizacao,
            )
            self.btn_localizar.grid(row=0, column=0, padx=(0, 10), pady=8, sticky="w")

            self.btn_iniciar = ctk.CTkButton(
                botoes,
                text="Proteger + Criar Rascunho",
                width=220,
                command=self._iniciar_rascunho,
            )
            self.btn_iniciar.grid(row=0, column=1, padx=10, pady=8, sticky="w")

            self.btn_validar_base = ctk.CTkButton(
                botoes,
                text="Validar Base Colada",
                width=180,
                command=self._validar_base_colada,
            )
            self.btn_validar_base.grid(row=0, column=2, padx=10, pady=8, sticky="w")

            self.btn_reiniciar = ctk.CTkButton(
                botoes,
                text="Reiniciar Processo",
                width=170,
                command=self._reiniciar,
            )
            self.btn_reiniciar.grid(row=0, column=3, padx=10, pady=8, sticky="w")

            self.btn_template = ctk.CTkButton(
                botoes,
                text="Gerar Template Excel",
                width=190,
                command=self._gerar_template_excel
            )
            self.btn_template.grid(row=0, column=4, padx=10, pady=8, sticky="w")

            self.btn_abrir_renomeador = ctk.CTkButton(
                botoes,
                text="Abrir Renomeador",
                width=170,
                command=self._abrir_renomeador
            )
            self.btn_abrir_renomeador.grid(row=0, column=5, padx=10, pady=8, sticky="w")

            self.progress = ctk.CTkProgressBar(frame_exec)
            self.progress.grid(row=1, column=0, padx=12, pady=(2, 6), sticky="ew")
            self.progress.set(0)

            self.lbl_prog = ctk.CTkLabel(frame_exec, text="0/0")
            self.lbl_prog.grid(row=1, column=0, padx=12, pady=(2, 6), sticky="e")

            ctk.CTkLabel(frame_exec, text="Log do processamento", font=ctk.CTkFont(size=14, weight="bold")).grid(
                row=2, column=0, padx=12, pady=(4, 4), sticky="nw"
            )
            self.txt_log = ctk.CTkTextbox(frame_exec)
            self.txt_log.grid(row=2, column=0, padx=12, pady=(30, 12), sticky="nsew")
            self.txt_log.configure(state="disabled")

            frame_ren = ctk.CTkFrame(frame_exec)
            frame_ren.grid(row=2, column=1, padx=12, pady=(4, 12), sticky="nsew")
            frame_ren.grid_columnconfigure(0, weight=1)
            frame_ren.grid_rowconfigure(5, weight=1)

            ctk.CTkLabel(frame_ren, text="Renomeador manual", font=ctk.CTkFont(size=14, weight="bold")).grid(
                row=0, column=0, padx=12, pady=(10, 6), sticky="w"
            )

            self.lbl_ren_status = ctk.CTkLabel(frame_ren, text="0/0")
            self.lbl_ren_status.grid(row=0, column=0, padx=12, pady=(10, 6), sticky="e")

            self.lbl_ren_arquivo = ctk.CTkLabel(frame_ren, text="Arquivo atual: -", justify="left")
            self.lbl_ren_arquivo.grid(row=1, column=0, padx=12, pady=4, sticky="w")

            self.lbl_ren_cpf = ctk.CTkLabel(frame_ren, text="CPF encontrado: -", justify="left")
            self.lbl_ren_cpf.grid(row=2, column=0, padx=12, pady=4, sticky="w")

            self.lbl_ren_sugestao = ctk.CTkLabel(frame_ren, text="Nome sugerido: -", justify="left")
            self.lbl_ren_sugestao.grid(row=3, column=0, padx=12, pady=4, sticky="w")

            frame_ren_btn = ctk.CTkFrame(frame_ren)
            frame_ren_btn.grid(row=4, column=0, padx=12, pady=8, sticky="ew")
            frame_ren_btn.grid_columnconfigure(3, weight=1)

            self.btn_ren_anterior = ctk.CTkButton(frame_ren_btn, text="Anterior", width=120, command=self._ren_pdf_anterior)
            self.btn_ren_anterior.grid(row=0, column=0, padx=(0, 8), pady=8)

            self.btn_ren_pular = ctk.CTkButton(frame_ren_btn, text="Pular", width=120, command=self._ren_pular_pdf)
            self.btn_ren_pular.grid(row=0, column=1, padx=8, pady=8)

            self.btn_ren_renomear = ctk.CTkButton(frame_ren_btn, text="Renomear + Próximo", width=170, command=self._ren_renomear_selecionado)
            self.btn_ren_renomear.grid(row=0, column=2, padx=8, pady=8)

            self.txt_ren_texto = ctk.CTkTextbox(frame_ren, height=320)
            self.txt_ren_texto.grid(row=5, column=0, padx=12, pady=(6, 6), sticky="nsew")

            self.txt_ren_log = ctk.CTkTextbox(frame_ren, height=120)
            self.txt_ren_log.grid(row=6, column=0, padx=12, pady=(6, 12), sticky="ew")

            self.txt_ren_texto.bind("<Return>", self._ren_enter_renomear)
            self.txt_ren_texto.bind("<KP_Enter>", self._ren_enter_renomear)
            self.txt_ren_texto.bind("<Shift-Return>", self._ren_shift_enter_pular)
            self.txt_ren_texto.bind("<Shift-KP_Enter>", self._ren_shift_enter_pular)
            self.txt_ren_texto.bind("<ButtonRelease-1>", self._ren_atualizar_sugestao)
            self.txt_ren_texto.bind("<KeyRelease>", self._ren_atualizar_sugestao)

            footer = ctk.CTkLabel(self, text=FOOTER_UI, font=ctk.CTkFont(size=11))
            footer.grid(row=5, column=0, padx=16, pady=(0, 10), sticky="e")

            self._ren_atualizar_botoes()

        def _set_entry_readonly(self, entry: CTkEntry, value: str):
            entry.configure(state="normal")
            entry.delete(0, "end")
            entry.insert(0, value)
            entry.configure(state="readonly")

        def _selecionar_pasta(self):
            p = filedialog.askdirectory(title="Selecione a pasta com PDFs")
            if p:
                self.origem_pdf = Path(p)
                self._set_entry_readonly(self.ent_origem, str(self.origem_pdf))
                self._log(f"Pasta selecionada: {self.origem_pdf}")

        def _selecionar_pdf(self):
            p = filedialog.askopenfilename(
                title="Selecione um PDF",
                filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
            )
            if p:
                self.origem_pdf = Path(p)
                self._set_entry_readonly(self.ent_origem, str(self.origem_pdf))
                self._log(f"PDF selecionado: {self.origem_pdf}")

        def _selecionar_base(self):
            p = filedialog.askopenfilename(
                title="Selecione a base",
                filetypes=[("Bases suportadas", "*.csv *.xls *.xlsx *.xlsm *.xlsb"), ("Todos", "*.*")],
            )
            if p:
                self.base_path = Path(p)
                self._set_entry_readonly(self.ent_base, str(self.base_path))
                self._log(f"Base selecionada: {self.base_path}")

        def _log(self, msg: str):
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", msg + "\n")
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")

        def _limpar_log(self):
            self.txt_log.configure(state="normal")
            self.txt_log.delete("1.0", "end")
            self.txt_log.configure(state="disabled")

        def _ren_log(self, msg: str):
            self.txt_ren_log.configure(state="normal")
            self.txt_ren_log.insert("end", msg + "\n")
            self.txt_ren_log.see("end")
            self.txt_ren_log.configure(state="disabled")

        def _ren_limpar_log(self):
            self.txt_ren_log.configure(state="normal")
            self.txt_ren_log.delete("1.0", "end")
            self.txt_ren_log.configure(state="disabled")
            if hasattr(self, "janela_renomeador") and self.janela_renomeador and self.janela_renomeador.winfo_exists():
                self.after(10, self._ren_janela_atualizar_conteudo)

        def _on_progress(self, atual: int, total: int):
            if total <= 0:
                self.progress.set(0)
                self.lbl_prog.configure(text="0/0")
                return
            self.progress.set(atual / total)
            self.lbl_prog.configure(text=f"{atual}/{total}")
            self.update_idletasks()

        def _texto_base(self) -> str:
            texto = self.txt_base.get("1.0", "end").strip()
            exemplo = (
                "CPF;Matrícula;Nome;Email Alternativo\n"
                "12345678900;123456;Nome e Sobrenome;email@exemplo.com\n"
            )
            if texto == exemplo:
                return ""
            return texto

        def _ao_mudar_tipo_documento(self, escolha: str):
            self._salvar_modelo_atual(silencioso=True)
            self.tipo_documento_atual = escolha
            self._carregar_template_na_tela()

        def _carregar_template_na_tela(self):
            template = self.templates_email.get(self.tipo_documento_atual, TEMPLATES_EMAIL["Informe de Rendimento"])
            self.ent_assunto.delete(0, "end")
            self.ent_assunto.insert(0, template["assunto"])
            self.txt_corpo.delete("1.0", "end")
            self.txt_corpo.insert("1.0", template["corpo"])

        def _salvar_modelo_atual(self, silencioso: bool = False):
            assunto = self.ent_assunto.get().strip()
            corpo = self.txt_corpo.get("1.0", "end").strip()
            self.templates_email[self.tipo_documento_atual] = {
                "assunto": assunto,
                "corpo": corpo,
            }
            if not silencioso:
                self._log(f"Modelo salvo em memória para o tipo: {self.tipo_documento_atual}")
                messagebox.showinfo("Modelo salvo", f"Modelo atualizado para: {self.tipo_documento_atual}")

        def _restaurar_modelo_padrao(self):
            self.templates_email[self.tipo_documento_atual] = json.loads(json.dumps(TEMPLATES_EMAIL[self.tipo_documento_atual]))
            self._carregar_template_na_tela()
            self._log(f"Modelo padrão restaurado para: {self.tipo_documento_atual}")

        def _validar_base_colada(self):
            try:
                texto = self._texto_base()
                base = obter_base(self.base_path, texto)
                self._log(f"Base validada com sucesso. Total de CPFs carregados: {len(base)}")
                messagebox.showinfo("Base validada", f"Base validada com sucesso.\n\nTotal de CPFs: {len(base)}")
            except Exception as e:
                self._log(f"Erro ao validar base: {e}")
                messagebox.showerror("Erro na base", str(e))

        def _alternar_botoes(self, habilitar: bool):
            estado = "normal" if habilitar else "disabled"
            self.btn_localizar.configure(state=estado)
            self.btn_iniciar.configure(state=estado)
            self.btn_validar_base.configure(state=estado)
            self.btn_reiniciar.configure(state=estado)
            self.btn_template.configure(state=estado)
            self.btn_abrir_renomeador.configure(state=estado)

        def _reiniciar(self):
            self.origem_pdf = None
            self.base_path = None
            self._set_entry_readonly(self.ent_origem, "")
            self._set_entry_readonly(self.ent_base, "")
            self.txt_base.delete("1.0", "end")
            self._limpar_log()
            self.progress.set(0)
            self.lbl_prog.configure(text="0/0")
            self._ren_reiniciar()
            self.cmb_tipo_documento.set("Informe de Rendimento")
            self.tipo_documento_atual = "Informe de Rendimento"
            self.templates_email = json.loads(json.dumps(TEMPLATES_EMAIL))
            self._carregar_template_na_tela()
            self._log("Tela reiniciada. Pronto para novo processamento.")

        def _ren_reiniciar(self):
            self.ren_pdfs = []
            self.ren_index = 0
            self.ren_pdf_atual = None
            self.ren_cpf_atual = ""
            self.txt_ren_texto.delete("1.0", "end")
            self._ren_limpar_log()
            self.lbl_ren_status.configure(text="0/0")
            self.lbl_ren_arquivo.configure(text="Arquivo atual: -")
            self.lbl_ren_cpf.configure(text="CPF encontrado: -")
            self.lbl_ren_sugestao.configure(text="Nome sugerido: -")
            self._ren_atualizar_botoes()

        def _gerar_template_excel(self):
            try:
                caminho = gerar_template_excel()

                self._log(f"Template Excel gerado: {caminho}")

                messagebox.showinfo(
                    "Template criado",
                    f"O template foi criado com sucesso.\n\n{caminho}"
                )

            except Exception as e:
                messagebox.showerror(
                    "Erro ao gerar template",
                    str(e)
                )

        def _executar(self, criar_rascunho: bool):
            if not self.origem_pdf:
                self._log("Erro: selecione uma pasta ou um PDF antes de iniciar.")
                messagebox.showerror("Erro", "Selecione uma pasta ou um PDF antes de iniciar.")
                return

            if criar_rascunho:
                texto = self._texto_base()
                if not self.base_path and not texto:
                    self._log("Erro: informe uma base por arquivo ou cole os dados na tela.")
                    messagebox.showerror(
                        "Erro",
                        "Para criar rascunhos, informe uma base por arquivo ou cole os dados diretamente na tela.",
                    )
                    return
            else:
                texto = self._texto_base()

            assunto_template = self.ent_assunto.get().strip()
            corpo_template = self.txt_corpo.get("1.0", "end").strip()

            self._alternar_botoes(False)
            self.progress.set(0)
            self.lbl_prog.configure(text="0/0")

            modo = "Proteger + Criar Rascunho" if criar_rascunho else "Encontrar CPF(s)"
            self._log("")
            self._log("Iniciando processamento...")
            self._log(f"Modo: {modo}")
            self._log(f"Tipo documento: {self.tipo_documento_atual}")
            self._log(f"Origem: {self.origem_pdf}")
            if criar_rascunho:
                self._log("Base usada: base colada na tela" if texto else f"Base usada: {self.base_path}")
            self._log("")

            def worker():
                try:
                    resultados, log_path = processar_arquivos(
                        origem_pdf=self.origem_pdf,
                        base_path=self.base_path,
                        base_texto=texto,
                        criar_rascunho=criar_rascunho,
                        on_progress=lambda a, t: self.after(0, self._on_progress, a, t),
                        on_log=lambda m: self.after(0, self._log, m),
                        subject_template=assunto_template,
                        body_template=corpo_template,
                    )

                    resumo = resumir_resultados(resultados, modo)
                    self.after(0, self._log, "")
                    self.after(0, self._log, resumo)
                    self.after(0, self._log, f"Log salvo em: {log_path}")
                    self.after(0, messagebox.showinfo, "Processo concluído", resumo + f"\n\nLog: {log_path}")

                except Exception as e:
                    msg = f"Ocorreu um erro: {e}\n\n{traceback.format_exc()}"
                    self.after(0, self._log, msg)
                    self.after(0, messagebox.showerror, "Erro", msg)
                finally:
                    self.after(0, self._alternar_botoes, True)

            threading.Thread(target=worker, daemon=True).start()

        def _iniciar_localizacao(self):
            self._executar(criar_rascunho=False)

        def _iniciar_rascunho(self):
            self._executar(criar_rascunho=True)

        def _abrir_renomeador(self):
            try:
                if not self.origem_pdf:
                    resposta = messagebox.askyesno(
                        "Abrir Renomeador",
                        "Nenhuma origem foi selecionada ainda.\n\nDeseja escolher uma pasta com PDFs agora?"
                    )
                    if not resposta:
                        return

                    p = filedialog.askdirectory(title="Selecione a pasta com PDFs para o renomeador")
                    if not p:
                        return

                    self.origem_pdf = Path(p)
                    self._set_entry_readonly(self.ent_origem, str(self.origem_pdf))
                    self._log(f"Pasta selecionada para o renomeador: {self.origem_pdf}")

                self.ren_pdfs = listar_pdfs(self.origem_pdf)
                self.ren_index = 0
                self.ren_pdf_atual = None
                self.ren_cpf_atual = ""
                self.txt_ren_texto.delete("1.0", "end")
                self._ren_limpar_log()
                self._ren_log(f"Renomeador carregado com {len(self.ren_pdfs)} PDF(s).")
                self._ren_carregar_pdf_atual()
                self._ren_atualizar_botoes()
                self._abrir_janela_renomeador()
                self._log(f"Renomeador aberto com {len(self.ren_pdfs)} PDF(s) em janela própria.")
            except Exception as e:
                self._ren_log(f"Erro ao abrir renomeador: {e}")
                self._log(f"Erro ao abrir renomeador: {e}")
                messagebox.showerror("Erro", str(e))

        def _ren_atualizar_botoes(self):
            total = len(self.ren_pdfs)
            atual_habilitado = total > 0 and self.ren_index < total
            anterior_habilitado = total > 0 and self.ren_index > 0

            self.btn_ren_anterior.configure(state="normal" if anterior_habilitado else "disabled")
            self.btn_ren_pular.configure(state="normal" if atual_habilitado else "disabled")
            self.btn_ren_renomear.configure(state="normal" if atual_habilitado else "disabled")
            if hasattr(self, "janela_renomeador") and self.janela_renomeador and self.janela_renomeador.winfo_exists():
                self.after(10, self._ren_janela_atualizar_botoes)

        def _ren_carregar_pdf_atual(self):
            if not self.ren_pdfs:
                self.ren_pdf_atual = None
                self.ren_cpf_atual = ""
                self.lbl_ren_status.configure(text="0/0")
                self.lbl_ren_arquivo.configure(text="Arquivo atual: -")
                self.lbl_ren_cpf.configure(text="CPF encontrado: -")
                self.lbl_ren_sugestao.configure(text="Nome sugerido: -")
                self.txt_ren_texto.delete("1.0", "end")
                self._ren_atualizar_botoes()
                self._ren_janela_atualizar_conteudo()
                return

            if self.ren_index >= len(self.ren_pdfs):
                self.ren_pdf_atual = None
                self.ren_cpf_atual = ""
                self.lbl_ren_status.configure(text=f"{len(self.ren_pdfs)}/{len(self.ren_pdfs)}")
                self.lbl_ren_arquivo.configure(text="Arquivo atual: -")
                self.lbl_ren_cpf.configure(text="CPF encontrado: -")
                self.lbl_ren_sugestao.configure(text="Nome sugerido: -")
                self.txt_ren_texto.delete("1.0", "end")
                self._ren_log("Todos os PDFs do renomeador foram percorridos.")
                self._ren_atualizar_botoes()
                self._ren_janela_atualizar_conteudo()
                return

            self.ren_pdf_atual = self.ren_pdfs[self.ren_index]

            texto = ""
            try:
                texto = extrair_texto_pdf(self.ren_pdf_atual)
            except Exception as e:
                self._ren_log(f"Erro ao extrair texto do PDF: {e}")

            self.txt_ren_texto.delete("1.0", "end")
            self.txt_ren_texto.insert("1.0", texto)

            try:
                identificacao = identificar_pdf(self.ren_pdf_atual)
                self.ren_cpf_atual = (
                    normalizar_cpf(identificacao.cpf_encontrado)
                    or extrair_cpf_do_nome_arquivo(self.ren_pdf_atual.name)
                    or encontrar_cpf_simples_no_texto(texto)
                    or "SEMCPF"
                )
            except Exception as e:
                self.ren_cpf_atual = (
                    extrair_cpf_do_nome_arquivo(self.ren_pdf_atual.name)
                    or encontrar_cpf_simples_no_texto(texto)
                    or "SEMCPF"
                )
                self._ren_log(f"Erro ao identificar PDF no renomeador: {e}")

            self.lbl_ren_status.configure(text=f"{self.ren_index + 1}/{len(self.ren_pdfs)}")
            self.lbl_ren_arquivo.configure(text=f"Arquivo atual: {self.ren_pdf_atual}")
            self.lbl_ren_cpf.configure(text=f"CPF encontrado: {self.ren_cpf_atual}")
            self.lbl_ren_sugestao.configure(text="Nome sugerido: -")
            self._ren_log(f"Abrindo PDF {self.ren_index + 1}/{len(self.ren_pdfs)}")
            self._ren_log(str(self.ren_pdf_atual))
            self._ren_atualizar_botoes()
            self.txt_ren_texto.focus_set()
            self._ren_janela_atualizar_conteudo()
            if hasattr(self, "janela_renomeador") and self.janela_renomeador and self.janela_renomeador.winfo_exists() and hasattr(self, "ren_win_texto"):
                self.ren_win_texto.focus_force()

        def _ren_obter_texto_selecionado(self) -> str:
            widgets = []
            if hasattr(self, "janela_renomeador") and self.janela_renomeador and self.janela_renomeador.winfo_exists() and hasattr(self, "ren_win_texto"):
                widgets.append(self.ren_win_texto)
            widgets.append(self.txt_ren_texto)

            for widget in widgets:
                try:
                    ranges = widget.tag_ranges("sel")
                    if ranges:
                        return widget.get(ranges[0], ranges[1]).strip()
                except Exception:
                    pass
            return ""

        def _ren_obter_nome_base_para_renomear(self) -> str:
            selecionado = self._ren_obter_texto_selecionado()
            if selecionado:
                nome_limpo = limpar_texto_selecionado_nome(selecionado)
                nome_final = normalizar_nome_arquivo_manual(nome_limpo)
                if nome_final:
                    return nome_final

            if self.ren_pdf_atual:
                nome_arquivo = extrair_nome_do_arquivo_sem_cpf(self.ren_pdf_atual.name, self.ren_cpf_atual)
                nome_final = normalizar_nome_arquivo_manual(nome_arquivo)
                if nome_final:
                    return nome_final

            return ""

        def _ren_atualizar_sugestao(self, event=None):
            nome_final = self._ren_obter_nome_base_para_renomear()
            if not nome_final:
                self.lbl_ren_sugestao.configure(text="Nome sugerido: -")
                return
            pasta = self.ren_pdf_atual.parent if self.ren_pdf_atual else Path.cwd()
            nome_sugerido = reduzir_nome_para_caminho(pasta, f"{self.ren_cpf_atual}_{nome_final}.pdf")
            self.lbl_ren_sugestao.configure(text=f"Nome sugerido: {nome_sugerido}")
            self._ren_janela_atualizar_conteudo()

        def _ren_construir_novo_caminho(self, nome_base: str) -> Tuple[str, Path]:
            if not self.ren_pdf_atual:
                raise ValueError("Nenhum PDF atual carregado no renomeador.")

            pasta = self.ren_pdf_atual.parent
            novo_nome = reduzir_nome_para_caminho(pasta, f"{self.ren_cpf_atual}_{nome_base}.pdf")
            novo_path = pasta / novo_nome

            contador = 1
            while novo_path.exists() and novo_path.resolve() != self.ren_pdf_atual.resolve():
                novo_nome = reduzir_nome_para_caminho(pasta, f"{self.ren_cpf_atual}_{nome_base}_{contador}.pdf")
                novo_path = pasta / novo_nome
                contador += 1

            return novo_nome, novo_path

        def _ren_renomear_selecionado(self):
            if not self.ren_pdf_atual:
                self._ren_log("Nenhum PDF carregado no renomeador.")
                return

            nome_final = self._ren_obter_nome_base_para_renomear()
            if not nome_final:
                self._ren_log("Não foi possível identificar um nome válido para renomear.")
                return

            try:
                novo_nome, novo_path = self._ren_construir_novo_caminho(nome_final)
                antigo_path = self.ren_pdf_atual
                antigo_nome = antigo_path.name
                mover_arquivo_windows_seguro(antigo_path, novo_path)
                self.ren_pdfs[self.ren_index] = novo_path
                self.ren_pdf_atual = novo_path

                if self.origem_pdf and self.origem_pdf.is_file() and self.origem_pdf.resolve() == antigo_path.resolve():
                    self.origem_pdf = novo_path
                    self._set_entry_readonly(self.ent_origem, str(self.origem_pdf))
                    self._log(f"Origem PDF atualizada para: {self.origem_pdf}")

                self._ren_log(f"Renomeado: {antigo_nome} -> {novo_nome}")
                self.ren_index += 1
                self._ren_carregar_pdf_atual()

            except Exception as e:
                self._ren_log(f"Erro ao renomear: {e}")
                messagebox.showerror("Erro", str(e))

        def _ren_pular_pdf(self):
            if not self.ren_pdf_atual:
                self._ren_log("Nenhum PDF carregado no renomeador.")
                return
            self._ren_log(f"PDF pulado: {self.ren_pdf_atual.name}")
            self.ren_index += 1
            self._ren_carregar_pdf_atual()

        def _ren_pdf_anterior(self):
            if not self.ren_pdfs:
                self._ren_log("Nenhum PDF carregado no renomeador.")
                return
            if self.ren_index <= 0:
                self._ren_log("Você já está no primeiro PDF.")
                return
            self.ren_index -= 1
            self._ren_carregar_pdf_atual()

        def _ren_enter_renomear(self, event=None):
            try:
                widget = getattr(event, "widget", None) if event else None
                if widget is not None:
                    widget.after(1, self._ren_renomear_selecionado)
                else:
                    self.after(1, self._ren_renomear_selecionado)
            except Exception:
                self._ren_renomear_selecionado()
            return "break"

        def _ren_shift_enter_pular(self, event=None):
            try:
                widget = getattr(event, "widget", None) if event else None
                if widget is not None:
                    widget.after(1, self._ren_pular_pdf)
                else:
                    self.after(1, self._ren_pular_pdf)
            except Exception:
                self._ren_pular_pdf()
            return "break"

        def _abrir_pdf_externo(self, pdf_path: Path):
            try:
                if os.name == "nt":
                    os.startfile(str(pdf_path))
                else:
                    import subprocess
                    subprocess.Popen(["xdg-open", str(pdf_path)])
            except Exception as e:
                messagebox.showerror("Erro ao abrir PDF", str(e))

        def _abrir_janela_renomeador(self):
            if hasattr(self, "janela_renomeador") and self.janela_renomeador and self.janela_renomeador.winfo_exists():
                self.janela_renomeador.focus_force()
                self._ren_janela_atualizar_conteudo()
                return

            self.janela_renomeador = ctk.CTkToplevel(self)
            self.janela_renomeador.title("CustomerThink | PDF Renamer RH")
            self.janela_renomeador.geometry("1100x800")
            self.janela_renomeador.minsize(980, 720)
            self.janela_renomeador.grid_columnconfigure(0, weight=1)
            self.janela_renomeador.grid_rowconfigure(4, weight=1)
            self.janela_renomeador.transient(self)

            topo = ctk.CTkFrame(self.janela_renomeador)
            topo.grid(row=0, column=0, padx=14, pady=(14, 8), sticky="ew")
            topo.grid_columnconfigure(0, weight=1)

            self.ren_win_lbl_status = ctk.CTkLabel(topo, text="0/0", font=ctk.CTkFont(size=16, weight="bold"))
            self.ren_win_lbl_status.grid(row=0, column=0, padx=12, pady=(10, 4), sticky="w")

            self.ren_win_lbl_arquivo = ctk.CTkLabel(topo, text="Arquivo atual: -", justify="left", wraplength=900)
            self.ren_win_lbl_arquivo.grid(row=1, column=0, padx=12, pady=4, sticky="w")

            self.ren_win_lbl_cpf = ctk.CTkLabel(topo, text="CPF encontrado: -", justify="left")
            self.ren_win_lbl_cpf.grid(row=2, column=0, padx=12, pady=4, sticky="w")

            self.ren_win_lbl_sugestao = ctk.CTkLabel(topo, text="Nome sugerido: -", justify="left", wraplength=900)
            self.ren_win_lbl_sugestao.grid(row=3, column=0, padx=12, pady=(4, 10), sticky="w")

            btns = ctk.CTkFrame(self.janela_renomeador)
            btns.grid(row=1, column=0, padx=14, pady=8, sticky="ew")
            btns.grid_columnconfigure(5, weight=1)

            self.ren_win_btn_selecionar_pasta = ctk.CTkButton(btns, text="Selecionar Pasta de PDFs", width=190, command=self._ren_escolher_pasta_na_janela)
            self.ren_win_btn_selecionar_pasta.grid(row=0, column=0, padx=(0, 8), pady=8)

            self.ren_win_btn_abrir_pdf = ctk.CTkButton(btns, text="Abrir PDF", width=120, command=lambda: self._abrir_pdf_externo(self.ren_pdf_atual) if self.ren_pdf_atual else None)
            self.ren_win_btn_abrir_pdf.grid(row=0, column=1, padx=8, pady=8)

            self.ren_win_btn_anterior = ctk.CTkButton(btns, text="Voltar Anterior", width=130, command=self._ren_pdf_anterior)
            self.ren_win_btn_anterior.grid(row=0, column=2, padx=8, pady=8)

            self.ren_win_btn_pular = ctk.CTkButton(btns, text="Pular PDF", width=120, command=self._ren_pular_pdf)
            self.ren_win_btn_pular.grid(row=0, column=3, padx=8, pady=8)

            self.ren_win_btn_renomear = ctk.CTkButton(btns, text="Renomear e Próximo", width=170, command=self._ren_renomear_selecionado)
            self.ren_win_btn_renomear.grid(row=0, column=4, padx=8, pady=8)

            self.ren_win_btn_fechar = ctk.CTkButton(btns, text="Fechar", width=120, command=self.janela_renomeador.destroy)
            self.ren_win_btn_fechar.grid(row=0, column=5, padx=8, pady=8)

            ctk.CTkLabel(self.janela_renomeador, text="Selecione o nome no texto abaixo e pressione Enter para renomear e seguir, ou use a sugestão automática baseada no nome do arquivo.", justify="left").grid(
                row=2, column=0, padx=14, pady=(0, 6), sticky="w"
            )

            self.ren_win_texto = ctk.CTkTextbox(self.janela_renomeador)
            self.ren_win_texto.grid(row=3, column=0, padx=14, pady=6, sticky="nsew")
            self.ren_win_texto.bind("<Return>", self._ren_enter_renomear)
            self.ren_win_texto.bind("<KP_Enter>", self._ren_enter_renomear)
            self.ren_win_texto.bind("<Shift-Return>", self._ren_shift_enter_pular)
            self.ren_win_texto.bind("<Shift-KP_Enter>", self._ren_shift_enter_pular)
            self.ren_win_texto.bind("<ButtonRelease-1>", self._ren_atualizar_sugestao)
            self.ren_win_texto.bind("<KeyRelease>", self._ren_atualizar_sugestao)
            self.janela_renomeador.bind("<Return>", self._ren_enter_renomear)
            self.janela_renomeador.bind("<KP_Enter>", self._ren_enter_renomear)
            self.janela_renomeador.bind("<Shift-Return>", self._ren_shift_enter_pular)
            self.janela_renomeador.bind("<Shift-KP_Enter>", self._ren_shift_enter_pular)

            ctk.CTkLabel(self.janela_renomeador, text="Log do renomeador", font=ctk.CTkFont(size=13, weight="bold")).grid(
                row=4, column=0, padx=14, pady=(6, 4), sticky="nw"
            )
            self.ren_win_log = ctk.CTkTextbox(self.janela_renomeador, height=140)
            self.ren_win_log.grid(row=5, column=0, padx=14, pady=(0, 14), sticky="ew")

            self.after(100, self._ren_janela_atualizar_conteudo)
            self.janela_renomeador.focus_force()

        def _ren_escolher_pasta_na_janela(self):
            p = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
            if not p:
                return

            self.origem_pdf = Path(p)
            self._set_entry_readonly(self.ent_origem, str(self.origem_pdf))
            self.ren_pdfs = listar_pdfs(self.origem_pdf)
            self.ren_index = 0
            self.ren_pdf_atual = None
            self.ren_cpf_atual = ""
            self.txt_ren_texto.delete("1.0", "end")
            self._ren_limpar_log()
            self._ren_log(f"{len(self.ren_pdfs)} PDFs encontrados em: {self.origem_pdf}")
            self._log(f"Pasta selecionada para o renomeador: {self.origem_pdf}")
            self._ren_carregar_pdf_atual()
            self._ren_atualizar_botoes()

        def _ren_janela_atualizar_conteudo(self):
            if not hasattr(self, "janela_renomeador") or not self.janela_renomeador or not self.janela_renomeador.winfo_exists():
                return

            if hasattr(self, "ren_win_lbl_status"):
                self.ren_win_lbl_status.configure(text=self.lbl_ren_status.cget("text"))
                self.ren_win_lbl_arquivo.configure(text=self.lbl_ren_arquivo.cget("text"))
                self.ren_win_lbl_cpf.configure(text=self.lbl_ren_cpf.cget("text"))
                self.ren_win_lbl_sugestao.configure(text=self.lbl_ren_sugestao.cget("text"))

            if hasattr(self, "ren_win_texto"):
                origem = self.txt_ren_texto.get("1.0", "end")
                atual = self.ren_win_texto.get("1.0", "end")
                if origem != atual:
                    self.ren_win_texto.delete("1.0", "end")
                    self.ren_win_texto.insert("1.0", origem)

            if hasattr(self, "ren_win_log"):
                origem_log = self.txt_ren_log.get("1.0", "end")
                atual_log = self.ren_win_log.get("1.0", "end")
                if origem_log != atual_log:
                    self.ren_win_log.delete("1.0", "end")
                    self.ren_win_log.insert("1.0", origem_log)

            self._ren_janela_atualizar_botoes()

        def _ren_janela_atualizar_botoes(self):
            if not hasattr(self, "janela_renomeador") or not self.janela_renomeador or not self.janela_renomeador.winfo_exists():
                return
            total = len(self.ren_pdfs)
            atual_habilitado = total > 0 and self.ren_index < total
            anterior_habilitado = total > 0 and self.ren_index > 0
            self.ren_win_btn_anterior.configure(state="normal" if anterior_habilitado else "disabled")
            self.ren_win_btn_pular.configure(state="normal" if atual_habilitado else "disabled")
            self.ren_win_btn_renomear.configure(state="normal" if atual_habilitado else "disabled")
            self.ren_win_btn_abrir_pdf.configure(state="normal" if self.ren_pdf_atual else "disabled")

else:
    class App:
        def __init__(self):
            raise RuntimeError(
                "customtkinter não está instalado. Instale com: pip install customtkinter"
            )


def gerar_template_excel(destino: Path = None):
    try:
        import pandas as pd
    except ImportError:
        raise RuntimeError("Instale as dependências: pip install pandas openpyxl")

    if destino is None:
        destino = Path.cwd() / "template_base_colaboradores.xlsx"

    dados = {
        "CPF": [
            "12345678900",
            "12345678901",
            "12345678902",
            "12345678903"
        ],
        "Matrícula": [
            "0001",
            "0002",
            "0003",
            "0004"
        ],
        "Nome": [
            "Fulano de Tal Sobrenome1",
            "Fulano de Tal Sobrenome2",
            "Fulano de Tal Sobrenome3",
            "Fulano de Tal Sobrenome4"
        ],
        "Email": [
            "fulano1@email.com",
            "fulano2@email.com",
            "fulano3@email.com",
            "fulano4@email.com"
        ]
    }

    df = pd.DataFrame(dados)

    with pd.ExcelWriter(destino, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="BaseEnvio", index=False)

    return destino


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
