# bu_orcamento_manager_dynamic.py
# Orçamento Folha - Manager (Pandas + CustomTkinter) - Versão Dinâmica (contas por novas colunas)
#
# Requisitos:
#   pip install pandas openpyxl customtkinter
#
# Rodar:
#   python bu_orcamento_manager_dynamic.py
#
# O que mudou (dinâmico):
# - Qualquer nova coluna de "conta" (ex.: VCC) pode ser calculada e exportada sem alterar o código:
#   - Se a coluna tiver valor numérico -> usa o valor
#   - Se tiver flag (SIM/X/1/TRUE) -> usa premissa extra do mesmo nome (FIXO ou PERCENTUAL_SALARIO)
# - Essas colunas viram fórmulas automaticamente na tela (após Carregar Base).
# - Opcional: incluir (ou não) contas dinâmicas na base de encargos (FGTS/INSS/RAT/Terceiros).
#
# Observação:
# - Para "ler novas colunas" após editar o Excel, salve o arquivo e clique em "Carregar Base" novamente.

from __future__ import annotations

import json
import re
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Callable, Any

import pandas as pd

try:
    import customtkinter as ctk
    from tkinter import filedialog, messagebox, ttk
except Exception as e:
    raise RuntimeError("Falha ao importar customtkinter/tkinter. Instale com: pip install customtkinter") from e


TRUTHY = {"sim", "s", "x", "yes", "y", "true", "1", "ok", "ativo", "ativa"}
FALSY = {"nao", "não", "n", "no", "false", "0", "inativo", "inativa", ""}

EXTRA_TIPOS = ["FIXO", "PERCENTUAL_SALARIO"]


def _safe_str(x) -> str:
    return "" if pd.isna(x) else str(x).strip()


def normalize_key(s: str) -> str:
    """Normaliza para chave (conta/premissa) robusta."""
    s = _safe_str(s)
    s = s.replace("\n", " ").strip()
    s = re.sub(r"\s+", " ", s)
    # remove acentos simples
    s2 = s.lower()
    s2 = (
        s2.replace("ã", "a").replace("á", "a").replace("à", "a").replace("â", "a")
        .replace("é", "e").replace("ê", "e")
        .replace("í", "i")
        .replace("ó", "o").replace("ô", "o").replace("õ", "o")
        .replace("ú", "u")
        .replace("ç", "c")
    )
    return s2


def sanitize_account_name(col_name: str) -> str:
    """
    Converte um nome de coluna em "nome de conta" consistente e curto.
    Mantém legível e evita caracteres problemáticos em abas do Excel.
    """
    s = _safe_str(col_name).replace("\n", " ").strip()
    s = re.sub(r"\s+", " ", s)
    # Conta final: espaço -> _, remove caracteres estranhos
    out = re.sub(r"[^0-9A-Za-zÀ-ÿ _\-\./]", "", s).strip()
    out = out.replace(" ", "_")
    # para UI e aba Excel: limite 40 aqui; aba é 31, mas export já corta
    return out[:40] if out else "CONTA"


def parse_number_br(x) -> float:
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
        try:
            return float(x)
        except Exception:
            return 0.0
    s = str(x).strip()
    if not s:
        return 0.0
    s = re.sub(r"[R$\s]", "", s, flags=re.IGNORECASE)
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def is_truthy(v) -> bool:
    if pd.isna(v):
        return False
    if isinstance(v, (int, float)):
        return float(v) != 0.0
    s = _safe_str(v).lower()
    s = (
        s.replace("ã", "a").replace("á", "a").replace("à", "a").replace("â", "a")
        .replace("é", "e").replace("ê", "e")
        .replace("í", "i")
        .replace("ó", "o").replace("ô", "o").replace("õ", "o")
        .replace("ú", "u")
        .replace("ç", "c")
    )
    if s in TRUTHY:
        return True
    if s in FALSY:
        return False
    return bool(s)


def parse_month_cell(v) -> Optional[pd.Timestamp]:
    if pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.normalize().replace(day=1)
    if isinstance(v, datetime):
        return pd.Timestamp(v).normalize().replace(day=1)
    s = _safe_str(v)
    if not s:
        return None
    for fmt in ("%b-%y", "%m/%Y", "%Y-%m", "%m-%y", "%Y/%m", "%d/%m/%Y"):
        try:
            ts = pd.to_datetime(s, format=fmt, errors="raise")
            return ts.normalize().replace(day=1)
        except Exception:
            pass
    ts = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return None
    return ts.normalize().replace(day=1)


def month_range(start: pd.Timestamp, periods: int = 12) -> List[pd.Timestamp]:
    start = pd.Timestamp(start).normalize().replace(day=1)
    return [start + pd.DateOffset(months=i) for i in range(periods)]


def ano_mes(ts: pd.Timestamp) -> str:
    ts = pd.Timestamp(ts)
    return f"{ts.year:04d}-{ts.month:02d}"


def truncate_cell(s: Any, max_len: int = 60) -> str:
    v = _safe_str(s)
    if len(v) > max_len:
        return v[:max_len - 3] + "..."
    return v


@dataclass
class Premissas:
    fgts_rate: float = 0.08
    inss_patronal_rate: float = 0.20
    rat_rate: float = 0.02
    terceiros_rate: float = 0.058

    dsr_rate: float = 1 / 6

    vr_valor_dia: float = 0.0
    vt_valor_mes: float = 0.0
    saude_custo_mes: float = 0.0
    odonto_custo_mes: float = 0.0
    seguro_vida_custo_mes: float = 0.0
    estacionamento_custo_mes: float = 0.0
    carro_custo_mes: float = 0.0
    creche_custo_mes: float = 0.0

    previdencia_rate: float = 0.0

    prov_ferias_rate: float = 1 / 12
    prov_terco_ferias_rate: float = (1 / 12) / 3
    prov_13_rate: float = 1 / 12

    # CCT (padrão global)
    cct_default_pct: float = 0.0          # ex.: 0.03 = 3%
    cct_start_mes: str = ""               # YYYY-MM

    # NOVO: contas dinâmicas entram na base de encargos?
    dynamic_in_base_encargos: int = 0     # 0 = não, 1 = sim

    # Dinâmicas contam como variável para DSR/reflexos (férias/13)?
    dynamic_as_variavel: int = 0            # 0 = não, 1 = sim


@dataclass
class Mapping:
    col_matricula: str = "Matrícula"
    col_nome: str = "Colaborador"
    col_cdc: str = "# CDC"
    col_salario_base: str = "Salário"
    col_adicional: str = "USO RH\nAdicional Salarial"
    col_salario_orc: str = "Salário Orçamento + CCT%"
    col_promo_salario: str = "Salário Contratação ou Promoção"
    col_promo_mes: str = "A partir Mês Contratação ou Promoção"
    col_admissao: str = "Admissão"

    col_media_variavel: str = "Média"

    # Contas contábeis (opcional)
    col_conta_contabil: str = "Conta Contábil"

    col_vr: str = "VR"
    col_vt: str = "VT"
    col_saude: str = "Plano Assistência Médica"
    col_odonto: str = "Plano Assistência Odontológica"
    col_previdencia: str = "Previdência Privada"
    col_seguro: str = "Seguro Vida"
    col_estacionamento: str = "Estacionamento"
    col_carros: str = "Carros"
    col_creche: str = "Auxílio Creche"

    # CCT por colaborador (opcional)
    col_cct_flag: str = "CCT Flag"
    col_cct_pct: str = "CCT %"
    col_cct_mes: str = "CCT Mês"


FORMULAS_BASE: List[str] = [
    "SALARIO",
    "VARIAVEL_MEDIA",
    "DSR",
    "PROV_FERIAS",
    "PROV_TERCO_FERIAS",
    "PROV_13",
    "AJUSTE_13",
    "AJUSTE_FERIAS",
    "AJUSTE_TERCO_FERIAS",
    "FGTS",
    "INSS_PATRONAL",
    "RAT",
    "TERCEIROS",
    "VR",
    "VT",
    "SAUDE",
    "ODONTO",
    "PREVIDENCIA",
    "SEGURO_VIDA",
    "ESTACIONAMENTO",
    "CARROS",
    "AUX_CRECHE",
]

FORMULA_HINTS: Dict[str, str] = {
    "SALARIO": "Salário mensal (com promoção/contratação e zera antes da admissão). Inclui CCT se marcado.",
    "VARIAVEL_MEDIA": "Média variável mensal (ex.: comissão).",
    "DSR": "DSR sobre a média variável (premissa dsr_rate).",
    "PROV_FERIAS": "Provisão de férias (premissa prov_ferias_rate).",
    "PROV_TERCO_FERIAS": "Provisão 1/3 férias (premissa prov_terco_ferias_rate).",
    "PROV_13": "Provisão 13º (premissa prov_13_rate).",
    "AJUSTE_13": "Ajuste no último mês para que o total de 13º no período feche no salário do último mês.",
    "AJUSTE_FERIAS": "Ajuste no último mês para que o total de férias no período feche no salário do último mês.",
    "AJUSTE_TERCO_FERIAS": "Ajuste no último mês para que o total do 1/3 de férias no período feche em (salário do último mês)/3.",
    "FGTS": "FGTS sobre base de encargos (somatório dos itens selecionados).",
    "INSS_PATRONAL": "INSS patronal sobre base de encargos (somatório dos itens selecionados).",
    "RAT": "RAT sobre base de encargos (somatório dos itens selecionados).",
    "TERCEIROS": "Terceiros sobre base de encargos (somatório dos itens selecionados).",
    "VR": "VR mensal: valor_dia * dias_uteis (ou valor direto na coluna).",
    "VT": "VT mensal: valor padrão mensal (ou valor direto na coluna).",
    "SAUDE": "Plano de Saúde mensal (flag ou valor direto na coluna).",
    "ODONTO": "Plano Odonto mensal (flag ou valor direto na coluna).",
    "PREVIDENCIA": "Previdência: percentual do salário se flag ativo (ou valor direto).",
    "SEGURO_VIDA": "Seguro de Vida mensal (flag ou valor direto).",
    "ESTACIONAMENTO": "Estacionamento mensal (flag ou valor direto).",
    "CARROS": "Carro mensal (flag ou valor direto).",
    "AUX_CRECHE": "Auxílio creche mensal (flag ou valor direto).",
}


def read_base_from_excel(
    file_path: Path,
    sheet_name: str,
    header_row_excel: int = 1,
    max_rows: int = 2000,
) -> pd.DataFrame:
    file_path = Path(file_path)
    header_idx = max(0, int(header_row_excel) - 1)
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_idx, nrows=max_rows, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _normalize_extras(raw_extras) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    if not raw_extras:
        return out
    if isinstance(raw_extras, dict):
        for k, v in raw_extras.items():
            nome = str(k).strip()
            if not nome:
                continue
            if isinstance(v, dict):
                tipo = str(v.get("tipo", "FIXO")).strip().upper()
                if tipo not in EXTRA_TIPOS:
                    tipo = "FIXO"
                valor = float(parse_number_br(v.get("valor", 0.0)))
            else:
                tipo = "FIXO"
                valor = float(parse_number_br(v))
            out[nome] = {"tipo": tipo, "valor": valor}
    return out


def load_profile_json(path: Path) -> Tuple[Premissas, Mapping, Dict[str, Dict[str, Any]], Dict[str, str], str, int, int, List[str]]:
    raw = json.loads(Path(path).read_text(encoding="utf-8"))
    p = Premissas(**raw.get("premissas", {}))
    extras = _normalize_extras(raw.get("premissas_extras", {}))
    m = Mapping(**raw.get("mapping", {}))
    mapping_extras = raw.get("mapping_extras", {}) or {}
    mapping_extras = {str(k).strip(): str(v).strip() for k, v in mapping_extras.items() if str(k).strip()}

    per = raw.get("periodo", {})
    start_mes = per.get("start_mes", "2026-04")
    periods = int(per.get("periods", 12))
    du_default = int(per.get("du_default", 22))
    formulas_selected = raw.get("formulas_selected", [])
    formulas_selected = [str(f) for f in formulas_selected if str(f).strip()]
    return p, m, extras, mapping_extras, start_mes, periods, du_default, formulas_selected


def save_profile_json(
    path: Path,
    premissas: Premissas,
    premissas_extras: Dict[str, Dict[str, Any]],
    mapping: Mapping,
    mapping_extras: Dict[str, str],
    start_mes: str,
    periods: int,
    du_default: int,
    formulas_selected: List[str],
) -> None:
    data = {
        "premissas": asdict(premissas),
        "premissas_extras": premissas_extras or {},
        "mapping": asdict(mapping),
        "mapping_extras": mapping_extras or {},
        "periodo": {"start_mes": start_mes, "periods": periods, "du_default": du_default},
        "formulas_selected": formulas_selected or [],
    }
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def export_excel(
    fact: pd.DataFrame,
    premissas: Premissas,
    premissas_extras: Dict[str, Dict[str, Any]],
    mapping_extras: Dict[str, str],
    meses: List[pd.Timestamp],
    dias_uteis: Dict[str, int],
    start_mes: str,
    periods: int,
    du_default: int,
    formulas_selected: List[str],
    out_path: Path,
    logger: Optional[Callable[[str], None]] = None,
) -> None:
    log = logger or (lambda msg: None)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    fat = fact.copy()
    fat["ano_mes"] = fat["ano_mes"].astype(str)

    consolidado = pd.pivot_table(
        fat, index="conta", columns="ano_mes", values="valor", aggfunc="sum", fill_value=0.0
    ).reset_index()

    cal = pd.DataFrame({"mes": [pd.Timestamp(m).normalize().replace(day=1) for m in meses]})
    cal["ano_mes"] = cal["mes"].apply(ano_mes)
    cal["dias_uteis"] = cal["ano_mes"].map(lambda x: int(dias_uteis.get(x, 22)))

    prem_df = pd.DataFrame([asdict(premissas)])

    extras_rows = []
    for nome, obj in (premissas_extras or {}).items():
        n = str(nome).strip()
        if not n:
            continue
        tipo = str(obj.get("tipo", "FIXO")).strip().upper()
        if tipo not in EXTRA_TIPOS:
            tipo = "FIXO"
        valor = float(parse_number_br(obj.get("valor", 0.0)))
        extras_rows.append({"nome": n, "tipo": tipo, "valor": valor})
    extras_df = pd.DataFrame(extras_rows, columns=["nome", "tipo", "valor"])

    config_df = pd.DataFrame([{
        "start_mes": start_mes,
        "periods": int(periods),
        "du_default": int(du_default),
        "formulas_selected": ", ".join(formulas_selected or []),
    }])

    mapex_rows = [{"chave": str(k), "coluna": str(v)} for k, v in (mapping_extras or {}).items()]
    mapex_df = pd.DataFrame(mapex_rows, columns=["chave", "coluna"])

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        prem_df.to_excel(writer, sheet_name="Premissas", index=False)
        extras_df.to_excel(writer, sheet_name="Premissas_Extras", index=False)
        config_df.to_excel(writer, sheet_name="Config", index=False)
        mapex_df.to_excel(writer, sheet_name="Mapeamento_Extras", index=False)
        cal.to_excel(writer, sheet_name="Calendario", index=False)
        fat.to_excel(writer, sheet_name="FatoOrcamento", index=False)
        # Aba separada com ajustes de provisões (auditoria)
        if "conta" in fat.columns:
            aj = fat[fat["conta"].astype(str).str.startswith("AJUSTE_")].copy()
        else:
            aj = pd.DataFrame()
        aj.to_excel(writer, sheet_name="Ajustes_Provisao", index=False)
        consolidado.to_excel(writer, sheet_name="Consolidado", index=False)

        contas = sorted(fat["conta"].unique().tolist())
        for conta in contas:
            # ajustes já vão para aba própria
            if str(conta).startswith("AJUSTE_"):
                continue
            df_c = fat[fat["conta"] == conta]
            piv = pd.pivot_table(
                df_c, index="cdc", columns="ano_mes", values="valor", aggfunc="sum", fill_value=0.0
            ).reset_index()
            piv.to_excel(writer, sheet_name=f"Conta_{conta}"[:31], index=False)

    log(f"Exportado: {out_path}")


class BudgetEngine:
    def __init__(
        self,
        df_base: pd.DataFrame,
        mapping: Mapping,
        premissas: Premissas,
        premissas_extras: Optional[Dict[str, Dict[str, Any]]],
        meses: List[pd.Timestamp],
        dias_uteis_por_mes: Optional[Dict[str, int]] = None,
        dynamic_account_columns: Optional[Dict[str, str]] = None,  # conta -> coluna
        logger: Optional[Callable[[str], None]] = None,
    ):
        self.df_base = df_base.copy()
        self.mapping = mapping
        self.p = premissas
        self.p_extras = premissas_extras or {}
        self.meses = [pd.Timestamp(m).normalize().replace(day=1) for m in meses]
        self.du = dias_uteis_por_mes or {ano_mes(m): 22 for m in self.meses}
        self.dynamic_cols = dynamic_account_columns or {}
        self.log = logger or (lambda msg: None)

        # cache: premissas_extras normalizadas por chave
        self._prem_extras_norm = {normalize_key(k): v for k, v in self.p_extras.items()}

    def _clean_base(self) -> pd.DataFrame:
        df = self.df_base.copy()

        mat = df.get(self.mapping.col_matricula, pd.Series([pd.NA] * len(df)))
        mat_s = mat.astype(str).str.replace(r"\D", "", regex=True)
        df["__matricula_digits__"] = mat_s
        df = df[df["__matricula_digits__"].str.len() >= 3].copy()

        cdc = df.get(self.mapping.col_cdc, pd.Series([pd.NA] * len(df)))
        df["__cdc__"] = cdc.apply(parse_number_br).fillna(0).astype(int)

        salario_base = self._pick_salary_base(df)
        df["__salario_base__"] = salario_base.astype(float)

        media = df.get(self.mapping.col_media_variavel, pd.Series([0] * len(df)))
        df["__media_variavel__"] = media.apply(parse_number_br).astype(float)

        adm = df.get(self.mapping.col_admissao, pd.Series([pd.NaT] * len(df)))
        df["__admissao__"] = pd.to_datetime(adm, errors="coerce", dayfirst=True)

        df["__matricula__"] = df.get(self.mapping.col_matricula, "").astype(str).str.strip()
        df["__nome__"] = df.get(self.mapping.col_nome, "").astype(str).str.strip()

        cc_col = self.mapping.col_conta_contabil
        if cc_col in df.columns:
            df["__conta_contabil__"] = df[cc_col].astype(str).str.strip()
        else:
            df["__conta_contabil__"] = ""

        return df

    def _pick_salary_base(self, df: pd.DataFrame) -> pd.Series:
        col_orc = self.mapping.col_salario_orc
        col_sal = self.mapping.col_salario_base
        col_add = self.mapping.col_adicional

        orc = df[col_orc].apply(parse_number_br) if col_orc in df.columns else pd.Series([0.0] * len(df))
        sal = df[col_sal].apply(parse_number_br) if col_sal in df.columns else pd.Series([0.0] * len(df))
        add = df[col_add].apply(parse_number_br) if col_add in df.columns else pd.Series([0.0] * len(df))

        base = sal + add
        return orc.where(orc > 0, base)

    def _get_cct_settings_for_row(self, df_row: pd.Series) -> Tuple[bool, float, Optional[pd.Timestamp]]:
        m = self.mapping
        flag = df_row.get(m.col_cct_flag, pd.NA)
        aplica = is_truthy(flag)

        pct = 0.0
        if m.col_cct_pct in df_row.index:
            pct = float(parse_number_br(df_row.get(m.col_cct_pct)))
        if pct <= 0:
            pct = float(self.p.cct_default_pct)

        start = None
        if m.col_cct_mes in df_row.index:
            start = parse_month_cell(df_row.get(m.col_cct_mes))
        if start is None and _safe_str(self.p.cct_start_mes):
            try:
                start = pd.to_datetime(self.p.cct_start_mes + "-01", errors="coerce")
                if not pd.isna(start):
                    start = pd.Timestamp(start).normalize().replace(day=1)
                else:
                    start = None
            except Exception:
                start = None

        return aplica and pct > 0 and start is not None, pct, start

    def _salary_by_month(self, df: pd.DataFrame) -> pd.DataFrame:
        m = self.mapping
        promo_sal = df.get(m.col_promo_salario, pd.Series([pd.NA] * len(df))).apply(parse_number_br)
        promo_mes = df.get(m.col_promo_mes, pd.Series([pd.NA] * len(df))).apply(parse_month_cell)

        base = df["__salario_base__"].astype(float)
        adm = df["__admissao__"]

        cct_apply, cct_pct, cct_start = [], [], []
        for _, r in df.iterrows():
            a, pct, st = self._get_cct_settings_for_row(r)
            cct_apply.append(bool(a))
            cct_pct.append(float(pct))
            cct_start.append(st if st is not None else pd.NaT)

        df["__cct_apply__"] = cct_apply
        df["__cct_pct__"] = cct_pct
        df["__cct_start__"] = pd.to_datetime(cct_start, errors="coerce")

        rows = []
        for mes in self.meses:
            mes_key = ano_mes(mes)

            sal_mes = base.copy()

            has_promo = promo_mes.notna() & (promo_sal > 0)
            promo_ok = has_promo & promo_mes.apply(lambda x: (x is not None) and (mes >= x))
            sal_mes = sal_mes.where(~promo_ok, promo_sal)

            has_adm = adm.notna()
            adm_mes = pd.to_datetime(adm.dt.to_period("M").astype(str), errors="coerce")
            sal_mes = sal_mes.where(~(has_adm & (mes < adm_mes)), 0.0)

            cct_ok = df["__cct_apply__"] & (df["__cct_start__"].notna()) & (mes >= df["__cct_start__"])
            sal_mes = sal_mes.where(~cct_ok, sal_mes * (1.0 + df["__cct_pct__"]))

            tmp = pd.DataFrame({
                "mes": mes,
                "ano_mes": mes_key,
                "matricula": df["__matricula__"],
                "nome": df["__nome__"],
                "cdc": df["__cdc__"],
                "admissao": df["__admissao__"],
                "ativo": (~(has_adm & (mes < adm_mes))).astype(int),
                "salario": sal_mes.astype(float),
                "media_variavel": (df["__media_variavel__"].astype(float) * (~(has_adm & (mes < adm_mes))).astype(int)),
                "conta_contabil": df.get("__conta_contabil__", "").astype(str),
            })
            rows.append(tmp)

        return pd.concat(rows, ignore_index=True)

    def _benefit_amount(self, raw_value, default_amount: float) -> float:
        if pd.isna(raw_value):
            return 0.0
        if isinstance(raw_value, (int, float)):
            v = float(raw_value)
            if v <= 1.0:
                return default_amount if v != 0.0 else 0.0
            return v
        s = _safe_str(raw_value).lower()
        s = (
            s.replace("ã", "a").replace("á", "a").replace("à", "a").replace("â", "a")
            .replace("é", "e").replace("ê", "e")
            .replace("í", "i")
            .replace("ó", "o").replace("ô", "o").replace("õ", "o")
            .replace("ú", "u")
            .replace("ç", "c")
        )
        if s in TRUTHY:
            return default_amount
        if s in FALSY:
            return 0.0
        # tenta número em string
        num = parse_number_br(s)
        if num != 0.0:
            return num
        return default_amount

    def _premissa_default_for_column(self, col_name: str, salario_mes: float) -> float:
        """
        Se coluna é flag, tenta buscar premissa extra com mesmo nome.
        - FIXO: valor
        - PERCENTUAL_SALARIO: valor * salario_mes
        """
        k_norm = normalize_key(col_name)
        obj = self._prem_extras_norm.get(k_norm)
        if not obj:
            # tenta por conta sanitizada também
            k2 = normalize_key(sanitize_account_name(col_name))
            obj = self._prem_extras_norm.get(k2)
        if not obj:
            return 0.0

        tipo = str(obj.get("tipo", "FIXO")).strip().upper()
        valor = float(parse_number_br(obj.get("valor", 0.0)))
        if tipo == "PERCENTUAL_SALARIO":
            return float(valor) * float(salario_mes)
        return float(valor)

    def compute(self, include_formulas: Optional[List[str]] = None) -> pd.DataFrame:
        formulas = include_formulas[:] if include_formulas else []
        formulas_set = set(formulas)

        df = self._clean_base()
        self.log(f"Base carregada: {len(df)} colaboradores válidos.")

        df_em = self._salary_by_month(df)
        self.log(f"Explodiu para colaborador x mês: {len(df_em)} linhas.")

        df_flags = df.set_index("__matricula__")

        def get_flag(matricula: str, col: str):
            if col not in df_flags.columns:
                return pd.NA
            try:
                return df_flags.at[matricula, col]
            except Exception:
                return pd.NA

        fact_rows = []

        salario = df_em["salario"]
        media = df_em["media_variavel"]
        active_mask = df_em.get("ativo", pd.Series([1]*len(df_em))).astype(float)

        # --- Dinâmicas: calculo 1x e reutilizo (para export e para reflexos/DSR se habilitado) ---
        dyn_selected = [f for f in formulas if f.startswith("DYN__")]
        dyn_series = {}  # conta -> Series alinhada em df_em

        if dyn_selected:
            for f in dyn_selected:
                conta = f.replace("DYN__", "", 1)
                col = self.dynamic_cols.get(conta)
                if (not col) or (col not in df.columns):
                    continue

                vals = []
                for _, r in df_em.iterrows():
                    mat = r["matricula"]
                    raw = get_flag(mat, col)

                    # número direto
                    if isinstance(raw, (int, float)):
                        v = float(raw)
                        if v <= 1.0:
                            vals.append(self._premissa_default_for_column(col, float(r["salario"])) if v != 0.0 else 0.0)
                        else:
                            vals.append(v)
                        continue

                    num = parse_number_br(raw)
                    if num != 0.0:
                        vals.append(float(num))
                    else:
                        vals.append(self._premissa_default_for_column(col, float(r["salario"])) if is_truthy(raw) else 0.0)

                dyn_series[conta] = (pd.Series(vals).astype(float) * active_mask)

        # --- DSR (sobre variável +, opcionalmente, dinâmicas) ---
        base_variavel = (media.copy() * active_mask)
        if int(getattr(self.p, "dynamic_as_variavel", 0)) == 1 and dyn_series:
            for s in dyn_series.values():
                base_variavel = base_variavel + s

        dsr_v = base_variavel * float(self.p.dsr_rate)

        # --- Reflexos (13 / férias / 1/3) sobre remuneração: salário + variável + DSR ---
        base_reflexos = salario + base_variavel + dsr_v
        prov_f = base_reflexos * float(self.p.prov_ferias_rate)
        prov_tf = base_reflexos * float(self.p.prov_terco_ferias_rate)
        prov_13 = base_reflexos * float(self.p.prov_13_rate)

        # --- Base de encargos (somatório do que estiver selecionado) ---
        # Mantém previsível: soma apenas as contas selecionadas; dinâmicas entram só se dynamic_in_base_encargos=1.
        base_enc = pd.Series(0.0, index=df_em.index)

        if "SALARIO" in formulas_set:
            base_enc = base_enc + salario
        if "VARIAVEL_MEDIA" in formulas_set:
            # Encargos sobre a média informada (coluna 'Média'); dinâmicas são tratadas separadamente via premissa.
            base_enc = base_enc + media
        if "DSR" in formulas_set:
            base_enc = base_enc + dsr_v
        if "PROV_FERIAS" in formulas_set:
            base_enc = base_enc + prov_f
        if "PROV_TERCO_FERIAS" in formulas_set:
            base_enc = base_enc + prov_tf
        if "PROV_13" in formulas_set:
            base_enc = base_enc + prov_13


        # linhas base
        if "SALARIO" in formulas_set:
            fact_rows.append(df_em.assign(conta="SALARIO", valor=salario))
        if "VARIAVEL_MEDIA" in formulas_set:
            fact_rows.append(df_em.assign(conta="VARIAVEL_MEDIA", valor=media))
        if "DSR" in formulas_set:
            fact_rows.append(df_em.assign(conta="DSR", valor=dsr_v))
        if "PROV_FERIAS" in formulas_set:
            fact_rows.append(df_em.assign(conta="PROV_FERIAS", valor=prov_f))
        if "PROV_TERCO_FERIAS" in formulas_set:
            fact_rows.append(df_em.assign(conta="PROV_TERCO_FERIAS", valor=prov_tf))
        if "PROV_13" in formulas_set:
            fact_rows.append(df_em.assign(conta="PROV_13", valor=prov_13))

        # --- Ajustes no último mês do período (true-up): total do período fecha em salário (e 1/3) ---
        if any(k in formulas_set for k in ["AJUSTE_13", "AJUSTE_FERIAS", "AJUSTE_TERCO_FERIAS"]):
            last_ano_mes = ano_mes(self.meses[-1])
            last_mask = (df_em["ano_mes"] == last_ano_mes)

            sum13 = prov_13.groupby(df_em["matricula"]).sum()
            sumf = prov_f.groupby(df_em["matricula"]).sum()
            sumtf = prov_tf.groupby(df_em["matricula"]).sum()

            df_last = df_em.loc[last_mask, ["mes", "ano_mes", "matricula", "nome", "cdc", "conta_contabil", "salario"]].copy()
            if not df_last.empty:
                adj_13_vals = []
                adj_f_vals = []
                adj_tf_vals = []

                for _, r in df_last.iterrows():
                    mat = r["matricula"]
                    sal_last = float(r["salario"])
                    if sal_last <= 0:
                        adj_13_vals.append(0.0)
                        adj_f_vals.append(0.0)
                        adj_tf_vals.append(0.0)
                        continue
                    adj_13_vals.append(float(sal_last - float(sum13.get(mat, 0.0))))
                    adj_f_vals.append(float(sal_last - float(sumf.get(mat, 0.0))))
                    adj_tf_vals.append(float((sal_last / 3.0) - float(sumtf.get(mat, 0.0))))

                df_last = df_last.drop(columns=["salario"])

                if "AJUSTE_13" in formulas_set:
                    df_a13 = df_last.copy()
                    df_a13["valor"] = pd.Series(adj_13_vals).astype(float)
                    df_a13 = df_a13[df_a13["valor"].fillna(0.0) != 0.0]
                    if not df_a13.empty:
                        fact_rows.append(df_a13.assign(conta="AJUSTE_13", valor=df_a13["valor"]))

                if "AJUSTE_FERIAS" in formulas_set:
                    df_af = df_last.copy()
                    df_af["valor"] = pd.Series(adj_f_vals).astype(float)
                    df_af = df_af[df_af["valor"].fillna(0.0) != 0.0]
                    if not df_af.empty:
                        fact_rows.append(df_af.assign(conta="AJUSTE_FERIAS", valor=df_af["valor"]))

                if "AJUSTE_TERCO_FERIAS" in formulas_set:
                    df_atf = df_last.copy()
                    df_atf["valor"] = pd.Series(adj_tf_vals).astype(float)
                    df_atf = df_atf[df_atf["valor"].fillna(0.0) != 0.0]
                    if not df_atf.empty:
                        fact_rows.append(df_atf.assign(conta="AJUSTE_TERCO_FERIAS", valor=df_atf["valor"]))
        # benefícios conhecidos
        needs_benefits = any(k in formulas_set for k in [
            "VR", "VT", "SAUDE", "ODONTO", "PREVIDENCIA", "SEGURO_VIDA",
            "ESTACIONAMENTO", "CARROS", "AUX_CRECHE"
        ])

        if needs_benefits:
            m = self.mapping
            vr_col = m.col_vr
            vt_col = m.col_vt
            saude_col = m.col_saude
            odonto_col = m.col_odonto
            prev_col = m.col_previdencia
            seg_col = m.col_seguro
            est_col = m.col_estacionamento
            car_col = m.col_carros
            cre_col = m.col_creche

            vr_vals, vt_vals, saude_vals, odonto_vals = [], [], [], []
            previd_vals, seguro_vals, est_vals, car_vals, cre_vals = [], [], [], [], []

            for _, r in df_em.iterrows():
                mat = r["matricula"]
                am = r["ano_mes"]
                du = int(self.du.get(am, 22))

                ativo_mult = float(r.get("ativo", 1))

                vr_default = float(self.p.vr_valor_dia) * du
                vt_default = float(self.p.vt_valor_mes)
                saude_default = float(self.p.saude_custo_mes)
                odonto_default = float(self.p.odonto_custo_mes)
                seguro_default = float(self.p.seguro_vida_custo_mes)
                est_default = float(self.p.estacionamento_custo_mes)
                car_default = float(self.p.carro_custo_mes)
                cre_default = float(self.p.creche_custo_mes)

                vr_vals.append((self._benefit_amount(get_flag(mat, vr_col), vr_default) if "VR" in formulas_set else 0.0) * ativo_mult)
                vt_vals.append((self._benefit_amount(get_flag(mat, vt_col), vt_default) if "VT" in formulas_set else 0.0) * ativo_mult)
                saude_vals.append((self._benefit_amount(get_flag(mat, saude_col), saude_default) if "SAUDE" in formulas_set else 0.0) * ativo_mult)
                odonto_vals.append((self._benefit_amount(get_flag(mat, odonto_col), odonto_default) if "ODONTO" in formulas_set else 0.0) * ativo_mult)
                seguro_vals.append((self._benefit_amount(get_flag(mat, seg_col), seguro_default) if "SEGURO_VIDA" in formulas_set else 0.0) * ativo_mult)
                est_vals.append((self._benefit_amount(get_flag(mat, est_col), est_default) if "ESTACIONAMENTO" in formulas_set else 0.0) * ativo_mult)
                car_vals.append((self._benefit_amount(get_flag(mat, car_col), car_default) if "CARROS" in formulas_set else 0.0) * ativo_mult)
                cre_vals.append((self._benefit_amount(get_flag(mat, cre_col), cre_default) if "AUX_CRECHE" in formulas_set else 0.0) * ativo_mult)

                if "PREVIDENCIA" in formulas_set:
                    prev_raw = get_flag(mat, prev_col)
                    if is_truthy(prev_raw):
                        previd_vals.append(float(self.p.previdencia_rate) * float(r["salario"]))
                    else:
                        if isinstance(prev_raw, (int, float)) and float(prev_raw) > 1.0:
                            previd_vals.append(float(prev_raw))
                        else:
                            previd_vals.append(0.0)
                else:
                    previd_vals.append(0.0)

            df_ben = df_em[["mes", "ano_mes", "matricula", "nome", "cdc", "conta_contabil"]].copy()
            df_ben["VR"] = vr_vals
            df_ben["VT"] = vt_vals
            df_ben["SAUDE"] = saude_vals
            df_ben["ODONTO"] = odonto_vals
            df_ben["SEGURO_VIDA"] = seguro_vals
            df_ben["ESTACIONAMENTO"] = est_vals
            df_ben["CARROS"] = car_vals
            df_ben["AUX_CRECHE"] = cre_vals
            df_ben["PREVIDENCIA"] = previd_vals

            for conta in ["VR", "VT", "SAUDE", "ODONTO", "SEGURO_VIDA", "ESTACIONAMENTO", "CARROS", "AUX_CRECHE", "PREVIDENCIA"]:
                if conta in formulas_set:
                    fact_rows.append(
                        df_ben[["mes", "ano_mes", "matricula", "nome", "cdc", "conta_contabil"]]
                        .assign(conta=conta, valor=df_ben[conta].astype(float))
                    )

        # contas dinâmicas (colunas novas, ex.: VCC)
        dyn_selected = [f for f in formulas if f.startswith("DYN__")]
        if dyn_selected:
            for f in dyn_selected:
                conta = f.replace("DYN__", "", 1)
                col = self.dynamic_cols.get(conta)
                if not col or col not in df.columns:
                    self.log(f"Aviso: conta dinâmica '{conta}' não encontrou coluna '{col}'.")
                    continue

                vals = []
                for _, r in df_em.iterrows():
                    mat = r["matricula"]
                    raw = get_flag(mat, col)
                    if isinstance(raw, (int, float)):
                        v = float(raw)
                        if v <= 1.0:
                            # 0/1 -> trata como flag
                            vals.append(self._premissa_default_for_column(col, float(r["salario"])) if v != 0.0 else 0.0)
                        else:
                            vals.append(v)
                        continue

                    # string/flag/número em texto
                    num = parse_number_br(raw)
                    if num != 0.0:
                        vals.append(float(num))
                    else:
                        if is_truthy(raw):
                            vals.append(self._premissa_default_for_column(col, float(r["salario"])))
                        else:
                            vals.append(0.0)

                df_dyn = df_em[["mes", "ano_mes", "matricula", "nome", "cdc", "conta_contabil"]].copy()
                df_dyn["valor"] = pd.Series(vals).astype(float)
                df_dyn["valor"] = df_dyn["valor"] * df_em.get("ativo", 1).astype(float)
                df_dyn = df_dyn[df_dyn["valor"].fillna(0.0) != 0.0]
                if not df_dyn.empty:
                    fact_rows.append(df_dyn.assign(conta=conta, valor=df_dyn["valor"]))

                # opcional: se a premissa estiver marcada para entrar na base de encargos
                if int(self.p.dynamic_in_base_encargos) == 1:
                    # adiciona na base de encargos SOMENTE se conta foi selecionada
                    # (para manter comportamento previsível)
                    base_enc = base_enc + pd.Series(vals).astype(float)

        # encargos (depois de montar a base)
        if "FGTS" in formulas_set:
            fact_rows.append(df_em.assign(conta="FGTS", valor=base_enc * float(self.p.fgts_rate)))
        if "INSS_PATRONAL" in formulas_set:
            fact_rows.append(df_em.assign(conta="INSS_PATRONAL", valor=base_enc * float(self.p.inss_patronal_rate)))
        if "RAT" in formulas_set:
            fact_rows.append(df_em.assign(conta="RAT", valor=base_enc * float(self.p.rat_rate)))
        if "TERCEIROS" in formulas_set:
            fact_rows.append(df_em.assign(conta="TERCEIROS", valor=base_enc * float(self.p.terceiros_rate)))

        if not fact_rows:
            self.log("Nenhuma fórmula selecionada. Nada a calcular.")
            return pd.DataFrame(columns=["mes", "ano_mes", "matricula", "nome", "cdc", "conta", "valor"])

        fact = pd.concat(fact_rows, ignore_index=True)
        # Remove zeros (mantém AJUSTE_* mesmo se zerado, para auditoria/validação)
        v0 = fact["valor"].fillna(0.0)
        keep = (v0 != 0.0) | fact["conta"].astype(str).str.startswith("AJUSTE_")
        fact = fact[keep].copy()
        self.log(f"Fato gerado: {len(fact)} linhas (sem zeros).")
        return fact


# -------------------------
# UI
# -------------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.title("Orçamento Folha - Manager (Pandas) - Dinâmico")
        self.geometry("1250x900")
        self.minsize(1100, 740)

        self.file_path: Optional[Path] = None
        self.sheet_name: str = "Cálculo Folha"
        self.header_row_excel: int = 1

        self.df_base: Optional[pd.DataFrame] = None
        self.fact: Optional[pd.DataFrame] = None

        self.mapping = Mapping()
        self.mapping_extras: Dict[str, str] = {}

        self.premissas = Premissas()
        self.premissas_extras: Dict[str, Dict[str, Any]] = {}

        self.start_mes = "2026-04"
        self.periods = 12
        self.du_default = 22
        self.dias_uteis: Dict[str, int] = {}

        # dinâmicas
        self.dynamic_account_columns: Dict[str, str] = {}  # conta -> coluna original
        self.dynamic_formulas: List[str] = []

        # lista de fórmulas é recomputada após carregar base
        self.formulas_all: List[str] = FORMULAS_BASE[:]
        self.formulas_selected: List[str] = FORMULAS_BASE[:]

        self.extra_rows: List[Tuple[ctk.CTkEntry, ctk.CTkComboBox, ctk.CTkEntry, ctk.CTkButton]] = []
        self.mapextra_rows: List[Tuple[ctk.CTkEntry, ctk.CTkComboBox, ctk.CTkButton]] = []

        self._build_layout()

    def log(self, msg: str):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _build_layout(self):
        top = ctk.CTkFrame(self, corner_radius=0)
        top.pack(side="top", fill="x")

        self.lbl_status = ctk.CTkLabel(top, text="Arquivo: (nenhum)")
        self.lbl_status.pack(side="left", padx=12, pady=10)

        ctk.CTkButton(top, text="Abrir Excel", command=self.on_open_excel, width=120).pack(side="right", padx=10, pady=10)
        ctk.CTkButton(top, text="Salvar Perfil JSON", command=self.on_save_profile, width=150).pack(side="right", padx=10, pady=10)
        ctk.CTkButton(top, text="Carregar Perfil JSON", command=self.on_load_profile, width=170).pack(side="right", padx=10, pady=10)

        main = ctk.CTkFrame(self, corner_radius=0)
        main.pack(fill="both", expand=True)

        nav = ctk.CTkFrame(main, width=240, corner_radius=0)
        nav.pack(side="left", fill="y")

        ctk.CTkButton(nav, text="1) Arquivo", command=lambda: self.show_tab("arquivo")).pack(fill="x", padx=10, pady=(10, 6))
        ctk.CTkButton(nav, text="2) Mapeamento", command=lambda: self.show_tab("mapeamento")).pack(fill="x", padx=10, pady=6)
        ctk.CTkButton(nav, text="3) Premissas", command=lambda: self.show_tab("premissas")).pack(fill="x", padx=10, pady=6)
        ctk.CTkButton(nav, text="4) Fórmulas e Exportação", command=lambda: self.show_tab("calcular")).pack(fill="x", padx=10, pady=6)

        ctk.CTkLabel(nav, text="Log").pack(anchor="w", padx=10, pady=(16, 4))
        self.txt_log = ctk.CTkTextbox(nav, height=280, wrap="word")
        self.txt_log.pack(fill="both", expand=False, padx=10, pady=(0, 10))
        self.txt_log.configure(state="disabled")

        self.content = ctk.CTkFrame(main, corner_radius=0)
        self.content.pack(side="left", fill="both", expand=True)

        self.tabs = {
            "arquivo": self._tab_arquivo(),
            "mapeamento": self._tab_mapeamento(),
            "premissas": self._tab_premissas(),
            "calcular": self._tab_calcular(),
        }
        self.show_tab("arquivo")

    def show_tab(self, name: str):
        for frame in self.tabs.values():
            frame.pack_forget()
        self.tabs[name].pack(fill="both", expand=True)

    # -------------------------
    # Tab Arquivo
    # -------------------------
    def _tab_arquivo(self):
        frame = ctk.CTkFrame(self.content)

        ctk.CTkLabel(frame, text="Arquivo e Base", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=16, pady=(16, 10))

        line = ctk.CTkFrame(frame)
        line.pack(fill="x", padx=16)

        ctk.CTkLabel(line, text="Aba da base:").pack(side="left", padx=(0, 10), pady=10)
        self.ent_sheet = ctk.CTkEntry(line, width=240)
        self.ent_sheet.insert(0, self.sheet_name)
        self.ent_sheet.pack(side="left", pady=10)

        ctk.CTkLabel(line, text="Linha do cabeçalho:").pack(side="left", padx=(18, 10), pady=10)
        self.ent_header = ctk.CTkEntry(line, width=70)
        self.ent_header.insert(0, str(self.header_row_excel))
        self.ent_header.pack(side="left", pady=10)

        ctk.CTkButton(line, text="Carregar Base", command=self.on_load_base, width=140).pack(side="left", padx=10, pady=10)
        ctk.CTkButton(line, text="Resumo", command=self.on_show_summary, width=100).pack(side="left", padx=10, pady=10)

        wrap = ctk.CTkFrame(frame)
        wrap.pack(fill="both", expand=True, padx=16, pady=16)

        self.preview_filter = ctk.CTkEntry(wrap, placeholder_text="Filtro de colunas (ex.: CDC, Salário, Matrícula)")
        self.preview_filter.pack(fill="x", padx=10, pady=(10, 6))
        self.preview_filter.bind("<KeyRelease>", lambda e: self._refresh_preview_table())

        table_frame = ctk.CTkFrame(wrap)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.tree = ttk.Treeview(table_frame, show="headings")
        self.vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        return frame

    def _refresh_preview_table(self):
        if self.df_base is None:
            return

        df = self.df_base.copy()
        df.columns = [str(c) for c in df.columns]

        flt = self.preview_filter.get().strip().lower()
        if flt:
            cols = [c for c in df.columns if flt in c.lower()]
            if not cols:
                cols = df.columns.tolist()[:12]
        else:
            cols = df.columns.tolist()[:12]

        view = df[cols].head(40).copy()
        for c in cols:
            view[c] = view[c].map(lambda x: truncate_cell(x, 60))

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=150, stretch=True, anchor="w")

        for _, row in view.iterrows():
            self.tree.insert("", "end", values=[row[c] for c in cols])

    def on_open_excel(self):
        fp = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xlsb *.xls")],
        )
        if not fp:
            return
        self.file_path = Path(fp)
        self.lbl_status.configure(text=f"Arquivo: {self.file_path}")
        self.log(f"Arquivo selecionado: {self.file_path}")

    def on_load_base(self):
        if not self.file_path:
            messagebox.showwarning("Atenção", "Selecione o arquivo Excel primeiro.")
            return

        self.sheet_name = self.ent_sheet.get().strip() or self.sheet_name
        try:
            self.header_row_excel = int(parse_number_br(self.ent_header.get()))
            if self.header_row_excel <= 0:
                self.header_row_excel = 1
        except Exception:
            self.header_row_excel = 1
        self.ent_header.delete(0, "end")
        self.ent_header.insert(0, str(self.header_row_excel))

        try:
            df = read_base_from_excel(self.file_path, self.sheet_name, self.header_row_excel, max_rows=5000)
            self.df_base = df
            self.log(f"Base carregada da aba '{self.sheet_name}'. Linhas lidas: {len(df)}. Colunas: {len(df.columns)}.")
            self._refresh_preview_table()
            self._apply_column_filter()
            self._set_mapping_defaults_if_possible()

            # NOVO: detectar contas dinâmicas e atualizar lista de fórmulas
            self._refresh_dynamic_accounts_and_formulas()
        except Exception as e:
            self.df_base = None
            self.log(f"Erro ao carregar base: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar base:\n{e}")

    def on_show_summary(self):
        if self.df_base is None:
            messagebox.showinfo("Resumo", "Nenhuma base carregada.")
            return
        cols = self.df_base.columns.tolist()
        msg = f"Linhas: {len(self.df_base)}\nColunas: {len(cols)}\n\nPrimeiras colunas:\n" + "\n".join([str(c) for c in cols[:30]])
        messagebox.showinfo("Resumo", msg)

    # -------------------------
    # Tab Mapeamento
    # -------------------------
    def _tab_mapeamento(self):
        frame = ctk.CTkFrame(self.content)

        ctk.CTkLabel(frame, text="Mapeamento de Colunas", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=16, pady=(16, 6))
        ctk.CTkLabel(frame, text="Carregue a base antes. Use o filtro para facilitar quando tiver muitas colunas.").pack(anchor="w", padx=16, pady=(0, 10))

        bar = ctk.CTkFrame(frame)
        bar.pack(fill="x", padx=16, pady=(0, 10))

        ctk.CTkLabel(bar, text="Filtro de colunas:").pack(side="left", padx=(10, 8), pady=10)
        self.ent_col_filter = ctk.CTkEntry(bar, width=320, placeholder_text="ex.: Salário, CDC, Matrícula, CCT")
        self.ent_col_filter.pack(side="left", padx=8, pady=10)
        self.ent_col_filter.bind("<KeyRelease>", lambda e: self._apply_column_filter())
        ctk.CTkButton(bar, text="Limpar filtro", command=self._clear_column_filter, width=120).pack(side="left", padx=8, pady=10)
        ctk.CTkButton(bar, text="Recarregar colunas", command=self._apply_column_filter, width=160).pack(side="left", padx=8, pady=10)

        scroll = ctk.CTkScrollableFrame(frame, height=420)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        self.cmbs: Dict[str, ctk.CTkComboBox] = {}

        def add_row(r, label, attr_name):
            ctk.CTkLabel(scroll, text=label).grid(row=r, column=0, sticky="w", padx=10, pady=8)
            cmb = ctk.CTkComboBox(scroll, values=["(carregue a base)"], width=560)
            cmb.grid(row=r, column=1, sticky="w", padx=10, pady=8)
            self.cmbs[attr_name] = cmb

        fields = [
            ("Matrícula", "col_matricula"),
            ("Nome", "col_nome"),
            ("CDC", "col_cdc"),
            ("Salário base", "col_salario_base"),
            ("Adicional salarial", "col_adicional"),
            ("Salário orçamento (opcional)", "col_salario_orc"),
            ("Salário promoção/contratação (opcional)", "col_promo_salario"),
            ("Mês promoção/contratação (opcional)", "col_promo_mes"),
            ("Admissão (opcional)", "col_admissao"),
            ("Média variável (opcional)", "col_media_variavel"),
            ("Conta Contábil (opcional)", "col_conta_contabil"),
            ("VR (flag/valor)", "col_vr"),
            ("VT (flag/valor)", "col_vt"),
            ("Saúde (flag/valor)", "col_saude"),
            ("Odonto (flag/valor)", "col_odonto"),
            ("Previdência (flag)", "col_previdencia"),
            ("Seguro de Vida (flag/valor)", "col_seguro"),
            ("Estacionamento (flag/valor)", "col_estacionamento"),
            ("Carros (flag/valor)", "col_carros"),
            ("Auxílio Creche (flag/valor)", "col_creche"),
            ("CCT flag (SIM/NÃO)", "col_cct_flag"),
            ("CCT % (opcional)", "col_cct_pct"),
            ("CCT a partir mês (opcional)", "col_cct_mes"),
        ]
        for i, (lbl, attr) in enumerate(fields):
            add_row(i, lbl, attr)

        actions = ctk.CTkFrame(frame)
        actions.pack(fill="x", padx=16, pady=(0, 10))

        ctk.CTkButton(actions, text="Aplicar Mapeamento", command=self.on_apply_mapping, width=200).pack(side="left", padx=0, pady=10)
        self.lbl_map_status = ctk.CTkLabel(actions, text="")
        self.lbl_map_status.pack(side="left", padx=12, pady=10)

        return frame

    def _clear_column_filter(self):
        self.ent_col_filter.delete(0, "end")
        self._apply_column_filter()

    def _apply_column_filter(self):
        if self.df_base is None:
            return
        cols_all = [str(c).strip() for c in self.df_base.columns.tolist() if str(c).strip()]
        if not cols_all:
            return

        flt = self.ent_col_filter.get().strip().lower()
        cols = cols_all
        if flt:
            cols = [c for c in cols_all if flt in c.lower()] or cols_all

        for cmb in self.cmbs.values():
            cur = cmb.get()
            cmb.configure(values=cols)
            cmb.set(cur if cur in cols else cols[0])

    def _set_mapping_defaults_if_possible(self):
        if self.df_base is None:
            return
        cols = [str(c).strip() for c in self.df_base.columns.tolist()]
        if not cols:
            return

        for attr, cmb in self.cmbs.items():
            desired = getattr(self.mapping, attr, "")
            cmb.configure(values=cols)
            if desired in cols:
                cmb.set(desired)
            else:
                cmb.set(cols[0])

    def on_apply_mapping(self):
        if self.df_base is None:
            messagebox.showwarning("Atenção", "Carregue a base primeiro.")
            return

        for attr, cmb in self.cmbs.items():
            val = cmb.get().strip()
            if val and val != "(carregue a base)":
                setattr(self.mapping, attr, val)

        self.lbl_map_status.configure(text="Mapeamento aplicado.")
        self.log("Mapeamento aplicado.")

        # após mapear, re-detecta contas dinâmicas (evita pegar colunas core)
        self._refresh_dynamic_accounts_and_formulas()

    # -------------------------
    # Tab Premissas
    # -------------------------
    def _tab_premissas(self):
        frame = ctk.CTkFrame(self.content)

        ctk.CTkLabel(frame, text="Premissas", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=16, pady=(16, 10))

        scroll = ctk.CTkScrollableFrame(frame, height=780)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        wrap = ctk.CTkFrame(scroll)
        wrap.pack(fill="x", padx=10, pady=10)

        self.ent_prem: Dict[str, ctk.CTkEntry] = {}

        def add_field(row, label, key, width=170):
            ctk.CTkLabel(wrap, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=6)
            ent = ctk.CTkEntry(wrap, width=width)
            ent.grid(row=row, column=1, sticky="w", padx=10, pady=6)
            self.ent_prem[key] = ent

        add_field(0, "FGTS (ex.: 0.08)", "fgts_rate")
        add_field(1, "INSS Patronal (ex.: 0.20)", "inss_patronal_rate")
        add_field(2, "RAT (ex.: 0.02)", "rat_rate")
        add_field(3, "Terceiros (ex.: 0.058)", "terceiros_rate")
        add_field(4, "DSR sobre variável (ex.: 0.1666)", "dsr_rate")

        add_field(5, "VR valor por dia", "vr_valor_dia")
        add_field(6, "VT valor mensal", "vt_valor_mes")
        add_field(7, "Saúde custo mensal", "saude_custo_mes")
        add_field(8, "Odonto custo mensal", "odonto_custo_mes")
        add_field(9, "Seguro vida custo mensal", "seguro_vida_custo_mes")
        add_field(10, "Estacionamento custo mensal", "estacionamento_custo_mes")
        add_field(11, "Carro custo mensal", "carro_custo_mes")
        add_field(12, "Creche custo mensal", "creche_custo_mes")
        add_field(13, "Previdência (% do salário) ex.: 0.05", "previdencia_rate")

        add_field(14, "Provisão férias (padrão 1/12)", "prov_ferias_rate")
        add_field(15, "Provisão 1/3 férias (padrão 1/36)", "prov_terco_ferias_rate")
        add_field(16, "Provisão 13º (padrão 1/12)", "prov_13_rate")

        add_field(17, "CCT % padrão (ex.: 0.03)", "cct_default_pct")
        add_field(18, "CCT mês inicial (YYYY-MM)", "cct_start_mes")

        # NOVO: dinâmicas em encargos
        add_field(19, "Dinâmicas entram na base de encargos? (0/1)", "dynamic_in_base_encargos")

        period_box = ctk.CTkFrame(scroll)
        period_box.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkLabel(period_box, text="Mês inicial (YYYY-MM)").grid(row=0, column=0, sticky="w", padx=10, pady=8)
        self.ent_start = ctk.CTkEntry(period_box, width=160)
        self.ent_start.grid(row=0, column=1, sticky="w", padx=10, pady=8)

        ctk.CTkLabel(period_box, text="Qtd meses").grid(row=0, column=2, sticky="w", padx=10, pady=8)
        self.ent_periods = ctk.CTkEntry(period_box, width=90)
        self.ent_periods.grid(row=0, column=3, sticky="w", padx=10, pady=8)

        ctk.CTkLabel(period_box, text="Dias úteis padrão").grid(row=0, column=4, sticky="w", padx=10, pady=8)
        self.ent_du = ctk.CTkEntry(period_box, width=90)
        self.ent_du.grid(row=0, column=5, sticky="w", padx=10, pady=8)

        act = ctk.CTkFrame(scroll)
        act.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(act, text="Carregar valores atuais", command=self.on_load_premissas, width=200).pack(side="left", padx=0, pady=10)
        ctk.CTkButton(act, text="Aplicar premissas", command=self.on_apply_premissas, width=200).pack(side="left", padx=10, pady=10)
        self.lbl_prem_status = ctk.CTkLabel(act, text="")
        self.lbl_prem_status.pack(side="left", padx=12, pady=10)

        # extras
        extra_box = ctk.CTkFrame(scroll)
        extra_box.pack(fill="x", padx=10, pady=(10, 10))

        ctk.CTkLabel(extra_box, text="Premissas extras", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=12, pady=(12, 6))
        ctk.CTkLabel(extra_box, text="Tipo FIXO ou PERCENTUAL_SALARIO. Use isso para colunas dinâmicas (ex.: VCC) quando a coluna for flag.", wraplength=900).pack(anchor="w", padx=12, pady=(0, 10))
        ctk.CTkButton(extra_box, text="Adicionar premissa", command=self.on_add_premissa_extra, width=160).pack(anchor="w", padx=12, pady=(0, 10))

        self.extra_list = ctk.CTkFrame(extra_box)
        self.extra_list.pack(fill="x", padx=12, pady=(0, 12))
        self.extra_list.grid_columnconfigure(0, weight=2)
        self.extra_list.grid_columnconfigure(1, weight=1)
        self.extra_list.grid_columnconfigure(2, weight=1)
        self.extra_list.grid_columnconfigure(3, weight=0)

        ctk.CTkLabel(self.extra_list, text="Nome").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 6))
        ctk.CTkLabel(self.extra_list, text="Tipo").grid(row=0, column=1, sticky="w", padx=(0, 10), pady=(0, 6))
        ctk.CTkLabel(self.extra_list, text="Valor").grid(row=0, column=2, sticky="w", padx=(0, 10), pady=(0, 6))

        self.on_load_premissas()

        return frame

    def on_load_premissas(self):
        p = self.premissas
        for k, ent in self.ent_prem.items():
            v = getattr(p, k, "")
            ent.delete(0, "end")
            ent.insert(0, str(v))

        self.ent_start.delete(0, "end")
        self.ent_start.insert(0, self.start_mes)
        self.ent_periods.delete(0, "end")
        self.ent_periods.insert(0, str(self.periods))
        self.ent_du.delete(0, "end")
        self.ent_du.insert(0, str(self.du_default))

        # limpar e reconstruir extras
        for r in self.extra_rows:
            for w in r[:3]:
                try:
                    w.destroy()
                except Exception:
                    pass
            try:
                r[3].destroy()
            except Exception:
                pass
        self.extra_rows = []
        for i, (nome, obj) in enumerate((self.premissas_extras or {}).items(), start=1):
            self._add_premissa_extra_ui_row(nome, obj.get("tipo", "FIXO"), obj.get("valor", 0.0), row=i)

        self.lbl_prem_status.configure(text="Valores carregados.")

    def on_apply_premissas(self):
        p = self.premissas
        for k, ent in self.ent_prem.items():
            txt = ent.get().strip()
            if k == "cct_start_mes":
                setattr(p, k, txt)
            elif k in ("dynamic_in_base_encargos", "dynamic_as_variavel"):
                setattr(p, k, int(parse_number_br(txt)))
            else:
                setattr(p, k, float(parse_number_br(txt)))

        self.start_mes = self.ent_start.get().strip() or self.start_mes
        try:
            self.periods = int(parse_number_br(self.ent_periods.get()))
            if self.periods <= 0:
                self.periods = 12
        except Exception:
            self.periods = 12
        try:
            self.du_default = int(parse_number_br(self.ent_du.get()))
            if self.du_default <= 0:
                self.du_default = 22
        except Exception:
            self.du_default = 22

        try:
            start_ts = pd.to_datetime(self.start_mes + "-01", errors="coerce")
            if pd.isna(start_ts):
                raise ValueError("Mês inicial inválido.")
            meses = month_range(start_ts, self.periods)
            self.dias_uteis = {ano_mes(m): int(self.du_default) for m in meses}
        except Exception as e:
            self.log(f"Erro ao configurar período: {e}")
            messagebox.showerror("Erro", f"Erro no período:\n{e}")
            return

        self.premissas_extras = self._collect_premissas_extras()
        self.lbl_prem_status.configure(text="Premissas aplicadas.")
        self.log("Premissas aplicadas.")

    def _add_premissa_extra_ui_row(self, nome="", tipo="FIXO", valor=0.0, row: Optional[int] = None):
        if row is None:
            row = 1 + len(self.extra_rows)

        ent_nome = ctk.CTkEntry(self.extra_list, width=320)
        ent_nome.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=6)
        ent_nome.insert(0, str(nome))

        cmb_tipo = ctk.CTkComboBox(self.extra_list, values=EXTRA_TIPOS, width=170)
        cmb_tipo.grid(row=row, column=1, sticky="w", padx=(0, 10), pady=6)
        cmb_tipo.set(tipo if tipo in EXTRA_TIPOS else "FIXO")

        ent_val = ctk.CTkEntry(self.extra_list, width=170)
        ent_val.grid(row=row, column=2, sticky="w", padx=(0, 10), pady=6)
        ent_val.insert(0, str(valor))

        btn_rm = ctk.CTkButton(self.extra_list, text="Remover", width=110,
                               command=lambda: self._remove_premissa_extra(ent_nome, cmb_tipo, ent_val, btn_rm))
        btn_rm.grid(row=row, column=3, sticky="w", padx=(0, 0), pady=6)

        self.extra_rows.append((ent_nome, cmb_tipo, ent_val, btn_rm))

    def on_add_premissa_extra(self):
        self._add_premissa_extra_ui_row()

    def _remove_premissa_extra(self, ent_nome, cmb_tipo, ent_val, btn_rm):
        try:
            ent_nome.destroy()
            cmb_tipo.destroy()
            ent_val.destroy()
            btn_rm.destroy()
        except Exception:
            pass
        self.extra_rows = [(a, b, c, d) for (a, b, c, d) in self.extra_rows if a != ent_nome]
        self.premissas_extras = self._collect_premissas_extras()
        self.log("Premissa extra removida.")

    def _collect_premissas_extras(self) -> Dict[str, Dict[str, Any]]:
        out: Dict[str, Dict[str, Any]] = {}
        for ent_nome, cmb_tipo, ent_val, _ in self.extra_rows:
            nome = ent_nome.get().strip()
            if not nome:
                continue
            tipo = cmb_tipo.get().strip().upper()
            if tipo not in EXTRA_TIPOS:
                tipo = "FIXO"
            valor = float(parse_number_br(ent_val.get().strip()))
            out[nome] = {"tipo": tipo, "valor": valor}
        return out

    # -------------------------
    # Tab Fórmulas
    # -------------------------
    def _tab_calcular(self):
        frame = ctk.CTkFrame(self.content)

        ctk.CTkLabel(frame, text="Fórmulas e Exportação", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=16, pady=(16, 10))

        box = ctk.CTkFrame(frame)
        box.pack(fill="x", padx=16, pady=(0, 10))

        ctk.CTkLabel(box, text="Adicionar fórmula:").pack(side="left", padx=(10, 8), pady=10)
        self.cmb_formula_add = ctk.CTkComboBox(box, values=self.formulas_all, width=260)
        self.cmb_formula_add.pack(side="left", padx=8, pady=10)
        ctk.CTkButton(box, text="Adicionar", command=self.on_add_formula, width=110).pack(side="left", padx=8, pady=10)
        ctk.CTkButton(box, text="Selecionar todas", command=self.on_select_all_formulas, width=150).pack(side="left", padx=8, pady=10)
        ctk.CTkButton(box, text="Limpar seleção", command=self.on_clear_formulas, width=150).pack(side="left", padx=8, pady=10)

        mid = ctk.CTkFrame(frame)
        mid.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        left = ctk.CTkFrame(mid)
        left.pack(side="left", fill="both", expand=True, padx=(0, 10), pady=10)

        ctk.CTkLabel(left, text="Fórmulas selecionadas").pack(anchor="w", padx=10, pady=(10, 6))
        self.list_formulas = ttk.Treeview(left, show="headings", columns=("formula", "descricao"), height=14)
        self.list_formulas.heading("formula", text="Fórmula")
        self.list_formulas.heading("descricao", text="Descrição")
        self.list_formulas.column("formula", width=180, anchor="w")
        self.list_formulas.column("descricao", width=520, anchor="w")
        self.list_formulas.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        btns = ctk.CTkFrame(left)
        btns.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(btns, text="Remover selecionada", command=self.on_remove_formula, width=200).pack(side="left", padx=0, pady=10)

        right = ctk.CTkFrame(mid, width=320)
        right.pack(side="left", fill="y", padx=(0, 0), pady=10)

        ctk.CTkLabel(right, text="Ações").pack(anchor="w", padx=10, pady=(10, 6))
        ctk.CTkButton(right, text="Calcular", command=self.on_calculate, width=200).pack(anchor="w", padx=10, pady=6)
        ctk.CTkButton(right, text="Exportar Excel", command=self.on_export_excel, width=200).pack(anchor="w", padx=10, pady=6)

        self.lbl_calc = ctk.CTkLabel(right, text="")
        self.lbl_calc.pack(anchor="w", padx=10, pady=(10, 6))

        ctk.CTkLabel(right, text="Dica: após editar o Excel (novas colunas), salve e recarregue a base.").pack(anchor="w", padx=10, pady=(10, 6))

        self._refresh_selected_formulas_table()
        return frame

    def _refresh_selected_formulas_table(self):
        for item in self.list_formulas.get_children():
            self.list_formulas.delete(item)

        for f in self.formulas_selected:
            if f.startswith("DYN__"):
                conta = f.replace("DYN__", "", 1)
                col = self.dynamic_account_columns.get(conta, "")
                desc = f"Conta dinâmica: coluna '{col}'. Número usa valor; flag usa premissa extra (mesmo nome)."
            else:
                desc = FORMULA_HINTS.get(f, "")
            self.list_formulas.insert("", "end", values=(f, desc))

    def on_add_formula(self):
        f = self.cmb_formula_add.get().strip()
        if not f:
            return
        if f not in self.formulas_selected and f in self.formulas_all:
            self.formulas_selected.append(f)
            self._refresh_selected_formulas_table()

    def on_remove_formula(self):
        sel = self.list_formulas.selection()
        if not sel:
            return
        vals = self.list_formulas.item(sel[0], "values")
        if not vals:
            return
        f = vals[0]
        self.formulas_selected = [x for x in self.formulas_selected if x != f]
        self._refresh_selected_formulas_table()

    def on_select_all_formulas(self):
        self.formulas_selected = self.formulas_all[:]
        self._refresh_selected_formulas_table()

    def on_clear_formulas(self):
        self.formulas_selected = []
        self._refresh_selected_formulas_table()

    # -------------------------
    # Dinâmicas
    # -------------------------
    def _core_columns_current_mapping(self) -> set:
        """Colunas que não devem virar contas dinâmicas (mapeadas ou claramente "meta")."""
        m = self.mapping
        core = {
            m.col_matricula, m.col_nome, m.col_cdc, m.col_salario_base, m.col_adicional, m.col_salario_orc,
            m.col_promo_salario, m.col_promo_mes, m.col_admissao, m.col_media_variavel,
            m.col_vr, m.col_vt, m.col_saude, m.col_odonto, m.col_previdencia, m.col_seguro,
            m.col_estacionamento, m.col_carros, m.col_creche,
            m.col_cct_flag, m.col_cct_pct, m.col_cct_mes,
            m.col_conta_contabil
        }
        # remove vazios
        core = {c for c in core if _safe_str(c)}
        return core

    def _is_candidate_dynamic_col(self, col: str) -> bool:
        """Heurística segura para detectar coluna de custo nova."""
        if not self.df_base is not None:
            return False
        s = _safe_str(col)
        if not s:
            return False
        # ignora colunas típicas de texto/descrição
        low = s.lower()
        banned_tokens = ["descricao", "descrição", "cargo", "chefia", "local", "regime", "situacao", "situação",
                         "bate ponto", "com promo", "tipo de aumento", "filtro", "sub", "matricula", "colaborador", "admissao"]
        if any(t in low for t in banned_tokens):
            return False

        # precisa ter "cara" de custo: algum número > 0 ou flags do tipo SIM/X
        ser = self.df_base[col].head(300)
        has_num = False
        has_flag = False
        for v in ser:
            if pd.isna(v):
                continue
            if isinstance(v, (int, float)):
                if float(v) != 0.0:
                    has_num = True
                    break
            else:
                num = parse_number_br(v)
                if num != 0.0:
                    has_num = True
                    break
                if is_truthy(v):
                    has_flag = True
        return bool(has_num or has_flag)

    def _refresh_dynamic_accounts_and_formulas(self):
        if self.df_base is None:
            self.dynamic_account_columns = {}
            self.dynamic_formulas = []
            return

        core_cols = self._core_columns_current_mapping()
        dyn_cols = []

        for col in self.df_base.columns:
            c = str(col).strip()
            if not c:
                continue
            if c in core_cols:
                continue
            if self._is_candidate_dynamic_col(c):
                dyn_cols.append(c)

        # monta mapa conta -> coluna
        account_map: Dict[str, str] = {}
        used = set()
        for col in dyn_cols:
            conta = sanitize_account_name(col)
            # evita colisão
            base = conta
            i = 2
            while conta in used:
                conta = f"{base}_{i}"
                i += 1
            used.add(conta)
            account_map[conta] = col

        self.dynamic_account_columns = account_map
        self.dynamic_formulas = [f"DYN__{k}" for k in sorted(account_map.keys())]

        # lista total de fórmulas
        prev_selected = set(self.formulas_selected)
        self.formulas_all = FORMULAS_BASE[:] + self.dynamic_formulas

        # mantém seleção anterior (remove as que não existem mais)
        self.formulas_selected = [f for f in self.formulas_selected if f in self.formulas_all]
        # se estava selecionando "todas" antes, não força; só preserva o que já estava

        try:
            self.cmb_formula_add.configure(values=self.formulas_all)
            if self.formulas_all:
                self.cmb_formula_add.set(self.formulas_all[0])
        except Exception:
            pass

        self._refresh_selected_formulas_table()

        if self.dynamic_formulas:
            self.log(f"Contas dinâmicas detectadas: {len(self.dynamic_formulas)} (ex.: {self.dynamic_formulas[:5]}).")
        else:
            self.log("Nenhuma conta dinâmica detectada.")

    # -------------------------
    # Cálculo/Export
    # -------------------------
    def _ensure_ready_for_calc(self) -> bool:
        if self.df_base is None:
            messagebox.showwarning("Atenção", "Carregue a base primeiro.")
            return False
        self.on_apply_premissas()
        self.on_apply_mapping()
        return True

    def on_calculate(self):
        if not self._ensure_ready_for_calc():
            return
        if not self.formulas_selected:
            messagebox.showwarning("Atenção", "Selecione ao menos uma fórmula.")
            return
        try:
            start_ts = pd.to_datetime(self.start_mes + "-01", errors="coerce")
            meses = month_range(start_ts, self.periods)
            self.dias_uteis = {ano_mes(m): int(self.du_default) for m in meses}

            eng = BudgetEngine(
                df_base=self.df_base,
                mapping=self.mapping,
                premissas=self.premissas,
                premissas_extras=self.premissas_extras,
                meses=meses,
                dias_uteis_por_mes=self.dias_uteis,
                dynamic_account_columns=self.dynamic_account_columns,
                logger=self.log,
            )
            self.fact = eng.compute(self.formulas_selected)

            total = float(self.fact["valor"].sum()) if self.fact is not None and len(self.fact) else 0.0
            total_txt = f"{total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            self.lbl_calc.configure(text=f"Calculado. Total: {total_txt}")
            self.log("Cálculo finalizado.")
        except Exception as e:
            self.fact = None
            self.log(f"Erro no cálculo: {e}")
            messagebox.showerror("Erro", f"Erro no cálculo:\n{e}")

    def on_export_excel(self):
        if self.fact is None or self.fact.empty:
            self.on_calculate()
            if self.fact is None or self.fact.empty:
                return

        out = filedialog.asksaveasfilename(
            title="Salvar orçamento",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not out:
            return

        try:
            start_ts = pd.to_datetime(self.start_mes + "-01", errors="coerce")
            meses = month_range(start_ts, self.periods)
            self.dias_uteis = {ano_mes(m): int(self.du_default) for m in meses}

            export_excel(
                fact=self.fact,
                premissas=self.premissas,
                premissas_extras=self.premissas_extras,
                mapping_extras=self.mapping_extras,
                meses=meses,
                dias_uteis=self.dias_uteis,
                start_mes=self.start_mes,
                periods=self.periods,
                du_default=self.du_default,
                formulas_selected=self.formulas_selected,
                out_path=Path(out),
                logger=self.log,
            )
            messagebox.showinfo("Exportação", f"Arquivo exportado:\n{out}")
        except Exception as e:
            self.log(f"Erro ao exportar: {e}")
            messagebox.showerror("Erro", f"Erro ao exportar:\n{e}")

    # -------------------------
    # Perfil JSON
    # -------------------------
    def on_save_profile(self):
        out = filedialog.asksaveasfilename(
            title="Salvar perfil JSON",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
        )
        if not out:
            return
        try:
            self.on_apply_premissas()
            self.on_apply_mapping()
            save_profile_json(
                path=Path(out),
                premissas=self.premissas,
                premissas_extras=self.premissas_extras,
                mapping=self.mapping,
                mapping_extras=self.mapping_extras,
                start_mes=self.start_mes,
                periods=self.periods,
                du_default=self.du_default,
                formulas_selected=self.formulas_selected,
            )
            messagebox.showinfo("Perfil", f"Perfil salvo:\n{out}")
            self.log(f"Perfil salvo: {out}")
        except Exception as e:
            self.log(f"Erro ao salvar perfil: {e}")
            messagebox.showerror("Erro", f"Erro ao salvar perfil:\n{e}")

    def on_load_profile(self):
        fp = filedialog.askopenfilename(
            title="Carregar perfil JSON",
            filetypes=[("JSON", "*.json")],
        )
        if not fp:
            return
        try:
            p, m, extras, mapping_extras, start_mes, periods, du_default, formulas_selected = load_profile_json(Path(fp))
            self.premissas = p
            self.mapping = m
            self.premissas_extras = extras
            self.mapping_extras = mapping_extras
            self.start_mes = start_mes
            self.periods = int(periods)
            self.du_default = int(du_default)

            # carrega premissas na tela (se já criada)
            try:
                self.on_load_premissas()
            except Exception:
                pass

            # seleciona fórmulas (depois que a base carregar vai casar com dinâmicas)
            self.formulas_selected = [f for f in (formulas_selected or []) if str(f).strip()]
            self.log(f"Perfil carregado: {fp}")
            messagebox.showinfo("Perfil", f"Perfil carregado:\n{fp}")

            # tenta aplicar mapeamento na UI
            try:
                self._set_mapping_defaults_if_possible()
            except Exception:
                pass

            # se já tem base carregada, atualiza dinâmicas e normaliza seleção
            if self.df_base is not None:
                self._refresh_dynamic_accounts_and_formulas()
                # depois de recomputar formulas_all, filtra seleção
                self.formulas_selected = [f for f in self.formulas_selected if f in self.formulas_all]
                self._refresh_selected_formulas_table()
        except Exception as e:
            self.log(f"Erro ao carregar perfil: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar perfil:\n{e}")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
