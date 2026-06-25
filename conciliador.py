import re
import unicodedata
import pandas as pd
import numpy as np
from datetime import datetime
from difflib import SequenceMatcher
from itertools import combinations


# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DE CAMPOS
# ══════════════════════════════════════════════════════════════════════════════

FIELD_HINTS = {
    "data":      ["data", "dtmovimento", "datamov", "dtlancamento",
                  "dtpagamento", "datapagamento", "competencia", "emissao", "baixa"],
    "valor":     ["valor", "vl", "value", "montante", "amount", "total",
                  "price", "preco", "pgto", "pago"],
    "debito":    ["debito", "deb", "saida", "debit", "db"],
    "credito":   ["credito", "cred", "entrada", "credit", "cr"],
    "descricao": ["descricao", "descrição", "hist", "memo", "narr", "obs", "observ",
                  "lancamento", "lançamento", "fornecedor", "cliente",
                  "beneficiario", "name", "nome"],
    "documento": ["documento", "doc", "nf", "nota", "numero", "número", "ref",
                  "referencia", "cheque", "boleto", "nsu", "codigo", "dcto"],
    "tipo":      ["tipo", "natureza", "d/c", "debcred", "modalidade", "operacao"],
}

FORMATOS_DATA = [
    "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%m/%d/%Y",
    "%d-%m-%Y", "%d-%m-%y", "%Y/%m/%d",
    "%d.%m.%Y", "%d.%m.%y", "%Y%m%d",
]


# ══════════════════════════════════════════════════════════════════════════════
# PARSERS DE DATA E VALOR
# ══════════════════════════════════════════════════════════════════════════════

def _parse_data(valor):
    if valor is None:
        return pd.NaT
    if isinstance(valor, float) and np.isnan(valor):
        return pd.NaT
    if isinstance(valor, pd.Timestamp):
        return valor.normalize()
    if isinstance(valor, datetime):
        return pd.Timestamp(valor).normalize()
    s = str(valor).strip().split(" ")[0].split("T")[0]
    if not s or s.lower() in ("nat", "nan", "none", ""):
        return pd.NaT
    for fmt in FORMATOS_DATA:
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except ValueError:
            continue
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def _parse_valor(valor) -> float:
    if valor is None:
        return np.nan
    if isinstance(valor, (int, float)):
        return float(valor) if not (isinstance(valor, float) and np.isnan(valor)) else np.nan
    s = str(valor).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return np.nan
    negativo = s.startswith("(") and s.endswith(")")
    if negativo:
        s = s[1:-1]
    s = re.sub(r"[R$€£\s\"']", "", s)
    sinal = -1 if (s.startswith("-") or negativo) else 1
    s = s.lstrip("-+")
    if not s:
        return np.nan
    np_ = s.count(".")
    nv  = s.count(",")
    if np_ == 0 and nv == 0:
        try:
            return sinal * float(s)
        except ValueError:
            return np.nan
    if np_ == 0 and nv == 1:
        partes = s.split(",")
        if len(partes[1]) == 3 and len(partes[0]) <= 3 and int(partes[1]) >= 100:
            return sinal * float(s.replace(",", ""))
        return sinal * float(s.replace(",", "."))
    if np_ == 1 and nv == 0:
        return sinal * float(s)
    if np_ > 1 and nv == 0:
        return sinal * float(s.replace(".", ""))
    if nv > 1 and np_ == 0:
        return sinal * float(s.replace(",", ""))
    if nv >= 1 and np_ >= 1:
        if s.rindex(".") < s.index(","):
            return sinal * float(s.replace(".", "").replace(",", "."))
        else:
            return sinal * float(s.replace(",", ""))
    try:
        return sinal * float(s.replace(",", "."))
    except ValueError:
        return np.nan


# ══════════════════════════════════════════════════════════════════════════════
# NORMALIZAÇÃO DE TEXTO
# ══════════════════════════════════════════════════════════════════════════════

def _normalizar_texto(texto) -> str:
    if not texto or str(texto).lower() in ("nan", "none", ""):
        return ""
    s = str(texto)
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.lower()
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _similaridade(a, b) -> float:
    na, nb = _normalizar_texto(a), _normalizar_texto(b)
    if not na or not nb:
        return 0.0
    return SequenceMatcher(None, na, nb).ratio()


# ══════════════════════════════════════════════════════════════════════════════
# DETECÇÃO DE COLUNAS E CARREGAMENTO
# ══════════════════════════════════════════════════════════════════════════════

def detect_columns(df: pd.DataFrame) -> dict:
    """Detecta automaticamente as colunas. Retorna mapping com chaves:
    data, valor, debito, credito, descricao, documento, tipo"""
    cols_lower = {c: c.lower().replace(" ", "").replace("_", "") for c in df.columns}
    mapping = {f: None for f in FIELD_HINTS}
    for field, hints in FIELD_HINTS.items():
        for col, col_clean in cols_lower.items():
            if any(h in col_clean for h in hints):
                mapping[field] = col
                break
    if not mapping["data"]:
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                mapping["data"] = col
                break
    if not mapping["valor"] and not mapping["debito"] and not mapping["credito"]:
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                mapping["valor"] = col
                break
    return mapping


def load_sheet(raw: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """Carrega planilha já lida (raw DataFrame) aplicando o mapping de colunas.
    Suporta coluna única de valor OU par débito/crédito separados."""
    df = pd.DataFrame()

    # Data
    col_data = mapping.get("data")
    df["data"] = raw[col_data].apply(_parse_data) if (col_data and col_data in raw.columns) else pd.NaT

    # Valor: prioriza coluna única; se não houver, combina débito e crédito
    col_valor   = mapping.get("valor")
    col_debito  = mapping.get("debito")
    col_credito = mapping.get("credito")

    if col_valor and col_valor in raw.columns:
        serie = raw[col_valor].apply(_parse_valor)
        # Detecta padrão Itaú: coluna única onde débito vem negativo e crédito positivo
        validos = serie.dropna()
        pct_neg = (validos < 0).sum() / len(validos) if len(validos) > 0 else 0
        if pct_neg >= 0.05:
            # Coluna mista D/C: separa e usa valor absoluto para conciliação
            df["valor_credito"] = serie.where(serie > 0, np.nan)
            df["valor_debito"]  = serie.where(serie < 0, np.nan).abs()
            df["valor"]         = serie.abs()
        else:
            df["valor"]         = serie
            df["valor_credito"] = np.nan
            df["valor_debito"]  = np.nan
    elif col_debito or col_credito:
        serie_deb  = raw[col_debito].apply(_parse_valor)  if (col_debito  and col_debito  in raw.columns) else pd.Series(np.nan, index=raw.index)
        serie_cred = raw[col_credito].apply(_parse_valor) if (col_credito and col_credito in raw.columns) else pd.Series(np.nan, index=raw.index)
        serie_deb  = serie_deb.abs()
        serie_cred = serie_cred.abs()
        df["valor_debito"]  = serie_deb
        df["valor_credito"] = serie_cred
        df["valor"] = serie_cred.fillna(serie_deb)
        df["valor"] = df["valor"].replace(0, np.nan)
    else:
        df["valor"]         = np.nan
        df["valor_credito"] = np.nan
        df["valor_debito"]  = np.nan

    # Demais campos
    for field in ("descricao", "documento", "tipo"):
        col = mapping.get(field)
        df[field] = raw[col] if (col and col in raw.columns) else np.nan

    # Detecta tipo D/C
    if "valor_debito" in df.columns and "valor_credito" in df.columns:
        def _tipo_dc(row):
            tem_cred = pd.notna(row.get("valor_credito")) and row.get("valor_credito", 0) > 0
            tem_deb  = pd.notna(row.get("valor_debito"))  and row.get("valor_debito", 0)  > 0
            if tem_cred and not tem_deb: return "C"
            if tem_deb  and not tem_cred: return "D"
            if tem_cred and tem_deb: return "C/D"
            return ""
        df["tipo_dc"] = df.apply(_tipo_dc, axis=1)
    else:
        df["tipo_dc"] = ""

    df["_idx"] = range(len(df))

    # ── Remove linhas completamente vazias (sem valor E sem data) ──
    tem_valor = df["valor"].notna() & (df["valor"] != 0)
    tem_data  = df["data"].notna()
    df = df[tem_valor | tem_data].copy()
    df = df.reset_index(drop=True)
    df["_idx"] = range(len(df))

    return df


# ══════════════════════════════════════════════════════════════════════════════
# SCORE DE CONFIABILIDADE
# ══════════════════════════════════════════════════════════════════════════════

def _calcular_score(b, e, tol_valor, tol_dias, usar_doc, usar_desc, min_sim) -> dict:
    resultado = {
        "score": 0,
        "ok_valor": False, "ok_data": False, "ok_doc": False, "ok_desc": False,
        "dif_valor": None, "dif_dias": None, "sim_desc": 0.0,
        "motivos": [], "dc_incompativel": False,
    }

    # ── Verifica compatibilidade D/C ──
    # Se ambos os lados têm D/C identificado e são opostos (um D, outro C),
    # não é a mesma natureza de lançamento — desqualifica o match.
    dc_b = (b.get("tipo_dc") or "").strip()
    dc_e = (e.get("tipo_dc") or "").strip()
    if dc_b in ("D", "C") and dc_e in ("D", "C") and dc_b != dc_e:
        resultado["dc_incompativel"] = True
        resultado["motivos"].append(f"⚠ D/C incompatível (Banco={dc_b} / ERP={dc_e})")
        return resultado  # score fica 0 — não concilia

    if pd.notna(b.get("valor")) and pd.notna(e.get("valor")):
        dif = abs(float(b["valor"]) - float(e["valor"]))
        resultado["dif_valor"] = round(dif, 2)
        if dif <= 0.001:
            resultado["score"] += 40
            resultado["ok_valor"] = True
            resultado["motivos"].append("Valor exato")
        elif dif <= tol_valor:
            resultado["score"] += 25
            resultado["ok_valor"] = True
            resultado["motivos"].append(f"Valor (dif R$ {dif:.2f})")

    if pd.notna(b.get("data")) and pd.notna(e.get("data")):
        try:
            dif_d = abs((pd.Timestamp(b["data"]) - pd.Timestamp(e["data"])).days)
            resultado["dif_dias"] = dif_d
            if dif_d == 0:
                resultado["score"] += 30
                resultado["ok_data"] = True
                resultado["motivos"].append("Data exata")
            elif dif_d <= tol_dias:
                resultado["score"] += 20
                resultado["ok_data"] = True
                resultado["motivos"].append(f"Data ({dif_d}d diferenca)")
        except Exception:
            pass

    if usar_doc:
        bd = str(b.get("documento", "")).strip()
        ed = str(e.get("documento", "")).strip()
        if bd and ed and bd not in ("nan", "None", "") and ed not in ("nan", "None", "") and bd == ed:
            resultado["score"] += 20
            resultado["ok_doc"] = True
            resultado["motivos"].append("Documento igual")

    if usar_desc:
        sim = _similaridade(b.get("descricao"), e.get("descricao"))
        resultado["sim_desc"] = round(sim * 100, 1)
        if sim >= min_sim:
            resultado["score"] += 10
            resultado["ok_desc"] = True
            resultado["motivos"].append(f"Descricao {resultado['sim_desc']}%")

    return resultado


def _nivel_confianca(score: int):
    if score >= 90:
        return "Alta",       "ok"
    if score >= 70:
        return "Média",      "warn"
    if score >= 50:
        return "Baixa",      "info"
    return "Muito Baixa",    "err"


def _tipo_match(sc: dict) -> str:
    if sc["ok_valor"] and sc["ok_data"] and sc["ok_doc"]:
        return "Match Exato"
    if sc["ok_valor"] and sc["ok_data"]:
        return "Match por Similaridade"
    if sc["ok_valor"]:
        return "Match por Tolerancia"
    return "Correspondencia Parcial"


# ══════════════════════════════════════════════════════════════════════════════
# DETECÇÃO DE DUPLICIDADES
# ══════════════════════════════════════════════════════════════════════════════

def detectar_duplicidades(df: pd.DataFrame, origem: str) -> pd.DataFrame:
    df = df.copy()
    df["_dup_key"] = df["valor"].astype(str) + "_" + df["data"].astype(str)
    contagem = df["_dup_key"].value_counts()
    duplicados = contagem[contagem > 1].index
    df_dup = df[df["_dup_key"].isin(duplicados)].copy()
    df_dup["origem"] = origem
    df_dup = df_dup.drop(columns=["_dup_key", "_idx"], errors="ignore")
    return df_dup


# ══════════════════════════════════════════════════════════════════════════════
# AGRUPAMENTO E DESMEMBRAMENTO
# ══════════════════════════════════════════════════════════════════════════════

def _buscar_combinacoes(valor_alvo: float, df_pool: pd.DataFrame,
                        tol: float = 0.01, max_itens: int = 5) -> list:
    candidatos = []
    pool = df_pool[pd.notna(df_pool["valor"])].copy()
    pool = pool[pool["valor"] <= valor_alvo + tol]
    if len(pool) == 0 or len(pool) > 50:
        return candidatos
    indices = list(pool.index)
    for n in range(2, min(max_itens + 1, len(indices) + 1)):
        for combo in combinations(indices, n):
            soma = pool.loc[list(combo), "valor"].sum()
            if abs(soma - valor_alvo) <= tol:
                candidatos.append(pool.loc[list(combo)])
                if len(candidatos) >= 3:
                    return candidatos
    return candidatos


def conciliar_agrupamentos(df_nao_conc_banco, df_nao_conc_erp,
                            tol_valor=0.01, tol_dias=5) -> list:
    resultados = []
    usado_erp   = set()
    usado_banco = set()

    for bi, b in df_nao_conc_banco.iterrows():
        if bi in usado_banco or pd.isna(b.get("valor")):
            continue
        pool_erp = df_nao_conc_erp[~df_nao_conc_erp.index.isin(usado_erp)]
        if pd.notna(b.get("data")):
            pool_erp = pool_erp[
                pool_erp["data"].apply(
                    lambda d: abs((pd.Timestamp(d) - pd.Timestamp(b["data"])).days) <= tol_dias
                    if pd.notna(d) else True
                )
            ]
        combos = _buscar_combinacoes(float(b["valor"]), pool_erp, tol_valor)
        for combo in combos:
            resultados.append({
                "tipo":        "Agrupamento",
                "descricao":   "1 lancamento bancario = soma de multiplos no ERP",
                "banco_idx":   [bi],
                "erp_idx":     list(combo.index),
                "banco_valor": b["valor"],
                "erp_valores": list(combo["valor"]),
                "soma_erp":    combo["valor"].sum(),
                "banco_data":  b.get("data"),
                "banco_desc":  b.get("descricao"),
                "score":       85,
                "confianca":   "Alta",
            })
            usado_banco.add(bi)
            usado_erp.update(combo.index)
            break

    for ei, e in df_nao_conc_erp.iterrows():
        if ei in usado_erp or pd.isna(e.get("valor")):
            continue
        pool_banco = df_nao_conc_banco[~df_nao_conc_banco.index.isin(usado_banco)]
        if pd.notna(e.get("data")):
            pool_banco = pool_banco[
                pool_banco["data"].apply(
                    lambda d: abs((pd.Timestamp(d) - pd.Timestamp(e["data"])).days) <= tol_dias
                    if pd.notna(d) else True
                )
            ]
        combos = _buscar_combinacoes(float(e["valor"]), pool_banco, tol_valor)
        for combo in combos:
            resultados.append({
                "tipo":          "Desmembramento",
                "descricao":     "1 lancamento ERP liquidado em multiplos no banco",
                "banco_idx":     list(combo.index),
                "erp_idx":       [ei],
                "banco_valores": list(combo["valor"]),
                "soma_banco":    combo["valor"].sum(),
                "erp_valor":     e["valor"],
                "erp_data":      e.get("data"),
                "erp_desc":      e.get("descricao"),
                "score":         80,
                "confianca":     "Alta",
            })
            usado_erp.add(ei)
            usado_banco.update(combo.index)
            break

    return resultados


# ══════════════════════════════════════════════════════════════════════════════
# ENGINE PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

def conciliar(
    df_banco,
    df_erp,
    tol_valor=0.01,
    tol_dias=2,
    usar_documento=True,
    usar_descricao=False,
    min_similaridade=0.70,
) -> dict:
    banco = df_banco.copy().reset_index(drop=True)
    erp   = df_erp.copy().reset_index(drop=True)
    usado_erp = set()
    rows = []

    for bi, b in banco.iterrows():
        melhor_sc    = None
        melhor_score = 0
        melhor_ei    = None
        melhor_e     = None

        for ei, e in erp.iterrows():
            if ei in usado_erp:
                continue
            sc = _calcular_score(b, e, tol_valor, tol_dias,
                                 usar_documento, usar_descricao, min_similaridade)
            if sc["score"] > melhor_score and sc["score"] >= 40:
                melhor_score = sc["score"]
                melhor_sc    = sc
                melhor_ei    = ei
                melhor_e     = e

        if melhor_e is not None:
            usado_erp.add(melhor_ei)
            sc       = melhor_sc
            nivel, _ = _nivel_confianca(sc["score"])
            tipo_m   = _tipo_match(sc)
            status   = "Conciliado" if sc["score"] >= 70 else "Divergente"

            motivo_div = []
            if sc["dif_valor"] is not None and sc["dif_valor"] > 0.001:
                motivo_div.append(f"Valor dif R$ {sc['dif_valor']:.2f}")
            if sc["dif_dias"] is not None and sc["dif_dias"] > 0:
                motivo_div.append(f"Data {sc['dif_dias']}d")
            if usar_descricao and not sc["ok_desc"]:
                motivo_div.append("Descricao divergente")

            rows.append({
                "banco_data":      b.get("data"),
                "banco_valor":     b.get("valor"),
                "banco_debito":    b.get("valor_debito"),
                "banco_credito":   b.get("valor_credito"),
                "banco_descricao": b.get("descricao"),
                "banco_documento": b.get("documento"),
                "erp_data":        melhor_e.get("data"),
                "erp_valor":       melhor_e.get("valor"),
                "erp_debito":      melhor_e.get("valor_debito"),
                "erp_credito":     melhor_e.get("valor_credito"),
                "erp_descricao":   melhor_e.get("descricao"),
                "erp_documento":   melhor_e.get("documento"),
                "status":          status,
                "tipo_match":      tipo_m,
                "score":           sc["score"],
                "confianca":       nivel,
                "motivos":         " | ".join(sc["motivos"]),
                "divergencias":    " | ".join(motivo_div) if motivo_div else "—",
                "diferenca_valor": sc["dif_valor"],
                "diferenca_dias":  sc["dif_dias"],
                "sim_descricao":   sc["sim_desc"],
                "_banco_idx":      bi,
                "_erp_idx":        melhor_ei,
            })
        else:
            rows.append({
                "banco_data":      b.get("data"),
                "banco_valor":     b.get("valor"),
                "banco_debito":    b.get("valor_debito"),
                "banco_credito":   b.get("valor_credito"),
                "banco_descricao": b.get("descricao"),
                "banco_documento": b.get("documento"),
                "erp_data":        None, "erp_valor":     None,
                "erp_debito":      None, "erp_credito":   None,
                "erp_descricao":   None, "erp_documento": None,
                "status":          "Nao Conciliado",
                "tipo_match":      "—",
                "score":           0,   "confianca": "—",
                "motivos":         "—", "divergencias": "Sem par no ERP",
                "diferenca_valor": None, "diferenca_dias": None,
                "sim_descricao":   0.0,
                "_banco_idx":      bi,  "_erp_idx": None,
            })

    for ei, e in erp.iterrows():
        if ei not in usado_erp:
            rows.append({
                "banco_data":      None, "banco_valor":     None,
                "banco_debito":    None, "banco_credito":   None,
                "banco_descricao": None, "banco_documento": None,
                "erp_data":        e.get("data"),
                "erp_valor":       e.get("valor"),
                "erp_debito":      e.get("valor_debito"),
                "erp_credito":     e.get("valor_credito"),
                "erp_descricao":   e.get("descricao"),
                "erp_documento":   e.get("documento"),
                "status":          "So no ERP",
                "tipo_match":      "—",
                "score":           0,   "confianca": "—",
                "motivos":         "—", "divergencias": "Sem par no banco",
                "diferenca_valor": None, "diferenca_dias": None,
                "sim_descricao":   0.0,
                "_banco_idx":      None, "_erp_idx": ei,
            })

    df_resultado = pd.DataFrame(rows)

    nc_banco_idx = df_resultado[df_resultado["status"] == "Nao Conciliado"]["_banco_idx"].dropna().astype(int).tolist()
    nc_erp_idx   = df_resultado[df_resultado["status"] == "So no ERP"]["_erp_idx"].dropna().astype(int).tolist()
    df_nc_banco  = banco.loc[banco.index.isin(nc_banco_idx)].copy()
    df_nc_erp    = erp.loc[erp.index.isin(nc_erp_idx)].copy()

    agrupamentos = []
    if len(df_nc_banco) > 0 and len(df_nc_erp) > 0:
        agrupamentos = conciliar_agrupamentos(df_nc_banco, df_nc_erp, tol_valor)

    sugestoes = []
    for bi, b in banco.loc[banco.index.isin(nc_banco_idx)].iterrows():
        for ei, e in erp.iterrows():
            if ei in usado_erp:
                continue
            sc = _calcular_score(b, e, tol_valor * 5, tol_dias * 3, usar_documento, True, 0.4)
            if 20 <= sc["score"] < 40:
                sugestoes.append({
                    "banco_data":      b.get("data"),
                    "banco_valor":     b.get("valor"),
                    "banco_descricao": b.get("descricao"),
                    "erp_data":        e.get("data"),
                    "erp_valor":       e.get("valor"),
                    "erp_descricao":   e.get("descricao"),
                    "score":           sc["score"],
                    "motivos":         " | ".join(sc["motivos"]) or "Similaridade parcial",
                })

    dup_banco = detectar_duplicidades(banco, "Banco")
    dup_erp   = detectar_duplicidades(erp,   "ERP")

    return {
        "resultado":          df_resultado,
        "agrupamentos":       agrupamentos,
        "sugestoes":          sugestoes[:20],
        "duplicidades_banco": dup_banco,
        "duplicidades_erp":   dup_erp,
        "resumo":             _calcular_resumo(df_resultado, agrupamentos, dup_banco, dup_erp),
    }


# ══════════════════════════════════════════════════════════════════════════════
# RESUMO EXECUTIVO
# ══════════════════════════════════════════════════════════════════════════════

def _calcular_resumo(df, agrupamentos, dup_banco, dup_erp) -> dict:
    total_banco  = len(df[df["banco_valor"].notna()])
    total_erp    = len(df[df["erp_valor"].notna()])
    conciliados  = len(df[df["status"] == "Conciliado"])
    divergentes  = len(df[df["status"] == "Divergente"])
    nao_conc     = len(df[df["status"] == "Nao Conciliado"])
    so_erp       = len(df[df["status"] == "So no ERP"])
    n_agrup      = len([a for a in agrupamentos if a["tipo"] == "Agrupamento"])
    n_desmemb    = len([a for a in agrupamentos if a["tipo"] == "Desmembramento"])
    taxa         = round(conciliados / total_banco * 100, 1) if total_banco else 0.0
    conc_df      = df[df["status"] == "Conciliado"]
    score_medio  = round(conc_df["score"].mean(), 1) if len(conc_df) > 0 else 0.0

    return {
        "total_banco":          total_banco,
        "total_erp":            total_erp,
        "conciliados":          conciliados,
        "divergentes":          divergentes,
        "nao_conciliados":      nao_conc,
        "so_erp":               so_erp,
        "agrupamentos":         n_agrup,
        "desmembramentos":      n_desmemb,
        "taxa_pct":             taxa,
        "score_medio":          score_medio,
        "valor_total_banco":    df["banco_valor"].sum(),
        "valor_conciliado":     df[df["status"] == "Conciliado"]["banco_valor"].sum(),
        "valor_divergente":     df[df["status"] == "Divergente"]["banco_valor"].sum(),
        "valor_nao_conciliado": df[df["status"] == "Nao Conciliado"]["banco_valor"].sum(),
        "valor_so_erp":         df[df["status"] == "So no ERP"]["erp_valor"].sum(),
        "duplicidades_banco":   len(dup_banco),
        "duplicidades_erp":     len(dup_erp),
    }


def resumo(resultado):
    if isinstance(resultado, dict):
        return resultado["resumo"]
    return _calcular_resumo(resultado, [], pd.DataFrame(), pd.DataFrame())
