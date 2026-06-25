import streamlit as st
import pandas as pd
import numpy as np

from conciliador import detect_columns, load_sheet, conciliar, resumo
from exportar import exportar_excel, exportar_csv_erp
from exportar_pdf import exportar_pdf

st.set_page_config(
    page_title="Conciliador Bancário",
    page_icon="⇄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Segoe UI', Arial, sans-serif; }
.stApp { background: #f8f9fb; }
.concil-header {
    background: linear-gradient(135deg, #085041 0%, #0F6E56 100%);
    color: white; padding: 1.5rem 2rem; border-radius: 12px; margin-bottom: 1.5rem;
    display: flex; align-items: center; gap: 1rem;
}
.concil-header h1 { margin: 0; font-size: 1.6rem; font-weight: 600; }
.concil-header p  { margin: 0; opacity: 0.85; font-size: 0.9rem; }
.steps-bar {
    display: flex; background: white; border-radius: 10px;
    overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 1.5rem;
}
.step-item { flex:1; padding:.75rem 1rem; text-align:center; font-size:.82rem; color:#888; border-right:1px solid #eee; }
.step-item:last-child { border-right:none; }
.step-item.active { background:#E1F5EE; color:#085041; font-weight:600; }
.step-item.done { color:#0F6E56; }
.step-num { font-size:.7rem; display:block; color:#bbb; margin-bottom:2px; }
.step-item.active .step-num, .step-item.done .step-num { color:#1D9E75; }
.metric-card { background:white; border-radius:10px; padding:1.2rem 1.5rem; text-align:center; box-shadow:0 1px 4px rgba(0,0,0,.06); }
.metric-val { font-size:2rem; font-weight:700; line-height:1; }
.metric-lbl { font-size:.78rem; color:#888; margin-top:4px; }
.metric-ok .metric-val { color:#0F6E56; }
.metric-warn .metric-val { color:#BA7517; }
.metric-err .metric-val { color:#A32D2D; }
.metric-info .metric-val { color:#185FA5; }
.badge { display:inline-block; padding:2px 10px; border-radius:20px; font-size:.78rem; font-weight:500; }
.badge-ok   { background:#E1F5EE; color:#085041; }
.badge-warn { background:#FAEEDA; color:#633806; }
.badge-err  { background:#FCEBEB; color:#791F1F; }
.badge-info { background:#E6F1FB; color:#0C447C; }
.result-table { width:100%; border-collapse:collapse; font-size:.85rem; }
.result-table th { background:#085041; color:white; padding:8px 10px; text-align:left; font-weight:500; position:sticky; top:0; }
.result-table td { padding:7px 10px; border-bottom:1px solid #f0f0f0; white-space:nowrap; max-width:180px; overflow:hidden; text-overflow:ellipsis; }
.result-table tr:hover td { background:#f5fbf8; }
.section-title { font-size:.9rem; font-weight:600; color:#333; margin-bottom:.5rem; display:flex; align-items:center; gap:6px; }
.progress-outer { background:#e8e8e8; border-radius:6px; height:10px; margin:6px 0 12px; overflow:hidden; }
.progress-inner { height:100%; background:linear-gradient(90deg,#0F6E56,#1D9E75); border-radius:6px; transition:width .6s ease; }
div[data-testid="stButton"] > button { border-radius:8px; font-weight:500; border:none; padding:.5rem 1.25rem; }
.info-banner { background:#E6F1FB; border-left:4px solid #378ADD; border-radius:0 8px 8px 0; padding:.75rem 1rem; font-size:.85rem; color:#185FA5; margin-bottom:1rem; }
</style>
""", unsafe_allow_html=True)

# ── Sessão ────────────────────────────────────────────────────────────────────
def _init():
    for k, v in {"step":0,"raw_banco":None,"raw_erp":None,"map_banco":{},
                 "map_erp":{},"df_result":None,"resumo":None,
                 "resultado_completo":None}.items():
        if k not in st.session_state:
            st.session_state[k] = v
_init()
S = st.session_state

STEPS = ["Importar","Mapear colunas","Critérios","Resultado","Exportar"]

def render_steps():
    html = '<div class="steps-bar">'
    for i, label in enumerate(STEPS):
        css = "active" if i==S.step else ("done" if i<S.step else "")
        icon = "✓ " if i < S.step else ""
        html += f'<div class="step-item {css}"><span class="step-num">Etapa {i+1}</span>{icon}{label}</div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)

# ── Palavras que indicam linha de cabeçalho real ──────────────────────────────
PALAVRAS_CAB = [
    "data","valor","debito","débito","credito","crédito","documento","doc",
    "descricao","descrição","historico","histórico","lancamento","lançamento",
    "fornecedor","cliente","razao","razão","social","cnpj","cpf","nf","nota",
    "numero","número","ref","tipo","natureza","saldo","conta","agencia",
    "agência","vencimento","pagamento","competencia","competência","dcto","nsu",
]

def _score_cabecalho(linha) -> int:
    score = 0
    preenchidos = 0
    for v in linha:
        s = str(v).strip().lower() if v is not None else ""
        if not s or s in ("nan","none"): continue
        preenchidos += 1
        if any(p in s for p in PALAVRAS_CAB): score += 3
        if len(s) > 40: score -= 3          # frases longas = não é cabeçalho
        s2 = s.replace(".","").replace(",","").replace("/","").replace("-","")
        if s2.isdigit(): score -= 2         # número = não é cabeçalho
    return score + min(preenchidos, 8)

def _ler_planilha(file, ext: str) -> pd.DataFrame:
    """Lê qualquer planilha detectando automaticamente a linha do cabeçalho real."""
    import io

    # Para Excel: lê todas as abas e usa a que tem mais dados
    if ext in ("xlsx", "xls", "xlsm"):
        conteudo = file.read() if hasattr(file, "read") else open(file, "rb").read()
        # XLS antigo usa xlrd, XLSX usa openpyxl
        engine = "xlrd" if ext == "xls" else "openpyxl"
        try:
            xf = pd.ExcelFile(io.BytesIO(conteudo), engine=engine)
        except Exception:
            # fallback
            engine = "xlrd" if engine == "openpyxl" else "openpyxl"
            xf = pd.ExcelFile(io.BytesIO(conteudo), engine=engine)
        melhor_aba, melhor_linhas = xf.sheet_names[0], 0
        for aba in xf.sheet_names:
            try:
                df_teste = pd.read_excel(io.BytesIO(conteudo), sheet_name=aba,
                                         dtype=str, header=None, nrows=5, engine=engine)
                n = df_teste.notna().sum().sum()
                if n > melhor_linhas:
                    melhor_linhas, melhor_aba = n, aba
            except: pass
        bruto = pd.read_excel(io.BytesIO(conteudo), sheet_name=melhor_aba,
                              dtype=str, header=None, engine=engine)
    elif ext == "csv":
        bruto = pd.read_csv(file, dtype=str, encoding="utf-8-sig",
                            sep=None, engine="python", header=None)
    else:
        conteudo = file.read() if hasattr(file, "read") else open(file, "rb").read()
        bruto = pd.read_excel(io.BytesIO(conteudo), dtype=str, header=None)

    # Avalia as primeiras 15 linhas para achar o cabeçalho
    melhor_idx, melhor_score = 0, -999
    for i in range(min(15, len(bruto))):
        sc = _score_cabecalho(bruto.iloc[i].tolist())
        if sc > melhor_score:
            melhor_score, melhor_idx = sc, i

    # Reconstrói usando a linha encontrada como cabeçalho
    cols = []
    vistos = {}
    for i, v in enumerate(bruto.iloc[melhor_idx].tolist()):
        s = str(v).strip() if v is not None and str(v).strip() not in ("nan","None","") else ""
        if not s:
            s = f"_vazia_{i}"   # coluna sem nome → será removida a seguir
        if s in vistos:
            vistos[s] += 1
            s = f"{s}_{vistos[s]}"
        else:
            vistos[s] = 0
        cols.append(s)

    df = bruto.iloc[melhor_idx + 1:].copy()
    df.columns = cols
    df = df.reset_index(drop=True)

    # Remove colunas sem nome (vazias no cabeçalho) e colunas 100% vazias
    df = df.loc[:, ~df.columns.str.startswith("_vazia_")]
    df = df.loc[:, df.notna().any(axis=0)]
    return df

def header():
    st.markdown("""
    <div class="concil-header">
        <div style="font-size:2rem">⇄</div>
        <div>
            <h1>Conciliador Bancário</h1>
            <p>Conciliação automática entre Extrato Bancário e ERP — qualquer formato de planilha</p>
        </div>
    </div>""", unsafe_allow_html=True)

def _badge_dc(dc):
    if not dc or str(dc).strip() == "": return "—"
    dc = str(dc).strip()
    cor = "#085041" if dc=="C" else "#A32D2D" if dc=="D" else "#BA7517"
    bg  = "#E1F5EE" if dc=="C" else "#FCEBEB" if dc=="D" else "#FAEEDA"
    return f'<span style="background:{bg};color:{cor};padding:1px 7px;border-radius:10px;font-size:11px;font-weight:600">{dc}</span>'

def badge_html(status):
    MAP = {
        "Conciliado":    ("ok",   "✅ Conciliado"),
        "Divergente":    ("warn", "⚠️ Divergente"),
        "Nao Conciliado":("err",  "❌ Não Conciliado"),
        "So no Banco":   ("err",  "🏦 Só no Banco"),
        "So no ERP":     ("info", "🗄️ Só no ERP"),
    }
    cls, label = MAP.get(str(status), ("ok", str(status)))
    return f'<span class="badge badge-{cls}">{label}</span>'

def conf_html(conf):
    mapa = {"Alta":"ok","Media":"warn","Média":"warn","Baixa":"err","Muito Baixa":"err"}
    cls = mapa.get(str(conf),"ok")
    label = {"Media":"Média"}.get(str(conf), str(conf))  # corrige grafia
    if str(conf) in mapa:
        return f'<span class="badge badge-{cls}">{label}</span>'
    return f'<span style="color:#bbb">{label}</span>'

def fmt_valor(v):
    try:
        if v is None: return "—"
        f = float(v)
        if np.isnan(f) or np.isinf(f): return "—"
        s = f"{abs(f):,.2f}".replace(",","X").replace(".",",").replace("X",".")
        return f"R$ {'-' if f<0 else ''}{s}"
    except: return "—"

def fmt_data(d):
    try:
        if d is None or pd.isna(d): return "—"
        return pd.Timestamp(d).strftime("%d/%m/%Y")
    except: return str(d) if d else "—"

# ── Etapa 0 — Importar ────────────────────────────────────────────────────────
def etapa_importar():
    st.markdown('<div class="info-banner">📌 Importe qualquer planilha. O sistema detecta automaticamente onde começa o cabeçalho real, ignorando linhas de título ou observação.</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="section-title">🏦 Extrato Bancário</div>', unsafe_allow_html=True)
        f = st.file_uploader("Banco", type=["xlsx","xls","csv"], label_visibility="collapsed", key="up_banco")
        if f:
            try:
                ext = f.name.rsplit(".",1)[-1].lower()
                raw = _ler_planilha(f, ext)
                S.raw_banco = raw
                S.map_banco = detect_columns(raw)
                st.success(f"✅ **{f.name}** — {len(raw)} linhas, {len(raw.columns)} colunas detectadas")
                with st.expander("Pré-visualização (5 primeiras linhas)"):
                    st.dataframe(raw.head(5), use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")

    with col2:
        st.markdown('<div class="section-title">🗄️ Financeiro / ERP</div>', unsafe_allow_html=True)
        f = st.file_uploader("ERP", type=["xlsx","xls","csv"], label_visibility="collapsed", key="up_erp")
        if f:
            try:
                ext = f.name.rsplit(".",1)[-1].lower()
                raw = _ler_planilha(f, ext)
                S.raw_erp = raw
                S.map_erp = detect_columns(raw)
                st.success(f"✅ **{f.name}** — {len(raw)} linhas, {len(raw.columns)} colunas detectadas")
                with st.expander("Pré-visualização (5 primeiras linhas)"):
                    st.dataframe(raw.head(5), use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")

    st.divider()
    _, btn_col = st.columns([5,1])
    with btn_col:
        if st.button("Próximo →", type="primary", use_container_width=True):
            if S.raw_banco is None or S.raw_erp is None:
                st.warning("Importe os dois arquivos antes de continuar.")
            else:
                S.step = 1; st.rerun()

# ── Etapa 1 — Mapear colunas ──────────────────────────────────────────────────
CAMPOS = ["data","valor","debito","credito","descricao","documento","tipo"]
CAMPOS_LABEL = {
    "data":      "📅 Data",
    "valor":     "💰 Valor (coluna única — ex: Itaú com negativos)",
    "debito":    "➖ Débito (coluna separada)",
    "credito":   "➕ Crédito (coluna separada)",
    "descricao": "📝 Descrição / Histórico / Fornecedor",
    "documento": "🔢 Nº Documento / NF",
    "tipo":      "🔄 Tipo (D/C)",
}
CAMPOS_HELP = {
    "valor":   "Use quando débito e crédito estão na mesma coluna. Ex: -100 (débito) e 200 (crédito).",
    "debito":  "Use quando há uma coluna só para saídas/débitos.",
    "credito": "Use quando há uma coluna só para entradas/créditos.",
}

def etapa_mapear():
    st.markdown('<div class="info-banner">🗺️ Confirme o mapeamento automático. Os nomes mostrados são exatamente os das colunas da sua planilha.</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    for side, label, raw_key, map_key in [
        (col1, "🏦 Extrato Bancário", "raw_banco", "map_banco"),
        (col2, "🗄️ ERP / Financeiro", "raw_erp",   "map_erp"),
    ]:
        raw = S[raw_key]
        if raw is None: continue
        # Filtra colunas com nomes claramente inválidos (frases longas, não são cabeçalhos)
        cols_validas = [
            col for col in raw.columns
            if len(str(col)) <= 50 and not str(col).startswith("_vazia_")
        ]
        opcoes = ["(não usar)"] + cols_validas

        with side:
            st.markdown(f'<div class="section-title">{label}</div>', unsafe_allow_html=True)
            for campo in CAMPOS:
                atual = S[map_key].get(campo)
                idx = opcoes.index(atual) if atual in opcoes else 0
                escolha = st.selectbox(
                    CAMPOS_LABEL[campo], opcoes, index=idx,
                    key=f"sel_{map_key}_{campo}",
                    help=CAMPOS_HELP.get(campo),
                )
                S[map_key][campo] = None if escolha == "(não usar)" else escolha
                if campo == "valor":
                    st.markdown("<div style='border-top:1px dashed #ddd;margin:6px 0;font-size:11px;color:#aaa;text-align:center'>— ou colunas separadas —</div>", unsafe_allow_html=True)
                if campo == "credito":
                    st.markdown("<div style='border-top:1px dashed #ddd;margin:6px 0'></div>", unsafe_allow_html=True)

    st.divider()
    b1, _, b2 = st.columns([1,4,1])
    with b1:
        if st.button("← Voltar", use_container_width=True): S.step=0; st.rerun()
    with b2:
        if st.button("Próximo →", type="primary", use_container_width=True): S.step=2; st.rerun()

# ── Etapa 2 — Critérios ───────────────────────────────────────────────────────
def etapa_criterios():
    st.markdown("### ⚙️ Configure os critérios de conciliação")
    col1, col2 = st.columns(2)
    with col1:
        usar_valor = st.checkbox("Bater por **Valor**", value=True)
        tol_valor  = st.number_input("Tolerância de valor (R$)", min_value=0.0, value=0.01, step=0.01, disabled=not usar_valor, format="%.2f")
        usar_doc   = st.checkbox("Bater por **Nº Documento / NF**", value=True)
    with col2:
        usar_data  = st.checkbox("Bater por **Data**", value=True)
        tol_dias   = st.number_input("Tolerância de data (dias)", min_value=0, value=2, step=1, disabled=not usar_data)
        usar_desc  = st.checkbox("Bater por **Descrição** (fuzzy)", value=False)
        min_sim    = st.slider("Similaridade mínima (%)", 50, 100, 70, disabled=not usar_desc)

    st.info("💡 Prioridade automática: Valor + Data + Documento → Valor + Data → Valor apenas.")
    st.divider()
    b1, _, b2 = st.columns([1,4,1])
    with b1:
        if st.button("← Voltar", use_container_width=True): S.step=1; st.rerun()
    with b2:
        if st.button("▶ Executar Conciliação", type="primary", use_container_width=True):
            with st.spinner("Conciliando lançamentos..."):
                df_b = load_sheet(S.raw_banco, S.map_banco)
                df_e = load_sheet(S.raw_erp,   S.map_erp)
                result = conciliar(df_b, df_e,
                    tol_valor=tol_valor if usar_valor else 0,
                    tol_dias=tol_dias   if usar_data  else 0,
                    usar_documento=usar_doc,
                    usar_descricao=usar_desc,
                    min_similaridade=min_sim/100,
                )
            if isinstance(result, dict):
                S.resultado_completo = result
                S.df_result = result["resultado"]
                S.resumo    = result["resumo"]
            else:
                S.df_result = result
                S.resumo    = resumo(result)
                S.resultado_completo = {
                    "resultado": result, "resumo": S.resumo,
                    "agrupamentos":[], "sugestoes":[],
                    "duplicidades_banco": pd.DataFrame(),
                    "duplicidades_erp":   pd.DataFrame(),
                }
            S.step = 3; st.rerun()

# ── Etapa 3 — Resultado ───────────────────────────────────────────────────────
def etapa_resultado():
    r  = S.resumo
    df = S.df_result

    m1,m2,m3,m4,m5 = st.columns(5)
    for col, val, lbl, css in [
        (m1, r["total_banco"],  "Total Banco",     "metric-card"),
        (m2, r["conciliados"],  "Conciliados",     "metric-card metric-ok"),
        (m3, r["divergentes"],  "Divergentes",     "metric-card metric-warn"),
        (m4, r.get("so_banco", r.get("nao_conciliados",0)), "Só no Banco", "metric-card metric-err"),
        (m5, r["so_erp"],       "Só no ERP",       "metric-card metric-info"),
    ]:
        with col:
            st.markdown(f'<div class="{css}"><div class="metric-val">{val}</div><div class="metric-lbl">{lbl}</div></div>', unsafe_allow_html=True)

    pct = r["taxa_pct"]
    st.markdown(f"""
    <div style="margin-top:1rem">
        <div style="display:flex;justify-content:space-between;font-size:.82rem;color:#666;margin-bottom:4px">
            <span>Taxa de conciliação</span><span><b>{pct}%</b></span>
        </div>
        <div class="progress-outer"><div class="progress-inner" style="width:{pct}%"></div></div>
    </div>""", unsafe_allow_html=True)

    st.divider()
    filtro = st.radio("Filtrar:", ["Todos","Conciliado","Divergente","So no Banco","So no ERP"],
        format_func=lambda x: {"Todos":"Todos","Conciliado":"✅ Conciliado",
            "Divergente":"⚠️ Divergente","So no Banco":"🏦 Só no Banco",
            "So no ERP":"🗄️ Só no ERP"}.get(x,x),
        horizontal=True, label_visibility="collapsed")

    if filtro == "Todos":
        filtrado = df
    elif filtro == "So no Banco":
        filtrado = df[df["status"].isin(["So no Banco","Nao Conciliado"])]
    else:
        filtrado = df[df["status"] == filtro]

    linhas = ""
    for _, row in filtrado.iterrows():
        linhas += f"""<tr>
            <td>{fmt_data(row.get("banco_data"))}</td>
            <td>{fmt_valor(row.get("banco_valor"))}</td>
            <td>{_badge_dc(row.get("banco_dc",""))}</td>
            <td title="{row.get("banco_descricao") or ""}">{str(row.get("banco_descricao") or "—")[:30]}</td>
            <td>{row.get("banco_documento") or "—"}</td>
            <td>{fmt_data(row.get("erp_data"))}</td>
            <td>{fmt_valor(row.get("erp_valor"))}</td>
            <td>{_badge_dc(row.get("erp_dc",""))}</td>
            <td title="{row.get("erp_descricao") or ""}">{str(row.get("erp_descricao") or "—")[:30]}</td>
            <td>{row.get("erp_documento") or "—"}</td>
            <td>{badge_html(str(row.get("status","—")))}</td>
            <td>{conf_html(str(row.get("confianca","—")))}</td>
            <td>{fmt_valor(row.get("diferenca_valor"))}</td>
            <td>{str(int(row["diferenca_dias"]))+"d" if pd.notna(row.get("diferenca_dias")) else "—"}</td>
        </tr>"""

    st.markdown(f"""
    <div style="overflow-x:auto;border-radius:10px;border:1px solid #e8e8e8;max-height:420px;overflow-y:auto">
    <table class="result-table"><thead><tr>
        <th>Data Banco</th><th>Valor Banco</th><th>D/C</th><th>Descr. Banco</th><th>Doc. Banco</th>
        <th>Data ERP</th><th>Valor ERP</th><th>D/C</th><th>Descr. ERP</th><th>Doc. ERP</th>
        <th>Status</th><th>Confiança</th><th>Dif. Valor</th><th>Dif. Dias</th>
    </tr></thead><tbody>{linhas}</tbody></table></div>""", unsafe_allow_html=True)
    st.caption(f"Exibindo {len(filtrado)} de {len(df)} lançamentos")

    st.divider()
    b1,_,b2 = st.columns([1,4,1])
    with b1:
        if st.button("← Voltar", use_container_width=True): S.step=2; st.rerun()
    with b2:
        if st.button("Exportar →", type="primary", use_container_width=True): S.step=4; st.rerun()

# ── Etapa 4 — Exportar ────────────────────────────────────────────────────────
def etapa_exportar():
    r  = S.resumo
    df = S.df_result
    st.markdown("### 📤 Exportar resultados")

    col1,col2,col3,col4 = st.columns(4)
    with col1:
        st.markdown("#### 📊 Excel completo")
        st.caption("Abas: Resumo, Conciliados, Divergentes, Não Conciliados, Só no ERP")
        try:
            st.download_button("⬇ Baixar Excel", data=exportar_excel(df,r),
                file_name="conciliacao_bancaria.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary")
        except Exception as e: st.error(f"Erro: {e}")

    with col2:
        st.markdown("#### 📋 CSV Input ERP")
        st.caption("Lançamentos conciliados para reimportação")
        try:
            st.download_button("⬇ Baixar CSV", data=exportar_csv_erp(df),
                file_name="conciliados_input_erp.csv", mime="text/csv",
                use_container_width=True)
        except Exception as e: st.error(f"Erro: {e}")

    with col3:
        st.markdown("#### 📈 CSV Divergências")
        st.caption("Divergentes e não conciliados para revisão")
        try:
            div = df[df["status"].isin(["Divergente","So no Banco","So no ERP"])].copy()
            st.download_button("⬇ Baixar Divergências",
                data=div.to_csv(index=False,sep=";",encoding="utf-8-sig").encode("utf-8-sig"),
                file_name="divergencias.csv", mime="text/csv", use_container_width=True)
        except Exception as e: st.error(f"Erro: {e}")

    with col4:
        st.markdown("#### 📄 Relatório PDF")
        st.caption("Relatório executivo completo")
        try:
            st.download_button("⬇ Baixar PDF", data=exportar_pdf(S.resultado_completo),
                file_name="relatorio_conciliacao.pdf", mime="application/pdf",
                use_container_width=True, type="primary")
        except Exception as e: st.error(f"Erro: {e}")

    st.divider()
    st.markdown("#### 💰 Resumo Financeiro")
    fc1,fc2,fc3 = st.columns(3)
    with fc1: st.metric("Valor Conciliado",     fmt_valor(r.get("valor_conciliado",0)))
    with fc2: st.metric("Valor Divergente",     fmt_valor(r.get("valor_divergente",0)))
    with fc3: st.metric("Valor Não Conciliado", fmt_valor(r.get("valor_so_banco", r.get("valor_nao_conciliado",0))))

    st.divider()
    if st.button("← Nova conciliação"):
        for k in ["raw_banco","raw_erp","df_result","resumo","resultado_completo"]:
            S[k] = None
        S.map_banco = {}; S.map_erp = {}; S.step = 0; st.rerun()

# ── Roteador ──────────────────────────────────────────────────────────────────
header()
render_steps()
if   S.step == 0: etapa_importar()
elif S.step == 1: etapa_mapear()
elif S.step == 2: etapa_criterios()
elif S.step == 3: etapa_resultado()
elif S.step == 4: etapa_exportar()
