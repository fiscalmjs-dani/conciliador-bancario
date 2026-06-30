import streamlit as st
import pandas as pd
import numpy as np
import hmac

from conciliador import detect_columns, load_sheet, conciliar, resumo
from exportar import exportar_excel, exportar_csv_erp
from exportar_pdf import exportar_pdf

# ─── Configuração da página ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Conciliador Bancário",
    page_icon="⇄",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ─── Proteção por senha ───────────────────────────────────────────────────────
def check_password():
    """Mostra uma tela de login simples. Retorna True se a senha estiver correta."""

    def password_entered():
        senha_digitada = st.session_state.get("password_input", "")
        senha_correta   = st.secrets.get("app_password", "")
        if senha_correta and hmac.compare_digest(senha_digitada, senha_correta):
            st.session_state["password_correct"] = True
            del st.session_state["password_input"]  # não guarda a senha em memória
        else:
            st.session_state["password_correct"] = False

    # Já autenticado nesta sessão
    if st.session_state.get("password_correct", False):
        return True

    st.markdown("""
    <div style="max-width:420px;margin:8rem auto 0;text-align:center">
        <div style="font-size:2.5rem">⇄</div>
        <h2 style="margin:0 0 0.3rem;color:#085041">Conciliador Bancário</h2>
        <p style="color:#888;font-size:0.9rem;margin-bottom:1.5rem">
            Acesso restrito — informe a senha para continuar
        </p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.text_input(
            "Senha", type="password",
            on_change=password_entered, key="password_input",
            label_visibility="collapsed", placeholder="Digite a senha de acesso",
        )
        if st.session_state.get("password_correct") is False:
            st.error("❌ Senha incorreta. Tente novamente.")

    return False


if not check_password():
    st.stop()

# ─── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Segoe UI', Arial, sans-serif; }
.stApp { background: #f8f9fb; }

.concil-header {
    background: linear-gradient(135deg, #085041 0%, #0F6E56 100%);
    color: white; padding: 1.5rem 2rem; border-radius: 12px;
    margin-bottom: 1.5rem; display: flex; align-items: center; gap: 1rem;
}
.concil-header h1 { margin: 0; font-size: 1.6rem; font-weight: 600; }
.concil-header p  { margin: 0; opacity: 0.85; font-size: 0.9rem; }

.steps-bar {
    display: flex; background: white; border-radius: 10px;
    overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 1.5rem;
}
.step-item {
    flex: 1; padding: 0.75rem 1rem; text-align: center;
    font-size: 0.82rem; color: #888; border-right: 1px solid #eee;
}
.step-item:last-child { border-right: none; }
.step-item.active { background: #E1F5EE; color: #085041; font-weight: 600; }
.step-item.done   { color: #0F6E56; }
.step-num { font-size: 0.7rem; display: block; color: #bbb; margin-bottom: 2px; }
.step-item.active .step-num { color: #1D9E75; }
.step-item.done   .step-num { color: #1D9E75; }

.metric-card {
    background: white; border-radius: 10px; padding: 1.2rem 1.5rem;
    text-align: center; box-shadow: 0 1px 4px rgba(0,0,0,.06);
}
.metric-val { font-size: 2rem; font-weight: 700; line-height: 1; }
.metric-lbl { font-size: 0.78rem; color: #888; margin-top: 4px; }
.metric-ok   .metric-val { color: #0F6E56; }
.metric-warn .metric-val { color: #BA7517; }
.metric-err  .metric-val { color: #A32D2D; }
.metric-info .metric-val { color: #185FA5; }

.badge {
    display: inline-block; padding: 2px 10px; border-radius: 20px;
    font-size: 0.78rem; font-weight: 500;
}
.badge-ok   { background: #E1F5EE; color: #085041; }
.badge-warn { background: #FAEEDA; color: #633806; }
.badge-err  { background: #FCEBEB; color: #791F1F; }
.badge-info { background: #E6F1FB; color: #0C447C; }

/* Tabela com cores por status */
.result-table { width: 100%; border-collapse: collapse; font-size: 0.83rem; }
.result-table th {
    background: #085041; color: white; padding: 8px 10px;
    text-align: left; font-weight: 500; position: sticky; top: 0; z-index: 1;
}
.result-table td {
    padding: 7px 10px; border-bottom: 1px solid #e8e8e8;
    white-space: nowrap; max-width: 180px;
    overflow: hidden; text-overflow: ellipsis;
}
.result-table tr.row-conciliado td   { background: #f0faf5; }
.result-table tr.row-divergente td   { background: #fdf6ea; }
.result-table tr.row-nao-conc td     { background: #fdf0f0; }
.result-table tr.row-so-erp td       { background: #f0f5fd; }
.result-table tr:hover td            { filter: brightness(0.96); }

.section-title {
    font-size: 0.9rem; font-weight: 600; color: #333;
    margin-bottom: 0.5rem; display: flex; align-items: center; gap: 6px;
}
.progress-outer {
    background: #e8e8e8; border-radius: 6px; height: 10px;
    margin: 6px 0 12px; overflow: hidden;
}
.progress-inner {
    height: 100%; background: linear-gradient(90deg, #0F6E56, #1D9E75);
    border-radius: 6px; transition: width .6s ease;
}
div[data-testid="stButton"] > button {
    border-radius: 8px; font-weight: 500; border: none; padding: 0.5rem 1.25rem;
}
.info-banner {
    background: #E6F1FB; border-left: 4px solid #378ADD;
    border-radius: 0 8px 8px 0; padding: 0.75rem 1rem;
    font-size: 0.85rem; color: #185FA5; margin-bottom: 1rem;
}
.legenda-cores {
    display: flex; gap: 1.2rem; flex-wrap: wrap;
    font-size: 0.8rem; margin-bottom: 0.5rem;
}
.legenda-item { display: flex; align-items: center; gap: 5px; }
.legenda-box  {
    width: 14px; height: 14px; border-radius: 3px; display: inline-block;
}
</style>
""", unsafe_allow_html=True)


# ─── Estado da sessão ─────────────────────────────────────────────────────────
def _init():
    defaults = {
        "step":      0,
        "map_banco": {},
        "map_erp":   {},
        "raw_banco": None,
        "raw_erp":   None,
        "df_result": None,
        "resumo":    None,
        "resultado_completo": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init()
S = st.session_state

STEPS = ["Importar", "Mapear colunas", "Critérios", "Resultado", "Exportar"]


# ─── Helpers UI ───────────────────────────────────────────────────────────────
def render_steps():
    html = '<div class="steps-bar">'
    for i, label in enumerate(STEPS):
        css  = "active" if i == S.step else ("done" if i < S.step else "")
        icon = "✓ " if i < S.step else ""
        html += f'<div class="step-item {css}"><span class="step-num">Etapa {i+1}</span>{icon}{label}</div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def header():
    col_h, col_btn = st.columns([6, 1])
    with col_h:
        st.markdown("""
        <div class="concil-header">
            <div style="font-size:2rem">⇄</div>
            <div>
                <h1>Conciliador Bancário</h1>
                <p>Conciliação automática entre Extrato Bancário e ERP — qualquer formato de planilha</p>
            </div>
        </div>""", unsafe_allow_html=True)
    with col_btn:
        st.write("")
        if st.button("🔒 Sair", use_container_width=True):
            st.session_state["password_correct"] = False
            st.rerun()


def badge_html(status: str) -> str:
    cls = {"Conciliado":"ok","Divergente":"warn","Nao Conciliado":"err","So no ERP":"info"}.get(status,"ok")
    label = {"Nao Conciliado":"Não Conciliado","So no ERP":"Só no ERP"}.get(status, status)
    return f'<span class="badge badge-{cls}">{label}</span>'


def conf_html(conf: str) -> str:
    cls = {"Alta":"ok","Media":"warn","Baixa":"err"}.get(conf,"ok")
    if conf in ("Alta","Media","Baixa"):
        label = {"Media":"Média"}.get(conf, conf)
        return f'<span class="badge badge-{cls}">{label}</span>'
    return f'<span style="color:#bbb">{conf}</span>'


def row_css(status: str) -> str:
    return {"Conciliado":"row-conciliado","Divergente":"row-divergente",
            "Nao Conciliado":"row-nao-conc","So no ERP":"row-so-erp"}.get(status,"")


def fmt_valor(v):
    try:
        if v is None: return "—"
        f = float(v)
        if np.isnan(f) or np.isinf(f): return "—"
        s = f"{abs(f):,.2f}".replace(",","X").replace(".",",").replace("X",".")
        return f"R$ {'-' if f < 0 else ''}{s}"
    except Exception:
        return "—"


def fmt_data(d):
    if d is None or pd.isna(d): return "—"
    try:
        return pd.Timestamp(d).strftime("%d/%m/%Y")
    except Exception:
        return str(d)


# ─── Etapa 0 — Importar ───────────────────────────────────────────────────────
def etapa_importar():
    st.markdown('<div class="info-banner">📌 Importe qualquer planilha sem se preocupar com a formatação. O sistema detecta as colunas automaticamente.</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="section-title">🏦 Extrato Bancário</div>', unsafe_allow_html=True)
        f_banco = st.file_uploader("Banco", type=["xlsx","xls","csv"],
                                   label_visibility="collapsed", key="up_banco")
        if f_banco:
            try:
                ext = f_banco.name.rsplit(".",1)[-1].lower()
                raw = pd.read_csv(f_banco, dtype=str, encoding="utf-8-sig", sep=None, engine="python") \
                      if ext == "csv" else pd.read_excel(f_banco, dtype=str)
                S.raw_banco = raw
                S.map_banco = detect_columns(raw)
                st.success(f"✅ **{f_banco.name}** — {len(raw)} linhas, {len(raw.columns)} colunas")
                with st.expander("Pré-visualização"):
                    st.dataframe(raw.head(5), use_container_width=True)
            except Exception as e:
                st.error(f"Erro: {e}")

    with col2:
        st.markdown('<div class="section-title">🗄️ Financeiro / ERP</div>', unsafe_allow_html=True)
        f_erp = st.file_uploader("ERP", type=["xlsx","xls","csv"],
                                 label_visibility="collapsed", key="up_erp")
        if f_erp:
            try:
                ext = f_erp.name.rsplit(".",1)[-1].lower()
                raw = pd.read_csv(f_erp, dtype=str, encoding="utf-8-sig", sep=None, engine="python") \
                      if ext == "csv" else pd.read_excel(f_erp, dtype=str)
                S.raw_erp = raw
                S.map_erp = detect_columns(raw)
                st.success(f"✅ **{f_erp.name}** — {len(raw)} linhas, {len(raw.columns)} colunas")
                with st.expander("Pré-visualização"):
                    st.dataframe(raw.head(5), use_container_width=True)
            except Exception as e:
                st.error(f"Erro: {e}")

    st.divider()
    _, btn_col = st.columns([5, 1])
    with btn_col:
        if st.button("Próximo →", type="primary", use_container_width=True):
            if S.raw_banco is None or S.raw_erp is None:
                st.warning("Importe os dois arquivos antes de continuar.")
            else:
                S.step = 1; st.rerun()


# ─── Etapa 1 — Mapear colunas ─────────────────────────────────────────────────
# Campos disponíveis para mapeamento — inclui débito e crédito separados
CAMPOS = ["data", "valor", "debito", "credito", "descricao", "documento", "tipo"]
CAMPOS_LABEL = {
    "data":      "Data",
    "valor":     "Valor (coluna única)",
    "debito":    "Débito (saída)",
    "credito":   "Crédito (entrada)",
    "descricao": "Descrição / Histórico",
    "documento": "Nº Documento",
    "tipo":      "Tipo (D/C)",
}
CAMPOS_HINT = {
    "valor":   "⚠️ Use esta OU Débito/Crédito, não os dois",
    "debito":  "Coluna de saídas/débitos",
    "credito": "Coluna de entradas/créditos",
}

def etapa_mapear():
    st.markdown('<div class="info-banner">🗺️ Mapeie as colunas. Se o extrato tiver <b>Débito</b> e <b>Crédito</b> separados, mapeie as duas e deixe "Valor" como (não usar).</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    opcoes_none = ["(não usar)"]

    for side, label, raw_key, map_key in [
        (col1, "🏦 Extrato Bancário", "raw_banco", "map_banco"),
        (col2, "🗄️ ERP / Financeiro", "raw_erp",   "map_erp"),
    ]:
        raw = S[raw_key]
        if raw is None:
            continue
        cols_disp = opcoes_none + list(raw.columns)

        with side:
            st.markdown(f'<div class="section-title">{label}</div>', unsafe_allow_html=True)
            for campo in CAMPOS:
                atual = S[map_key].get(campo)
                idx   = cols_disp.index(atual) if atual in cols_disp else 0
                hint  = CAMPOS_HINT.get(campo, "")
                lbl   = CAMPOS_LABEL[campo] + (f" — {hint}" if hint else "")
                escolha = st.selectbox(lbl, cols_disp, index=idx,
                                       key=f"sel_{map_key}_{campo}")
                S[map_key][campo] = None if escolha == "(não usar)" else escolha

            # Aviso se valor E débito/crédito forem mapeados ao mesmo tempo
            tem_valor  = bool(S[map_key].get("valor"))
            tem_debcred = bool(S[map_key].get("debito") or S[map_key].get("credito"))
            if tem_valor and tem_debcred:
                st.warning("⚠️ Você mapeou 'Valor' e também 'Débito/Crédito'. O sistema usará apenas 'Valor'. Deixe 'Valor' como (não usar) para usar as colunas separadas.")

    st.divider()
    b1, _, b2 = st.columns([1, 4, 1])
    with b1:
        if st.button("← Voltar", use_container_width=True):
            S.step = 0; st.rerun()
    with b2:
        if st.button("Próximo →", type="primary", use_container_width=True):
            S.step = 2; st.rerun()


# ─── Etapa 2 — Critérios ──────────────────────────────────────────────────────
def etapa_criterios():
    st.markdown("### ⚙️ Configure os critérios de conciliação")
    col1, col2 = st.columns(2)
    with col1:
        usar_valor = st.checkbox("Bater por **Valor**", value=True)
        tol_valor  = st.number_input("Tolerância de valor (R$)", min_value=0.0, value=0.01,
                                     step=0.01, disabled=not usar_valor, format="%.2f")
        usar_doc   = st.checkbox("Bater por **Nº Documento / NF**", value=True)
    with col2:
        usar_data  = st.checkbox("Bater por **Data**", value=True)
        tol_dias   = st.number_input("Tolerância de data (dias)", min_value=0, value=2,
                                     step=1, disabled=not usar_data)
        usar_desc  = st.checkbox("Bater por **Descrição** (fuzzy)", value=False)
        min_sim    = st.slider("Similaridade mínima (%)", 50, 100, 70, disabled=not usar_desc)

    st.info("💡 Quanto mais critérios combinados, maior a confiança do match.")

    st.divider()
    b1, _, b2 = st.columns([1, 4, 1])
    with b1:
        if st.button("← Voltar", use_container_width=True):
            S.step = 1; st.rerun()
    with b2:
        if st.button("▶ Executar Conciliação", type="primary", use_container_width=True):
            with st.spinner("Conciliando lançamentos..."):
                df_b = load_sheet(S.raw_banco, S.map_banco)
                df_e = load_sheet(S.raw_erp,   S.map_erp)
                result_dict = conciliar(
                    df_b, df_e,
                    tol_valor=tol_valor if usar_valor else 0,
                    tol_dias=tol_dias   if usar_data  else 0,
                    usar_documento=usar_doc,
                    usar_descricao=usar_desc,
                    min_similaridade=min_sim / 100,
                )
                S.df_result = result_dict["resultado"]
                S.resumo    = result_dict["resumo"]
                S.resultado_completo = result_dict
            S.step = 3; st.rerun()


# ─── Etapa 3 — Resultado ──────────────────────────────────────────────────────
def etapa_resultado():
    r  = S.resumo
    df = S.df_result

    # Métricas
    m1, m2, m3, m4, m5 = st.columns(5)
    for col, val, lbl, css in [
        (m1, r["total_banco"],     "Total Banco",     "metric-card"),
        (m2, r["conciliados"],     "Conciliados",     "metric-card metric-ok"),
        (m3, r["divergentes"],     "Divergentes",     "metric-card metric-warn"),
        (m4, r["nao_conciliados"], "Não Conciliados", "metric-card metric-err"),
        (m5, r["so_erp"],          "Só no ERP",       "metric-card metric-info"),
    ]:
        with col:
            st.markdown(f'<div class="{css}"><div class="metric-val">{val}</div>'
                        f'<div class="metric-lbl">{lbl}</div></div>', unsafe_allow_html=True)

    pct = r["taxa_pct"]
    st.markdown(f"""
    <div style="margin-top:1rem">
        <div style="display:flex;justify-content:space-between;font-size:.82rem;color:#666;margin-bottom:4px">
            <span>Taxa de conciliação</span><span><b>{pct}%</b></span>
        </div>
        <div class="progress-outer"><div class="progress-inner" style="width:{pct}%"></div></div>
    </div>""", unsafe_allow_html=True)

    st.divider()

    # Legenda de cores
    st.markdown("""
    <div class="legenda-cores">
        <span class="legenda-item"><span class="legenda-box" style="background:#d4edda;border:1px solid #0F6E56"></span>Conciliado</span>
        <span class="legenda-item"><span class="legenda-box" style="background:#fff3cd;border:1px solid #BA7517"></span>Divergente</span>
        <span class="legenda-item"><span class="legenda-box" style="background:#f8d7da;border:1px solid #A32D2D"></span>Não Conciliado</span>
        <span class="legenda-item"><span class="legenda-box" style="background:#cce5ff;border:1px solid #185FA5"></span>Só no ERP</span>
    </div>""", unsafe_allow_html=True)

    # Filtros
    FILTRO_OPTS = {
        "Todos": None,
        "Conciliado": "Conciliado",
        "Divergente": "Divergente",
        "Não Conciliado": "Nao Conciliado",
        "Só no ERP": "So no ERP",
    }
    filtro_label = st.radio("Filtrar:", list(FILTRO_OPTS.keys()),
                            horizontal=True, label_visibility="collapsed")
    filtro_val = FILTRO_OPTS[filtro_label]
    filtrado   = df if filtro_val is None else df[df["status"] == filtro_val]

    # Detecta se há colunas de débito/crédito
    tem_debcred_banco = "banco_debito" in df.columns and df["banco_debito"].notna().any()
    tem_debcred_erp   = "erp_debito"   in df.columns and df["erp_debito"].notna().any()

    # Monta cabeçalho dinâmico
    th_banco = "<th>Data Banco</th><th>Débito Banco</th><th>Crédito Banco</th><th>Descr. Banco</th><th>Doc. Banco</th>" \
               if tem_debcred_banco else \
               "<th>Data Banco</th><th>Valor Banco</th><th>Descr. Banco</th><th>Doc. Banco</th>"
    th_erp   = "<th>Data ERP</th><th>Débito ERP</th><th>Crédito ERP</th><th>Descr. ERP</th><th>Doc. ERP</th>" \
               if tem_debcred_erp else \
               "<th>Data ERP</th><th>Valor ERP</th><th>Descr. ERP</th><th>Doc. ERP</th>"

    linhas = ""
    for _, row in filtrado.iterrows():
        status = str(row.get("status", ""))
        css    = row_css(status)

        if tem_debcred_banco:
            td_banco = (f"<td>{fmt_data(row.get('banco_data'))}</td>"
                        f"<td>{fmt_valor(row.get('banco_debito'))}</td>"
                        f"<td>{fmt_valor(row.get('banco_credito'))}</td>"
                        f"<td title='{row.get('banco_descricao') or ''}'>{str(row.get('banco_descricao') or '—')[:30]}</td>"
                        f"<td>{row.get('banco_documento') or '—'}</td>")
        else:
            td_banco = (f"<td>{fmt_data(row.get('banco_data'))}</td>"
                        f"<td>{fmt_valor(row.get('banco_valor'))}</td>"
                        f"<td title='{row.get('banco_descricao') or ''}'>{str(row.get('banco_descricao') or '—')[:30]}</td>"
                        f"<td>{row.get('banco_documento') or '—'}</td>")

        if tem_debcred_erp:
            td_erp = (f"<td>{fmt_data(row.get('erp_data'))}</td>"
                      f"<td>{fmt_valor(row.get('erp_debito'))}</td>"
                      f"<td>{fmt_valor(row.get('erp_credito'))}</td>"
                      f"<td title='{row.get('erp_descricao') or ''}'>{str(row.get('erp_descricao') or '—')[:30]}</td>"
                      f"<td>{row.get('erp_documento') or '—'}</td>")
        else:
            td_erp = (f"<td>{fmt_data(row.get('erp_data'))}</td>"
                      f"<td>{fmt_valor(row.get('erp_valor'))}</td>"
                      f"<td title='{row.get('erp_descricao') or ''}'>{str(row.get('erp_descricao') or '—')[:30]}</td>"
                      f"<td>{row.get('erp_documento') or '—'}</td>")

        dif_dias_str = f"{row.get('diferenca_dias')}d" if row.get('diferenca_dias') is not None else "—"

        linhas += (f"<tr class='{css}'>"
                   f"{td_banco}{td_erp}"
                   f"<td>{badge_html(status)}</td>"
                   f"<td>{conf_html(str(row.get('confianca','—')))}</td>"
                   f"<td>{fmt_valor(row.get('diferenca_valor'))}</td>"
                   f"<td>{dif_dias_str}</td>"
                   f"</tr>")

    st.markdown(f"""
    <div style="overflow-x:auto;border-radius:10px;border:1px solid #e8e8e8;max-height:430px;overflow-y:auto">
    <table class="result-table">
        <thead><tr>{th_banco}{th_erp}
            <th>Status</th><th>Confiança</th><th>Dif. Valor</th><th>Dif. Dias</th>
        </tr></thead>
        <tbody>{linhas}</tbody>
    </table>
    </div>""", unsafe_allow_html=True)

    st.caption(f"Exibindo {len(filtrado)} de {len(df)} lançamentos")

    st.divider()
    b1, _, b2 = st.columns([1, 4, 1])
    with b1:
        if st.button("← Voltar", use_container_width=True):
            S.step = 2; st.rerun()
    with b2:
        if st.button("Exportar →", type="primary", use_container_width=True):
            S.step = 4; st.rerun()


# ─── Etapa 4 — Exportar ───────────────────────────────────────────────────────
def etapa_exportar():
    r  = S.resumo
    df = S.df_result

    st.markdown("### 📤 Exportar resultados")
    st.markdown(f"Conciliação concluída: **{r['conciliados']} conciliados**, "
                f"**{r['divergentes']} divergentes**, **{r['nao_conciliados']} não conciliados**.")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown("#### 📊 Excel completo")
        st.caption("Planilha com abas: Resumo, Conciliados, Divergentes, Não Conciliados, Só no ERP, Todos")
        try:
            xlsx_bytes = exportar_excel(df, r)
            st.download_button(
                "⬇ Baixar Excel", data=xlsx_bytes,
                file_name="conciliacao_bancaria.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary",
            )
        except Exception as e:
            st.error(f"Erro ao gerar Excel: {e}")

    with col2:
        st.markdown("#### 📋 CSV — Input ERP")
        st.caption("Lançamentos conciliados para reimportação no ERP")
        try:
            csv_bytes = exportar_csv_erp(df)
            st.download_button(
                "⬇ Baixar CSV", data=csv_bytes,
                file_name="conciliados_input_erp.csv",
                mime="text/csv", use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erro ao gerar CSV: {e}")

    with col3:
        st.markdown("#### 📈 CSV — Divergências")
        st.caption("Divergentes e não conciliados para revisão manual")
        try:
            div = df[df["status"].isin(["Divergente","Nao Conciliado","So no ERP"])].copy()
            csv_div = div.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                "⬇ Baixar Divergências", data=csv_div,
                file_name="divergencias_para_revisao.csv",
                mime="text/csv", use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erro ao gerar CSV: {e}")

    with col4:
        st.markdown("#### 📄 Relatório PDF")
        st.caption("Relatório executivo completo com métricas e tabelas")
        try:
            pdf_bytes = exportar_pdf(S.resultado_completo)
            st.download_button(
                "⬇ Baixar PDF", data=pdf_bytes,
                file_name="relatorio_conciliacao.pdf",
                mime="application/pdf", use_container_width=True, type="primary",
            )
        except Exception as e:
            st.error(f"Erro ao gerar PDF: {e}")

    st.divider()
    st.markdown("#### 💰 Resumo Financeiro")
    fc1, fc2, fc3 = st.columns(3)
    with fc1: st.metric("Valor Conciliado",     fmt_valor(r["valor_conciliado"]))
    with fc2: st.metric("Valor Divergente",      fmt_valor(r["valor_divergente"]))
    with fc3: st.metric("Valor Não Conciliado",  fmt_valor(r["valor_nao_conciliado"]))

    st.divider()
    b1, _, b2 = st.columns([1, 4, 1])
    with b1:
        if st.button("← Voltar ao Resultado", use_container_width=True):
            S.step = 3; st.rerun()
    with b2:
        if st.button("🔄 Nova Conciliação", use_container_width=True):
            for k in list(S.keys()):
                del S[k]
            st.rerun()


# ─── Roteador ─────────────────────────────────────────────────────────────────
header()
render_steps()

if   S.step == 0: etapa_importar()
elif S.step == 1: etapa_mapear()
elif S.step == 2: etapa_criterios()
elif S.step == 3: etapa_resultado()
elif S.step == 4: etapa_exportar()
