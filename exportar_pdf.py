"""
exportar_pdf.py
Gera relatório PDF executivo da conciliação bancária.
Dependência: reportlab  (pip install reportlab)
"""
import io
from datetime import datetime
import pandas as pd
import numpy as np

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether,
)
from reportlab.platypus import Flowable


# ── Paleta ────────────────────────────────────────────────────────────────────
VERDE       = colors.HexColor("#085041")
VERDE_CLARO = colors.HexColor("#E1F5EE")
AMARELO     = colors.HexColor("#BA7517")
AMAR_CLARO  = colors.HexColor("#FAEEDA")
VERMELHO    = colors.HexColor("#A32D2D")
VERM_CLARO  = colors.HexColor("#FCEBEB")
AZUL        = colors.HexColor("#185FA5")
AZUL_CLARO  = colors.HexColor("#E6F1FB")
CINZA       = colors.HexColor("#F1EFE8")
CINZA_ESC   = colors.HexColor("#444441")
BRANCO      = colors.white
PRETO       = colors.black


def _fmt(v):
    try:
        f = float(v)
        if np.isnan(f):
            return "—"
        s = f"{abs(f):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {'-' if f < 0 else ''}{s}"
    except Exception:
        return "—"


def _fdata(d):
    try:
        return pd.Timestamp(d).strftime("%d/%m/%Y")
    except Exception:
        return "—"


def _trunc(texto, n=35):
    s = str(texto) if texto and str(texto) not in ("nan", "None", "") else "—"
    return s[:n] + "…" if len(s) > n else s


# ── Estilos ───────────────────────────────────────────────────────────────────
def _estilos():
    base = getSampleStyleSheet()
    estilos = {}

    estilos["titulo"] = ParagraphStyle(
        "titulo", fontSize=20, textColor=BRANCO,
        fontName="Helvetica-Bold", alignment=TA_LEFT, leading=24,
    )
    estilos["subtitulo"] = ParagraphStyle(
        "subtitulo", fontSize=10, textColor=BRANCO,
        fontName="Helvetica", alignment=TA_LEFT, leading=14,
    )
    estilos["secao"] = ParagraphStyle(
        "secao", fontSize=12, textColor=VERDE,
        fontName="Helvetica-Bold", spaceBefore=14, spaceAfter=6,
    )
    estilos["corpo"] = ParagraphStyle(
        "corpo", fontSize=9, textColor=CINZA_ESC,
        fontName="Helvetica", leading=13,
    )
    estilos["label"] = ParagraphStyle(
        "label", fontSize=8, textColor=colors.HexColor("#888888"),
        fontName="Helvetica", alignment=TA_CENTER,
    )
    estilos["metrica"] = ParagraphStyle(
        "metrica", fontSize=18, textColor=VERDE,
        fontName="Helvetica-Bold", alignment=TA_CENTER,
    )
    estilos["cabec_tab"] = ParagraphStyle(
        "cabec_tab", fontSize=8, textColor=BRANCO,
        fontName="Helvetica-Bold", alignment=TA_CENTER,
    )
    estilos["celula"] = ParagraphStyle(
        "celula", fontSize=8, textColor=CINZA_ESC,
        fontName="Helvetica", alignment=TA_LEFT, leading=10,
    )
    estilos["celula_c"] = ParagraphStyle(
        "celula_c", fontSize=8, textColor=CINZA_ESC,
        fontName="Helvetica", alignment=TA_CENTER, leading=10,
    )
    return estilos


# ── Cabeçalho de página ───────────────────────────────────────────────────────
class _Cabecalho(Flowable):
    def __init__(self, largura):
        super().__init__()
        self.largura = largura
        self.altura  = 70

    def draw(self):
        c = self.canv
        # Fundo verde
        c.setFillColor(VERDE)
        c.rect(0, 0, self.largura, self.altura, fill=1, stroke=0)
        # Ícone
        c.setFillColor(VERDE_CLARO)
        c.setFont("Helvetica-Bold", 22)
        c.drawString(14, 28, "⇄")
        # Título
        c.setFont("Helvetica-Bold", 16)
        c.drawString(46, 42, "Conciliador Bancário")
        c.setFont("Helvetica", 9)
        c.setFillColor(colors.HexColor("#A8D8C8"))
        c.drawString(46, 28, "Relatório Executivo de Conciliação")
        # Data geração
        c.setFont("Helvetica", 8)
        c.setFillColor(VERDE_CLARO)
        data_str = datetime.now().strftime("Gerado em %d/%m/%Y às %H:%M")
        c.drawRightString(self.largura - 14, 28, data_str)


# ── Cartões de métricas ───────────────────────────────────────────────────────
def _cartoes_metricas(res: dict, largura: float, estilos: dict) -> Table:
    items = [
        ("Total Banco",       str(res["total_banco"]),    CINZA,      CINZA_ESC),
        ("Total ERP",         str(res["total_erp"]),      CINZA,      CINZA_ESC),
        ("Conciliados",       str(res["conciliados"]),    VERDE_CLARO, VERDE),
        ("Divergentes",       str(res["divergentes"]),    AMAR_CLARO,  AMARELO),
        ("Não Conciliados",   str(res.get("nao_conciliados", res.get("so_banco",0))),
                                                          VERM_CLARO,  VERMELHO),
        ("Taxa (%)",          f"{res['taxa_pct']}%",      VERDE_CLARO, VERDE),
    ]

    celulas = []
    for label, valor, bg, fg in items:
        t = Table(
            [[Paragraph(valor, ParagraphStyle("mv", fontSize=17, fontName="Helvetica-Bold",
                                              textColor=fg, alignment=TA_CENTER))],
             [Paragraph(label, ParagraphStyle("ml", fontSize=7.5, fontName="Helvetica",
                                              textColor=fg, alignment=TA_CENTER))]],
            colWidths=[largura / 6 - 6],
        )
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), bg),
            ("ROUNDEDCORNERS", [6]),
            ("TOPPADDING",    (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
        celulas.append(t)

    tabela = Table([celulas], colWidths=[largura / 6] * 6, hAlign="LEFT")
    tabela.setStyle(TableStyle([("LEFTPADDING", (0,0),(-1,-1), 3),
                                ("RIGHTPADDING",(0,0),(-1,-1), 3)]))
    return tabela


# ── Barra de progresso ────────────────────────────────────────────────────────
def _barra_progresso(taxa: float, largura: float) -> Table:
    pct    = min(max(taxa / 100, 0), 1)
    preenc = largura * pct
    vazio  = largura - preenc

    dados = [[""]]
    t = Table(dados, colWidths=[largura], rowHeights=[10])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), CINZA),
        ("ROUNDEDCORNERS", [5]),
    ]))

    if preenc > 0:
        dados2 = [[""]]
        t2 = Table(dados2, colWidths=[preenc], rowHeights=[10])
        t2.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), VERDE),
            ("ROUNDEDCORNERS", [5]),
        ]))
        linha = Table([[t2, ""]], colWidths=[preenc, vazio if vazio > 0 else 0.01],
                      rowHeights=[10])
        linha.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),
                                   ("RIGHTPADDING",(0,0),(-1,-1),0),
                                   ("TOPPADDING",(0,0),(-1,-1),0),
                                   ("BOTTOMPADDING",(0,0),(-1,-1),0)]))
        return linha
    return t


# ── Tabela genérica estilizada ────────────────────────────────────────────────
def _tabela_dados(cabecalhos, linhas, larguras, estilos, cor_cab=VERDE) -> Table:
    E = estilos
    cab = [Paragraph(h, E["cabec_tab"]) for h in cabecalhos]
    dados = [cab]

    for i, linha in enumerate(linhas):
        linha_fmt = []
        for cell in linha:
            if isinstance(cell, tuple):
                texto, cor_bg = cell
                linha_fmt.append(Paragraph(str(texto), E["celula_c"]))
            else:
                linha_fmt.append(Paragraph(str(cell), E["celula"]))
        dados.append(linha_fmt)

    t = Table(dados, colWidths=larguras, repeatRows=1)
    estilo = [
        ("BACKGROUND",    (0, 0), (-1, 0),  cor_cab),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  BRANCO),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 7.5),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#DDDDDD")),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [BRANCO, CINZA]),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]
    t.setStyle(TableStyle(estilo))
    return t


def _badge(texto, tipo="ok") -> str:
    cores_map = {
        "ok":   ("#085041", "#E1F5EE"),
        "warn": ("#633806", "#FAEEDA"),
        "err":  ("#791F1F", "#FCEBEB"),
        "info": ("#0C447C", "#E6F1FB"),
        "gray": ("#444441", "#F1EFE8"),
    }
    fg, bg = cores_map.get(tipo, cores_map["gray"])
    return (f'<font color="{fg}"><b> {texto} </b></font>')


# ── Função principal ──────────────────────────────────────────────────────────
def exportar_pdf(resultado_dict: dict) -> bytes:
    """
    Recebe o dict retornado por conciliar() e gera o PDF em bytes.
    """
    df        = resultado_dict["resultado"]
    res       = resultado_dict["resumo"]
    agrup     = resultado_dict.get("agrupamentos", [])
    sugestoes = resultado_dict.get("sugestoes", [])
    dup_banco = resultado_dict.get("duplicidades_banco", pd.DataFrame())
    dup_erp   = resultado_dict.get("duplicidades_erp",   pd.DataFrame())

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=1.5*cm,  bottomMargin=1.5*cm,
        title="Relatório de Conciliação Bancária",
    )
    largura = A4[0] - 3*cm
    E = _estilos()
    story = []

    # ── 1. Cabeçalho ──
    story.append(_Cabecalho(largura))
    story.append(Spacer(1, 14))

    # ── 2. Cartões de métricas ──
    story.append(Paragraph("Resumo Executivo", E["secao"]))
    story.append(_cartoes_metricas(res, largura, E))
    story.append(Spacer(1, 8))

    # Barra progresso
    taxa = res["taxa_pct"]
    story.append(Paragraph(f"Taxa de Conciliação: <b>{taxa}%</b>", E["corpo"]))
    story.append(Spacer(1, 4))
    story.append(_barra_progresso(taxa, largura))
    story.append(Spacer(1, 4))

    # Linha de resumo financeiro
    fin_dados = [
        ["Valor Total Banco", "Valor Conciliado", "Valor Divergente", "Valor Não Conc."],
        [_fmt(res["valor_total_banco"]), _fmt(res["valor_conciliado"]),
         _fmt(res["valor_divergente"]),  _fmt(res.get("valor_nao_conciliado",
                                               res.get("valor_so_banco", 0)))],
    ]
    fin_tab = Table(fin_dados, colWidths=[largura/4]*4)
    fin_tab.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0), VERDE),
        ("TEXTCOLOR",     (0, 0), (-1, 0), BRANCO),
        ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("BACKGROUND",    (0, 1), (-1, 1), VERDE_CLARO),
        ("FONTNAME",      (0, 1), (-1, 1), "Helvetica-Bold"),
        ("TEXTCOLOR",     (0, 1), (-1, 1), VERDE),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
    ]))
    story.append(Spacer(1, 8))
    story.append(fin_tab)
    story.append(Spacer(1, 14))
    story.append(HRFlowable(width=largura, color=CINZA, thickness=1))

    # ── 3. Lançamentos Conciliados ──
    conc = df[df["status"] == "Conciliado"]
    if len(conc) > 0:
        story.append(Paragraph(f"✅  Conciliados ({len(conc)})", E["secao"]))
        cab  = ["Data Banco", "Valor Banco", "Descr. Banco", "Data ERP",
                "Valor ERP",  "Descr. ERP",  "Score", "Tipo Match"]
        larg = [1.8*cm, 2.2*cm, 4*cm, 1.8*cm, 2.2*cm, 4*cm, 1.2*cm, 2.8*cm]
        linhas = []
        for _, r in conc.head(100).iterrows():
            linhas.append([
                _fdata(r.get("banco_data")),   _fmt(r.get("banco_valor")),
                _trunc(r.get("banco_descricao"), 28), _fdata(r.get("erp_data")),
                _fmt(r.get("erp_valor")),       _trunc(r.get("erp_descricao"), 28),
                str(r.get("score", "")),         str(r.get("tipo_match", "")),
            ])
        story.append(_tabela_dados(cab, linhas, larg, E, cor_cab=VERDE))
        if len(conc) > 100:
            story.append(Paragraph(f"… e mais {len(conc)-100} registros conciliados.", E["corpo"]))
        story.append(Spacer(1, 10))

    # ── 4. Divergentes ──
    div = df[df["status"] == "Divergente"]
    if len(div) > 0:
        story.append(Paragraph(f"⚠️  Divergentes ({len(div)})", E["secao"]))
        cab  = ["Data Banco", "Valor Banco", "Descr. Banco",
                "Data ERP",   "Valor ERP",   "Descr. ERP",
                "Dif. Valor", "Dif. Dias", "Divergência"]
        larg = [1.6*cm, 2*cm, 3.2*cm, 1.6*cm, 2*cm, 3.2*cm, 1.8*cm, 1.5*cm, 3.1*cm]
        linhas = []
        for _, r in div.head(100).iterrows():
            linhas.append([
                _fdata(r.get("banco_data")),    _fmt(r.get("banco_valor")),
                _trunc(r.get("banco_descricao"),22), _fdata(r.get("erp_data")),
                _fmt(r.get("erp_valor")),        _trunc(r.get("erp_descricao"),22),
                _fmt(r.get("diferenca_valor")),
                f"{int(r['diferenca_dias'])}d" if pd.notna(r.get("diferenca_dias")) else "—",
                _trunc(r.get("divergencias", ""), 28),
            ])
        story.append(_tabela_dados(cab, linhas, larg, E, cor_cab=colors.HexColor("#BA7517")))
        story.append(Spacer(1, 10))

    # ── 5. Não Conciliados (só banco) ──
    status_nc = "Nao Conciliado" if "Nao Conciliado" in df["status"].values else "So no Banco"
    nc = df[df["status"] == status_nc]
    if len(nc) > 0:
        story.append(Paragraph(f"❌  Não Conciliados — Só no Banco ({len(nc)})", E["secao"]))
        cab  = ["Data", "Valor", "Descrição", "Documento"]
        larg = [2.5*cm, 3*cm, 8*cm, 3.5*cm]
        linhas = [[_fdata(r.get("banco_data")), _fmt(r.get("banco_valor")),
                   _trunc(r.get("banco_descricao"), 50), str(r.get("banco_documento") or "—")]
                  for _, r in nc.head(100).iterrows()]
        story.append(_tabela_dados(cab, linhas, larg, E, cor_cab=VERMELHO))
        story.append(Spacer(1, 10))

    # ── 6. Só no ERP ──
    soerp = df[df["status"] == "So no ERP"]
    if len(soerp) > 0:
        story.append(Paragraph(f"🔵  Só no ERP ({len(soerp)})", E["secao"]))
        cab  = ["Data", "Valor", "Descrição", "Documento"]
        larg = [2.5*cm, 3*cm, 8*cm, 3.5*cm]
        linhas = [[_fdata(r.get("erp_data")), _fmt(r.get("erp_valor")),
                   _trunc(r.get("erp_descricao"), 50), str(r.get("erp_documento") or "—")]
                  for _, r in soerp.head(100).iterrows()]
        story.append(_tabela_dados(cab, linhas, larg, E, cor_cab=AZUL))
        story.append(Spacer(1, 10))

    # ── 7. Agrupamentos / Desmembramentos ──
    if agrup:
        story.append(PageBreak())
        story.append(Paragraph(f"🔗  Agrupamentos e Desmembramentos ({len(agrup)})", E["secao"]))
        for a in agrup:
            tipo = a["tipo"]
            cor  = VERDE if tipo == "Agrupamento" else AZUL
            if tipo == "Agrupamento":
                desc = (f"<b>{tipo}</b> — Banco: {_fmt(a['banco_valor'])} "
                        f"= ERP: {' + '.join(_fmt(v) for v in a['erp_valores'])} "
                        f"(soma: {_fmt(a['soma_erp'])}) | Score: {a['score']}")
            else:
                desc = (f"<b>{tipo}</b> — ERP: {_fmt(a['erp_valor'])} "
                        f"= Banco: {' + '.join(_fmt(v) for v in a['banco_valores'])} "
                        f"(soma: {_fmt(a['soma_banco'])}) | Score: {a['score']}")
            story.append(Paragraph(desc, E["corpo"]))
            story.append(Spacer(1, 4))
        story.append(Spacer(1, 10))

    # ── 8. Sugestões ──
    if sugestoes:
        story.append(Paragraph(f"💡  Possíveis Correspondências Não Confirmadas ({len(sugestoes)})",
                                E["secao"]))
        cab  = ["Data Banco", "Valor Banco", "Descr. Banco",
                "Data ERP",   "Valor ERP",   "Descr. ERP", "Score", "Motivo"]
        larg = [1.8*cm, 2*cm, 3.5*cm, 1.8*cm, 2*cm, 3.5*cm, 1.2*cm, 4.2*cm]
        linhas = []
        for s in sugestoes:
            linhas.append([
                _fdata(s.get("banco_data")),    _fmt(s.get("banco_valor")),
                _trunc(s.get("banco_descricao"),25), _fdata(s.get("erp_data")),
                _fmt(s.get("erp_valor")),        _trunc(s.get("erp_descricao"),25),
                str(s.get("score", "")),          _trunc(s.get("motivos",""), 35),
            ])
        story.append(_tabela_dados(cab, linhas, larg, E,
                                   cor_cab=colors.HexColor("#5A7FA8")))
        story.append(Spacer(1, 10))

    # ── 9. Duplicidades ──
    if len(dup_banco) > 0 or len(dup_erp) > 0:
        story.append(Paragraph("⚠️  Alertas de Duplicidade", E["secao"]))
        for df_dup, origem in [(dup_banco, "Banco"), (dup_erp, "ERP")]:
            if len(df_dup) > 0:
                story.append(Paragraph(f"Duplicidades no {origem}: {len(df_dup)} registros",
                                        E["corpo"]))
                cab  = ["Data", "Valor", "Descrição"]
                larg = [2.5*cm, 3*cm, 11.5*cm]
                col_data = "banco_data" if origem == "Banco" else "erp_data"
                linhas = []
                for _, r in df_dup.head(20).iterrows():
                    linhas.append([
                        _fdata(r.get("data")),
                        _fmt(r.get("valor")),
                        _trunc(r.get("descricao", ""), 60),
                    ])
                story.append(_tabela_dados(cab, linhas, larg, E,
                                           cor_cab=colors.HexColor("#BA7517")))
                story.append(Spacer(1, 6))

    # ── 10. Rodapé ──
    story.append(Spacer(1, 14))
    story.append(HRFlowable(width=largura, color=CINZA, thickness=1))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        f"Relatório gerado pelo <b>Conciliador Bancário</b> em "
        f"{datetime.now().strftime('%d/%m/%Y às %H:%M')} · "
        f"Score médio dos conciliados: <b>{res.get('score_medio', 0)}%</b>",
        ParagraphStyle("rodape", fontSize=7.5, textColor=colors.HexColor("#999999"),
                       alignment=TA_CENTER),
    ))

    doc.build(story)
    return buf.getvalue()
