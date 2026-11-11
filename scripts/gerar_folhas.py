#!/usr/bin/env python3
"""Folha de conferência já preenchida, uma por loja, a partir da planilha do dia.

Vem impresso: marca, referência, descrição e quantidade do romaneio daquela loja.
As vendedoras só escrevem a cor e as quantidades por tamanho (uma linha por cor).

Uso: python3 scripts/gerar_folhas.py 11_10_Cadastro_de_Produtos.xlsx saida/ [--data 10/11]
Saída: Folha_Conferencia_Loja1.pdf, Loja2, Loja3 (A4 deitado)
"""
import argparse
import os
import sys

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mangabe as mg  # noqa: E402

LOJAS = ["Loja 1 · Praiamar", "Loja 2 · Balneário", "Loja 3 · Miramar"]
TAMS = ["XPP\n34", "PP\n36", "P\n38", "M\n40", "G\n42", "GG\n44"]
CINZA = colors.HexColor("#9a9a9a")
FUNDO = colors.HexColor("#ececec")
LINHA = colors.HexColor("#444444")
CORES = mg.cores_tab()
TAMSET = set(mg.tamanhos())

st_desc = ParagraphStyle("d", fontName="Helvetica", fontSize=8.5, leading=10)
st_ref = ParagraphStyle("r", fontName="Helvetica-Bold", fontSize=10.5, leading=12)
st_cont = ParagraphStyle("c", fontName="Helvetica", fontSize=7, leading=8, textColor=CINZA)
st_tit = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=15, leading=18)
st_sub = ParagraphStyle("s", fontName="Helvetica", fontSize=9, leading=12)

# colunas: Ref | Descrição | Rom. | Cor | 6 tamanhos | Outro | Total | Obs
LARG = [24 * mm, 56 * mm, 13 * mm, 36 * mm] + [13 * mm] * 6 + [13 * mm, 14 * mm, 39 * mm]  # 273 mm
ALT = 7.8 * mm


def linhas_por_ref(q, unico):
    if unico:
        return 1
    return 1 if q <= 2 else 2 if q <= 9 else 3


def cabecalho_tabela():
    return ["Ref.", "Descrição", "Rom.", "Cor (por extenso)"] + TAMS + ["Outro", "Total", "Obs."]


def bloco_marca(marca, itens, loja, ex=None):
    unico = mg.norm(marca) == "VICKY"
    dados = [[Paragraph(f"<b>{marca}</b>", ParagraphStyle("m", fontName="Helvetica-Bold", fontSize=11)),
              "", "", ""] + [""] * 8 + [f"esperado: {sum(i['lj'][loja] for i in itens)} pç"],
             cabecalho_tabela()]
    estilo = [
        ("SPAN", (0, 0), (11, 0)), ("BACKGROUND", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (12, 0), (12, 0), "Helvetica-Bold"), ("FONTSIZE", (12, 0), (12, 0), 9),
        ("ALIGN", (12, 0), (12, 0), "RIGHT"),
        ("BACKGROUND", (0, 1), (-1, 1), FUNDO), ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 1), (-1, 1), 7.5), ("LEADING", (0, 1), (-1, 1), 8.5),
        ("ALIGN", (2, 1), (-1, 1), "CENTER"),
        ("LINEBELOW", (0, 0), (-1, 0), 1.2, colors.black),
    ]
    r = 2
    for it in itens:
        q = it["lj"][loja]
        n = linhas_por_ref(q, unico)
        for k in range(n):
            if k == 0:
                desc = it["desc"]
                if ex is not None:
                    tm = mg.marcas().get(mg.norm(marca))
                    dig = mg.colecoes().get(mg.norm(it["col"]))
                    aux = mg.aux_existente(ex, it["ref"], dig, tm["id"]) if tm and dig else None
                    if aux:
                        cs = sorted({mg.separar(d, CORES, TAMSET)[1] or "?" for d in ex.desc[ex.aux == aux]})
                        desc += f'<br/><font size="6.5" color="#777777">já no sistema: {", ".join(cs)}</font>'
                linha = [Paragraph(it["ref"], st_ref), Paragraph(desc, st_desc), str(q)]
            else:
                linha = [Paragraph(it["ref"], st_cont), "", ""]
            linha += [""] * 10
            dados.append(linha)
        estilo += [("LINEABOVE", (0, r), (-1, r), 1.0, colors.black)]
        if n > 1:
            estilo += [("SPAN", (1, r), (1, r + n - 1)), ("SPAN", (2, r), (2, r + n - 1)),
                       ("VALIGN", (1, r), (2, r + n - 1), "MIDDLE")]
        if unico:
            estilo += [("SPAN", (4, r), (10, r))]
            dados[r][4] = "qtd recebida: ______"
            estilo += [("TEXTCOLOR", (4, r), (4, r), CINZA), ("FONTSIZE", (4, r), (4, r), 7.5)]
        r += n
    # linhas para peças fora do romaneio
    for k in range(2):
        dados.append([Paragraph("fora do rom.", st_cont)] + [""] * 12)
    estilo += [("LINEABOVE", (0, r), (-1, r), 1.0, colors.black),
               ("BACKGROUND", (0, r), (-1, r + 1), colors.HexColor("#f7f7f7"))]
    estilo += [
        ("GRID", (0, 1), (-1, -1), 0.5, LINHA),
        ("BOX", (0, 1), (-1, -1), 1.2, colors.black),
        ("LINEBEFORE", (4, 1), (4, -1), 1.0, colors.black),
        ("LINEBEFORE", (10, 1), (10, -1), 1.0, colors.black),
        ("VALIGN", (0, 2), (-1, -1), "MIDDLE"),
        ("ALIGN", (2, 2), (2, -1), "CENTER"),
        ("FONTNAME", (2, 2), (2, -1), "Helvetica-Bold"), ("FONTSIZE", (2, 2), (2, -1), 11),
        ("BACKGROUND", (2, 2), (2, r - 1), FUNDO),
        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]
    alturas = [7 * mm, 8 * mm] + [ALT] * (len(dados) - 2)
    t = Table(dados, colWidths=LARG, rowHeights=alturas, repeatRows=2)
    t.setStyle(TableStyle(estilo))
    return t


def gerar(planilha, saida, data=None, cadastrados=None):
    itens = mg.ler_planilha(planilha)
    ex = mg.ler_cadastrados(cadastrados) if cadastrados else None
    marcas = {}
    for it in itens.values():
        marcas.setdefault(it["marca"], []).append(it)
    os.makedirs(saida, exist_ok=True)
    arquivos = []
    for loja, nome in enumerate(LOJAS):
        blocos = []
        total = 0
        for marca, its in marcas.items():
            its = [i for i in its if i["lj"][loja] > 0]
            if not its:
                continue
            total += sum(i["lj"][loja] for i in its)
            t = bloco_marca(marca, its, loja, ex)
            blocos.append(KeepTogether([t]) if len(its) <= 8 else t)
            blocos.append(Spacer(1, 5 * mm))
        if not blocos:
            continue
        path = os.path.join(saida, f"Folha_Conferencia_Loja{loja + 1}.pdf")
        titulo = f"Compra {data}" if data else "Compra"

        def topo(canvas, doc, nome=nome, titulo=titulo, total=total):
            canvas.saveState()
            w, h = landscape(A4)
            canvas.setFont("Helvetica-Bold", 14)
            canvas.drawString(12 * mm, h - 12 * mm, "CONFERÊNCIA DE MERCADORIA")
            x = 12 * mm + canvas.stringWidth("CONFERÊNCIA DE MERCADORIA ", "Helvetica-Bold", 14)
            canvas.setFont("Helvetica", 14)
            canvas.drawString(x, h - 12 * mm, "mangabê")
            canvas.setFont("Helvetica-Bold", 12)
            canvas.drawRightString(w - 12 * mm, h - 12 * mm, f"{nome}   ·   {titulo}")
            canvas.setFont("Helvetica", 8.5)
            canvas.drawString(12 * mm, h - 18 * mm,
                              "Uma linha por cor, escrita por extenso (OFF WHITE, PRETO, BEGE…).  Se o total não bater com "
                              "“Rom.”, explique em Obs.  Peça que não está na lista: linhas “fora do rom.”, com a referência.")
            canvas.drawRightString(w - 12 * mm, h - 7 * mm, f"Total esperado na loja: {total} peças")
            canvas.setFont("Helvetica", 8)
            canvas.drawString(12 * mm, 7 * mm, "Conferido por: ______________________   Data: ____/____")
            canvas.drawRightString(w - 12 * mm, 7 * mm, f"{nome} · página {doc.page}")
            canvas.restoreState()

        doc = SimpleDocTemplate(path, pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
                                topMargin=22 * mm, bottomMargin=12 * mm,
                                title=f"Conferência {nome} {titulo}")
        doc.build(blocos, onFirstPage=topo, onLaterPages=topo)
        arquivos.append(path)
        print(f"{nome}: {total} peças -> {path}")
    return arquivos


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("planilha")
    ap.add_argument("saida")
    ap.add_argument("--data")
    ap.add_argument("--cadastrados", help="relatório Produtos Cadastrados: mostra as cores já existentes nas reposições")
    a = ap.parse_args()
    gerar(a.planilha, a.saida, a.data, a.cadastrados)
