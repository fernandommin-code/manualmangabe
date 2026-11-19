#!/usr/bin/env python3
"""Folha de conferência do cadastro (PDF), uma por loja: o que entrou no Microvix para cada peça conferida.

Por marca e referência: Cód. Auxiliar, Código do produto (barras), descrição do Microvix (já com cor e
tamanho), preço de venda e quantidade da loja, com uma caixa para a vendedora marcar.
Serve para a loja conferir a etiqueta e o produto de cada peça antes/depois da entrada.

Uso: python3 scripts/folha_cadastro.py --conf conf/ --cadastrados Produtos_Cadastrados.xlsx \
        --planilha 11_17_Cadastro_de_Produtos.xlsx --titulo "Compra 17/11" --saida pasta/
"""
import argparse
import os
import sys
import tempfile

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gerar_entrada  # noqa: E402
import mangabe as mg  # noqa: E402

LOJAS = {1: "Loja 1 · Praiamar", 2: "Loja 2 · Balneário", 3: "Loja 3 · Miramar"}
VERDE = colors.HexColor("#0d6b5f")


def gerar(conf_dir, cadastrados, planilha, titulo, saida):
    ex = mg.ler_cadastrados(cadastrados).drop_duplicates("cod").set_index("cod")
    itens = mg.ler_planilha(planilha) if planilha else {}
    marca_da_ref = {mg.ref_key(r): it["marca"] for (_, r), it in itens.items()}
    with tempfile.TemporaryDirectory() as tmp:
        res = gerar_entrada.gerar(conf_dir, cadastrados, tmp, planilha)
        por_loja = {}
        for n in (1, 2, 3):
            p = os.path.join(tmp, f"Entrada_Loja{n}.csv")
            if os.path.exists(p):
                import csv
                por_loja[n] = [(r["Cod.Produto"], int(r["Quantidade"]))
                               for r in csv.DictReader(open(p, encoding="utf-8-sig"), delimiter=";")]
    os.makedirs(saida, exist_ok=True)
    st_t = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=15, leading=19, textColor=VERDE)
    st_s = ParagraphStyle("s", fontName="Helvetica", fontSize=9, leading=12, textColor=colors.HexColor("#555555"))
    st_m = ParagraphStyle("m", fontName="Helvetica-Bold", fontSize=11, leading=14, spaceBefore=8, textColor=VERDE)
    st_c = ParagraphStyle("c", fontName="Helvetica", fontSize=8, leading=10)
    arquivos = []
    for n, linhas in por_loja.items():
        grupos = {}
        for cod, q in linhas:
            p = ex.loc[cod]
            marca = marca_da_ref.get(mg.ref_key(p.ref), "OUTROS")
            grupos.setdefault(marca, []).append((p.ref, p.aux, cod, " ".join(p.desc.split()), mg.num(p.venda), q))
        total = sum(q for _, q in linhas)
        story = [Paragraph(f"Conferência do cadastro — {titulo}", st_t),
                 Paragraph(f"<b>{LOJAS[n]}</b> · {total} peças · confira em cada peça: referência, "
                           "cód. auxiliar e código da etiqueta. Marque OK quando bater.", st_s), Spacer(1, 4)]
        for marca in sorted(grupos):
            rows = sorted(grupos[marca], key=lambda x: (mg.ref_key(x[0]), x[2]))
            pecas = sum(r[5] for r in rows)
            story.append(Paragraph(f"{marca} — {pecas} peças", st_m))
            data = [["Referência", "Cód. auxiliar", "Código (etiqueta)", "Produto no Microvix", "Preço", "Qtd", "OK"]]
            for ref, aux, cod, desc, preco, q in rows:
                data.append([ref, aux or "—", cod, Paragraph(desc, st_c), f"R$ {preco:,.2f}".replace(",", "X")
                             .replace(".", ",").replace("X", "."), str(q), ""])
            tb = Table(data, colWidths=[20 * mm, 20 * mm, 24 * mm, 72 * mm, 19 * mm, 10 * mm, 9 * mm], repeatRows=1)
            tb.setStyle(TableStyle([
                ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8), ("FONT", (0, 1), (-1, -1), "Helvetica", 8),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e6f0ee")),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#b9c7c3")),
                ("ALIGN", (4, 1), (5, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
            story.append(tb)
        nome = os.path.join(saida, f"Folha_cadastro_{LOJAS[n].split('· ')[1].replace('á', 'a')}.pdf")
        SimpleDocTemplate(nome, pagesize=A4, leftMargin=12 * mm, rightMargin=12 * mm, topMargin=12 * mm,
                          bottomMargin=12 * mm, title=f"Conferência do cadastro {LOJAS[n]}").build(story)
        arquivos.append((nome, total))
    if res["nao_encontrados"]:
        print("Atenção: peças sem produto no Microvix ficaram fora da folha:", len(res["nao_encontrados"]))
    return arquivos


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--conf", required=True)
    ap.add_argument("--cadastrados", required=True)
    ap.add_argument("--planilha")
    ap.add_argument("--titulo", default="")
    ap.add_argument("--saida", required=True)
    a = ap.parse_args()
    for f, t in gerar(a.conf, a.cadastrados, a.planilha, a.titulo, a.saida):
        print(f, t)
