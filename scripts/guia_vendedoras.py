#!/usr/bin/env python3
"""Guia de uma página para as vendedoras usarem o app de conferência. Uso: python3 scripts/guia_vendedoras.py saida.pdf"""
import sys

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

URL = "https://claude.ai/artifact/Bw297fhPHqWy35JJEU1hds"
VERDE = colors.HexColor("#0d6b5f")

t = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=VERDE)
h = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=11.5, leading=15, spaceBefore=8, textColor=VERDE)
p = ParagraphStyle("p", fontName="Helvetica", fontSize=10, leading=13.5)
b = ParagraphStyle("b", parent=p, leftIndent=12, bulletIndent=2)
small = ParagraphStyle("s", parent=p, fontSize=8.5, textColor=colors.HexColor("#555555"))


def passos(lista):
    rows = [[Paragraph(f"<b>{i}</b>", ParagraphStyle("n", parent=p, alignment=1, textColor=colors.white)),
             Paragraph(x, p)] for i, x in enumerate(lista, 1)]
    tb = Table(rows, colWidths=[9 * mm, 165 * mm])
    tb.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, -1), VERDE), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor("#d9dfdc")),
                            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    return tb


def gerar(path):
    s = [Paragraph("Conferência de mercadoria no app — guia rápido", t), Spacer(1, 4),
         Paragraph(f"Link: <b>{URL}</b> — abrir no celular, logada na sua conta Claude. "
                   "Se aparecer “Somente leitura”, peça ao Fernando acesso de <b>Colaborador</b>.", p),
         Paragraph("Começar", h),
         passos(["No topo, escolha a <b>compra</b> (ex.: Compra 10/11).",
                 "Toque na sua <b>loja</b>: Loja 1 Praiamar, Loja 2 Balneário ou Loja 3 Miramar.",
                 "Abra a <b>marca</b> que está conferindo. Aparecem só as referências que vieram para a sua loja, "
                 "com a quantidade do romaneio."]),
         Paragraph("Conferir cada referência", h),
         passos(["Toque na referência. Escreva a <b>cor</b> (escolha da lista, por extenso: OFF WHITE, PRETO, "
                 "VERMELHO…). Se aparecer <i>“já no sistema: …”</i>, use o mesmo nome.",
                 "Toque <b>+</b> em cada tamanho até chegar na quantidade. Errou? <b>−</b>.",
                 "Outra cor da mesma referência: <b>+ cor</b>. Jeans numerado ou tamanho único: botão <b>grade</b>. "
                 "Tamanho que não aparece: <b>+ tamanho</b>.",
                 "Toque <b>Salvar</b>. Se o total não bater com o romaneio, escreva o motivo na "
                 "<b>observação</b> (ex.: “faltou 1 P preto”, “veio 1 M a mais”)."]),
         Paragraph("Terminar a marca", h),
         passos(["Peça que veio e não está na lista: no fim da marca, <b>+ referência</b> (escreva a referência "
                 "da etiqueta) e confira como as outras.",
                 "Quando todas as referências estiverem verdes (ok) ou com observação, toque <b>Concluir marca</b>."
                 " Só depois disso o cadastro e a entrada são gerados."]),
         Paragraph("Dicas", h),
         Paragraph("• Confira pela <b>etiqueta da peça</b> (referência, cor e tamanho), não pela embalagem.", b),
         Paragraph("• Duas pessoas podem conferir ao mesmo tempo; o app salva na hora e mostra para todo mundo.", b),
         Paragraph("• Na lista: <b>verde</b> = bateu · <b>vermelho</b> = falta ou sobra · <b>amarelo</b> = pendente.", b),
         Paragraph("• Reposição: se a cor não estiver em “já no sistema”, tudo bem — é cor nova, escreva normalmente.", b),
         Spacer(1, 10),
         Paragraph("Dúvida ou problema no app: fale com o Fernando.", small)]
    SimpleDocTemplate(path, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm, topMargin=15 * mm,
                      bottomMargin=12 * mm, title="Guia da conferência no app").build(s)


if __name__ == "__main__":
    gerar(sys.argv[1])
