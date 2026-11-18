#!/usr/bin/env python3
"""Planilha de etiquetas a partir das listas de entrada (Entrada_LojaN.csv / *_lista.csv).

Uma aba de resumo (produto × quantidade; quantidade 1 fica em branco, como o sistema de etiquetas espera)
e uma aba com uma linha por etiqueta, por loja.
Código = código do produto no Microvix (é o código de barras). Preço = preço de venda do Microvix.

Uso: python3 scripts/gerar_etiquetas.py Produtos_Cadastrados.xlsx saida.xlsx "Praiamar=lista_loja1.csv" "Miramar=lista_loja3.csv"
"""
import csv
import os
import sys

import openpyxl
from openpyxl.styles import Font

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mangabe as mg  # noqa: E402


def gerar(cadastrados, saida, listas, precos=None):
    ex = mg.ler_cadastrados(cadastrados).drop_duplicates("cod").set_index("cod")
    cores, tams = mg.cores_tab(), set(mg.tamanhos())
    precos = precos or {}
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    cab = ["Código (barras)", "Descrição", "Referência", "Cor", "Tamanho", "Preço", "Qtd etiquetas"]
    totais = {}
    for nome, path in listas:
        rows = list(csv.DictReader(open(path, encoding="utf-8-sig"), delimiter=";"))
        res = wb.create_sheet(f"{nome} resumo"[:31])
        et = wb.create_sheet(f"{nome} uma por etiqueta"[:31])
        res.append(cab)
        et.append(cab[:-1])
        tot = 0
        for r in rows:
            c = r["Cod.Produto"]
            p = ex.loc[c]
            d = " ".join(p.desc.split())
            _, cor, t = mg.separar(d, cores, tams)
            preco = precos.get(c, mg.num(p.venda))
            q = int(r["Quantidade"])
            res.append([int(c), d, p.ref, cor, t, preco, q if q > 1 else None])  # 1 etiqueta = em branco
            for _ in range(q):
                et.append([int(c), d, p.ref, cor, t, preco])
            tot += q
        res.append([])
        res.append(["", "TOTAL", "", "", "", "", tot])
        for ws in (res, et):
            for cell in ws[1]:
                cell.font = Font(bold=True)
            for col, w in zip("ABCDEFG", [16, 40, 10, 12, 8, 10, 13]):
                ws.column_dimensions[col].width = w
            for row in ws.iter_rows(min_row=2):
                if len(row) > 5 and isinstance(row[5].value, (int, float)):
                    row[5].number_format = '"R$" #,##0.00'
        totais[nome] = tot
    wb.save(saida)
    print("etiquetas:", totais, "->", saida)
    return totais


if __name__ == "__main__":
    gerar(sys.argv[1], sys.argv[2], [a.split("=", 1) for a in sys.argv[3:]])
