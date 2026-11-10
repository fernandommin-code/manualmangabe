#!/usr/bin/env python3
"""Etapa 1½ — prévia da planilha do dia, ANTES da conferência.

Mostra, para cada referência, o Cód. Auxiliar que ela vai receber (novo ou reposição),
se é parte de cima ou de baixo, e o que precisa de ajuste antes do cadastro:
marca sem ID, tipo de peça desconhecido, descrição genérica ou cortada, NCM faltando,
referência repetida, totais que não fecham.

Uso: python3 scripts/previa_codigos.py 11_10_Cadastro_de_Produtos.xlsx Produtos_Cadastrados.xlsx [saida.xlsx]
"""
import os
import sys
from collections import Counter

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gerar_cadastro  # noqa: E402
import mangabe as mg  # noqa: E402


def previa(planilha, cadastrados, saida=None):
    itens = mg.ler_planilha(planilha)
    ex = mg.ler_cadastrados(cadastrados)
    tab_marcas, tab_col = mg.marcas(), mg.colecoes()
    ultimos = mg.ultimos_codigos(ex)
    pend, linhas = [], []
    nomes = Counter((mg.norm(it["marca"]), mg.norm(it["desc"])) for it in itens.values())
    for k, it in itens.items():
        r = gerar_cadastro.atribuir_codigo(it, ex, ultimos, tab_marcas, tab_col, pend)
        aux, origem = (r[0], r[1]) if r else ("", "")
        parte = "cima" if aux and int(aux[3:]) >= 500 else "baixo/única" if aux else ""
        linhas.append([it["marca"], it["ref"], it["desc"], *it["lj"], sum(it["lj"]), aux, parte, origem])
        palavras = mg.norm(it["desc"]).split()
        if len(palavras) <= 1 and mg.norm(it["marca"]) != "VICKY":
            pend.append(["Descrição genérica (no caixa fica igual a outros modelos)", it["marca"], it["ref"], it["desc"]])
        if nomes[(mg.norm(it["marca"]), mg.norm(it["desc"]))] > 1 and mg.norm(it["marca"]) != "VICKY":
            pend.append(["Mesmo nome em mais de uma referência da marca", it["marca"], it["ref"], it["desc"]])
        if not it["ncm"]:
            pend.append(["Sem NCM (aguardando nota)", it["marca"], it["ref"], ""])
        if not sum(it["lj"]):
            pend.append(["Referência sem quantidade", it["marca"], it["ref"], it["desc"]])
        if aux and origem.startswith("reposição"):
            ref_mv = ex.ref[ex.aux == aux].iloc[0]
            if ref_mv != it["ref"]:
                pend.append(["Referência escrita diferente: o cadastro usa a do Microvix", it["marca"], it["ref"],
                             f"Microvix: {ref_mv}"])
    df = pd.DataFrame(linhas, columns=["Marca", "Referência", "Descrição", "Lj 1", "Lj 2", "Lj 3", "Total",
                                       "Cód. Auxiliar", "Parte", "Origem"])
    pdf = pd.DataFrame(pend, columns=["Pendência", "Marca", "Referência", "Detalhe"])
    if saida:
        with pd.ExcelWriter(saida) as w:
            df.to_excel(w, sheet_name="Códigos", index=False)
            pdf.to_excel(w, sheet_name="Pendências", index=False)
    return df, pdf


if __name__ == "__main__":
    df, pdf = previa(*sys.argv[1:4])
    pd.set_option("display.width", 200)
    pd.set_option("display.max_rows", 500)
    print(df.to_string(index=False))
    print()
    print(pdf.to_string(index=False))
