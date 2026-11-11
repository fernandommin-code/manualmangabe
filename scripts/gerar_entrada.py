#!/usr/bin/env python3
"""Etapa 7 — gera os arquivos de entrada de estoque por loja, DEPOIS da importação no Microvix.

O código de barras é gerado pelo próprio Microvix, então a entrada usa o Cod.Produto.
Rode com o relatório Produtos Cadastrados exportado de novo, já com os produtos importados.

Uso:
  python3 scripts/gerar_entrada.py --conf conf/ --cadastrados Produtos_Cadastrados_pos_importacao.xlsx \
      --saida saida/ [--planilha 11_10_Cadastro_de_Produtos.xlsx]

Saídas em --saida:
  Entrada_Loja1.csv / Loja2 / Loja3   Cod.Produto;Descrição;Quantidade
  Entrada_Loja1.txt / ...             ARQUIVO PARA O MICROVIX: um código por linha, uma linha por peça
                                      (testado na GULE em 11/11; "cod;qtd" não é aceito)
  Entrada_nao_encontrados.csv         grade conferida que não achou produto no Microvix
"""
import argparse
import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mangabe as mg  # noqa: E402


def gerar(conf_dir, cadastrados, saida, planilha=None):
    ex = mg.ler_cadastrados(cadastrados)
    grade = mg.grade_conferida(mg.ler_conferencia(conf_dir))
    itens = mg.ler_planilha(planilha) if planilha else {}
    tab_marcas, tab_col = mg.marcas(), mg.colecoes()
    lojas = {0: {}, 1: {}, 2: {}}
    nao = []
    for (m, ref), g in sorted(grade.items()):
        aux = None
        tm = tab_marcas.get(m)
        it = itens.get((m, ref))
        if tm:
            digs = [tab_col[mg.norm(it["col"])]] if it and mg.norm(it["col"]) in tab_col else sorted(
                set(tab_col.values()), reverse=True)
            for d in digs:
                aux = mg.aux_existente(ex, ref, d, tm["id"])
                if aux:
                    break
        for (cor, tam), q in sorted(g.items(), key=lambda x: (x[0][0], mg.ordem_tam(x[0][1]))):
            cod = mg.produto_existente(ex, ref, cor, tam, aux) if aux else None
            if not cod:
                nao.append([m, ref, cor, tam, q.get(0, 0), q.get(1, 0), q.get(2, 0)])
                continue
            desc = ex.loc[ex.cod == cod, "desc"].iloc[0]
            for loja, n in q.items():
                lojas[loja][cod] = (desc, lojas[loja].get(cod, ("", 0))[1] + n)

    os.makedirs(saida, exist_ok=True)
    resumo = {}
    for loja, prods in lojas.items():
        if not prods:
            continue
        base = os.path.join(saida, f"Entrada_Loja{loja + 1}")
        with open(base + ".csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Cod.Produto", "Descrição", "Quantidade"])
            for cod, (desc, n) in sorted(prods.items()):
                w.writerow([cod, desc, n])
        with open(base + ".txt", "w", encoding="utf-8") as f:
            for cod, (_, n) in sorted(prods.items()):
                f.write((cod + "\n") * n)
        resumo[loja + 1] = sum(n for _, n in prods.values())
    if nao:
        with open(os.path.join(saida, "Entrada_nao_encontrados.csv"), "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Marca", "Referência", "Cor", "Tamanho", "Lj 1", "Lj 2", "Lj 3"])
            w.writerows(nao)
    print("Peças por loja:", resumo)
    print("Não encontrados no Microvix:", len(nao))
    for n in nao:
        print("  -", " ".join(str(x) for x in n[:4]))
    return dict(pecas=resumo, nao_encontrados=nao)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--conf", required=True)
    ap.add_argument("--cadastrados", required=True)
    ap.add_argument("--saida", required=True)
    ap.add_argument("--planilha")
    a = ap.parse_args()
    gerar(a.conf, a.cadastrados, a.saida, a.planilha)
