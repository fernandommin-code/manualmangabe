#!/usr/bin/env python3
"""Etapa 5 — gera os arquivos de importação de produtos do Microvix (um produto por cor × tamanho).

Uso:
  python3 scripts/gerar_cadastro.py \
      --planilha 11_10_Cadastro_de_Produtos.xlsx \
      --conf  conf/            (JSON exportados de compras/<id>/conf) \
      --fechamento fech/       (opcional: JSON de compras/<id>/fechamento) \
      --modelo ModeloImportacaoProduto.xlsx \
      --cadastrados Produtos_Cadastrados.xlsx \
      --saida saida/

Saídas em --saida:
  Importacao_<MARCA>.xlsx   um arquivo por marca, no layout do modelo do Microvix
  Resumo_cadastro.xlsx      códigos auxiliares criados, divergências, já cadastrados e pendências
"""
import argparse
import os
import sys

import openpyxl
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mangabe as mg  # noqa: E402


def atribuir_codigo(it, ex, ultimos, tab_marcas, tab_col, pend):
    """Cód. Auxiliar da referência: o já usado (reposição) ou o próximo livre da faixa.
    Retorna (aux, origem, marca_info) ou None (pendência registrada)."""
    tm = tab_marcas.get(mg.norm(it["marca"]))
    if not tm:
        pend.append(["Marca sem ID em tabelas/marcas.csv", it["marca"], it["ref"], ""])
        return None
    dig = tab_col.get(mg.norm(it["col"]))
    if not dig:
        pend.append(["Coleção sem dígito em tabelas/colecoes.csv", it["marca"], it["ref"], it["col"]])
        return None
    aux = mg.aux_existente(ex, it["ref"], dig, tm["id"])
    if aux:
        nomes = ex.desc_n[ex.aux == aux]
        if not any(d.startswith(mg.norm(it["desc"]) + " ") or d == mg.norm(it["desc"]) for d in nomes):
            pend.append(["Reposição com nome diferente: o cadastro usa o nome do Microvix",
                         it["marca"], it["ref"], f"planilha: {it['desc']} | Microvix: {mg.nome_base_microvix(ex, aux)}"])
        return aux, "reposição (código existente)", tm
    parte = mg.parte_da_peca(it["desc"])
    if parte is None:
        parte = "baixo"
        pend.append(["Tipo de peça desconhecido: assumido parte de baixo/peça única — conferir",
                     it["marca"], it["ref"], it["desc"]])
    u = ultimos.setdefault((dig, tm["id"]), {"baixo": -1, "cima": 499})
    u[parte] += 1
    if (parte == "baixo" and u[parte] > 499) or u[parte] > 999:
        raise SystemExit(f"Faixa esgotada: {it['marca']} {parte}")
    return f"{dig}{tm['id']}{u[parte]:03d}", "novo", tm


def gerar(planilha, conf_dir, modelo, cadastrados, saida, fech_dir=None):
    itens = mg.ler_planilha(planilha)
    ex = mg.ler_cadastrados(cadastrados)
    confs = mg.ler_conferencia(conf_dir)
    grade = mg.grade_conferida(confs)
    tab_marcas, tab_col = mg.marcas(), mg.colecoes()
    ultimos = mg.ultimos_codigos(ex)

    wb_modelo = openpyxl.load_workbook(modelo)
    hdr = [(c.value or "").strip() for c in wb_modelo.active[1]]
    col = {h: i for i, h in enumerate(hdr)}

    por_marca, codigos, ja_cad, pend = {}, [], [], []
    padroes = mg.padroes_microvix()
    faltando = [c for c in mg.OBRIGATORIAS if c not in ("Fornecedor", "Linha") and not padroes.get(c)]
    if faltando:
        pend.append(["Coluna obrigatória do Microvix sem valor em tabelas/padroes_microvix.csv", "", "",
                     ", ".join(faltando)])

    # pendências de fechamento / conferência
    if fech_dir:
        fech = mg.ler_fechamento(fech_dir)
        for (m, r), it in itens.items():
            for loja, q in enumerate(it["lj"]):
                if q and not fech.get((loja, m)):
                    pend.append(["Marca não concluída", it["marca"], "", f"Loja {loja + 1}"])
    for k, it in itens.items():
        if k not in grade and sum(it["lj"]):
            pend.append(["Referência sem conferência (não cadastrada)", it["marca"], it["ref"], it["desc"]])

    ordem = {k: i for i, k in enumerate(itens)}  # mesma ordem da planilha (= ordem do romaneio)
    for k in sorted(grade, key=lambda k: (ordem.get(k, len(ordem)), k)):
        g = grade[k]
        it = itens.get(k)
        if not it:
            pend.append(["Conferida mas fora da planilha (sem custo/preço)", k[0], k[1], "incluir na planilha e rodar de novo"])
            continue
        r = atribuir_codigo(it, ex, ultimos, tab_marcas, tab_col, pend)
        if not r:
            continue
        aux, origem, tm = r
        desc_base, ref_out = it["desc"], it["ref"]
        if origem.startswith("reposição"):  # mantém nome e referência como já estão no Microvix
            desc_base = mg.nome_base_microvix(ex, aux) or desc_base
            ref_out = ex.ref[ex.aux == aux].iloc[0]
        codigos.append([it["marca"], ref_out, desc_base, aux, origem])
        if not mg.linha_da_peca(desc_base):
            pend.append(["Linha do Microvix desconhecida (incluir a 1ª palavra em tabelas/tipos_peca.csv)",
                         it["marca"], it["ref"], desc_base])
        if not tm["fornecedor"] and not any(p[0].startswith("Marca sem fornecedor") and p[1] == it["marca"] for p in pend):
            pend.append(["Marca sem fornecedor em tabelas/marcas.csv (coluna obrigatória no Microvix)",
                         it["marca"], "", ""])

        for cor, tam in sorted(g, key=lambda x: (x[0], mg.ordem_tam(x[1]))):
            desc = mg.descricao_produto(desc_base, cor, tam)
            q = g[(cor, tam)]
            cod = mg.produto_existente(ex, ref_out, cor, tam, aux)
            if cod:
                ja_cad.append([cod, desc, it["ref"], aux, q.get(0, 0), q.get(1, 0), q.get(2, 0)])
                continue
            row = [None] * len(hdr)

            def put(nome, v):
                row[col[nome]] = v

            put("Descrição", " ".join(str(desc_base).upper().split()))  # o Microvix acrescenta cor e tamanho
            put("Referência", ref_out)
            put("Cód. Auxiliar", aux)
            put("Marca", it["marca"])
            put("Coleção", it["col"])
            put("Tamanho", tam)
            put("Cores", cor or None)
            put("Contabiliza saldo em estoque", "Sim")
            put("Custo com ICMS (R$)", it["custo"])
            put("Desconto (%)", tm["desconto"] or None)
            put("Preço de venda R$", it["venda"])
            if it["ncm"]:
                put("NCM", str(it["ncm"]))
            for nome, v in padroes.items():
                if v and nome in col:
                    put(nome, v)
            if tm["fornecedor"]:
                put("Fornecedor", tm["fornecedor"])
            linha = mg.linha_da_peca(desc_base)
            if linha:
                put("Linha", linha)
            por_marca.setdefault(it["marca"], []).append(row)

    os.makedirs(saida, exist_ok=True)
    arquivos = []
    for m, rows in por_marca.items():
        wb = openpyxl.load_workbook(modelo)
        ws = wb.active
        for r in rows:
            ws.append(r)
        nome = os.path.join(saida, f"Importacao_{mg.slug(m).replace('-', '_')}.xlsx")
        wb.save(nome)
        arquivos.append(nome)

    resumo = os.path.join(saida, "Resumo_cadastro.xlsx")
    with pd.ExcelWriter(resumo) as w:
        pd.DataFrame(codigos, columns=["Marca", "Referência", "Descrição", "Cód. Auxiliar", "Origem"]).to_excel(
            w, sheet_name="Códigos auxiliares", index=False)
        pd.DataFrame(mg.divergencias(confs), columns=["loja", "marca", "ref", "desc", "esperado", "contado",
                                                      "diferenca", "obs"]).to_excel(
            w, sheet_name="Divergências", index=False)
        pd.DataFrame(ja_cad, columns=["Cod.Produto", "Descrição", "Referência", "Cód. Auxiliar", "Lj 1", "Lj 2",
                                      "Lj 3"]).to_excel(w, sheet_name="Já cadastrados", index=False)
        pd.DataFrame(pend, columns=["Pendência", "Marca", "Referência", "Detalhe"]).to_excel(
            w, sheet_name="Pendências", index=False)

    for ref, cor, tam, cods, usado in mg.REPETIDOS:
        print(f"  ! Produto repetido no Microvix: ref {ref} {cor} {tam} -> {', '.join(cods)} (usado {usado})")
    novos = {m: len(r) for m, r in por_marca.items()}
    print("Produtos novos (cor × tamanho) por marca:", novos)
    print("Já cadastrados (só entrada):", len(ja_cad))
    print("Pendências:", len(pend))
    for p in pend:
        print("  -", " | ".join(str(x) for x in p if x))
    return dict(arquivos=arquivos, resumo=resumo, novos=novos, codigos=codigos, ja_cadastrados=ja_cad, pendencias=pend)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--planilha", required=True)
    ap.add_argument("--conf", required=True)
    ap.add_argument("--fechamento")
    ap.add_argument("--modelo", required=True)
    ap.add_argument("--cadastrados", required=True)
    ap.add_argument("--saida", required=True)
    a = ap.parse_args()
    gerar(a.planilha, a.conf, a.modelo, a.cadastrados, a.saida, a.fechamento)
