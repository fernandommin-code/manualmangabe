#!/usr/bin/env python3
"""Teste com o histórico real, sem tocar no Microvix.

Pega as últimas N referências cadastradas de cada marca (parte de baixo e de cima),
finge que são compra nova (tira do relatório, monta a planilha e a conferência)
e roda o gerador. Compara o que o script gera com o que já está no Microvix:
  - Código Auxiliar na faixa certa (baixo/cima) e na sequência certa
  - Descrição de cada cor × tamanho igual à cadastrada
Também revisa o cadastro inteiro da coleção: cores fora da tabela, tamanhos estranhos,
referência com dois códigos, código com dois nomes.

Uso: python3 scripts/simular_com_historico.py Produtos_Cadastrados.xlsx ModeloImportacaoProduto.xlsx saida_teste/ [--n 3]
"""
import argparse
import json
import os
import sys
from collections import Counter, defaultdict

import openpyxl
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gerar_cadastro  # noqa: E402
import mangabe as mg  # noqa: E402

TIT = ["Lj 1", "Lj 2", "Lj 3", "Cód. Interno", "Descrição do Produto", "Coleção", "Referência",
       "Preço Custo", "Preço Venda", "NCM"]


def num(x):
    x = str(x or "0").strip()
    if "," in x:
        x = x.replace(".", "").replace(",", ".")
    return float(x)


def cores_tab():
    with open(os.path.join(mg.TABELAS, "cores.csv"), encoding="utf-8") as f:
        return sorted((mg.norm(l) for l in f.read().split("\n")[1:] if l.strip()), key=len, reverse=True)


def separar(desc, cores, tams):
    """'VESTIDO MIDI LINHO OFF WHITE PP' -> ('VESTIDO MIDI LINHO', 'OFF WHITE', 'PP')"""
    w = desc.split()
    tam = w[-1] if w and mg.norm_tam(w[-1]) in tams else None
    corpo = " ".join(w[:-1] if tam else w)
    for c in cores:
        if mg.norm(corpo).endswith(" " + c):
            return corpo[: len(corpo) - len(c)].strip(), c, mg.norm_tam(tam) if tam else "ÚNICO"
    return corpo, None, mg.norm_tam(tam) if tam else "ÚNICO"


def salvar_cadastrados(ex, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Produtos Cadastrados - Microvix ERP"])
    ws.append(["Cod.Produto", "Descrição", "Referência", "Cód. Auxiliar", "Custo Produto(R$)", "Mark-up(%)",
               "Preço Venda(R$)"])
    for r in ex[["cod", "desc", "ref", "aux", "custo", "mk", "venda"]].itertuples(index=False):
        ws.append(list(r))
    wb.save(path)


def revisar_catalogo(v, cores, tams):
    achados = []
    cor_fora = Counter()
    for d in v.desc:
        _, c, _ = separar(d, cores, tams)
        if c is None:
            cor_fora[" ".join(d.split()[-2:-1])] += 1
    for k, n in cor_fora.most_common(15):
        achados.append(["Cor não reconhecida (penúltima palavra)", k, n])
    for t, n in Counter(d.split()[-1] for d in v.desc if mg.norm_tam(d.split()[-1]) not in tams).most_common(10):
        achados.append(["Última palavra não é tamanho", t, n])
    g = v.groupby(["ref_n", v.aux.str[1:3]]).aux.nunique()
    for (ref, m), n in g[g > 1].items():
        achados.append(["Referência com mais de um Cód. Auxiliar na coleção", f"marca {m} ref {ref}",
                        ", ".join(sorted(v[(v.ref_n == ref) & (v.aux.str[1:3] == m)].aux.unique()))])
    for aux, gr in v.groupby("aux"):
        nomes = {separar(d, cores, tams)[0] for d in gr.desc}
        if len(nomes) > 1:
            achados.append(["Cód. Auxiliar com nomes diferentes", aux, " / ".join(sorted(nomes))[:120]])
        if gr.ref_n.nunique() > 1:
            achados.append(["Cód. Auxiliar em mais de uma referência", aux, ", ".join(sorted(gr.ref.unique()))])
    return achados


def sem_aux(ex, v):
    refs = set(v.ref_n)
    m = ex[(ex.aux == "") & ex.ref_n.isin(refs)]
    return [["Produto sem Cód. Auxiliar (ref. da coleção)", r.cod, r.desc] for r in m.itertuples()]


def main(cadastrados, modelo, saida, n, colecao="VERÃO 27"):
    os.makedirs(saida, exist_ok=True)
    ex = mg.ler_cadastrados(cadastrados)
    dig = mg.colecoes()[mg.norm(colecao)]
    tams, cores = set(mg.tamanhos()), cores_tab()
    v = ex[ex.aux.str.fullmatch(dig + r"\d{5}")]
    id2marca = {m["id"]: nome for nome, m in mg.marcas().items()}

    # 1. escolhe as últimas N de cada faixa por marca
    escolhidos = []
    for idm in sorted(v.aux.str[1:3].unique()):
        auxs = sorted(v.aux[v.aux.str[1:3] == idm].unique())
        escolhidos += [a for a in auxs if int(a[3:]) < 500][-n:] + [a for a in auxs if int(a[3:]) >= 500][-n:]
    sel = v[v.aux.isin(escolhidos)]
    resto = ex[~ex.aux.isin(escolhidos)]

    # 2. monta planilha e conferência
    wb = openpyxl.Workbook()
    ws = wb.active
    r = 3
    conf_dir = os.path.join(saida, "conf")
    os.makedirs(conf_dir, exist_ok=True)
    real, sem_marca, nao_separou = {}, set(), []
    for idm, grp in sel.groupby(sel.aux.str[1:3]):
        marca = id2marca.get(idm)
        if not marca:
            sem_marca.add(idm)
            continue
        ws.cell(r, 1, marca)
        for i, t in enumerate(TIT):
            ws.cell(r + 1, i + 1, t)
        r += 2
        for (aux, ref), g in grp.groupby(["aux", "ref"]):  # ordem em que foram cadastradas
            partes = [separar(d, cores, tams) for d in g.desc]
            nome = Counter(p[0] for p in partes).most_common(1)[0][0]
            linhas = defaultdict(dict)
            for (nm, c, t), d in zip(partes, g.desc):
                if c is None and t != "ÚNICO" or nm != nome:
                    nao_separou.append([marca, ref, d])
                linhas[c or ""][t] = 1
            for i, val in enumerate([1, 0, 0, None, nome, colecao, ref, num(g.custo.iloc[0]),
                                     num(g.venda.iloc[0]), None]):
                ws.cell(r, i + 1, val)
            r += 1
            real[(mg.norm(marca), mg.norm(ref))] = dict(aux=aux, descs=set(mg.norm(d) for d in g.desc))
            with open(os.path.join(conf_dir, f"0~{mg.slug(marca)}~{mg.slug(ref)}.json"), "w") as f:
                json.dump(dict(loja=0, marca=marca, ref=ref, esperado=len(g),
                               linhas=[{"cor": c, "q": q} for c, q in linhas.items()]), f)
        r += 1
    plan = os.path.join(saida, "planilha_simulada.xlsx")
    wb.save(plan)
    cad = os.path.join(saida, "cadastrados_sem_as_ultimas.xlsx")
    salvar_cadastrados(resto, cad)

    # 3. roda o gerador
    res = gerar_cadastro.gerar(plan, conf_dir, modelo, cad, os.path.join(saida, "importacao"))
    gerados = defaultdict(set)
    for f in res["arquivos"]:
        w2 = openpyxl.load_workbook(f).active
        h = [(c.value or "").strip() for c in w2[1]]
        for row in w2.iter_rows(min_row=2, values_only=True):
            d = dict(zip(h, row))
            gerados[(mg.norm(d["Marca"]), mg.norm(d["Referência"]))].add(mg.norm(d["Descrição"]))
    aux_ger = {(mg.norm(m), mg.norm(ref)): a for m, ref, _, a, _ in res["codigos"]}

    # 4. compara
    rel = []
    for k, rv in sorted(real.items()):
        a = aux_ger.get(k)
        faixa_ok = a and (int(a[3:]) >= 500) == (int(rv["aux"][3:]) >= 500)
        falta = rv["descs"] - gerados.get(k, set())
        sobra = gerados.get(k, set()) - rv["descs"]
        rel.append(dict(marca=k[0], ref=k[1], aux_real=rv["aux"], aux_script=a,
                        codigo_igual=a == rv["aux"], faixa_ok=bool(faixa_ok),
                        descricoes_iguais=not falta and not sobra,
                        so_no_microvix=" | ".join(sorted(falta))[:200], so_no_script=" | ".join(sorted(sobra))[:200]))
    df = pd.DataFrame(rel)
    achados = revisar_catalogo(v, cores, tams) + sem_aux(ex, v)
    out = os.path.join(saida, "Relatorio_teste.xlsx")
    with pd.ExcelWriter(out) as w:
        df.to_excel(w, sheet_name="Comparação", index=False)
        pd.DataFrame(achados, columns=["Achado", "Item", "Detalhe"]).to_excel(w, sheet_name="Revisão do cadastro",
                                                                             index=False)
        pd.DataFrame(nao_separou, columns=["Marca", "Ref", "Descrição no Microvix"]).to_excel(
            w, sheet_name="Nome fora do padrão", index=False)
    print(f"Referências testadas: {len(df)}  marcas: {df.marca.nunique()}")
    print(f"  Cód. Auxiliar idêntico: {df.codigo_igual.sum()}   faixa baixo/cima certa: {df.faixa_ok.sum()}")
    print(f"  Descrições idênticas:   {df.descricoes_iguais.sum()}")
    print(f"  Marcas sem ID na tabela: {sorted(sem_marca)}")
    print(f"  Nomes fora do padrão <DESC> <COR> <TAM>: {len(nao_separou)}")
    print(f"  Achados na revisão do cadastro: {len(achados)}")
    print("Relatório:", out)
    return df, achados, nao_separou


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("cadastrados")
    ap.add_argument("modelo")
    ap.add_argument("saida")
    ap.add_argument("--n", type=int, default=3)
    a = ap.parse_args()
    main(a.cadastrados, a.modelo, a.saida, a.n)
