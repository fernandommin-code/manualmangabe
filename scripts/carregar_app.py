#!/usr/bin/env python3
"""Prepara os documentos do app de conferência (https://claude.ai/artifact/Bw297fhPHqWy35JJEU1hds).

Gera um JSON por documento, no formato que o Claude grava com ArtifactData (batch, file_path):
  compra.json                 -> compras/<id>            {titulo, data, aberta}
  marcas/<SLUG>.json          -> compras/<id>/marcas      {marca, ordem, itens:[{ref, desc, lj, tam, exist}]}
  conf/<loja>~<MARCA>~<REF>.json -> compras/<id>/conf     (só com --conf: conferências feitas em papel)

`exist` = cores que já existem no Microvix para a referência (reposição), para a vendedora
escrever a cor com o mesmo nome do cadastro.

Uso:
  python3 scripts/carregar_app.py PLANILHA.xlsx saida/ --data 2026-11-10 \
      [--cadastrados Produtos_Cadastrados.xlsx] [--conf dir_com_conferencias_transcritas]
"""
import argparse
import datetime
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mangabe as mg  # noqa: E402


def gerar(planilha, saida, data, cadastrados=None, conf_dir=None):
    itens = mg.ler_planilha(planilha)
    ex = mg.ler_cadastrados(cadastrados) if cadastrados else None
    tab_m, tab_c = mg.marcas(), mg.colecoes()
    cores, tams = mg.cores_tab(), set(mg.tamanhos())
    os.makedirs(os.path.join(saida, "marcas"), exist_ok=True)
    d = datetime.date.fromisoformat(data)
    json.dump({"titulo": f"Compra {d:%d/%m}", "data": data, "aberta": True},
              open(os.path.join(saida, "compra.json"), "w"), ensure_ascii=False)
    marcas = {}
    for it in itens.values():
        marcas.setdefault(it["marca"], []).append(it)
    for ordem, (marca, its) in enumerate(marcas.items(), start=1):
        unico = mg.norm(marca) == "VICKY"
        lista = []
        for it in its:
            item = {"ref": it["ref"], "desc": it["desc"], "lj": it["lj"], "tam": "U" if unico else "L"}
            tm, dig = tab_m.get(mg.norm(marca)), tab_c.get(mg.norm(it["col"]))
            if ex is not None and tm and dig:
                aux = mg.aux_existente(ex, it["ref"], dig, tm["id"])
                if aux:
                    item["exist"] = sorted({mg.separar(x, cores, tams)[1] or "?" for x in ex.desc[ex.aux == aux]})
            lista.append(item)
        json.dump({"marca": marca, "ordem": ordem, "itens": lista},
                  open(os.path.join(saida, "marcas", mg.slug(marca) + ".json"), "w"), ensure_ascii=False)
    n_conf = 0
    if conf_dir:
        os.makedirs(os.path.join(saida, "conf"), exist_ok=True)
        refs = {(mg.norm(it["marca"]), mg.ref_key(it["ref"])): it for it in itens.values()}
        for c in mg.ler_conferencia(conf_dir):
            it = refs.get((mg.norm(c["marca"]), mg.ref_key(c["ref"])))
            if not it:
                continue
            loja = int(c["loja"])
            linhas = [{"cor": l.get("cor") or "SEM COR", "q": {k: int(v) for k, v in l["q"].items() if int(v)}}
                      for l in c.get("linhas", []) if any(int(v) for v in l["q"].values())]
            total = sum(q for l in linhas for q in l["q"].values())
            doc = {"loja": loja, "marca": it["marca"], "ref": it["ref"], "desc": it["desc"],
                   "esperado": it["lj"][loja], "extra": False, "tamKey": "U" if mg.norm(it["marca"]) == "VICKY" else "L",
                   "extraSizes": [], "linhas": linhas, "total": total,
                   "obs": c.get("obs") or ("transcrito da folha de papel" if total == it["lj"][loja]
                                           else "transcrito da folha de papel - conferir"),
                   "em": datetime.datetime.now().isoformat(timespec="seconds")}
            nome = f"{loja}~{mg.slug(it['marca'])}~{mg.slug(it['ref'])}.json"
            json.dump(doc, open(os.path.join(saida, "conf", nome), "w"), ensure_ascii=False)
            n_conf += 1
    print(f"{len(marcas)} marcas, {len(itens)} referências, {n_conf} conferências -> {saida}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("planilha")
    ap.add_argument("saida")
    ap.add_argument("--data", required=True, help="AAAA-MM-DD da compra (id no app)")
    ap.add_argument("--cadastrados")
    ap.add_argument("--conf")
    a = ap.parse_args()
    gerar(a.planilha, a.saida, a.data, a.cadastrados, a.conf)
