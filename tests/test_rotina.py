"""Teste de ponta a ponta: planilha do dia + conferência do app -> importação -> entrada.

Rodar: python3 -m pytest tests/
"""
import csv
import json
import os
import sys

import openpyxl
import pandas as pd

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(AQUI), "scripts"))
import gerar_cadastro  # noqa: E402
import gerar_entrada  # noqa: E402
import mangabe as mg  # noqa: E402

MODELO = os.path.join(AQUI, "fixtures", "ModeloImportacaoProduto.xlsx")
TIT = ["Lj 1", "Lj 2", "Lj 3", "Cód. Interno", "Descrição do Produto", "Coleção", "Referência",
       "Preço Custo", "Preço Venda", "NCM"]


def planilha(path, blocos):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["E1"] = "CADASTRO DE PRODUTOS"
    r = 3
    for marca, linhas in blocos:
        ws.cell(r, 1, marca)
        for i, t in enumerate(TIT):
            ws.cell(r + 1, i + 1, t)
        r += 2
        for l in linhas:
            for i, v in enumerate(l):
                ws.cell(r, i + 1, v)
            r += 1
        ws.cell(r, 4, "TOTAL PEÇAS")
        r += 2
    wb.save(path)


def cadastrados(path, linhas):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Produtos Cadastrados - Microvix ERP"])
    ws.append(["Cod.Produto", "Descrição", "Referência", "Cód. Auxiliar", "Custo Produto(R$)", "Mark-up(%)",
               "Preço Venda(R$)"])
    for l in linhas:
        ws.append(l)
    wb.save(path)


def conf(d, loja, marca, ref, esperado, linhas, obs=""):
    os.makedirs(d, exist_ok=True)
    total = sum(q for l in linhas for q in l["q"].values())
    doc = dict(loja=loja, marca=marca, ref=ref, desc="", esperado=esperado, extra=False, tamKey="L",
               linhas=linhas, obs=obs, total=total)
    with open(os.path.join(d, f"{loja}~{mg.slug(marca)}~{mg.slug(ref)}.json"), "w") as f:
        json.dump({"id": "x", "data": doc}, f)


EXISTENTES = [
    ["20001", "SAIA MIDI FENDAS OFF WHITE PP", "9449", "639032", "119", "100", "299"],
    ["20002", "CROPPED ALEGRIA VERDE PP", "8559", "639527", "99", "100", "229"],
    ["20003", "BLUSA BRUNA LASTEX OFF WHITE PP", "9469", "639529", "89", "100", "199"],
    ["20004", "BLUSA BRUNA LASTEX OFF WHITE P", "9469", "639529", "89", "100", "199"],
    ["20005", "BLUSA ANTIGA PRETO P", "9469", "539501", "89", "100", "199"],  # mesma ref, outra coleção
]


def test_rotina_completa(tmp_path):
    p = tmp_path / "plan.xlsx"
    planilha(p, [("FPM", [
        [3, 1, 0, None, "VESTIDO MIDI LINHO", "VERÃO 27", "9500", 159.9, 399, "62044200"],
        [1, 1, 0, None, "BLUSA BRUNA LASTEX LISA", "VERÃO 27", "09469", 89, 199, None],  # zero à esquerda e nome diferente
        [2, 0, 0, None, "REGATA CANELADA", "VERÃO 27", "9501", 59.9, 149, None],
        [1, 0, 0, None, "PONCHO TRICO", "VERÃO 27", "9502", 99, 249, None],
    ])])
    cad = tmp_path / "cad.xlsx"
    cadastrados(cad, EXISTENTES)
    c = tmp_path / "conf"
    conf(c, 0, "FPM", "9500", 3, [{"cor": "PRETO", "q": {"P": 1, "M": 1}}, {"cor": "bege", "q": {"P": 1}}])
    conf(c, 1, "FPM", "9500", 1, [{"cor": "PRETO", "q": {"M": 1}}])
    conf(c, 0, "FPM", "09469", 1, [{"cor": "OFF WHITE", "q": {"P": 1}}])            # já existe
    conf(c, 1, "FPM", "09469", 1, [{"cor": "LILÁS", "q": {"XPP": 1}}])             # cor nova, mesmo aux
    conf(c, 0, "FPM", "9501", 2, [{"cor": "BRANCO", "q": {"PP": 1}}], obs="faltou 1")
    conf(c, 0, "FPM", "9502", 1, [{"cor": "CRU", "q": {"U": 1}}])

    out = tmp_path / "saida"
    r = gerar_cadastro.gerar(str(p), str(c), MODELO, str(cad), str(out))

    aux = {ref: a for _, ref, _, a, _ in r["codigos"]}
    assert aux["9469"] == "639529"      # reposição: reaproveita o código da coleção atual, não o 539501
    assert aux["9500"] == "639033"      # vestido: parte de baixo, maior usado 032
    assert aux["9501"] == "639530"      # regata: parte de cima, maior usado 529
    assert aux["9502"] == "639531"      # poncho: parte de cima
    ws = openpyxl.load_workbook(out / "Importacao_FPM.xlsx").active
    hdr = [(x.value or "").strip() for x in ws[1]]
    rows = [dict(zip(hdr, [x.value for x in row])) for row in ws.iter_rows(min_row=2)]
    assert rows[0]["Descrição"] == "VESTIDO MIDI LINHO"  # o Microvix acrescenta cor e tamanho
    descs = [mg.descricao_produto(x["Descrição"], x["Cores"], x["Tamanho"]) for x in rows]  # nome final no Microvix
    assert descs == ["VESTIDO MIDI LINHO BEGE P", "VESTIDO MIDI LINHO PRETO P", "VESTIDO MIDI LINHO PRETO M",
                     "BLUSA BRUNA LASTEX LILAS XPP", "REGATA CANELADA BRANCO PP", "PONCHO TRICO CRU ÚNICO"]
    v = rows[0]
    assert v["Cód. Auxiliar"] == "639033" and v["Referência"] == "9500" and v["Tamanho"] == "P"
    assert v["Cores"] == "BEGE" and v["Custo com ICMS (R$)"] == 159.9 and v["Desconto (%)"] == 15
    assert v["Preço de venda R$"] == 399 and v["NCM"] == "62044200" and v["Código de barras"] is None
    assert rows[3]["Cód. Auxiliar"] == "639529" and rows[3]["Referência"] == "9469"  # nome e ref do Microvix
    assert len(r["ja_cadastrados"]) == 1 and r["ja_cadastrados"][0][0] == "20004"
    div = pd.read_excel(out / "Resumo_cadastro.xlsx", sheet_name="Divergências")
    assert list(div.ref.astype(str)) == ["9501"]

    # simula o Microvix depois da importação
    novos = [[str(30000 + i), d, x["Referência"], x["Cód. Auxiliar"], "", "", ""] for i, (d, x) in
             enumerate(zip(descs, rows))]
    cad2 = tmp_path / "cad2.xlsx"
    cadastrados(cad2, EXISTENTES + novos)
    e = gerar_entrada.gerar(str(c), str(cad2), str(out), str(p))
    assert e["nao_encontrados"] == []
    assert e["pecas"] == {1: 6, 2: 2}
    with open(out / "Entrada_Loja1.csv", encoding="utf-8-sig") as f:
        l1 = {r["Cod.Produto"]: int(r["Quantidade"]) for r in csv.DictReader(f, delimiter=";")}
    assert l1["20004"] == 1 and "20005" not in l1
    assert open(out / "Entrada_Loja2.txt").read().split() == ["30002", "30003"]


def test_parte_da_peca():
    assert mg.parte_da_peca("Blusa Bruna") == "cima"
    assert mg.parte_da_peca("CALÇA PANTALONA") == "baixo"
    assert mg.parte_da_peca("MACACÃO LONGO") == "baixo"
    assert mg.parte_da_peca("PONCHO") == "cima"
    assert mg.parte_da_peca("KAFTAN") is None
