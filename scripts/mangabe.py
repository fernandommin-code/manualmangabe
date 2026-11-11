"""Funções comuns da rotina de cadastro por grade (Mangabê / Microvix)."""
import csv
import glob
import json
import os
import re
import unicodedata
from collections import Counter

import openpyxl
import pandas as pd

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TABELAS = os.path.join(RAIZ, "tabelas")


def norm(s):
    """Maiúsculo, sem acento, espaços simples."""
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode()
    return " ".join(s.upper().split())


def ref_key(s):
    """Chave de comparação da referência: sem acento, maiúscula e sem zeros à esquerda (09404 = 9404)."""
    return norm(s).lstrip("0") or "0"


def slug(s):
    return re.sub(r"[^A-Z0-9]+", "-", norm(s)).strip("-") or "X"


def _csv(nome):
    with open(os.path.join(TABELAS, nome), encoding="utf-8") as f:
        return list(csv.DictReader(f, delimiter=";"))


def marcas():
    """{MARCA normalizada: {'id': '39', 'desconto': 15.0}}"""
    out = {}
    for r in _csv("marcas.csv"):
        d = r["desconto_compra"].strip()
        out[norm(r["marca"])] = {"id": r["id"].zfill(2), "desconto": float(d) if d else 0.0,
                                 "fornecedor": (r.get("fornecedor_microvix") or "").strip()}
    return out


def padroes_microvix():
    """Valores fixos que o Microvix exige em todo produto (tabelas/padroes_microvix.csv)."""
    return {r["coluna"].strip(): r["valor"].strip() for r in _csv("padroes_microvix.csv")}


OBRIGATORIAS = ["Fornecedor", "Configuração tributária", "Unidade de venda", "Setor", "Linha", "Espessura"]


def colecoes():
    """{COLEÇÃO normalizada: dígito}"""
    return {norm(r["colecao"]): r["digito"] for r in _csv("colecoes.csv")}


def tamanhos():
    return [r["tamanho"] for r in sorted(_csv("tamanhos.csv"), key=lambda r: int(r["ordem"]))]


def tipos_peca():
    return {norm(r["palavra"]): r["parte"] for r in _csv("tipos_peca.csv")}


def parte_da_peca(desc):
    """'cima' (500–999) ou 'baixo' (000–499, inclui peça única). None se a 1ª palavra não estiver na tabela."""
    tipos = tipos_peca()
    for w in norm(desc).split():  # 1ª palavra conhecida: "MAXI COLAR" -> COLAR, "MAXI CROPPED" -> CROPPED
        if w in tipos:
            return tipos[w]
    return None


TAM_ALIAS = {"U": "ÚNICO", "UNICO": "ÚNICO", "UN": "ÚNICO", "UNICA": "ÚNICO", "ÚNICA": "ÚNICO"}


def norm_tam(t):
    t = " ".join(str(t).upper().split())
    return TAM_ALIAS.get(t, TAM_ALIAS.get(norm(t), t))


def norm_cor(c):
    c = norm(c)
    return "" if c in ("", "SEM COR", "UNICA", "UNICO") else c


def cores_tab():
    with open(os.path.join(TABELAS, "cores.csv"), encoding="utf-8") as f:
        return sorted((norm(l) for l in f.read().split("\n")[1:] if l.strip()), key=len, reverse=True)


def separar(desc, cores, tams):
    """'VESTIDO MIDI LINHO OFF WHITE PP' -> ('VESTIDO MIDI LINHO', 'OFF WHITE', 'PP')"""
    w = desc.split()
    tam = w[-1] if w and norm_tam(w[-1]) in tams else None
    corpo = " ".join(w[:-1] if tam else w)
    for c in cores:
        if norm(corpo).endswith(" " + c):
            return corpo[: len(corpo) - len(c)].strip(), c, norm_tam(tam) if tam else "ÚNICO"
    return corpo, None, norm_tam(tam) if tam else "ÚNICO"


def ordem_tam(t):
    lista = tamanhos()
    return lista.index(t) if t in lista else len(lista)


def descricao_produto(desc, cor, tam):
    """Descrição no padrão Microvix: <DESCRIÇÃO> <COR> <TAMANHO> (ex.: BLUSA LINKS PRETO ÚNICO).
    Acessório sem cor e tamanho único fica só com a descrição (ex.: COLAR CAMADA METAL)."""
    base = " ".join(str(desc).upper().split())
    if not cor and tam == "ÚNICO":
        return base
    return " ".join(p for p in (base, cor, tam) if p)


# ---------------------------------------------------------------- planilha do dia
def ler_planilha(path):
    """Lê a planilha do dia (um bloco por marca). Retorna {(MARCA, REF): item}."""
    ws = openpyxl.load_workbook(path, data_only=True).active
    itens, marca = {}, None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if isinstance(a, str) and ws.cell(r + 1, 1).value == "Lj 1":
            marca = a.strip()
            continue
        ref = ws.cell(r, 7).value
        if marca and ref not in (None, "", "Referência"):
            ref = str(ref).strip()
            if isinstance(ws.cell(r, 7).value, float) and ws.cell(r, 7).value.is_integer():
                ref = str(int(ws.cell(r, 7).value))
            itens[(norm(marca), ref_key(ref))] = dict(
                marca=marca, ref=ref, desc=ws.cell(r, 5).value or "",
                col=ws.cell(r, 6).value or "VERÃO 27", custo=ws.cell(r, 8).value,
                venda=ws.cell(r, 9).value, ncm=ws.cell(r, 10).value,
                lj=[int(ws.cell(r, c).value or 0) for c in (1, 2, 3)])
    return itens


# ---------------------------------------------------------------- relatório Produtos Cadastrados
def ler_cadastrados(path):
    """Relatório Produtos Cadastrados do Microvix (não abre no openpyxl: usa calamine)."""
    ex = pd.read_excel(path, header=1, engine="calamine", dtype=str).dropna(subset=["Cod.Produto"])
    ex = ex.iloc[:, :7]
    ex.columns = ["cod", "desc", "ref", "aux", "custo", "mk", "venda"]
    ex = ex.fillna("")
    for c in ("cod", "desc", "ref", "aux"):
        ex[c] = ex[c].str.strip()
    ex["desc_n"] = ex["desc"].map(norm)
    ex["ref_n"] = ex["ref"].map(ref_key)
    return ex


def ultimos_codigos(ex):
    """{(coleção, marca): {'baixo': maior nº usado, 'cima': maior nº usado}} só com Cód. Auxiliar de 6 dígitos."""
    out = {}
    for a in ex.aux[ex.aux.str.fullmatch(r"\d{6}")].unique():
        k = (a[0], a[1:3])
        n = int(a[3:])
        p = out.setdefault(k, {"baixo": -1, "cima": 499})
        lado = "cima" if n >= 500 else "baixo"
        p[lado] = max(p[lado], n)
    return out


def aux_existente(ex, ref, dig_col, id_marca):
    """Cód. Auxiliar já usado por esta referência, na mesma coleção e marca (reposição)."""
    m = ex[(ex.ref_n == ref_key(ref)) & ex.aux.str.fullmatch(dig_col + id_marca + r"\d{3}")]
    return m.aux.iloc[0] if len(m) else None


def produto_existente(ex, ref, cor, tam, aux=None):
    """Cod.Produto do Microvix para referência + cor + tamanho, se já cadastrado."""
    m = ex[ex.ref_n == ref_key(ref)]
    if aux:
        m = m[m.aux == aux]
    if not cor and tam == "ÚNICO":
        return m.cod.iloc[0] if len(m) == 1 else None
    fim = norm(f"{cor} {tam}" if cor else tam)
    for d, c in zip(m.desc_n, m.cod):
        if d == fim or d.endswith(" " + fim):
            return c
    return None


# ---------------------------------------------------------------- conferência do app
def ler_conferencia(conf_dir):
    """Lê os JSON exportados de compras/<id>/conf. Retorna lista de dicts de conferência."""
    out = []
    for f in sorted(glob.glob(os.path.join(conf_dir, "**", "*.json"), recursive=True)):
        with open(f, encoding="utf-8") as fh:
            c = json.load(fh)
        c = c.get("data", c)
        if "marca" in c and "ref" in c and "loja" in c:
            out.append(c)
    return out


def grade_conferida(confs):
    """{(MARCA, REF): {(cor, tam): {loja: qtd}}}"""
    grade = {}
    for c in confs:
        k = (norm(c["marca"]), ref_key(c["ref"]))
        for linha in c.get("linhas", []):
            cor = norm_cor(linha.get("cor", ""))
            for t, q in (linha.get("q") or {}).items():
                q = int(q or 0)
                if q:
                    cel = grade.setdefault(k, {}).setdefault((cor, norm_tam(t)), {})
                    cel[int(c["loja"])] = cel.get(int(c["loja"]), 0) + q
    return grade


def ler_fechamento(fech_dir):
    out = {}
    for f in glob.glob(os.path.join(fech_dir, "**", "*.json"), recursive=True):
        with open(f, encoding="utf-8") as fh:
            c = json.load(fh)
        c = c.get("data", c)
        if "marca" in c and "loja" in c:
            out[(int(c["loja"]), norm(c["marca"]))] = bool(c.get("concluida"))
    return out


def divergencias(confs):
    """Linhas em que o contado não bateu com o romaneio."""
    out = []
    for c in confs:
        total = sum(int(q or 0) for l in c.get("linhas", []) for q in (l.get("q") or {}).values())
        esp = int(c.get("esperado") or 0)
        if total != esp or c.get("extra"):
            out.append(dict(loja=int(c["loja"]) + 1, marca=c["marca"], ref=c["ref"], desc=c.get("desc", ""),
                            esperado=esp, contado=total, diferenca=total - esp, obs=c.get("obs", "")))
    return out


def nome_base_microvix(ex, aux):
    """Nome do modelo como está no Microvix (sem cor e tamanho), o mais usado no Cód. Auxiliar."""
    cores, tams = cores_tab(), set(tamanhos())
    nomes = Counter(" ".join(separar(d, cores, tams)[0].split()) for d in ex.desc[ex.aux == aux])
    return nomes.most_common(1)[0][0] if nomes else None
