#!/usr/bin/env python3
"""Mostra o próximo Cód. Auxiliar livre por marca (parte de baixo/peça única e parte de cima).

Uso: python3 scripts/proximos_codigos.py Produtos_Cadastrados.xlsx [--colecao "VERÃO 27"]
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mangabe as mg  # noqa: E402

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("cadastrados")
    ap.add_argument("--colecao", default="VERÃO 27")
    a = ap.parse_args()
    dig = mg.colecoes()[mg.norm(a.colecao)]
    ult = mg.ultimos_codigos(mg.ler_cadastrados(a.cadastrados))
    print(f"{'Marca':<16}{'ID':<4}{'Baixo/única':<14}{'Cima':<8}")
    for nome, m in sorted(mg.marcas().items(), key=lambda x: x[1]["id"]):
        u = ult.get((dig, m["id"]), {"baixo": -1, "cima": 499})
        print(f"{nome:<16}{m['id']:<4}{dig}{m['id']}{u['baixo'] + 1:03d}{'':<8}{dig}{m['id']}{u['cima'] + 1:03d}")
