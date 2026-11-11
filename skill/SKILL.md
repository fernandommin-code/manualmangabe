---
name: "cadastro-compras-mangabe"
description: "Rotina semanal de compras da Mangabê, do romaneio à entrada de estoque no Microvix: planilha de cadastro a partir de fotos de romaneios/cupons e NF-e com preço por marca; conferência das lojas por cor × tamanho (app ou folha impressa); checagem de repique (produto que já existe); arquivo de importação de produtos do Microvix; arquivo de entrada por loja. Usar sempre que Fernando mandar pedidos, notas, a planilha do dia, fotos de folhas de conferência, o relatório Produtos Cadastrados, inconsistências de importação do Microvix ou pedir cadastro/entrada de alguma marca — mesmo que ele não diga 'cadastro'."
---

# Cadastro de compras Mangabê

Fernando compra às terças e manda, marca por marca, fotos dos romaneios/cupons e as NF-e ("<MARCA>, margem de X%" + imagens). O trabalho vai até o fim:

1. planilha do dia (sem grade) → 2. prévia de códigos e repiques → 3. conferência das lojas por cor × tamanho (app; folha de papel como reserva) → 4. divergências → 5. arquivo de cadastro do Microvix → 6. a moça do cadastro importa → 7. arquivo de entrada por loja → 8. ela dá a entrada.

Os scripts e tabelas estão nesta skill (`scripts/`, `tabelas/`) e também no repositório `manualmangabe` (GitHub fernandommin-code/manualmangabe, com o manual em README.md). Rode os scripts com `python3 <skill>/scripts/...`; eles leem `tabelas/` da pasta ao lado. Dependências: `pip install openpyxl pandas python-calamine reportlab`.

## 1. Planilha do dia

- Uma planilha por dia de compra: `MM_DD_Cadastro_de_Produtos.xlsx`. "Nova planilha" = copiar o layout da anterior e limpar as linhas de produtos.
- Layout (não mudar): E1 "CADASTRO DE PRODUTOS", J1 a data; cada marca é um bloco: nome da marca na coluna A, linha de títulos `Lj 1 | Lj 2 | Lj 3 | Cód. Interno | Descrição do Produto | Coleção | Referência | Preço Custo | Preço Venda | NCM`, uma linha por referência e `TOTAL PEÇAS` com `=SUM()` em A–C. Uma linha em branco entre blocos, todas as marcas na mesma aba.
- Descrição centralizada, Calibri 11, bordas finas, preços em R$ contábil, NCM como texto, Cód. Interno em branco. Rodar o recalc do skill xlsx e reenviar.
- Planilha editada à mão por Fernando passa a ser a base: só mudar o que ele pedir.

## 2. Transcrição do romaneio

- **Loja pelo CNPJ:** Loja 1 = 15.478.784/0001-35 (Praiamar); Loja 2 = 15.478.784/0002-16 (Balneário); Loja 3 = 63.860.624/0001-07 (Miramar). Observação do pedido ("LJ 02") prevalece; avisar a divergência.
- **Uma linha por referência**, somando cores e tamanhos. Devoluções fora (só mencionar o crédito).
- **Descrição** em até 3 palavras, com o nome do modelo (evitar "BLUSA", "CALÇA" sozinhos: no caixa ficam iguais). Nome cortado = palpite em amarelo; trocar pela descrição da NF quando chegar.
- **Coleção:** VERÃO 27 salvo indicação. **Preço Custo:** sempre bruto. **NCM:** só da NF.
- Conferir soma de peças e subtotal bruto por loja contra o cupom; apontar desconto não aplicado, valor diferente na nota, marca misturada, reposição.

## 3. Preço de venda

Ler `/topics/precificacao-fornecedores.md` na memória antes de precificar; a instrução do dia prevalece.
- Margem M sobre custo líquido: venda = custo bruto × (1 − desconto) ÷ (1 − M). "margem de X%" → terminado em 9 mais próximo; "margem mínima" → para cima até terminar em 9.
- Markup K: custo bruto × K, terminado em 9. Vicky: bijuteria ×2,5 arredondado para cima até 5 ou 9, mínimo R$ 25; bolsas 2,2×.
- Preço fixado por Fernando: aplicar e informar a margem resultante.
- Descontos de compra: em `tabelas/marcas.csv` (TOVA 15, HIT 10, Manacá 15, MYSK 7, Donna Ritz 15, Salgunamu 20, REKIM 10, GULE 15, Be Free 15, Kaele 7, Other Girls 10, Vicky 10, FPM 15, Esmeral 13, Hush 10). Marca nova: ler no cupom e acrescentar na tabela.

Resposta curta: tabela da marca, totais por loja, só os pontos que pedem ação.

## 4. Prévia: códigos e repiques (logo depois da planilha)

`scripts/previa_codigos.py PLANILHA PRODUTOS_CADASTRADOS [saida.xlsx]` mostra o Código Auxiliar de cada referência (novo ou reposição) e aponta descrição genérica, nome repetido, NCM faltando, referência escrita diferente do Microvix, reposição com nome diferente.

**Relatório Produtos Cadastrados:** Suprimentos > Estoque > Relatórios > Produtos Cadastrados, Excel. Para numerar, ele tem que ser **do mesmo dia e completo da marca (sem filtro de data nem de coleção)** — vários relatórios que chegam vêm filtrados "só hoje" com poucas linhas; nesse caso, juntar com o relatório completo anterior (os códigos do Microvix são sequenciais, então dá para ver se algo foi criado no meio). Não abre no openpyxl: o script usa calamine.

**Código Auxiliar** (6 dígitos, igual para toda a grade do modelo): coleção (6 = Verão 27, 5 = Inverno 26) + ID da marca (`tabelas/marcas.csv`) + nº do modelo (000–499 parte de baixo/peça única; 500–999 parte de cima). Ex.: primeiro vestido FPM do Verão 27 = 639000, primeira blusa = 639500. Parte de cima/baixo pela 1ª palavra conhecida da descrição (`tabelas/tipos_peca.csv`; "MAXI COLAR" → colar). O próximo número é o maior usado + 1.

**Repique (reposição):** referência que já existe **na mesma marca e coleção** reaproveita o Código Auxiliar existente (comparação sem zeros à esquerda: 09404 = 9404; mesma referência em outra marca, ex. 08297 FPM × KAELE, não conta). Cor × tamanho que já existe **não vai para o cadastro**, só para a entrada. Cor ou tamanho novo vira produto novo com o mesmo Código Auxiliar, **o nome e a referência como estão no Microvix**. O script detecta produto repetido no Microvix (mesma ref/cor/tamanho duas vezes), usa o de nome principal e avisa.

## 5. Conferência das lojas

**App (padrão):** https://claude.ai/artifact/Bw297fhPHqWy35JJEU1hds (capability db). Vendedoras precisam de acesso de Colaborador. Por referência, escolhem a cor (lista com os nomes do Microvix) e tocam + nos tamanhos (XPP–GG; botão troca para numerado ou único); se não bater com o romaneio, observação obrigatória; "Concluir marca" por loja. Reposições mostram "já no sistema: <cores>".

- Abrir a compra: `scripts/carregar_app.py PLANILHA saida/ --data AAAA-MM-DD --cadastrados RELATORIO` gera `compra.json`, `marcas/<SLUG>.json` (com `exist`) e, com `--conf`, conferências transcritas. Gravar com `ArtifactData` batch (até 50 por chamada, `file_path`): `compras/<id>`, `compras/<id>/marcas/<SLUG>`, `compras/<id>/conf/<loja>~<SLUG>~<REF>`. Documento que já existe exige `if_version`.
- Ler: `ArtifactData` list de `compras/<id>/conf` com `out_dir` e `query.limit` 1000 (o "~" vira "@" no nome do arquivo; os scripts leem a pasta direto). Fechamento em `compras/<id>/fechamento/<loja>~<SLUG>` {concluida}.

**Folha impressa (reserva):** `scripts/gerar_folhas.py PLANILHA saida/ --data DD/MM --cadastrados RELATORIO` gera um PDF por loja já preenchido (marca, ref, descrição, quantidade do romaneio, cores já cadastradas). Ao transcrever foto: `"` / `/` = igual à linha de cima; nomes do fornecedor viram os do Microvix ("pepper red" → VERMELHO, "fresh yellow" → AMARELO, "off" → OFF WHITE); ampliar e girar a foto com PIL antes de ler números duvidosos; números rabiscados: escolher o que fecha com o total circulado/romaneio e avisar. Transcrição vai para JSON no formato do app (e pode ser carregada no app com `--conf`).

## 6. Divergências

Antes de gerar o cadastro: listar por marca e loja o que não bateu (contado × romaneio, com a observação) e referências sem conferência, para Fernando cobrar o fornecedor ou mandar recontar. Pode-se gerar o cadastro de uma marca só com as lojas já conferidas; o que vier depois entra num cadastro pequeno.

## 7. Arquivo de cadastro do Microvix

`scripts/gerar_cadastro.py --planilha P --conf DIR --modelo ModeloImportacaoProduto.xlsx --cadastrados R --saida S` → `Importacao_<MARCA>.xlsx` (um produto por cor × tamanho conferido) + `Resumo_cadastro.xlsx` (códigos, divergências, já cadastrados, pendências). O modelo vem de Suprimentos > Estoque > Importação de Produtos > Nova Importação (67 colunas).

O que aprendemos testando (GULE e FPM, 11/11/2026):
- **Descrição = só o modelo** ("VESTIDO LONGO RENDA"). O Microvix acrescenta cor e tamanho sozinho; mandando completo, fica duplicado.
- **Colunas obrigatórias** além de descrição/ref/marca/coleção/tamanho/cor/custo/preço: Fornecedor (`tabelas/marcas.csv`, coluna fornecedor_microvix: GULE 22, FPM 19), Configuração tributária 1, Unidade de venda UN, Setor GERAL, Espessura GERAL (`tabelas/padroes_microvix.csv`), Linha (tipo da peça, `tabelas/tipos_peca.csv`) e **NCM**. Sem NCM a linha volta; sem nota, usar o NCM que o mesmo fornecedor deu a peça parecida na mesma compra e avisar para conferir com a nota.
- **Custo com ICMS = custo líquido** (bruto × (1 − desconto da marca)) e **Desconto (%) = 0**, como a moça cadastra à mão. A entrada de estoque usa o Custo com ICMS e ignora o Desconto (%): com bruto + desconto, a FPM entrou a preço cheio. Código de barras em branco: o Microvix usa o próprio código do produto.
- **Na tela de importação, Identificador dos Produtos = "Código do produto"**. Linha sem Código cria; linha com Código altera aquele produto (serve para corrigir nome/NCM). Nunca "Referência" nem "Código auxiliar": a grade inteira tem a mesma referência e o mesmo auxiliar e o Microvix junta tudo num produto só (aconteceu na GULE).
- Resultado: verde = criados, amarelo = alterados, vermelho = inconsistências (pedir print de "Visualizar inconsistências" — o arquivo de inconsistências não traz o motivo). Esperado num cadastro novo: tudo verde.
- Primeira vez de uma marca: perguntar o código do fornecedor e conferir se as Linhas existem com esse nome.

## 8. Entrada de estoque

Depois da importação, pedir o Produtos Cadastrados de novo (só da marca, sem filtro de data). `scripts/gerar_entrada.py --conf DIR --cadastrados R --planilha P --saida S` → por loja `Entrada_LojaN.txt` (**um código por linha, uma linha por peça** — único formato que a tela aceitou; "código;quantidade" é recusado) e `Entrada_LojaN.csv` (lista para conferir). Inclui repiques e produtos novos numa entrada só por loja; `Entrada_nao_encontrados.csv` lista o que não achou código. Lembrar: conferir o total na tela antes de encerrar, entrada encerrada não se altera.

## Outros

- `scripts/proximos_codigos.py RELATORIO` — próximo código livre por marca.
- `scripts/simular_com_historico.py RELATORIO MODELO saida/` — refaz as últimas referências de cada marca e compara com o Microvix; também lista repetidos e nomes fora do padrão.
- Cor nova no cadastro: acrescentar em `tabelas/cores.csv` (e na lista do app). Marca nova: linha em `tabelas/marcas.csv` com ID, desconto e fornecedor.
- Ao terminar uma mudança nas regras, atualizar também o repositório `manualmangabe` (README e scripts) para os dois não divergirem.
