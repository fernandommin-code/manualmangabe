# Manual Mangabê — cadastro de produtos por grade (cor × tamanho)

Rotina semanal que vai do romaneio de compras até a entrada de estoque no Microvix,
com a conferência das lojas feita por cor e tamanho no app.

## Visão geral

| # | Etapa | Quem | Resultado |
|---|---|---|---|
| 1 | Transcrever o romaneio | Fernando + Claude | Planilha do dia, uma linha por referência, total por loja |
| 2 | Abrir a conferência no app | Claude | Compra carregada no app, com o esperado de cada referência por loja |
| 3 | Conferir por cor × tamanho | Vendedoras | Grade contada lançada no app; marca concluída por loja |
| 4 | Divergências | Claude → Fernando | Lista do que faltou ou sobrou, para cobrar o fornecedor |
| 5 | Gerar o cadastro | Claude (`scripts/gerar_cadastro.py`) | `Importacao_<MARCA>.xlsx`, um por marca |
| 6 | Importar no Microvix | Moça do cadastro | Produtos criados, um por cor × tamanho |
| 7 | Gerar a entrada | Claude (`scripts/gerar_entrada.py`) | `Entrada_LojaN.csv` / `.txt` por loja |
| 8 | Dar entrada | Moça do cadastro | Estoque nas 3 lojas |

---

## 1. Planilha do dia (sem grade)

- Arquivo `MM_DD_Cadastro_de_Produtos.xlsx`, uma aba, uma marca abaixo da outra.
- Cada bloco tem a linha com o nome da marca na coluna A, os títulos
  `Lj 1 | Lj 2 | Lj 3 | Cód. Interno | Descrição do Produto | Coleção | Referência | Preço Custo | Preço Venda | NCM`,
  uma linha por referência e uma linha `TOTAL PEÇAS`.
- **Preço Custo** é o custo **bruto** do romaneio. O desconto da marca entra no cadastro (etapa 5).
- **Descrição** curta, sem cor e sem tamanho (ex.: `VESTIDO MIDI LINHO`). Cor e tamanho entram no cadastro.
- **NCM** só da nota fiscal.

**Romaneio que já vem por cor e tamanho:** a grade do romaneio também é carregada no app.
As cores já aparecem preenchidas e cada tamanho mostra o que o romaneio diz (`P · rom. 2`).
Quando o romaneio não traz grade, a vendedora digita as cores.

### Prévia (logo depois da planilha)

`scripts/previa_codigos.py` mostra o Código Auxiliar que cada referência vai receber e o que é reposição.
Também aponta o que precisa de ajuste antes do cadastro:
- descrição genérica ("CALÇA", "TOP")
- o mesmo nome em duas referências
- NCM faltando
- referência escrita diferente do Microvix (ex.: `09404` × `9404`)
- reposição com nome diferente do cadastrado

**Reposição:** o cadastro usa o nome e a referência que já estão no Microvix, para a grade não ficar com dois nomes.
A comparação de referência ignora zeros à esquerda.

## 2–3. Conferência — folha impressa (padrão) ou app

### Folha já preenchida

`scripts/gerar_folhas.py` gera um PDF por loja (A4 deitado) a partir da planilha do dia.
Vêm impressos a marca, a referência, a descrição e a quantidade do romaneio daquela loja.
As vendedoras escrevem só a **cor** (uma linha por cor) e as quantidades por tamanho
(XPP/34 até GG/44, "Outro" para tamanhos fora da grade).
Peça que não está na lista vai nas linhas "fora do rom." de cada marca.
A Vicky tem uma linha por referência, só com a quantidade.

Depois de conferir, mandam a foto das folhas. O Claude transcreve para o mesmo formato do app
(`compras/<id>/conf`) e segue para as divergências e o cadastro.

### App (padrão a partir de 11/11)

`scripts/carregar_app.py` prepara a compra para o app (com as cores já cadastradas nas reposições)
e pode carregar conferências feitas em papel. O Claude grava com ArtifactData e lê de volta direto
para os scripts de cadastro e entrada. A skill da rotina está em `skill/SKILL.md`.


App: <https://claude.ai/artifact/Bw297fhPHqWy35JJEU1hds> (cópia em `app/conferencia.html`).
As vendedoras precisam de acesso de **Colaborador**.

1. Escolher a compra e a loja.
2. Em cada referência, para cada cor: escrever a cor e tocar `+` / `−` em cada tamanho.
   A grade padrão é XPP–GG. O botão de grade troca para numerado (34–46) ou tamanho único,
   e "+ tamanho" acrescenta outros.
3. Salvar. Se não bater com o romaneio, a observação é obrigatória.
4. Peça que veio sem estar no romaneio vai em "+ referência".
5. Ao terminar a marca na loja, tocar em **Concluir marca**.

**Escrever as cores sempre com o mesmo nome do Microvix** (a lista aparece no campo):
OFF WHITE, PRETO, AZUL, AZUL MARINHO, MARROM, AMARELO, VERDE, BEGE, ROSA, VERMELHO, CRU,
LARANJA, BRANCO, MANTEIGA, PEROLA, LILAS, ROXO, CARAMELO, VINHO, DOURADO, TERRACOTA, AREIA, CREME.
O nome da cor vira parte do nome do produto: "AZUL" e "AZUL MARINHO" viram produtos diferentes.

Dados gravados pelo app (coleção `compras/<AAAA-MM-DD>`):

- `marcas/<SLUG>`: `{marca, ordem, itens:[{ref, desc, lj:[l1,l2,l3], tam:"L"|"N"|"U", grade?:[{cor, q:{tam:qtd}}]}]}`
- `conf/<loja>~<MARCA>~<REF>`: `{loja, marca, ref, esperado, linhas:[{cor, q:{tam:qtd}}], obs, total}`
- `fechamento/<loja>~<MARCA>`: `{concluida}`

## 4. Divergências

O `Resumo_cadastro.xlsx` (gerado na etapa 5) tem a aba **Divergências**, com loja, marca,
referência, esperado, contado e observação, para cobrar o fornecedor.

## 5. Cadastro — Código Auxiliar e arquivo de importação

### Regra do Código Auxiliar (6 dígitos, igual para toda a grade da referência)

```
6        39        000
│        │         └── nº do produto: 000–499 parte de baixo e peça única
│        │                            500–999 parte de cima
│        └── ID da marca (tabelas/marcas.csv) — 39 = FPM
└── coleção (tabelas/colecoes.csv) — 6 = Verão 27, 5 = Inverno 26
```

Exemplo: o primeiro vestido da FPM no Verão 27 é `639000`, o próximo é `639001`, e assim por diante.
A primeira blusa é `639500`.

- O script continua a partir do **maior número já usado** no relatório Produtos Cadastrados,
  separado por marca, coleção e faixa (baixo/cima). `scripts/proximos_codigos.py` mostra o próximo livre.
- **Parte de cima ou de baixo** sai da 1ª palavra da descrição (`tabelas/tipos_peca.csv`).
  BLUSA, CAMISA, REGATA, TOP, CROPPED, BODY, JAQUETA, BLAZER, COLETE, KIMONO, CORSET etc. são de cima.
  VESTIDO, CALÇA, SHORTS, SAIA, MACACÃO e CONJUNTO são de baixo ou peça única.
  Palavra fora da tabela: o script assume parte de baixo e avisa em Pendências.
- **Reposição:** se a referência já existe na mesma marca e coleção, o script reaproveita o código.
  Cor × tamanho que já existe não é cadastrado de novo, só recebe entrada.

### O que vai no arquivo (um produto por cor × tamanho)

| Coluna do modelo | Valor |
|---|---|
| Descrição | só o modelo, ex.: `VESTIDO MIDI LINHO`. **O Microvix acrescenta a cor e o tamanho sozinho** e o produto fica `VESTIDO MIDI LINHO PRETO P` (teste da GULE) |
| Referência | referência do fornecedor |
| Cód. Auxiliar | regra acima |
| Marca, Coleção, Tamanho, Cores | da planilha e da conferência |
| Custo com ICMS (R$) | custo bruto |
| Desconto (%) | desconto da marca (`tabelas/marcas.csv`), para o produto já nascer com custo líquido |
| Preço de venda R$ | da planilha |
| NCM | da planilha, quando houver |
| Contabiliza saldo em estoque | Sim |
| Código de barras | **em branco: o Microvix gera** |

Só entra o que foi **conferido**. Tamanho que não chegou não é cadastrado: produto sem entrada
fica com o custo cheio no Microvix.

> **Atenção — o Microvix reconhece o produto pelo Código Auxiliar.** Se o arquivo trouxer um
> Código Auxiliar que já existe, a importação **altera** os produtos existentes (amarelo = alterados)
> em vez de criar novos. Por isso o relatório Produtos Cadastrados usado para numerar tem que ser
> **exportado no mesmo dia, completo (Verão 27, sem filtro de data)**, depois de qualquer cadastro manual.
> Com a coluna Código preenchida, a importação altera aquele produto (serve para corrigir nomes).
>
> **Identificador dos Produtos = "Código do produto"** (sempre). Linhas sem Código são criadas;
> linhas com Código alteram aquele produto. Nunca usar "Referência" nem "Código auxiliar": toda a grade
> de um modelo tem a mesma referência e o mesmo código auxiliar, e o Microvix junta tudo num produto só.
> Resultado certo da GULE (11/11): 18 verdes (criados), 11 amarelos (alterados), 0 vermelhos.

## 6. Importar no Microvix (moça do cadastro)

Suprimentos > Estoque > Importação de Produtos > Nova Importação, **uma importação por marca**.

O Microvix recusa a linha se faltar **Fornecedor, Configuração tributária, Unidade de venda,
Setor, Linha ou Espessura** (teste da GULE em 10/11). Por isso o arquivo já sai com essas colunas preenchidas:
- `tabelas/padroes_microvix.csv`: valores iguais para todos os produtos (escritos exatamente como no Microvix)
- `tabelas/marcas.csv`, coluna `fornecedor_microvix`: o fornecedor de cada marca

Se alguma estiver vazia, o `Resumo_cadastro.xlsx` avisa em Pendências.
Se der erro, clicar em "Visualizar inconsistências" e mandar um print com o motivo de cada linha.

## 7–8. Entrada de estoque

1. Exportar de novo o relatório **Produtos Cadastrados** (Coleção Verão 27, Excel), já com os produtos importados.
2. Rodar `scripts/gerar_entrada.py`. Ele troca cada cor × tamanho conferido pelo **Cod.Produto** do Microvix e gera:
   - `Entrada_LojaN.csv`: `Cod.Produto;Descrição;Quantidade`
   - `Entrada_LojaN.txt`: um Cod.Produto por peça, para leitura em lote
   - `Entrada_nao_encontrados.csv`: o que não achou produto (normalmente cor escrita diferente)
3. Dar a entrada em cada loja importando o `Entrada_LojaN.txt`: **um código por linha, uma linha por peça** (testado na GULE em 11/11). O formato `código;quantidade` não é aceito.

---

## Como rodar

```bash
pip install -r requirements.txt

# prévia da planilha do dia (antes da conferência): códigos, reposições e pendências
python3 scripts/previa_codigos.py 11_10_Cadastro_de_Produtos.xlsx Produtos_Cadastrados.xlsx Previa.xlsx

# folhas de conferência já preenchidas (uma por loja)
python3 scripts/gerar_folhas.py 11_10_Cadastro_de_Produtos.xlsx folhas/ --data 10/11

# próximo código livre por marca
python3 scripts/proximos_codigos.py Produtos_Cadastrados.xlsx

# etapa 5
python3 scripts/gerar_cadastro.py --planilha 11_10_Cadastro_de_Produtos.xlsx \
  --conf conf/ --fechamento fech/ --modelo ModeloImportacaoProduto.xlsx \
  --cadastrados Produtos_Cadastrados.xlsx --saida saida/

# etapa 7 (depois de importar)
python3 scripts/gerar_entrada.py --conf conf/ --planilha 11_10_Cadastro_de_Produtos.xlsx \
  --cadastrados Produtos_Cadastrados_pos_importacao.xlsx --saida saida/

# testes
python3 -m pytest tests/

# teste com o histórico real (não mexe no Microvix): refaz as últimas referências
# de cada marca e compara com o que está cadastrado; também revisa o cadastro
python3 scripts/simular_com_historico.py Produtos_Cadastrados.xlsx ModeloImportacaoProduto.xlsx saida_teste/
```

`conf/` e `fech/` são os JSON exportados do app (`compras/<id>/conf` e `compras/<id>/fechamento`).

## Tabelas (`tabelas/`)

| Arquivo | Conteúdo |
|---|---|
| `marcas.csv` | ID de 2 dígitos e desconto de compra de cada marca |
| `colecoes.csv` | 1º dígito do Código Auxiliar por coleção |
| `tipos_peca.csv` | 1ª palavra da descrição → parte de cima ou de baixo |
| `cores.csv` | cores no padrão do Microvix |
| `tamanhos.csv` | ordem dos tamanhos |

Marca nova ou coleção nova: acrescentar uma linha na tabela, sem mexer no script.
