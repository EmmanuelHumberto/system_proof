# Manutencao e Solucao de Problemas

## Regra De Ouro

Mantenha a planilha fechada antes de rodar scripts administrativos em `tools/`. Muitos deles abrem o Excel por automacao e podem falhar se o arquivo ja estiver aberto.

## Processos Excel Invisiveis

Se o Windows disser que a pasta ou a planilha esta aberta, rode:

```bat
tools\limpar_excel_invisivel.bat
```

Esse utilitario foi criado para encerrar processos Excel invisiveis usados pela automacao. Antes de forcar qualquer encerramento manual, confirme se nao existe uma planilha aberta com trabalho nao salvo.

## Recriar Ou Aplicar Layout

Scripts principais:

- `tools\aplicar_visual_pastel_planilha.vbs`: aplica visual geral, limites de area util e formatacoes.
- `tools\formatar_proteger_formulario_seguro.vbs`: formata campos, protecao e formulario seguro.
- `tools\reorganizar_despesas_12_meses.vbs`: reorganiza a aba `Despesas` com meses, anual, cota parte e observacoes.
- `tools\aplicar_cota_anual_resumo.vbs`: ajusta resumo/rateio para incluir cota parte anual.
- `tools\limpar_planilha_principal_total.vbs`: limpa dados principais preservando estrutura.

Execute com a planilha fechada:

```bat
cscript //Nologo tools\aplicar_visual_pastel_planilha.vbs
```

## Reimportar VBA

O codigo-fonte das macros fica em:

```text
vba\modulos_vba.bas
```

Para reimportar o modulo no workbook:

```bat
cscript //Nologo tools\importar_modulo_excel.vbs
```

No Excel, pode ser necessario habilitar a opcao de confiar no acesso ao modelo de objeto do projeto VBA.

## Diagnosticar Dependencias

Use:

```bat
VerificarDependencias_Windows.bat
```

Problemas comuns:

- `tesseract` nao encontrado: adicionar a pasta de instalacao ao `PATH`.
- idioma `por` ausente: instalar pacote de idioma Portugues do Tesseract.
- `pdftoppm` nao encontrado: instalar Poppler e adicionar `Library\bin` ao `PATH`.
- `tkinter` ausente: reinstalar Python com `tcl/tk and IDLE` marcado.
- `pdfplumber`, `openpyxl` ou `PIL` ausentes: rodar `python -m pip install -r requirements.txt`.

## Quando A Extracao Nao Encontra Valor Ou Recebedor

1. Abra o comprovante original e veja se o texto esta legivel.
2. Rode pela ferramenta grafica e confira a pre-visualizacao.
3. Se for PDF escaneado, confirme se Poppler e Tesseract funcionam.
4. Verifique se a categoria escolhida esta correta.
5. Confira `dados_extraidos.json` da pasta da categoria.
6. Confira se `resumo_extraidos.xlsx` foi atualizado.
7. Se o texto foi extraido mas o campo ficou vazio, a regra da categoria em `_scripts/extrair_categoria.py` precisa ser melhorada.

## JSON E Planilha Por Categoria

Cada pasta de categoria deve conseguir manter:

- `dados_extraidos.json`
- `resumo_extraidos.xlsx`

Quando a extracao roda para uma categoria, estes arquivos sao criados se nao existirem e atualizados se ja existirem.

A mesclagem usa hash do arquivo para identificar comprovantes ja conhecidos. Comprovante novo entra; comprovante igual e atualizado/mesclado; arquivo removido pode continuar como registro se ainda estiver no indice antigo, salvo quando a rotina de limpeza for executada.

## Pasta `_extraidos`

A pasta `_extraidos` e area tecnica. Ela pode atrapalhar testes apenas se voce estiver usando dados antigos como fonte de alguma rotina legada. No fluxo atual, a fila oficial vem dos `dados_extraidos.json` por categoria e depois de `comprovantes.json`.

Para teste limpo, remova:

- `_extraidos\dados.json`
- `comprovantes.json`
- `comprovantes.csv`
- `dados_extraidos.json` de cada pasta de categoria
- `resumo_extraidos.xlsx` de cada pasta de categoria

Nao remova os arquivos originais dos comprovantes.

## Adicionar Nova Categoria

1. Crie a pasta da categoria dentro de `comprovantes/`.
2. Atualize `CONFIG` em `_scripts/config.py` com `pasta_top` e `sub`.
3. Crie ou ajuste a funcao de extracao em `_scripts/extrair_categoria.py`.
4. Adicione a categoria ao dicionario `EXTRACTORS`.
5. Ajuste `_scripts/classificacao.py` para classificar despesas da categoria.
6. Rode a extracao da categoria.
7. Abra o Excel, carregue a fila e importe.

## Adicionar Nova Despesa

Use a aba `Cadastros` quando possivel. Tecnicamente, a despesa tambem precisa existir ou ser criada na aba `Despesas`, e o resumo precisa reconhecer a categoria.

Campos importantes:

- Categoria
- Despesa
- Cota parte? (`Sim` quando divide por 2)

## Estrutura Esperada Da Aba Despesas

- Linha 4: cabecalhos.
- Linhas 5 a 120: despesas.
- Meses: colunas C a N.
- Despesa eventual/anual: coluna O.
- Valor mes: coluna P.
- Valor anual: coluna Q.
- Cota parte?: coluna R.
- Valor cota parte anual: coluna S.
- Valor cota parte mes: coluna T.
- Comprovante: coluna U.
- Observacoes: coluna V.

## Git

O repositorio esta em `main` com remote `origin` apontando para GitHub. Antes de enviar para o GitHub, lembre que comprovantes e planilha podem conter dados pessoais. Faça push somente quando tiver certeza de que esses dados devem ir para o repositorio remoto.
