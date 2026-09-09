# Manual de Uso

## 1. Colocar Novos Comprovantes

Coloque os arquivos na pasta correta dentro de `comprovantes/`. Os formatos aceitos sao:

- `.png`
- `.jpg`
- `.jpeg`
- `.pdf`

Cada comprovante deve ficar na categoria correspondente. A categoria define qual conjunto de regex sera usado na extracao.

## 2. Verificar Dependencias

Execute:

```bat
VerificarDependencias_Windows.bat
```

Se alguma biblioteca Python faltar, aceite a instalacao pelo proprio verificador. Se faltar Tesseract ou Poppler, instale manualmente conforme `docs/INSTALACAO.md`.

## 3. Abrir A Ferramenta Grafica

Execute:

```bat
ExtrairComprovantes_Windows.bat
```

Ou, pelo terminal:

```bat
python _scripts\extrair_gui.py
```

Na tela da ferramenta:

1. Escolha um arquivo ou uma pasta.
2. Escolha a categoria correta.
3. Clique em `Extrair`.
4. Revise data, valor, pagador, recebedor, despesa e periodicidade.
5. Use `Ver comprovante` para conferir o arquivo original.
6. Use `Remover item` se algum item nao deve ir para a fila.
7. Salve a fila para gerar `comprovantes.json` e `comprovantes.csv`.

## 4. Extrair Por Linha De Comando

Extrair um arquivo:

```bat
python _scripts\extrair_ferramenta.py --arquivo "C:\caminho\arquivo.pdf" --categoria "Educação"
```

Extrair uma pasta de uma categoria:

```bat
python _scripts\extrair_ferramenta.py --lote "C:\caminho\pasta" --categoria "Alimentação"
```

Extrair todas as categorias configuradas:

```bat
python _scripts\extrair_ferramenta.py --tudo
```

Ou:

```bat
integrar_extracao.bat
```

## 5. Importar No Excel

1. Abra `Planilha_Comprovantes_VBA.xlsm`.
2. Habilite macros.
3. Va para a aba `Import`.
4. Clique em `Fila JSON` para carregar `comprovantes.json`.
5. Confira a aba `Fila`.
6. Use `Linha selecionada` para importar apenas uma linha.
7. Use `Pendentes` para importar todos os registros ainda nao importados.

## 6. O Que A Importacao Preenche

Ao importar, o VBA:

- cria uma linha em `Controle` com os dados do comprovante;
- registra o hash oculto para evitar duplicidade;
- encontra ou cria a despesa em `Despesas`;
- soma o valor no mes correto ou em `Despesa eventual/anual`;
- marca `Comprovante` como `Anexado`;
- inclui os numeros de controle em `Observacoes`;
- atualiza formulas de `Despesas` e `Resumo e Rateio`.

## 7. Cota Parte

Na aba `Despesas`, a coluna `Cota parte?` controla o calculo:

- `Sim`: divide a despesa por 2.
- vazio ou outro valor: considera o valor integral.

A planilha calcula:

- `Valor mes`: soma dos meses mais despesa eventual/anual, dividido por 12.
- `Valor anual`: soma dos meses mais despesa eventual/anual.
- `Valor cota parte anual`: se `Sim`, divide o valor anual por 2; se nao, usa o valor anual integral.
- `Valor cota parte mes`: se `Sim`, divide o valor mensal por 2; se nao, usa o valor mensal integral.

## 8. Abrir Comprovantes E Registros

- Na aba `Controle`, selecione uma linha e use `Abrir arquivo` para abrir o comprovante original.
- Na aba `Despesas`, ao clicar em uma celula de `Comprovante` ou `Observacoes`, a macro pode abrir/filtrar os registros relacionados na aba `Controle`.
- Quando uma despesa possui varios numeros em `Observacoes`, a aba `Controle` deve ser filtrada para mostrar somente aqueles registros.

## 9. Excluir Um Registro Importado

Use a exclusao a partir da aba `Controle`:

1. Selecione a linha do registro que deseja excluir.
2. Acione a macro/botao de exclusao configurado para controle.
3. O sistema remove o registro do `Controle`.
4. O valor correspondente e abatido da aba `Despesas`.
5. O numero de controle e removido de `Observacoes`.
6. O item tambem pode ser removido da `Fila` quando encontrado por hash ou arquivo.

## 10. Cadastros

A aba `Cadastros` existe para manutencao de categorias/despesas e marcador de cota parte. Use-a para:

- criar nova despesa;
- ajustar nome de despesa;
- marcar se a despesa tem cota parte;
- sincronizar as despesas com a aba `Despesas`.

## 11. Limpeza

`Limpar dados` limpa os dados importados da planilha principal. A fila pode permanecer como historico operacional da ultima extracao, mas o controle efetivo de importacao fica em `Controle`, por hash.

Para limpar dados extraidos e indices por categoria antes de um teste grande, apague somente os arquivos gerados (`comprovantes.json`, `comprovantes.csv`, `dados_extraidos.json`, `resumo_extraidos.xlsx` e `_extraidos/dados.json`). Os comprovantes originais nao devem ser apagados.
