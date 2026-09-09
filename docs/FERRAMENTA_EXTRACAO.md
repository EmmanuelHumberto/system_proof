# Manual Da Ferramenta De Extracao

A ferramenta de extracao e a porta de entrada dos comprovantes. Ela fica em `_scripts/extrair_gui.py` e pode ser aberta por `ExtrairComprovantes_Windows.bat`.

## Quando Usar

Use a ferramenta quando quiser:

- importar um novo comprovante isolado;
- importar uma pasta inteira de uma categoria;
- conferir visualmente dados antes de mandar para o Excel;
- gerar `comprovantes.json` somente com a selecao revisada;
- atualizar `dados_extraidos.json` e `resumo_extraidos.xlsx` da categoria.

## Campos Da Tela

- `Arquivo ou pasta`: caminho do arquivo ou diretorio selecionado.
- `Categoria`: define qual regra de extracao sera usada.
- `Filtrar categoria`: filtra os itens exibidos na grade.
- `Arquivo`: nome/caminho do comprovante.
- `Data`: data extraida no formato `AAAA-MM-DD`.
- `Valor`: valor monetario extraido.
- `Recebedor`: estabelecimento, pessoa ou empresa recebedora.
- `Despesa`: classificacao usada na planilha.
- `Periodicidade`: `Mensal` ou `Anual`.
- `Preview`: imagem previa do comprovante quando suportado.

## Botoes Principais

- `Arquivo`: seleciona um comprovante individual.
- `Pasta`: seleciona uma pasta com varios comprovantes.
- `Extrair`: roda OCR/leitura de PDF e aplica as regras da categoria.
- `Ver comprovante`: abre o arquivo selecionado.
- `Remover item`: tira o item da fila antes de salvar.
- `Aplicar edicao`: aplica alteracoes manuais feitas nos campos da tela.
- `Salvar tudo`: grava `comprovantes.json` e `comprovantes.csv` com os itens revisados.

## Fluxo Para Um Arquivo

1. Clique em `Arquivo`.
2. Escolha o comprovante.
3. Selecione a categoria correta.
4. Clique em `Extrair`.
5. Confira os campos.
6. Corrija manualmente se necessario.
7. Clique em `Aplicar edicao`.
8. Clique em `Salvar tudo`.
9. Abra o Excel, carregue a fila e importe.

## Fluxo Para Uma Pasta

1. Clique em `Pasta`.
2. Escolha a pasta que contem comprovantes de uma unica categoria.
3. Selecione a categoria correspondente.
4. Clique em `Extrair`.
5. Revise a grade.
6. Remova registros indevidos.
7. Ajuste campos incorretos.
8. Clique em `Salvar tudo`.
9. No Excel, carregue `comprovantes.json` e importe.

## Fluxo Para Todas As Categorias

Use o botao/macro de extracao completa ou rode:

```bat
python _scripts\extrair_ferramenta.py --tudo
```

Esse fluxo atualiza os indices de todas as categorias configuradas e recria a fila consolidada.

## Como A Ferramenta Evita Duplicidade

Cada arquivo recebe um hash MD5. Esse hash e salvo no indice da categoria e tambem na fila consolidada. No Excel, a importacao compara o hash com a aba `Controle` antes de gravar.

## O Que Fazer Quando Um Campo Sair Vazio

- Confira se a categoria esta correta.
- Abra o comprovante e veja se o texto esta legivel.
- Se for PDF escaneado, teste `pdftoppm` e `tesseract`.
- Veja se o dado aparece em `texto_extraido` no JSON da categoria.
- Se o texto existe mas o campo nao preenche, ajuste a regra da categoria em `_scripts/extrair_categoria.py`.

## Arquivos Gerados

Ao salvar/extrair, a ferramenta pode gerar ou atualizar:

- `comprovantes.json`
- `comprovantes.csv`
- `dados_extraidos.json` na pasta da categoria
- `resumo_extraidos.xlsx` na pasta da categoria

## Limpeza Para Teste

Para testar uma carga do zero, feche o Excel e remova arquivos gerados, mantendo os comprovantes originais:

```text
comprovantes.json
comprovantes.csv
_extraidos\dados.json
comprovantes\...\dados_extraidos.json
comprovantes\...\resumo_extraidos.xlsx
```

Depois rode a extracao novamente.
