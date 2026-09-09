# Arquitetura

## Visao Geral

O sistema tem duas partes principais:

- Python: extrai texto, aplica regras por categoria, gera JSON/CSV e mantem indices por pasta.
- Excel/VBA: carrega a fila, permite revisao, importa para `Controle` e soma nas linhas de `Despesas`.

O Python nao grava diretamente nas abas principais da planilha. Ele gera a fila. A gravacao final na planilha e feita pelas macros VBA para manter rastreabilidade e evitar duplicidade.

## Diagrama Do Fluxo

```text
comprovantes/ categoria
        |
        v
_scripts/extrair_ferramenta.py ou extrair_gui.py
        |
        v
OCR / leitura PDF / regex por categoria
        |
        v
dados_extraidos.json + resumo_extraidos.xlsx por pasta
        |
        v
_scripts/extrair_tudo.py
        |
        v
comprovantes.json + comprovantes.csv
        |
        v
Excel VBA: CarregarFilaJson
        |
        v
Aba Fila -> Importar pendentes/linha selecionada
        |
        v
Aba Controle + Aba Despesas + Resumo e Rateio
```

## Componentes Python

### `_scripts/config.py`

Define caminhos principais e o mapa de categorias. A base e detectada automaticamente pelo diretorio pai de `_scripts`.

Arquivos principais:

- `PLANILHA_VBA`: caminho da planilha `.xlsm`.
- `CSV`: `comprovantes.csv`.
- `EXTRAIDOS`: pasta `_extraidos`.
- `DADOS`: `_extraidos/dados.json`.
- `COMPROVANTES_DIR`: pasta `comprovantes/`.
- `CONFIG`: relacao entre categoria e pasta fisica.

### `_scripts/verificar_deps.py`

Diagnostica dependencias Python e binarios externos. Pode instalar bibliotecas Python com `--instalar`, mas nao instala Tesseract/Poppler automaticamente.

### `_scripts/extrair_gui.py`

Interface grafica de extracao. Permite selecionar arquivo ou pasta, escolher categoria, revisar campos, pre-visualizar comprovante, remover item e salvar a fila para o Excel.

### `_scripts/extrair_ferramenta.py`

Versao de linha de comando/interativa da extracao. Modos:

```bat
python _scripts\extrair_ferramenta.py
python _scripts\extrair_ferramenta.py --arquivo "caminho" --categoria "Educação"
python _scripts\extrair_ferramenta.py --lote "caminho_da_pasta" --categoria "Alimentação"
python _scripts\extrair_ferramenta.py --tudo
```

### `_scripts/extrair_categoria.py`

Contem os extratores por categoria. Cada categoria tem uma funcao propria de regex e heuristicas:

- `extrair_alimentacao`
- `extrair_comunicacao`
- `extrair_lazer`
- `extrair_moradia`
- `extrair_saude`
- `extrair_transporte`
- `extrair_educacao`
- `extrair_vestuario`

### `_scripts/indices_categoria.py`

Mantem o indice por categoria:

- lista arquivos validos (`.png`, `.jpg`, `.jpeg`, `.pdf`);
- calcula hash MD5;
- normaliza campos;
- salva `dados_extraidos.json`;
- salva `resumo_extraidos.xlsx`;
- mescla novos resultados sem perder registros existentes.

### `_scripts/extrair_tudo.py`

Le os `dados_extraidos.json` das categorias e gera a fila consolidada:

- `comprovantes.json`
- `comprovantes.csv`

### `_scripts/integrar_extracao_planilha.py`

Pipeline completo usado pela macro e pelos `.bat`:

1. Reprocessa arquivos fisicos.
2. Atualiza JSON e planilha de cada categoria.
3. Gera fila consolidada para o Excel.

## Componentes VBA

O modulo principal fica em `vba/modulos_vba.bas` e tambem esta importado dentro de `Planilha_Comprovantes_VBA.xlsm`.

Macros principais:

- `ExtrairDados`: executa o pipeline Python.
- `AbrirFerramentaExtracao`: abre a ferramenta grafica Python.
- `CarregarFilaJson`: carrega `comprovantes.json` para a aba `Fila`.
- `ImportarSelecionadoDaFila`: importa apenas a linha ativa da aba `Fila`.
- `ImportarTodosPendentes`: importa todas as linhas pendentes.
- `AbrirComprovante`: abre o arquivo vinculado ao registro selecionado.
- `AbrirRegistrosDaObservacao`: filtra a aba `Controle` pelos numeros citados em `Observacoes`.
- `ExcluirRegistroControleSelecionado`: remove um registro do controle e desfaz o valor na despesa correspondente.
- `LimparDados`: limpa registros importados de `Controle` e valores de `Despesas`.
- `AbrirCadastros`: abre a area de cadastro de categorias/despesas.
- `SalvarCadastroDespesa`: cria ou atualiza despesa e marcador de cota parte.
- `SincronizarCadastrosDespesas`: sincroniza cadastros com a aba `Despesas`.
- `DesbloquearEdicao` e `BloquearEdicao`: alternam protecao das abas.

## Abas Da Planilha

- `Import`: painel de operacao.
- `Fila`: registros extraidos aguardando importacao.
- `Controle`: trilha detalhada dos comprovantes importados.
- `Despesas`: matriz por despesa e mes, com totais, cota parte e observacoes.
- `Resumo e Rateio`: resumo por categoria e calculo de capacidade contributiva/rateio.
- `Config`: parametros operacionais.
- `Cadastros`: manutencao de categoria, despesa e cota parte.

## Modelo De Dados Da Fila

`comprovantes.json` usa uma lista `comprovantes` com campos:

- `id`
- `status`
- `categoria`
- `despesa`
- `data`
- `mes`
- `valor`
- `pagador`
- `recebedor`
- `tipo`
- `arquivo`
- `hash`
- `periodicidade`

O hash e o principal controle contra duplicidade.

## Dados Por Categoria

Cada pasta de categoria pode conter:

- comprovantes originais;
- `dados_extraidos.json`;
- `resumo_extraidos.xlsx`.

Quando a extracao da categoria roda, esses arquivos sao criados ou atualizados automaticamente.
