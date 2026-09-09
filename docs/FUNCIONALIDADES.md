# Funcionalidades Do Sistema

## Extracao De Comprovantes

- Extrair um unico arquivo (`png`, `jpg`, `jpeg`, `pdf`).
- Extrair todos os arquivos de uma pasta selecionada.
- Extrair todas as categorias configuradas.
- Ler texto embutido em PDFs digitais com `pdfplumber`.
- Converter PDFs escaneados em imagem com `pdftoppm`.
- Aplicar OCR em imagens com Tesseract (`por+eng`).
- Aplicar regras especificas por categoria.
- Preencher campos extraidos: data, mes, valor, pagador, recebedor, tipo, despesa e periodicidade.
- Permitir revisao manual antes de mandar para o Excel.
- Remover item da fila antes de salvar.
- Abrir/pre-visualizar comprovante na ferramenta grafica.

## Indices Por Categoria

- Criar `dados_extraidos.json` dentro da pasta da categoria.
- Criar `resumo_extraidos.xlsx` dentro da pasta da categoria.
- Atualizar indices ja existentes sem perder registros conhecidos.
- Detectar comprovante pelo hash MD5.
- Listar comprovantes novos ainda nao presentes no indice.
- Normalizar caminho relativo para manter portabilidade.

## Fila Consolidada

- Gerar `comprovantes.json` para importacao no Excel.
- Gerar `comprovantes.csv` como saida auxiliar.
- Consolidar registros vindos dos `dados_extraidos.json` por categoria.
- Gerar metadados de data/hora, base do projeto e total de comprovantes.
- Manter status inicial `pendente`.

## Planilha Excel

- Carregar fila JSON na aba `Fila`.
- Importar uma linha selecionada.
- Importar todos os pendentes.
- Evitar duplicidade por hash ja importado.
- Preencher aba `Controle` com historico detalhado.
- Preencher aba `Despesas` com valor no mes correto.
- Preencher despesa eventual/anual quando a periodicidade for anual.
- Atualizar status de comprovante como `Anexado`.
- Registrar numeros de controle em `Observacoes`.
- Abrir comprovante original a partir do registro.
- Filtrar registros do `Controle` a partir dos numeros em `Observacoes`.
- Excluir registro do `Controle` e abater valor da despesa.
- Limpar dados importados preservando estrutura.
- Bloquear/desbloquear edicao.
- Congelar referencias da aba `Despesas`.
- Atualizar formulas de totais.

## Cadastros

- Cadastrar nova categoria/despesa operacional.
- Definir se uma despesa entra com cota parte.
- Sincronizar cadastros com a aba `Despesas`.
- Criar despesa ausente quando necessario.
- Garantir categoria no resumo.

## Calculos Financeiros

- Valor mensal: `(soma dos meses + despesa eventual/anual) / 12`.
- Valor anual: `soma dos meses + despesa eventual/anual`.
- Cota parte anual: divide por 2 quando `Cota parte? = Sim`; caso contrario usa valor integral.
- Cota parte mes: divide por 2 quando `Cota parte? = Sim`; caso contrario usa valor integral.
- Resumo por categoria em `Resumo e Rateio`.
- Total filtrado na aba `Controle` via subtotal.

## Manutencao Administrativa

- Aplicar tema e layout por script VBS.
- Recriar/ajustar areas de trabalho das abas.
- Reimportar modulo VBA para a planilha.
- Limpar processos Excel invisiveis.
- Limpar planilha principal para testes.
- Versionar fonte VBA e scripts no Git.
