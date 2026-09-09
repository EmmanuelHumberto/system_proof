# System Proof - Planilha de Comprovantes

Sistema local para extrair dados de comprovantes, revisar uma fila de importacao e alimentar uma planilha Excel com macros VBA.

O projeto foi feito para ser portavel: a pasta pode ser movida para outro diretorio e os scripts recalculam o caminho base automaticamente a partir da propria pasta do projeto.

## O Que Existe No Projeto

- `Planilha_Comprovantes_VBA.xlsm`: planilha principal com macros.
- `comprovantes/`: pastas de comprovantes separadas por categoria.
- `comprovantes.json`: fila consolidada lida pela aba `Fila` do Excel.
- `comprovantes.csv`: saida auxiliar em CSV.
- `_scripts/`: scripts Python de OCR, extracao, classificacao e geracao da fila.
- `_extraidos/`: area tecnica de dados brutos/temporarios da extracao global.
- `vba/`: fonte versionavel dos modulos VBA.
- `tools/`: scripts administrativos para ajustar layout, importar VBA e limpar processos Excel.
- `docs/`: documentacao tecnica e manuais.

## Comeco Rapido No Windows

1. Instale Python 3 e marque a opcao `Add Python to PATH`.
2. Instale o Tesseract OCR com idioma Portugues.
3. Instale o Poppler para Windows e deixe `pdftoppm.exe` no `PATH`.
4. Execute `VerificarDependencias_Windows.bat`.
5. Execute `ExtrairComprovantes_Windows.bat` para abrir a ferramenta grafica.
6. Abra `Planilha_Comprovantes_VBA.xlsm` no Excel.
7. Na aba `Import`, use `Fila JSON` para carregar a fila e depois importe pendentes ou a linha selecionada.

## Documentacao

- [Instalacao e dependencias](docs/INSTALACAO.md)
- [Arquitetura](docs/ARQUITETURA.md)
- [Manual da ferramenta de extracao](docs/FERRAMENTA_EXTRACAO.md)
- [Manual de uso](docs/MANUAL_USO.md)
- [Funcionalidades do sistema](docs/FUNCIONALIDADES.md)
- [Manutencao e solucao de problemas](docs/MANUTENCAO.md)

## Fluxo Principal

1. Colocar comprovantes nas pastas de categoria dentro de `comprovantes/`.
2. Rodar a extracao pela ferramenta grafica, por lote ou por todas as categorias.
3. Conferir/ajustar os dados antes da importacao.
4. Salvar a fila `comprovantes.json`.
5. Abrir o Excel e carregar a fila.
6. Importar os registros para `Controle` e `Despesas`.
7. Conferir os totais em `Resumo e Rateio`.

## Observacao Importante Sobre Duplicidade

A importacao usa o hash do arquivo para evitar duplicidade. Se o mesmo comprovante ja foi importado para `Controle`, ele nao deve ser importado novamente. Se o arquivo for alterado, o hash muda e ele passa a ser tratado como outro arquivo.