# Instalacao e Dependencias

Este documento descreve como preparar uma maquina para rodar o extrator e a planilha.

## Requisitos

- Windows com Microsoft Excel habilitado para macros (`.xlsm`).
- Python 3.10 ou superior recomendado.
- Tesseract OCR instalado e acessivel pelo `PATH`.
- Idiomas `por` e `eng` disponiveis no Tesseract.
- Poppler instalado e `pdftoppm` acessivel pelo `PATH`.
- Bibliotecas Python listadas em `requirements.txt`.

## Instalar Python

Baixe o Python em `https://www.python.org/downloads/windows/` e marque:

- `Add Python to PATH`
- `pip`
- `tcl/tk and IDLE`, para habilitar `tkinter`

Teste no Prompt de Comando:

```bat
python --version
py -3 --version
```

## Instalar Bibliotecas Python

Dentro da pasta do projeto:

```bat
python -m pip install -r requirements.txt
```

Ou use o verificador:

```bat
VerificarDependencias_Windows.bat
```

O verificador detecta bibliotecas faltantes e oferece instalar as dependencias Python automaticamente.

## Instalar Tesseract OCR

No Windows, use o instalador mantido em:

`https://github.com/UB-Mannheim/tesseract/wiki`

Durante a instalacao, inclua o idioma Portugues. Depois confirme:

```bat
tesseract --version
tesseract --list-langs
```

A lista deve conter `por` e, preferencialmente, `eng`.

## Instalar Poppler

O Poppler e usado para converter PDF escaneado em imagem antes do OCR.

1. Baixe uma versao Windows em `https://github.com/oschwartz10612/poppler-windows/releases`.
2. Extraia a pasta.
3. Adicione a subpasta `Library\bin` ao `PATH` do Windows.
4. Teste:

```bat
pdftoppm -h
```

## Descoberta de Bibliotecas Pendentes

A ferramenta oficial e:

```bat
VerificarDependencias_Windows.bat
```

Por baixo, ela chama:

```bat
python _scripts\verificar_deps.py
```

Para instalar automaticamente apenas as bibliotecas Python faltantes:

```bat
python _scripts\verificar_deps.py --instalar
```

Ela verifica:

- `openpyxl`: gera as planilhas `resumo_extraidos.xlsx` por categoria.
- `pdfplumber`: le texto embutido em PDFs digitais.
- `pillow`: exibe pre-visualizacao de imagens na ferramenta grafica.
- `tkinter`: interface grafica nativa do Python.
- `tesseract`: OCR de imagens.
- `pdftoppm`: conversao de PDF escaneado.

## Configuracao Do Excel

No Excel:

1. Abra `Planilha_Comprovantes_VBA.xlsm`.
2. Habilite macros quando o Excel perguntar.
3. Se necessario, em `Arquivo > Opcoes > Central de Confiabilidade`, habilite macros para a pasta confiavel.
4. Mantenha o arquivo dentro da pasta do projeto para preservar caminhos relativos.

## Portabilidade

O projeto detecta a pasta base em `_scripts/config.py`. Normalmente nao e preciso editar caminho nenhum.

Se for necessario forcar outro caminho, defina a variavel de ambiente `DAISY_BASE` apontando para a pasta raiz do projeto.
