# Plano de Desenvolvimento — Planilha Automatizada de Comprovantes (Excel + VBA)

## Objetivo

Refazer a planilha estratégica usando **VBA**, de forma que ela funcione como um
pequeno aplicativo: extrair os dados dos comprovantes, inserir na planilha,
calcular a média mensal correta (contando só os meses preenchidos), calcular a
cota-parte (rateio entre genitores), preencher automaticamente todos os campos
e criar links para abrir cada comprovante na hora da conferência.

## 1. Arquitetura

```
[Comprovantes PNG/PDF]
        │  (Python: OCR + parsing)          ← já temos (extrair_categoria.py)
        ▼
[JSON padronizado  comprovantes.json]
        │  (VBA: macro Importar)
        ▼
[Planilha Excel + VBA]
        │  (macros de cálculo, links, preenchimento)
        ▼
[Planilha final preenchida + links clicáveis]
```

Divisão de responsabilidades:
- **Python** → extrai e padroniza (fora do Excel, é o que já foi feito).
- **VBA** → importa, calcula, preenche, gera links (dentro do Excel).

## 2. Formato de dados padronizado (JSON oficial na implementação atual)

Status de implementação: a ferramenta Python gera `comprovantes.json` como
fonte oficial para alimentar a planilha principal. O arquivo `comprovantes.csv`
continua sendo gerado apenas por compatibilidade com macros/rotinas antigas.

Pipeline implementado:

```bash
python3 _scripts/integrar_extracao_planilha.py
```

Esse comando:
1. consolida os `dados_extraidos.json` por categoria;
2. gera `comprovantes.json`;
3. marca cada registro como `status: "pendente"`;
4. gera `comprovantes.csv` compatível.

Importante: a extração Python **não preenche** `Despesas` nem `Controle`. O
preenchimento é comandado no Excel:

1. `CarregarFilaJson` lê `comprovantes.json` e preenche apenas a aba **Fila**.
2. O usuário revisa/edita os registros pendentes na aba **Fila**.
3. `ImportarSelecionadoDaFila` ou `ImportarTodosPendentes` grava na planilha
   principal e cria os links clicáveis na aba **Controle**.

CSV compatível, com `;` como separador:

```csv
id;categoria;despesa;data;mes;valor;pagador;recebedor;tipo;arquivo;hash;periodicidade
```

O campo `periodicidade` aceita `Mensal`, `Anual` e `Eventual`.

### Formato JSON padronizado

Um único arquivo `comprovantes.json` com todos os comprovantes de todas as pastas:

```json
{
  "comprovantes": [
    {
      "id": 1,
      "categoria": "Alimentação",
      "despesa": "Supermercado e alimentação domiciliar",
      "data": "2026-05-08",
      "mes": "2026-05",
      "valor": 69.24,
      "pagador": "Daisy Ferreira Braga de Souza",
      "recebedor": "iFood",
      "tipo": "Pix",
      "arquivo": "Alimentação-.../Comprovante_....png",
      "hash": "abc123..."
    }
  ]
}
```

O campo `despesa` (já classificado) é gerado pelo Python e consumido pelo VBA.

## 3. Estrutura da planilha (abas)

1. **Despesas** — despesas por categoria; colunas de mês dinâmicas; média mensal.
2. **Controle** — inventário de comprovantes (1 linha por comprovante) + link.
3. **Resumo e Rateio** — totais por categoria + cota-parte dos genitores.
4. **Config** — parâmetros (renda líquida, obrigações, nº de moradores, etc.).
5. **Import** — botões (Extrair / Importar / Limpar) e status da importação.

## 4. Resolvendo a média mensal (sem média fixa)

**Problema:** a planilha atual usa `=Total/9 + anual/12`, mas nem toda despesa
tem os 9 meses preenchidos (algumas têm 4, 2 ou 1 mês). A média precisa contar
apenas os meses com valor.

**Solução 1 — fórmula nativa (sem VBA):**
`=AVERAGE(C5:K5)` já ignora células vazias. Ou `=SUM(C5:K5)/COUNT(C5:K5)`.

**Solução 2 — UDF VBA (mais controle, recomendada):**

```vba
Function MediaMeses(rng As Range) As Double
    Dim c As Range, soma As Double, n As Long
    soma = 0: n = 0
    For Each c In rng
        If IsNumeric(c.Value) And c.Value > 0 Then
            soma = soma + c.Value
            n = n + 1
        End If
    Next c
    MediaMeses = IIf(n > 0, soma / n, 0)
End Function
```

Uso na célula de média mensal: `=MediaMeses(C5:K5) + M5/12`
(média dos meses preenchidos + parcela anual/12).

Vantagem da UDF: dá para contar "meses preenchidos" com critério configurável
(ex.: ignorar valores <= 0, ignorar meses específicos).

## 5. Cota-parte (rateio entre genitores)

Campos em **Config** (entrada manual do usuário):
- Genitor 1: renda líquida mensal, outras obrigações essenciais.
- Genitor 2: renda líquida mensal, outras obrigações essenciais.

Fórmulas em **Resumo e Rateio**:
```
Renda disponível(G1) = renda_liquida(G1) - obrigacoes(G1)
Renda disponível(G2) = renda_liquida(G2) - obrigacoes(G2)
Percentual(G1)       = disponivel(G1) / (disponivel(G1) + disponivel(G2))
Percentual(G2)       = disponivel(G2) / (disponivel(G1) + disponivel(G2))
Cota(G1)             = TotalGeral * Percentual(G1)
Cota(G2)             = TotalGeral * Percentual(G2)
```

## 6. Links para conferência dos comprovantes

Na aba **Controle**, coluna "Arquivo", usar hyperlink nativo:

`=HYPERLINK("caminho\comprovante.png"; "abrir")`

Ou uma macro que abre o arquivo da célula selecionada:

```vba
Sub AbrirComprovante()
    Dim p As String
    p = ActiveCell.Value
    If p <> "" And Dir(p) <> "" Then
        ThisWorkbook.FollowHyperlink p
    Else
        MsgBox "Arquivo não encontrado: " & p, vbExclamation
    End If
End Sub
```

## 7. Importação automática

- **Botão "Extrair"** → chama o Python (via `Shell`) para rodar a extração e
  gerar o `comprovantes.json`.
- **Botão "Importar"** → lê o `comprovantes.json` (via Power Query ou macro
  com JSON parser) e preenche as abas Despesas e Controle.
- **Botão "Limpar"** → limpa os dados importados (para reprocessar).

Exemplo de chamada do Python pelo VBA:
```vba
Sub ExtrairDados()
    Shell "python C:\...\extrair_tudo.py", vbNormalFocus
End Sub
```

## 8. Fases de desenvolvimento

- **Fase 1 — Estrutura:** criar a planilha com as abas, cabeçalhos e colunas de
  mês dinâmicas (que crescem conforme aparecem novos meses).
- **Fase 2 — Fórmulas/UDF:** média mensal dinâmica, cota-parte, totais.
- **Fase 3 — Importação:** VBA lê o JSON e preenche Despesas + Controle.
- **Fase 4 — Links:** hyperlinks clicáveis para os comprovantes.
- **Fase 5 — Integração:** botão que dispara a extração Python + importa.
- **Fase 6 — Validação:** comparar totais (JSON × planilha), testar duplicatas,
  testar meses vazios na média.

## 9. Critérios de aceite

1. Média mensal calculada só com os meses preenchidos (sem ÷ fixo).
2. Cota-parte recalculada automaticamente ao editar rendas em Config.
3. Importação do JSON preenche Despesas e Controle sem erros.
4. Cada comprovante tem link clicável que abre o arquivo.
5. Totais da planilha batem com os totais dos JSONs (validação).
6. Reprocessar (Limpar + Importar) não duplica linhas.

## 10. Interface de importação (UserForm)

A importação não é "cega": antes de gravar, o usuário revisa cada comprovante.

- **ListBox** com todos os comprovantes importados (id, data, valor, recebedor).
- **Campos editáveis** (TextBoxes): data, valor, pagador, recebedor, tipo.
- **ComboBox de categoria** + **ComboBox de despesa** — para o usuário corrigir a
  classificação quando o extrator não acertou (ou não classificou).
- **ComboBox de periodicidade** (Mensal / Anual / Eventual).
- **Botão "Inserir"** grava o comprovante na planilha e avança; **"Pular"** pula.

Arquivo: `userform_importacao.bas` (código pronto; montar o UserForm no editor VBA).

## 11. Valores esporádicos / eventuais / anuais

Há uma coluna "Valor anual/eventual (R$)" (coluna M na aba Despesas). A regra:

- Valor **mensal** → vai para a coluna do mês correspondente.
- Valor **anual/eventual** → vai para a coluna anual (M) e entra na média como
  `M/12` (12 meses do ano), ou seja, vira o valor mensal equivalente.
- A média mensal final = `média dos meses preenchidos + anual/12`.

Assim, um IPTU anual de R$ 760,15 contribui com R$ 63,35/mês na média e na
cota-parte. A periodicidade vem do CSV (`periodicidade`) e é editável na
interface (ComboBox Mensal/Anual/Eventual).

## 12. Armazenamento dos comprovantes

Opções avaliadas:

| Opção | Uso recomendado |
|---|---|
| Arquivos + JSON/CSV | Fase inicial (atual) |
| **SQLite** | **Recomendado** — SQL completo em 1 arquivo, sem servidor |
| PostgreSQL | Multiusuário / acesso remoto / web |

**Recomendação:** SQLite como fonte canônica (o Python grava no SQLite e exporta
o CSV que o Excel/VBA consome). PostgreSQL fica como evolução futura se surgir
acesso multiusuário.

Esquema SQLite proposto:

```sql
CREATE TABLE comprovantes (
    id INTEGER PRIMARY KEY,
    categoria TEXT,
    despesa TEXT,
    data TEXT,
    mes TEXT,
    valor REAL,
    pagador TEXT,
    recebedor TEXT,
    tipo TEXT,
    periodicidade TEXT,   -- 'Mensal', 'Anual', 'Eventual'
    arquivo TEXT,
    hash TEXT             -- para detectar duplicatas
);
```
