# Correcoes aplicadas em 14/09/2026

Escopo: somente a versao preenchida em
`C:\Users\User\Documents\Projetos\Daisy\system_proof`.
A planilha foi salva diretamente, sem criar copia. O programa oficial nao foi alterado.

## Resultado

- Cadastro tem botao Novo cadastro; Carregar/Salvar edita sem duplicar nem apagar valores.
- Identificadores estaveis vinculam Cadastros, Despesas e Controle.
- Nomes anteriores sao preservados na aba tecnica Aliases, inclusive telefonia.
- Sincronizar preserva periodicidade e situacao; despesas inativas nao aceitam novas importacoes.
- Novas despesas ficam visiveis; o cadastro e os totais podem crescer alem da linha 120.
- Resumo inclui todas as categorias existentes, inclusive Mesada, sem limite de dez.
- Removido o conjunto de totais redundantes da linha 121; total atual em 122.
- Importacoes validam valor, competencia, ano, cadastro ativo e hash antes de gravar.
- Duplicatas identificadas por hash nao somam nem subtraem valores automaticamente.
- Edicao e exclusao usam vinculos explicitos, atualizam os saldos e bloqueiam registros pendentes.
- IDs repetidos foram diferenciados; IDs anteriores permanecem na coluna M de Controle.
- Observacoes anteriores de Despesas foram preservadas na coluna W.
- Competencias das linhas 447, 448 e 666 foram normalizadas para 2026-01.
- O extrator Python consulta o cadastro salvo para nomes, despesas ativas e periodicidade da interface.
- Novos extratores de Esporte e Mesada nao inventam pagador ou recebedor quando ausentes.
- A normalizacao de nomes deixou de modificar o texto original durante comparacoes.

## Valores preservados

Foram conferidos individualmente os valores de 667 registros e todas as entradas
mensais/eventuais de Despesas, antes e depois da gravacao.

| Item | Valor |
| --- | ---: |
| Controle | R$ 59.067,48 |
| Despesas | R$ 59.268,45 |
| Diferenca historica a conferir | R$ 200,97 |
| Valor subtraido por estas correcoes | R$ 0,00 |

585 lancamentos possuem vinculo e saldo consistentes. 82 foram marcados PENDENTE
em Controle e detalhados em Conciliacao. Os valores desses registros nao foram
alterados. Nao troque o status para OK sem conferir comprovante, despesa e saldo.

A diferenca de R$ 200,97 nao foi atribuida automaticamente a duplicatas. Ajustar
essa diferenca sem identificar a origem mudaria os dados sem fundamento.

## Duplicatas

A varredura inicial examinou 844 arquivos. Encontrou 62 pares de comprovantes
com conteudo identico e um par de arquivos vazios .gitkeep. Nenhum foi excluido.

As linhas 669 e 671 de Controle usam o mesmo arquivo
`comprovantes\Moradia\condominio-fevereiro.pdf`, mas com valores diferentes:
R$ 792,23 e R$ 136,49. Isso exige conferencia: pode ser discriminacao de itens
do documento, nao necessariamente pagamento repetido.

Lista integral com caminhos e hashes:
`C:\Users\User\Desktop\Daisy_repo\_auditoria_system_proof_20260914\DUPLICADOS.md`.

## Verificacao

24 testes executados no Excel 2010 instalado, usando planilha vazia em memoria,
sem salvar arquivo de teste. Foram verificados cadastro, renomeacao, nomes
anteriores, sincronizacao, crescimento, resumo, importacao, duplicatas, edicao,
mudanca de competencia, exclusao e bloqueio de ambiguidades.

O arquivo salvo foi relido: 57 cadastros unicos, 667 IDs de controle unicos,
nenhuma celula de erro Excel, macros embutidas correspondentes aos fontes,
Mesada presente no resumo e integracao Python com telefonia validada.

Evidencias tecnicas e scripts:
`C:\Users\User\Desktop\Daisy_repo\_auditoria_system_proof_20260914`.

Nao execute os antigos scripts de reconstruir/limpar/importar modulo como parte
desta atualizacao: eles sao ferramentas legadas de reconstrucao, nao a migracao
validada, e podem apagar dados ou remover o novo modulo de integridade.
