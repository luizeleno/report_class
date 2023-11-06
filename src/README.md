# `report_class`

## Classe em python3 para gerar relatório de notas de alguma turma

**Otimizado para cursos da USP usando a "Lista de Apoio ao Docente" do Sistema Júpiter**

---

### Exemplo de uso

Formato da planilha: as notas da disciplina *LOM3260*, do 2º semestre de 2020, estão na primeira aba de uma planilha chamada `Notas.xlsx`, que tem as colunas `P1`, `P2`, `NF` e `Freq`. As cinco primeiras linhas da planilha devem ser ignoradas. Necessariamente, a planilha deve conter também as colunas `Código` e `Nome`.

Você quer criar um relatório contendo uma tabela com todas as colunas, sem mostrar os nomes dos alunos, apenas seus códigos (NUSP). Você quer também um histograma da coluna `NF`. Use então o código abaixo:

```python
import report_class

turma = report_class.turma('Notas.xlsx', sr=5, print_nomes=False, aba=0,
                            disciplina='LOM3260', semestre=2, ano=2020)

turma.create_table('P1', 'P2', 'NF', 'Freq')
turma.create_cabecalho('Resultado final')

turma.create_latex_report(['NF'], landscape=False)
```

O resultado será um arquivo $\LaTeX$ chamado `report-NF.tex`, que deve ser processado manualmente para gerar o relatório final (se quiser rodar o $\LaTeX$ automaticamente para gerar um arquivo pdf, use a opção `runlatex=True` na última linha).
