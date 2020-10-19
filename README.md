## class turma()

Cria um relatório de notas com base numa planilha eletrônica (e.g., excel)

Prof. Luiz T. F. Eleno
Escola de Engenharia de Lorena da USP
[sites.usp.br/computeel](https://sites.usp.br/computeel)

* Argumentos de inicialização:
  * `planilha`: string - planilha com as notas. Necessariamente deve conter as colunas 'Código' (com os NUSP) e 'Nome'
  * `aba`: int/string - aba com as notas (padrão: 0, ou seja, a primeira aba)
  * `sr` (skiprows): int - linhas a serem ignoradas no começo da planilha
  * `print_nomes`: bool - imprimir ou não o nome dos alunos (padrão: `False`)
  * `disciplina`: string
  * `ano`: int
  * `semestre`: int

### Rotinas para gera o conteúdo do relatório:

#### create_histogram: cria um pdf com um histograma da coluna fornecidas

* Argumentos:
  * `prova`: string - coluna usada para produzir o histograma
  * `ymax`: float - máximo valor da ordenada do gráfico
  * `create_table`: produz uma tabela com as colunas da planilha fornecidas como argumento

#### create_table: produz uma tabela com as colunas da planilha fornecidas como argumento

* Argumentos:
  * `strings` - nomes das colunas da planilha que se deseja reproduzir

### Rotinas para formatar e gerar o relatório:

#### create_cabecalho: formata o cabeçalho do relatório 

* Argumento:
  * `titulo`: string
        
#### create_latex_report: cria o relatório em pdf, juntando tabela + histograma. Exije o uso prévio das rotinas acima

* Argumentos:
  * `prova`: string - coluna escolhida para o relatório. Deve ser uma coluna previamente usada pelo comando create_histogram
  * `papersize`: string - um dos formatos de papersize aceitos pelo pacote geometry (default: a4paper).
                
### Exemplo de aplicação:
    
```python
import report_class

turma = report_class.turma('Notas.xls', aba=1, disciplina='LOM3213', sr=3)
turma.create_table('Nota P2', 'Trabalho', 'Lista 2', 'Lista 4', 'Bonus', 'Prova 2') 
turma.create_histogram('Prova 2', ymax=10.5)
turma.create_cabecalho('Resultado da P2')
turma.create_latex_report('Prova 2', papersize='a3paper')
```
