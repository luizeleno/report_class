import pandas 
import math
import sys
import os
import re
import numpy as np
import scipy.stats as st
import matplotlib
import matplotlib.pyplot as plt

'''
Classe em python3 para gerar relatório de notas de alguma turma

Otimizado para cursos da USP usando a "Lista de Apoio ao Docente" do Sistema Júpiter**

---

Exemplo de uso

Formato da planilha: as notas da disciplina _LOM3260_, do 2º semestre de 2020, estão na primeira aba de uma planilha chamada `Notas.xlsx`, que tem as colunas `P1`, `P2`, `NF` e `Freq`. As cinco primeiras linhas da planilha devem ser ignoradas. Necessariamente, a planilha deve conter também as colunas `Código` e `Nome` (esses nomes de coluna são configuráveis).

Você quer criar um relatório contendo uma tabela com todas as colunas, sem mostrar os nomes dos alunos, apenas seus códigos (NUSP). Você quer também um histograma da coluna `NF`. Use então o código abaixo:

import report_class

turma = report_class.turma('Notas.xlsx', sr=5, print_nomes=False, aba=0,
                            disciplina='LOM3260', semestre=2, ano=2020)

turma.create_table('P1', 'P2', 'NF', 'Freq')
turma.create_cabecalho('Resultado final')

turma.create_latex_report(['NF'], landscape=False)

O resultado será um arquivo $\LaTeX$ chamado `report-NF.tex`, que deve ser processado manualmente para gerar o relatório final (se quiser rodar o $\LaTeX$ automaticamente para gerar um arquivo pdf, use a opção `runlatex=True` na última linha).
'''

class turma:
    '''
    Cria um relatório de notas com base numa planilha eletrônica (e.g., excel)

        Argumentos:
            * planilha (obrigatório): string - nome da planilha com as notas. Necessariamente deve conter as colunas 'Código' (com os NUSP) e 'Nome'
                ** (esses nomes são configuráveis com as variáveis opcionais codigo e nome)
            * aba: int/string - aba com as notas (padrão: 0, ou seja, a primeira aba)
            * sr (skiprows): int - linhas a serem ignoradas no começo da planilha
            * codigo: nome da coluna da planilha com os IDs (por exemplo, NUSP) dos alunos
            * print_nomes: bool - imprimir ou não o nome dos alunos (padrão: False)
            * disciplina: string
            * ano: int
            * semestre: int
    '''
    def __init__(self, planilha, aba=0, sr=0, print_nomes=False, codigo='Código', nome='Nome', disciplina='LOM3260', ano=2021, semestre=1):

        turma = pandas.read_excel(planilha, sheet_name=aba, skiprows=sr)

        self.print_nomes = print_nomes
        
        if self.print_nomes :
            self.turma = turma.sort_values(by=[nome]) 
        else:
            self.turma = turma.sort_values(by=[codigo]) 

        self.turma.reset_index(inplace=True)

        self.nusp = self.turma[codigo]
        self.nomes = self.turma[nome]
        self.disciplina = disciplina
        self.ano = ano
        self.semestre = semestre
        
        self.histogram = {}
        self.cabecalho = ''
        self.alunos = ''
        self.tabela_notas = ''
    
    def create_table(self, *args):
        '''
            create_table: produz uma tabela com as colunas da planilha fornecidas como argumento

            Argumentos:
                *args: strings - nomes das colunas da planilha que se deseja reproduzir
                Obs: 'Freq' é uma coluna especial para as frequências. A função read_notas() tratará essa coluna diferentemente.
        '''

        N = len(args) 
        if  N: notas = []
        for prova in args:
            notas.append(self.read_notas(P = str(prova)))
        
        for i in range(self.nusp.size):
            
            #print(self.nomes[i])
            
            aluno = str(self.nusp[i]) 
            if self.print_nomes:  
                aluno += ' & ' + self.nomes[i]
            for prova in notas:
                if math.isnan(prova[i]):
                    pri = ' '
                else:
                    if prova[i] >=5:
                        color = 'blue'
                    else:
                        color = 'red'
                    pri = f'\\textcolor{{{color}}}{{{prova[i]}}}'
                pri = pri.replace('.', ',')
                aluno += ' & ' + pri 
            aluno += ' \\\\ \n'
            if sys.version_info[0] > (2):
                self.alunos += aluno
            else:
                self.alunos += aluno.encode('utf-8')

        self.tabela_notas += '\\rowcolors{2}{gray!25}{white} \n'
        if self.print_nomes :
            columns = 'll' +  'c' * N
        else :
            columns = 'l' +  'c' * N
        self.tabela_notas += '\\begin{tabular}{' + columns + '} \n'
        self.tabela_notas += '\\hline \n'
        self.tabela_notas += '\\rowcolor{gray!50}'
        provas_nomes = "} & \\textbf{".join(str(x) for x in args)
        header = '\\textbf{NUSP}' 
        if self.print_nomes:  
            header += ' & \\textbf{Nome}' 
        header += ' & \\textbf{' +  provas_nomes + '} \\\\ \n'
        self.tabela_notas += header
        self.tabela_notas += '\\hline \n'
        self.tabela_notas += self.alunos + '\n'
        self.tabela_notas += '\\hline \n'
        self.tabela_notas += '\\end{tabular} \n'

    def create_cabecalho(self, titulo='Resultados de avaliações'):
        '''
            create_cabecalho: formata o cabeçalho do report

                Argumento:
                    * titulo: string
        '''

        self.cabecalho = '\\begin{center}\n'
        self.cabecalho += '{\\LARGE \\bfseries ' + str(self.disciplina) + ' --- ' + str(self.semestre) + '\\textordmasculine{} semestre de ' + str(self.ano) + '}\\\\[4mm] \n'
        self.cabecalho += '{\\Large \\bfseries ' + titulo + '} \n'
        self.cabecalho += '\\end{center} \n'
        
    def create_latex_report(self, provas, filetitle_append='', papersize='a4paper', landscape=True, runlatex=False, cols=2, args=[]):
        '''
            create_latex_report: cria o relatório em pdf, juntando tabela + histograma.
            Exige o uso prévio de create_table()

            Argumentos:
                * provas: list of strings - colunas escolhidas para o report. Devem ser uma ou mais dentre as colunas fornecidas à create_table()
                * papersize: string - um dos formatos de papersize aceitos pelo pacote geometry (default: a4paper).
                * landscape: bool - se a orientação do papel é landscape (default: True)
                * runlatex: bool - se deve rodar automaticamente o LaTeX (default: False)
                * cols: número de colunas do relatório (para o pacote multicol)
                * args: lista de argumentos opcionais passados ao pacote LaTeX geometry (default: [])
        '''
        
        texfile = 'report-' + self.disciplina + '-' + filetitle_append + '.tex'
        pattern = re.compile(r'\s+')
        texfile = re.sub(pattern, '-', texfile)
        
        with open(texfile, 'w') as f:
            f.write('\\documentclass[12pt]{article}\n\n')
            #f.write('\\usepackage[bitstream-charter]{mathdesign}\n')
            f.write('\\usepackage{mathpazo}\n')
            if landscape:
                f.write(f'\\usepackage[{papersize}, landscape, margin=10mm]{{geometry}}\n')
            else:
                f.write(f'\\usepackage[{papersize}, margin=10mm]{{geometry}}\n')
            if len(args):
                f.write('\\geometry{')
                for arg in args:
                    f.write(f'{arg}')
                    f.write(',')
                f.write('}\n')
                    
            f.write('\\usepackage[utf8]{inputenc}\n')
            f.write('\\usepackage[table]{xcolor}\n')
            f.write('\\usepackage{icomma}\n')
            f.write('\\usepackage{graphicx}\n')
            f.write('\\usepackage{multicol}\n')
            f.write('\\usepackage{textcomp}\n\n')

            f.write('\\begin{document}\n\n')

            f.write('\\thispagestyle{empty}\n')
            f.write('\\centering\n')
            
            f.write(self.cabecalho)
            
            f.write('\\vfill\n')
            f.write(f'\\begin{{multicols}}{{{cols}}}\n')
            f.write('\\centering \n')
            
            f.write(self.tabela_notas)
            f.write('\n')
            
            for prova in provas:
                self.histogram[prova] = histogram_class(self, prova)
                figfile = 'hist-' + prova + '.pdf'
                figfile = figfile.replace(" ", "_")
                f.write('\\includegraphics[width=\columnwidth]{' + figfile + '}\n')
                f.write('\\begin{minipage}{\\columnwidth}\n')
                f.write('\\begin{flushright}\n')
                f.write(f'Azuis: {self.histogram[prova].azuis:.0f}\n\n')
                f.write(f'Vermelhas: {self.histogram[prova].vermelhas:.0f} \n')
                f.write('\\end{flushright}\n')
                f.write('\\vskip 5mm\n')                
                f.write('\\end{minipage}\n')
            
            f.write('\\end{multicols}\n\n')
            f.write('\\end{document}\n')
            
        if runlatex:
            os.system('latexmk -pdf ' + texfile)
    
    # Funções auxiliares

    def read_notas(self, P='P1'):
        '''
            Função para ler uma coluna específica
            é usada apenas internamente
        '''
        notas = self.turma[P]

        if P in ('Freq', 'Freq.', 'freq', 'freq.'):
            print('Aviso: Coluna de frequências identificada')
            notas *= 100
            notas = np.around(notas, decimals=0)
            notas = np.array(notas, dtype=int)
        else:
            notas = self.arredonda_notas(notas)

        return notas

    def filtra_notas(self, notas):
        '''
            filtro para identificar células vazias ou não numéricas
            usada apenas internamente
        '''
        notas = filter(lambda x: math.isnan(x) is False, notas)
        if sys.version_info[0] > 2 :
            ''' in python 3, filter returns a filter object
                and requires list() to get an array
                not necessary in python 2 '''
            notas = list(notas)
        return notas

    def arredonda_notas(self, notas):
        '''
            arredonda as notas para uma casa decimal
            usada apenas internamente
        '''
        return np.around(notas, decimals=1)


class histogram_class:
    
    def __init__(self, turma, prova):
        self.histogram(turma, prova)

    def histogram(self, turma, prova):
        '''
            histogram: cria um pdf com um histograma da coluna da turma fornecida

            Argumentos:
                * table: class - turma (nome da classe turma)
                * prova: string - coluna para o histograma
                
            classe usada apenas internamente
        '''

        matplotlib.rcParams['font.family'] = 'STIXGeneral'
        matplotlib.rcParams['font.size'] = 14
        #hfont = {'fontname': }

        plt.figure()
        ax = plt.axes()

        p1 = turma.read_notas(prova)
        p1 = turma.filtra_notas(p1)
        p1 = turma.arredonda_notas(p1)

        bins = range(11)
        n, bins, patches = ax.hist(p1, bins, rwidth=.8, histtype='bar', zorder=2, align='right')
        self.vermelhas = sum(n[:5])
        self.azuis = sum(n)- self.vermelhas
        N, m, s = p1.size, np.mean(p1), np.std(p1)
        info_hist = f'Estudantes: {N}\nMédia: {m:.1f} $\pm$ {s:.1f}'

        z = st.norm(loc = m, scale = s)
        x = np.linspace( 0, 10, 100 )
        n_dist = z.pdf(x) * N
        ax.plot(x, n_dist, zorder = 2, label=info_hist)
        plt.legend(loc=0, framealpha=1.)
        ym = max( max(n_dist), max(n))
        ym = float(ym)
        ym = np.ceil(ym) + .5

        ax.set_xlim(0.5, 10.5)
        ax.set_xticks(range(1,11))
        ax.set_xticklabels([f'[{i}-{i+1}{"]" if i==9 else ")"}' for i in range(10)], rotation=45)
        ax.set_yticks(range(0, int(ym+1)))
        ax.set_ylim( 0, ym )
        plt.grid( axis='y', zorder=1 )
        ax.set_title(prova)
        ax.set_xlabel('Nota')
        ax.set_ylabel('Número de estudantes')

        figfile = 'hist-' + prova + '.pdf'
        figfile = figfile.replace(" ", "_")
        plt.savefig(figfile,bbox_inches='tight')
