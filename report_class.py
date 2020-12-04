# -*- coding: utf-8 -*-

import pandas 
import math
import sys
import os
import re
import numpy as np
import scipy.stats as st
import matplotlib
import matplotlib.pyplot as plt

class turma:
    '''
    Cria um relatório de notas com base numa planilha eletrônica (e.g., excel)

        Argumentos:
            * planilha (obrigatório): string - nome da planilha com as notas. Necessariamente deve conter as colunas 'Código' (com os NUSP) e 'Nome'
            * aba: int/string - aba com as notas (padrão: 0, ou seja, a primeira aba)
            * sr (skiprows): int - linhas a serem ignoradas no começo da planilha
            * print_nomes: bool - imprimir ou não o nome dos alunos (padrão: False)
            * disciplina: string
            * ano: int
            * semestre: int
    '''
    def __init__(self, planilha, aba=4, sr=0, print_nomes=False, disciplina='LOM3260', ano=2021, semestre=1):

        turma = pandas.read_excel(planilha, sheet_name=aba, skiprows=sr)

        self.print_nomes = print_nomes
        
        if self.print_nomes :
            self.turma = turma.sort_values(by=['Nome']) 
        else:
            self.turma = turma.sort_values(by=[u'Código']) 

        self.turma.reset_index(inplace=True)

        self.nusp = self.turma["Código"]
        self.nomes = self.turma['Nome']
        self.disciplina = disciplina
        self.ano = ano
        self.semestre = semestre
        
        self.histogram = {}
    
    def create_table(self, *args):
        '''
            create_table: produz uma tabela com as colunas da planilha fornecidas como argumento

            Argumentos:
                *args: strings - nomes das colunas da planilha que se deseja reproduzir
        '''

        N = len(args) 
        if  N: notas = []
        for i in range(N):
            notas.append(self.read_notas(P = str(args[i])))
        
        self.alunos = ''
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
                    pri = '\\textcolor{%s}{%.1f}' % (color, prova[i])
                pri = pri.replace('.', ',')
                aluno += ' & ' + pri 
            aluno += ' \\\\ \n'
            if sys.version_info[0] > (2):
                self.alunos += aluno
            else:
                self.alunos += aluno.encode('utf-8')

        self.tabela_notas = ''
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
            
    def create_histogram(self, prova):
        '''
            cria um histogram com a prova escolhida
            
            Argumentos:
                * prova: string - coluna escolhida para o histograma.
        '''
        self.histogram[prova] = histogram_class(self, prova)
        
    def create_cabecalho(self, titulo='Resultado da Prova'):
        '''
            create_cabecalho: formata o cabeçalho do report

                Argumento:
                    * titulo: string
        '''
        self.cabecalho = ''
        self.cabecalho += '\\begin{center}\n'
        self.cabecalho += '{\\LARGE \\bfseries ' + str(self.disciplina) + ' --- ' + str(self.semestre) + '\\textordmasculine{} semestre de ' + str(self.ano) + '}\\\\[4mm] \n'
        self.cabecalho += '{\\Large \\bfseries ' + titulo + '} \n'
        self.cabecalho += '\\end{center} \n'

    def create_latex_report(self, prova, papersize='a4paper', landscape=True, runlatex=False):
        '''
            create_latex_report: cria o relatório em pdf, juntando tabela + histograma. Exije o uso prévio das rotinas acima

            Argumentos:
                * prova: string - coluna escolhida para o report. Deve ser uma coluna previamente usada pelo comando create_histogram
                * papersize: string - um dos formatos de papersize aceitos pelo pacote geometry (default: a4paper).
                * landscape: bool - se a orientação do papel é landscape (default: True)
                * runlatex: bool - se deve rodar automaticamente o LaTeX (default: False)
        '''
        
        texfile = 'report-' + prova + '.tex'
        pattern = re.compile(r'\s+')
        texfile = re.sub(pattern, '-', texfile)
        
        with open(texfile, 'w') as f:
            f.write('\\documentclass[12pt]{article}\n\n')
            #f.write('\\usepackage[bitstream-charter]{mathdesign}\n')
            f.write('\\usepackage{mathpazo}\n')
            if landscape:
                f.write('\\usepackage[%s, landscape, margin=10mm]{geometry}\n' % papersize)
            else:
                f.write('\\usepackage[%s, margin=10mm]{geometry}\n' % papersize)
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
            f.write('\\begin{multicols}{2}\n')
            f.write('\\centering \n')
            
            f.write(self.tabela_notas)
            f.write('\n')
            
            figfile = 'hist-' + prova + '.pdf'
            figfile = figfile.replace(" ", "_")
            f.write('\\includegraphics[width=\columnwidth]{' + figfile + '}\n')
            f.write('\\begin{flushright}\n')
            f.write('Azuis: %d\n\n' % (self.histogram[prova].azuis))
            f.write('Vermelhas: %d \n' % (self.histogram[prova].vermelhas))
            f.write('\\end{flushright}\n')
            f.write('\\end{multicols}\n\n')

            f.write('\\end{document}\n')
            
        if runlatex:
            os.system('pdflatex ' + texfile)
    
    # Funções auxiliares

    def read_notas(self, P='P1'):
        '''
            Função para ler uma coluna específica
            é usada apenas internamente
        '''

        notas = self.turma[P]
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
            create_histogram: cria um pdf com um histograma da coluna da turma fornecida

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
        n, bins, patches = plt.hist(p1, bins, rwidth=.8, histtype='bar', zorder=2)
        self.vermelhas = sum(n[:5])
        self.azuis = sum(n)- self.vermelhas
        N, m, s = p1.size, np.mean(p1), np.std(p1)
        info_hist = 'Alunos: %d \nMédia: %.1f $\pm$ %.1f' % (N, m, s)

        z = st.norm(loc = m, scale = s)
        x = np.linspace( 0, 10, 100 )
        n_dist = z.pdf(x) * N
        plt.plot(x, n_dist, zorder = 2, label=info_hist)
        plt.legend(loc=0, framealpha=1.)
        ym = max( max(n_dist), max(n))
        ym = float(ym)
        ym = np.ceil(ym) + .5

        plt.xlim(0, 10)
        plt.xticks(range(0,11))
        plt.yticks(range(0, int(ym+1)))
        plt.ylim( 0, ym )
        plt.grid( axis='y', zorder=1 )
        plt.xlabel('Nota')
        plt.ylabel('Número de alunos')

        figfile = 'hist-' + prova + '.pdf'
        figfile = figfile.replace(" ", "_")
        plt.savefig(figfile,bbox_inches='tight')
