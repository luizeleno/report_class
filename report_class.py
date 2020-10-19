import pandas 
import math
import sys
import numpy as np
import scipy.stats as st
import matplotlib
import matplotlib.pyplot as plt

class turma:
	'''
	
	class turma()
	
	Cria um relatório de notas com base numa planilha eletrônica (e.g., excel)

		Argumentos de inicialização:
			* planilha: string - planilha com as notas. Necessariamente deve conter as colunas 'Código' (com os NUSP) e 'Nome'
			* aba: int/string - aba com as notas (padrão: 0, ou seja, a primeira aba)
			* sr (skiprows): int - linhas a serem ignoradas no começo da planilha
			* print_nomes: bool - imprimir ou não o nome dos alunos (padrão: False)
			* disciplina: string
			* ano: int
			* semestre: int
		
	* Rotinas para gera o conteúdo do relatório:

		* create_histogram: cria um pdf com um histograma da coluna fornecidas
			
			Argumentos:
				* prova: string - coluna usada para produzir o histograma
				* ymax: float - máximo valor da ordenada do gráfico
		
		* create_table: produz uma tabela com as colunas da planilha fornecidas como argumento

			Argumentos:
				* strings - nomes das colunas da planulha que se deseja reproduzir

	* Rotinas para formatar e gerar o relatório:

		* create_cabecalho: formata o cabeçalho do report

				Argumento:
					* titulo: string
		
		* create_latex_report: cria o relatório em pdf, juntando tabela + histograma. Exije o uso prévio das rotinas acima

			Argumentos:
				* prova: string - coluna escolhida para o report. Deve ser uma coluna previamente usada pelo comando create_histogram
				* papersize: string - um dos formatos de papersize aceitos pelo pacote geometry (default: a4paper).
				
	* Exemplo de aplicação:
	
		import report_class

		turma = report_class.turma('Notas.xls', aba=1, disciplina='LOM3213', sr=3)
		turma.create_table('Nota P2', 'Trabalho', 'Lista 2', 'Lista 4', 'Bonus', 'Prova 2') 
		turma.create_histogram('Prova 2', ymax=10.5)
		turma.create_cabecalho('Resultado da P2')
		turma.create_latex_report('Prova 2', papersize='a3paper')

'''
	def __init__(self, planilha, aba=4, sr=0, print_nomes=False, disciplina='LOM3226', ano=2018, semestre=1):

		turma = pandas.read_excel(planilha, sheet_name=aba, skiprows=sr)

		self.print_nomes = print_nomes
		
		if self.print_nomes :
			self.turma = turma.sort_values(by=['Nome']) 
		else:
			self.turma = turma.sort_values(by=[u'Código']) 

		self.turma.reset_index(inplace=True)

		self.nusp = self.turma[u"Código"]
		self.nomes = self.turma['Nome']
		self.disciplina = disciplina
		self.ano = ano
		self.semestre = semestre

	def read_notas(self, P='P1'):

		notas = self.turma[P]
		notas = self.arredonda_notas(notas)
		return notas
	
	def create_table(self, *args):
		'''
			create_table: produz uma tabela com as colunas da planilha fornecidas como argumento

			Argumentos:
				* strings - nomes das colunas da planulha que se deseja reproduzir
		'''

		N = len(args) 
		if  N: notas = []
		for i in range(N):
			notas.append( self.read_notas( P = str( args[i] ) ) )
		
		f = open( 'alunos.tex', 'w')

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
			if sys.version_info[0] > (2) :
				f.write(aluno)
			else:
				f.write(aluno.encode('utf-8'))
		f.close()
		#
		f = open( 'tabela-notas.tex', 'w')
		f.write('\\rowcolors{2}{gray!25}{white} \n')
		if self.print_nomes :
			columns = 'll' +  'c' * N
		else :
			columns = 'l' +  'c' * N
		f.write('\\begin{tabular}{' + columns + '} \n')
		f.write('\\hline \n')
		f.write('\\rowcolor{gray!50}')
		provas_nomes = "} & \\textbf{".join(str(x) for x in args)
		header = '\\textbf{NUSP}' 
		if self.print_nomes:  
			header += ' & \\textbf{Nome}' 
		header += ' & \\textbf{' +  provas_nomes + '} \\\\ \n'
		f.write(header)
		f.write('\\hline \n')
		f.write('\\input{alunos} \n')
		f.write('\\hline \n')
		f.write('\\end{tabular} \n')
		f.close()
		
	def filtra_notas(self, notas):
		
		notas = filter(lambda x: math.isnan(x) is False, notas)
		if sys.version_info[0] > 2 :
			''' in python 3, filter returns a filter object
				and requires list() to get an array
				not necessary in python 2 '''
			notas = list(notas)
		return notas

	def arredonda_notas(self, notas):
		
			return np.around(notas, decimals=1)
	
	def create_histogram(self, prova):
		'''
			create_histogram: cria um pdf com um histograma da coluna fornecidas
			
			Argumentos:
				* prova: string - coluna usada para produzir o histograma
				* ymax: float - máximo valor da ordenada do gráfico
		'''
		
		matplotlib.rcParams['font.family'] = 'STIXGeneral'
		matplotlib.rcParams['font.size'] = 14
		#hfont = {'fontname': }
		
		plt.figure()
		ax = plt.axes()

		p1 = self.read_notas(prova)
		p1 = self.filtra_notas(p1)
		p1 = self.arredonda_notas(p1)

		bins = range( 11 )
		n, bins, patches = plt.hist( p1, bins, rwidth=.8, histtype='bar', zorder=2 )
		self.vermelhas = sum(n[:5])
		self.azuis = sum(n)- self.vermelhas
		N, m, s = p1.size, np.mean(p1), np.std(p1)
		info_hist = u'Alunos: %d \nMédia: %.1lf $\pm$ %.1lf' % ( N, m, s )
														   
		z = st.norm( loc = m, scale = s )
		x = np.linspace( 0, 10, 100 )
		n_dist = z.pdf( x ) * N
		plt.plot( x, n_dist, zorder = 2, label=info_hist )
		plt.legend(loc=0, framealpha=1.)
		ym = max( max(n_dist), max(n) )
		ym = float(ym)
		ym = np.ceil(ym) + .5
	
		plt.xlim( 0, 10 )
		plt.xticks( range(0,11) )
		plt.yticks( range(0, int(ym+1) ) )
		plt.ylim( 0, ym )
		plt.grid( axis='y', zorder=1 )
		plt.xlabel( 'Nota' )
		plt.ylabel( u'Número de alunos' )
		
		figfile = 'hist-' + prova + '.pdf'
		figfile = figfile.replace(" ", "_")
		plt.savefig(figfile,bbox_inches='tight')
		
	def create_latex_report(self, prova, papersize='a4paper'):
		'''
			create_latex_report: cria o relatório em pdf, juntando tabela + histograma. Exije o uso prévio das rotinas acima

			Argumentos:
				* prova: string - coluna escolhida para o report. Deve ser uma coluna previamente usada pelo comando create_histogram
				* papersize: string - um dos formatos de papersize aceitos pelo pacote geometry (default: a4paper).
		'''
		
		f = open('report-' + prova + '.tex', 'w')
		f.write('\\documentclass[12pt]{article}\n\n')
		#f.write('\\usepackage[bitstream-charter]{mathdesign}\n')
		f.write('\\usepackage{mathpazo}\n')
		f.write('\\usepackage[%s, landscape, margin=10mm]{geometry}\n' % papersize)
		f.write('\\usepackage[utf8]{inputenc}\n')
		f.write('\\usepackage[table]{xcolor}\n')
		f.write('\\usepackage{icomma}\n')
		f.write('\\usepackage{graphicx}\n')
		f.write('\\usepackage{multicol}\n')
		f.write('\\usepackage{textcomp}\n\n')

		f.write('\\begin{document}\n\n')
		
		f.write('\\thispagestyle{empty}\n')
		f.write('\\centering\n')
		f.write('\\input{cabecalho}\n')
		f.write('\\vfill\n')
		f.write('\\begin{multicols}{2}\n')
		f.write('\\centering \n')
		f.write('\\input{tabela-notas}\n')
		figfile = 'hist-' + prova + '.pdf'
		figfile = figfile.replace(" ", "_")
		f.write('\\includegraphics[width=\columnwidth]{' + figfile + '}\n')
		f.write('\\begin{flushright}\n')
		f.write('Azuis: %d\n\n' % (self.azuis))
		f.write('Vermelhas: %d \n' % (self.vermelhas))
		f.write('\\end{flushright}\n')
		f.write('\\end{multicols}\n\n')
		
		f.write('\\end{document}\n')
		f.close()
		
	def create_cabecalho(self, titulo='Resultado da Prova'):
		'''
			create_cabecalho: formata o cabeçalho do report

				Argumento:
					* titulo: string
		'''
		
		f = open('cabecalho.tex', 'w')
		f.write('\\begin{center}\n')
		f.write('{\\LARGE \\bfseries ' + str(self.disciplina) + ' --- ' + str(self.semestre) + '\\textordmasculine{} semestre de ' + str(self.ano) + '}\\\\[4mm] \n')
		f.write('{\\Large \\bfseries ' + titulo + '} \n')
		f.write('\\end{center} \n')
		f.close()
