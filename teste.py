import src.report_class as rc

dados = rc.turma(planilha='Notas.xls', sr=5, disciplina='LOM3226', ano=2023, semestre=1)
dados.create_cabecalho('Resultado Final')
dados.create_table('P1', 'P2', 'NF')
dados.create_latex_report(['P1', 'P2'], runlatex=True, args=['papersize={220mm, 200mm}'])
