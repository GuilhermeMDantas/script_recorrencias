import xlrd, xlsxwriter		# Manage xls file
import unidecode			# Remover acentos
import re 					# Tratamento específico de string
import collections			# Criar nested dictionary (relação solicitante/divisao palavra)

def main():

	# Abre os sats
	excel_sats = xlrd.open_workbook(r'')

	# Abre o spreadsheet dos sats
	planilha_sats = excel_sats.sheet_by_index(0)

	# dc_ocorrencia
	col_ocorrencia = planilha_sats.col(8)

	# nm_solicitante
	col_solicitante = planilha_sats.col(1)

	# fk_cd_divisao
	col_divisao = planilha_sats.col(3)

	# Pegas todas as ocorrências e já as transforma em string
	ocorrencias = [str(ocorrencia) for ocorrencia in col_ocorrencia]
	del ocorrencias[0]  # Deleta o nome da coluna (dc_ocorrencia) da lista

	# Pega todos os solicitantes e já os transforma em string
	solicitantes = [str(solicitante) for solicitante in col_solicitante]
	del solicitantes[0]  # Deleta o nome da coluna (nm_solicitante) da lista

	# Pega todos as divisoes e já as transforma em string
	divisoes = [str(divisoes) for divisoes in col_divisao]
	del divisoes[0]  # Deleta o nome da coluna (fk_cd_divisao) da lista

	# Cria os arquivos que serão escritos as ocorrencias
	txt_ocorrencias		= open(r'total de ocorrencias2.txt', 		'w+')
	txt_solicitantes	= open(r'ocorrencia por solicitante2.txt', 	'w+')
	txt_divisao			= open(r'ocorrencia por divisao2.txt', 		'w+')


	#########################################
	#										#
	#	Faz o tratamento das ocorrencias 	#
	#										#
	#########################################
	
	# Limpa as ocorrencias de caracteres especiais e deixa tudo lowercase
	ocorrencias = limpa_ocorrencias(ocorrencias)

	# Dicionario de ocorrencias
	dic_ocorrencias = {}

	# Adiciona as palavras e suas ocorrencias ao dicionario
	dic_ocorrencias = relaciona_elementos(ocorrencias, dic_ocorrencias)
	#print('Ocorrencias Total:\n{' + '\n' + '\n'.join('\t{}:{}'.format(k, v) for k, v in dic_ocorrencias.items()) + '\n' + '}')

	# Escreve as ocorrências no formato CSV
	for palavra, ocorrencia in dic_ocorrencias.items():
		txt_ocorrencias.write('{};{};'.format(palavra, ocorrencia))


	#########################################
	#										#
	#	Faz o tratamento dos solicitantes 	#
	#										#
	#########################################

	# Limpa os nomes e deixa tudo lowercase
	solicitantes = limpa_solicitantes(solicitantes)

	# Cria o dicionario principal
	dic_solicitantes = collections.defaultdict(dict)

	# Adiciona os solicitantes como [key]s do dicionario
	dic_solicitantes = adiciona_elementos(solicitantes, dic_solicitantes)

	# Relaciona quem fez a ocorrencia e quais foram as palavras
	dic_solicitantes = relaciona_elementos(ocorrencias, dic_solicitantes, solicitantes)

	# Escreve as ocorrências por solicitante no formato CSV
	escreve_txt(txt_solicitantes, dic_solicitantes)		

	#####################################
	#									#
	#	Faz o tratamento das divisoes	#
	#									#
	#####################################
	
	# Remove texto dos numeros
	divisoes = limpa_divisoes(divisoes)

	# Criação do dicionario de divisoes
	dic_divisoes = collections.defaultdict(dict)

	# Juntas as recorrencias de uma mesma divisão numa só entrada
	dic_divisoes = adiciona_elementos(divisoes, dic_divisoes)

	# Relaciona as palavras com as divisões
	dic_divisoes = relaciona_elementos(ocorrencias, dic_divisoes, divisoes)

	# Escreve as ocorrências por divisão no formato CSV
	escreve_txt(txt_divisao, dic_divisoes)

	return


#####################################################################
#																	#
#						limpa_ocorrencias(col)						#
#						 Padroniza as strings						#
#	(Tira acento, caracteres especiais, trailling characters, etc)	#
#																	#
#####################################################################


def limpa_ocorrencias(col):

	# Iterador
	i = 0

	"""
	replace('text:\'', ''){
		remove [text:'] do começo de todas a strings (vem por default assim)
	}

	rstrip('\'!?'){
		remove [', !, ?] do final das strings
	}
	"""

	for string in col:

		# Explicado no bloco acima
		col[i] = string.replace('text:\'', '').rstrip('\'!?')

		# Remove [\r\n] das strings NESSA ORDEM
		col[i] = col[i].replace(r'\r\n', ' ')

		# Remove [\r] soltos no meio do texto
		col[i] = col[i].replace(r'\r', ' ')

		# Deixa tudo lower case
		col[i] = col[i].lower()

		# Remove acentos das strings
		col[i] = unidecode.unidecode(col[i])

		# Remove caracteres especiais das strings
		col[i] = re.sub('[\.\+\-\*\'\(\)\\\\/=",@!:_]+', ' ', col[i])

		# Remove trailling whitespace [' '] do começo e fim das strings
		col[i] = col[i].strip()

		# Remove espaços extras entre as strings
		col[i] = re.sub(' +', ' ', col[i])

		i += 1

	return col


#####################################################
#													#
#				limpa_solicitantes(col)				#
#				  Padroniza os nomes				#
#	(Tira caracteres especiais, deixa tudo lower)	#
#													#
#####################################################


def limpa_solicitantes(col):
	
	# Iterador
	i = 0

	"""
	replace('text\'', '') {
		Mesma coisa do limpa_ocorrencias()
		remove [text:'] do começo de todos os nomes (vem por default assim)
	}

	rstrip('\'') {
		Praticamente igual ao do limpa_ocorrencias(), porém só checa por '
		remove os ['] do final dos nomes
	}

	lower = tudo lowercase
	"""

	for solicitante in col:

		# Explicado acima
		col[i] = solicitante.replace('text:\'', '').rstrip('\'').lower()

		# Remove caracteres especiais dos nomes
		col[i] = re.sub('[\.\+\-\*\'\(\)\\\\/=",@!:_]+', ' ', col[i])

		i += 1

	return col


#####################################
#									#
#		 limpa_divisoes(col)		#
#	  remove a string [number:]		#
#	e o decimal [.0] das divisoes 	#
#									#
#####################################

def limpa_divisoes(col):

	# Iterador
	i = 0

	"""
	replace('number:', '') {
		remove [number:] do começo de todas as divisoes (vem por default assim)
	}

	replace('.0') {
		remove o decimal point dos numeros
	}
	"""

	for divisao in col:
		col[i] = divisao.replace('number:', '').replace('.0', '')
		i += 1
	
	return col


#############################################
#											#
#	 adiciona_elementos(col, dicionario)	#
#	   adiciona o nome do solicitante 		#
#	ou o numero da divisão no dicionario 	#
#											#
#############################################

def adiciona_elementos(col, dicionario):

	# Solicitante/Divisão
	for elemento in col:
		
		try:
			# Se solicitante/divisao não está no dicionario, coloca junto com um novo dicionario vazio
			if elemento not in dicionario:
				dicionario[elemento] = {}
		except Exception as e:
			print('Houve um erro\n{}'.format(e))


	return dicionario


#########################################################################
#																		#
#				relaciona_elementos(ocorrencias, col, dicionario)		#
#	relaciona as palavras usadas pelo elemento (solicitante/divisao)	#
#				e a quantidade de vezes que elas foram usadas			#
#																		#
#########################################################################

def relaciona_elementos(ocorrencias, dicionario, col = None):

	# Iterador
	i = 0

	# Total de Ocorrencias
	if col == None:
		# Pega todas as ocorrências individuais
		for texto in ocorrencias:

			# Separa cada palavra de uma ocorrencia numa lista que pode ser iterada
			palavras = texto.split()

			# Itera na lista de palavras das ocorrências
			for palavra in palavras:

				# Se palavra NÃO for menor igual a 2 caracteres
				# Exceto a palavra rg
				if not len(palavra) <= 2 or palavra == 'rg':
					try:
						# Se a [palavra] JÁ está no dicionario, adiciona +1 ocorrência a ela
						if palavra in dicionario:
							dicionario[palavra] += 1

						# Se a [palavra] NÃO está no dicionario, adiciona uma entrada a ela
						else:
							dicionario[palavra] = 1

					except Exception as e:
						print('Houve um erro\n{}'.format(e))


		return dicionario

	# Ocorrencias por solicitante/divisao
	else:

		# Por solicitante/divisao em col
		for elemento in col:

			# Por ocorrencia
			for texto in ocorrencias[i:]:

				# Divide as ocorrencias de [uma grande string] para ['uma', 'grande', 'string'] para ser mais fácil iterar
				palavras = texto.split()

				# Por palavra
				for palavra in palavras:

					# Se palavra NÃO for menor igual a 2 caracteres
					# Exceto a palavra rg
					if not len(palavra) <= 2 or palavra == 'rg':
						try:
							# Se a [palavra] JÁ está no dicionario, adiciona +1 ocorrência a ela
							if palavra in dicionario[elemento]:
								dicionario[elemento][palavra] += 1

							# Se a [palavra] NÃO está no dicionario, adiciona uma entrada a ela
							else:
								dicionario[elemento][palavra] = 1

						except Exception as e:
							print("Houve um erro\n{}".format(e))

				# Sincroniza as ocorrencias com os solicitantes
				i += 1

				# Próximo solicitante
				break

		return dicionario
	

#####################################################
#													#
#			escreve_txt(txt, dicionario)			#
#	Escreve no .txt as informações das colunas		#
#	Apenas no caso dos solicitantes ou divisão		#
#	O código para apenas ocorrências é diferente	#
#													#
#####################################################

def escreve_txt(txt, dicionario):

	# Por Solicitante/Divisão, palavras{}
	for chave, valor in dicionario.items():

		# Escreve "[chave];"
		# Nesse caso, o solicitante ou divisao
		txt.write('{};'.format(chave))

		# Separa as palavras do formato "[chave]:[valor]"
		# E mantém apenas o [chave]
		# Nesse caso, a palavra
		for palavra in valor:

			# Escreve "[palavra];[ocorrencias];"
			txt.write('{};{};'.format(palavra, dicionario[chave][palavra]))

	return


if __name__ == '__main__':
	main()
