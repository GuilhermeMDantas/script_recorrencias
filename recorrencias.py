import xlrd, xlsxwriter		# Manage xls file
import unidecode			# Remover acentos
import re 					# Tratamento específico de string
import collections			# Criar nested dictionary (relação solicitante/divisao palavra)

def main():

	# Abre os sats
	workbook_sats = xlrd.open_workbook('')

	# Abre o spreadsheet dos sats
	sheet_sats = workbook_sats.sheet_by_index(0)

	# dc_ocorrencia
	col_ocorrencia = sheet_sats.col(8)

	# nm_solicitante
	col_solicitante = sheet_sats.col(1)

	# fk_cd_divisao
	col_divisao = sheet_sats.col(3)

	# Pegas todas as ocorrências e já as transforma em string
	ocorrencias = [str(ocorrencia) for ocorrencia in col_ocorrencia]
	del ocorrencias[0]  # Deleta o nome da coluna (dc_ocorrencia) da lista

	# Pega todos os solicitantes e já os transforma em string
	solicitantes = [str(solicitante) for solicitante in col_solicitante]
	del solicitantes[0]  # Deleta o nome da coluna (nm_solicitante) da lista

	# Pega todos as divisoes e já as transforma em string
	divisoes = [str(divisoes) for divisoes in col_divisao]
	del divisoes[0]  # Deleta o nome da coluna (fk_cd_divisao) da lista


	txt_ocorrencias		= open(r'total de ocorrencias.txt', 			'w+')
	txt_solicitantes	= open(r'ocorrencia por solicitante.txt', 	'w+')
	txt_divisao			= open(r'ocorrencia por divisao.txt', 		'w+')


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
	dic_ocorrencias = conta_palavras(ocorrencias, dic_ocorrencias)
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
	dic_solicitantes = adiciona_solicitantes(solicitantes, dic_solicitantes)

	# Relaciona quem fez a ocorrencia e quais foram as palavras
	dic_solicitantes = relaciona_solicitante_palavras(ocorrencias, dic_solicitantes, solicitantes)

	# Escreve as ocorrências por solicitante no formato CSV
	for solicitante, palavras in dic_solicitantes.items():

		# Escreve "[solicitante];"
		txt_solicitantes.write('{};'.format(solicitante))

		# Separa as palavras do formato "[key]:[value]"
		# E mantém apenas o [key]
		for palavra in palavras:

			# Escreve "[palavra];[ocorrencias];"
			txt_solicitantes.write('{};{};'.format(palavra, dic_solicitantes[solicitante][palavra]))
			
		

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
	dic_divisoes = agrupa_divisoes(divisoes, dic_divisoes)

	# Relaciona as palavras com as divisões
	dic_divisoes = relaciona_divisao_palavras(ocorrencias, dic_divisoes, divisoes)

	# Escreve as ocorrências por divisao formato CSV
	for divisao, palavras in dic_divisoes.items():

		# Escreve "[divisao];"
		txt_divisao.write('{};'.format(divisao))

		# Separa as palavras do formato "[key]:[value]"
		# E mantém apenas o [key]
		for palavra in palavras:

			# Escreve "[palavra];[ocorrencias];"
			txt_divisao.write('{};{};'.format(palavra, dic_divisoes[divisao][palavra]))
			
		
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


#####################################################################
#																	#
#						conta_palavras(ocorrencias)					#
#			Conta a quantidade de vezes que uma palavra aparece		#
#																	#
#####################################################################


def conta_palavras(ocorrencias, dic_ocorrencias):

	# Pega todas as ocorrências individuais
	for string in ocorrencias:

		# Separa cada palavra de uma ocorrencia numa lista que pode ser iterada
		words = string.split()

		# Itera na lista de palavras das ocorrências
		for word in words:

			# Se o tamanho da palavra for IGUAL ou MENOR que 2 caracteres (pula "de", "a", "o", etc)
			# Com exceção da string 'rg'
			if not len(word) <= 2 or word == 'rg':
				try:
					if dic_ocorrencias[word]:
						dic_ocorrencias[word] += 1
				except KeyError:
					dic_ocorrencias[word] = 1

	return dic_ocorrencias


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


#####################################################################
#																	#
#		adiciona_solicitantes(solicitantes, dic_solicitantes)		#
#			   Adiciona os solicitantes [solicitantes]				#
#			No dicionario [dic_solicitantes] e retorna-o			#
#																	#
#####################################################################


def adiciona_solicitantes(solicitantes, dic_solicitantes):

	# Solicitante na coluna de solicitantes
	for solicitante in solicitantes:

		try:
			if solicitante in dic_solicitantes:
				# Se [solicitante] está no dicionarío não tem que oq fazer
				pass
			else:
				# Se [solicitante] não está no dicionario, coloca junto com as palavras
				dic_solicitantes[solicitante] = {}
		except KeyError as e:
			print('error: ' + e)
			dic_solicitantes[solicitante] = 1

	return dic_solicitantes


#####################################################################
#																	#
#	relaciona_solicitante_palavras(ocorrencias, dic_solicitantes)	#
#					Relaciona quais palavras foram 					#
#			usadas por [solicitante] numa [ocorrencia]				#
#																	#
#####################################################################


def relaciona_solicitante_palavras(ocorrencias, dic_solicitantes, solicitantes):
	# Pra iterar em cada ocorrencia
	i = 0

	# Itera em todos os solicitantes
	for solicitante in solicitantes:

		# Itera em todas as strings
		# [i:] da ocorrencia de index I até o fim do item
		for string in ocorrencias[i:]:
			words = string.split()

			# Pra cada palavra nessa ocorrencia
			for word in words:				

				# Ignora palavras com 2 ou menos caracteres
				# Exceção: rg
				if not len(word) <= 2 or word == 'rg':
					# Se a [word] já está no dicionario do [solicitante] que está fazendo a ocorrencia
					# Adiciona +1 ocorrencia pra [word] nesse [solicitante]
					if word in dic_solicitantes[solicitante]:
						dic_solicitantes[solicitante][word] += 1

					# Se a [word] não está no dicionario do [solicitante]
					# Adiciona-a e já coloca que ela tem 1 ocorrencia
					else:
						dic_solicitantes[solicitante][word] = 1


			# Para ficar sincronizado a ocorrencia com o solicitante
			i+= 1

			break
	return dic_solicitantes


#########################################
#										#
#		agrupa_divisoes(col)			#
#	agrupa todas as divisoes iguais		#
#	numa só ocorrencia do dicionario	#
#										#
#########################################


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
#		 agrupa_divisoes(divisoes)			#
#	 Agrupa duas ou mais ocorrências de 	#
#	  uma mesma divisão numa só entrada		#
#											#
#############################################


def agrupa_divisoes(divisoes, dic_divisoes):

	# numero na coluna de divisoes
	for numero in divisoes:
		
		try:
			if numero in dic_divisoes:
				# Se o número já está no dicionario
				# Não tem porque colocar ele novamente
				# Ou, pior, somar ele com ele mesmo
				pass
			else:
				# Se [numero] não está no dicionario, coloca junto com o dicionario ['palavras']
				dic_divisoes[numero] = {}
		except KeyError as e:
			print('error: ' + e)
			dic_divisoes[numero] = 1


	return dic_divisoes


#############################################################################
#																			#
#	  relaciona_divisao_palavras(ocorrencias, dic_divisoes, divisoes)		#
#	Relaciona quais foram as [palavra] usadas por uma [divisao] ao todo 	#
#																			#
#############################################################################


def relaciona_divisao_palavras(ocorrencias, dic_divisoes, divisoes):

	# Pra iterar em cada ocorrencia
	i = 0

	# Itera em todos os solicitantes
	for divisao in divisoes:

		# Itera em todas as strings
		# [i:] da ocorrencia de index I até o fim do item
		for string in ocorrencias[i:]:
			words = string.split()

			# Pra cada palavra nessa ocorrencia
			for word in words:				

				# Ignora palavras iguais ou menores a 2 caracteres
				# Exceção palavra rg
				if not len(word) <= 2 or word == 'rg':
					# Se a [word] já está no dicionario da [divisao] que está fazendo a ocorrencia
					# Adiciona +1 ocorrencia pra [word] nessa [divisao]
					if word in dic_divisoes[divisao]:
						dic_divisoes[divisao][word] += 1

					# Se a [word] não está no dicionario das [divisoes]
					# Adiciona-a e já coloca que ela tem 1 ocorrencia
					else:
						dic_divisoes[divisao][word] = 1


			# Para ficar sincronizado a ocorrencia com as divisoes
			i+= 1

			break

	return dic_divisoes




if __name__ == '__main__':
	main()
