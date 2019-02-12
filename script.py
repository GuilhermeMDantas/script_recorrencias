import os, os.path			# Funções do os
import xlutils, xlrd, xlwt	# Manage xls file
import logging				# Logs
import	unidecode			# Remover acentos
import re 					# Tratamento específico de string
import collections			# Criar nested dictionary (relação solicitante/divisao palavra)

def main():

	# Abre os sats
	workbook = xlrd.open_workbook('')

	# Abre o spreadsheet dos sats
	sheet = workbook.sheet_by_index(0)

	# dc_ocorrencia
	col_ocorrencia = sheet.col(8)

	# nm_solicitante
	col_solicitante = sheet.col(1)

	# fk_cd_divisao
	col_divisao = sheet.col(3)

	# Pegas todas as ocorrências e já as transforma em string
	ocorrencias = [str(ocorrencia) for ocorrencia in col_ocorrencia]
	del ocorrencias[0]  # Deleta o nome da coluna (dc_ocorrencia) da lista

	# Pega todos os solicitantes e já os transforma em string
	solicitantes = [str(solicitante) for solicitante in col_solicitante]
	del solicitantes[0]  # Deleta o nome da coluna (nm_solicitante) da lista

	# Pega todos as divisoes e já as transforma em string
	divisoes = [str(divisoes) for divisoes in col_divisao]
	del divisoes[0]  # Deleta o nome da coluna (fk_cd_divisao) da lista

	# Arquivos que vão ser escrito os dados extraídos
	arquivo_ocorrencia 	= open('total de ocorrencias.txt', 			'w+')
	arquivo_solicitante	= open('ocorrencia por solicitante.txt', 	'w+')
	arquivo_divisao		= open('ocorrencia por divisao.txt',		'w+')


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

	# Adiciona todas as palavras encontradas por ocorrencia
	arquivo_ocorrencia.write('Ocorrencias Total:\n{' + '\n' + '\n'.join('\t{}:{}'.format(k, v) for k, v in dic_ocorrencias.items()) + '\n' + '}')
	arquivo_ocorrencia.close()
	

	#########################################
	#										#
	#	Faz o tratamento dos solicitantes 	#
	#										#
	#########################################

	
	# Limpa os nomes e deixa tudo lowercase
	solicitantes = limpa_solicitantes(solicitantes)

	# Cria o dicionario principal
	palavra_por_solicitante = collections.defaultdict(dict)

	# Adiciona os solicitantes como [key]s do dicionario
	palavra_por_solicitante = adiciona_solicitantes(solicitantes, palavra_por_solicitante)

	# Relaciona quem fez a ocorrencia e quais foram as palavras
	palavra_por_solicitante = relaciona_solicitante_palavras(ocorrencias, palavra_por_solicitante, solicitantes)
	
	arquivo_solicitante.write('Palavras Por Solicitante\n{' + '\n' + '\n'.join('\t{}:\n\t\t{}'.format(k, v) for k, v in palavra_por_solicitante.items()) + '\n' + '}')
	arquivo_solicitante.close()

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


	arquivo_divisao.write('{' + '\n' + '\n'.join('\t{}:\n\t\t{}'.format(k, v) for k, v in dic_divisoes.items()) + '\n' + '}')
	arquivo_divisao.close()


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

	rstrip('\'!?\\r'){
		remove os [', \r, !, ?] do final das strings
	}

	replace('\\r\\n', ' '){
		remove os trailling characters [\r\n] que não foram pegos por estarem no meio da string (nessa ordem \r\n, pois parece ser o padrão)
		e adiciona um espaço no lugar deles
	}

	replace(',' | ':' | '!' | '"' | '(' | ')', ''){
		remove [', :, !, ", (, )] do meio das strings. Teve que ser 1 por vez porque o replace não pega [char] e sim uma literal
	}

	lower = tudo lowercase
	unidecode = tira os acentos

	# Serve para manter ips e emails
	re.sub('\.(?![a-zA-Z0-9]{2})') {
		'\.' 				Considera '.' como uma literal
		'()' 				Retorna o começo e fim da string dentro dos paranteses
		'?![a-z0-9]' 		Considera '.' APENAS se ele NÃO estiver seguido de caracteres hexadecimais (a-z, 0-9)
		?! 					Matches string1(.) ONLY if it is not followed by string2(range(a-z0-9))
		'{2}' 				Especifica que EXATAMENTE 3 cópias da expressão (no caso [a-z0-9]), se não, desconsidera a string
	}
	"""

	for string in col:

		col[i] = string.replace('text:\'', '').rstrip('\'!?\\r').replace('\\r\\n', ' ').replace(',', '').replace(':', '').replace('!', '').replace('"', '').replace('(', '').replace(')', '').lower()
		col[i] = re.sub('\.(?![a-z0-9]{2})', '', col[i])
		col[i] = unidecode.unidecode(col[i])
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

		col[i] = solicitante.replace('text:\'', '').rstrip('\'').lower()
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
				dic_solicitantes[solicitante]['palavras'] = {}
		except KeyError as e:
			print('error: ' + e)
			dic_solicitantes[solicitante] = 1

	return dic_solicitantes


#####################################################################
#																	#
#	relaciona_solicitante_palavras(ocorrencias, dic_solicitantes)	#
#						padroniza as strings						#
#	(tira acento, caracteres especiais, trailling characters, etc)	#
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
					if word in dic_solicitantes[solicitante]['palavras']:
						dic_solicitantes[solicitante]['palavras'][word] += 1

					# Se a [word] não está no dicionario do [solicitante]
					# Adiciona-a e já coloca que ela tem 1 ocorrencia
					else:
						dic_solicitantes[solicitante]['palavras'][word] = 1


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
				dic_divisoes[numero]['palavras'] = {}
		except KeyError as e:
			print('error: ' + e)
			dic_divisoes[numero] = 1


	return dic_divisoes


#
#
#
#
#


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
					if word in dic_divisoes[divisao]['palavras']:
						dic_divisoes[divisao]['palavras'][word] += 1

					# Se a [word] não está no dicionario das [divisoes]
					# Adiciona-a e já coloca que ela tem 1 ocorrencia
					else:
						dic_divisoes[divisao]['palavras'][word] = 1


			# Para ficar sincronizado a ocorrencia com as divisoes
			i+= 1

			break

	return dic_divisoes
	
if __name__ == '__main__':
	main()
