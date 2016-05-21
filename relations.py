from xlutils.copy import copy
import xlrd
import xlwt
import os

def pegaLinha(arq, i):
	xls = xlrd.open_workbook(arq)
	plan = xls.sheets()[0]
	return plan.row_values(i)

def leSugestoes(arq):
	a = open(arq, 'r')
	sugestoes = []
	for i in range(10):
		linha = a.readline()
		sugestoes.append(linha.split(' ')[1])
	return sugestoes

def insere(book, linha, n):
	i = 0
	row = book.get_sheet(0).row(n)
	for item in linha:
		row.write(i, item)
		i = i + 1
	book.save("Relations.xls")

def cria(name):
	book = xlwt.Workbook()
	sheet = book.add_sheet("Sheet")
	book.save(name)
	return book

def nLinhas(book, encontrou):
	if encontrou:
		return book.sheets()[0].nrows
	else:
		return 0

def main():
	continuar = True
	encontrou = False
	listDir = os.listdir('../New-Relations-PrOntExt')
	for item in listDir:
		if(item == "contagem.txt"):
			encontrou = True
	if encontrou:	
		cont = open("contagem.txt", 'r')
	else:
		cont = open("contagem.txt", 'w+')
		cont.write('0');
	cont.seek(0)
	i = int(cont.read())
	cont.close()
	encontrou = False
	for item in listDir:
		if(item =="Relations.xls"):
			encontrou = True
	if not encontrou:
		cria("Relations.xls")
	bookRd = xlrd.open_workbook("Relations.xls")
	bookWt = copy(bookRd)
	j = nLinhas(bookRd, encontrou)
	while continuar:
		lista = pegaLinha("RelationsSheet.xls",i)
		print "Nome---------------------------------------------------------"
		print lista[0]
		inst = lista[14]
		instList = inst.split(' ')
		print "Instancias---------------------------------------------------"
		for item in instList:
			print item
		print "Sugestoes----------------------------------------------------"
		sugestList = leSugestoes("Relations-Names/Pr_"+str(i))
		for item in sugestList:
			print item
		op = raw_input("Deseja inserir essa relacao? (y/n)\n")
		if(op == 'y'):
			relation = raw_input("Digite o nome da relacao\n")
			lista[0] = relation	
			insere(bookWt, lista, j)
			j = j + 1
		i = i + 1
		prox = raw_input("Deseja analisar a proxima relacao?(y/n)\n")
		if(prox == 'n'):
			continuar = False;
	
	cont = open("contagem.txt", 'w')
	cont.write(str(i))
	cont.close()

if __name__ == "__main__":
	main()
