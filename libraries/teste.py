from data import Data
from conectSheets import ConectSheets


conect = ConectSheets()
conect.connect()

test = Data()
print("insira o valor")
print(conect.Lista)