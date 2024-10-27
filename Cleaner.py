from unidecode import unidecode
import os

def folders(caminho_pasta):
    lista_pastas = os.listdir(caminho_pasta)

    for pasta in lista_pastas:
        if unidecode(pasta) != pasta:
            print(pasta,'\n')
            os.rename(caminho_pasta + '\\' + pasta,caminho_pasta + '\\' + unidecode(pasta))