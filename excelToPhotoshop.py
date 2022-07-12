from attr import mutable
import pandas as pd
import numpy as np
import matplotlib.pyplot as plb
import win32com.client
import os

dataName = '' #Caminho do arquivo .xlsx
df = pd.read_excel(dataName)

#dados da planilha divididos entre as layers no photoshop
#df[nome da coluna na planilha]
coluna01 = df['coluna1'].tolist()
coluna02 = df['coluna2'].tolist()
coluna03 = df['coluna3'].tolist()

contador = 0

#Por algum motivo não deu certo a modificação do caminho do arquivo, o que facilitaria alterar o numero no arquivo psd
#Então pode ser feita uma função para cada modelo .psd copiando esta função.
def modelo1(contador):
        psApp = win32com.client.Dispatch("Photoshop.Application")
        psApp.Open(r'resources\teste.psd') #Caminho do arquivo PSD para ser editado
        doc = psApp.Application.ActiveDocument

        layer_01(doc, contador)
        contador += 1
        saveImg(doc, contador)

        doc.Close(2)
        return contador + 1



def layer_01(doc, contador1):
    #Label que vai ser alterada no arquivo .psd.
    layers_01 = doc.ArtLayers["Layer1"].TextItem
    layers_02 = doc.ArtLayers["Layer2"].TextItem
    layers_03 = doc.ArtLayers["Layer3"].TextItem

    #Primeira metade da edição
    layers_01.contents = coluna01[contador1]
    layers_02.contents = coluna02[contador1]
    layers_03.contents = coluna03[contador1]


def saveImg(doc, contador):
    #Salva a imagem png
    options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    options.Format = 13   # PNG Formato
    options.PNG8 = False  # Setar para PNG-24 bit
    pngfile = '/resources/img_result/Image_'+ str(contador+1) +'.png'
    doc.Export(ExportIn=pngfile, ExportAs=2, Options=options)


#Quantas imagens criar;
for x in range(9):

    valueAtual1 = modelo1(contador)
    contador = valueAtual1

 