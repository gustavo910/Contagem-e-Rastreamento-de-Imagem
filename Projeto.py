import cv2
import imutils
import numpy as np
import matplotlib.pyplot as plt
from tkinter import *
import tkinter.filedialog
from openpyxl import Workbook
wb = Workbook()

#O arquivo será a imagem cuja escolheu na interface grafica
arquivo = tkinter.filedialog.askopenfilenames()

soma=0
qtd =0
i=0
q=1
media=0.0


def display(img, count, cmap="gray"):
#Essa funçao tem o objetivo de mostrar a imagem pronta com os resultados de rastreamento e contagem
        f, axs = plt.subplots(1,2, figsize=(12, 5))
        axs[0].imshow(img, cmap="gray")
        axs[1].imshow(img, cmap="gray")
        axs[1].set_title("Contador = {}".format(count))

        #segunda imagem


def img2(figura, posicaofigura):
#essa funçao é a principal e tem o objetivo de fazer o tratamento de imagem,contador e o rastreamento dos objetos
    global media
    global soma
    wb.create_sheet("Imagem")
    #planilha = wb.worksheets[0]
    if posicaofigura == 0:
        planilha = wb["Imagem"]
    else:
        titulo = "Imagem" + str(posicaofigura)
        planilha = wb[titulo]
    

    #image = cv2.imread(arquivo[i])
    image = figura
    image = cv2.resize(image, (3000, 3000), interpolation=cv2.INTER_AREA)

    image_blur = cv2.medianBlur(image, 25)
    image_blur_gray = cv2.cvtColor(image_blur, cv2.COLOR_BGR2GRAY)
    image_res, image_thresh = cv2.threshold(image_blur_gray, 240, 250, cv2.THRESH_BINARY_INV)
    kernel = np.ones((5, 5), np.uint8)

    opening = cv2.morphologyEx(image_thresh, cv2.MORPH_OPEN, kernel)

    dilatacao= cv2.dilate(opening,kernel, iterations = 8)

    dist_transform = cv2.distanceTransform(opening, cv2.DIST_L2, 5)
    ret, last_image = cv2.threshold(dist_transform, 0.3 * dist_transform.max(), 255, 0)
    last_image = np.uint8(last_image)

    cnts = cv2.findContours(last_image.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = imutils.grab_contours(cnts)

    for (i, c) in enumerate(cnts):
        #esse for é onde se obtem a contagem de objetos e rastreamento destes
        ((x, y), _) = cv2.minEnclosingCircle(c)
        cv2.putText(image, "#{}".format(i + 1), (int(x) - 45, int(y) + 20), cv2.FONT_HERSHEY_SIMPLEX, 2, (0, 255, 0), 5)
        cv2.drawContours(image, [c], -1, (0, 255, 0), 2)
        m = len(cnts)

        comparacao(x, y,i,planilha,m)

    if posicaofigura == 0:
        media = m
        soma = m
    else:
        soma = soma + m
        media = soma/(cont+1)
    
    planilha['D1']="media"
    planilha['D2']= media
    #wb.save("meuArquivo.xlsx")

    display(image, len(cnts))
    return m

def comparacao(x,y,i,planilha,m):
    #essa funçao é a de saida de resultados no excel onde passamos da função anterior a posição do objeto na imagem
    #a quantidade e a media de objeto de tal imagem com as imagens anteriores
    a='A'
    b='B'

    #print(media, "media1")
    planilha['A1'] = "Posição X"
    planilha['B1'] = "Posição Y"
    planilha['C1'] = "Quantidade"
    planilha['C2'] =  m
    d=str(i+2)
    r=a+d
    s=b+d

    planilha[r] = x
    planilha[s] = y

for cont in range(0, len(arquivo)):
    #esse for é necessario para se quiser fazer a contagem de mais de uma imagem
    qtd=cont
    imagem = cv2.imread(arquivo[cont])
    somaImagem = img2(imagem, cont)

    i += 1
    q+=1


wb.save("meuArquivo.xlsx")
#saida dos resultados em excel