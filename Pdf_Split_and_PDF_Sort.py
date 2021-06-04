from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
import os

ruta = r'G:\Mi unidad\WSP Ingenieria\P1378\2021\Reportes'
ruta_salida = r'G:\Mi unidad\WSP Ingenieria\P1378\2021\Reportes\Salida'
ruta_salida_PDF = r'G:\Mi unidad\WSP Ingenieria\P1378\2021\Reportes\Salida\PDF'

orden = [1,2,3,4,5,6,14,7,8,9,10,11,12,13]

def Split_sheets(Archivo_pdf):

    pdf = PdfFileReader(Archivo_pdf)
    Num_paginas = pdf.getNumPages()
    Nom_archivo = Archivo_pdf.replace(".pdf","")
    ruta_salida_por_archivo = os.path.join(ruta_salida,Nom_archivo)
    os.mkdir(ruta_salida_por_archivo)
    for pagina in range(Num_paginas):

        pdf_writer = PdfFileWriter()
        Hoja_Actual = pdf.getPage(pagina)
        pdf_writer.addPage(Hoja_Actual)

        outputFilename = str(Nom_archivo)+"-{}.pdf".format(pagina + 1)

        with open(os.path.join(ruta_salida_por_archivo,outputFilename),"wb") as salida:
            pdf_writer.write(salida)
            print("created", outputFilename)
    return ruta_salida_por_archivo
    

def sort_pages(orden, ruta_archivos, Archivo_pdf):
    
    unidor = PdfFileMerger()
    i = 0
    for i,pagina in enumerate(os.listdir(ruta_archivos)):
        pagina = PdfFileReader( os.path.join(ruta_archivos,pagina) , 'r')
        unidor.append(pagina)
        i += 1
        if i == 6:
            break

    for pagina in os.listdir(ruta_archivos):
        if pagina.endswith("14.pdf"):
            pagina = PdfFileReader( os.path.join(ruta_archivos,pagina) , 'r')
            unidor.append(pagina)    

    for pagina in os.listdir(ruta_archivos):
        if pagina[-6:] != "14.pdf" and pagina[-6:] != "-1.pdf" and pagina[-6:] != "-2.pdf" and pagina[-6:] != "-3.pdf" and pagina[-6:] != "-4.pdf" and pagina[-6:] != "-5.pdf" and pagina[-6:] != "-6.pdf":
            pagina = PdfFileReader( os.path.join(ruta_archivos,pagina) , 'r')
            unidor.append(pagina)
    unidor.write(os.path.join(ruta_salida_PDF, Archivo_pdf))
    unidor.close()


for archivo in os.listdir(ruta):
    if archivo.endswith("pdf"):
        ruta_archivo = Split_sheets(archivo)
        #nom_archivo = archivo.replace(".pdf","")
        sort_pages(orden, ruta_archivo,archivo)




