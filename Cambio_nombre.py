import os

nombres_ori = []
nombres_ajustados = []
directorio = "G:\\Mi unidad\\WSP Ingenieria\\P1378\\2021\\Ajuste Anexos Laboratorios\\Torres Cuestecitas - La Loma\\TCCL- Especiales\\"
#directorio = directorio
os.chdir(directorio)
pos = 6  #Posición donde inician los 0 de los nombres  

for archivo in os.listdir(directorio):
    if  archivo[-4:].lower() == ".pdf":
        nombres_ori.append(archivo)

        if archivo[pos] == "0" and archivo[pos+1] == "0":
            if archivo.count("-") == 0:
                nombre_nuevo = (str(archivo[:pos]) +"-"+ (archivo[pos+2:])).upper()
            else:
                nombre_nuevo = (str(archivo[:pos]) + (archivo[pos+2:])).upper()
        elif archivo[pos] == "0":
            if archivo.count("-") == 0:
                nombre_nuevo = (str(archivo[:pos]) +"-"+ (archivo[pos+1:])).upper()
            else:
                nombre_nuevo = (str(archivo[:pos]) + (archivo[pos+1:])).upper()
        elif archivo[pos] != "0":
            if archivo.count("-") == 0:
                nombre_nuevo = (str(archivo[:pos]) + "-" + archivo[pos:]).upper()
            else:
                nombre_nuevo = (str(archivo[:pos]) + archivo[pos:]).upper()
        nombres_ajustados.append(nombre_nuevo)

for i in range(len(nombres_ajustados)):
    os.rename(os.path.join(directorio + nombres_ori[i]) , os.path.join(directorio + nombres_ajustados[i]))
    
    


 

    