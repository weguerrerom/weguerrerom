import xlwings as xw
import os
import time

start_time = time.time()
Path_Folder =               "G:\\Mi unidad\\WSP Ingenieria\\P1378\\2021\\Archivos a PDF\\E.7.2 MEM_CALCULO\\sulfatos\\"
Path_Folder_Modificados =   "G:\\Mi unidad\\WSP Ingenieria\\P1378\\2021\\Archivos a PDF\\E.7.2 MEM_CALCULO\\sulfatos\\MODIFICADOS\\"
Path_Folder_Pdf =           "G:\\Mi unidad\\WSP Ingenieria\\P1378\\2021\\Archivos a PDF\\wetransfer-b94340\\h.natural\\Modificados Python\\PDF\\"
Des_Estratos = [
                ["0.00 - 0.50 - KEhn. Formación Hato Nuevo - Sc. Cuesta" , " ARCILLA LIMOSA Y ARENOSA DE COLOR AMARILLENTO MARRÓN, PLASTICIDAD MEDIA, CONSISTENCIA MUY RIGIDA Y HUMEDAD BAJA"],
                ["0.50 - 1.00 - KEhn. Formación Hato Nuevo - Sc. Cuesta" , " ARCILLA LIMOSA Y ARENOSA DE COLOR MARRÓN AMARILLO, PLASTICIDAD MEDIA, CONSISTENCIA MUY RIGIDA Y HUMEDAD BAJA"],
                ["1.00 - 1.50 - KEhn. Formación Hato Nuevo - Sc. Cuesta" , " ARCILLA LIMOSA Y ARENOSA DE COLOR GRIS, PLASTICIDAD MEDIA A ALTA, CONSISTENCIA DURA Y HUMEDAD BAJA"],
                ["1.50 - 2.00 - KEhn. Formación Hato Nuevo - Sc. Cuesta" , " ARCILLA LIMOSA CON ALGO DE ARENA, DE COLOR GRIS CLARO, PLASTICIDAD MEDIA A ALTA, CONSISTENCIA DURA Y HUMEDAD BAJA"],
                ["2.00 - 2.50 - KEhn. Formación Hato Nuevo - Sc. Cuesta" , " ARCILLA LIMOSA CON ALGO DE ARENA, DE COLOR GRIS CLARO, PLASTICIDAD MEDIA A ALTA, CONSISTENCIA DURA Y HUMEDAD BAJA"],
                ["2.50 - 3.00 - KEhn. Formación Hato Nuevo - Sc. Cuesta" , " ARCILLA LIMOSA CON ALGO DE ARENA FINA, DE COLOR GRIS CLARO, PLASTICIDAD MEDIA, CONSISTENCIA DURA Y HUMEDAD MEDIA"],
                ["0.00 - 0.50 - Nm. Formación Monguí - Sc. Cuesta" , " ARCILLA ARENOSA CON GRAVAS, DE COLOR AMARILLENTO, PLASTICIDAD MEDIA, CONSISTENCIA MUY RIGIDA Y HUMEDAD BAJA"],
                ["0.50 - 1.00 - Nm. Formación Monguí - Sc. Cuesta" , " ARCILLA LIMOARENOSA CON GRAVAS, DE COLOR AMARILLENTO, PLASTICIDAD MEDIA, CONSISTENCIA MUY RIGIDA Y HUMEDAD BAJA"],
                ["1.00 - 1.50 - Nm. Formación Monguí - Sc. Cuesta" , " ARCILLA LIMOSA Y ARENOSA CON GRAVAS, DE COLOR AMARILLENTO Y GRISÁCEO, PLASTICIDAD MEDIA, CONSISTENCIA RIGIDA Y HUMEDAD BAJA"],
                ["1.50 - 2.00 - Nm. Formación Monguí - Sc. Cuesta" , " ARCILLA LIMOSA CON ALGO DE ARENA, DE COLOR GRIS AMARILLENTO, PLASTICIDAD MEDIA, CONSISTENCIA DURA Y HUMEDAD BAJA"],
                ["2.00 - 2.50 - Nm. Formación Monguí - Sc. Cuesta" , " ARCILLA LIMOSA CON ALGO DE ARENA, DE COLOR GRIS CLARO AMARILLENTO, PLASTICIDAD MEDIA, CONSISTENCIA DURA Y HUMEDAD BAJA"],
                ["0.00 - 0.50 - Nc. Formación Castilletes - Klc. Lomo de carstificacion" , " ARCILLA ARENOSA DE COLOR AMARILLO HABANO, DE PLASTICIDAD MEDIA, CONSISTENCIA MEDIA Y HUMEDAD BAJA"],
                ["0.50 - 1.00 - Nc. Formación Castilletes - Klc. Lomo de carstificacion" , " ARCILLA ARENOSA DE COLOR AMARILLO HABANO, DE PLASTICIDAD MEDIA, CONSISTENCIA MEDIA Y HUMEDAD BAJA"],
                ["1.00 - 1.50 - Nc. Formación Castilletes - Klc. Lomo de carstificacion" , " ARCILLA LIMOSA CON ARENA DE COLOR AMARILLO Y OCRE, DE PLASTICIDAD MEDIA, CONSISTENCIA DURA Y HUMEDAD BAJA A MEDIA"],
                ["1.50 - 2.00 - Nc. Formación Castilletes - Klc. Lomo de carstificacion" , " ARCILLA LIMOSA CON ARENA DE COLOR AMARILLO Y OCRE, DE PLASTICIDAD MEDIA, CONSISTENCIA DURA Y HUMEDAD BAJA A MEDIA"],
                ["2.00 - 2.50 - Nc. Formación Castilletes - Klc. Lomo de carstificacion" , " ARENAS Y ARCILLAS LIMOSAS, DE COLOR HABANO, PLASTICIDAD BAJA, COMPACIDAD MEDIA Y HUMEDAD BAJA"],
                ["2.50 - 3.00 - Nc. Formación Castilletes - Klc. Lomo de carstificacion" , " ARENAS Y ARCILLAS LIMOSAS, DE COLOR HABANO, PLASTICIDAD BAJA, COMPACIDAD MEDIA Y HUMEDAD BAJA"],
                ["3.00 - 3.50 - Nc. Formación Castilletes - Klc. Lomo de carstificacion" , " ARENAS Y ARCILLAS LIMOSAS, DE COLOR AMARILLO Y HABANO, PLASTICIDAD BAJA, COMPACIDAD DENSA A MUY DENSA Y HUMEDAD BAJA"],
                ["3.50 - 4.00 - Nc. Formación Castilletes - Klc. Lomo de carstificacion" , " ARENAS ARCILLOSAS, DE COLOR AMARILLO Y HABANO, COMPACIDAD DENSA A MUY DENSA Y HUMEDAD BAJA"],
                ["4.00 - 4.50 - Nc. Formación Castilletes - Klc. Lomo de carstificacion" , " ARENAS LIMOSA, DE COLOR AMARILLO, COMPACIDAD DURA A MUY DENSA Y HUMEDAD BAJA"],
                ["4.50 - 5.00 - Nc. Formación Castilletes - Klc. Lomo de carstificacion" , " ARENA FINA LIMOSA DE COLOR AMARILLO OSCURO, COMPACIDAD MUY DENSA Y HUMEDAD BAJA A MEDIA"],
                ["5.00 - 5.50 - Nc. Formación Castilletes - Klc. Lomo de carstificacion" , " ARENA FINA LIMOSA DE COLOR AMARILLO OSCURO, COMPACIDAD MUY DENSA Y HUMEDAD BAJA A MEDIA"],
                ["5.50 - 6.00 - Nc. Formación Castilletes - Klc. Lomo de carstificacion" , " ARENA FINA LIMOSA DE COLOR AMARILLO OSCURO, COMPACIDAD MUY DENSA Y HUMEDAD BAJA A MEDIA"],
                ["0.00 - 0.50 - Pznm. Neis de Macuira - Dp. Planicie","ARENA Y ARCILLA LIMOSA DE COLOR GRIS Y OCRE , DE PLASTICIDAD MEDIA, CONSISTENCIA MEDIA Y HUMEDAD BAJA"],
                ["0.50 - 1.00 - Pznm. Neis de Macuira - Dp. Planicie","ARENA Y ARCILLA LIMOSA DE COLOR GRIS Y OCRE , DE PLASTICIDAD MEDIA, CONSISTENCIA MEDIA Y HUMEDAD BAJA"],
                ["1.00 - 1.50 - Pznm. Neis de Macuira - Dp. Planicie","ARENA Y ARCILLA LIMOSA DE COLOR GRIS Y AMARILLO, DE COMPACIDAD DENSA, PLASTICIDAD BAJA A MEDIA Y HUMEDAD BAJA"],
                ["1.50 - 2.00 - Pznm. Neis de Macuira - Dp. Planicie","ARENA Y ARCILLA LIMOSA DE COLOR GRIS Y AMARILLO, DE COMPACIDAD DENSA, PLASTICIDAD BAJA A MEDIA Y HUMEDAD BAJA"],
                ["2.00 - 2.50 - Pznm. Neis de Macuira - Dp. Planicie","ARENA LIMOSA DE COLOR GRIS Y CAFÉ, DE COMPACIDAD MUY DENSA Y HUMEDAD BAJA"],
                ["2.50 - 3.00 - Pznm. Neis de Macuira - Dp. Planicie","ARENA ARCILLOSA DE COLOR GRIS Y CAFÉ, DE COMPACIDAD DENSA A MUY DENSA Y HUMEDAD BAJA A MEDIA"],
                ["3.00 - 3.50 - Pznm. Neis de Macuira - Dp. Planicie","ARENA ARCILLOSA DE COLOR GRIS, DE COMPACIDAD DENSA A MUY DENSA Y HUMEDAD BAJA"],
                ["3.50 - 4.00 - Pznm. Neis de Macuira - Dp. Planicie","ARENA ARCILLOSA DE COLOR GRIS, COMPACIDAD DENSA A MUY DENSA, PLASTICIDAD MEDIA Y HUMEDAD BAJA"],
                ["4.00 - 4.50 - Pznm. Neis de Macuira - Dp. Planicie","ARENA LIMOSA DE COLOR GRIS, COMPACIDAD MUY DENSA, PLASTICIDAD MEDIA Y HUMEDAD BAJA"],
                ["4.50 - 5.00 - Pznm. Neis de Macuira - Dp. Planicie","NEIS DE COLOR GRIS DE GRANO GRUESO A MEDIO, CON OXIDACIONES, FRACTURADA Y DIACLASADA"],
                ["5.00 - 5.50 - Pznm. Neis de Macuira - Dp. Planicie","NEIS DE COLOR GRIS DE GRANO GRUESO A MEDIO, CON OXIDACIONES, FRACTURADA Y DIACLASADA"],
                ["5.50 - 6.00 - Pznm. Neis de Macuira - Dp. Planicie","NEIS DE COLOR GRIS DE GRANO GRUESO A MEDIO, CON OXIDACIONES, FRACTURADA Y DIACLASADA"],
                ["0.00 - 0.50 - Qal. Depósitos Aluviales - Fpi. Plano de inundación", "ARCILLA LIMOSA Y ARENOSA, DE COLOR GRIS Y MARRÓN, DE CONSISTENCIA MEDIA, PLASTICIDAD MEDIA Y HUMEDAD BAJA"],
                ["0.50 - 1.00 - Qal. Depósitos Aluviales - Fpi. Plano de inundación", "ARCILLA LIMOSA Y ARENOSA, DE COLOR GRIS Y MARRÓN, DE CONSISTENCIA MEDIA, PLASTICIDAD MEDIA Y HUMEDAD BAJA"],
                ["1.00 - 1.50 - Qal. Depósitos Aluviales - Fpi. Plano de inundación", "ARCILLA ARENOSA Y ARENA ARCILLOSA DE COLORES MARRÓN Y GRIS, DE CONSISTENCIA DURA, PLASTICIDAD MEDIA Y HUMEDAD BAJA"],
                ["1.50 - 2.00 - Qal. Depósitos Aluviales - Fpi. Plano de inundación", "ARCILLA ARENOSA Y ARENA ARCILLOSA DE COLORES MARRÓN Y GRIS, DE CONSISTENCIA DURA, PLASTICIDAD MEDIA Y HUMEDAD BAJA"],
                ["2.00 - 2.50 - Qal. Depósitos Aluviales - Fpi. Plano de inundación", "ARCILLAS LIMOARENOSAS DE COLOR MARRÓN, PLASTICIDAD MEDIA, CONSISTENCIA MEDIA Y HUMEDAD MEDIA"],
                ["2.50 - 3.00 - Qal. Depósitos Aluviales - Fpi. Plano de inundación", "ARCILLAS LIMOARENOSAS DE COLOR MARRÓN, PLASTICIDAD MEDIA, CONSISTENCIA MEDIA Y HUMEDAD MEDIA"],
                ["3.00 - 3.50 - Qal. Depósitos Aluviales - Fpi. Plano de inundación", "ARCILLA ARENOSA Y ARENA ARCILLOSA, DE COLOR CAFÉ Y HABANO, DE CONSISTENCIA RIGIDA, PLASTICIDAD BAJA Y HUMEDAD BAJA A MEDIA"],
                ["3.50 - 4.00 - Qal. Depósitos Aluviales - Fpi. Plano de inundación", "ARCILLA ARENOSA Y ARENA ARCILLOSA, DE COLOR CAFÉ Y HABANO, DE CONSISTENCIA RIGIDA, PLASTICIDAD BAJA Y HUMEDAD BAJA A MEDIA"],
                ["4.00 - 4.50 - Qal. Depósitos Aluviales - Fpi. Plano de inundación", "ARENA ARCILLOSA DE COLOR GRIS CLARO, DE COMPACIDAD DENSA A MUY DENSA Y HUMEDAD BAJA"],
                ["4.50 - 5.00 - Qal. Depósitos Aluviales - Fpi. Plano de inundación", "ARENA ARCILLOSA DE COLOR GRIS CLARO, DE COMPACIDAD DENSA A MUY DENSA Y HUMEDAD BAJA"],
                ["5.00 - 5.50 - Qal. Depósitos Aluviales - Fpi. Plano de inundación", "ARENA LIMOSA DE COLOR GRIS CLARO Y CAFÉ, DE COMPACIDAD MEDIA A MUY DENSA Y HUMEDAD BAJA"],
                ["5.50 - 6.00 - Qal. Depósitos Aluviales - Fpi. Plano de inundación", "ARENA LIMOSA DE COLOR GRIS CLARO Y CAFÉ, DE COMPACIDAD MEDIA A MUY DENSA Y HUMEDAD BAJA"],
                ["0.00 - 0.50 - Qal. Depósitos Aluviales - Dmo. Monticulo y ondulaciones denudadas","ARCILLA LIMOSA Y ARENOSA, DE COLOR MARRÓN, DE CONSISTENCIA DURA, PLASTICIDAD ALTA Y HUMEDAD MEDIA"],
                ["0.50 - 1.00 - Qal. Depósitos Aluviales - Dmo. Monticulo y ondulaciones denudadas","ARCILLA LIMOSA Y ARENOSA, DE COLOR MARRÓN, DE CONSISTENCIA DURA, PLASTICIDAD ALTA Y HUMEDAD MEDIA"],
                ["1.00 - 1.50 - Qal. Depósitos Aluviales - Dmo. Monticulo y ondulaciones denudadas","ARCILLA LIMOSA CON ALGO DE ARENA, DE COLOR MARRÓN Y GRIS, DE CONSISTENCIA DURA, PLASTICIDAD MEDIA A ALTA Y HUMEDAD MEDIA"],
                ["1.50 - 2.00 - Qal. Depósitos Aluviales - Dmo. Monticulo y ondulaciones denudadas","ARCILLA LIMOSA CON ALGO DE ARENA, DE COLOR MARRÓN Y GRIS, DE CONSISTENCIA DURA, PLASTICIDAD MEDIA A ALTA Y HUMEDAD MEDIA"],
                ["2.00 - 2.50 - Qal. Depósitos Aluviales - Dmo. Monticulo y ondulaciones denudadas","ARENAS Y ARCILLAS LIMOSAS, DE COLOR GRIS MARRÓN, PLASTICIDAD MEDIA, COMPACIDAD MEDIA Y HUMEDAD BAJA"],
                ["2.50 - 3.00 - Qal. Depósitos Aluviales - Dmo. Monticulo y ondulaciones denudadas","ARENAS Y ARCILLAS LIMOSAS, DE COLOR GRIS MARRÓN, PLASTICIDAD MEDIA, COMPACIDAD MEDIA Y HUMEDAD BAJA"],
                ["3.00 - 3.50 - Qal. Depósitos Aluviales - Dmo. Monticulo y ondulaciones denudadas","ARENA ARCILLOSA DE COLOR MARRÓN Y GRIS, DE COMPACIDAD DENSA Y HUMEDAD BAJA"],
                ["3.50 - 4.00 - Qal. Depósitos Aluviales - Dmo. Monticulo y ondulaciones denudadas","ARENA ARCILLOSA DE COLOR MARRÓN Y GRIS, DE COMPACIDAD DENSA Y HUMEDAD BAJA"],
                ["4.00 - 4.50 - Qal. Depósitos Aluviales - Dmo. Monticulo y ondulaciones denudadas","ARENAS Y ARCILLAS DE COLOR MARRÓN Y GRIS, DE COMPACIDAD DENSA A MUY DENSA Y HUMEDAD BAJA"],
                ["4.50 - 5.00 - Qal. Depósitos Aluviales - Dmo. Monticulo y ondulaciones denudadas","ARENAS Y ARCILLAS DE COLOR MARRÓN Y GRIS, DE COMPACIDAD DENSA A MUY DENSA Y HUMEDAD BAJA"],
                ["5.00 - 5.50 - Qal. Depósitos Aluviales - Dmo. Monticulo y ondulaciones denudadas","ARENA ARCILLOSA DE COLOR CAFÉ Y GRIS, DE COMPACIDAD MUY DENSA Y HUMEDAD BAJA"],
                ["5.50 - 6.00 - Qal. Depósitos Aluviales - Dmo. Monticulo y ondulaciones denudadas","ARCILLA ARENOSA DE COLOR CAFÉ AMARILLENTO, DE COMPACIDAD MUY DENSA Y HUMEDAD BAJA"],
                ["0.00 - 0.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Edl. Dunas longitudinales" , "ARCILLA ARENOSA, DE COLOR MARRÓN Y GRIS, DE CONSISTENCIA RIGIDA, PLASTICIDAD MEDIA Y HUMEDAD BAJA"],
                ["0.50 - 1.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Edl. Dunas longitudinales" , "ARCILLA ARENOSA, DE COLOR MARRÓN Y GRIS, DE CONSISTENCIA RIGIDA, PLASTICIDAD MEDIA Y HUMEDAD BAJA"],
                ["1.00 - 1.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Edl. Dunas longitudinales" , "ARCILLA ARENOSA, DE COLOR MARRÓN Y GRIS, DE CONSISTENCIA RIGIDA, PLASTICIDAD MEDIA Y HUMEDAD BAJA"],
                ["1.50 - 2.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Edl. Dunas longitudinales" , "ARCILLA ARENOSA, DE COLOR MARRÓN Y GRIS, DE CONSISTENCIA RIGIDA, PLASTICIDAD MEDIA Y HUMEDAD BAJA"],
                ["2.00 - 2.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Edl. Dunas longitudinales" , "ARCILLA Y ARENA DE COLOR MARRÓN Y GRIS, DE PLASTICIDAD MEDIA, CONSISTENCIA DURA Y HUMEDAD BAJA"],
                ["2.50 - 3.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Edl. Dunas longitudinales" , "ARCILLA Y ARENA DE COLOR MARRÓN Y GRIS, DE PLASTICIDAD MEDIA, CONSISTENCIA DURA Y HUMEDAD BAJA"],
                ["3.00 - 3.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Edl. Dunas longitudinales" , "ARCILLA LIMOSA CON ARENA DE COLOR AMARILLO ANARANJADO, DE CONSISTENCIA DURA, PLASTICIDAD MEDIA Y HUMEDAD MEDIA"],
                ["3.50 - 4.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Edl. Dunas longitudinales" , "ARCILLA LIMOSA CON ARENA DE COLOR AMARILLO ANARANJADO, DE CONSISTENCIA DURA, PLASTICIDAD MEDIA Y HUMEDAD MEDIA"],
                ["4.00 - 4.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Edl. Dunas longitudinales" , "ARENA ARCILLOSA DE COLOR ANARANJADO CLARO, DE COMPACIDAD MUY DENSA Y HUMEDAD BAJA"],
                ["4.50 - 5.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Edl. Dunas longitudinales" , "ARENA ARCILLOSA DE COLOR ANARANJADO CLARO, DE COMPACIDAD MUY DENSA Y HUMEDAD BAJA"],
                ["0.00 - 0.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ema. Mantos de arena eolica" ,"ARCILLA Y ARENA LIMOSA, DE COLOR AMARILLENTO Y ANARANJADO, DE CONSISTENCIA RIGIDA, PLASTICIDAD MEDIA Y HUMEDAD BAJA"], 
                ["0.50 - 1.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ema. Mantos de arena eolica" ,"ARCILLA Y ARENA LIMOSA, DE COLOR AMARILLENTO Y ANARANJADO, DE CONSISTENCIA RIGIDA, PLASTICIDAD MEDIA Y HUMEDAD BAJA"],
                ["1.00 - 1.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ema. Mantos de arena eolica" ,"ARCILLA ARENOSA DE COLOR MARRÓN Y AMARILLO, DE CONSISTENCIA DURA, PLASTICIDAD MEDIA Y HUMEDAD BAJA A MEDIA"],
                ["1.50 - 2.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ema. Mantos de arena eolica" ,"ARCILLA ARENOSA DE COLOR MARRÓN Y AMARILLO, DE CONSISTENCIA DURA, PLASTICIDAD MEDIA Y HUMEDAD BAJA A MEDIA"],
                ["2.00 - 2.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ema. Mantos de arena eolica" ,"ARCILLAS DE COLORES GRISÁCEOS Y CAFÉS, CON ALGO DE ARENA, PLASTICIDAD MEDIA A ALTA, CONSISTENCIA DURA Y HUMEDAD BAJA"],
                ["2.50 - 3.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ema. Mantos de arena eolica" ,"ARCILLAS DE COLORES GRISÁCEOS Y CAFÉS, CON ALGO DE ARENA, PLASTICIDAD MEDIA A ALTA, CONSISTENCIA DURA Y HUMEDAD BAJA"],
                ["3.00 - 3.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ema. Mantos de arena eolica" ,"ARCILLA ARENOSA Y ARENA ARCILLOSA, DE COLOR AMARILLENTO Y MARRÓN, DE CONSISTENCIA DURA, PLASTICIDAD MEDIA A ALTA Y HUMEDAD BAJA A MEDIA"],
                ["3.50 - 4.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ema. Mantos de arena eolica" ,"ARCILLA ARENOSA Y ARENA ARCILLOSA, DE COLOR AMARILLENTO Y MARRÓN, DE CONSISTENCIA DURA, PLASTICIDAD MEDIA A ALTA Y HUMEDAD BAJA A MEDIA"],
                ["4.00 - 4.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ema. Mantos de arena eolica" ,"ARENA ARCILLOSA Y ARCILLA ARENOSA, DE COLOR MARRÓN Y GRIS, DE CONSISTENCIA DURA, PLASTICIDAD BAJA A MEDIA Y HUMEDAD BAJA"],
                ["4.50 - 5.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ema. Mantos de arena eolica" ,"ARENA ARCILLOSA Y ARCILLA ARENOSA, DE COLOR MARRÓN Y GRIS, DE CONSISTENCIA DURA, PLASTICIDAD BAJA A MEDIA Y HUMEDAD BAJA"],
                ["5.00 - 5.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ema. Mantos de arena eolica" ,"ARENA ARCILLOLIMOSA, DE COLOR GRIS Y MARRÓN, DE COMPACIDAD MEDIA Y HUMEDAD BAJA"],
                ["5.50 - 6.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ema. Mantos de arena eolica" ,"ARENA ARCILLOLIMOSA, DE COLOR GRIS Y MARRÓN, DE COMPACIDAD MEDIA Y HUMEDAD BAJA"],
                ["0.00 - 0.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ftas. Terraza aluvial subreciente" , "ARCILLA ARENOSA DE COLOR MARRÓN ANARANJADO, DE PLASTICIDAD MEDIA, CONSISTENCIA MUY RIGIDA Y HUMEDAD BAJA"],
                ["0.50 - 1.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ftas. Terraza aluvial subreciente" , "ARCILLA ARENOSA DE COLOR MARRÓN ANARANJADO, DE PLASTICIDAD MEDIA, CONSISTENCIA RIGIDA A MUY RIGIDA Y HUMEDAD BAJA A MEDIA"],
                ["1.00 - 1.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ftas. Terraza aluvial subreciente" , "ARCILLA ARENOSA DE COLOR MARRÓN ANARANJADO Y GRIS, DE PLASTICIDAD MEDIA, CONSISTENCIA RIGIDA A MUY RIGIDA Y HUMEDAD BAJA A MEDIA"],
                ["1.00 - 1.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ftas. Terraza aluvial subreciente" , "ARENA ARCILLOSA CON ALGO DE LIMO, DE COLOR MARRÓN AMARILLENTO Y GRISÁCEO, DE COMPACIDAD MEDIA A DENSA Y HUMEDA BAJA"],
                ["1.50 - 2.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ftas. Terraza aluvial subreciente" , "ARENA ARCILLOSA ALGO LIMOSA, DE COLOR MARRÓN AMARILLENTO Y ANARANJADO, DE COMPACIDAD MEDIA A DENSA Y HUMEDA BAJA"],
                ["2.00 - 2.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ftas. Terraza aluvial subreciente" , "ARENA ARCILLOSA ALGO LIMOSA, DE COLOR MARRÓN AMARILLENTO Y ANARANJADO, DE COMPACIDAD MEDIA A MUY DENSA Y HUMEDA BAJA"],
                ["2.50 - 3.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ftas. Terraza aluvial subreciente" , "ARENA ARCILLOSA ALGO LIMOSA, DE COLOR MARRÓN AMARILLENTO Y ANARANJADO, DE COMPACIDAD MEDIA A MUY DENSA Y HUMEDA BAJA"],
                ["3.00 - 3.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ftas. Terraza aluvial subreciente" , "ARENA LIMOSA, DE COLOR MARRÓN AMARILLENTO, DE COMPACIDAD MUY DENSA Y HUMEDA BAJA"],
                ["3.50 - 4.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ftas. Terraza aluvial subreciente" , "ARENA LIMOSA CON PRESENCIA DE GRAVAS, DE COLOR MARRÓN AMARILLENTO, DE COMPACIDAD MUY DENSA Y HUMEDA BAJA"],
                ["4.00 - 4.50 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ftas. Terraza aluvial subreciente" , "ARENA DE GRANO FINO CON ALGO DE FINO, DE COLOR CAFÉ, COMPACIDAD DENSA Y HUMEDAD BAJA."],
                ["4.50 - 5.00 - Qale. Depósitos Aluviales recientes (influencia Eolica) - Ftas. Terraza aluvial subreciente" , "ARENA DE GRANO FINO CON ALGO DE FINO, DE COLOR CAFÉ, COMPACIDAD DENSA Y HUMEDAD BAJA"]
                ]


for Archivo in os.listdir(Path_Folder):
    if Archivo[-5:].lower() == ".xlsx" and Archivo[-11:] != "BONDAD.xlsx":
        app = xw.App(visible = False)
        Libro = xw.Book(str(Path_Folder + Archivo))
        Hoja1 = Libro.sheets["tamaño_muestra"] 
        Hoja2 = Libro.sheets["Datos"] 
        Hoja3 = Libro.sheets["estadística"] 
        Hoja4 = Libro.sheets["fdp"] 
        Hoja5 = Libro.sheets["fdp_acum"] 
        Hoja6 = Libro.sheets["fdp_graficas"] 

        validacion_inter = round(Hoja1.range("C8").value / Hoja1.range("C7").value,2) >= 0.25 and Hoja1.range("C8").value >= 3 and Hoja2.range("A15").value is not None

        if validacion_inter:
            Estrato_buscado = Hoja1.range("K5").value + " - " + Hoja1.range("C5").value +" - "+Hoja1.range("C6").value
            for info_estrato,des_estrato in Des_Estratos:
                if info_estrato == Estrato_buscado:
                    Hoja1.range("I6").value = des_estrato

            Hoja1.api.PageSetup.PrintArea= ("$A$1:$L$102")
            Hoja2.api.PageSetup.PrintArea= ("$A$1:$L$72")
            Hoja3.api.PageSetup.PrintArea= ("$A$1:$M$74")
            Hoja4.api.PageSetup.PrintArea= ("$A$1:$M$254")
            Hoja5.api.PageSetup.PrintArea= ("$A$1:$N$127")
            Hoja6.api.PageSetup.PrintArea= ("$A$1:$O$189")
            Libro.save(str(Path_Folder_Modificados + Archivo))
            Libro.close()
        app.quit()
    
    if Archivo[-5:].lower() == ".xlsx" and Archivo[-11:] == "BONDAD.xlsx":
        app = xw.App(visible = False)
        Libro = xw.Book(str(Path_Folder + Archivo))
        Hoja1 = Libro.sheets["CDF_medidos"] 
        Hoja2 = Libro.sheets["CDF_interpolados"] 
        Hoja3 = Libro.sheets["Prueba de Bondad"] 

        Validacion_bondad = round(Hoja1.range("C8").value / Hoja1.range("C7").value,2) >= 0.25 and Hoja1.range("C8").value >= 3 and Hoja2.range("A13").value is not None

        if Validacion_bondad:
            Hoja1.api.PageSetup.PrintArea= ("$A$1:$L$122")
            Hoja2.api.PageSetup.PrintArea= ("$A$1:$L$122")
            Hoja3.api.PageSetup.PrintArea= ("$A$1:$L$60")
            Libro.save(str(Path_Folder_Modificados + Archivo))
            Libro.close()
        app.quit()

end = time.time()
elapsed_time = end - start_time
print(elapsed_time)
