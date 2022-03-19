
import os

import csv
import shutil

from datetime import datetime, time
import time

import smtplib, ssl

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

ARCHIVO_CORREOS = "reporte_encuestas.csv"
ARCHIVO_FORMS = "formularios.csv"
ARCHIVO_ARTICULOS = "articulos.csv"

ARCHIVO_BIEN_ENVIADOS = "enviados_correctamente.csv"
ARCHIVO_NO_ENVIADOS = "no_califican.csv"
ARCHIVO_MAL_ENVIADOS = "ENVIOS_FALLIDOS.csv"
ARCHIVO_STATS = "informe_envio.txt"

ARCHIVO_ENVIOS_HISTORICOS = "envios_historicos.csv"

ARCHIVO_TOKEN = "token.json"


def validar_opcion(opc_minimas: int, opc_maximas: int, texto: str = '') -> str:
    """
    PRE: Recibe los int "opc_minimas" y "opc_maximas" que 
    simbolizan la cantidad de opciones posibles.
    
    POST: Devuelve en formato string la var "opc" con un número 
    entero dentro del rango de opciones.
    """
    opc = input("{}".format(texto))
    while not opc.isnumeric() or int(opc) > opc_maximas or int(opc) < opc_minimas:
        opc = input("\tPor favor, ingrese una opcion valida: ")
    
    return opc


def enviar_email( correo:str, cuerpo:str, asunto:str, mail: smtplib.SMTP) -> bool:
    """Envia emails con el msj y asunto pasados por parametro, a los correos recibidos
    
    Args:
        correo (str): Correo al que se le desea enviar el mail    
        cuerpo (str): texto que ira en el cuerpo del mail. Debe estar en formato HTML
        asunto (str): el asunto (subject) del mail
        mail(smtplib.SMTP): Objeto de tipo smptlib.SMTP que contiene la info de inicio de sesion del usuario
    
    Returns:
        res (bool): El booleano res es true si se envio correctamente el mail, false en caso contrario
    """
    res = None    
    try:  
        # me == my email address
        # you == recipient's email address
        me = mail.user
        you = correo

        # Create message container - the correct MIME type is multipart/alternative.
        msg = MIMEMultipart('alternative')
        msg['Subject'] = asunto
        msg['From'] = me
        msg['To'] = you

        # Create the body of the message (a plain-text and an HTML version).
        # Record the MIME types of both parts - text/plain and text/html.
        #part1 = MIMEText(text, 'plain')
        part_html = MIMEText(cuerpo, 'html')

        # Attach parts into message container.
        # According to RFC 2046, the last part of a multipart message, in this case
        # the HTML message, is best and preferred.
        msg.attach(part_html)
        
        # Send the message via local SMTP server.
        mail.sendmail(me,you, msg.as_string())
        res = True
    except:
        res = False
        print("\n\tOCURRIO UN PROBLEMA AL ENVIAR CORREO A:",correo,"\n")    
    return res 
    

def cargar_ventas(dict_ventas: dict, path_correos: str) -> bool:
    """Carga el diccionario ventas con la info del archivo ARCHIVO_CORREOS.csv
    
    Args:
        dict_ventas (dict): Tendra la informacion de las ventas. Cada venta es un diccionario que
        esta como valor de la primary key "remito" del dict_ventas.
        
        dict_ventas = {remito: venta = {campo:info} }
    
    Returns:
        res (bool): El booleano res es true si se leyo correctamente el mail, false en caso contrario
    """
        
    res = True
    lista_errores = list()
    try:
        with open(path_correos, newline='') as archivo:
            
            csv_reader = csv.reader(archivo, delimiter=';')
            encabezado = next(csv_reader)
            cont = 2                
            
            for linea in archivo:
                
                linea = linea.strip()
                linea = linea.split(";")
                if len(linea) == 9:
                    try:
                        remito = linea[3]
                                    
                        if remito not in dict_ventas.keys():
                            
                            venta = dict()
                            
                            venta["contacto"] = linea[6]
                            venta["correo"] = linea[8]
                            venta["remito"] = remito
                            venta["articulos"] = [ linea[0] ] #una lista de articulos
                            venta["tipos_articulos"] = [] #una lista de tipos de articulos x ahora vacia
                            venta["formulario"] = "" #x ahora es un string vacio
                            venta["link_formulario"] = "" #x ahora es un string vacio
                            venta["asunto_correo"] = "" #x ahora es un string vacio
                            venta["tipo_servicio"] = "" #x ahora es un string vacio
                            
                            #venta["vendedores"] = [ linea[5] ]
                            # TODO AGREGAR CAMPO CON NOMBRE ASISTENTE DE VENTAS!!!
                            
                            dict_ventas[remito] = venta #agrego la venta al dict de ventas
                            
                            ultimo_remito = remito
                                                
                        else:   #agrego el "producto o tipo producto a la venta correspondiente"
                            #vendedor = linea[5]
                            
                            articulo = linea[0]
                            dict_ventas[remito]["articulos"].append(articulo)
                                                                        
                    except IndexError as faltan_columnas_reporte_encuestas:
                        error = "\t - Se produjo un error al leer la linea: " + str(cont)
                        lista_errores.append( error )
                        res = False
                        
                else:
                    error = "\t - Se produjo un error al leer la linea: " + str(cont)
                    lista_errores.append( error )
                    res = False
                
                cont +=1
                                    
            if lista_errores:
                print("\t\t\t","FALTAN O SOBRAN COLUMNAS EN EL ARCHIVO",ARCHIVO_CORREOS,"\n")
                for error in lista_errores:
                    print (error)
                    
    except FileNotFoundError:
        print("\t\t\t","NO SE PUDO ABRIR EL ARCHIVO reporte_encuestas.csv","\n")
        res = False
                
    return res

                  
def cargar_dict_forms(dict_forms: dict, path_forms: str):
    """Carga el diccionario dict_forms leyendo el archivo ARCHIVO_FORMS

    Args:
        dict_forms (dict): Tiene como claves los tipos de articulo y como valores el formulario q le
        corresponden a cada uno.
        
        dict_ventas = {tipo_articulo: formulario}
    
    Returns:
        res (bool): El booleano res es true si se leyo correctamente el mail, false en caso contrario
    """
    res = True
    lista_errores = list()
    
    try:
        with open(path_forms, newline='') as archivo:
            
            csv_reader = csv.reader(archivo, delimiter=';')
            encabezado = next(csv_reader)
            
            cont = 2
            for linea in archivo:     
                   
                linea = linea.strip().split(";")
                if len(linea) == 2:
                        
                    try:        
                    
                        tipo_articulo = linea[0]
                        link_formulario = linea[1]

                        dict_forms[tipo_articulo] = link_formulario
                    
                    except IndexError as faltan_columnas_reporte_encuestas:
                        error = "\t - Se produjo un error al leer la linea: " + str(cont)
                        lista_errores.append( error )
                        res = False
                else:
                    error = "\t - Se produjo un error al leer la linea: " + str(cont)
                    lista_errores.append( error )
                    res = False
                
                cont+=1                
            
            if lista_errores:
                print("\t\t\t","FALTAN DATOS O SOBRAN COLUMNAS EN EL ARCHIVO",ARCHIVO_FORMS,"\n")
                for error in lista_errores:
                    print (error)
                         
    except FileNotFoundError:
        res = False
        print("\t\t\t","NO SE PUDO ABRIR EL ARCHIVO formularios.csv","\n")
    
    return res
    

def cargar_dict_articulos(dict_articulos: dict, path_articulos:str):
    """Carga el diccionario articulos con la info del archivo ARCHIVO ARTICULOS.

    Args:
        dict_articulos (dict): Tiene como claves los articulos y como valore el tipo de articulo que es cada
        uno (Implementacion, Mantenimiento, Producto, -)
    
    Returns:
        res (bool): El booleano res es true si se leyo correctamente el mail, false en caso contrario.
    """
    res = True
    lista_errores = list()
    
    try:
        with open(path_articulos, newline='') as archivo:
            
            csv_reader = csv.reader(archivo, delimiter=';')
            encabezado = next(csv_reader)
            
            cont = 2
            for linea in archivo:   
                
                linea = linea.strip().split(";")
                if len(linea) == 3:
                    try:     
                        articulo = linea[0]
                        clase_articulo = linea[1]
                        tipo_articulo = linea[2]
                        #comentario = linea[3]
                        #descripcion = linea[4]
                        
                        dict_articulos[articulo] = tipo_articulo
                    
                    except IndexError as faltan_columnas_reporte_encuestas:
                        error = "\t - Se produjo un error al leer la linea: " + str(cont)
                        lista_errores.append( error )
                        res = False
                else:
                    error = "\t - Se produjo un error al leer la linea: " + str(cont)
                    lista_errores.append( error )
                    res = False
                
                cont +=1
                
            if lista_errores:
                print("\t\t\t","FALTAN O SOBRAN COLUMNAS EN EL ARCHIVO",ARCHIVO_ARTICULOS,"\n")
                for error in lista_errores:
                    print (error)                    
                    
    except FileNotFoundError:
        res = False
        print("\t\t\t","NO SE PUDO ABRIR EL ARCHIVO",ARCHIVO_ARTICULOS,"\n")
    
    return res

def armar_mensaje(venta: dict) -> str:
    """Arma el mensaje a enviar por coreo

    Args:
        Venta (dict): tiene la info de la venta
    
    Returns:
        mensaje(str): Es el mensaje a enviar
    """
    
    contacto = venta["contacto"]
    remito = venta["remito"]
    link_formulario = venta["link_formulario"]
    tipo_servicio = venta["tipo_servicio"]
    
    mensaje =  """
    <font size="2" face="arial" color="black">
    Estimado/a {0},<br><br>En base al servicio de {1} realizado y correspondiente al remito nº 
    {2} le enviamos nuestra encuesta de satisfacción al cliente.<br><br>
    Su opinión nos ayudará a mejorar nuestro servicio y no le tomará mas de un minuto.<br><br>
    </font>
    
    <font size="4" face="arial" color=0000FF>
    <strong> <a href="{3}">CLICK AQUÍ PARA RESPONDER LA ENCUESTA</a> </strong>
    </font>
    
    <br><br>Muchísimas gracias por su colaboración <br><br>
    
    <i>Encuestas de Satisfacción al Cliente<i> <br>
    <i>12SOFTWARE S.R.L <i> <br>
    
    <font size="1" face="arial" color="black">
    Capital federal <br>
    Buenos Aires - Argentina <br>
    E-mail: germandouce@gmail.com<br>
    Tel.: (+54 911) 6424-5970 <br>
    Web: www.12software.com<br>
    </font>
    <img src="https://upload.wikimedia.org/wikipedia/commons/5/56/Rick_Astley.jpg"> """.format(contacto,tipo_servicio,remito, link_formulario)
    
    return mensaje


def enviar_correos(dict_ventas: dict, lista_bien_enviados: list,lista_no_enviados:list, 
lista_mal_enviados: list, lista_advertencias: list, correo_remitente: smtplib.SMTP):
    """Emvia los correos leyendo el diccionario ventas

    Args:
        dict_ventas (dict): Contiene la info de todas las ventas y remitos
    """
    cont = 0
    for remito, venta in dict_ventas.items():
        
        correo = venta["correo"]
        asunto = venta["asunto_correo"]
        tipo_encuesta = venta["formulario"]
        
        renglon = remito + ";" + correo + ";" + venta["contacto"] + ";" + tipo_encuesta
        
        enviar = (venta["formulario"] != "-" and venta["formulario"] != "ART_NO_ENCONTRADO" and remito[:6] == "R00007")
        #print(remito[:6])
        
        if enviar: 
            
            mensaje = armar_mensaje(venta)    
            
            time.sleep(2)            

            bien_enviado = enviar_email(correo, mensaje, asunto, correo_remitente)
            
            envios = 1
            while not bien_enviado and envios <=5:
                time.sleep(2)            
                envios+=1
                print("\tenviando por",envios,"a vez a", correo)
                bien_enviado = enviar_email(correo, mensaje, asunto, correo_remitente)                
                            
            if bien_enviado:
                lista_bien_enviados.append(renglon)
                print("\tSe envio encuesta a", venta["contacto"],":", correo)
            else:
                lista_mal_enviados.append(renglon)
                if correo == "":
                    correo = "no_habia_correo"
                error = "\n\tSE PRODUJO UN ERROR DESPUES DE 6 INTENTOS DE ENVIAR UN CORREO A" + " " + correo + "\n"
                lista_advertencias.append(error)                
        else:
            lista_no_enviados.append(renglon)
            #print("No se envio un correo porque no calificaba a", correo)
            
        cont += 1


def agregar_formularios(dict_ventas: dict, dict_forms: dict, dict_articulos: dict, lista_advertencias: list):
    
    for remito, venta in dict_ventas.items():
            formulario = ""
            link_formulario = ""
            asunto_correo = ""
            tipo_servicio = ""
                       
            for articulo in venta["articulos"]:
                
                if articulo in dict_articulos.keys():
                
                    tipo_articulo = dict_articulos[articulo]
                    
                    venta["tipos_articulos"].append(tipo_articulo)
                else:
                    venta["tipos_articulos"].append("ART_NO_ENCONTRADO")
                    error = "\n\tEL ARTICULO " + articulo + " NO SE ENCUENTRA EN EL LISTADO DE ARTICULOS LOCAL\n"
                    lista_advertencias.append(error)
            #fin carga tipos de articulos
            
            if "ART_NO_ENCONTRADO" in venta["tipos_articulos"]: #era una lista
            #Al menos un articulo no encontrado no envia mail
                formulario = "ART_NO_ENCONTRADO"
                link_formulario = "ART_NO_ENCONTRADO"
                asunto_correo = "ART_NO_ENCONTRADO"
                tipo_servicio = "ART_NO_ENCONTRADO"
            
            elif "Implementacion" in venta["tipos_articulos"] :
            #Al menos un articulo de implementaciones - envia form implementacion
                formulario = "Implementacion"
                link_formulario = dict_forms["Implementacion"]
                asunto_correo = "Encuesta - Implementacion de sistemas- 12SOFTWARE SRL"
                tipo_servicio = "implementacion de sistema"
            
            elif "Mantenimiento" in venta["tipos_articulos"]: #era una lista
            #Al menos un articulo de Servicio tecnico - envia form servicio tecnico
                formulario = "Mantenimiento"
                link_formulario = dict_forms["Mantenimiento"]
                asunto_correo = "Encuesta - Mantenimiento de sistemas 12SOFTARE SRL"
                tipo_servicio = "mantenimiento de sistema"
                
            elif "-"  in venta["tipos_articulos"]:
                #Al menos un Articulo sin formulario - no envia email
                formulario = "-"
                link_formulario = dict_forms["-"]
                asunto_correo = "-"
                tipo_servicio = "-"  
            
            else:
                #ADVERTENCIA OJO QUE SI NO ES NADA DE LO ANTERIOR ENTRA AQUI!!!!
                #Else, envia form productos
                formulario = "Producto" 
                link_formulario = dict_forms["Producto"]
                asunto_correo = "Encuesta - Venta de de Productos - 12SOFTWARE SRL"
                tipo_servicio = "venta de producto"
                
            venta["formulario"] = formulario
            venta["link_formulario"] = link_formulario
            venta["asunto_correo"] = asunto_correo
            venta["tipo_servicio"] = tipo_servicio


def crear_archivo_resumen_correos(lista_correos: list, nombre_archivo: str) -> None:
    """Crea y cargala lista enviada por parametro colocando cada elemento de la misma
    en un renglon del archivo nombre_archivo

    Args:
        lista_correos (list): Cada elemento de la lista sera una linea del archivo y
        tiene info del remito al que e le envio ese correo
        
        nombre_archivo (str): un string con el nombre del archivo que se crea temporalmente 
        en el directorio de ejcucion del progrma
    """
    with open(nombre_archivo, "w", encoding="UTF-8") as archivo:
        for rengon in lista_correos:
            archivo.write(rengon + "\n")
   

def crear_archivo_stats(inicio_envio: datetime, fin_envio: datetime, tiempo_envio: datetime,
total_procesados: int ,bien_enviados: int, no_enviados: int, mal_enviados: int, ) -> None:
    """Crea un archivo con un resumen de la info del envio de correos electronicos de la corrida

    Args:
        inicio_envio (int): Objeto datetime con hora de inicio
        bien_enviados (int): numero de correso bien enviados
        no_enviados (int): numero de correso no enviados
        mal_enviados (int): numero de correso mal enviados
        tiempo_envio (int): Objeto datetime con tiempo total de envio
    """
    with open(ARCHIVO_STATS, "w", encoding="UTF-8") as archivo:
        archivo.write("----------------- INFORME DEL ENVIO ---------------------\n\n")
        archivo.write("TOTAL PROCESADOS: " + str(total_procesados) + "\n\n" )
        archivo.write("ENVIADOS CORRECTAMENTE: " + str(bien_enviados) + "\n\n" )
        archivo.write("NO EVIADOS PORQUE NO CALIFICAN: " + str(no_enviados) + "\n\n" )
        archivo.write("ENVIOS FALLIDOS: " + str(mal_enviados) + "\n\n" )
        archivo.write("_________________________________________\n\n")
        archivo.write("INICIO ENVIO: " + str(inicio_envio) + "\n\n" )
        archivo.write("HORA FIN: " + str(fin_envio) + "\n\n" )
        archivo.write("TIEMPO ENVIO: " + str(tiempo_envio) + "\n\n")


def archivar_archivos_de_registro_generados(no_enviados: int, mal_enviados: int, fin_envio: datetime)-> None:
    """[summary]

    Args:
        no_enviados (int): cantidad de corres no enviados
        mal_enviados (int): cantidad de corres mal enviados (fallo el envio)
        fin_envio (datetime): objeto datetime con la hora del final del envio
    """
    #ruta_actual = "..\\"+os.getcwd()
    ruta_actual = os.path.normpath(os.getcwd() + os.sep + os.pardir)

    path_correos = ruta_actual + "\\" + ARCHIVO_CORREOS


    
    ruta_registro_reportes = ruta_actual + "\\registro_reportes"
    ruta_reporte_nuevo_dia = ruta_registro_reportes + "\\" + str(fin_envio.year) +"-"+ str(fin_envio.month)+"-"+str(fin_envio.day)
        
    if not ( os.path.exists(ruta_registro_reportes) ):
        os.mkdir(ruta_registro_reportes)
            
    if not ( os.path.exists(ruta_reporte_nuevo_dia) ):
        os.mkdir(ruta_reporte_nuevo_dia)    
    
    ruta_nuevo_reporte = ruta_reporte_nuevo_dia + "\\" +"reporte"

    ruta_nuevo_reporte = ruta_reporte_nuevo_dia + "\\" +"reporte " + str(fin_envio.hour) +"h " + str(fin_envio.minute)
    ruta_nuevo_reporte = ruta_nuevo_reporte + "m " + str(fin_envio.second) + "s"

    os.mkdir(ruta_nuevo_reporte)
    
    path_archivo_reporte_encuestas = ruta_nuevo_reporte + "\\" + ARCHIVO_CORREOS
    path_archivo_bien_enviados = ruta_nuevo_reporte + "\\" + ARCHIVO_BIEN_ENVIADOS
    path_archivo_no_enviados = ruta_nuevo_reporte + "\\" + ARCHIVO_NO_ENVIADOS
    path_archivo_mal_enviados = ruta_nuevo_reporte + "\\" + ARCHIVO_MAL_ENVIADOS
    path_archivo_stats = ruta_nuevo_reporte + "\\" + ARCHIVO_STATS
    
    shutil.move(ARCHIVO_BIEN_ENVIADOS, path_archivo_bien_enviados)
    
    if no_enviados:
        shutil.move(ARCHIVO_NO_ENVIADOS, path_archivo_no_enviados)
    else:
        os.remove(ARCHIVO_NO_ENVIADOS)
    
    if mal_enviados:    
        shutil.move(ARCHIVO_MAL_ENVIADOS, path_archivo_mal_enviados)
    else:
        os.remove(ARCHIVO_MAL_ENVIADOS)
    
    shutil.move(ARCHIVO_STATS, path_archivo_stats)
    
    #muevo reporte encuestas para no generar problemas si se olvidan de cambiarlo
    shutil.move(path_correos, path_archivo_reporte_encuestas)


def escribir_registro_historico(lista_bien_enviados: list, fin_envio: datetime):
    """Escribe un archivo con un registro historico de correso enviados agregandole la info que esta
    en list_bien_enviados y la fecha de fin_envio. Si el archivo no existe,lo crea, caso contrario
    agrega al final los nuevos registros.

    Args:
        lista_bien_enviados (list): Cada elemento sera un renglon del archivo y tiene info de ese correo enviado
        fin_envio (datetime): objeto datetime con la hora de fin del envvio de los correos.
    """
    
    ruta_actual = os.path.normpath(os.getcwd() + os.sep + os.pardir)
    ruta_registro_reportes = ruta_actual + "\\registro_reportes" 
    
    path_archivo_envios_historicos = ruta_registro_reportes + "\\" + ARCHIVO_ENVIOS_HISTORICOS
    
    if os.path.exists(ruta_registro_reportes): 
        with open(path_archivo_envios_historicos, "a+", encoding="UTF-8") as archivo:
            for renglon in lista_bien_enviados:
                archivo.write( str(fin_envio) + ";" + renglon + "\n" )
    else:
        print("-- No se creo correctamente el direcorio de resgistro de reportes --")


def mostrar_advertencia_existencia_archivos(path_correos,path_forms, path_articulos) -> None:
    """Muestra por pantalla un mesnaje al ususario para chequear existencia de archivos y luego
    el estado actual de los mismos
    """
    print("\n\tPor favor verifique que los siguientes tres archivos estan en la carpeta en la que se encuentra y") 
    print("\tque estan actualizados a la última version:\n\n")
        
    existe_correos = os.path.exists(path_correos)
    existe_forms = os.path.exists(path_forms)
    existe_articulos = os.path.exists(path_articulos)  
    
    print("\t\tARCHIVO \t\t\tENCONTRADO \t\tÚLTIMA FECHA DE MODIFICACION\n")
    print("\t",ARCHIVO_CORREOS,end = "")
    if not existe_correos:
        print("NO".rjust(24),"-".rjust(30))
    else:
        print("SI".rjust(24), end ="")
        date = datetime.strptime( time.ctime(os.path.getmtime(path_correos) ), "%a %b %d %H:%M:%S %Y" )
        mostrar_fecha(date)
    print()
    
    print("\t",ARCHIVO_FORMS,end = "")
    if not existe_forms:
        print("NO".rjust(30),"-".rjust(30))
    else:
        print("SI".rjust(30), end = "")
        date = datetime.strptime( time.ctime(os.path.getmtime(path_forms) ), "%a %b %d %H:%M:%S %Y" )
        mostrar_fecha(date)
    print()
    
    print("\t",ARCHIVO_ARTICULOS,end = "")
    if not existe_articulos:
        print("NO".rjust(32),"-".rjust(30))
    else:
        print("SI".rjust(32),end = "")
        date = datetime.strptime( time.ctime(os.path.getmtime(path_articulos) ), "%a %b %d %H:%M:%S %Y" )
        mostrar_fecha(date)
        
    print("\n")


def obtener_mes(num_mes:int) ->str:
    """Obtiene el nombre del mes teniendo su numero

    Args:
        num_mes (int): El numero de mes del anio

    Returns:
        str: Nombre del mes correspondiente al numero
    """
    meses={
        1:"Enero",
        2:"Febrero",
        3:"Marzo",
        4:"Abril",
        5:"Mayo",
        6:"Junio",
        7:"Julio",
        8:"Agosto",
        9:"Septiembre",
        10:"Octubre",
        11:"Noviembre",       
        12:"Diciembre"    
    }
    nombre_mes = meses[num_mes]
    
    return nombre_mes

def mostrar_fecha(date: datetime)-> None:
    """Recibe una fecha en objeto datetime, lo castea a string y lo imprime por patalla en expresion horaria "argentina"

    Args:
        date (datetime): Una fecha determinada
    """
    
    date = str(date)
    dia = date[8:10]
    mes = date[5:7]
    num_mes = int(mes)
    nombre_mes = obtener_mes(num_mes)
    anio = date[0:4]
    fecha = (dia +" de " + nombre_mes + " de " + anio)
    print("\t\t\t",fecha)


def chequear_carga_articulos(dict_articulos: dict, path_articulos: str) -> bool:
    """chequea carga articulos

    Args:
        dict_ventas (dict): articuos y y su tipo
        path_correos (str): el path del archivo con articulos y su tipo articulo

    Returns:
        bool: true si estaba bien cargado inicialmete false en caso contrario
    """
    articulos_bien_cargados = cargar_dict_articulos(dict_articulos, path_articulos)
    if not articulos_bien_cargados:
        dict_articulos.clear()
        
        print("\n\n\tCorrija los errores y presione enter")
        time.sleep(2)
        os.startfile(path_articulos)
        
        input("\t")
        
    return articulos_bien_cargados


def chequear_carga_forms(dict_forms: dict, path_forms: str) -> bool:
    """chequea carga formularios

    Args:
        dict_ventas (dict): forms
        path_correos (str): el path del archuvo con tipo articulos y formularios

    Returns:
        bool: true si estaba bien cargado inicialmete false en caso contrario
    """
    formularios_bien_cargados = cargar_dict_forms(dict_forms, path_forms)
    if not formularios_bien_cargados:
        dict_forms.clear()
        
        print("\n\n\tCorrija los errores y presione enter")
        time.sleep(2)
        os.startfile(path_forms)
        
        input("\t")
        
    return formularios_bien_cargados


def chequear_carga_ventas(dict_ventas: dict, path_correos: str) -> bool:
    """chequea carga ventas

    Args:
        dict_ventas (dict): ventas
        path_correos (str): el path del archuvo con correos y ventas

    Returns:
        bool: true si estaba bien cargado inicialmete false en caso contrario
    """
    ventas_bien_cargadas = cargar_ventas(dict_ventas, path_correos)
    if not ventas_bien_cargadas:
        dict_ventas.clear()
        
        print("\n\n\tCorrija los errores y presione enter")
        time.sleep(2)
        os.startfile(path_correos)
        
        input("\t")
        
    return ventas_bien_cargadas


def iniciar_envio(dict_ventas: dict, dict_forms: dict, dict_articulos: dict) -> bool:
    """
    Devuelve true si el usuario acpeta inicar el envio. False en caso contrario. Chequea la existencia
    de los 3 archivos y la correcta lectura de los mismos.
    """
    #ruta_actual = "..\\"+os.getcwd()
    ruta_actual = os.path.normpath(os.getcwd() + os.sep + os.pardir)

    confirmar = False
    
    path_correos = ruta_actual + "\\" + ARCHIVO_CORREOS
    path_forms = ruta_actual + "\\" + ARCHIVO_FORMS
    path_articulos = ruta_actual + "\\" + ARCHIVO_ARTICULOS

    comenzar_envio = False    
    
    while ( not comenzar_envio ):
        
        mostrar_advertencia_existencia_archivos(path_correos,path_forms, path_articulos)
                
        existe_correos = os.path.exists(path_correos)
        existe_forms = os.path.exists(path_forms)
        existe_articulos = os.path.exists(path_articulos)  
        
        if ( existe_correos and existe_forms and existe_articulos):
            print("\tSe encontraron todos los archivos necesarios para la ejecucion de la campaña.\n") 
            print("\t¿Los archivos se encuentran actualizados a su última versión? Ingrese [ si ] o [ no ] en caso contrario.")            
            opc = input("\t-> ")
            print("\n")
                            
            if opc == "si":
                
                ventas_bien_cargadas = False
                formularios_bien_cargados = False
                articulos_bien_cargados = False
                
                if os.path.exists(path_correos):
                    ventas_bien_cargadas = chequear_carga_ventas(dict_ventas, path_correos)
                
                if os.path.exists(path_forms):
                    formularios_bien_cargados = chequear_carga_forms(dict_forms, path_forms)
                
                if os.path.exists(path_articulos): 
                    articulos_bien_cargados = chequear_carga_articulos(dict_articulos, path_articulos)
            
                comenzar_envio = ventas_bien_cargadas and formularios_bien_cargados and articulos_bien_cargados
                        
            else:
                print("\tActualice los archivos y presione la tecla enter")
                time.sleep(1.5)
                os.startfile(ruta_actual)
                
                actualizados = input("\t")
                            
        else:
            print("\tNo se encontro alguno de los archivos necesarios para la ejecucion de la campaña")
            print("\tPor favor, coloquelos en la carpeta en la que se encuentra y luego presione enter")
            time.sleep(2)
            os.startfile(ruta_actual)
            
            actualizados = input("\t")
            
        os.system("cls")
        time.sleep(1)

    time.sleep(1)
        
    print("\n\n\tIngrese [ si ] para iniciar el envio, o cualquier otra tecla para cancelarlo.")
    print("\n\n\t\tTENGA EN CUENTA QUE UNA VEZ INICIADO EL ENVIO EL MISMO NO PODRÁ SER INTERRUMPIDO\n\n")
    opc = input("\tINICIAR: ")
    if opc == "si":
        confirmar = True
    else:
        print("\n\tENVIO CANCELADO")
    print()
    
    return confirmar


def obtener_password()-> str:
    """Pide la contrasenia del correo al ussuario oculatndola mientras la escribe y la devuelve

    Returns:
        str: la contrasenia ingresada por el usuario
    """
    
    print("\n\tIngrese su contraseña y presione enter")
    password = input("\t-> ")
    
    return password


def obtener_correo() -> bool:
    """verifica que el correo electronico contenga el @ y
    que se @gmail.com emitiendo mensajes de error en cada caso

    Args:
        correo (str): coreo electronico del remitete

    Returns:
        bool: ]
    """
    correo = input("\t-> ")
    correcto = False
    while (not correcto):
        if correo[-10:] != "@gmail.com":
        #if correo[-10:] != "@empresa.com.ar":
            print("\n\t Por favor incluya '@gmail.com' en la direccion de correo electrónico\n")            
            correo = input("\t-> ")
        else:
            correcto = True
            
    return correo
            
            
def iniciar_sesion() -> smtplib.SMTP:
    
    """Permite un incio de sesion en google con usuario y contraseña y 
    devuelve n objeto con informacion del correo remitente

    Returns:
        [smtplib.SMTP]: Devuelve la variable mail con infomacion del correo remitente
    """
    
    print("\t\t\t\t\t\tINICIO DE SESION\n")
    
    inicio_correcto = False
    while (not inicio_correcto):
        
        print("\tIngrese su correo electronico INCLUYENDO @gmail.com.ar AL FINAL DEL MISMO y presione enter")
        
        correo = obtener_correo()
        password = obtener_password()

        print("\n\tIniciando sesión...")
        
        mail = smtplib.SMTP('smtp.gmail.com', 587)
        mail.ehlo()   
        mail.starttls()
        try:
            inicio_correcto =  mail.login(correo, password)
        except:
            print("\n\t\tEl correo electronico y / o la constraseña no son correctos\n")
    
    print("\n\t\t\t\t\tInicio de sesión exitoso!\n")
    time.sleep(2)
    os.system("cls")
    
    return mail
    

def main()-> None:
    """Script de envio automatico de correos electronicos con encuestas de satisfaccion correspondientes
    al tipo de articulo adquirido.
    """    
    print("\n\n\tBIENVENIDO AL PROGRAMA DE ENVIO AUTOMATICO DE ENCUESTAS POR CORREO ELECTRONICO DE 12SOFTWARE\n\n")
    
    correo_remitente = iniciar_sesion()
    
    dict_ventas = dict()
    dict_forms = dict()
    dict_articulos = dict()    
    
    if iniciar_envio(dict_ventas, dict_forms, dict_articulos):
                
        lista_advertencias = list()
        
        agregar_formularios(dict_ventas, dict_forms, dict_articulos, lista_advertencias)

        lista_bien_enviados = list() #cada elemento sera un correo enviado
        lista_no_enviados = list()
        lista_mal_enviados = list()

        inicio_envio =  datetime.now()
        enviar_correos(dict_ventas, lista_bien_enviados, lista_no_enviados, lista_mal_enviados, lista_advertencias, correo_remitente)
        fin_envio = datetime.now()

        tiempo_envio = fin_envio - inicio_envio    
        bien_enviados = len(lista_bien_enviados)
        no_enviados = len(lista_no_enviados)
        mal_enviados = len(lista_mal_enviados)
        total_enviados = bien_enviados + no_enviados + mal_enviados

        crear_archivo_resumen_correos(lista_bien_enviados, ARCHIVO_BIEN_ENVIADOS)
        crear_archivo_resumen_correos(lista_no_enviados, ARCHIVO_NO_ENVIADOS)
        crear_archivo_resumen_correos(lista_mal_enviados, ARCHIVO_MAL_ENVIADOS)
        crear_archivo_stats(inicio_envio, fin_envio, tiempo_envio, total_enviados, bien_enviados, no_enviados, mal_enviados)

        if lista_advertencias:
            print("\n\t_______________________________ADVERTENCIAS________________________________")
            for advertencia in lista_advertencias:
                print(advertencia)
                            
        archivar_archivos_de_registro_generados(no_enviados, mal_enviados,fin_envio)

        escribir_registro_historico(lista_bien_enviados, fin_envio)

        #cierre de sesion
        correo_remitente.quit()
        
        print("\n\tFIN DEL ENVIO\n")
                
        print("\tPara ver un informa mas detallado de su envío dirigase al registro correspondiente a la fecha y ")
        print("\thora de hoy de la carpeta 'registro_reportes'")
                
        print("\n\tAdios!")
    
    enter = input("\n\t\t\t\tPRESIONE ENTER PARA CERRAR EL PROGRAMA\n\t\t\t\t\t\t")


main()