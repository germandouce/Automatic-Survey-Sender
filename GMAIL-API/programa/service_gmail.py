import os

from googleapiclient.discovery import Resource, build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials


SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send'
]

ruta_actual = os.getcwd()

# Archivo generado para la API
ARCHIVO_SECRET_CLIENT = "client_secret.json"
PATH_TOKEN = ruta_actual+"\\"+"autenticacion\\token.json"
PATH_ARCHIVO_SECRET_CLIENT = ruta_actual+"\\"+"autenticacion\\client_secret.json"

def cargar_credenciales() -> Credentials:
    credencial = None
    
    if os.path.exists(PATH_TOKEN):
        with open(PATH_TOKEN, 'r'):
            credencial = Credentials.from_authorized_user_file(PATH_TOKEN, SCOPES)

    return credencial


def guardar_credenciales(credencial: Credentials) -> None:
    with open(PATH_TOKEN, 'w') as token:
        token.write(credencial.to_json())


def son_credenciales_invalidas(credencial: Credentials) -> bool:
    return not credencial or not credencial.valid


def son_credenciales_expiradas(credencial: Credentials) -> bool:
    return credencial and credencial.expired and credencial.refresh_token


def autorizar_credenciales() -> Credentials:
    print("""\n\tDebido a que es la primera vez que utiliza el programa debera AUTENTICARSE y habilitar a la aplicacion
    \tcon su cuenta de gmail.\n""")
    print("\tEste proceso deberá realizarlo una única vez puesto que el programa guardara su token de autenticación.\n")
    print("\tSe le mostrará un link por pantalla que deberá copiar y pegar en su navegador.")
    print("\tSi desea continuar, presione enter. Luego, siga las instrucciones que se le presentan por pantalla.\n")
    input("\t¿continuar?: ")
    print()
    flow = InstalledAppFlow.from_client_secrets_file(PATH_ARCHIVO_SECRET_CLIENT, SCOPES)

    return flow.run_local_server(open_browser=False, port=0)


def generar_credenciales() -> Credentials:
    credencial = cargar_credenciales()

    if son_credenciales_invalidas(credencial):

        if son_credenciales_expiradas(credencial):
            credencial.refresh(Request())

        else:
            if os.path.exists(PATH_ARCHIVO_SECRET_CLIENT):
                credencial = autorizar_credenciales()
        
        if not (credencial == None): #Si no es nula, es decir, la pudo crear o ya existia
            guardar_credenciales(credencial)

    return credencial


def obtener_servicio() -> Resource:
    """
    Creador de la conexion a la API Gmail
    """
    credentials=generar_credenciales()
    while (credentials == None):
        print("\n\tPOR FAVOR COLOQUE EL ACRHIVO 'client_secret.json' DENTRO DE LA CARPETA 'autenticacion', LUEGO PRESIONE ENTER\n")
        input("")
        credentials=generar_credenciales()

    return build('gmail', 'v1', credentials=generar_credenciales())

obtener_servicio()