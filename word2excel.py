"""
Author: Jesus Melendez @jmelendezgeo
Date last update : 29-04-2021
Description: Este codigo procesa los documentos de word (que siguen el formato especificado) en determinada ruta
            para guardar los registros a una forma mas estructurada (csv y xlsx). Aplica un primer flujo de limpieza de datos

    Ejemplo de registros en .docx:

    Claim Number:	  S0M3C0D3C3Details
    Claim Number Cross Reference:
    Name:	  PEDRO I PEREZ
    Birth Date:	  05/24/1949
    Date of Death:
    Sex:	  M
    Address:	  12345 112TH ST
      JAMAICA, NY  12345-6789
    Most recent State:	  NY (33)
    Most recent County:	  QUEENS (590)
            """


# Importamos librerias necesarias
import pandas as pd
import docx2txt
import re
import warnings
warnings.filterwarnings('ignore')
import os

def leer_documento(nombre_archivo):
    """Funcion que recibe el path de un archivo de word para analizarlo, buscar por un patron y retorna un DataFrame de pandas
    Con la informacion agrupada segun el patron """
    #Analizar el archivo de word y almacenarlo en un string
    my_text = docx2txt.process(nombre_archivo)

    pattern="""
    (Number\:\\t)(?P<ClaimNumber>.*)\\n{2}
    .*
    (Reference\:\\t)(?P<CrossReference>.*)\\n{2}
    .*
    (Name\:\\t)(?P<Name>.*)\\n{2}
    .*
    (Date\:\\t)(?P<BirthDate>.*)\\n{2}
    .*
    (Death\:\\t)(?P<DeathDate>.*)\\n{2}
    .*
    (Sex\:\\t)(?P<Sex>.*)\\n{2}
    .*
    (Address\:\\t)(?P<Address>.*\\n*.*)
    """

    lista=list()#Lista vacia para almacenar la informacion como diccionarios
    # Aplicamos busqueda de patron y agrupamos en una lista de diccionarios
    for item in re.finditer(pattern,my_text,re.VERBOSE):
        lista.append(item.groupdict())

    #Guardar diccionarios en un Dataframe
    df=pd.DataFrame(lista)
    return df


def remover_empty(df):
    """Funcion que recibe un DataFrame y aplica strip() para remover espacios en blanco en cada columna del DataFrame"""
    # Para cada columna del df, remover espacios en blanco al inicio, al final y guardar en Dataframe
    for column in df.columns:
        df[column]=df[column].str.strip()
    #Retornar DataFrame sin espacios blancos
    return df


def remover_strings(df):
    """Funcion que remueve strings identificados de la columna ClaimNumber que contiene los codigos de interes """
    # Lista con strings para reemplazar
    to_replace=['ROLL','ROL','Details','NO SIRVE','PASO','CAMRA','NO SIRVIO','CASA','RETRY']

    for element in to_replace: # Para cada string en la lista, remover y guardar en DataFrame
        df['ClaimNumber']=df['ClaimNumber'].str.replace(element,'')

    return df


def separar_columnas(df):
    """Funcion que recibe DataFrame y separa la informacion de ClaimNumber en dos columnas (Code1_mc y Code2_mi).
    Tambien separa Address en Direccion, Condado, Estado y Codigo postal"""

    codigos= df['ClaimNumber'].str.split(expand=True) #Separar ClaimNumber
    df['Code1_mc']=codigos[0] # Primer elemento es Code1_mc
    df['Code2_mi']=codigos[1] # Segundo elemento es Code2_mi. Si no hay, se guarda como vacio
    df.drop(columns='ClaimNumber',inplace= True) # Remover columna ClaimNumber

    # Separar Direccion
    new= df['Address'].str.split('\n\n',expand=True) #La separacion es por los saltos de linea
    df['Direccion'] = new[0]

    # Tomar Condado
    df['Condado'] = new[1].str.split(',',expand=True)[0] # La separacion es por coma ','
    # Tomar Estado y Zipcode
    new=new[1].str.split(',',expand=True)
    df['Estado']=new[1].str.split(' ',expand=True)[1] #Estado es columna 1
    df['ZipCode']= new[1].str.split(' ',expand=True)[3] # Zipcode es columna 3
    df.drop(columns='Address',inplace=True)

    return df


def depurar_datos(df):
    """Funcion que reemplaza vacios con None, remueve registros con informacion en DeathDate y elimina columnas DeathDate y CrossReference """
    for column in df.columns:
        df[column]=df[column].replace([''],[None])
    df.drop(df[~df['DeathDate'].isnull()].index, inplace=True)
    df.drop(columns=['CrossReference','DeathDate'],inplace=True)

    return df


def depurar_codigos(df):
    """Funcion que depura codigos de Code1_mc y Code2_mi que no fueron correctamente separados """
    # Crear DataFrame con los registros que tengan un Code1_mc distinto a 10 caracteres, Code1_mc >= 18 caracteres
    df_test = df[(df['Code1_mc'].str.len()!=10) | (df['Code2_mi'].str.len()!=8)]
    # Localizar registros con >= 18 caracteres y asigna los 9-17 al Code2_mi
    df_test.loc[df_test['Code1_mc'].str.len()>=18,'Code2_mi']=df_test['Code1_mc'].str.slice(start=10,stop=18)
    # Reemplazar Code1_mc con los caracteres 0-9 (Los 10 primeros)
    df_test.loc[df_test['Code1_mc'].str.len()>10,'Code1_mc']=df_test['Code1_mc'].str.slice(stop=10)
    # En los Code2_mi mas largos de 8 caracteres, reemplazar Code2_mi con los primeros 8 caracteres
    df_test.loc[df_test['Code2_mi'].str.len()>8,'Code2_mi']=df_test['Code2_mi'].str.slice(stop=8)
    df.update(df_test)

    return df


def guardar(df_limpio):
    """Funcion que guarda el DataFrame recibido en un archivo de excel y en un archivo csv """
    df_limpio.to_excel('nydb.xlsx',sheet_name='test1')
    df_limpio.to_csv('nydb.csv')


def limpieza_datos(df):
    """Funcion que controla el flujo de limpieza de datos """
    df=remover_empty(df)
    df=remover_strings(df)
    df=separar_columnas(df)
    df=depurar_datos(df)
    df=depurar_codigos(df)

    return df


def main(mypath):
    """Funcion main que controla el flujo de lectura de documentos, limpieza de datos y guardado """

    # Creamos DataFrame vacio para concatenar la informacion resultante de leer_documento
    df=pd.DataFrame(columns=['ClaimNumber','CrossReference','Name','BirthDate','DeathDate','Sex','Address'])

    #Para cada archivo que se encuentre en mypath, leer documentos y concatenar registros en df
    for root,dirs,files in os.walk(mypath):
        for filename in files:
            df=pd.concat([df,leer_documento(os.path.join(root,filename))],ignore_index=True)

    # Aplicar flujo de limpieza
    df_limpio=limpieza_datos(df)

    #Guardar registro limpio
    guardar(df_limpio)

    #Imprimimos alguna informacion de interes.
    print(f"Terminamos. En total tuvimos {len(files)} documentos.")
    print(f"Esto nos dio un total de {len(df_limpio)} registros.")
    print(f"Ya hemos procesado los siguientes archivos: {files}")



if __name__ == '__main__':
    #Definir Ruta
    mypath="PATH"

    main(mypath)
