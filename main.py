import pandas as pd
from datetime import datetime

from imports.formatoTdC_Nombre_CC import formatoTdC_Nombre_CC
from imports.formatoD_L_CP_M import formatoD_L_CP_M
from imports.formatoDirector import formatoDirector
from imports.formatoPersonaContacto import formatoPersonaContacto

# Abrir Excel
df = pd.read_excel('./excels/copia.xlsx',  skiprows=2, engine="openpyxl")

# Ajusto los indices para que los indices del array coincidan con los indices de las filas del excel
compensacionArray = 4
cambioFormato = 461 - compensacionArray
##
arrayFilas = []

ahora = datetime.now()
ahoraStr = str(ahora)


for i in range(0, cambioFormato):
    arrayFilas.append(df.iloc[i])

arrayFilasFinal = []
arrayFilasContactos = []





# Iterar sobre las filas del excel
for i in arrayFilas:
    print(f"\nProcesando fila {i.name}")
    # division campos TdP, Nombre, CC
    if isinstance(i["Tipo de Centro\nNombre\nCódigo Centro"], str):
        if str(i["Tipo de Centro\nNombre\nCódigo Centro"])[-1].isdigit():
            tipoCentro, nombre, codigoCentro = formatoTdC_Nombre_CC(i["Tipo de Centro\nNombre\nCódigo Centro"]) ##OUTPUT: Array=[Tipo de Centro, Nombre, Código Centro]
    else:
        tipoCentro = ""
        nombre = ""
        codigoCentro = ""
        # campo email 
    if isinstance(i["e-mail I"], str):
        email= str(i["e-mail I"])
    else:
        email = ""
    #division campos direccion Localidad Provincia Cp
    if isinstance(i["Dirección\nLocalidad\nC.P. - Municipio"] , str):
        direccion, localidad, cp, municipio = formatoD_L_CP_M(i["Dirección\nLocalidad\nC.P. - Municipio"]) ##OUTPUT: Array=[Direccion, Localidad, CP, Municipio]
    else: 
        direccion = ""
        localidad = ""
        cp = ""
        municipio = ""
    #campo telefono
    if isinstance(i["Teléfonos\nFax"], str):
        tlfnRaw = i["Teléfonos\nFax"].replace("\n", "")
        posicion = tlfnRaw.find("Fax")
        if posicion != -1:
            tlfn = f"{tlfnRaw[:posicion]}  "
            fax = tlfnRaw[posicion:]
        else:
            tlfn = tlfnRaw
            fax = ""
    else:
        tlfn = ""
        fax = ""      
    #Campo Director de Contactos
    if isinstance(i["Director/a\ne-Mail"], str):
        director, emailDirector = formatoDirector(i["Director/a\ne-Mail"])
    else:    
        director = ""
        emailDirector = ""

    #Campo Persona de contacto
    if isinstance(i["Persona de contacto"], str):
        arrayContactos = formatoPersonaContacto(i["Persona de contacto"]) ##OUTPUT: Array=[{Nombre, Cargo}]
    else:
        contacto = ""
 
    #Elimino los campos de los que ya he sacado info para iterar sobre los campos restantes y crear el campo "Notas internas"
    del i["Tipo de Centro\nNombre\nCódigo Centro"]
    del i["e-mail I"]
    del i["Dirección\nLocalidad\nC.P. - Municipio"]
    del i["Teléfonos\nFax"]
    del i["Unnamed: 0"]
    del i["Director/a\ne-Mail"]
    del i["Persona de contacto"]

    #Creo el campo "Notas internas"
    notasInternas = []

    notasInternas.append(f"Codigo de Centro    ----->    {codigoCentro}")
    if fax != "":
        notasInternas.append(f"Fax    ----->    {fax}")
    for clave, valor in i.items():
        valorSinSaltos = str(valor).replace("\n", " ")
        claveValor = f"{clave}    ----->    {valorSinSaltos}"
        notasInternas.append(claveValor)
    
    strNotasInternas = "\n".join(notasInternas)


    fila = {
        "Nombre" : nombre , 
        "Dirección 1" : direccion,
        "Dirección 2" : "",
        "Código Postal" : cp,
        "Ciudad/Localidad" : localidad,
        "Provincia" : municipio,
        "País" : "", 
        "NIF" : "",
        "Tipo de Centro" : tipoCentro,
        "Teléfono" : tlfn,
        "Móvil" : "",
        "Correo Electrónico" : email,
        "Sitio Web" : "",
        "Notas internas" :  strNotasInternas
        }
    if fila["Nombre"] == "":
        continue

    arrayFilasFinal.append(fila)

    filaContactoDirector = {
        "Nombre Centro":nombre,
        "Nombre":director,
        "email":emailDirector,
        "Cargo":"Director/a"
    }

    arrayFilasContactos.append(filaContactoDirector)

    for contacto in arrayContactos:
        if len(contacto) < 1:
            contacto[1] = ""
        filaContacto = {
            "Nombre Centro":nombre,
            "Nombre":contacto[0],
            "Cargo":contacto[1]
        }   

        arrayFilasContactos.append(filaContacto)

    
    

#Excel contactos
dfExcel = pd.DataFrame(arrayFilasContactos)
dfExcel.to_excel(f"excelFormatoContactos.xlsx", index=False)

#Excel contenido

dfExcel = pd.DataFrame(arrayFilasFinal)
dfExcel.to_excel(f"excelFormato.xlsx", index=False)

print("Excel creado exitosamente")