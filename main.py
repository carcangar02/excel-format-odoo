import pandas as pd


from imports.formatoTdC_Nombre_CC import formatoTdC_Nombre_CC
from imports.formatoD_L_CP_M import formatoD_L_CP_M
from imports.formatoDirector import formatoDirector
from imports.formatoPersonaContacto import formatoPersonaContacto
from imports.formatoTelefono import formatoTelefono

# Abrir Excel
df = pd.read_excel('./excels/copia.xlsx',  skiprows=2, engine="openpyxl")

# Ajusto los indices para que los indices del array coincidan con los indices de las filas del excel
compensacionArray = 4
# cambioFormato = 461 - compensacionArray
cambioFormato = len(df)
##
arrayFilas = []



for i in range(0, cambioFormato):
    arrayFilas.append(df.iloc[i])

arrayFilasFinal = []
arrayFilasContactos = []
emailRepetidos = []





# Iterar sobre las filas del excel
for i in arrayFilas:
    activador = False
    if (i["Unnamed: 0"]==1.0):
        activador = True
        


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
        if email == "#VALUE!":
            email = ""
    else:
        email = ""


    #division campos direccion Localidad Provincia Cp
    if activador==True:
        horario = f"{i['Dirección\nLocalidad\nC.P. - Municipio']}"
        paNotas = i['Horario']
    else:
        try:

            direccion, localidad, cp, municipio = formatoD_L_CP_M(i["Dirección\nLocalidad\nC.P. - Municipio"]) ##OUTPUT: Array=[Direccion, Localidad, CP, Municipio]
            localidad = localidad[0] + localidad[1:].lower()
            municipio = municipio[0] + municipio[1:].lower()
            direccion = f'{direccion}, {localidad}'

        except:
                direccion = ""
                localidad = ""
                cp = ""
                municipio = ""

        
        






    #campo telefono
    if isinstance(i["Teléfonos\nFax"], str):
        tlfnRaw = i["Teléfonos\nFax"].replace("\n", "")
        posicion = tlfnRaw.find("Fax")
        if posicion != -1:
            tlfnCombinado = f"{tlfnRaw[:posicion]} "
            fax = tlfnRaw[posicion:]
            tlfn ,movil =formatoTelefono(tlfnCombinado)
        else:
            tlfn ,movil =formatoTelefono(tlfnRaw)
            fax = ""
    else:
        tlfn = ""
        movil = ""
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
    del i["Horario"]
    del i["Director/a\ne-Mail"]
    del i["Persona de contacto"]

    #Creo el campo "Notas internas"
    notasInternas = []

    notasInternas.append(f"Codigo de Centro    ----->    {codigoCentro}")
    if activador==True:
        notasInternas.append(f"Observaciones    ----->    {paNotas}")
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
        "Ciudad/Localidad" : municipio,
        "Provincia" : "Cantabria",
        "País" : "", 
        "NIF" : "",
        "Tipo de Centro" : tipoCentro,
        "Teléfono" : tlfn,
        "Móvil" : movil,
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
    if filaContactoDirector["Nombre"]=="":
        print(filaContactoDirector)

    if filaContactoDirector["Nombre"] != "" and filaContactoDirector["email"] != "":
        arrayFilasContactos.append(filaContactoDirector)

    for contacto in arrayContactos:
        if len(contacto) < 1:
            contacto[1] = ""
        filaContacto = {
            "Nombre Centro":nombre,
            "Nombre":contacto[0],
            "email":"",
            "Cargo":contacto[1]
        }   
        if filaContacto["Nombre"] != "" and filaContacto["Cargo"] != "":
            arrayFilasContactos.append(filaContacto)

        if filaContactoDirector["Nombre"]=="" and filaContactoDirector["email"]!="":
            emailRepetidos.append(filaContactoDirector)

arrayFilasContactosSin_duplicados = []
valores_vistos = set()
valores_vistosEmail = set()

for d in arrayFilasContactos:
        clave_unica = (d["Nombre"], d["Cargo"],d["email"])  # Solo Nombre y Cargo

        if clave_unica not in valores_vistos:
            valores_vistos.add(clave_unica)
            arrayFilasContactosSin_duplicados.append(d)

for d in emailRepetidos:
        clave_unica2 = (d["Nombre Centro"],d["email"])  # Solo Nombre y Cargo

        if clave_unica2 not in valores_vistosEmail:
            valores_vistosEmail.add(clave_unica2)
            arrayFilasContactosSin_duplicados.append(d)



        

      





    
    

#Excel contactos

dfExcel = pd.DataFrame(arrayFilasContactosSin_duplicados)
dfExcel.to_excel(f"excelFormatoContactos.xlsx", index=False)

#Excel contenido

dfExcel = pd.DataFrame(arrayFilasFinal)
dfExcel.to_excel(f"excelFormato.xlsx", index=False)

print("Excel creado exitosamente")