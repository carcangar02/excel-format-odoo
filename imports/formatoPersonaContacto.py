import re
def formatoPersonaContacto(param):

    contactos = re.findall(r'([\w\s]+)\s\(([\w\s]+)\)', param)
    resultado = [[palabra, parentesis] if parentesis else [palabra] for palabra, parentesis in contactos]
    
    print(f"Elementos: {resultado}")


    return contactos
