def formatoTdC_Nombre_CC(param):
    paramSplit = param.split("\n")
    tipoCentro = paramSplit[0]
    nombre = paramSplit[1]
    codigoCentro = paramSplit[2]
    return [tipoCentro, nombre, codigoCentro]