def formatoD_L_CP_M(param):
    paramSplit = param.split("\n")
    direccion = paramSplit[0]
    localidad = paramSplit[1]
    cpMunicipio = paramSplit[2]
    cpMunicipioSplit = cpMunicipio.split(" - ")
    cp = cpMunicipioSplit[0]
    municipio = cpMunicipioSplit[1]
    return [direccion, localidad, cp, municipio]