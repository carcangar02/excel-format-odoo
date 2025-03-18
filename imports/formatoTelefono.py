def formatoTelefono(telefono):
    if telefono[:9].find(" ") != -1:
        corte = telefono.find("942") + 12
    else: 
        corte = telefono.find("942") + 10

    tlfn = telefono[:corte].replace(" ", "")
    if len(tlfn) == 10:
        movil = tlfn[11:] +telefono[corte:].replace(" ", "") 
        movil = telefono[corte:].replace(" ", "")
    else:
        movil = ""




    return [tlfn, movil]