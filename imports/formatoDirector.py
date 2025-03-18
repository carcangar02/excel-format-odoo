import re

def formatoDirector(param):

    partes = re.split(r'(?=[a-z])', param, maxsplit=1)  


    #solo email
    if partes[0] == "":
            emailDirectorRaw = re.sub(r'[;,]', '', partes[1])
            if partes[1].count("@") > 1:
                emailDirector = emailDirectorRaw.replace("\n", ", ")
            else:
                emailDirector = emailDirectorRaw.replace("\n", "")
    else:
        if len(partes) == 2:
            if "@" in partes[1]:
                emailDirectorRaw = re.sub(r'[;,]', '', partes[1])
                if partes[1].count("@") > 1:
                    emailDirector = emailDirectorRaw.replace("\n", ", ")
                else:
                    emailDirector = emailDirectorRaw.replace("\n", "")
            else:
                partes[0] = f"{partes[0]}{partes[1]}"
                emailDirector = ""
        else:
            emailDirector = ""
    director = partes[0]

    return [director, emailDirector]


