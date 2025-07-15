texto = ".0"
if(texto[-1] == "0" and texto[-2] == "."):
    texto = texto[0:-2]

print(texto)

if(texto.isnumeric()):
    print("sexo")