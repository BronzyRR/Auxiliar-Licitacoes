from num2words import num2words

def retornarStringItem(*args):
    resultado = ""
    for item in args:
        if(item != 0):
            resultado += "{}.".format(item)

    resultado = resultado[0:len(resultado)-1]
    return resultado

def number_to_long_number(number_p):
    if number_p.find(',') != -1:
        number_p = number_p.split(',')
        number_p1 = int(number_p[0].replace('.', ''))
        number_p2 = int(number_p[1])
    else:
        number_p1 = int(number_p.replace('.', ''))
        number_p2 = 0

    if number_p1 == 1:
        aux1 = ' real'
    else:
        aux1 = ' reais'

    if number_p2 == 1:
        aux2 = ' centavo'
    else:
        aux2 = ' centavos'

    text1 = ''
    if number_p1 > 0:
        text1 = num2words(number_p1, lang='pt_BR') + str(aux1)
    else:
        text1 = ''

    if number_p2 > 0:
        text2 = num2words(number_p2, lang='pt_BR') + str(aux2)
    else:
        text2 = ''

    if (number_p1 > 0 and number_p2 > 0):
        result = text1 + ' e ' + text2
    else:
        result = text1 + text2

    return result

def adequar_string(numero):
    s = list(numero)

    for i in range(len(s)):
        if s[i] == ".":
            s[i] = ","
        elif s[i] == ",":
            s[i] = "."

    resultado = "".join(s)
    return resultado

def capitalizar_letras(numero):
    s = list(numero)

    if(len(s) != 0):
        s[0] = s[0].upper()

        for i in range(1, len(s)):
            if s[i-1] == " " and s[i+1] != " ":
                s[i] = s[i].upper()
    else:
        return

    resultado = "".join(s)
    return resultado

def adicionar_pontos(numero):
    s = list(numero)
    s.reverse()

    r = []
    contador = 0

    #o primeiro passo é verificar os dois primeiros caracteres
    if(s[2] == ","):
        r.append(s[0])
        r.append(s[1])
        r.append(s[2])

        for i in range(3, len(s)):
            contador += 1
            r.append(s[i])
            if(contador == 3):
                r.append(".")
    else:
        r.append("0")
        r.append("0")
        r.append(",")

        for i in range(len(s)):
            contador += 1
            r.append(s[i])

            if(contador == 3):
                r.append(".")

    r.reverse()
    resultado = "".join(r)
    return resultado


