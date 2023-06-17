#Algoritmo de lhun
class Tarjeta:
    def __init__(self, numero):
        self.numero = numero

    def ValidarTarjeta(self):
        tarjetaInvertida = self.numero[::-1]  # Se invierte el número de la tarjeta

        valorAValidar = 0
        for i in range(len(tarjetaInvertida)):
            digito = int(tarjetaInvertida[i])
            if i % 2 == 1:  # Multiplicar por 2 los dígitos en posición impar
                numeroMultiplicado = digito * 2
                if numeroMultiplicado > 9:
                    numeroMultiplicado = numeroMultiplicado - 9  # Restar 9 si el resultado es mayor a 9
                valorAValidar += numeroMultiplicado
            else:
                valorAValidar += digito

        if valorAValidar % 10 == 0:
            return "Tarjeta valida."
        else:
            return "Tarjeta no valida."




