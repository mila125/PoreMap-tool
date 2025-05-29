import numpy as np  # Para operaciones numéricas avanzadas (opcional)
import matplotlib.pyplot as plt  # Para graficar la secuencia (opcional)
fibonacci_dual = []  # Array global para almacenar la secuencia combinada

def fibonacci_dual(max_x, max_y):
    """
    Genera dos secuencias de Fibonacci dentro de un mismo array.
    - Elementos en índices pares corresponden a la distancia en X.
    - Elementos en índices impares corresponden a la distancia en Y.

    :param max_x: Número máximo de términos en la secuencia X.
    :param max_y: Número máximo de términos en la secuencia Y.
    :return: Lista con ambas secuencias intercaladas.
    """
    def fibonacci(n):
        """Genera la secuencia de Fibonacci hasta n términos."""
        if n <= 0:
            return []
        fib = [0, 1]
        for _ in range(2, n):
            fib.append(fib[-1] + fib[-2])
        return fib[:n]

    # Obtener ambas secuencias
    seq_x = fibonacci(max_x)
    seq_y = fibonacci(max_y)

    # Intercalar ambas secuencias en un solo array
    resultado = []
    i, j = 0, 0

    while i < len(seq_x) or j < len(seq_y):
        if i < len(seq_x):
            resultado.append(seq_x[i])  # Elemento para X (posición par)
            i += 1
        if j < len(seq_y):
            resultado.append(seq_y[j])  # Elemento para Y (posición impar)
            j += 1

    return resultado

# Ejemplo de uso
max_x = 5  # Número máximo de términos para X
max_y = 6  # Número máximo de términos para Y
secuencia = fibonacci_dual(max_x, max_y)
print(secuencia)