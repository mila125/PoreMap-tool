import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import math

def leer_dimension_poro_desde_excel(ruta_excel, hoja="BJH", columna="Pore Width (nm)"):
    try:
        df = pd.read_excel(ruta_excel, sheet_name=hoja)
        valores = df[columna].dropna().tolist()
        if not valores:
            raise ValueError("No se encontraron valores válidos en la columna.")
        return valores[0]  # solo el primero para visualizar
    except Exception as e:
        print(f" Error al leer Excel: {e}")
        return 5  # valor por defecto en nm


def cesaro_fractal(ax, p1, p2, nivel):
    if nivel == 0:
        ax.plot([p1[0], p2[0]], [p1[1], p2[1]], color='blue')
    else:
        dx = p2[0] - p1[0]
        dy = p2[1] - p1[1]

        # puntos intermedios
        A = (p1[0] + dx / 3, p1[1] + dy / 3)
        B = (p1[0] + dx * 2 / 3, p1[1] + dy * 2 / 3)

        # punto pico (ángulo Cesàro: 60°)
        angle = math.atan2(dy, dx) - math.pi / 3
        length = math.hypot(dx, dy) / 3
        C = (A[0] + length * math.cos(angle), A[1] + length * math.sin(angle))

        cesaro_fractal(ax, p1, A, nivel - 1)
        cesaro_fractal(ax, A, C, nivel - 1)
        cesaro_fractal(ax, C, B, nivel - 1)
        cesaro_fractal(ax, B, p2, nivel - 1)


def visualizar_poro_fractal(ruta_excel, hoja="BJH", columna="Pore Width (nm)", nivel=3):
    ancho_poro_nm = leer_dimension_poro_desde_excel(ruta_excel, hoja, columna)
    escala = ancho_poro_nm  # puedes escalar a gusto

    fig, ax = plt.subplots(figsize=(8, 2))
    p1 = (0, 0)
    p2 = (escala, 0)

    cesaro_fractal(ax, p1, p2, nivel)
    ax.set_aspect('equal')
    ax.set_title(f"Fractal tipo Cesàro (simulación de poro: {ancho_poro_nm:.2f} nm)")
    ax.axis('off')
    plt.show()