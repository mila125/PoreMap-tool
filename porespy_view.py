import pandas as pd

def visualizar_poro_promedio(tamano_nm):
    import matplotlib.pyplot as plt

    radio = tamano_nm / 2
    escala = 2
    radio_pix = radio * escala

    fig, ax = plt.subplots()
    ax.set_aspect('equal')
    circulo = plt.Circle((0, 0), radio_pix, color='skyblue', edgecolor='black')
    ax.add_artist(circulo)

    ax.set_xlim(-radio_pix * 1.2, radio_pix * 1.2)
    ax.set_ylim(-radio_pix * 1.2, radio_pix * 1.2)
    ax.set_title(f"Poro circular (Radio ~ {radio:.2f} nm)")
    plt.axis('off')
    plt.show()

def visualizar_poro_con_porespy(ruta_excel, hoja_excel):
    try:
        df = pd.read_excel(ruta_excel, sheet_name=hoja_excel)

        # Detectar automáticamente la columna de radios o diámetros
        posibles_columnas = [col for col in df.columns if 'diam' in col.lower() or 'radius' in col.lower()]
        if not posibles_columnas:
            print(" Columna de radio o diámetro no encontrada.")
            return

        columna = posibles_columnas[0]  # Usamos la primera coincidencia
        print(f" Usando columna: {columna}")

        valores = df[columna].dropna()

        for tamano in valores:
            visualizar_poro_promedio(tamano)

    except Exception as e:
        print(f" Error en visualización con porespy: {e}")