#Graficos a partir de los dataframes

import openpyxl
import os
import csv
import traceback
from datetime import datetime
from pywinauto import Application, findwindows
import time
import threading
from methods_to_df import generar_nombre_unico
from novawinmng import manejar_novawin
from pywinauto.keyboard import send_keys
from openpyxl import Workbook
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import subprocess
from openpyxl import load_workbook
from pandas import ExcelWriter
from openpyxl.drawing.image import Image

def draw_comparison_bar_chart(bjhd, bjha):
    # Filtrar columnas necesarias
    radius_bjhd = bjhd['Radius']
    dV_logr_bjhd = bjhd['dV(logr)']
    
    radius_bjha = bjha['Radius']
    dV_logr_bjha = bjha['dV(logr)']

    # Asegurar mismo tamaño
    min_length = min(len(radius_bjhd), len(radius_bjha))
    radius_bjhd = radius_bjhd[:min_length]
    dV_logr_bjhd = dV_logr_bjhd[:min_length]
    radius_bjha = radius_bjha[:min_length]
    dV_logr_bjha = dV_logr_bjha[:min_length]

    # === Gráfico de desorción (BJHD) ===
    plt.figure(figsize=(10, 5))
    plt.bar(radius_bjhd, dV_logr_bjhd, color='blue', edgecolor='black', alpha=0.7)
    plt.xlabel('Radius (Å)')
    plt.ylabel('dV(logr)')
    plt.title('BJH Desorción')
    plt.xticks(rotation=90)
    plt.grid(axis='y', linestyle='--', alpha=0.6)
    img_bjhd_path = "grafico_bjhd.png"
    plt.tight_layout()
    plt.savefig(img_bjhd_path, dpi=300)
    plt.close()

    # === Gráfico de adsorción (BJHA) ===
    plt.figure(figsize=(10, 5))
    plt.bar(radius_bjha, dV_logr_bjha, color='green', edgecolor='black', alpha=0.7)
    plt.xlabel('Radius (Å)')
    plt.ylabel('dV(logr)')
    plt.title('BJH Adsorción')
    plt.xticks(rotation=90)
    plt.grid(axis='y', linestyle='--', alpha=0.6)
    img_bjha_path = "grafico_bjha.png"
    plt.tight_layout()
    plt.savefig(img_bjha_path, dpi=300)
    plt.close()

    print(f"Gráficos BJH guardados en: {img_bjhd_path} y {img_bjha_path}")
    return img_bjhd_path, img_bjha_path

def draw_DFT(df):
    half_pore_width = pd.to_numeric(df['Half pore width'], errors='coerce')
    dVr = df['dV(r)']
    
    valid_indices = ~half_pore_width.isna()
    dVr = dVr[valid_indices]
    half_pore_width = half_pore_width[valid_indices]

    plt.figure(figsize=(12, 6))
    plt.bar(half_pore_width, dVr, width=0.8, color='green', alpha=0.7, label='dV(r)')
    
    plt.xlabel('Half pore width (Å)')
    plt.ylabel('dV(r) (cc/Å/g)')
    plt.title('DFT Analysis: Half Pore Width vs dV(r)')
    plt.legend()
    plt.grid(axis='y', linestyle='--', alpha=0.6)

    image_path = "grafico_dft.png"
    plt.tight_layout()
    plt.savefig(image_path, dpi=300)
    plt.close()

    print(f"Gráfico DFT guardado en: {image_path}")
    return image_path
def draw_HK(df):
    # Filtrar y convertir las columnas necesarias
    dVr = df['dV()']
    half_pore_width = pd.to_numeric(df['Half pore width'], errors='coerce')

    # Eliminar filas no válidas
    valid_indices = ~half_pore_width.isna()
    dVr = dVr[valid_indices]
    half_pore_width = half_pore_width[valid_indices]

    # Crear el gráfico de barras
    plt.figure(figsize=(14, 6))
    bar_width = 0.8
    plt.bar(half_pore_width, dVr, width=bar_width, color='green', alpha=0.7, label='dV(r)')

    # Etiquetas y título
    plt.xlabel('Half pore width , A')
    plt.ylabel('dV cc A g')
    plt.title('Gráfico HK')
    plt.xticks(half_pore_width, rotation=90)
    plt.legend()

    # Guardar la imagen
    temp_image_path = "grafico_hk.png"
    plt.tight_layout()
    plt.savefig(temp_image_path, dpi=300)
    plt.close()

    print(f"Imagen del gráfico HK guardada en: {temp_image_path}")
    return temp_image_path


