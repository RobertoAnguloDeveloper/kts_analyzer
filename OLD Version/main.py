import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Cargar el archivo Excel
file_path = 'RGM-Fuel-and-Haulage-data-20-24.xlsx'
df = pd.read_excel(file_path)

# Preparar los datos (asumiendo que la primera columna contiene categorías y el resto son meses)
# Ajusta según la estructura real de tu archivo
df_transposed = df.set_index(df.columns[0]).T

# Convertir el índice a fechas si es necesario
df_transposed.index = pd.to_datetime(df_transposed.index)

# ----- GRÁFICA 1: Barras agrupadas para Ore Mined y Overburden -----
fig, ax = plt.subplots(figsize=(14, 6))

# Extraer las categorías de interés
ore_mined_rgm = df_transposed['Ore Mined']['RGM']
overburden_rgm = df_transposed['Overburden']['RGM']
ore_mined_sar = df_transposed['Ore Mined']['Sar']
overburden_sar = df_transposed['Overburden']['Sar']

# Configurar posiciones de las barras
x = np.arange(len(df_transposed.index))
width = 0.2

# Crear barras agrupadas
ax.bar(x - 1.5*width, ore_mined_rgm, width, label='Ore Mined RGM', color='#21808d')
ax.bar(x - 0.5*width, overburden_rgm, width, label='Overburden RGM', color='#5e5240')
ax.bar(x + 0.5*width, ore_mined_sar, width, label='Ore Mined Sar', color='#32b8c6')
ax.bar(x + 1.5*width, overburden_sar, width, label='Overburden Sar', color='#a77b2f')

# Configurar etiquetas y título
ax.set_xlabel('Mes', fontsize=12, fontweight='bold')
ax.set_ylabel('Kilotoneladas (kt)', fontsize=12, fontweight='bold')
ax.set_title('Ore Mined y Overburden por Mes (RGM y Sar)', fontsize=14, fontweight='bold')
ax.set_xticks(x)
ax.set_xticklabels(df_transposed.index.strftime('%Y-%m'), rotation=45, ha='right')
ax.legend()
ax.grid(axis='y', alpha=0.3)

plt.tight_layout()
plt.savefig('grafica_barras_agrupadas.png', dpi=300)
plt.show()

# ----- GRÁFICA 2: Líneas para Diesel y Flota Activa -----
fig, ax1 = plt.subplots(figsize=(14, 6))

# Eje izquierdo: Consumo de Diesel
diesel = df_transposed['Liter of Diesel Consumed']
color1 = '#21808d'
ax1.set_xlabel('Mes', fontsize=12, fontweight='bold')
ax1.set_ylabel('Litros de Diesel Consumido', fontsize=12, fontweight='bold', color=color1)
ax1.plot(df_transposed.index, diesel, color=color1, marker='o', linewidth=2, label='Diesel Consumido')
ax1.tick_params(axis='y', labelcolor=color1)
ax1.grid(alpha=0.3)

# Eje derecho: Flota Activa
ax2 = ax1.twinx()
fleet = df_transposed['Active Fleet Count (Aprox)']
color2 = '#a84b2f'
ax2.set_ylabel('Flota Activa (Aprox)', fontsize=12, fontweight='bold', color=color2)
ax2.plot(df_transposed.index, fleet, color=color2, marker='s', linewidth=2, label='Flota Activa')
ax2.tick_params(axis='y', labelcolor=color2)

# Título
plt.title('Consumo de Diesel y Flota Activa Mensual (2020-2024)', fontsize=14, fontweight='bold')

# Rotar etiquetas del eje x
plt.xticks(rotation=45, ha='right')

plt.tight_layout()
plt.savefig('grafica_lineas_diesel_flota.png', dpi=300)
plt.show()

# ----- GRÁFICA 3: Combinada (Barras + Línea) -----
fig, ax1 = plt.subplots(figsize=(14, 6))

# Barras: Total Ore Mined
total_ore_mined = ore_mined_rgm + ore_mined_sar
color1 = '#21808d'
ax1.bar(df_transposed.index, total_ore_mined, color=color1, alpha=0.7, label='Total Ore Mined')
ax1.set_xlabel('Mes', fontsize=12, fontweight='bold')
ax1.set_ylabel('Ore Mined (kt)', fontsize=12, fontweight='bold', color=color1)
ax1.tick_params(axis='y', labelcolor=color1)

# Línea: Diesel Consumido
ax2 = ax1.twinx()
color2 = '#c0152f'
ax2.plot(df_transposed.index, diesel, color=color2, marker='o', linewidth=2, label='Diesel Consumido')
ax2.set_ylabel('Litros de Diesel', fontsize=12, fontweight='bold', color=color2)
ax2.tick_params(axis='y', labelcolor=color2)

# Título
plt.title('Producción vs Consumo de Diesel', fontsize=14, fontweight='bold')
plt.xticks(rotation=45, ha='right')

# Leyendas
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left')

plt.tight_layout()
plt.savefig('grafica_combinada_produccion_diesel.png', dpi=300)
plt.show()

print("✅ Gráficas generadas exitosamente")
