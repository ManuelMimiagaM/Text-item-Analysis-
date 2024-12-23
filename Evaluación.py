# -*- coding: utf-8 -*-
"""
Created on Wed Jun 19 11:29:07 2024

@author: Juan Manuel Mimiaga Morales
"""
import pandas as pd


# Ruta del archivo Excel para Reactivos
archivo_excel_reactivos = r"C:\Users\UNISA_96\Documents\Analisis PAI\PAI24-M7. Actividad sumativa-respuestas Miguel Zúñiga González.xlsx"
# Leer el archivo Excel de Reactivos
Reactivos = pd.read_excel(archivo_excel_reactivos)

# Crear la nueva columna "Pregunta_1" en Reactivos
Reactivos['P_1'] = Reactivos['Pregunta 1'].str.split(':').str[0].str.strip()
Reactivos['P_2'] = Reactivos['Pregunta 2'].str.split(':').str[0].str.strip()
Reactivos['P_3'] = Reactivos['Pregunta 3'].str.split(':').str[0].str.strip()
Reactivos['P_4'] = Reactivos['Pregunta 4'].str.split(':').str[0].str.strip()
Reactivos['P_5'] = Reactivos['Pregunta 5'].str.split(':').str[0].str.strip()
Reactivos['P_6'] = Reactivos['Pregunta 6'].str.split(':').str[0].str.strip()
Reactivos['P_7'] = Reactivos['Pregunta 7'].str.split(':').str[0].str.strip()
Reactivos['P_8'] = Reactivos['Pregunta 8'].str.split(':').str[0].str.strip()
Reactivos['P_9'] = Reactivos['Pregunta 9'].str.split(':').str[0].str.strip()
Reactivos['P_10'] = Reactivos['Pregunta 10'].str.split(':').str[0].str.strip()

ReactivosS = Reactivos[['Nombre', 'P_1', 'Respuesta 1', 
                                           'P_2','Respuesta 2', 
                                           'P_3','Respuesta 3',  
                                           'P_4','Respuesta 4', 
                                           'P_5','Respuesta 5', 
                                           'P_6','Respuesta 6', 
                                           'P_7','Respuesta 7', 
                                           'P_8','Respuesta 8', 
                                           'P_9','Respuesta 9', 
                                           'P_10','Respuesta 10']]

# Guardar el DataFrame resultante en un nuevo archivo Excel
archivo_salida = r'C:/Users/UNISA_96/Documents/Analisis PAI/Preguntas_7.xlsx'
ReactivosS.to_excel(archivo_salida, index=False)

# Ruta del archivo Excel para CVE Preguntas
Hoja1 = 'Hoja1'  # Nombre de la hoja que quieres leer
Cve_preguntas = pd.read_excel(r"C:\Users\UNISA_96\Documents\Analisis PAI\CVE_Preguntas-M7.xlsx", sheet_name='Hoja1')

Hoja2 = 'Hoja2'  # Nombre de la hoja que quieres leer
Cve_preguntas2 = pd.read_excel(r"C:\Users\UNISA_96\Documents\Analisis PAI\CVE_Preguntas-M7.xlsx", sheet_name='Hoja2')

Hoja3 = 'Hoja3'  # Nombre de la hoja que quieres leer
Cve_preguntas3 = pd.read_excel(r"C:\Users\UNISA_96\Documents\Analisis PAI\CVE_Preguntas-M7.xlsx", sheet_name='Hoja3')

Hoja4 = 'Hoja4'  # Nombre de la hoja que quieres leer
Cve_preguntas4 = pd.read_excel(r"C:\Users\UNISA_96\Documents\Analisis PAI\CVE_Preguntas-M7.xlsx", sheet_name='Hoja4')

Hoja5 = 'Hoja5'  # Nombre de la hoja que quieres leer
Cve_preguntas5 = pd.read_excel(r"C:\Users\UNISA_96\Documents\Analisis PAI\CVE_Preguntas-M7.xlsx", sheet_name='Hoja5')

Hoja6 = 'Hoja6'  # Nombre de la hoja que quieres leer
Cve_preguntas6 = pd.read_excel(r"C:\Users\UNISA_96\Documents\Analisis PAI\CVE_Preguntas-M7.xlsx", sheet_name='Hoja6')

Hoja7 = 'Hoja7'  # Nombre de la hoja que quieres leer
Cve_preguntas7 = pd.read_excel(r"C:\Users\UNISA_96\Documents\Analisis PAI\CVE_Preguntas-M7.xlsx", sheet_name='Hoja7')

Hoja8 = 'Hoja8'  # Nombre de la hoja que quieres leer
Cve_preguntas8 = pd.read_excel(r"C:\Users\UNISA_96\Documents\Analisis PAI\CVE_Preguntas-M7.xlsx", sheet_name='Hoja8')

Hoja9 = 'Hoja9'  # Nombre de la hoja que quieres leer
Cve_preguntas9 = pd.read_excel(r"C:\Users\UNISA_96\Documents\Analisis PAI\CVE_Preguntas-M7.xlsx", sheet_name='Hoja9')

Hoja10 = 'Hoja10'  # Nombre de la hoja que quieres leer
Cve_preguntas10 = pd.read_excel(r"C:\Users\UNISA_96\Documents\Analisis PAI\CVE_Preguntas-M7.xlsx", sheet_name='Hoja10')

# Realizar el merge basado en la columna 'Pregunta_1' de Cve_preguntas y 'Pregunta_1' de Reactivos
resultado_merge = pd.merge(Cve_preguntas, Reactivos, on='P_1', how='outer')
resultado_merge = pd.merge(Cve_preguntas2, resultado_merge, left_on='P_2', right_on='P_2', how='outer')
resultado_merge = pd.merge(Cve_preguntas3, resultado_merge, left_on='P_3', right_on='P_3', how='outer')
resultado_merge = pd.merge(Cve_preguntas4, resultado_merge, left_on='P_4', right_on='P_4', how='outer')
resultado_merge = pd.merge(Cve_preguntas5, resultado_merge, left_on='P_5', right_on='P_5', how='outer')
resultado_merge = pd.merge(Cve_preguntas6, resultado_merge, left_on='P_6', right_on='P_6', how='outer')                       
resultado_merge = pd.merge(Cve_preguntas7, resultado_merge, left_on='P_7', right_on='P_7', how='outer')
resultado_merge = pd.merge(Cve_preguntas8, resultado_merge, left_on='P_8', right_on='P_8', how='outer')
resultado_merge = pd.merge(Cve_preguntas9, resultado_merge, left_on='P_9', right_on='P_9', how='outer')
resultado_merge = pd.merge(Cve_preguntas10, resultado_merge, left_on='P_10', right_on='P_10', how='outer')
                           



# Buscar cuál columna (A, B o C) contiene el valor de 'Respuesta 1'
resultado_merge['R_1'] = resultado_merge.apply(lambda row: 
    'A' if row['A1'] == row['Respuesta 1'] 
    else 'B' if row['B1'] == row['Respuesta 1'] 
    else 'C' if row['C1'] == row['Respuesta 1'] 
    else None, axis=1)
    
resultado_merge['R_2'] = resultado_merge.apply(lambda row: 
    'A' if row['A2'] == row['Respuesta 2'] 
    else 'B' if row['B2'] == row['Respuesta 2'] 
    else 'C' if row['C2'] == row['Respuesta 2'] 
    else None, axis=1)

resultado_merge['R_3'] = resultado_merge.apply(lambda row: 
    'A' if row['A3'] == row['Respuesta 3'] 
    else 'B' if row['B3'] == row['Respuesta 3'] 
    else 'C' if row['C3'] == row['Respuesta 3'] 
    else None, axis=1)    

resultado_merge['R_4'] = resultado_merge.apply(lambda row: 
    'A' if row['A4'] == row['Respuesta 4'] 
    else 'B' if row['B4'] == row['Respuesta 4'] 
    else 'C' if row['C4'] == row['Respuesta 4'] 
    else None, axis=1)  
    
resultado_merge['R_5'] = resultado_merge.apply(lambda row: 
    'A' if row['A5'] == row['Respuesta 5'] 
    else 'B' if row['B5'] == row['Respuesta 5'] 
    else 'C' if row['C5'] == row['Respuesta 5'] 
    else None, axis=1)      
    
resultado_merge['R_6'] = resultado_merge.apply(lambda row: 
    'A' if row['A6'] == row['Respuesta 6'] 
    else 'B' if row['B6'] == row['Respuesta 6'] 
    else 'C' if row['C6'] == row['Respuesta 6'] 
    else None, axis=1)      
    
resultado_merge['R_7'] = resultado_merge.apply(lambda row: 
    'A' if row['A7'] == row['Respuesta 7'] 
    else 'B' if row['B7'] == row['Respuesta 7'] 
    else 'C' if row['C7'] == row['Respuesta 7'] 
    else None, axis=1)  

resultado_merge['R_8'] = resultado_merge.apply(lambda row: 
    'A' if row['A8'] == row['Respuesta 8'] 
    else 'B' if row['B8'] == row['Respuesta 8'] 
    else 'C' if row['C8'] == row['Respuesta 8'] 
    else None, axis=1)  

resultado_merge['R_9'] = resultado_merge.apply(lambda row: 
    'A' if row['A9'] == row['Respuesta 9'] 
    else 'B' if row['B9'] == row['Respuesta 9'] 
    else 'C' if row['C9'] == row['Respuesta 9'] 
    else None, axis=1)  
    
resultado_merge['R_10'] = resultado_merge.apply(lambda row: 
    'A' if row['A10'] == row['Respuesta 10'] 
    else 'B' if row['B10'] == row['Respuesta 10'] 
    else 'C' if row['C10'] == row['Respuesta 10'] 
    else None, axis=1)      
    
# Guardar el DataFrame resultante en un nuevo archivo Excel
archivo_salida = r'C:/Users/UNISA_96/Documents/Analisis PAI/Respuestas_CorrespondidasM7.xlsx'
resultado_merge.to_excel(archivo_salida, index=False)
    
    
# Seleccionar las columnas requeridas
Reactivos_Seleccionados = resultado_merge[['Nombre', 'CVE_REAC1', 'R_1', 
                                           'CVE_REAC2','R_2', 
                                           'CVE_REAC3','R_3',  
                                           'CVE_REAC4','R_4', 
                                           'CVE_REAC5','R_5', 
                                           'CVE_REAC6','R_6', 
                                           'CVE_REAC7','R_7', 
                                           'CVE_REAC8','R_8', 
                                           'CVE_REAC9','R_9', 
                                           'CVE_REAC10','R_10']]

archivo_salida = r'C:/Users/UNISA_96/Documents/Analisis PAI/CalibracionesEx_M7.xlsx'
Reactivos_Seleccionados.to_excel(archivo_salida, index=False)

import matplotlib.pyplot as plt


# Contar la cantidad de respuestas por opción (A, B, C) para cada pregunta
respuestas_por_pregunta = {
    'R_1': resultado_merge['R_1'].value_counts(),
    'R_2': resultado_merge['R_2'].value_counts(),
    'R_3': resultado_merge['R_3'].value_counts(),
    'R_4': resultado_merge['R_4'].value_counts(),
    'R_5': resultado_merge['R_5'].value_counts(),
    'R_6': resultado_merge['R_6'].value_counts(),
    'R_7': resultado_merge['R_7'].value_counts(),
    'R_8': resultado_merge['R_8'].value_counts(),
    'R_9': resultado_merge['R_9'].value_counts(),
    'R_10': resultado_merge['R_10'].value_counts()
}

# Convertir el diccionario a un DataFrame y rellenar valores faltantes con 0
df_respuestas = pd.DataFrame(respuestas_por_pregunta).fillna(0)

# Crear gráfico de barras
fig, ax = plt.subplots(figsize=(10, 7))

# Posibles respuestas
respuestas = ['A', 'B', 'C']

# Graficar cada respuesta (A, B, C)
for respuesta in respuestas:
    ax.bar(df_respuestas.columns, df_respuestas.loc[respuesta], label=f'Respuesta {respuesta}')

# Añadir título y etiquetas
ax.set_title('Distribución de Respuestas por Pregunta')
ax.set_xlabel('Preguntas')
ax.set_ylabel('Cantidad de Respuestas')

# Rotar las etiquetas del eje X para mejor legibilidad
plt.xticks(rotation=45)

# Añadir leyenda
ax.legend(title='Respuestas')

# Mostrar gráfico
plt.tight_layout()
plt.show()
