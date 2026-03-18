import os
import pandas as pd
import matplotlib.pyplot as plt

directorio_actual = os.path.dirname(os.path.abspath(__file__))
ruta_csv = os.path.join(directorio_actual, '..', 'Results', 'Final Results.csv')
df_raw = pd.read_csv(ruta_csv)

df = df_raw[~df_raw['label'].str.contains('-')].copy()

df['Estado'] = df['success'].map({True: 'PASSED', False: 'FAILED'})

resumen = df.groupby('label').agg({
    'elapsed': ['count', 'mean', 'min', 'max', 'std'],
    'bytes': 'mean'
}).reset_index()

resumen.columns = ['Peticion', 'Total Muestras', 'Promedio (ms)', 'Minimo (ms)', 'Maximo (ms)', 'Desviacion Std', 'Tamano Promedio (Bytes)']

ruta_excel = os.path.join(directorio_actual, '..', 'Results', 'Reporte_Final.xlsx')

with pd.ExcelWriter(ruta_excel, engine='xlsxwriter') as writer:
    resumen.to_excel(writer, sheet_name='Resumen', index=False)
    df[['timeStamp', 'elapsed', 'label', 'responseCode', 'Estado', 'bytes', 'URL']].to_excel(writer, sheet_name='Detalle_Completo', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets['Resumen']
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    
    for col_num, value in enumerate(resumen.columns.values):
        worksheet.write(0, col_num, value, header_format)
        worksheet.set_column(col_num, col_num, 18) 

print(f"✅ Excel organizado generado en: {ruta_excel}")

promedio = df['elapsed'].mean()
maximo = df['elapsed'].max()
minimo = df['elapsed'].min()
exitosos = (df['success'] == True).sum()

print(f"--- REPORTE DE RENDIMIENTO ---")
print(f"Total de peticiones filtradas: {len(df)}")
print(f"Peticiones exitosas: {exitosos}")
print(f"Tiempo promedio: {promedio:.2f} ms")
print(f"Tiempo maximo: {maximo} ms")

plt.figure(figsize=(10, 6))
df_success = df[df['success'] == True].reset_index(drop=True)

plt.bar(df_success.index, df_success['elapsed'], color='skyblue', label='Tiempo de Respuesta')
plt.axhline(y=promedio, color='red', linestyle='--', label=f'Promedio: {promedio:.2f}ms')

plt.title('Analisis de Latencia por Usuario - Load Test KaraokeIvan', fontsize=14)
plt.xlabel('Numero de Peticion (Thread)', fontsize=12)
plt.ylabel('Tiempo de Respuesta (ms)', fontsize=12)
plt.legend()
plt.grid(axis='y', linestyle='--', alpha=0.7)

ruta_grafico = os.path.join(directorio_actual, '..', 'Results', 'GraficoRendimiento.png')

plt.savefig(ruta_grafico)
print(f"✅ Grafico generado exitosamente en: {ruta_grafico}")
plt.show()
