import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
import io
import os

# Configuración estética de la página web
st.set_page_config(page_title="Procesador de Asistencias", page_icon="📊", layout="centered")
st.title("📊 Sistema de Procesamiento de Asistencias")

def redondear_hora(hora_str, es_entrada=True):
    if hora_str in ['SIN MARCA', 'FALTA MARCA SALIDA'] or not hora_str:
        return hora_str
    
    dt = datetime.strptime(hora_str, "%H:%M:%S")
    minutos = dt.minute
    
    if es_entrada:
        # REGLA ENTRADA: Si pasa de los 13 minutos, llegó tarde 30 minutos (se penaliza)
        if minutos > 13:
            if minutos <= 30:
                dt = dt.replace(minute=30, second=0)
            else:
                dt = (dt + timedelta(hours=1)).replace(minute=0, second=0)
        else:
            # Llegó a tiempo (tolerancia de 13 min), se clava en la hora en punto
            dt = dt.replace(minute=0, second=0)
    else:
        # REGLA SALIDA: Si se pasa de 46 minutos (incluido el 46), se toma la hora siguiente
        if minutos >= 46:
            dt = (dt + timedelta(hours=1)).replace(minute=0, second=0)
        else:
            if minutos >= 30:
                dt = dt.replace(minute=30, second=0)
            else:
                dt = dt.replace(minute=0, second=0)
                
    return dt.strftime("%H:%M:%S")

# 1. Zona interactiva para subir el archivo Excel
uploaded_file = st.file_uploader("Arrastra aquí el archivo de Excel", type=["xlsx"])

if uploaded_file is not None:
    nombre_original = uploaded_file.name 
    nombre_base, extension = os.path.splitext(nombre_original) 
    nombre_salida = f"{nombre_base}_PROCESADO{extension}" 
    
    with st.spinner("Procesando los datos de asistencia... Por favor espera."):
        try:
            # Leer el archivo subido directamente desde la memoria
            df = pd.read_excel(uploaded_file)
            
            # Leer y Renombrar columnas dinámicamente
            if len(df.columns) >= 5:
                df = df.iloc[:, :5]
                df.columns = ['ID_Persona', 'Nombre', 'Departamento', 'Fecha_Hora_Original', 'Quinta_Columna']
            else:
                df = df.iloc[:, :4]
                df.columns = ['ID_Persona', 'Nombre', 'Departamento', 'Fecha_Hora_Original']
                df['Quinta_Columna'] = ''
                
            # 4. Eliminar filas vacías
            df = df.dropna(subset=['Fecha_Hora_Original', 'ID_Persona'])
            
            # 5. Limpieza de Textos
            df['Departamento'] = df['Departamento'].astype(str).str.replace('CARMONA/', '', case=False).str.strip()
            df['ID_Persona'] = df['ID_Persona'].astype(str).str.replace("'", "", regex=False).str.strip()
            df['Quinta_Columna'] = df['Quinta_Columna'].astype(str).fillna('').str.upper().str.strip()
            
            # 6. Separar Fecha y Hora de forma segura
            df['Fecha_Hora_Formateada'] = pd.to_datetime(df['Fecha_Hora_Original'], errors='coerce')
            df = df.dropna(subset=['Fecha_Hora_Formateada'])
            df['Solo_Fecha'] = df['Fecha_Hora_Formateada'].dt.date
            df['Solo_Hora'] = df['Fecha_Hora_Formateada'].dt.time
            
            
            # Cargar registros en la clase
            lista_de_objetos = []
            
            class datosTrabajador:
                def __init__(self, id, nombre, departamento, fecha, hora, quinta_columna):
                    self.id_registro = id
                    self.nombre = nombre
                    self.departamento = departamento
                    self.fecha = fecha
                    self.hora = hora
                    self.quinta_columna = quinta_columna

            for fila in df.itertuples(index=False):
                nuevo_objeto = datosTrabajador(
                    id=fila.ID_Persona,
                    nombre=fila.Nombre,
                    departamento=fila.Departamento,
                    fecha=fila.Solo_Fecha,
                    hora=fila.Solo_Hora,
                    quinta_columna=fila.Quinta_Columna
                )
                lista_de_objetos.append(nuevo_objeto)
                
            # 8. Agrupar por (Trabajador, Fecha)
            asistencias_por_dia = {}
            for obj in lista_de_objetos:
                clave = (obj.id_registro, obj.nombre, obj.departamento, obj.fecha, obj.quinta_columna)
                if clave not in asistencias_por_dia:
                    asistencias_por_dia[clave] = []
                
                if isinstance(obj.hora, time):
                    hora_texto = obj.hora.strftime("%H:%M:%S")
                else:
                    hora_texto = str(obj.hora).strip()
                asistencias_por_dia[clave].append(hora_texto)
                
            # 9. Calcular horas aplicando filtros y lógica
            reporte_horas = []
            formato = "%H:%M:%S"
            
            for (id_reg, nombre, depto, fecha, quinta_col), horas in asistencias_por_dia.items():
                horas_totales_dia = sorted(horas)
                textos_limpios = []
                
                for h in horas_totales_dia:
                    if not textos_limpios:
                        textos_limpios.append(h)
                    else:
                        t_anterior = datetime.strptime(textos_limpios[-1], formato)
                        t_actual = datetime.strptime(h, formato)
                        if (t_actual - t_anterior).total_seconds() > 120:
                            textos_limpios.append(h)
                
                horas_ordenadas = textos_limpios
                
                # CASO DE MARCA INCOMPLETA (Menos de 2 marcas en el día)
                if len(horas_ordenadas) < 2:
                    entrada_orig = horas_ordenadas[0] if horas_ordenadas else 'SIN MARCA'
                    reporte_horas.append({
                        'ID': id_reg, 'Nombre': nombre, 'Departamento': depto, 'Fecha': str(fecha),
                        'Entrada Original': entrada_orig, 
                        'Salida Original': 'FALTA MARCA SALIDA',
                        'Entrada Redondeada': redondear_hora(entrada_orig, es_entrada=True),
                        'Salida Redondeada': 'FALTA MARCA SALIDA',
                        'Horas Totales': 0.0,
                        'Horas Restadas Descanso': 0.0,
                        'Horas Netas': 0.0, 'Horas Extra': 0.0
                    })
                    continue
                    
                hora_inicio = horas_ordenadas[0]
                hora_final = horas_ordenadas[-1]
                
                # Se obtienen las horas redondeadas oficiales de Entrada y Salida principal
                hora_inicio_red = redondear_hora(hora_inicio, es_entrada=True)
                hora_final_red = redondear_hora(hora_final, es_entrada=False)
                
                dt_inicio_red = datetime.strptime(hora_inicio_red, formato)
                dt_final_red = datetime.strptime(hora_final_red, formato)
                
                diferencia_total = dt_final_red - dt_inicio_red
                horas_totales = diferencia_total.total_seconds() / 3600
                
                horas_a_restar = 0.0
                total_marcas = len(horas_ordenadas)
                
                if total_marcas == 2 and horas_totales >= 7:
                    horas_a_restar = 1.0
                    
                elif total_marcas >= 2: # and total_marcas % 2 == 0:
                    for i in range(1, total_marcas - 1, 2):
                        # Quitamos 'redondear_hora' para usar los minutos reales del almuerzo
                        dt_salida_medio = datetime.strptime(horas_ordenadas[i], formato)
                        dt_regreso_medio = datetime.strptime(horas_ordenadas[i+1], formato)
                        horas_a_restar += (dt_regreso_medio - dt_salida_medio).total_seconds() / 3600
                        print(f"horas a restar {horas_a_restar:.2f} a la persona {nombre} en fecha {fecha}")

                        
                elif total_marcas == 3:
                    dt_salida_medio = datetime.strptime(horas_ordenadas[1], formato) 
                    
                    tiempo_hasta_fin = (dt_final_red - dt_salida_medio).total_seconds() / 3600
                    if tiempo_hasta_fin > 1.0: 
                        horas_a_restar = 1.0 
                    else:
                        horas_a_restar = tiempo_hasta_fin

                        
                #elif total_marcas > 3 and total_marcas % 2 != 0:
                   #horas_a_restar = 1.0
                    #print(f"entro al elif total_marcas > 3 and total_marcas % 2 != 0: {}")
                
                # print(f"se le restan {horas_a_restar} a {nombre}")
                    
                horas_netas = max(0.0, horas_totales - horas_a_restar)
                
                # Cálculo de horas extra sobre la base laboral reglamentaria
                horas_laborales = 7.5 if "MIXTO" in quinta_col else 8.0
                horas_extra = max(0.0, horas_netas - horas_laborales)
                
                # Guardar los registros en el diccionario
                reporte_horas.append({
                    'ID': id_reg, 'Nombre': nombre, 'Departamento': depto, 'Fecha': str(fecha),
                    'Entrada Original': hora_inicio, 'Salida Original': hora_final,
                    'Entrada Redondeada': hora_inicio_red, 'Salida Redondeada': hora_final_red,
                    'Horas Totales': round(horas_totales, 2),
                    'Horas Restadas Descanso': round(horas_a_restar, 2),  # <--- NUEVA COLUMNA AÑADIDA
                    'Horas Netas': round(horas_netas, 2), 'Horas Extra': round(horas_extra, 2)
                })
                
            # --- RENDERIZADO DEL REPORTE UNIFICADO ---
            df_resultado = pd.DataFrame(reporte_horas)
            
            if not df_resultado.empty:
                st.success("¡Procesamiento completado exitosamente!")
                st.subheader("📊 Vista Previa del Reporte")
                st.dataframe(df_resultado, use_container_width=True)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Asistencias')
                buffer.seek(0)
                
                st.download_button(
                    label="📥 Descargar Reporte en Excel",
                    data=buffer,
                    file_name=nombre_salida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("El archivo se leyó pero no se generaron registros válidos para el reporte.")
                
        except Exception as e:
            st.error(f"Ocurrió un error al procesar el archivo: {e}")
