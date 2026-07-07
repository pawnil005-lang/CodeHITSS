import pandas as pd
import glob
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ==========================================
#         CONFIGURACIÓN MANUAL
# ==========================================
# 1. Definir la ruta de tu escritorio
ruta_carpeta = r'C:\Users\mendozapa\HITSS\Angel Jesus Zavala Ubillus - Fotos Compartido\Cortes Julio'

# 2. Definir las 'E' que aparecerán en la tabla de fotos
usuarios_permitidos = ['E759708', 'E759762', 'E759763', 'E760214', 'E760642','E760568', 'E761375']

# 3. Definir la HORA LÍMITE de forma manual (Formato 24 horas 'HH:MM')
# Ejemplo: '11:00' procesará todo lo hecho HASTA las 10:59:59.
hora_corte_manual = '13:00'
# ==========================================

ruta_salida = os.path.join(ruta_carpeta, 'Reporte_Consolidado_Final.xlsx')

# --- PARTE 1: PROCESAR "CORTES TOA" ---
archivos_cortes = glob.glob(os.path.join(ruta_carpeta, "CORTES TOA*.xlsx"))
datos_reporte = []

meses_espanol = {1:'ene', 2:'feb', 3:'mar', 4:'abr', 5:'may', 6:'jun', 
                 7:'jul', 8:'ago', 9:'sep', 10:'oct', 11:'nov', 12:'dic'}

for archivo in archivos_cortes:
    try:
        nombre_archivo = os.path.basename(archivo)
        match = re.search(r'CORTES TOA\s*(.*?)\.xlsx', nombre_archivo, re.IGNORECASE)
        fecha_texto = match.group(1).strip() if match else "Desconocida"
        
        try:
            fecha_obj = pd.to_datetime(fecha_texto, format='%Y%m%d')
            fecha = f"{fecha_obj.day:02d}-{meses_espanol[fecha_obj.month]}"
        except Exception:
            fecha = fecha_texto
        
        df_altas = pd.read_excel(archivo, sheet_name='BASE ALTAS')
        df_mantos = pd.read_excel(archivo, sheet_name='BASE MANTOS')
        df_control = pd.read_excel(archivo, sheet_name='CONTROL X ASESOR')
        
        altas = len(df_altas)
        mantos = len(df_mantos)
        cargados = altas + mantos
        
        columna_pendientes = [col for col in df_control.columns if 'pendientes' in str(col).lower()][0]
        pendientes = pd.to_numeric(df_control[columna_pendientes], errors='coerce').sum()
        
        # === NUEVA CONDICIÓN: IGNORAR SI PENDIENTES ES 0 ===
        if int(pendientes) == 0:
            print(f"Omitido (Pendientes = 0): {nombre_archivo}")
            continue  # Salta a la siguiente iteración sin agregar el documento
        # ===================================================
        
        gestionados = cargados - pendientes
        
        datos_reporte.append({
            'FECHA': fecha, 'CARGADOS': cargados, 'ALTAS': altas,
            'MANTOS': mantos, 'PENDIENTES': int(pendientes), 'GESTIONADOS': int(gestionados)
        })
    except Exception as e:
        print(f"Error procesando {archivo}: {e}")

df_cortes = pd.DataFrame(datos_reporte)
if not df_cortes.empty:
    totales = df_cortes[['CARGADOS', 'ALTAS', 'MANTOS', 'PENDIENTES', 'GESTIONADOS']].sum()
    fila_total = pd.DataFrame({'FECHA': ['TOTAL'], **totales.to_dict()})
    df_cortes = pd.concat([df_cortes, fila_total], ignore_index=True)


# --- PARTE 2: PROCESAR "REGISTRO DE RF FOTOS" ---
archivos_fotos = glob.glob(os.path.join(ruta_carpeta, "*REGISTRO DE RF FOTOS*.xlsx"))
df_fotos_dinamica = pd.DataFrame()

if archivos_fotos:
    try:
        archivo_foto = archivos_fotos[0]
        df_f = pd.read_excel(archivo_foto)
        
        df_f.columns = df_f.columns.astype(str).str.strip()
        
        col_fecha = next((col for col in df_f.columns if 'hora de fi' in col.lower()), None)
        col_hora = next((col for col in df_f.columns if 'nalización' in col.lower()), None)
        
        if 'HORA DE FINALIZACIÓN' in df_f.columns:
            col_fecha = 'HORA DE FINALIZACIÓN'
            col_hora = 'HORA DE FINALIZACIÓN'
            
        if col_fecha and col_hora:
            if 'USUARIO E' in df_f.columns:
                df_f['USUARIO E'] = df_f['USUARIO E'].astype(str).str.strip()
                if len(usuarios_permitidos) > 0:
                    df_f = df_f[df_f['USUARIO E'].isin(usuarios_permitidos)]
            
            horas_str = df_f[col_hora].astype(str).str.replace('a. m.', 'AM', regex=False).str.replace('p. m.', 'PM', regex=False)
            horas_24h = pd.to_datetime(horas_str, errors='coerce').dt.strftime('%H:%M:%S')
            limite_str = pd.to_datetime(hora_corte_manual, format='%H:%M').strftime('%H:%M:%S')
            
            df_f = df_f[horas_24h.fillna('23:59:59') < limite_str]
            
            if not df_f.empty:
                df_f['HORA_INT'] = pd.to_datetime(horas_str, errors='coerce').dt.hour
                df_f['FECHA_SOLO'] = pd.to_datetime(df_f[col_fecha], errors='coerce').dt.strftime('- %d/%m/%Y')
                
                df_grouped = df_f.groupby(['FECHA_SOLO', 'USUARIO E', 'HORA_INT']).size().reset_index(name='SOT')
                
                df_ajustado_list = []
                
                for (fecha_aux, usuario_aux), group in df_grouped.groupby(['FECHA_SOLO', 'USUARIO E']):
                    conteo_por_hora = dict(zip(group['HORA_INT'], group['SOT']))
                    
                    horas_bajas = [h for h, c in conteo_por_hora.items() if 1 <= c <= 5]
                    
                    for h_baja in horas_bajas:
                        cant = conteo_por_hora[h_baja]
                        if cant == 0:
                            continue 
                        
                        objetivos = [h for h, c in conteo_por_hora.items() if h != h_baja and c > 0]
                        
                        if objetivos:
                            hora_cercana = min(objetivos, key=lambda x: abs(x - h_baja))
                            conteo_por_hora[hora_cercana] += cant
                            conteo_por_hora[h_baja] = 0
                        else:
                            conteo_por_hora[h_baja] = 0
                    
                    for h, c in conteo_por_hora.items():
                        if c > 5:
                            df_ajustado_list.append({
                                'FECHA_SOLO': fecha_aux,
                                'USUARIO E': usuario_aux,
                                'HORA_INT': h,
                                'SOT': c
                            })
                
                if df_ajustado_list:
                    df_ajustada = pd.DataFrame(df_ajustado_list)
                    
                    def mapear_hora_string(h_int):
                        am_pm = 'a. m.' if h_int < 12 else 'p. m.'
                        display_hour = h_int if h_int <= 12 else h_int - 12
                        if display_hour == 0:
                            display_hour = 12
                        return f"{display_hour:02d}:00 {am_pm}"
                    
                    df_ajustada['HORA_FORMATO'] = df_ajustada['HORA_INT'].apply(mapear_hora_string)
                    
                    df_fotos_dinamica = pd.pivot_table(
                        df_ajustada, index=['USUARIO E'], columns='HORA_FORMATO', 
                        values='SOT', aggfunc='sum', margins=True, margins_name='Total'
                    ).fillna("")
        else:
            print(f"No se detectaron las columnas de fecha y hora. Columnas halladas: {list(df_f.columns)}")
            
    except Exception as e:
        print(f"Error procesando fotos: {e}")


# --- PARTE 3: EXPORTAR Y APLICAR FORMATO CONJUNTO ---
with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
    if not df_cortes.empty:
        df_cortes.to_excel(writer, sheet_name='Resumen', index=False, startrow=0)
        fila_inicio_segunda_tabla = len(df_cortes) + 4 
    else:
        fila_inicio_segunda_tabla = 0
        
    if not df_fotos_dinamica.empty:
        df_fotos_dinamica.to_excel(writer, sheet_name='Resumen', index=True, startrow=fila_inicio_segunda_tabla)

# --- PARTE 4: MEJORAS VISUALES ---
if os.path.exists(ruta_salida):
    wb = load_workbook(ruta_salida)
    ws = wb['Resumen']

    blue_fill = PatternFill(start_color="3B5E94", end_color="3B5E94", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    center_align = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = center_align

    if not df_cortes.empty:
        for col_idx in range(1, len(df_cortes.columns) + 1):
            ws.cell(row=1, column=col_idx).fill = blue_fill
            ws.cell(row=1, column=col_idx).font = white_font

        fila_total_t1 = len(df_cortes) + 1
        for col_idx in range(1, len(df_cortes.columns) + 1):
            ws.cell(row=fila_total_t1, column=col_idx).fill = blue_fill
            ws.cell(row=fila_total_t1, column=col_idx).font = white_font

    if not df_fotos_dinamica.empty:
        header_row_t2 = fila_inicio_segunda_tabla + 1
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=header_row_t2, column=col_idx)
            if cell.value:
                cell.fill = blue_fill
                cell.font = white_font

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 14

    wb.save(ruta_salida)
    print(f"\n¡Éxito! Ambos reportes consolidados en:\n{ruta_salida}")