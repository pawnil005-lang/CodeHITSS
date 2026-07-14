import pandas as pd
import glob
import os
import re
import shutil
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ==========================================
#         CONFIGURACIÓN MANUAL
# ==========================================
ruta_carpeta = r'C:\Users\mendozapa\HITSS\Angel Jesus Zavala Ubillus - Fotos Compartido\Cortes Julio'
# ruta_carpeta = r'C:\Users\huaysarar\OneDrive - HITSS\Fotos Compartido\Cortes Julio'

usuarios_permitidos = [
    'E759763',  # ERICK JONATHAN NÚÑEZ LUDEÑA
    'E759708',  # FRESIA CLEMENTE RODRIGUEZ
    'E759762',  # RENATO PAULO CALLAN ANDRADE
    'E759747',  # JORGELUIS ARMANDO CORDOVA TORRES
    #'E759761',  # PABLO ANDRÉS MENDOZA CALLE
    #'E759899',  # AARON ADRIANO GUZMAN PHATTE
    'E760214',  # RENZO ALFREDO ULLOA BOZETA
    'E760568',  # RICHARD BRUSS HUAYSARA ROMAN
    'E760642',  # MANUEL ALEJANDRO ANGELES RAMON
    # 'E761375' ,  # ANGEL ZAVALA
    #'E759960',
    #'E760197',
    #'E760189',
    #'E760570'

]

# --- CÁLCULO AUTOMÁTICO DE FECHA Y HORA DE CORTE ---
ahora = datetime.now()
hora_corte_manual = ahora.strftime('%H:00')  # Extrae la hora actual en punto (ej. "15:00")
fecha_hoy_str = ahora.strftime('%d/%m/%Y')   # Extrae la fecha de hoy para el filtro (ej. "12/07/2026")

print(f"[INFO] Fecha actual detectada: {fecha_hoy_str} -> Hora de corte aplicada: {hora_corte_manual}")
# ---------------------------------------------------

ruta_salida = os.path.join(ruta_carpeta, 'Reporte_Consolidado_Final.xlsx')
# ==========================================


# --- PARTE 0: AUTO-REUBICACIÓN DESDE DESCARGAS ---
ruta_descargas = os.path.join(os.path.expanduser('~'), 'Downloads')
archivos_descargados = glob.glob(os.path.join(ruta_descargas, "*REGISTRO DE RF FOTOS*.xlsx"))

if archivos_descargados:
    archivo_reciente = max(archivos_descargados, key=os.path.getmtime)
    try:
        shutil.move(archivo_reciente, os.path.join(ruta_carpeta, "REGISTRO DE RF FOTOS.xlsx"))
        print(f"[ÉXITO] Archivo actualizado desde Descargas: {os.path.basename(archivo_reciente)}")
    except Exception as e:
        print(f"[ERROR] No se pudo mover el archivo: {e}")
else:
    print("[INFO] No hay nuevas descargas. Se usará el 'REGISTRO DE RF FOTOS' existente en carpeta.")


# --- PARTE 1: PROCESAR "CORTES TOA" ---
archivos_cortes = glob.glob(os.path.join(ruta_carpeta, "CORTES TOA*.xlsx"))
datos_reporte = []
meses_espanol = {1:'ene', 2:'feb', 3:'mar', 4:'abr', 5:'may', 6:'jun', 
                 7:'jul', 8:'ago', 9:'sep', 10:'oct', 11:'nov', 12:'dic'}

for archivo in archivos_cortes:
    try:
        match = re.search(r'CORTES TOA\s*(.*?)\.xlsx', os.path.basename(archivo), re.IGNORECASE)
        fecha_texto = match.group(1).strip() if match else "Desconocida"
        
        try:
            f_obj = pd.to_datetime(fecha_texto, format='%Y%m%d')
            fecha = f"{f_obj.day:02d}-{meses_espanol[f_obj.month]}"
        except Exception:
            fecha = fecha_texto
        
        df_altas = pd.read_excel(archivo, sheet_name='BASE ALTAS')
        df_mantos = pd.read_excel(archivo, sheet_name='BASE MANTOS')
        df_control = pd.read_excel(archivo, sheet_name='CONTROL X ASESOR')
        
        # Búsqueda de pendientes en CONTROL X ASESOR
        col_pend_altas = next((c for c in df_control.columns if str(c).strip().upper().replace('_', ' ').startswith('PENDIENTE A')), None)
        col_pend_mantos = next((c for c in df_control.columns if str(c).strip().upper().replace('_', ' ').startswith('PENDIENTE M')), None)
        
        pend_altas = pd.to_numeric(df_control[col_pend_altas], errors='coerce').sum() if col_pend_altas else 0
        pend_mantos = pd.to_numeric(df_control[col_pend_mantos], errors='coerce').sum() if col_pend_mantos else 0

        cargados = len(df_altas) + len(df_mantos)
        
        col_pend = next((c for c in df_control.columns if str(c).strip().upper() == 'PENDIENTES'), None)
        if col_pend:
            pendientes = pd.to_numeric(df_control[col_pend], errors='coerce').sum()
        else:
            pendientes = pend_altas + pend_mantos
        
        if int(pendientes) == 0:
            print(f"[OMITIDO] Pendientes = 0: {os.path.basename(archivo)}")
            continue
            
        datos_reporte.append({
            'FECHA': fecha, 
            'CARGADOS': cargados, 
            'ALTAS': len(df_altas),
            'PEND. ALTAS': int(pend_altas),
            'MANTOS': len(df_mantos), 
            'PEND. MANTOS': int(pend_mantos),
            'PENDIENTES': int(pendientes),
            'GESTIONADOS': int(cargados - pendientes)
        })
    except Exception as e:
        print(f"Error procesando {archivo}: {e}")

df_cortes = pd.DataFrame(datos_reporte)
if not df_cortes.empty:
    column_order = ['FECHA', 'CARGADOS', 'ALTAS', 'PEND. ALTAS', 'MANTOS', 'PEND. MANTOS', 'PENDIENTES', 'GESTIONADOS']
    df_cortes = df_cortes[column_order]
    
    totales = df_cortes[['CARGADOS', 'ALTAS', 'PEND. ALTAS', 'MANTOS', 'PEND. MANTOS', 'PENDIENTES', 'GESTIONADOS']].sum()
    df_cortes = pd.concat([df_cortes, pd.DataFrame({'FECHA': ['TOTAL'], **totales.to_dict()})], ignore_index=True)


# --- PARTE 2: PROCESAR "REGISTRO DE RF FOTOS" ---
archivos_fotos = glob.glob(os.path.join(ruta_carpeta, "*REGISTRO DE RF FOTOS*.xlsx"))
df_fotos_dinamica = pd.DataFrame()

if archivos_fotos:
    try:
        df_f = pd.read_excel(archivos_fotos[0])
        df_f.columns = df_f.columns.astype(str).str.strip()
        
        col_fecha = next((col for col in df_f.columns if 'hora de fi' in col.lower()), None)
        col_hora = next((col for col in df_f.columns if 'nalización' in col.lower()), None)
        
        if 'HORA DE FINALIZACIÓN' in df_f.columns:
            col_fecha = col_hora = 'HORA DE FINALIZACIÓN'
            
        if col_fecha and col_hora:
            # 1. Filtro por usuarios permitidos
            if 'USUARIO E' in df_f.columns and usuarios_permitidos:
                df_f['USUARIO E'] = df_f['USUARIO E'].astype(str).str.strip()
                df_f = df_f[df_f['USUARIO E'].isin(usuarios_permitidos)]
            
            # 2. NUEVO FILTRO: Estrictamente la fecha de hoy (ignora días anteriores como el 11)
            fechas_evaluadas = pd.to_datetime(df_f[col_fecha], errors='coerce').dt.strftime('%d/%m/%Y')
            df_f = df_f[fechas_evaluadas == fecha_hoy_str]
            
            # 3. Filtro por hora máxima (hora de corte automática en punto)
            horas_str = df_f[col_hora].astype(str).str.replace('a. m.', 'AM', regex=False).str.replace('p. m.', 'PM', regex=False)
            horas_24h = pd.to_datetime(horas_str, errors='coerce').dt.strftime('%H:%M:%S')
            limite_str = pd.to_datetime(hora_corte_manual, format='%H:%M').strftime('%H:%M:%S')
            
            df_f = df_f[horas_24h.fillna('23:59:59') < limite_str]
            
            if not df_f.empty:
                df_f['HORA_INT'] = pd.to_datetime(horas_str, errors='coerce').dt.hour
                df_f['FECHA_SOLO'] = f"- {fecha_hoy_str}"
                
                df_grouped = df_f.groupby(['FECHA_SOLO', 'USUARIO E', 'HORA_INT']).size().reset_index(name='SOT')
                df_ajustado_list = []
                
                for (fecha_aux, usuario_aux), group in df_grouped.groupby(['FECHA_SOLO', 'USUARIO E']):
                    conteo = dict(zip(group['HORA_INT'], group['SOT']))
                    for h_baja in [h for h, c in conteo.items() if 1 <= c <= 5]:
                        if conteo[h_baja] == 0:
                            continue
                        objetivos = [h for h, c in conteo.items() if h != h_baja and c > 0]
                        if objetivos:
                            conteo[min(objetivos, key=lambda x: abs(x - h_baja))] += conteo[h_baja]
                        conteo[h_baja] = 0
                        
                    for h, c in conteo.items():
                        if c > 5:
                            df_ajustado_list.append({'FECHA_SOLO': fecha_aux, 'USUARIO E': usuario_aux, 'HORA_INT': h, 'SOT': c})
                
                if df_ajustado_list:
                    df_ajustada = pd.DataFrame(df_ajustado_list)
                    
                    def mapear_hora_string(h_int):
                        am_pm = 'a. m.' if h_int < 12 else 'p. m.'
                        h_disp = h_int if h_int <= 12 else h_int - 12
                        return f"{h_disp if h_disp != 0 else 12:02d}:00 {am_pm}"
                    
                    df_ajustada['HORA_FORMATO'] = df_ajustada['HORA_INT'].apply(mapear_hora_string)
                    horas_ordenadas = [mapear_hora_string(h) for h in sorted(df_ajustada['HORA_INT'].unique())]
                    
                    df_ajustada['HORA_FORMATO'] = pd.Categorical(df_ajustada['HORA_FORMATO'], categories=horas_ordenadas, ordered=True)
                    df_fotos_dinamica = pd.pivot_table(df_ajustada, index=['USUARIO E'], columns='HORA_FORMATO', values='SOT', aggfunc='sum', margins=True, margins_name='Total').fillna("")
            else:
                print(f"[AVISO] Después de filtrar por fecha de hoy ({fecha_hoy_str}) y hora (<{hora_corte_manual}), no quedaron registros en RF FOTOS.")
        else:
            print(f"No se detectaron las columnas requeridas. Halladas: {list(df_f.columns)}")
    except Exception as e:
        print(f"Error procesando fotos: {e}")


# --- PARTE 3: EXPORTAR Y APLICAR FORMATO CONJUNTO ---
with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
    if not df_cortes.empty:
        df_cortes.to_excel(writer, sheet_name='Resumen', index=False, startrow=0)
        fila_t2 = len(df_cortes) + 4
    else:
        fila_t2 = 0
        
    if not df_fotos_dinamica.empty:
        df_fotos_dinamica.to_excel(writer, sheet_name='Resumen', index=True, startrow=fila_t2)


# --- PARTE 4: MEJORAS VISUALES ---
if os.path.exists(ruta_salida):
    wb = load_workbook(ruta_salida)
    ws = wb['Resumen']
    
    blue_fill = PatternFill(start_color="3B5E94", end_color="3B5E94", fill_type="solid") # Azul Base
    pend_fill = PatternFill(start_color="558ED5", end_color="558ED5", fill_type="solid") # Azul Claro
    white_font = Font(color="FFFFFF", bold=True)
    center_align = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = center_align

    if not df_cortes.empty:
        for row_idx in [1, len(df_cortes) + 1]:
            for col_idx, col_name in enumerate(df_cortes.columns, 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if 'PEND' in str(col_name).upper():
                    cell.fill = pend_fill
                else:
                    cell.fill = blue_fill
                cell.font = white_font

    if not df_fotos_dinamica.empty:
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=fila_t2 + 1, column=col_idx)
            if cell.value is not None:
                cell.fill = blue_fill
                cell.font = white_font
        
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=ws.max_row, column=col_idx)
            if cell.value is not None:
                cell.fill = blue_fill
                cell.font = white_font

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 14

    wb.save(ruta_salida)
    print(f"\n¡Éxito! Ambos reportes consolidados con el nuevo formato en:\n{ruta_salida}")