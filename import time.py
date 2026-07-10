import os
import re
import pandas as pd

# ############################################################################
# 🟢 CONFIGURACIÓN DE RUTAS
# ############################################################################
# Carpeta donde está tu archivo CSV
RUTA_CARPETA = r"C:\Users\mendozapa\Desktop\CODEPABLO\DATA-REPORTES\MDA-LARAIGO"
NOMBRE_SALIDA = "2026_REPORTELARAIGO_MATERIALES_PROCESADO.xlsx"

# ############################################################################
# 🔵 LÓGICA DE EXTRACCIÓN (REGEX)
# ############################################################################
def extraer_datos(texto):
    datos = {"tecnico": None, "departamento": None, "contrata": None, "grupo": None, "SOT": None, "Tipo": None, "Motivo": None, "Submotivo": None}
    if pd.isna(texto): return pd.Series(datos)

    def get_match(patron, txt, group=1):
        m = re.search(patron, txt, re.IGNORECASE | re.DOTALL)
        return m.group(group).strip() if m else None

    datos["tecnico"] = get_match(r"t[eé]cnico\s+(.*?)\s+(?:del departamento|departamento)", texto)
    datos["departamento"] = get_match(r"departamento\s+(.*?)(?=\s+(?:de la contrata|contrata))", texto)
    datos["contrata"] = get_match(r"contrata\s+(.*?)(?=\s*(?:,?\s*necesito|\s+grupo|\s+con la SOT))", texto)
    datos["grupo"] = get_match(r"grupo\s+(.*?)\s+con la SOT", texto)
    datos["SOT"] = get_match(r"\b(\d{8})\b", texto)
    datos["Tipo"] = get_match(r"Tipo:\s*(.*?)(?=\s*Motivo:)", texto)
    datos["Motivo"] = get_match(r"Motivo:\s*(.*?)(?=\s*Submotivo:)", texto)
    datos["Submotivo"] = get_match(r"Submotivo:\s*(.*)", texto)
    return pd.Series(datos)

# ############################################################################
# 🟡 PROCESAMIENTO
# ############################################################################
try:
    # 1. Buscar el primer archivo CSV en la carpeta
    archivos_csv = [f for f in os.listdir(RUTA_CARPETA) if f.endswith('.csv')]
    
    if not archivos_csv:
        print(f"❌ Error: No se encontró ningún archivo .csv en {RUTA_CARPETA}")
    else:
        archivo_entrada = os.path.join(RUTA_CARPETA, archivos_csv[0])
        print(f"📂 Procesando archivo: {archivo_entrada}")
        
        # 2. Carga y limpieza inicial
        df = pd.read_csv(archivo_entrada, encoding="cp1252", sep="|")
        df.columns = df.columns.str.strip()
        
        # Filtrado
        df = df[df["Texto de interacción"].fillna("").str.contains(r"soy el t[eé]cnico.*del departamento", case=False, regex=True)]
        
        if not df.empty:
            # Aplicar extracción
            nuevas_cols = df["Texto de interacción"].apply(extraer_datos)
            df_final = pd.concat([df, nuevas_cols], axis=1)
            
            # Fecha y ordenamiento
            df_final["fecha_hora"] = pd.to_datetime(df_final["Fecha de interacción - Hora de interacción"], errors="coerce", dayfirst=True)
            df_final = df_final.sort_values(by=["Ticket", "fecha_hora"], ascending=[True, False])
            df_final["row"] = df_final.groupby("Ticket").cumcount() + 1
            
            if "fecha_hora" in df_final.columns: df_final.drop(columns=["fecha_hora"], inplace=True)
            
            # 3. Guardado
            ruta_salida_excel = os.path.join(RUTA_CARPETA, NOMBRE_SALIDA)
            df_final.to_excel(ruta_salida_excel, index=False)
            print(f"🚀 Proceso completado. Archivo generado en: {ruta_salida_excel}")
        else:
            print("⚠️ El filtro no devolvió ninguna coincidencia en el archivo.")

except Exception as e:
    print(f"❌ Error crítico: {e}")