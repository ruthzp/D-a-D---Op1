import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.formatting.rule import CellIsRule
from io import BytesIO


# ---------------------------------------------------------
# NORMALIZADOR
# ---------------------------------------------------------
def limpiar(s):
    if pd.isna(s):
        return ""
    return (
        str(s)
        .replace("\xa0", " ")
        .replace("\t", " ")
        .replace("\n", " ")
        .strip()
        .upper()
    )


def _to_int(v):
    """Convierte valores a enteros sin romper."""
    try:
        if pd.isna(v):
            return 0
    except Exception:
        pass
    try:
        return int(v)
    except Exception:
        try:
            return int(float(v))
        except Exception:
            return 0


# ---------------------------------------------------------
# DETECTAR ENCABEZADO PARA ASC – CAJAS – SEDE
# ---------------------------------------------------------
def detectar_fila_encabezados_cajas_sede(df_raw):

    for i in range(min(200, len(df_raw))):
        fila = df_raw.iloc[i].astype(str).apply(limpiar)
        valores = list(fila.values)

        if (
            "SEDE OPERATIVA" in valores
            and "TIPO" in valores
            and "TOTAL INVENTARIO IMPRENTA" in valores
            and "INGRESO" in valores
            and "SALIDA" in valores
        ):
            return i

    raise ValueError("No se encontró encabezado en ASC - CAJAS SEDE.")


# ---------------------------------------------------------
# CARGAR ASC – CAJAS – SEDE
# ---------------------------------------------------------
def cargar_asc_cajas_sede(archivo_asc):

    df_raw = pd.read_excel(archivo_asc, sheet_name="Reporte", header=None)
    fila = detectar_fila_encabezados_cajas_sede(df_raw)

    df = pd.read_excel(
        archivo_asc,
        sheet_name="Reporte",
        header=fila
    )

    df.columns = [limpiar(c) for c in df.columns]

    obligatorias = [
        "SEDE OPERATIVA",
        "TIPO",
        "TOTAL INVENTARIO IMPRENTA",
        "INGRESO",
        "SALIDA",
    ]

    faltan = [c for c in obligatorias if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan columnas obligatorias en ASC - CAJAS SEDE: {faltan}")

    df["SEDE OPERATIVA"] = df["SEDE OPERATIVA"].apply(limpiar)
    df["TIPO"] = df["TIPO"].apply(limpiar)

    return df


# ---------------------------------------------------------
# CLASIFICAR TIPO
# ---------------------------------------------------------
def clasificar_tipo(tipo):
    t = limpiar(tipo)

    if "APLIC" in t:
        return "INSTRUMENTO"

    if "ADIC" in t:
        return "ADICIONAL"

    if "CAND" in t:
        return "CANDADO"

    return None


# ---------------------------------------------------------
# GENERAR CAJAS-SEDE
# ---------------------------------------------------------
def generar_cajas_sede(ruta_plantilla_temp, archivo_asc_cajas_sede):

    try:
        with st.spinner("Generando hoja CAJAS-SEDE..."):

            # --- 1) Cargar ASC
            df = cargar_asc_cajas_sede(archivo_asc_cajas_sede)

            # --- 2) Agrupar por SEDE + TIPO
            index = {}

            for _, row in df.iterrows():
                tipo_cl = clasificar_tipo(row["TIPO"])
                if not tipo_cl:
                    continue

                sede = row["SEDE OPERATIVA"]

                key = (sede, tipo_cl)

                if key not in index:
                    index[key] = {"T": 0, "I": 0, "S": 0}

                index[key]["T"] += _to_int(row["TOTAL INVENTARIO IMPRENTA"])
                index[key]["I"] += _to_int(row["INGRESO"])
                index[key]["S"] += _to_int(row["SALIDA"])

            # --- 3) Cargar plantilla
            wb = load_workbook(ruta_plantilla_temp)
            ws = wb["CAJAS-SEDE"]

            # Mapear encabezados fila 1
            header_map = {}
            for c in ws[1]:
                if c.value:
                    header_map[limpiar(c.value)] = c.column_letter

            def col(name):
                k = limpiar(name)
                if k not in header_map:
                    raise ValueError(f"Columna no encontrada en plantilla: {name}")
                return header_map[k]

            # Sede
            col_sede = col("SEDE")

            # Columnas T
            col_T_INSTR = col("CAJA DE INSTRUMENTO DE APLICACIÓN[T]")  # C
            col_T_ADIC = col("CAJA DE INSTRUMENTO ADICIONAL[T]")       # D
            col_T_CAND = col("CAJA DE CANDADO[T]")                     # E

            # Columnas I
            col_I_INSTR = col("CAJA DE INSTRUMENTO DE APLICACIÓN-I")   # G
            col_I_ADIC = col("CAJA DE INSTRUMENTO ADICIONAL-I")        # I
            col_I_CAND = col("CAJA DE CANDADO-I")                      # K

            # Columnas I%
            col_IP_INSTR = col("CAJA DE INSTRUMENTO DE APLICACIÓN-I[P]")  # H
            col_IP_ADIC = col("CAJA DE INSTRUMENTO ADICIONAL-I[P]")       # J
            col_IP_CAND = col("CAJA DE CANDADO-I[P]")                     # L

            # Columnas S
            col_S_INSTR = col("CAJA DE INSTRUMENTO DE APLICACIÓN-S")   # N
            col_S_ADIC = col("CAJA DE INSTRUMENTO ADICIONAL-S")        # P
            col_S_CAND = col("CAJA DE CANDADO-S")                      # R

            # Columnas S%
            col_SP_INSTR = col("CAJA DE INSTRUMENTO DE APLICACIÓN-S[P]")  # O
            col_SP_ADIC = col("CAJA DE INSTRUMENTO ADICIONAL-S[P]")       # Q
            col_SP_CAND = col("CAJA DE CANDADO-S[P]")                     # S

            # Totales
            col_TOTAL_T = col("CAJAS[T]")     # F
            col_TOTAL_I = col("CAJAS-I[T]")   # M
            col_TOTAL_S = col("CAJAS-S[T]")   # T

            max_row = ws.max_row

            # --- 4) Procesar filas de la plantilla
            for r in range(2, max_row + 1):

                sede_pl = limpiar(ws[f"{col_sede}{r}"].value)
                if not sede_pl:
                    continue

                datos = {
                    "INSTRUMENTO": index.get((sede_pl, "INSTRUMENTO"), {"T": 0, "I": 0, "S": 0}),
                    "ADICIONAL": index.get((sede_pl, "ADICIONAL"), {"T": 0, "I": 0, "S": 0}),
                    "CANDADO": index.get((sede_pl, "CANDADO"), {"T": 0, "I": 0, "S": 0}),
                }

                # --- T ---
                ws[f"{col_T_INSTR}{r}"] = datos["INSTRUMENTO"]["T"]
                ws[f"{col_T_ADIC}{r}"] = datos["ADICIONAL"]["T"]
                # col_T_CAND no se toca

                # --- I ---
                ws[f"{col_I_INSTR}{r}"] = datos["INSTRUMENTO"]["I"]
                ws[f"{col_I_ADIC}{r}"] = datos["ADICIONAL"]["I"]
                ws[f"{col_I_CAND}{r}"] = datos["CANDADO"]["I"]

                # --- S ---
                ws[f"{col_S_INSTR}{r}"] = datos["INSTRUMENTO"]["S"]
                ws[f"{col_S_ADIC}{r}"] = datos["ADICIONAL"]["S"]
                ws[f"{col_S_CAND}{r}"] = datos["CANDADO"]["S"]

                # --- Totales ---
                ws[f"{col_TOTAL_T}{r}"] = f"={col_T_INSTR}{r}+{col_T_ADIC}{r}+{col_T_CAND}{r}"
                ws[f"{col_TOTAL_I}{r}"] = f"={col_I_INSTR}{r}+{col_I_ADIC}{r}+{col_I_CAND}{r}"
                ws[f"{col_TOTAL_S}{r}"] = f"={col_S_INSTR}{r}+{col_S_ADIC}{r}+{col_S_CAND}{r}"

                # --- Porcentajes I ---
                ws[f"{col_IP_INSTR}{r}"] = f"={col_I_INSTR}{r}/{col_T_INSTR}{r}"
                ws[f"{col_IP_ADIC}{r}"] = f"={col_I_ADIC}{r}/{col_T_ADIC}{r}"
                ws[f"{col_IP_CAND}{r}"] = f"={col_I_CAND}{r}/{col_T_CAND}{r}"

                # --- Porcentajes S ---
                ws[f"{col_SP_INSTR}{r}"] = f"={col_S_INSTR}{r}/{col_T_INSTR}{r}"
                ws[f"{col_SP_ADIC}{r}"] = f"={col_S_ADIC}{r}/{col_T_ADIC}{r}"
                ws[f"{col_SP_CAND}{r}"] = f"={col_S_CAND}{r}/{col_T_CAND}{r}"

            # --- 5) Formato condicional
            for nombre, letra in header_map.items():
                if "[P]" in nombre:
                    rango = f"{letra}2:{letra}{max_row}"
                    regla = CellIsRule(
                        operator="lessThan",
                        formula=["1"],
                        font=Font(color="FFFF0000")
                    )
                    ws.conditional_formatting.add(rango, regla)

            # --- 6) Guardar SOLO la hoja CAJAS-SEDE
            for hoja in wb.sheetnames:
                if hoja != "CAJAS-SEDE":
                    del wb[hoja]

            out = BytesIO()
            wb.save(out)
            out.seek(0)

            st.session_state["cajas_sede_generada"] = out

        st.success("Hoja CAJAS-SEDE generada correctamente ✔")

    except Exception as e:
        st.error(f"Error al generar CAJAS-SEDE: {e}")