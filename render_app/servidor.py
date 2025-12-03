# ============================================================
# PY_INICIAL(03.12) - Versión corregida con integración de fotos
# ============================================================
import os
import tempfile
import unicodedata
from datetime import datetime
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = Flask(__name__, template_folder="templates")
CORS(app)


# =========================
# normalizar texto
def normalizar(texto):
    if texto is None:
        return ""
    return (
        unicodedata.normalize("NFKD", str(texto))
        .encode("ascii", "ignore")
        .decode("ascii")
    )


def valOrDash(v):
    return v if (v is not None and str(v).strip() != "") else "-"


# FUNS DOCX (mantenidas y adaptadas)
def set_cell_style(cell, text, font_size=10, bold=False, align_center=True):
    cell.text = str(text)
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(font_size)
            run.bold = bold
        paragraph.alignment = (
            WD_PARAGRAPH_ALIGNMENT.CENTER
            if align_center
            else WD_PARAGRAPH_ALIGNMENT.LEFT
        )
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def add_subtitle(doc, text, indent=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(11)
    run.font.name = "Calibri"
    run.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    if indent:
        p.paragraph_format.left_indent = Inches(0.3)


def add_note(doc, note="*NO CUENTA CON DICHO ELEMENTO"):
    p = doc.add_paragraph()
    run = p.add_run(note)
    run.italic = True
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(255, 0, 0)
    run.font.name = "Calibri"
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p.paragraph_format.space_after = Pt(6)


def create_table(doc, rows, cols, font_size=10, indent=False):
    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    if indent:
        tbl = table._element  # <w:tbl>
        tblPr_elems = tbl.xpath("w:tblPr")
        if tblPr_elems:
            tblPr = tblPr_elems[0]
        else:
            tblPr = OxmlElement("w:tblPr")
            tbl.append(tblPr)
        tblInd = OxmlElement("w:tblInd")
        tblInd.set(qn("w:w"), "300")
        tblInd.set(qn("w:type"), "dxa")
        tblPr.append(tblInd)
    for row in table.rows:
        for cell in row.cells:
            set_cell_style(cell, "-", font_size=font_size)
    return table


def insertar_recuadro_foto(doc, ancho_cm=15, alto_cm=10):
    """
    Inserta un recuadro de imagen de ancho_cm x alto_cm con borde.
    """
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    ancho_in = ancho_cm / 2.54
    alto_twips = int(alto_cm * 567)  # 1 cm ≈ 567 twips

    cell = table.cell(0, 0)
    try:
        cell.width = Cm(ancho_cm)
    except Exception:
        pass

    # altura exacta de fila
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(alto_twips))
    trHeight.set(qn("w:hRule"), "exact")
    trPr.append(trHeight)

    p = cell.paragraphs[0]
    run = p.add_run("ESPACIO PARA IMAGEN")
    run.font.name = "Calibri"
    run.font.size = Pt(11)
    run.bold = True
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # fondo + borde
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), "FFFFFF")
    tcPr.append(shd)

    borders = OxmlElement("w:tcBorders")
    for side in ["top", "left", "bottom", "right"]:
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "12")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")
        borders.append(border)
    tcPr.append(borders)


def insertar_foto_doc(doc, img_path, ancho_cm=15, alto_cm=10):
    """
    Inserta imagen con tamaño exacto ancho_cm x alto_cm (Cm).
    """
    try:
        doc.add_picture(img_path, width=Cm(ancho_cm), height=Cm(alto_cm))
    except Exception:
        insertar_recuadro_foto(doc, ancho_cm=ancho_cm, alto_cm=alto_cm)


# =========================
# Convierte payload JSON -> DataFrames
def build_dfs_from_json(payload):
    df_info = pd.DataFrame([payload.get("general", {})])
    df_tanques = pd.DataFrame(payload.get("tanques", []))
    df_accesorios = pd.DataFrame(payload.get("accesoriosTanque", []))
    df_red = pd.DataFrame(payload.get("accesoriosRed", []))
    df_equipos = pd.DataFrame(payload.get("equipos", []))
    # Observaciones: puede venir como dict de subpuntos
    obs = payload.get("observaciones", {})
    # si observaciones viene como lista/df, el comportamiento original ya lo manejaba
    if isinstance(obs, dict):
        # convertimos a df estilo [ {Subpunto: '7.1', Observación: '...'}, ... ] si es necesario
        df_obs = pd.DataFrame([obs])
    else:
        df_obs = pd.DataFrame(obs)
    return df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs


# =========================
# Helper: búsqueda tolerante de fotos en fotos_dict
def normalize_key(s):
    return str(s).strip().lower().replace(" ", "_").replace("-", "_")


def find_photos_for_title(fotos_dict, title, fallback_keys=None):
    """
    Dada una 'title' textual (por ejemplo "9.1. FOTO PANORÁMICA DE LA ZONA")
    intenta encontrar en fotos_dict las imágenes correspondientes.
    Estrategias:
      1) Buscar key exacta en fotos_dict
      2) Normalizar title y buscar keys que contengan la normalized title o viceversa
      3) Probar con fallback_keys (lista) si se proporcionan (ej: ['tanque_1','tanque_1_panoramica'])
    Devuelve lista de rutas (posible vacía).
    """
    if fotos_dict is None:
        return []

    # 1) keys exactas
    if title in fotos_dict:
        return fotos_dict[title]

    title_norm = normalize_key(title)

    # try fallback keys first
    if fallback_keys:
        for k in fallback_keys:
            if k in fotos_dict:
                return fotos_dict[k]

    matches = []
    # 2) key contains title norm OR title norm contains key
    for k, arr in fotos_dict.items():
        kn = normalize_key(k)
        if kn == title_norm or kn in title_norm or title_norm in kn:
            matches.extend(arr)

    # 3) try numeric tank/equipment heuristics:
    # if title contains "TANQUE {n}" (case-insensitive) try "tanque_n"
    import re

    m = re.search(r"tanque\s*([0-9]+)", title, re.IGNORECASE)
    if m:
        idx = m.group(1)
        key_candidate = f"tanque_{idx}"
        if key_candidate in fotos_dict:
            matches.extend(fotos_dict[key_candidate])

    # remove duplicates and return
    unique = []
    for p in matches:
        if p not in unique:
            unique.append(p)
    return unique


# =========================
# Agrega fotos para subtítulo con matching flexible
def add_fotos_para_subtitulo_with_match(doc, fotos_dict, subtitle_id_candidates, texto, cantidad_recuadros=1):
    """
    subtitle_id_candidates: puede ser una lista de posibles keys (ej: ['tanque_1', 'tanque_1_panoramica'])
    texto: título que se colocará en el doc
    """
    # insertar subtitulo (formato como en original)
    p = doc.add_paragraph()
    run = p.add_run(texto)
    run.bold = True
    run.font.size = Pt(11)
    run.font.name = "Calibri"
    run.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.keep_with_next = True

    # Attempt 1: explicit candidates
    fotos = []
    if subtitle_id_candidates:
        for cand in subtitle_id_candidates:
            if cand in fotos_dict:
                fotos.extend(fotos_dict[cand])

    # Attempt 2: try find_photos_for_title using the textual title
    if not fotos:
        fotos = find_photos_for_title(fotos_dict, texto, fallback_keys=subtitle_id_candidates)

    # Insert photos or recuadros
    if fotos:
        for img in fotos:
            insertar_foto_doc(doc, img, ancho_cm=15, alto_cm=10)
    else:
        for _ in range(cantidad_recuadros):
            insertar_recuadro_foto(doc, ancho_cm=15, alto_cm=10)


# =========================
# generar_docx_desde_dfs (mantiene original y añade fotos_dict)
def generar_docx_desde_dfs(
    df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs, fotos_dict=None
):
    if fotos_dict is None:
        fotos_dict = {}

    doc = Document()

    # --- Título
    titulo = doc.add_paragraph()
    run = titulo.add_run(
        "INFORME DE MANTENIMIENTO PREVENTIVO Y CUMPLIMIENTO NORMATIVO"
    )
    run.bold = True
    run.underline = True
    run.font.size = Pt(14)
    run.font.name = "Calibri"
    titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # === 1. Información de cliente
    add_subtitle(doc, "1. INFORMACIÓN DE CLIENTE")
    campos = [
        "Nombre o razón social del cliente",
        "Fecha de inspección",
        "Dirección",
        "RUC o DNI",
        "Número de instalación",
        "Distrito",
        "Departamento",
        "Coordenadas",
        "Nombre del contacto",
        "Número del contacto",
        "Correo de contacto",
    ]
    tabla1 = create_table(doc, len(campos), 2)
    datos_generales = {}
    if df_info is not None and not df_info.empty:
        datos_generales = df_info.iloc[0].to_dict()
    for i, campo in enumerate(campos):
        valor = datos_generales.get(campo, None)
        set_cell_style(tabla1.cell(i, 0), campo, align_center=False)
        set_cell_style(tabla1.cell(i, 1), valOrDash(valor), align_center=False)

    # === 2. Tipo de instalación
    add_subtitle(doc, "2. TIPO DE INSTALACION")
    tabla2 = create_table(doc, 3, 8)
    tabla2.cell(0, 0).merge(tabla2.cell(0, 1)).text = "DOMESTICO"
    tabla2.cell(0, 2).merge(tabla2.cell(0, 5)).text = "INDUSTRIAL"
    tabla2.cell(0, 6).merge(tabla2.cell(0, 7)).text = "CANALIZADO"
    for cell in tabla2.rows[0].cells:
        set_cell_style(cell, cell.text, bold=True)
    subtipos = [
        "Doméstico",
        "Comercio",
        "Industrial",
        "Agroindustrial",
        "Minera",
        "Avícola",
        "Residencial",
        "Comercial",
    ]
    for i, subtipo in enumerate(subtipos):
        set_cell_style(tabla2.cell(1, i), subtipo)
    for i in range(8):
        set_cell_style(tabla2.cell(2, i), " ")

    # === 3. Tanques inspeccionados
    add_subtitle(doc, "3. TANQUES INSPECCIONADOS")
    headers3 = [
        "Tanque",
        "Capacidad (gal)",
        "N° de serie",
        "Año de fabricación",
        "Tipo de tanque",
        "Fabricante de Tanque",
        "% Actual",
    ]
    num_tanques = len(df_tanques) if df_tanques is not None else 0
    tabla3 = create_table(doc, num_tanques + 1, len(headers3))
    for j, h in enumerate(headers3):
        set_cell_style(tabla3.cell(0, j), h, bold=True)
    for i in range(num_tanques):
        for j, col in enumerate(headers3):
            valor = None
            if df_tanques is not None and col in df_tanques.columns:
                valor = df_tanques.iloc[i][col]
            else:
                if col == "N° de serie" and df_tanques is not None:
                    if "serie" in df_tanques.columns:
                        valor = df_tanques.iloc[i].get("serie", None)
            if j == 0:
                valor = str(i + 1)
            set_cell_style(tabla3.cell(i + 1, j), valOrDash(valor))

    # === 4. Accesorios de los tanques
    add_subtitle(doc, "4. ACCESORIOS DE LOS TANQUES")
    accesorios = [
        "Válvula de llenado",
        "Medidor de porcentaje",
        "Válvula de seguridad",
        "Válvula de drenaje",
        "Multiválvula",
        "Válvula exceso de flujo (Retorno)",
        "Válvula exceso de flujo (Bypass)",
        "Val 3",
    ]
    atributos = ["Marca", "Código", "Serie", "Mes/Año de fabricación"]
    if df_accesorios is None or df_accesorios.empty:
        df_accesorios = pd.DataFrame(columns=["Tanque", "Atributo"] + accesorios)
    unique_tanques = (
        sorted(df_accesorios["Tanque"].unique()) if not df_accesorios.empty else []
    )
    nrows = 1 + len(atributos) * len(unique_tanques) if unique_tanques else 2
    tabla4 = doc.add_table(rows=nrows, cols=2 + len(accesorios))
    tabla4.style = "Table Grid"
    tabla4.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_cell_style(tabla4.cell(0, 0), "N", font_size=7, bold=True)
    set_cell_style(tabla4.cell(0, 1), "Tanques", font_size=7, bold=True)
    for i, acc in enumerate(accesorios):
        set_cell_style(tabla4.cell(0, i + 2), acc, font_size=7, bold=True)
    row_idx = 1
    for tanque in unique_tanques:
        grupo = df_accesorios[df_accesorios["Tanque"] == tanque]
        for attr in atributos:
            set_cell_style(tabla4.cell(row_idx, 1), attr, font_size=7)
            for i, acc in enumerate(accesorios):
                try:
                    val = grupo[grupo["Atributo"] == attr][acc].values
                    if len(val) and pd.notna(val[0]) and str(val[0]).strip() != "":
                        valor = val[0]
                    else:
                        valor = "-"
                except Exception:
                    valor = "-"
                set_cell_style(tabla4.cell(row_idx, i + 2), str(valor), font_size=7)
            row_idx += 1
        # merge number column
        if row_idx - 1 >= row_idx - 4:
            tabla4.cell(row_idx - 4, 0).merge(tabla4.cell(row_idx - 1, 0))
            set_cell_style(tabla4.cell(row_idx - 4, 0), str(tanque), font_size=7)

    # === 5. Accesorios en redes ===
    add_subtitle(doc, "5. ACCESORIOS EN REDES")
    df_red_local = (
        df_red.copy()
        if df_red is not None
        else pd.DataFrame(
            columns=["Tipo", "Marca", "Serie", "Código", "Mes/Año de fabricación"]
        )
    )
    if "Tipo" in df_red_local.columns:
        df_red_local["Tipo"] = df_red_local["Tipo"].astype(str).str.lower().fillna("")
    else:
        df_red_local["Tipo"] = ""
    mapa_accesorios = {
        "llenado_toma_desplazada": "5.1. Válvula de llenado (toma desplazada)",
        "retorno_toma_desplazada": "5.2. Válvula de retorno (toma desplazada)",
        "alivio_hidrostatico": "5.3. Válvula de alivio hidrostático",
        "regulador_primera_etapa": "5.4. Regulador de primera etapa",
        "alivio": "5.5. Válvula de alivio",
        "regulador_2da": "5.6. Regulador de segunda etapa",
        "pull_away": "5.7. Válvula Pull Away",
    }
    grupos = df_red_local.groupby("Tipo") if not df_red_local.empty else {}

    accesorios_red_dict = {}
    for clave, titulo in mapa_accesorios.items():
        add_subtitle(doc, titulo, indent=True)
        lista = (
            grupos.get_group(clave).to_dict(orient="records")
            if (hasattr(grupos, "groups") and clave in grupos.groups)
            else []
        )
        filas = max(2, len(lista) + 1)
        tabla = create_table(doc, filas, 5, indent=True)
        headers = [
            "Válvula",
            "Marca",
            "Serie",
            "Código",
            "Mes/Año de fabricación",
        ]
        for j, h in enumerate(headers):
            set_cell_style(tabla.cell(0, j), h, bold=True)
        if lista:
            for idx, acc in enumerate(lista):
                set_cell_style(tabla.cell(idx + 1, 0), str(idx + 1))
                set_cell_style(tabla.cell(idx + 1, 1), valOrDash(acc.get("Marca")))
                set_cell_style(tabla.cell(idx + 1, 2), valOrDash(acc.get("Serie")))
                set_cell_style(tabla.cell(idx + 1, 3), valOrDash(acc.get("Código")))
                set_cell_style(
                    tabla.cell(idx + 1, 4), valOrDash(acc.get("Mes/Año de fabricación"))
                )
        else:
            for j in range(5):
                set_cell_style(tabla.cell(1, j), "-")
        accesorios_red_dict[clave] = lista

    # Zona medidores 
    zona_medidores_bool = False
    try:
        zona_medidores_bool = (
            df_red_local[df_red_local["Tipo"] == "zona_medidores"]["Código"]
            .astype(str)
            .str.lower()
            .str.contains("true")
            .any()
        )
    except Exception:
        zona_medidores_bool = False
    accesorios_red_dict["zona_medidores"] = zona_medidores_bool

    # === 6. Equipos de la instalación ===
    add_subtitle(doc, "6. EQUIPOS DE LA INSTALACIÓN")
    df_equipos_local = df_equipos.copy() if df_equipos is not None else pd.DataFrame()
    if "Tipo de equipo" in df_equipos_local.columns:
        df_equipos_local["Tipo de equipo"] = (
            df_equipos_local["Tipo de equipo"].astype(str).str.lower().fillna("")
        )
    else:
        df_equipos_local["Tipo de equipo"] = ""

    estructura_equipos = {
        "vaporizador": ["Equipo", "Marca", "Tipo", "Serie", "Año de fabricación", "Capacidad"],
        "quemador": ["Equipo", "Marca", "Modelo", "Tipo", "Serie", "Año de fabricación", "Capacidad (kW)"],
        "decantador": ["Equipo", "Fabricante", "Modelo", "Tipo", "Serie", "Año de fabricación", "Capacidad (gal)"],
        "dispensador_de_gas": ["Equipo", "Marca", "Modelo", "Serie"],
        "bomba": ["Equipo", "Marca", "Modelo", "Serie"],
        "tablero": ["Equipo", "TAG"],
        "estabilizador": ["Equipo", "Marca", "Modelo", "Serie"],
        "detector": ["Equipo", "Marca", "Modelo", "Serie"],
        "extintor": ["Equipo", "Marca", "Serie", "Año de fabricación", "Próxima PH", "Fecha de próxima recarga"],
    }

    grupos_e = df_equipos_local.groupby("Tipo de equipo") if not df_equipos_local.empty else {}

    for idx, (tipo_equipo, columnas) in enumerate(estructura_equipos.items(), start=1):
        nombre_limpio = tipo_equipo.replace("_", " ").capitalize()
        subtitulo = f"6.{idx}. {nombre_limpio}"
        add_subtitle(doc, subtitulo, indent=True)
        datos = (
            grupos_e.get_group(tipo_equipo)
            if (hasattr(grupos_e, "groups") and tipo_equipo in grupos_e.groups)
            else pd.DataFrame(columns=columnas)
        )
        filas = max(2, len(datos) + 1)
        tabla = create_table(doc, filas, len(columnas), indent=True)
        for j, col in enumerate(columnas):
            set_cell_style(tabla.cell(0, j), col, bold=True)
        if not datos.empty:
            for i, (_, fila) in enumerate(datos.iterrows()):
                for j, col in enumerate(columnas):
                    valor = fila.get(col, None)
                    if j == 0:
                        set_cell_style(tabla.cell(i + 1, j), str(i + 1))
                    else:
                        set_cell_style(tabla.cell(i + 1, j), valOrDash(valor))
        else:
            for j in range(len(columnas)):
                set_cell_style(tabla.cell(1, j), "-")

    equipos_instalacion = {
        k: grupos_e.get_group(k).to_dict(orient="records")
        if (hasattr(grupos_e, "groups") and k in grupos_e.groups)
        else []
        for k in estructura_equipos.keys()
    }

    # === 7. Observaciones generales ===
    add_subtitle(doc, "7. OBSERVACIONES GENERALES")
    df_obs_local = df_obs.copy() if df_obs is not None else pd.DataFrame(columns=["Subpunto", "Observación"])
    subtitulos_7 = {
        "7.1": "7.1. Observaciones al cliente",
        "7.2": "7.2. Observaciones en red de llenado y retorno",
        "7.3": "7.3. Observaciones en zona de tanque",
        "7.4": "7.4. Observaciones en red de consumo",
    }
    for clave, titulo in subtitulos_7.items():
        add_subtitle(doc, titulo, indent=True)
        texto = df_obs_local[df_obs_local["Subpunto"] == clave]["Observación"].values
        if len(texto) and str(texto[0]).strip() != "":
            doc.add_paragraph(str(texto[0]).strip())
        else:
            doc.add_paragraph("-")

    add_subtitle(doc, "7.5. Observaciones en equipos varios (Vaporizador, Quemador, Decantador, etc)", indent=True)
    equipos_obs = [
        "Vaporizador",
        "Quemador",
        "Decantador",
        "Dispensador de gas",
        "Bomba de abastecimiento",
        "Tablero eléctrico",
        "Estabilizador",
        "Detector de gases",
        "Extintor",
    ]
    tabla_obs = create_table(doc, len(equipos_obs) + 1, 2, indent=True)
    set_cell_style(tabla_obs.cell(0, 0), "Equipo", bold=True)
    set_cell_style(tabla_obs.cell(0, 1), "Observación", bold=True)
    texto_75 = df_obs_local[df_obs_local["Subpunto"] == "7.5"]["Observación"].values
    observaciones_75 = []
    if len(texto_75) and str(texto_75[0]).strip():
        observaciones_75 = [x.strip() for x in str(texto_75[0]).split(".") if x.strip()]
    for i, equipo in enumerate(equipos_obs):
        set_cell_style(tabla_obs.cell(i + 1, 0), equipo)
        set_cell_style(tabla_obs.cell(i + 1, 1), observaciones_75[i] if i < len(observaciones_75) else "-")

    # === 8. Evidencia general ===
    add_subtitle(doc, "8. EVIDENCIA FOTOGRÁFICA (del establecimiento)")
    insertar_recuadro_foto(doc)

    # === 9. Evidencia fotográfica de elementos de la instalación ===
    add_subtitle(doc, "9. Evidencia fotográfica de elementos de la instalación")

    # Preparar bloques para punto 9 (idéntico al original)
    tanques_for_block = (
        df_tanques.to_dict(orient="records")
        if df_tanques is not None and not df_tanques.empty
        else []
    )
    df_equipos_local = df_equipos.copy() if df_equipos is not None else pd.DataFrame()
    equipos_for_block = (
        df_equipos_local.to_dict(orient="records")
        if df_equipos_local is not None and not df_equipos_local.empty
        else []
    )
    accesorios_red_for_block = accesorios_red_dict if 'accesorios_red_dict' in locals() else {}

    bloque_9 = []
    contador = 1

    # 9.1 Panorámica general
    bloque_9.append((f"9.{contador}. FOTO PANORÁMICA DE LA ZONA", True, 1)); contador += 1

    # Placas por tanque
    for i, t in enumerate(tanques_for_block):
        serie = valOrDash(t.get("N° de serie") or t.get("serie"))
        bloque_9.append((f"9.{contador}. PLACA DE TANQUE {i+1} DE SERIE: {serie}", True, 1)); contador += 1

    # Panorámica de alrededores por tanque (ejemplo con 4 recuadros)
    for i, t in enumerate(tanques_for_block):
        serie = valOrDash(t.get("N° de serie") or t.get("serie"))
        bloque_9.append((f"9.{contador}. FOTO PANORÁMICA DE ALREDEDORES DE TANQUE {i+1} DE SERIE: {serie}", True, 4)); contador += 1

    # Bloque iterativo según tanques (varios subtítulos por tanque)
    for i, t in enumerate(tanques_for_block):
        serie = valOrDash(t.get("N° de serie") or t.get("serie"))
        bloque_9.append((f"9.{contador}. FOTO DE BASES DE CONCRETO DE TANQUE {i + 1} DE SERIE: {serie}", True, 1)); contador += 1
        bloque_9.append((f"9.{contador}. FOTO DE MANÓMETROS 0-60 PSI DE TANQUE {i + 1} DE SERIE: {serie}", True, 1)); contador += 1
        bloque_9.append((f"9.{contador}. FOTO DE MANÓMETROS 0-300 PSI DE TANQUE {i + 1} DE SERIE: {serie}", True, 1)); contador += 1
        bloque_9.append((f"9.{contador}. FOTO DE CONEXIÓN DE CHICOTE A LA MULTIVÁLVULA DE TANQUE {i + 1} DE SERIE: {serie}", True, 1)); contador += 1
        bloque_9.append((f"9.{contador}. STICKERS DEL TANQUE {i + 1} DE SERIE: {serie} Y PINTADO", True, 1)); contador += 1
        bloque_9.append((f"9.{contador}. FOTO DE LOS 04 ANCLAJES, PERNOS, TORNILLOS DEL TANQUE {i + 1} DE SERIE: {serie}", True, 1)); contador += 1
        bloque_9.append((f"9.{contador}. FOTO DE VÁLVULA DE LLENADO DE TANQUE {i + 1} DE SERIE: {serie}", True, 1)); contador += 1
        bloque_9.append((f"9.{contador}. FOTO DE VÁLVULA DE SEGURIDAD DE TANQUE {i + 1} DE SERIE: {serie}", True, 1)); contador += 1
        bloque_9.append((f"9.{contador}. FOTO DE VÁLVULA DE DRENAJE DE TANQUE {i + 1} DE SERIE: {serie}", True, 1)); contador += 1
        bloque_9.append((f"9.{contador}. FOTO DE VÁLVULA DE MULTIVÁLVULA DE TANQUE {i + 1} DE SERIE: {serie}", True, 1)); contador += 1
        bloque_9.append((f"9.{contador}. FOTO DE VÁLVULA DE MEDIDOR DE PORCENTAJE DE TANQUE {i + 1} DE SERIE: {serie}", True, 1)); contador += 1

    # Equipos específicos (manteniendo la lista que tenías)
    for tipo in ["estabilizador", "quemador", "vaporizador", "tablero", "bomba", "dispensador_de_gas", "decantador", "detector"]:
        lista_eq = equipos_instalacion.get(tipo, [])
        if lista_eq:
            for eq in lista_eq:
                serie = valOrDash(eq.get("Serie"))
                bloque_9.append((f"9.{contador}. FOTO DE PLACA DE {tipo.upper()} DE SERIE: {serie}", True, 1)); contador += 1
                bloque_9.append((f"9.{contador}. FOTO DE {tipo.upper()}", True, 1)); contador += 1
        else:
            bloque_9.append((f"9.{contador}. FOTO DE PLACA DE {tipo.upper()} DE SERIE: -", False, 1)); contador += 1
            bloque_9.append((f"9.{contador}. FOTO DE {tipo.upper()}", False, 1)); contador += 1

    # Toma desplazada (llenado_toma_desplazada)
    tiene_toma = bool(accesorios_red_for_block.get("llenado_toma_desplazada")) if isinstance(accesorios_red_for_block, dict) else False
    bloque_9.append((f"9.{contador}. FOTO DEL PUNTO DE TRANSFERENCIA DESPLAZADO", tiene_toma, 1)); contador += 1
    bloque_9.append((f"9.{contador}. FOTO DE LA CAJA DE LA TOMA DESPLAZADA", tiene_toma, 1)); contador += 1
    bloque_9.append((f"9.{contador}. FOTO DEL RECORRIDO DESDE TOMA DESPLAZADA HASTA TANQUE", tiene_toma, 1)); contador += 1

    # Accesorios individuales (mapa)
    mapa = {
        "llenado_toma_desplazada": "VÁLVULA DE LLENADO TOMA DESPLAZADA",
        "retorno_toma_desplazada": "VÁLVULA DE RETORNO TOMA DESPLAZADA",
        "alivio_hidrostatico": "VÁLVULA DE ALIVIO HIDROSTÁTICO",
        "regulador_primera_etapa": "REGULADOR DE PRIMERA ETAPA",
        "alivio": "VÁLVULA DE ALIVIO",
        "regulador_2da": "REGULADOR DE SEGUNDA ETAPA",
        "pull_away": "VÁLVULA PULL AWAY",
    }
    for clave, nombre in mapa.items():
        lista = accesorios_red_for_block.get(clave, []) if (isinstance(accesorios_red_for_block, dict) and clave in accesorios_red_for_block) else []
        cantidad = max(1, len(lista))
        for idx in range(cantidad):
            if idx < len(lista):
                codigo = valOrDash(lista[idx].get("Código"))
                existe = True
            else:
                codigo = "-"
                existe = False
            bloque_9.append((f"9.{contador}. FOTO DE {nombre} {idx+1} DE CÓDIGO: {codigo}", existe, 1))
            contador += 1

    # Zona de medidores
    bloque_9.append((f"9.{contador}. FOTO DE ZONA MEDIDORES", bool(accesorios_red_for_block.get("zona_medidores")) if isinstance(accesorios_red_for_block, dict) else False, 1)); contador += 1

    # === Subtitulos internos para punto 9 ===
    def add_foto_con_subtitulo(doc, texto, incluir_imagen=True, num_recuadros=1):
        p = doc.add_paragraph()
        run = p.add_run(texto)
        run.bold = True
        run.font.size = Pt(11)
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.left_indent = Inches(0.3)
        p.paragraph_format.keep_with_next = True
        if incluir_imagen:
            for _ in range(num_recuadros):
                insertar_recuadro_foto(doc)
        else:
            add_note(doc, "*NO CUENTA CON DICHO ELEMENTO")

    def insertar_dos_fotos(doc, bloques):
        # Inserta dos subtítulos seguidos (cada uno con su recuadro)
        for texto, incluir in bloques:
            add_foto_con_subtitulo(doc, texto, incluir_imagen=incluir)

    # === Recorrer bloque_9 con lógica extendida, pero reemplazando recuadros por fotos si existen ===
    i = 0
    while i < len(bloque_9):
        item = bloque_9[i]

        # Compatibilidad: item puede tener 2 o 3 valores
        if len(item) == 2:
            texto, incluir = item
            num_recuadros = 1
        else:
            texto, incluir, num_recuadros = item

        # Construir posibles keys que el frontend podría haber usado:
        # Ejemplo: si el texto contiene "TANQUE 3" generamos keys: 'tanque_3', 'tanque_3_placa', etc.
        # Esto ayuda a emparejar fotos subidas con títulos generados aquí.
        subtitle_candidates = []
        # Normalize text to search numbers
        text_lower = str(texto).lower()
        import re
        m = re.search(r"tanque\s*([0-9]+)", text_lower)
        if m:
            n = m.group(1)
            subtitle_candidates.append(f"tanque_{n}")
            subtitle_candidates.append(f"tanque_{n}_placa")
            subtitle_candidates.append(f"tanque_{n}_panoramica")
            subtitle_candidates.append(f"tanque_{n}_base")
        # equipos: buscar palabra 'placa de <TIPO>' por ejemplo 'PLACa DE QUEMADOR'
        m2 = re.search(r"placa de\s+([a-z0-9_]+)", text_lower)
        if m2:
            tipo_eq = normalize_key(m2.group(1))
            subtitle_candidates.append(f"equipo_{tipo_eq}")
            subtitle_candidates.append(f"{tipo_eq}")

        # Adicional: candidate basado en contador 9.X (por si frontend usó numeración)
        # extra: remove numbering '9.X.' from start and normalize
        without_number = re.sub(r"^9\.\d+\.\s*", "", texto)
        subtitle_candidates.append(normalize_key(without_number))

        if incluir:
            if num_recuadros > 1:
                # intento de insertar fotos con match, si no hay, inserta num_recuadros recuadros
                fotos_found = []
                # check candidates and title
                for cand in subtitle_candidates:
                    fotos_found.extend(fotos_dict.get(cand, []))
                fotos_found = list(dict.fromkeys(fotos_found))
                if not fotos_found:
                    fotos_found = find_photos_for_title(fotos_dict, texto, fallback_keys=subtitle_candidates)
                if fotos_found:
                    for img in fotos_found:
                        insertar_foto_doc(doc, img, ancho_cm=15, alto_cm=10)
                else:
                    for _ in range(num_recuadros):
                        insertar_recuadro_foto(doc, ancho_cm=15, alto_cm=10)
                doc.add_paragraph()
                i += 1
                continue

            # Caso normal: intentar agrupar el siguiente si también es 1 recuadro
            if i + 1 < len(bloque_9):
                item2 = bloque_9[i + 1]
                if len(item2) == 2:
                    texto2, incluir2 = item2
                    num_recuadros2 = 1
                else:
                    texto2, incluir2, num_recuadros2 = item2

                if incluir2 and num_recuadros == 1 and num_recuadros2 == 1:
                    # intentar insertar para texto y texto2 en pareja, buscando fotos para cada uno
                    # primero build candidates for texto2
                    subtitle_candidates2 = []
                    m = re.search(r"tanque\s*([0-9]+)", str(texto2).lower())
                    if m:
                        n = m.group(1)
                        subtitle_candidates2.append(f"tanque_{n}")
                    subtitle_candidates2.append(normalize_key(re.sub(r"^9\.\d+\.\s*", "", texto2)))
                    # buscar fotos para ambos
                    fotos1 = []
                    for cand in subtitle_candidates:
                        fotos1.extend(fotos_dict.get(cand, []))
                    fotos1 = list(dict.fromkeys(fotos1))
                    if not fotos1:
                        fotos1 = find_photos_for_title(fotos_dict, texto, fallback_keys=subtitle_candidates)
                    fotos2 = []
                    for cand in subtitle_candidates2:
                        fotos2.extend(fotos_dict.get(cand, []))
                    fotos2 = list(dict.fromkeys(fotos2))
                    if not fotos2:
                        fotos2 = find_photos_for_title(fotos_dict, texto2, fallback_keys=subtitle_candidates2)

                    if fotos1 or fotos2:
                        # insertar conjunto: si una de las dos no tiene fotos, insertar recuadro(s) en su lugar
                        # insertar fotos1 (si existen) o recuadro
                        if fotos1:
                            for img in fotos1:
                                insertar_foto_doc(doc, img, ancho_cm=15, alto_cm=10)
                        else:
                            insertar_recuadro_foto(doc)
                        # insertar fotos2
                        if fotos2:
                            for img in fotos2:
                                insertar_foto_doc(doc, img, ancho_cm=15, alto_cm=10)
                        else:
                            insertar_recuadro_foto(doc)
                        doc.add_paragraph()
                        i += 2
                        continue
                    else:
                        # si ninguno tiene fotos, usar el comportamiento original (dos recuadros)
                        insertar_dos_fotos(doc, [(texto, True), (texto2, True)])
                        i += 2
                        doc.add_paragraph()
                        continue

            # Sino, intentar insertar una foto para este subtitulo
            fotos_found = []
            for cand in subtitle_candidates:
                fotos_found.extend(fotos_dict.get(cand, []))
            fotos_found = list(dict.fromkeys(fotos_found))
            if not fotos_found:
                fotos_found = find_photos_for_title(fotos_dict, texto, fallback_keys=subtitle_candidates)
            if fotos_found:
                for img in fotos_found:
                    insertar_foto_doc(doc, img, ancho_cm=15, alto_cm=10)
            else:
                add_foto_con_subtitulo(doc, texto, incluir_imagen=True)
            doc.add_paragraph()
            i += 1
            continue

        # Caso: no incluir imagen → agrupar todos seguidos con nota
        while i < len(bloque_9) and not bloque_9[i][1]:
            texto_no = bloque_9[i][0]
            add_foto_con_subtitulo(doc, texto_no, incluir_imagen=False)
            i += 1

    # === 10. EVIDENCIA FOTOGRÁFICA (MANTENIMIENTO REALIZADO) ===
    add_subtitle(doc, "10. EVIDENCIA FOTOGRÁFICA (MANTENIMIENTO REALIZADO)")
    add_note(doc, "NOTA 1: SE DEBERÁ MENCIONAR LOS TRABAJOS EJECUTADOS POR TANQUE (INCLUIR LAS INSPECCIONES QUE SE REALICEN)")
    add_note(doc, "NOTA 2: LAS IMÁGENES DEBEN TENER UN TAMAÑO DE 15CM (LARGO) X 10CM (ALTO) MÁXIMO Y SE DEBERÁ VISUALIZAR CLARAMENTE LOS DATOS RELEVANTES (OBSERVACIONES, DESCRIPCIONES DE ESTADO DE ELEMENTOS, TRABAJO REALIZADO, ETC) DE LOS ELEMENTOS EN LOS TRABAJOS REALIZADOS (TANQUES, ACCESORIOS, REDES)")

    i_trab = 1
    series_tanques = [valOrDash(row.get("N° de serie") or row.get("serie")) for row in tanques_for_block]
    for idx, serie in enumerate(series_tanques, start=1):
        add_subtitle(doc, f"10.{i_trab}. TRABAJOS REALIZADOS EN EL TANQUE {idx} DE SERIE: {serie}", indent=True)
        doc.add_paragraph("(ACTIVIDAD 1: FOTO ANTES Y FOTO DESPUÉS; DESCRIPCIÓN DEL TRABAJO REALIZADO)")
        insertar_recuadro_foto(doc)
        insertar_recuadro_foto(doc)
        doc.add_paragraph("(ACTIVIDAD 2: FOTO ANTES Y FOTO DESPUÉS; DESCRIPCIÓN DEL TRABAJO REALIZADO)")
        insertar_recuadro_foto(doc)
        insertar_recuadro_foto(doc)
        i_trab += 1

    add_subtitle(doc, f"10.{i_trab}. OBSERVACIONES FINALES", indent=True)
    doc.add_paragraph("...")

    # Guardar archivo temporal y devolver path
    tmpdir = tempfile.mkdtemp(prefix="informe_")
    filename = f"Informe_Mantenimiento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    output_path = os.path.join(tmpdir, filename)
    doc.save(output_path)
    return output_path


# =========================
# Endpoint /generar (acepta json o multipart con fotos)
@app.route("/generar", methods=["POST"])
def generar():
    try:
        fotos_dict = {}
        fotos_tmpdir = None

        if request.content_type and request.content_type.startswith("multipart/form-data"):
            json_str = request.form.get("json")
            if not json_str:
                return jsonify({"error": "Falta campo 'json' en multipart"}), 400
            # parse JSON (pandas helper o json.loads)
            try:
                import json
                data = json.loads(json_str)
            except Exception:
                data = pd.io.json.loads(json_str)

            # archivos enviados (campo 'files' repetido)
            file_items = request.files.getlist("files")
            if not file_items:
                file_items = list(request.files.values())

            if file_items:
                fotos_tmpdir = tempfile.mkdtemp(prefix="fotos_")
                for f in file_items:
                    filename = f.filename or "foto.jpg"
                    parts = filename.split("__", 2)
                    subId = parts[0] if len(parts) >= 1 else "otros"
                    safe_name = os.path.basename(filename)
                    save_path = os.path.join(fotos_tmpdir, safe_name)
                    f.save(save_path)
                    fotos_dict.setdefault(subId, []).append(save_path)
        else:
            data = request.get_json()
            fotos_dict = {}

        # build dfs
        df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs = build_dfs_from_json(data)

        # generar docx
        output_path = generar_docx_desde_dfs(
            df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs, fotos_dict=fotos_dict
        )

        # enviar y (opcional) limpiar temporales después de enviar
        return send_file(output_path, as_attachment=True, download_name=os.path.basename(output_path))
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500
        
@app.route("/")
def home():
    return "Backend activo. Usa la ruta /generar."

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT",5001))
    app.run(host="0.0.0.0", port=port)


