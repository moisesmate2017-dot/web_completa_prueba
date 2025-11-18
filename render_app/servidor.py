# ============================================
# SERVIDOR.PY — PARTE 1
# ============================================

import os
import io
import tempfile
import json
from datetime import datetime
from flask import Flask, request, send_file, jsonify, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename

import pandas as pd
import unicodedata

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = Flask(__name__, template_folder="templates")
CORS(app)

# =====================================================
# CONFIGURACIÓN
# =====================================================

ALLOWED_EXTENSIONS = {"jpg", "jpeg", "png"}
MAX_IMAGE_BYTES = 6 * 1024 * 1024    # 6 MB por imagen

def allowed_file(filename):
    """Verifica si el archivo tiene extensión válida."""
    if not filename or "." not in filename:
        return False
    ext = filename.rsplit(".", 1)[1].lower()
    return ext in ALLOWED_EXTENSIONS

def safe_normalize(s):
    if s is None:
        return ""
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")

def valOrDash(v):
    return v if (v is not None and str(v).strip() != "") else "-"

# =====================================================
# ESTILO Y UTILIDADES DOCX
# =====================================================

def set_cell_style(cell, text, font_size=10, bold=False, align_center=True):
    cell.text = str(text)
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(font_size)
            run.bold = bold
        paragraph.alignment = (
            WD_PARAGRAPH_ALIGNMENT.CENTER if align_center else WD_PARAGRAPH_ALIGNMENT.LEFT
        )
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def add_subtitle(doc, text, indent=False):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(11)
    r.font.name = "Calibri"
    r.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)

    if indent:
        p.paragraph_format.left_indent = Inches(0.3)

def add_note(doc, note="*NO CUENTA CON DICHO ELEMENTO"):
    p = doc.add_paragraph()
    r = p.add_run(note)
    r.italic = True
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(255, 0, 0)
    r.font.name = "Calibri"
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p.paragraph_format.space_after = Pt(6)

def create_table(doc, rows, cols, font_size=10, indent=False):
    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    if indent:
        tbl = table._element
        tblPr = tbl.xpath("w:tblPr")
        if tblPr:
            tblPr = tblPr[0]
        else:
            tblPr = OxmlElement("w:tblPr")
            tbl.append(tblPr)

        ind = OxmlElement("w:tblInd")
        ind.set(qn("w:w"), "300")
        ind.set(qn("w:type"), "dxa")
        tblPr.append(ind)

    # Inicializar celdas con guion
    for row in table.rows:
        for cell in row.cells:
            set_cell_style(cell, "-", font_size=font_size)

    return table


# =====================================================
# FUNCIÓN CRÍTICA: INSERTAR FOTO EN RECUADRO
# =====================================================

def insertar_recuadro_foto_cell(cell, ancho_cm=15, alto_cm=10, image_path=None):
    """
    Inserta dentro de una celda un recuadro de borde con:
      - Imagen (si image_path existe)
      - Texto 'ESPACIO PARA IMAGEN' si no hay imagen

    Mantiene tamaño 15 × 10 cm (ancho se ajusta a 15 cm).
    """

    ancho_in = ancho_cm / 2.54

    # Limpiar el contenido previo de la celda
    for p in cell.paragraphs:
        p.clear()

    p = cell.paragraphs[0]

    if image_path and os.path.exists(image_path):
        try:
            run = p.add_run()
            run.add_picture(image_path, width=Inches(ancho_in))
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        except:
            # Si ocurre error, mostrar placeholder
            run = p.add_run("ESPACIO PARA IMAGEN")
            run.bold = True
            run.font.name = "Calibri"
            run.font.size = Pt(11)
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    else:
        # Placeholder
        run = p.add_run("ESPACIO PARA IMAGEN")
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Bordes
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), "FFFFFF")
    tcPr.append(shd)

    borders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "12")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")
        borders.append(border)
    tcPr.append(borders)
# ============================================
# SERVIDOR.PY — PARTE 2
# ============================================

# =====================================================
# CONSTRUCCIÓN DE DATAFRAMES DESDE EL JSON DEL FRONTEND
# =====================================================

def build_dfs_from_json(payload):
    """
    Construye DataFrames desde el payload JSON.
    Retorna:
      df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs
    """

    general = payload.get("general", {}) or {}
    tanques = payload.get("tanques", []) or []
    accesorios_tanque = payload.get("accesoriosTanque", {}) or {}
    accesorios_red = payload.get("accesoriosRed", []) or []
    equipos = payload.get("equipos", []) or []
    observaciones = payload.get("observaciones", {}) or {}

    # ---------------------
    # df_info
    # ---------------------
    campos_info = [
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
    row_info = {c: general.get(c, "") for c in campos_info}
    df_info = pd.DataFrame([row_info])

    # ---------------------
    # df_tanques
    # ---------------------
    tanques_rows = []
    for t in tanques:
        row = {
            "N° de serie": t.get("serie") or t.get("N° de serie") or "",
            "Capacidad (gal)": t.get("capacidad") or "",
            "Año de fabricación": t.get("anio") or t.get("Año de fabricación") or "",
            "Tipo de tanque": t.get("tipo") or "",
            "Fabricante de Tanque": t.get("fabricante") or "",
            "% Actual": t.get("porcentaje") or "",
        }
        tanques_rows.append(row)
    df_tanques = pd.DataFrame(tanques_rows)

    # ---------------------
    # df_accesorios por tanque
    # ---------------------
    accesorios_cols = [
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

    acc_rows = []
    for tank_key, accs in accesorios_tanque.items():
        try:
            tk = int(tank_key)
        except:
            tk = tank_key

        for attr in atributos:
            row = {"Tanque": tk, "Atributo": attr}

            for acc_name in accesorios_cols:
                acc_entry = accs.get(acc_name, {}) if isinstance(accs, dict) else {}
                row_val = acc_entry.get(attr, "") if isinstance(acc_entry, dict) else ""
                row[acc_name] = row_val if row_val is not None else ""

            acc_rows.append(row)

    df_accesorios = pd.DataFrame(acc_rows)

    # ---------------------
    # df_red
    # ---------------------
    df_red = pd.DataFrame([
        {
            "Tipo": r.get("Tipo", ""),
            "Marca": r.get("Marca", ""),
            "Serie": r.get("Serie", ""),
            "Código": r.get("Código", ""),
            "Mes/Año de fabricación": r.get("Mes/Año de fabricación", ""),
        }
        for r in accesorios_red
    ])

    # ---------------------
    # df_equipos
    # ---------------------
    df_equipos = pd.DataFrame(equipos)

    # ---------------------
    # df_obs
    # ---------------------
    obs_rows = []
    for sp in ["7.1", "7.2", "7.3", "7.4", "7.5"]:
        obs_rows.append({
            "Subpunto": sp,
            "Observación": observaciones.get(sp, "")
        })

    df_obs = pd.DataFrame(obs_rows)

    return df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs



# =====================================================
# CREAR LOS SUBTÍTULOS DE FOTOS QUE EL FRONTEND NECESITA
# =====================================================

def build_slots_for_frontend(df_tanques, df_red, df_equipos):
    """
    Construye lista de subtítulos (slots) donde SÍ se permite subir fotos.
    No incluye los subtítulos con include=False.
    Retorna una lista de dicts:
       [{"code": "9.1", "label": "...", "slots": N}, ...]
    """

    slots = []
    contador = 1

    # 9.1 Panorámica general (siempre existe)
    slots.append({
        "code": f"9.{contador}",
        "label": "FOTO PANORÁMICA DE LA ZONA",
        "cantidad": 1
    })
    contador += 1

    # Placas por tanque
    tanques = df_tanques.to_dict(orient="records") if df_tanques is not None else []
    for i, t in enumerate(tanques):
        serie = valOrDash(t.get("N° de serie") or t.get("serie"))
        slots.append({
            "code": f"9.{contador}",
            "label": f"PLACA DE TANQUE {i+1} DE SERIE: {serie}",
            "cantidad": 1
        })
        contador += 1

    # Panorámica alrededores (4 fotos)
    for i, t in enumerate(tanques):
        serie = valOrDash(t.get("N° de serie") or t.get("serie"))
        slots.append({
            "code": f"9.{contador}",
            "label": f"FOTO PANORÁMICA DE ALREDEDORES DE TANQUE {i+1} DE SERIE: {serie}",
            "cantidad": 4
        })
        contador += 1

    # Accesorios del tanque
    for i, t in enumerate(tanques):
        serie = valOrDash(t.get("N° de serie") or t.get("serie"))

        etiquetas = [
            "BASES DE CONCRETO",
            "MANÓMETRO 0-60 PSI",
            "MANÓMETRO 0-300 PSI",
            "CONEXIÓN DE CHICOTE A LA MULTIVÁLVULA",
            "STICKERS Y PINTADO",
            "04 ANCLAJES, PERNOS, TORNILLOS",
            "VÁLVULA DE LLENADO",
            "VÁLVULA DE SEGURIDAD",
            "VÁLVULA DE DRENAJE",
            "MULTIVÁLVULA",
            "MEDIDOR DE PORCENTAJE"
        ]

        for et in etiquetas:
            slots.append({
                "code": f"9.{contador}",
                "label": f"FOTO DE {et} DE TANQUE {i+1} DE SERIE: {serie}",
                "cantidad": 1
            })
            contador += 1

    # Equipos diversos
    equipos_dict = df_equipos.to_dict(orient="records") if not df_equipos.empty else []
    tipos = ["estabilizador", "quemador", "vaporizador", "tablero", "bomba",
             "dispensador_de_gas", "decantador", "detector"]

    for tp in tipos:
        eq_filtrados = [e for e in equipos_dict if (e.get("Tipo de equipo") or "").lower() == tp]
        if eq_filtrados:
            for eq in eq_filtrados:
                serie = valOrDash(eq.get("Serie"))
                slots.append({
                    "code": f"9.{contador}",
                    "label": f"PLACA DE {tp.upper()} DE SERIE: {serie}",
                    "cantidad": 1
                })
                contador += 1
                slots.append({
                    "code": f"9.{contador}",
                    "label": f"FOTO DE {tp.upper()}",
                    "cantidad": 1
                })
                contador += 1

    # Toma desplazada en red (si existe)
    df_red_local = df_red.copy() if df_red is not None else pd.DataFrame()
    tiene_toma = "llenado_toma_desplazada" in df_red_local["Tipo"].astype(str).str.lower().tolist()
    if tiene_toma:
        slots.append({
            "code": f"9.{contador}",
            "label": "FOTO DEL PUNTO DE TRANSFERENCIA DESPLAZADO",
            "cantidad": 1
        })
        contador += 1
        slots.append({
            "code": f"9.{contador}",
            "label": "FOTO DE LA CAJA DE LA TOMA DESPLAZADA",
            "cantidad": 1
        })
        contador += 1
        slots.append({
            "code": f"9.{contador}",
            "label": "FOTO DEL RECORRIDO DESDE TOMA DESPLAZADA HASTA TANQUE",
            "cantidad": 1
        })
        contador += 1

    # Zona medidores
    zona_bool = "zona_medidores" in df_red_local["Tipo"].astype(str).str.lower().tolist()
    if zona_bool:
        slots.append({
            "code": f"9.{contador}",
            "label": "FOTO DE ZONA MEDIDORES",
            "cantidad": 1
        })
        contador += 1

    # -------- BLOQUE 10 (ACTIVIDADES) --------
    # 10.1, 10.2... por cada tanque
    for idx, t in enumerate(tanques, start=1):
        serie = valOrDash(t.get("N° de serie") or t.get("serie"))
        slots.append({
            "code": f"10.{idx}",
            "label": f"TRABAJOS REALIZADOS EN EL TANQUE {idx} DE SERIE: {serie}",
            "cantidad": 0,   # dinámico (actividades)
            "tipo": "actividad"
        })

    next_code = len(tanques) + 1
    # Redes de llenado y retorno
    slots.append({
        "code": f"10.{next_code}",
        "label": "TRABAJOS REALIZADOS EN REDES DE LLENADO Y RETORNO",
        "cantidad": 0,
        "tipo": "actividad"
    })
    next_code += 1

    slots.append({
        "code": f"10.{next_code}",
        "label": "TRABAJOS REALIZADOS EN REDES DE CONSUMO",
        "cantidad": 0,
        "tipo": "actividad"
    })

    return slots
# ============================================
# SERVIDOR.PY — PARTE 3
# ============================================

# =====================================================
# ENDPOINT: /photo_slots
# Devuelve al frontend qué subtítulos requieren fotos
# =====================================================

@app.post("/photo_slots")
def photo_slots():
    try:
        payload = request.json
        if not payload:
            return jsonify({"error": "Payload vacío"}), 400

        df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs = build_dfs_from_json(payload)
        slots = build_slots_for_frontend(df_tanques, df_red, df_equipos)

        return jsonify({"slots": slots})

    except Exception as e:
        return jsonify({"error": str(e)}), 500



# =====================================================
# ENDPOINT: /generar
# Acepta JSON o multipart/form-data (para fotos)
# =====================================================

@app.post("/generar")
def generar():
    try:
        # =================================================
        # 1. DETERMINAR SI ES JSON O FORM-DATA
        # =================================================
        if request.content_type.startswith("multipart/form-data"):
            if "payload" not in request.form:
                return jsonify({"error": "No se envió payload en form-data"}), 400
            payload = json.loads(request.form["payload"])
        else:
            payload = request.json or {}
        
        # Validación básica
        if not payload:
            return jsonify({"error": "Payload vacío"}), 400

        # =================================================
        # 2. CONSTRUIR DATAFRAMES
        # =================================================
        df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs = build_dfs_from_json(payload)

        # =================================================
        # 3. RECIBIR FOTOS DEL FORM-DATA
        # =================================================
        uploaded_photos = {}

        if request.content_type.startswith("multipart/form-data"):
            for key in request.files:
                file = request.files[key]
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename)
                    temp_path = os.path.join(tempfile.gettempdir(), filename)
                    file.save(temp_path)
                    uploaded_photos[key] = temp_path

        # =================================================
        # 4. CREAR DOCUMENTO
        # =================================================
        doc = Document()

        # =================================================
        # 5. TÍTULO PRINCIPAL
        # =================================================
        title = doc.add_paragraph()
        run = title.add_run("ACTA DE MANTENIMIENTO PREVENTIVO DE INSTALACIÓN DE GLP")
        run.bold = True
        run.font.size = Pt(14)
        run.font.name = "Calibri"
        title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # =================================================
        # 6. PUNTO 1 — INFORMACIÓN GENERAL
        # =================================================
        add_subtitle(doc, "1. INFORMACIÓN DEL CLIENTE")

        table = create_table(doc, rows=len(df_info.columns), cols=2)
        for idx, col in enumerate(df_info.columns):
            set_cell_style(table.cell(idx, 0), col, bold=True)
            set_cell_style(table.cell(idx, 1), valOrDash(df_info.iloc[0][col]))

        # =================================================
        # 7. PUNTO 2 — TANQUES
        # =================================================
        add_subtitle(doc, "2. INFORMACIÓN DE TANQUES")

        if df_tanques.empty:
            add_note(doc, "*NO CUENTA CON TANQUES REGISTRADOS")
        else:
            cols = list(df_tanques.columns)
            table = create_table(doc, rows=len(df_tanques) + 1, cols=len(cols))
            for j, c in enumerate(cols):
                set_cell_style(table.cell(0, j), c, bold=True)

            for i, row in df_tanques.iterrows():
                for j, c in enumerate(cols):
                    set_cell_style(table.cell(i+1, j), valOrDash(row[c]))

        # =================================================
        # 8. PUNTO 3 — ACCESORIOS DE TANQUE
        # =================================================
        add_subtitle(doc, "3. ACCESORIOS DEL TANQUE")

        if df_accesorios.empty:
            add_note(doc)
        else:
            cols = list(df_accesorios.columns)
            table = create_table(doc, rows=len(df_accesorios) + 1, cols=len(cols))
            for j, c in enumerate(cols):
                set_cell_style(table.cell(0, j), c, bold=True)

            for i, row in df_accesorios.iterrows():
                for j, c in enumerate(cols):
                    set_cell_style(table.cell(i+1, j), valOrDash(row[c]))

        # =================================================
        # 9. PUNTO 4 — REDES
        # =================================================
        add_subtitle(doc, "4. REDES DE LLENADO, RETORNO Y CONSUMO")

        if df_red.empty:
            add_note(doc)
        else:
            cols = list(df_red.columns)
            table = create_table(doc, rows=len(df_red) + 1, cols=len(cols))
            for j, c in enumerate(cols):
                set_cell_style(table.cell(0, j), c, bold=True)

            for i, row in df_red.iterrows():
                for j, c in enumerate(cols):
                    set_cell_style(table.cell(i+1, j), valOrDash(row[c]))

        # =================================================
        # 10. PUNTO 5 — EQUIPOS
        # =================================================
        add_subtitle(doc, "5. EQUIPOS")

        if df_equipos.empty:
            add_note(doc)
        else:
            cols = list(df_equipos.columns)
            table = create_table(doc, rows=len(df_equipos) + 1, cols=len(cols))
            for j, c in enumerate(cols):
                set_cell_style(table.cell(0, j), c, bold=True)

            for i, row in df_equipos.iterrows():
                for j, c in enumerate(cols):
                    set_cell_style(table.cell(i+1, j), valOrDash(row[c]))

        # =================================================
        # 11. PUNTO 6 — OBSERVACIONES
        # =================================================
        add_subtitle(doc, "6. OBSERVACIONES")

        if df_obs.empty:
            add_note(doc)
        else:
            table = create_table(doc, rows=len(df_obs)+1, cols=2)
            set_cell_style(table.cell(0, 0), "Subpunto", bold=True)
            set_cell_style(table.cell(0, 1), "Observación", bold=True)
            for i, row in df_obs.iterrows():
                set_cell_style(table.cell(i+1, 0), row["Subpunto"])
                set_cell_style(table.cell(i+1, 1), valOrDash(row["Observación"]))

        # =================================================
        #     ATENCIÓN: AQUÍ TERMINA PARTE 3
        # =================================================

        # La parte 4 continúa desde aquí,
        # generando el PUNTO 7, 8, 9.x (fotos),
        # y finalmente el bloque 10.x de actividades (con fotos antes/después).
# ============================================
# SERVIDOR.PY — PARTE 4
# (continuación)
# ============================================

# ============================
# UTIL: PARSEAR ARCHIVOS SUBIDOS
# ============================
def parse_uploaded_photos_dict(request_files):
    """
    Espera keys con estos formatos (flexible):
      - photo__9_3__1
      - photo__9_3__img1
      - photo__10_1__act1__antes
      - photo__10_1__act1__despues

    Retorna:
      photos_map = {
        "9_3": [path1, path2, ...],   # para slots normales
        "10_1": {
            "activities": {
                1: {"antes": path_or_none, "despues": path_or_none, "otros": [..]},
                2: {...}
            }
        }
      }
    """
    photos_map = {}
    for key in request_files:
        if not key.startswith("photo__"):
            # aceptar cualquier otro file key ignorado
            continue
        parts = key.split("__")[1:]  # quitar prefijo 'photo'
        # parts examples:
        # ['9_3','1']  OR ['10_1','act1','antes']
        if len(parts) == 0:
            continue
        slot = parts[0]  # e.g., "9_3" or "10_1"
        file = request_files.get(key)
        if not file:
            continue
        filename = secure_filename(file.filename)
        # Guardar en temp
        fd, tmp_path = tempfile.mkstemp(prefix="img_", suffix="_" + filename)
        os.close(fd)
        file.save(tmp_path)
        # Asegurar tamaño permitido
        try:
            if os.path.getsize(tmp_path) > MAX_IMAGE_BYTES:
                # eliminar y saltar
                os.remove(tmp_path)
                continue
        except:
            pass

        # Activity pattern
        if len(parts) >= 3 and parts[1].startswith("act"):
            # parts[1] ejemplo: act1 -> obtener indice
            try:
                act_idx = int(parts[1].replace("act", ""))
            except:
                act_idx = 1
            typ = parts[2].lower() if len(parts) >= 3 else "antes"
            photos_slot = photos_map.setdefault(slot, {})
            acts = photos_slot.setdefault("activities", {})
            act_entry = acts.setdefault(act_idx, {"antes": None, "despues": None, "otros": []})
            if typ in ("antes", "before", "bef"):
                act_entry["antes"] = tmp_path
            elif typ in ("despues", "después", "after", "aft"):
                act_entry["despues"] = tmp_path
            else:
                # store in otros
                act_entry["otros"].append(tmp_path)
            continue

        # Otherwise treat as normal slot image (append)
        arr = photos_map.setdefault(slot, {}).setdefault("images", [])
        arr.append(tmp_path)

    return photos_map


# ============================
# UTIL: Insertar N recuadros horizontales en el doc con imagenes si existen
# ============================
def insert_n_recuadros(doc, texto_subtitulo, slot_key, cantidad, photos_map):
    """
    Agrega el subtitulo y luego 'cantidad' recuadros.
    Si photos_map tiene imágenes para slot_key -> las inserta en orden.
    Si no -> deja placeholder.
    """
    add_subtitle(doc, texto_subtitulo, indent=True)
    imgs = []
    if photos_map and slot_key in photos_map and photos_map[slot_key].get("images"):
        imgs = photos_map[slot_key]["images"]
    # insertar en una tabla de 1 fila x cantidad
    table = doc.add_table(rows=1, cols=cantidad)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c in range(cantidad):
        cell = table.cell(0, c)
        if c < len(imgs):
            insertar_recuadro_foto_cell(cell, image_path=imgs[c])
        else:
            insertar_recuadro_foto_cell(cell, image_path=None)
    doc.add_paragraph()  # espacio


# ============================
# CONTINUAR GENERACIÓN DEL DOCUMENTO (PUNTO 7, 8, 9 con fotos, 10 actividades)
# ============================

# (continuando dentro del endpoint generar() después del punto 6 que dejamos en la PARTE 3)

        # =================================================
        # 12. PUNTO 7 — OBSERVACIONES GENERALES (detalladas)
        # =================================================
        add_subtitle(doc, "7. OBSERVACIONES GENERALES")

        # Subpuntos 7.1 - 7.4
        subtitulos_7 = {
            "7.1": "7.1. Observaciones al cliente",
            "7.2": "7.2. Observaciones en red de llenado y retorno",
            "7.3": "7.3. Observaciones en zona de tanque",
            "7.4": "7.4. Observaciones en red de consumo",
        }
        df_obs_local = df_obs.copy() if df_obs is not None else pd.DataFrame(columns=["Subpunto", "Observación"])
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

        # =================================================
        # PUNTO 8 — EVIDENCIA GENERAL (RECUADROS)
        # =================================================
        add_subtitle(doc, "8. EVIDENCIA FOTOGRÁFICA (del establecimiento)")
        # Insertar un recuadro grande (15x10) centrado
        tabla_8 = doc.add_table(rows=1, cols=1)
        tabla_8.alignment = WD_TABLE_ALIGNMENT.CENTER
        insertar_recuadro_foto_cell(tabla_8.cell(0,0), ancho_cm=15, alto_cm=10, image_path=None)
        doc.add_paragraph()

        # =================================================
        # PUNTO 9 — EVIDENCIA FOTOGRÁFICA DE ELEMENTOS DE LA INSTALACIÓN (varios subtítulos)
        # Usaremos la misma lógica que build_slots_for_frontend para respetar orden
        # =================================================

        add_subtitle(doc, "9. Evidencia fotográfica de elementos de la instalación")

        # Generar lista de subtítulos completa (incluye include False) para mantener estructura
        # Pero para insertar imágenes usaremos photos_map (solo incluye las que subió el usuario)
        all_slots = []  # lista de tuples (code,label,cantidad,include_bool)
        # Vamos a construirla replicando build_slots_for_frontend pero incluyendo include flag:
        contador_local = 1
        # 9.1
        all_slots.append((f"9.{contador_local}", "FOTO PANORÁMICA DE LA ZONA", 1, True)); contador_local += 1

        # placas por tanque
        tanques_list = df_tanques.to_dict(orient="records") if df_tanques is not None and not df_tanques.empty else []
        for i, t in enumerate(tanques_list):
            serie = valOrDash(t.get("N° de serie") or t.get("serie"))
            all_slots.append((f"9.{contador_local}", f"PLACA DE TANQUE {i+1} DE SERIE: {serie}", 1, True))
            contador_local += 1

        # panoramica alrededores (4)
        for i, t in enumerate(tanques_list):
            serie = valOrDash(t.get("N° de serie") or t.get("serie"))
            all_slots.append((f"9.{contador_local}", f"FOTO PANORÁMICA DE ALREDEDORES DE TANQUE {i+1} DE SERIE: {serie}", 4, True))
            contador_local += 1

        # bloque iterativo por tanque (varios subtítulos por tanque)
        etiquetas = [
            "FOTO DE BASES DE CONCRETO",
            "FOTO DE MANÓMETROS 0-60 PSI",
            "FOTO DE MANÓMETROS 0-300 PSI",
            "FOTO DE CONEXIÓN DE CHICOTE A LA MULTIVÁLVULA",
            "STICKERS Y PINTADO",
            "FOTO DE LOS 04 ANCLAJES, PERNOS, TORNILLOS",
            "FOTO DE VÁLVULA DE LLENADO",
            "FOTO DE VÁLVULA DE SEGURIDAD",
            "FOTO DE VÁLVULA DE DRENAJE",
            "FOTO DE VÁLVULA DE MULTIVÁLVULA",
            "FOTO DE VÁLVULA DE MEDIDOR DE PORCENTAJE"
        ]
        for i, t in enumerate(tanques_list):
            serie = valOrDash(t.get("N° de serie") or t.get("serie"))
            for et in etiquetas:
                all_slots.append((f"9.{contador_local}", f"{et} DE TANQUE {i+1} DE SERIE: {serie}", 1, True))
                contador_local += 1

        # equipos especificos
        equipos_list = df_equipos.to_dict(orient="records") if df_equipos is not None and not df_equipos.empty else []
        for tipo in ["estabilizador", "quemador", "vaporizador", "tablero", "bomba", "dispensador_de_gas", "decantador", "detector"]:
            lista_eq = [e for e in equipos_list if (e.get("Tipo de equipo") or "").lower() == tipo]
            if lista_eq:
                for eq in lista_eq:
                    serie = valOrDash(eq.get("Serie"))
                    all_slots.append((f"9.{contador_local}", f"FOTO DE PLACA DE {tipo.upper()} DE SERIE: {serie}", 1, True))
                    contador_local += 1
                    all_slots.append((f"9.{contador_local}", f"FOTO DE {tipo.upper()}", 1, True))
                    contador_local += 1
            else:
                # si no existe, mantener el subtítulo con include False (aparecerá en doc con nota)
                all_slots.append((f"9.{contador_local}", f"FOTO DE PLACA DE {tipo.upper()} DE SERIE: -", 1, False))
                contador_local += 1
                all_slots.append((f"9.{contador_local}", f"FOTO DE {tipo.upper()}", 1, False))
                contador_local += 1

        # accesorios en redes (mapa)
        mapa = {
            "llenado_toma_desplazada": "VÁLVULA DE LLENADO TOMA DESPLAZADA",
            "retorno_toma_desplazada": "VÁLVULA DE RETORNO TOMA DESPLAZADA",
            "alivio_hidrostatico": "VÁLVULA DE ALIVIO HIDROSTÁTICO",
            "regulador_primera_etapa": "REGULADOR DE PRIMERA ETAPA",
            "alivio": "VÁLVULA DE ALIVIO",
            "regulador_2da": "REGULADOR DE SEGUNDA ETAPA",
            "pull_away": "VÁLVULA PULL AWAY",
        }

        # construir grupos desde df_red
        df_red_local = df_red.copy() if df_red is not None else pd.DataFrame()
        tipos_en_red = df_red_local["Tipo"].astype(str).tolist() if not df_red_local.empty else []
        for clave, nombre in mapa.items():
            lista = [r for r in equipos_list if False]  # placeholder, mantenemos include True si aparece en df_red
            include = any(clave in t.lower() for t in tipos_en_red)
            if include:
                # si hay N items, generamos N subtítulos
                matches = [r for _, r in df_red_local.iterrows() if (r.get("Tipo") or "").lower() == clave]
                cantidad = max(1, len(matches))
                for idx in range(cantidad):
                    codigo = valOrDash(matches[idx]["Código"]) if idx < len(matches) else "-"
                    all_slots.append((f"9.{contador_local}", f"FOTO DE {nombre} {idx+1} DE CÓDIGO: {codigo}", 1, True))
                    contador_local += 1
            else:
                # Si no existe, mantener el titulo con include False (aparecerá con nota)
                all_slots.append((f"9.{contador_local}", f"FOTO DE {nombre}", 1, False))
                contador_local += 1

        # Zona medidores
        zona_med_bool = any("zona_medidores" in str(t).lower() for t in tipos_en_red)
        all_slots.append((f"9.{contador_local}", "FOTO DE ZONA MEDIDORES", 1, zona_med_bool))
        contador_local += 1

        # =================================================
        # PARSEAR ARCHIVOS SUBIDOS (si los hay)
        # =================================================
        photos_map = {}
        if request.content_type.startswith("multipart/form-data"):
            photos_map = parse_uploaded_photos_dict(request.files)

        # =================================================
        # INSERTAR TODOS LOS SUBTITULOS 9.x CON SUS RECUADROS (si include True) O NOTA (si False)
        # =================================================
        for code, label, cantidad, include in all_slots:
            slot_key = code.replace('.', '_')  # ej "9.3" -> "9_3"
            if include:
                # si hay fotos en photos_map para este slot_key -> insertar
                insert_n_recuadros(doc, f"{code}. {label}", slot_key, cantidad, photos_map)
            else:
                # insertar subtitulo y nota de no cuenta
                add_subtitle(doc, f"{code}. {label}", indent=True)
                add_note(doc, "*NO CUENTA CON DICHO ELEMENTO")
                doc.add_paragraph()

        # =================================================
        # BLOQUE 10: ACTIVIDADES DINÁMICAS (por tanque + redes)
        # payload puede incluir 'actividades' con estructura:
        # {
        #   "10.1": [ { "descripcion": "...", "antes": file?, "despues": file? }, ... ],
        #   "10.2": ...
        # }
        # =================================================

        add_subtitle(doc, "10. EVIDENCIA FOTOGRÁFICA (MANTENIMIENTO REALIZADO)")

        # Nota instructiva
        add_note(doc, "NOTA 1: SE DEBERÁ MENCIONAR LOS TRABAJOS EJECUTADOS POR TANQUE (INCLUIR LAS INSPECCIONES QUE SE REALICEN)")
        add_note(doc, "NOTA 2: LAS IMÁGENES DEBEN TENER UN TAMAÑO DE 15CM X 10CM MÁXIMO Y SE DEBERÁ VISUALIZAR CLARAMENTE LOS DATOS RELEVANTES")

        # actividades desde payload
        actividades_payload = payload.get("actividades", {}) or {}
        # Para generar códigos de 10.x en orden, necesitamos el número de tanques
        n_tanques = len(tanques_list)
        # Procesar por cada 10.x generado arriba: 10.1..10.N por tanques, luego redes
        for idx in range(1, n_tanques+1):
            code = f"10.{idx}"
            titulo = f"{code}. TRABAJOS REALIZADOS EN EL TANQUE {idx} DE SERIE: {valOrDash(tanques_list[idx-1].get('N° de serie') or tanques_list[idx-1].get('serie'))}"
            add_subtitle(doc, titulo, indent=True)
            actividades = actividades_payload.get(code, [])
            if not actividades:
                doc.add_paragraph("-")
                continue
            # Para cada actividad: mostrar descripcion y recuadros antes/despues
            for act_i, act in enumerate(actividades, start=1):
                desc = act.get("descripcion", "").strip()
                if desc:
                    doc.add_paragraph(f"(ACTIVIDAD {act_i}: {desc})")
                else:
                    doc.add_paragraph(f"(ACTIVIDAD {act_i})")
                # insertar dos recuadros: antes y despues
                # buscar en photos_map: key slot "10_<idx>" -> activities -> act_i
                slotkey = code.replace('.', '_')
                antes_path = None
                despues_path = None
                if slotkey in photos_map and "activities" in photos_map[slotkey]:
                    act_entry = photos_map[slotkey]["activities"].get(act_i, {})
                    if act_entry:
                        antes_path = act_entry.get("antes")
                        despues_path = act_entry.get("despues")
                # insertar recuadros lado a lado
                tbl = doc.add_table(rows=1, cols=2)
                tbl.style = "Table Grid"
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                insertar_recuadro_foto_cell(tbl.cell(0,0), image_path=antes_path)
                insertar_recuadro_foto_cell(tbl.cell(0,1), image_path=despues_path)
                doc.add_paragraph()

        # Redes: siguiente código
        code_red_llenado = f"10.{n_tanques+1}"
        add_subtitle(doc, f"{code_red_llenado}. TRABAJOS REALIZADOS EN REDES DE LLENADO Y RETORNO", indent=True)
        actividades = actividades_payload.get(code_red_llenado, [])
        if not actividades:
            doc.add_paragraph("-")
        else:
            for act_i, act in enumerate(actividades, start=1):
                desc = act.get("descripcion", "").strip()
                if desc:
                    doc.add_paragraph(f"(ACTIVIDAD {act_i}: {desc})")
                else:
                    doc.add_paragraph(f"(ACTIVIDAD {act_i})")
                slotkey = code_red_llenado.replace('.', '_')
                antes_path = despues_path = None
                if slotkey in photos_map and "activities" in photos_map[slotkey]:
                    ent = photos_map[slotkey]["activities"].get(act_i, {})
                    antes_path = ent.get("antes")
                    despues_path = ent.get("despues")
                tbl = doc.add_table(rows=1, cols=2)
                tbl.style = "Table Grid"
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                insertar_recuadro_foto_cell(tbl.cell(0,0), image_path=antes_path)
                insertar_recuadro_foto_cell(tbl.cell(0,1), image_path=despues_path)
                doc.add_paragraph()

        code_red_consumo = f"10.{n_tanques+2}"
        add_subtitle(doc, f"{code_red_consumo}. TRABAJOS REALIZADOS EN REDES DE CONSUMO", indent=True)
        actividades = actividades_payload.get(code_red_consumo, [])
        if not actividades:
            doc.add_paragraph("-")
        else:
            for act_i, act in enumerate(actividades, start=1):
                desc = act.get("descripcion", "").strip()
                if desc:
                    doc.add_paragraph(f"(ACTIVIDAD {act_i}: {desc})")
                else:
                    doc.add_paragraph(f"(ACTIVIDAD {act_i})")
                slotkey = code_red_consumo.replace('.', '_')
                antes_path = despues_path = None
                if slotkey in photos_map and "activities" in photos_map[slotkey]:
                    ent = photos_map[slotkey]["activities"].get(act_i, {})
                    antes_path = ent.get("antes")
                    despues_path = ent.get("despues")
                tbl = doc.add_table(rows=1, cols=2)
                tbl.style = "Table Grid"
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                insertar_recuadro_foto_cell(tbl.cell(0,0), image_path=antes_path)
                insertar_recuadro_foto_cell(tbl.cell(0,1), image_path=despues_path)
                doc.add_paragraph()

        # =================================================
        # 11. PUNTO 11, 12, 13 — EVIDENCIA E INFORMES FINALES
        # =================================================
        add_subtitle(doc, "11. EVIDENCIA FOTOGRÁFICA DE LA INSTALACIÓN")
        tabla_11 = doc.add_table(rows=1, cols=1)
        tabla_11.alignment = WD_TABLE_ALIGNMENT.CENTER
        insertar_recuadro_foto_cell(tabla_11.cell(0,0), image_path=None)
        doc.add_paragraph()

        add_subtitle(doc, "12. Conclusiones")
        doc.add_paragraph("-")
        add_subtitle(doc, "13. Recomendaciones")
        doc.add_paragraph("-")

        # =================================================
        # GUARDAR DOCX EN TEMP
        # =================================================
        fd, path_out = tempfile.mkstemp(prefix="Informe_Mantenimiento_", suffix=".docx")
        os.close(fd)
        doc.save(path_out)

        # =================================================
        # LIMPIEZA: eliminar archivos temporales subidos (photos_map)
        # =================================================
        try:
            # remover temporales guardados en photos_map
            if request.content_type.startswith("multipart/form-data"):
                for k, v in photos_map.items():
                    # v puede contener 'images' list o 'activities' dict
                    if isinstance(v, dict):
                        imgs = v.get("images", []) or []
                        for p in imgs:
                            try:
                                os.remove(p)
                            except:
                                pass
                        acts = v.get("activities", {}) or {}
                        for idx_act, act in acts.items():
                            for key_img in ("antes", "despues"):
                                p = act.get(key_img)
                                if p:
                                    try:
                                        os.remove(p)
                                    except:
                                        pass
        except Exception:
            pass

        # =================================================
        # ENVIAR ARCHIVO AL CLIENTE
        # =================================================
        try:
            return send_file(path_out, as_attachment=True, download_name=os.path.basename(path_out))
        except TypeError:
            return send_file(path_out, as_attachment=True)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": "Error interno", "detail": str(e)}), 500

# ============================================
# Fin PARTE 4
# ============================================

# DIME "CONTINÚA" PARA LA PARTE 5/5 (CIERRE, AJUSTES FINALES Y FIN DEL ARCHIVO)
# ============================================
# SERVIDOR.PY — PARTE 5 (FINAL)
# ============================================

from werkzeug.utils import secure_filename

# Límite max para imágenes subidas (5 MB)
MAX_IMAGE_BYTES = 5 * 1024 * 1024


# ============================
# FUNCIÓN PARA INSERTAR IMAGEN O PLACEHOLDER EN UNA CELDA
# ============================
def insertar_recuadro_foto_cell(cell, image_path=None, ancho_cm=15, alto_cm=10):
    """
    Inserta un recuadro con borde y tamaño fijo.
    Si image_path existe, inserta la imagen redimensionada.
    Si no, coloca el texto 'ESPACIO PARA IMAGEN'.
    """
    # limpiar celda
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # tamaño en unidades Word
    width_in = ancho_cm / 2.54
    height_twips = int(alto_cm * 567)

    # set row height
    tr = cell._tc.getparent()
    if hasattr(tr, "get_or_add_trPr"):
        trPr = tr.get_or_add_trPr()
        trHeight = OxmlElement("w:trHeight")
        trHeight.set(qn("w:val"), str(height_twips))
        trHeight.set(qn("w:hRule"), "exact")
        trPr.append(trHeight)

    # borde
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for side in ["top", "left", "bottom", "right"]:
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "12")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")
        borders.append(border)
    tcPr.append(borders)

    if image_path and os.path.exists(image_path):
        try:
            # insertar imagen
            run = p.add_run()
            run.add_picture(image_path, width=Inches(width_in))
            return
        except Exception:
            pass  # si falla imagen → colocar placeholder

    # placeholder
    run = p.add_run("ESPACIO PARA IMAGEN")
    run.font.name = "Calibri"
    run.font.size = Pt(11)
    run.bold = True


# ============================================
# RUTA DE PRUEBA
# ============================================
@app.route("/")
def index():
    try:
        return render_template("pagina.html")
    except Exception:
        return "<h3>Servidor Flask funcionando. Envía POST JSON o multipart a /generar</h3>"


# ============================================
# MAIN — Necesario para Render y también Local
# ============================================
if __name__ == "__main__":
    PORT = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=PORT)
