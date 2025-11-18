# ============================================
# servidor.py - Versión completa y consolidada
# ============================================
# Requiere:
# flask, flask-cors, werkzeug, python-docx, pandas
# Instálalos: pip install flask flask-cors python-docx pandas openpyxl
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
# UTILIDADES DOCX (estilo)
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

    # Inicializar celdas con guion
    for row in table.rows:
        for cell in row.cells:
            set_cell_style(cell, "-", font_size=font_size)

    return table

# =====================================================
# FUNCIÓN PARA INSERTAR FOTO EN RECUADRO (VERSIÓN CONSERVADA)
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
        try:
            p.clear()
        except Exception:
            p.text = ""

    p = cell.paragraphs[0]

    if image_path and os.path.exists(image_path):
        try:
            run = p.add_run()
            run.add_picture(image_path, width=Inches(ancho_in))
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        except Exception:
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

    # Bordes y fondo
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

# =====================================================
# CONSTRUIR DATAFRAMES DESDE JSON
# =====================================================

def build_dfs_from_json(payload):
    """
    Construye DataFrames desde el payload JSON esperado.
    Retorna: df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs
    """
    general = payload.get("general", {}) or {}
    tanques = payload.get("tanques", []) or []
    accesorios_tanque = payload.get("accesoriosTanque", {}) or {}
    accesorios_red = payload.get("accesoriosRed", []) or []
    equipos = payload.get("equipos", []) or []
    observaciones = payload.get("observaciones", {}) or {}

    # df_info
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
    row_info = {c: general.get(c, "") for c in campos}
    df_info = pd.DataFrame([row_info])

    # df_tanques
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

    # df_accesorios: construir filas (Tanque, Atributo, <accesorio columns>)
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
    rows = []
    # accesorios_tanque esperado: { "1": { "Válvula de llenado": {"Marca":..., "Código":...}, ... }, ... }
    for tank_key, accs in accesorios_tanque.items():
        try:
            tk = int(tank_key)
        except Exception:
            tk = tank_key
        for attr in atributos:
            row = {"Tanque": tk, "Atributo": attr}
            for acc_name in accesorios_cols:
                acc_entry = accs.get(acc_name, {}) if isinstance(accs, dict) else {}
                row_val = acc_entry.get(attr, "") if isinstance(acc_entry, dict) else ""
                row[acc_name] = row_val
            rows.append(row)
    df_accesorios = pd.DataFrame(rows)

    # df_red: list of dicts
    df_red = pd.DataFrame(
        [
            {
                "Tipo": r.get("Tipo", ""),
                "Marca": r.get("Marca", ""),
                "Serie": r.get("Serie", ""),
                "Código": r.get("Código", ""),
                "Mes/Año de fabricación": r.get("Mes/Año de fabricación", ""),
            }
            for r in accesorios_red
        ]
    )

    # df_equipos
    df_equipos = pd.DataFrame(equipos)

    # df_obs
    obs_rows = []
    for sp in ["7.1", "7.2", "7.3", "7.4", "7.5"]:
        obs_rows.append({"Subpunto": sp, "Observación": observaciones.get(sp, "")})
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
       [{"code": "9.1", "label": "...", "cantidad": N, "tipo": optional}, ...]
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

    # Accesorios del tanque (varios subtítulos)
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
    for i, t in enumerate(tanques):
        serie = valOrDash(t.get("N° de serie") or t.get("serie"))
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
    tiene_toma = "llenado_toma_desplazada" in df_red_local["Tipo"].astype(str).str.lower().tolist() if not df_red_local.empty else False
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
    zona_bool = "zona_medidores" in df_red_local["Tipo"].astype(str).str.lower().tolist() if not df_red_local.empty else False
    if zona_bool:
        slots.append({
            "code": f"9.{contador}",
            "label": "FOTO DE ZONA MEDIDORES",
            "cantidad": 1
        })
        contador += 1

    # -------- BLOQUE 10 (ACTIVIDADES) --------
    for idx, t in enumerate(tanques, start=1):
        serie = valOrDash(t.get("N° de serie") or t.get("serie"))
        slots.append({
            "code": f"10.{idx}",
            "label": f"TRABAJOS REALIZADOS EN EL TANQUE {idx} DE SERIE: {serie}",
            "cantidad": 0,
            "tipo": "actividad"
        })

    next_code = len(tanques) + 1
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

# =====================================================
# UTIL: PARSEAR ARCHIVOS SUBIDOS
# =====================================================

def parse_uploaded_photos_dict(request_files):
    """
    Espera keys con estos formatos (flexible):
      - photo__9_3__1
      - photo__9_3__img1
      - photo__10_1__act1__antes
      - photo__10_1__act1__despues

    Retorna:
      photos_map = {
        "9_3": { "images":[path,...] },
        "10_1": {
            "activities": { 1: {"antes":path,"despues":path}, ... }
        }
      }
    """
    photos_map = {}
    for key in request_files:
        if not str(key).startswith("photo__"):
            continue
        parts = key.split("__")[1:]
        if len(parts) == 0:
            continue
        slot = parts[0]
        file = request_files.get(key)
        if not file:
            continue
        filename = secure_filename(file.filename)
        fd, tmp_path = tempfile.mkstemp(prefix="img_", suffix="_" + filename)
        os.close(fd)
        file.save(tmp_path)
        try:
            if os.path.getsize(tmp_path) > MAX_IMAGE_BYTES:
                os.remove(tmp_path)
                continue
        except Exception:
            pass

        if len(parts) >= 3 and parts[1].startswith("act"):
            try:
                act_idx = int(parts[1].replace("act", ""))
            except Exception:
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
                act_entry["otros"].append(tmp_path)
            continue

        arr = photos_map.setdefault(slot, {}).setdefault("images", [])
        arr.append(tmp_path)

    return photos_map

# =====================================================
# UTIL: Insertar N recuadros horizontales en el doc con imagenes si existen
# =====================================================

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
    table = doc.add_table(rows=1, cols=max(1, cantidad))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c in range(cantidad):
        cell = table.cell(0, c)
        if c < len(imgs):
            insertar_recuadro_foto_cell(cell, image_path=imgs[c])
        else:
            insertar_recuadro_foto_cell(cell, image_path=None)
    doc.add_paragraph()

# =====================================================
# ENDPOINT: /photo_slots
# =====================================================

@app.post("/photo_slots")
def photo_slots():
    try:
        payload = request.get_json()
        if not payload:
            return jsonify({"error": "Payload vacío"}), 400

        df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs = build_dfs_from_json(payload)
        slots = build_slots_for_frontend(df_tanques, df_red, df_equipos)

        return jsonify({"slots": slots})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# =====================================================
# ENDPOINT: /generar
# =====================================================

@app.post("/generar")
def generar():
    try:
        # 1. DETERMINAR SI ES JSON O FORM-DATA
        if request.content_type and request.content_type.startswith("multipart/form-data"):
            if "payload" not in request.form:
                return jsonify({"error": "No se envió payload en form-data"}), 400
            payload = json.loads(request.form["payload"])
        else:
            payload = request.get_json() or {}

        if not payload:
            return jsonify({"error": "Payload vacío"}), 400

        # Construir DataFrames
        df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs = build_dfs_from_json(payload)

        # Parsear archivos subidos
        photos_map = {}
        if request.content_type and request.content_type.startswith("multipart/form-data"):
            photos_map = parse_uploaded_photos_dict(request.files)

        # Crear documento
        doc = Document()

        # TITULO
        titulo = doc.add_paragraph()
        run = titulo.add_run("INFORME DE MANTENIMIENTO PREVENTIVO Y CUMPLIMIENTO NORMATIVO")
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

        # === 5. Accesorios en redes
        add_subtitle(doc, "5. ACCESORIOS EN REDES")
        df_red_local = df_red.copy() if df_red is not None else pd.DataFrame(
            columns=["Tipo", "Marca", "Serie", "Código", "Mes/Año de fabricación"]
        )
        if "Tipo" in df_red_local.columns:
            df_red_local["Tipo"] = df_red_local["Tipo"].astype(str).str.lower().fillna("")
        else:
            df_red_local["Tipo"] = ""
        # Mapa de accesorios
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

        # === 6. Equipos de la instalación
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
        tabla_8 = doc.add_table(rows=1, cols=1)
        tabla_8.alignment = WD_TABLE_ALIGNMENT.CENTER
        insertar_recuadro_foto_cell(tabla_8.cell(0,0), ancho_cm=15, alto_cm=10, image_path=None)
        doc.add_paragraph()

        # === 9. Evidencia fotográfica de elementos de la instalación ===
        add_subtitle(doc, "9. Evidencia fotográfica de elementos de la instalación")

        # Preparar lista completa (incluye include flag conceptual) y luego insertar
        all_slots = []
        contador_local = 1
        all_slots.append((f"9.{contador_local}", "FOTO PANORÁMICA DE LA ZONA", 1, True)); contador_local += 1

        tanques_for_block = df_tanques.to_dict(orient="records") if df_tanques is not None and not df_tanques.empty else []
        for i, t in enumerate(tanques_for_block):
            serie = valOrDash(t.get("N° de serie") or t.get("serie"))
            all_slots.append((f"9.{contador_local}", f"PLACA DE TANQUE {i+1} DE SERIE: {serie}", 1, True)); contador_local += 1

        for i, t in enumerate(tanques_for_block):
            serie = valOrDash(t.get("N° de serie") or t.get("serie"))
            all_slots.append((f"9.{contador_local}", f"FOTO PANORÁMICA DE ALREDEDORES DE TANQUE {i+1} DE SERIE: {serie}", 4, True)); contador_local += 1

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
        for i, t in enumerate(tanques_for_block):
            serie = valOrDash(t.get("N° de serie") or t.get("serie"))
            for et in etiquetas:
                all_slots.append((f"9.{contador_local}", f"{et} DE TANQUE {i+1} DE SERIE: {serie}", 1, True)); contador_local += 1

        equipos_list = df_equipos.to_dict(orient="records") if df_equipos is not None and not df_equipos.empty else []
        for tipo in ["estabilizador", "quemador", "vaporizador", "tablero", "bomba", "dispensador_de_gas", "decantador", "detector"]:
            lista_eq = [e for e in equipos_list if (e.get("Tipo de equipo") or "").lower() == tipo]
            if lista_eq:
                for eq in lista_eq:
                    serie = valOrDash(eq.get("Serie"))
                    all_slots.append((f"9.{contador_local}", f"FOTO DE PLACA DE {tipo.upper()} DE SERIE: {serie}", 1, True)); contador_local += 1
                    all_slots.append((f"9.{contador_local}", f"FOTO DE {tipo.upper()}", 1, True)); contador_local += 1
            else:
                all_slots.append((f"9.{contador_local}", f"FOTO DE PLACA DE {tipo.upper()} DE SERIE: -", 1, False)); contador_local += 1
                all_slots.append((f"9.{contador_local}", f"FOTO DE {tipo.upper()}", 1, False)); contador_local += 1

        mapa = {
            "llenado_toma_desplazada": "VÁLVULA DE LLENADO TOMA DESPLAZADA",
            "retorno_toma_desplazada": "VÁLVULA DE RETORNO TOMA DESPLAZADA",
            "alivio_hidrostatico": "VÁLVULA DE ALIVIO HIDROSTÁTICO",
            "regulador_primera_etapa": "REGULADOR DE PRIMERA ETAPA",
            "alivio": "VÁLVULA DE ALIVIO",
            "regulador_2da": "REGULADOR DE SEGUNDA ETAPA",
            "pull_away": "VÁLVULA PULL AWAY",
        }
        df_red_local = df_red.copy() if df_red is not None else pd.DataFrame()
        tipos_en_red = df_red_local["Tipo"].astype(str).tolist() if not df_red_local.empty else []
        for clave, nombre in mapa.items():
            include = any(clave in t.lower() for t in tipos_en_red)
            if include:
                matches = [r for _, r in df_red_local.iterrows() if (r.get("Tipo") or "").lower() == clave]
                cantidad = max(1, len(matches))
                for idx in range(cantidad):
                    codigo = valOrDash(matches[idx]["Código"]) if idx < len(matches) else "-"
                    all_slots.append((f"9.{contador_local}", f"FOTO DE {nombre} {idx+1} DE CÓDIGO: {codigo}", 1, True)); contador_local += 1
            else:
                all_slots.append((f"9.{contador_local}", f"FOTO DE {nombre}", 1, False)); contador_local += 1

        zona_med_bool = any("zona_medidores" in str(t).lower() for t in tipos_en_red)
        all_slots.append((f"9.{contador_local}", "FOTO DE ZONA MEDIDORES", 1, zona_med_bool)); contador_local += 1

        # Insertar subtítulos 9.x con sus recuadros o nota
        for code, label, cantidad, include in all_slots:
            slot_key = code.replace('.', '_')
            if include:
                insert_n_recuadros(doc, f"{code}. {label}", slot_key, cantidad, photos_map)
            else:
                add_subtitle(doc, f"{code}. {label}", indent=True)
                add_note(doc, "*NO CUENTA CON DICHO ELEMENTO")
                doc.add_paragraph()

        # === 10. EVIDENCIA FOTOGRÁFICA (MANTENIMIENTO REALIZADO) - actividades dinámicas
        add_subtitle(doc, "10. EVIDENCIA FOTOGRÁFICA (MANTENIMIENTO REALIZADO)")
        add_note(doc, "NOTA 1: SE DEBERÁ MENCIONAR LOS TRABAJOS EJECUTADOS POR TANQUE (INCLUIR LAS INSPECCIONES QUE SE REALICEN)")
        add_note(doc, "NOTA 2: LAS IMÁGENES DEBEN TENER UN TAMAÑO DE 15CM X 10CM MÁXIMO Y SE DEBERÁ VISUALIZAR CLARAMENTE LOS DATOS RELEVANTES")

        actividades_payload = payload.get("actividades", {}) or {}
        n_tanques = len(tanques_for_block)

        # Actividades por tanque
        for idx in range(1, n_tanques+1):
            code = f"10.{idx}"
            titulo = f"{code}. TRABAJOS REALIZADOS EN EL TANQUE {idx} DE SERIE: {valOrDash(tanques_for_block[idx-1].get('N° de serie') or tanques_for_block[idx-1].get('serie'))}"
            add_subtitle(doc, titulo, indent=True)
            actividades = actividades_payload.get(code, [])
            if not actividades:
                doc.add_paragraph("-")
                continue
            for act_i, act in enumerate(actividades, start=1):
                desc = act.get("descripcion", "").strip()
                if desc:
                    doc.add_paragraph(f"(ACTIVIDAD {act_i}: {desc})")
                else:
                    doc.add_paragraph(f"(ACTIVIDAD {act_i})")
                slotkey = code.replace('.', '_')
                antes_path = None
                despues_path = None
                if slotkey in photos_map and "activities" in photos_map[slotkey]:
                    act_entry = photos_map[slotkey]["activities"].get(act_i, {})
                    if act_entry:
                        antes_path = act_entry.get("antes")
                        despues_path = act_entry.get("despues")
                tbl = doc.add_table(rows=1, cols=2)
                tbl.style = "Table Grid"
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                insertar_recuadro_foto_cell(tbl.cell(0,0), image_path=antes_path)
                insertar_recuadro_foto_cell(tbl.cell(0,1), image_path=despues_path)
                doc.add_paragraph()

        # Redes de llenado
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

        # Redes de consumo
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

        # 11, 12, 13
        add_subtitle(doc, "11. EVIDENCIA FOTOGRÁFICA DE LA INSTALACIÓN")
        tabla_11 = doc.add_table(rows=1, cols=1)
        tabla_11.alignment = WD_TABLE_ALIGNMENT.CENTER
        insertar_recuadro_foto_cell(tabla_11.cell(0,0), image_path=None)
        doc.add_paragraph()

        add_subtitle(doc, "12. Conclusiones")
        doc.add_paragraph("-")
        add_subtitle(doc, "13. Recomendaciones")
        doc.add_paragraph("-")

        # Guardar docx en archivo temporal y devolver
        fd, path_out = tempfile.mkstemp(prefix="Informe_Mantenimiento_", suffix=".docx")
        os.close(fd)
        doc.save(path_out)

        # limpieza: eliminar archivos temporales guardados en photos_map
        try:
            if photos_map:
                for k, v in photos_map.items():
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

        try:
            return send_file(path_out, as_attachment=True, download_name=os.path.basename(path_out))
        except TypeError:
            return send_file(path_out, as_attachment=True)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": "Error interno", "detail": str(e)}), 500

# =====================================================
# Ruta simple para debug / index
# =====================================================

@app.route("/")
def index():
    try:
        return render_template("pagina.html")
    except Exception:
        return "<h3>Servidor Flask funcionando. Envía POST JSON o multipart a /generar</h3>"

# =====================================================
# MAIN
# =====================================================

if __name__ == "__main__":
    PORT = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=PORT)

