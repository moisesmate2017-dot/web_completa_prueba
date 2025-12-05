# servidor.py
import os
import tempfile
import unicodedata
from datetime import datetime
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

import pandas as pd
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = Flask(__name__)
CORS(app)

# ---------------------------
# Helpers
# ---------------------------
def normalizar(texto):
    if texto is None:
        return ""
    if not isinstance(texto, str):
        texto = str(texto)
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join([c for c in nfkd if not unicodedata.combining(c)])

def valOrDash(v):
    if v is None:
        return "-"
    if isinstance(v, str) and v.strip()=="":
        return "-"
    return v

# ---------------------------
# DOCX helpers
# ---------------------------
def set_cell_style(cell, text, bold=False, font_size=10, align_center=False):
    cell.text = str(text) if text is not None else "-"
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.name = "Calibri"
            r.font.size = Pt(font_size)
            r.bold = bold
        if align_center:
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        else:
            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    try:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    except Exception:
        pass

def create_table(doc, rows, cols):
    t = doc.add_table(rows=rows, cols=cols)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    return t

def insertar_recuadro_foto(doc, ancho_cm=15, alto_cm=10):
    # creates table 1x1 with "ESPACIO PARA IMAGEN"
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0,0)
    try:
        cell.width = Cm(ancho_cm)
    except:
        pass
    p = cell.paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run("ESPACIO PARA IMAGEN")
    run.bold = True
    # set row height
    tr = tbl.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(int(alto_cm*567)))
    trPr.append(trHeight)
    return tbl

def insertar_foto_doc(doc, img_path, ancho_cm=15, alto_cm=10):
    try:
        doc.add_picture(img_path, width=Cm(ancho_cm), height=Cm(alto_cm))
    except Exception:
        insertar_recuadro_foto(doc, ancho_cm, alto_cm)

# ---------------------------
# Build dataframes from JSON payload
# ---------------------------
def build_dfs_from_json(payload):
    if not payload:
        payload = {}
    df_info = pd.DataFrame([payload.get("general", {})])
    df_tanques = pd.DataFrame(payload.get("tanques", []))
    df_accesorios = pd.DataFrame(payload.get("accesoriosTanque", []))
    df_red = pd.DataFrame(payload.get("accesoriosRed", []))
    df_equipos = pd.DataFrame(payload.get("equipos", []))
    df_obs = pd.DataFrame([payload.get("observaciones", {})])
    return df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs

# ---------------------------
# Matching helper
# ---------------------------
def normalize_key(s):
    return str(s).strip().lower().replace(" ", "_").replace("-", "_")

def find_photos_for_key_or_title(fotos_dict, candidates, title=None):
    found = []
    if not fotos_dict:
        return found
    for c in candidates:
        if c in fotos_dict:
            found.extend(fotos_dict.get(c,[]))
    if not found and title:
        tnorm = normalize_key(title)
        for k,arr in fotos_dict.items():
            kn = normalize_key(k)
            if kn == tnorm or kn in tnorm or tnorm in kn:
                found.extend(arr)
    # unique
    unique = []
    for x in found:
        if x not in unique:
            unique.append(x)
    return unique

# ---------------------------
# GENERADOR DOCX (estructura respetando la tuya)
# ---------------------------
def generar_docx_desde_dfs(df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs, fotos_dict=None):
    if fotos_dict is None:
        fotos_dict = {}

    doc = Document()

    # PORTADA
    p = doc.add_paragraph()
    r = p.add_run("INFORME DE MANTENIMIENTO PERIODICO")
    r.bold = True
    r.font.size = Pt(16)
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    doc.add_paragraph()
    doc.add_paragraph()

    info = df_info.iloc[0] if not df_info.empty else {}

    p = doc.add_paragraph()
    p.add_run(f"CLIENTE: {info.get('cliente', '-')}\n").bold = True
    p.add_run(f"DIRECCIÓN DE INSTALACIÓN: {info.get('direccion', '-')}\n")
    p.add_run(f"RUC: {info.get('ruc', '-')}\n")
    p.add_run(f"FECHA: {info.get('fecha', '-')}\n")

    doc.add_page_break()

    # 1. INFORMACIÓN CLIENTE
    p = doc.add_paragraph(); p.add_run("1. INFORMACIÓN DEL CLIENTE").bold = True
    table = create_table(doc, rows=4, cols=2)
    set_cell_style(table.cell(0,0), "Cliente", True); set_cell_style(table.cell(0,1), info.get("cliente","-"))
    set_cell_style(table.cell(1,0), "Dirección", True); set_cell_style(table.cell(1,1), info.get("direccion","-"))
    set_cell_style(table.cell(2,0), "RUC", True); set_cell_style(table.cell(2,1), info.get("ruc","-"))
    set_cell_style(table.cell(3,0), "Fecha", True); set_cell_style(table.cell(3,1), info.get("fecha","-"))

    doc.add_page_break()

    # 2. TANQUES INSPECCIONADOS
    p = doc.add_paragraph(); p.add_run("2. TANQUES INSPECCIONADOS").bold = True
    if df_tanques is not None and not df_tanques.empty:
        cols = list(df_tanques.columns)
        table = create_table(doc, rows=len(df_tanques)+1, cols=len(cols))
        for ci,c in enumerate(cols):
            set_cell_style(table.cell(0,ci), c, True, align_center=True)
        for ri,row in df_tanques.iterrows():
            for ci,c in enumerate(cols):
                set_cell_style(table.cell(ri+1,ci), valOrDash(row[c]))
    else:
        doc.add_paragraph("No se registraron tanques.")
    doc.add_page_break()

    # 3. ACCESORIOS DE TANQUE
    p = doc.add_paragraph(); p.add_run("3. ACCESORIOS DE TANQUE").bold = True
    if df_accesorios is not None and not df_accesorios.empty:
        cols = list(df_accesorios.columns)
        table = create_table(doc, rows=len(df_accesorios)+1, cols=len(cols))
        for ci,c in enumerate(cols):
            set_cell_style(table.cell(0,ci), c, True, align_center=True)
        for ri,row in df_accesorios.iterrows():
            for ci,c in enumerate(cols):
                set_cell_style(table.cell(ri+1,ci), valOrDash(row[c]))
    else:
        doc.add_paragraph("No se registraron accesorios en tanques.")
    doc.add_page_break()

    # 4. RED DE ACCESORIOS
    p = doc.add_paragraph(); p.add_run("4. RED DE ACCESORIOS").bold = True
    if df_red is not None and not df_red.empty:
        cols = list(df_red.columns)
        table = create_table(doc, rows=len(df_red)+1, cols=len(cols))
        for ci,c in enumerate(cols):
            set_cell_style(table.cell(0,ci), c, True, align_center=True)
        for ri,row in df_red.iterrows():
            for ci,c in enumerate(cols):
                set_cell_style(table.cell(ri+1,ci), valOrDash(row[c]))
    else:
        doc.add_paragraph("No se registraron accesorios de red.")
    doc.add_page_break()

    # 5. EQUIPOS
    p = doc.add_paragraph(); p.add_run("5. EQUIPOS").bold = True
    if df_equipos is not None and not df_equipos.empty:
        cols = list(df_equipos.columns)
        table = create_table(doc, rows=len(df_equipos)+1, cols=len(cols))
        for ci,c in enumerate(cols):
            set_cell_style(table.cell(0,ci), c, True, align_center=True)
        for ri,row in df_equipos.iterrows():
            for ci,c in enumerate(cols):
                set_cell_style(table.cell(ri+1,ci), valOrDash(row[c]))
    else:
        doc.add_paragraph("No se registraron equipos.")
    doc.add_page_break()

    # 6. OBSERVACIONES
    p = doc.add_paragraph(); p.add_run("6. OBSERVACIONES").bold = True
    if df_obs is not None and not df_obs.empty:
        obs = df_obs.iloc[0].to_dict()
        for k,v in obs.items():
            doc.add_paragraph(f"{k}: {valOrDash(v)}")
    else:
        doc.add_paragraph("-")
    doc.add_page_break()

    # 7. EVIDENCIA FOTOGRÁFICA (GENERAL)
    p = doc.add_paragraph(); p.add_run("7. EVIDENCIA FOTOGRÁFICA").bold = True
    insertar_recuadro_foto(doc)

    # 8 & 9: EVIDENCIA DETALLADA / PUNTO 9
    p = doc.add_paragraph(); p.add_run("9. EVIDENCIA FOTOGRÁFICA DE ELEMENTOS DE LA INSTALACIÓN").bold = True

    # Build a list of subtitles like in your original code (we keep it generic)
    subtitulos = []

    # Add general panorama
    subtitulos.append(("panoramica", "9.1 FOTO PANORÁMICA DE LA ZONA", 1))

    # Per-tank subtitles: use df_tanques rows
    tanques = df_tanques.to_dict(orient="records") if (df_tanques is not None and not df_tanques.empty) else []
    for i,t in enumerate(tanques):
        serie = t.get("N° de serie") or t.get("serie") or "-"
        subtitulos.append((f"tanque_{i+1}_placa", f"9.X PLACA DE TANQUE {i+1} - SERIE: {serie}", 1))
        subtitulos.append((f"tanque_{i+1}_panoramica", f"9.X FOTO PANORÁMICA ALREDEDORES TANQUE {i+1}", 4))
        subtitulos.append((f"tanque_{i+1}_base", f"9.X FOTO DE BASES DE CONCRETO TANQUE {i+1}", 1))
        subtitulos.append((f"tanque_{i+1}_manometro_60", f"9.X MANÓMETRO 0-60 PSI TANQUE {i+1}", 1))
        subtitulos.append((f"tanque_{i+1}_manometro_300", f"9.X MANÓMETRO 0-300 PSI TANQUE {i+1}", 1))
        subtitulos.append((f"tanque_{i+1}_chicote", f"9.X CONEXIÓN CHICOTE MULTIVÁLVULA TANQUE {i+1}", 1))
        subtitulos.append((f"tanque_{i+1}_stickers", f"9.X STICKERS TANQUE {i+1}", 1))
        subtitulos.append((f"tanque_{i+1}_anclajes", f"9.X ANCLAJES/PERNOS TANQUE {i+1}", 1))
        subtitulos.append((f"tanque_{i+1}_valv_ll", f"9.X VÁLVULA DE LLENADO TANQUE {i+1}", 1))
        subtitulos.append((f"tanque_{i+1}_valv_seg", f"9.X VÁLVULA DE SEGURIDAD TANQUE {i+1}", 1))
        subtitulos.append((f"tanque_{i+1}_valv_drena", f"9.X VÁLVULA DE DRENAJE TANQUE {i+1}", 1))
        subtitulos.append((f"tanque_{i+1}_multivalvula", f"9.X MULTIVÁLVULA TANQUE {i+1}", 1))
        subtitulos.append((f"tanque_{i+1}_medidor", f"9.X MEDIDOR DE PORCENTAJE TANQUE {i+1}", 1))

    # Equipos: iterate df_equipos if present
    equipos = df_equipos.to_dict(orient="records") if (df_equipos is not None and not df_equipos.empty) else []
    for i,eq in enumerate(equipos):
        serie = eq.get("Serie") or eq.get("serie") or "-"
        tipo = eq.get("Tipo de equipo") or eq.get("Tipo") or "EQUIPO"
        subtitulos.append((f"equipo_{i+1}_placa", f"9.X PLACA DE {tipo} {i+1} - SERIE: {serie}", 1))
        subtitulos.append((f"equipo_{i+1}_foto", f"9.X FOTO DE {tipo} {i+1}", 1))

    # Now render subtitles: insert photos if photos exist in fotos_dict, else recuadros
    for sid, title, count in subtitulos:
        p = doc.add_paragraph(); p.add_run(title).bold = True
        fotos = fotos_dict.get(sid, [])
        if not fotos:
            # try fuzzy matching by normalized key or title
            fotos = find_photos_for_key_or_title(fotos_dict, [sid], title=title)
        if fotos:
            for img in fotos:
                insertar_foto_doc(doc, img, ancho_cm=15, alto_cm=10)
        else:
            # insert requested number of recuadros
            for _ in range(count):
                insertar_recuadro_foto(doc, ancho_cm=15, alto_cm=10)
        doc.add_paragraph()

    # 10. trabajos realizados - ejemplo simple por tanque
    p = doc.add_paragraph(); p.add_run("10. EVIDENCIA DE TRABAJOS REALIZADOS").bold = True
    for i,t in enumerate(tanques, start=1):
        serie = t.get("N° de serie") or t.get("serie") or "-"
        add_sub = doc.add_paragraph(); add_sub.add_run(f"10.{i} TRABAJOS EN TANQUE {i} - SERIE: {serie}").bold = True
        # two before/after pairs
        insertar_recuadro_foto(doc); insertar_recuadro_foto(doc)
        insertar_recuadro_foto(doc); insertar_recuadro_foto(doc)
        doc.add_paragraph()

    # Save docx
    tmpdir = tempfile.mkdtemp(prefix="informe_")
    filename = f"INFORME_MANTENIMIENTO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    output_path = os.path.join(tmpdir, filename)
    doc.save(output_path)
    return output_path

# ---------------------------
# /generar handler (compatible JSON + multipart)
# ---------------------------
@app.route("/generar", methods=["POST"])
def generar():
    """
    Endpoint compatible con tu flujo original.
    - Accepts JSON body (application/json)
    - Accepts multipart/form-data with field 'json' (string) and files 'files' (repeated)
      File names must be: <subId>__<idx>__<origName>
    """
    try:
        fotos_dict = {}
        data = None
        content_type = request.content_type or ""
        if content_type.startswith("multipart/form-data"):
            json_str = request.form.get('json')
            if not json_str:
                return jsonify({'error':'Missing json form field'}), 400
            try:
                import json as _json
                data = _json.loads(json_str)
            except Exception:
                data = pd.io.json.loads(json_str)

            files = request.files.getlist('files')
            if not files:
                files = list(request.files.values())

            if files:
                tmpd = tempfile.mkdtemp(prefix='fotos_')
                for f in files:
                    fname = f.filename or 'file.jpg'
                    parts = fname.split('__',2)
                    subId = parts[0] if parts else 'otros'
                    safe = os.path.basename(fname)
                    dest = os.path.join(tmpd, safe)
                    f.save(dest)
                    fotos_dict.setdefault(subId, []).append(dest)
        else:
            # JSON normal
            data = request.get_json()
            fotos_dict = {}

        # Build DataFrames
        try:
            df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs = build_dfs_from_json(data)
        except Exception:
            df_info = pd.DataFrame([data.get("general", {})])
            df_tanques = pd.DataFrame(data.get("tanques", []))
            df_accesorios = pd.DataFrame(data.get("accesoriosTanque", []))
            df_red = pd.DataFrame(data.get("accesoriosRed", []))
            df_equipos = pd.DataFrame(data.get("equipos", []))
            df_obs = pd.DataFrame([data.get("observaciones", {})])

        # Generate docx
        try:
            output_path = generar_docx_desde_dfs(df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs, fotos_dict)
        except TypeError:
            output_path = generar_docx_desde_dfs(df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs)

        return send_file(output_path, as_attachment=True)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

# Ensure server uses PORT env (Render)
if __name__ == "__main__":
    import os
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)), debug=True)

