#!/usr/bin/env python3
# servidor.py - Backend para Solgas (bloques fotos + generar .docx)
# Requiere: Flask, python-docx, Pillow, werkzeug
from flask import Flask, request, jsonify, send_file
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Cm, Pt
from PIL import Image
import os, io, json, tempfile, traceback

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 300 * 1024 * 1024  # 300MB
UPLOAD_TMP = "/tmp/solgas_uploads"
os.makedirs(UPLOAD_TMP, exist_ok=True)

def safe_str(x, default='-'):
    if x is None: return default
    s = str(x).strip()
    return s if s else default

# -----------------------
# Generador de bloques (mismo algoritmo que frontend)
# Recibe payload (general/tanks/accesorios/equipos/actividades)
# -----------------------
def build_default_bloques_from_payload(payload):
    bloques = []
    order = 0
    general = payload.get("general") or payload.get("generalInfo") or {}
    tanks = payload.get("tanques", []) or payload.get("tanks", [])
    accesoriosByTank = payload.get("accesoriosTanque", {}) or payload.get("accesoriosByTank", {}) or {}
    redes = payload.get("accesoriosRed", []) or payload.get("redes", [])
    equipos = payload.get("equipos", []) or payload.get("equipos", [])
    actividades = payload.get("actividades", {}) or payload.get("activities", {})

    def push(titulo, clave, fotos=1, aplica=True):
        nonlocal order
        order += 1
        bloques.append({"titulo": titulo, "clave": clave, "fotos": fotos, "aplica": aplica, "order": order})

    # 8 y 9 básicos
    push("8. Evidencia fotográfica (del establecimiento) - Panorámica", "foto_8_1", 1, True)
    push("9.1 FOTO PANORÁMICA DE LA ZONA", "foto_9_panoramica", 1, True)

    # por tanque (9.x)
    fixed_titles = [
      "FOTO DE BASES DE CONCRETO",
      "FOTO DE MANÓMETROS 0-60 PSI",
      "FOTO DE MANÓMETROS 0-300 PSI",
      "FOTO DE CONEXIÓN DE CHICOTE A LA MULTIVÁLVULA",
      "STICKERS DEL TANQUE Y PINTADO",
      "FOTO DE LOS 04 ANCLAJES, PERNOS, TORNILLOS",
      "FOTO DE VÁLVULA DE LLENADO",
      "FOTO DE VÁLVULA DE SEGURIDAD",
      "FOTO DE VÁLVULA DE DRENAJE",
      "FOTO DE VÁLVULA DE MULTIVÁLVULA",
      "FOTO DE VÁLVULA DE MEDIDOR DE PORCENTAJE"
    ]
    for idx, t in enumerate(tanks):
        i = idx+1
        serie = safe_str(t.get("serie"))
        push(f"PLACA DE TANQUE {i} - Serie: {serie}", f"foto_9_placa_tank_{i}", 1, True)
        push(f"FOTO PANORÁMICA DE ALREDEDORES DE TANQUE {i} - Serie: {serie}", f"foto_9_panoramica_tank_{i}", 4, True)
        for j, title in enumerate(fixed_titles, start=1):
            push(f"9.{j} {title} - Tanque {i} - Serie: {serie}", f"foto_9_tank{i}_{j}", 1, True)

    # Equipos: placas y general. Si no hay equipos, se crean placeholders (aplica False)
    tipos_encontrados = {}
    for eq in equipos:
        tipo = eq.get("Tipo de equipo") or eq.get("tipo") or "equipo"
        tipos_encontrados.setdefault(tipo, []).append(eq)
    if tipos_encontrados:
        for tipo, lista in tipos_encontrados.items():
            for idx, eq in enumerate(lista, start=1):
                serie = safe_str(eq.get("Serie") or eq.get("serie"))
                push(f"FOTO DE PLACA DE {tipo.upper()} {idx} - Serie: {serie}", f"foto_9_{tipo}_placa_{idx}", 1, True)
                push(f"FOTO DE {tipo.upper()} {idx} (general)", f"foto_9_{tipo}_general_{idx}", 1, True)
    else:
        # placeholder generic (no aplica)
        push("FOTO DE PLACA DE EQUIPO - Serie: -", "foto_9_equipo_placa_1", 1, False)
        push("FOTO DE EQUIPO (general)", "foto_9_equipo_general_1", 1, False)

    # toma desplazada y accesorios
    for key, title in [
        ("foto_9_toma_transferencia", "FOTO DEL PUNTO DE TRANSFERENCIA DESPLAZADO"),
        ("foto_9_toma_caja", "FOTO DE LA CAJA DE LA TOMA DESPLAZADA"),
        ("foto_9_toma_recorrido", "FOTO DEL RECORRIDO DESDE TOMA DESPLAZADA HASTA TANQUE")
    ]:
        push(title, key, 1, True)

    # mapa accesorios en redes
    mapa = {
        "llenado_toma_desplazada": "VÁLVULA DE LLENADO TOMA DESPLAZADA",
        "retorno_toma_desplazada": "VÁLVULA DE RETORNO TOMA DESPLAZADA",
        "alivio": "VÁLVULA DE ALIVIO",
        "regulador_2da": "REGULADOR DE SEGUNDA ETAPA",
        "pull_away": "VÁLVULA PULL AWAY",
        "alivio_hidrostatico": "VÁLVULA DE ALIVIO HIDROSTÁTICO",
        "regulador_primera_etapa": "REGULADOR DE PRIMERA ETAPA"
    }
    for clave_map, readable in mapa.items():
        lista = [r for r in redes if (r.get("Tipo")==clave_map or r.get("tipo")==clave_map)]
        cantidad = max(1, len(lista))
        for idx in range(cantidad):
            codigo = lista[idx].get("Código") if idx < len(lista) else None
            aplica = len(lista) > 0
            push(f"FOTO DE {readable} {idx+1} - Código: {safe_str(codigo,'-')}", f"foto_9_acc_{clave_map}_{idx+1}", 1, aplica)

    push("FOTO DE ZONA MEDIDORES", "foto_9_zona_medidores", 1, True)
    push("11. EVIDENCIA FOTOGRÁFICA DE LA INSTALACIÓN", "foto_11_1", 1, True)

    # PUNTO 10: Trabajos realizados (dinámicos) - por tanque y redes
    actividades_obj = actividades or {}
    # por tanque
    for idx, t in enumerate(tanks):
        i = idx+1
        serie = safe_str(t.get("serie"))
        # posibles llaves donde el frontend puede mandar actividades: "tank1", "tank_1", "tanque1", "1", etc.
        actividades_tank = None
        # check common possibilities
        for key_candidate in (f"tank{i}", f"tank_{i}", f"tanque{i}", f"tanque_{i}", str(i)):
            if key_candidate in actividades_obj:
                actividades_tank = actividades_obj[key_candidate]
                break
        # check numeric as list
        if actividades_tank is None:
            actividades_tank = actividades_obj.get("tanks", {}).get(str(i)) if isinstance(actividades_obj.get("tanks"), dict) else None
        # fallback if top-level activities is dict keyed by indexes as strings
        if actividades_tank is None:
            actividades_tank = actividades_obj.get(str(i), [])
        if actividades_tank and isinstance(actividades_tank, list) and len(actividades_tank)>0:
            for a_idx, act in enumerate(actividades_tank, start=1):
                desc = safe_str(act.get("descripcion") or act.get("desc") or f"Actividad {a_idx}")
                push(f"10. Tanque {i} - {desc} - ANTES - Serie: {serie}", f"foto_10_tank{i}_act{a_idx}__antes", 1, True)
                push(f"10. Tanque {i} - {desc} - DESPUÉS - Serie: {serie}", f"foto_10_tank{i}_act{a_idx}__despues", 1, True)
        else:
            push(f"10. Tanque {i} - Trabajos: No se realizaron actividades (o no ameritó)", f"foto_10_tank{i}_no_actividades", 0, True)

    # trabajos en redes (llenado_retorno y consumo)
    redes_works = [("red_llenado_retorno","Red - Llenado/Retorno"), ("red_consumo","Red - Consumo")]
    for rw_key, rw_label in redes_works:
        acts_rw = actividades_obj.get(rw_key) or actividades_obj.get(rw_key.replace("red_","")) or actividades_obj.get(rw_key.replace("red_","")) or []
        if acts_rw and isinstance(acts_rw, list) and len(acts_rw)>0:
            for a_idx, act in enumerate(acts_rw, start=1):
                desc = safe_str(act.get("descripcion") or act.get("desc") or f"Actividad {a_idx}")
                push(f"10. {rw_label} - {desc} - ANTES", f"foto_10_{rw_key}_act{a_idx}__antes", 1, True)
                push(f"10. {rw_label} - {desc} - DESPUÉS", f"foto_10_{rw_key}_act{a_idx}__despues", 1, True)
        else:
            push(f"10. {rw_label} - No se realizaron actividades", f"foto_10_{rw_key}_no_actividades", 0, True)

    return bloques

@app.route("/bloques_fotos", methods=["POST"])
def bloques_fotos():
    try:
        payload = request.get_json(force=True)
    except Exception as e:
        return jsonify({"error":"JSON inválido","detail":str(e)}), 400
    bloques = build_default_bloques_from_payload(payload)
    return jsonify(bloques), 200

def save_temp_file(file_storage, prefix="f"):
    filename = secure_filename(file_storage.filename or "upload")
    fd, path = tempfile.mkstemp(prefix=prefix + "_", suffix="_"+filename, dir=UPLOAD_TMP)
    os.close(fd)
    file_storage.save(path)
    return path

@app.route("/generate_docx", methods=["POST"])
def generate_docx():
    # validate payload
    if 'payload' not in request.form:
        return jsonify({"error":"No se recibió campo 'payload' en form-data"}), 400
    try:
        payload = json.loads(request.form['payload'])
    except Exception as e:
        return jsonify({"error":"Payload JSON inválido","detail":str(e)}), 400

    # Recolectar archivos del request (puede haber claves foto_x / foto_x__1 / foto_10_...)
    files_map = {}
    for key in request.files:
        # getlist para permitir múltiples archivos por la misma clave
        files_map[key] = request.files.getlist(key)

    # Crear documento
    doc = Document()
    # Normalizar font
    try:
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Arial'
        font.size = Pt(10)
    except Exception:
        pass
    doc.add_heading('Informe de Inspección - Solgas', level=1)

    # Información general
    gen = payload.get("general") or payload.get("generalInfo") or {}
    if gen:
        p = doc.add_paragraph()
        p.add_run("Información general:\n").bold = True
        for k, v in gen.items():
            p.add_run(f"{k}: ").bold = True
            p.add_run(f"{v}\n")

    # Tanques y accesorios
    tanques = payload.get("tanques", []) or payload.get("tanks", [])
    if tanques:
        doc.add_heading("Tanques", level=2)
        for idx, t in enumerate(tanques, start=1):
            doc.add_paragraph(f"Tanque {idx} — Serie: {safe_str(t.get('serie'))} — Tipo: {safe_str(t.get('tipo'))}")
            acc_map = (payload.get("accesoriosTanque") or payload.get("accesoriosByTank") or {}).get(str(idx), {})
            if acc_map:
                tbl = doc.add_table(rows=1, cols=4)
                hdr = tbl.rows[0].cells
                hdr[0].text = "Accesorio"; hdr[1].text = "Marca"; hdr[2].text = "Código"; hdr[3].text = "Serie / Mes-Año"
                for acc_name, fields in acc_map.items():
                    row = tbl.add_row().cells
                    row[0].text = acc_name
                    row[1].text = safe_str(fields.get("Marca"))
                    row[2].text = safe_str(fields.get("Código"))
                    row[3].text = f"{safe_str(fields.get('Serie'))} / {safe_str(fields.get('Mes/Año de fabricación'))}"
            doc.add_paragraph()

    # Accesorios en redes
    redes = payload.get("accesoriosRed", []) or payload.get("redes", [])
    if redes:
        doc.add_heading("Accesorios en redes", level=2)
        tbl = doc.add_table(rows=1, cols=4)
        hdr = tbl.rows[0].cells
        hdr[0].text = "Tipo"; hdr[1].text = "Marca"; hdr[2].text = "Serie"; hdr[3].text = "Código / Mes-Año"
        for r in redes:
            row = tbl.add_row().cells
            row[0].text = safe_str(r.get("Tipo"))
            row[1].text = safe_str(r.get("Marca"))
            row[2].text = safe_str(r.get("Serie"))
            row[3].text = f"{safe_str(r.get('Código'))} / {safe_str(r.get('Mes/Año de fabricación'))}"
        doc.add_paragraph()

    # Equipos
    equipos = payload.get("equipos", []) or payload.get("equipos", [])
    if equipos:
        doc.add_heading("Equipos", level=2)
        for eq in equipos:
            info = ", ".join([f"{k}: {v}" for k, v in eq.items() if v])
            doc.add_paragraph(info)

    # Observaciones
    obs = payload.get("observaciones") or payload.get("obs") or {}
    if obs:
        doc.add_heading("Observaciones", level=2)
        for k, v in obs.items():
            doc.add_paragraph(f"{k}: {safe_str(v, '')}")

    # BLOQUES FOTOGRÁFICOS
    bloques = build_default_bloques_from_payload(payload)
    doc.add_page_break()
    doc.add_heading("Evidencias fotográficas", level=2)

    # Flatten file keys so clave and clave__N both map
    flattened = {}
    for key, flist in files_map.items():
        flattened.setdefault(key, []).extend(flist)
        # if key contains '__', also map base
        if "__" in key:
            base = key.split("__")[0]
            flattened.setdefault(base, []).extend(flist)

    # helper: insert image into a cell with forced size 15x10 cm.
    def insert_image_in_cell(cell, file_storage):
        try:
            tmp_path = save_temp_file(file_storage, prefix="img")
            # ensure image mode and size acceptable (Pillow)
            try:
                with Image.open(tmp_path) as im:
                    # convert to RGB if needed
                    if im.mode in ("RGBA","P"):
                        bg = Image.new("RGB", im.size, (255,255,255))
                        bg.paste(im, mask=im.split()[3] if im.mode=="RGBA" else None)
                        bg.save(tmp_path, format="JPEG")
                    # optionally can resize preserving aspect to fit within 15x10 cm while keeping aspect
            except Exception:
                pass
            paragraph = cell.paragraphs[0]
            run = paragraph.add_run()
            run.add_picture(tmp_path, width=Cm(15), height=Cm(10))
            try: os.remove(tmp_path)
            except: pass
            return True
        except Exception as e:
            print("Error insert image:", e)
            traceback.print_exc()
            return False

    for block in bloques:
        titulo = block.get("titulo")
        clave = block.get("clave")
        fotos_needed = block.get("fotos", 1)
        aplica = block.get("aplica", True)
        doc.add_heading(titulo, level=3)
        if fotos_needed == 0:
            doc.add_paragraph("(Sin imágenes solicitadas para este subtítulo).")
            continue
        # Gather files for this clave
        found_files = []
        # exact key
        if clave in files_map:
            found_files = files_map[clave]
        else:
            # collect keys starting with clave (e.g., clave__1, clave__2, clave_act..., etc)
            candidates = [k for k in files_map.keys() if k.startswith(clave)]
            for kc in candidates:
                found_files.extend(files_map.get(kc, []))
            # flattened fallback
            if not found_files and clave in flattened:
                found_files = flattened[clave]
        if not found_files:
            doc.add_paragraph("No se subió fotografía para este subtítulo.")
            continue
        # layout: try to place up to fotos_needed per row
        cols = min(fotos_needed, max(1, len(found_files)))
        rows = (len(found_files) + cols -1)//cols
        tbl = doc.add_table(rows=rows, cols=cols)
        f_idx = 0
        for r in range(rows):
            cells = tbl.rows[r].cells
            for c in range(cols):
                if f_idx < len(found_files):
                    insert_image_in_cell(cells[c], found_files[f_idx])
                else:
                    cells[c].text = ""
                f_idx += 1
        doc.add_paragraph()

    # DETALLE DE ACTIVIDADES (10.X) - mostrar ANTES / DESPUÉS en tabla
    doc.add_page_break()
    doc.add_heading("Trabajos realizados (detalle ANTES / DESPUÉS)", level=2)
    actividades_obj = payload.get("actividades", {}) or payload.get("activities", {})
    if actividades_obj and isinstance(actividades_obj, dict) and len(actividades_obj)>0:
        # iterate keys in actividades_obj: can be tank1, tank_1, red_..., or numeric keys
        for key, acts in actividades_obj.items():
            # acts is expected to be a list of activities for that key
            doc.add_heading(f"Actividades — {key}", level=3)
            if not acts:
                doc.add_paragraph("No se registraron actividades.")
                continue
            for a_idx, act in enumerate(acts, start=1):
                desc = safe_str(act.get("descripcion") or act.get("desc") or f"Actividad {a_idx}")
                doc.add_paragraph(f"- {desc}")
                # expected file keys for before/after as in frontend/backend conventions
                before_keys = [
                    f"foto_10_{key}_act{a_idx}__antes",
                    f"foto_10_{key}_act{a_idx}__1",
                    f"foto_10_{key}__antes",
                    f"foto_10_{key}__1",
                    f"foto_10_{key}_act{a_idx}__1"
                ]
                after_keys = [
                    f"foto_10_{key}_act{a_idx}__despues",
                    f"foto_10_{key}_act{a_idx}__2",
                    f"foto_10_{key}__despues",
                    f"foto_10_{key}__2",
                    f"foto_10_{key}_act{a_idx}__2"
                ]
                before_files = []
                after_files = []
                for k in before_keys:
                    if k in files_map: before_files.extend(files_map[k])
                for k in after_keys:
                    if k in files_map: after_files.extend(files_map[k])
                # As fallback, check flattened base names
                base_before = f"foto_10_{key}_act{a_idx}"
                if not before_files:
                    for k in flattened.get(base_before, []):
                        before_files.append(k)
                base_after = f"foto_10_{key}_act{a_idx}"
                if not after_files:
                    for k in flattened.get(base_after, []):
                        after_files.append(k)
                # Create 1-row table with 2 cols: Antes / Después
                tbl = doc.add_table(rows=1, cols=2)
                hdr = tbl.rows[0].cells
                hdr[0].text = "ANTES"
                hdr[1].text = "DESPUÉS"
                row = tbl.add_row().cells
                # insert first before + first after if exist; else put text
                if before_files:
                    insert_image_in_cell(row[0], before_files[0])
                else:
                    row[0].text = "(No hubo fotografía ANTES)"
                if after_files:
                    insert_image_in_cell(row[1], after_files[0])
                else:
                    row[1].text = "(No hubo fotografía DESPUÉS)"
                doc.add_paragraph()
    else:
        doc.add_paragraph("No se registraron actividades dinámicas en el payload.")

    # finalize doc
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return send_file(
        bio,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='informe_generado.docx'
    )

@app.route("/", methods=["GET"])
def index():
    return jsonify({"ok": True, "note": "Servidor Solgas — endpoints: POST /bloques_fotos , POST /generate_docx"}), 200

if __name__ == "__main__":
    # production: use gunicorn/uWSGI; this is dev
    app.run(host="0.0.0.0", port=5000, debug=True)
