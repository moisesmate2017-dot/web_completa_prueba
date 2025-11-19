#!/usr/bin/env python3
# servidor.py - Backend ampliado para Solgas (bloques fotos + generar .docx + servir pagina.html)
# Requisitos: Flask, python-docx, Pillow
# python -m pip install flask python-docx pillow

from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Cm, Pt, Inches
from PIL import Image
import os, io, json, math, tempfile

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=None
)

# Config
app.config['MAX_CONTENT_LENGTH'] = 400 * 1024 * 1024  # 400MB
UPLOAD_TMP = os.path.join(tempfile.gettempdir(), "solgas_uploads")
os.makedirs(UPLOAD_TMP, exist_ok=True)


def safe_str(x, default='-'):
    if x is None:
        return default
    s = str(x).strip()
    return s if s else default


def save_temp_file(file_storage, prefix="f"):
    """
    Guarda temporalmente un FileStorage en disco y retorna la ruta.
    """
    filename = secure_filename(file_storage.filename or "upload")
    path = os.path.join(UPLOAD_TMP, f"{prefix}_{filename}")
    file_storage.save(path)
    return path


def compute_picture_size_for_docx(img_path, max_cm_w=15.0, max_cm_h=10.0):
    """
    Devuelve (width, height) en unidades docx.shared (Cm) manteniendo la proporción de la imagen
    y maximizándola dentro del recuadro de max_cm_w x max_cm_h.
    """
    try:
        with Image.open(img_path) as im:
            w_px, h_px = im.size
            # DPI fallback (some images don't include DPI)
            dpi = im.info.get('dpi', (96, 96))
            dpi_x = dpi[0] if isinstance(dpi, (list, tuple)) else dpi
            # convert px -> cm using dpi: cm = px / dpi * 2.54
            w_cm = (w_px / dpi_x) * 2.54
            h_cm = (h_px / dpi_x) * 2.54
            # if DPI is nonsense (0), fallback to ratio only
            if dpi_x <= 1:
                ratio = w_px / max(1, h_px)
                # assume some baseline and scale by ratio -> but simpler: use px ratio
                w_cm = w_px / 37.7952755906  # px->cm using 96dpi approx
                h_cm = h_px / 37.7952755906
            # Scale to fit within max_cm_w x max_cm_h while keeping aspect ratio
            scale = min(max_cm_w / max(1e-9, w_cm), max_cm_h / max(1e-9, h_cm), 1.0)
            target_w_cm = w_cm * scale
            target_h_cm = h_cm * scale
            # Ensure not zero
            if target_w_cm <= 0 or target_h_cm <= 0:
                target_w_cm, target_h_cm = max_cm_w, max_cm_h
            return Cm(target_w_cm), Cm(target_h_cm)
    except Exception as e:
        # Fallback: return the box size
        return Cm(max_cm_w), Cm(max_cm_h)


# -----------------------
# Generador de bloques (frontend/backend deben coincidir)
# -----------------------
def build_default_bloques_from_payload(payload):
    bloques = []
    order = 0

    general = payload.get("general") or payload.get("generalInfo") or {}
    tanks = payload.get("tanques", []) or payload.get("tanks", [])
    accesoriosByTank = payload.get("accesoriosTanque", {}) or payload.get("accesoriosByTank", {}) or {}
    redes = payload.get("accesoriosRed", []) or payload.get("redes", []) or payload.get("networks", [])
    equipos = payload.get("equipos", []) or payload.get("equipos", [])
    actividades = payload.get("actividades", {}) or payload.get("activities", {}) or payload.get("actividades_por_tanque", {})

    def push(titulo, clave, fotos=1, aplica=True):
        nonlocal order
        order += 1
        bloques.append({"titulo": titulo, "clave": clave, "fotos": fotos, "aplica": aplica, "order": order})

    # 8 & 9 generales
    push("8. Evidencia fotográfica (del establecimiento) - Panorámica", "foto_8_1", 1, True)
    push("9.1 FOTO PANORÁMICA DE LA ZONA", "foto_9_panoramica", 1, True)

    # por tanque (9.x) - placa + panorámicas + lista fija
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
        i = idx + 1
        serie = safe_str(t.get("serie"))
        push(f"PLACA DE TANQUE {i} - Serie: {serie}", f"foto_9_placa_tank_{i}", 1, True)
        push(f"FOTO PANORÁMICA DE ALREDEDORES DE TANQUE {i} - Serie: {serie}", f"foto_9_panoramica_tank_{i}", 4, True)
        for j, title in enumerate(fixed_titles, start=1):
            push(f"9.{j} {title} - Tanque {i} - Serie: {serie}", f"foto_9_tank{i}_{j}", 1, True)

    # Equipos (placa + general por equipo) - si no hay equipos, placeholders no aplica
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
        push("FOTO DE PLACA DE EQUIPO - Serie: -", "foto_9_equipo_placa_1", 1, False)
        push("FOTO DE EQUIPO (general)", "foto_9_equipo_general_1", 1, False)

    # toma desplazada + recorrido
    push("FOTO DEL PUNTO DE TRANSFERENCIA DESPLAZADO", "foto_9_toma_transferencia", 1, True)
    push("FOTO DE LA CAJA DE LA TOMA DESPLAZADA", "foto_9_toma_caja", 1, True)
    push("FOTO DEL RECORRIDO DESDE TOMA DESPLAZADA HASTA TANQUE", "foto_9_toma_recorrido", 1, True)

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
        lista = [r for r in redes if (r.get("Tipo") == clave_map or r.get("tipo") == clave_map)]
        cantidad = max(1, len(lista))
        for idx in range(cantidad):
            codigo = lista[idx].get("Código") if idx < len(lista) else None
            push(f"FOTO DE {readable} {idx+1} - Código: {safe_str(codigo,'-')}",
                 f"foto_9_acc_{clave_map}_{idx+1}",
                 1,
                 True if len(lista) > 0 else False)

    push("FOTO DE ZONA MEDIDORES", "foto_9_zona_medidores", 1, True)
    push("11. EVIDENCIA FOTOGRÁFICA DE LA INSTALACIÓN", "foto_11_1", 1, True)

    # ---------------- Punto 10: Actividades dinámicas por tanque y por redes ----------------
    actividades_obj = actividades or {}

    # Por tanque: se permite múltiples actividades por tanque
    for idx, t in enumerate(tanks):
        i = idx + 1
        serie = safe_str(t.get("serie"))
        # buscar en distintos nombres (compatibilidad)
        possible_keys = [f"tank{i}", f"tank_{i}", f"tanque{i}", f"tanque_{i}", str(i)]
        activities_for_tank = None
        for k in possible_keys:
            if k in actividades_obj:
                activities_for_tank = actividades_obj[k]
                break
        # También se acepta estructura actividades: { "tanks": { "1": [ ... ] } }
        if not activities_for_tank:
            by_index = actividades_obj.get("tanks") or actividades_obj.get("tanques")
            if isinstance(by_index, dict) and str(i) in by_index:
                activities_for_tank = by_index[str(i)]

        if activities_for_tank and isinstance(activities_for_tank, list) and len(activities_for_tank) > 0:
            for a_idx, act in enumerate(activities_for_tank, start=1):
                desc = safe_str(act.get("descripcion") or act.get("desc") or f"Actividad {a_idx}")
                push(f"10. Tanque {i} - {desc} - ANTES - Serie: {serie}",
                     f"foto_10_tank{i}_act{a_idx}__antes", 1, True)
                push(f"10. Tanque {i} - {desc} - DESPUÉS - Serie: {serie}",
                     f"foto_10_tank{i}_act{a_idx}__despues", 1, True)
        else:
            # Placeholder: no actividades
            push(f"10. Tanque {i} - Trabajos: No se realizaron actividades (o no ameritó)",
                 f"foto_10_tank{i}_no_actividades", 0, True)

    # Redes trabajos: red_llenado_retorno y red_consumo
    redes_work_keys = ["red_llenado_retorno", "red_consumo", "llenado_retorno", "consumo"]
    # prefer explicit keys if provided
    for rw in ["red_llenado_retorno", "red_consumo"]:
        acts_rw = actividades_obj.get(rw) or actividades_obj.get(rw.replace("red_", "")) or []
        if acts_rw and isinstance(acts_rw, list) and len(acts_rw) > 0:
            for a_idx, act in enumerate(acts_rw, start=1):
                desc = safe_str(act.get("descripcion") or act.get("desc") or f"Actividad {a_idx}")
                push(f"10. Redes {rw} - {desc} - ANTES", f"foto_10_{rw}_act{a_idx}__antes", 1, True)
                push(f"10. Redes {rw} - {desc} - DESPUÉS", f"foto_10_{rw}_act{a_idx}__despues", 1, True)
        else:
            push(f"10. Redes {rw} - No se realizaron actividades", f"foto_10_{rw}_no_actividades", 0, True)

    return bloques


# -------------------------
# Rutas
# -------------------------
@app.route("/", methods=["GET"])
def index():
    # sirve tu frontend en templates/pagina.html
    return render_template("pagina.html")


@app.route("/bloques_fotos", methods=["POST"])
def bloques_fotos():
    try:
        payload = request.get_json(force=True)
    except Exception as e:
        return jsonify({"error": "JSON inválido", "detail": str(e)}), 400
    bloques = build_default_bloques_from_payload(payload)
    return jsonify(bloques), 200


@app.route("/generate_docx", methods=["POST"])
def generate_docx():
    # payload (form-data)
    if 'payload' not in request.form:
        return jsonify({"error": "No se recibió campo 'payload' en form-data"}), 400
    try:
        payload = json.loads(request.form['payload'])
    except Exception as e:
        return jsonify({"error": "Payload JSON inválido", "detail": str(e)}), 400

    # archivos recibidos
    files_map = {}
    for key in request.files:
        files_map[key] = request.files.getlist(key)

    # construir documento
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = "Arial"
    font.size = Pt(10)
    doc.add_heading("Informe de Inspección - Solgas", level=1)

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
                hdr[0].text = "Accesorio"
                hdr[1].text = "Marca"
                hdr[2].text = "Código"
                hdr[3].text = "Serie / Mes-Año"
                for acc_name, fields in acc_map.items():
                    row = tbl.add_row().cells
                    row[0].text = acc_name
                    row[1].text = safe_str(fields.get("Marca"))
                    row[2].text = safe_str(fields.get("Código"))
                    row[3].text = f"{safe_str(fields.get('Serie'))} / {safe_str(fields.get('Mes/Año de fabricación'))}"
            doc.add_paragraph()

    # Redes
    redes = payload.get("accesoriosRed", []) or payload.get("redes", [])
    if redes:
        doc.add_heading("Accesorios en redes", level=2)
        tbl = doc.add_table(rows=1, cols=4)
        hdr = tbl.rows[0].cells
        hdr[0].text = "Tipo"
        hdr[1].text = "Marca"
        hdr[2].text = "Serie"
        hdr[3].text = "Código / Mes-Año"
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
            # mostrar solo campos con valor
            pairs = [f"{k}: {v}" for k, v in eq.items() if v and str(v).strip()]
            doc.add_paragraph(", ".join(pairs))

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

    # Aplanar claves: cuando se reciben foto__1 foto__2 mapeamos tambien a base clave
    flattened = {}
    for key, flist in files_map.items():
        flattened.setdefault(key, []).extend(flist)
        if "__" in key:
            base = key.split("__")[0]
            flattened.setdefault(base, []).extend(flist)

    def insert_image_into_cell(cell, file_storage_obj):
        """
        Guarda temporalmente file_storage_obj, calcula tamaño proporcional para 15x10 cm
        e inserta la imagen en la celda.
        """
        try:
            tmp = save_temp_file(file_storage_obj, prefix="img")
            w, h = compute_picture_size_for_docx(tmp, max_cm_w=15.0, max_cm_h=10.0)
            run = cell.paragraphs[0].add_run()
            pic = run.add_picture(tmp)
            # ajustar en unidades
            pic.width = w
            pic.height = h
            try:
                os.remove(tmp)
            except:
                pass
            return True
        except Exception as e:
            cell.text = "(error insertando imagen)"
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

        # reunir archivos: exact match, claves con sufijos __1, __2, etc., y flattened base
        found_files = []
        if clave in files_map:
            found_files = files_map[clave]
        else:
            # buscar startswith en request.files keys
            for k in list(files_map.keys()):
                if k.startswith(clave):
                    found_files.extend(files_map[k])
            # fallback flattened
            if not found_files and clave in flattened:
                found_files = flattened[clave]

        if not found_files:
            doc.add_paragraph("No se subió fotografía para este subtítulo.")
            continue

        # col layout: si el bloque pide >1 fotos (ej. 4 panorámicas) intentamos ponerlas en columnas
        cols = min(fotos_needed, max(1, len(found_files)))
        rows = math.ceil(len(found_files) / cols)
        tbl = doc.add_table(rows=rows, cols=cols)
        f_idx = 0
        for r in range(rows):
            cells = tbl.rows[r].cells
            for c in range(cols):
                if f_idx < len(found_files):
                    insert_image_into_cell(cells[c], found_files[f_idx])
                else:
                    cells[c].text = ""
                f_idx += 1
        doc.add_paragraph()

    # DETALLE ACTIVIDADES (10.X) - ANTES / DESPUÉS
    doc.add_page_break()
    doc.add_heading("Trabajos realizados (detalle antes / después)", level=2)
    actividades_obj = payload.get("actividades") or payload.get("activities") or payload.get("actividades_por_tanque") or {}

    if actividades_obj and isinstance(actividades_obj, dict) and len(actividades_obj) > 0:
        # mostrar por key (puede ser 'tank1', '1', 'red_llenado_retorno', etc.)
        for key, acts in actividades_obj.items():
            # Normalize key string for title
            title_key = str(key)
            doc.add_heading(f"Actividades — {title_key}", level=3)
            if not acts or (isinstance(acts, list) and len(acts) == 0):
                doc.add_paragraph("No se registraron actividades.")
                continue
            if not isinstance(acts, list):
                doc.add_paragraph("Formato de actividades inválido (se esperaba lista).")
                continue
            for a_idx, act in enumerate(acts, start=1):
                desc = safe_str(act.get("descripcion") or act.get("desc") or f"Actividad {a_idx}")
                doc.add_paragraph(f"- {desc}", style='Intense Quote')
                # claves preferidas
                before_keys = [
                    f"foto_10_{key}_act{a_idx}__antes", f"foto_10_{key}_act{a_idx}__1",
                    f"foto_10_{key}_act{a_idx}__1_antes", f"foto_10_{key}_act{a_idx}__before"
                ]
                after_keys = [
                    f"foto_10_{key}_act{a_idx}__despues", f"foto_10_{key}_act{a_idx}__2",
                    f"foto_10_{key}_act{a_idx}__2_despues", f"foto_10_{key}_act{a_idx}__after"
                ]
                before_files = []
                after_files = []

                for k in before_keys:
                    if k in files_map:
                        before_files.extend(files_map[k])
                # also check startswith and flattened
                for k in list(files_map.keys()):
                    if any(k.startswith(bk) for bk in before_keys):
                        before_files.extend(files_map[k])
                if len(before_files) == 0:
                    for bk in before_keys:
                        if bk in flattened:
                            before_files.extend(flattened[bk])

                for k in after_keys:
                    if k in files_map:
                        after_files.extend(files_map[k])
                for k in list(files_map.keys()):
                    if any(k.startswith(ak) for ak in after_keys):
                        after_files.extend(files_map[k])
                if len(after_files) == 0:
                    for ak in after_keys:
                        if ak in flattened:
                            after_files.extend(flattened[ak])

                # tabla 2 columnas ANTES / DESPUÉS
                tbl = doc.add_table(rows=1, cols=2)
                hdr = tbl.rows[0].cells
                hdr[0].text = "ANTES"
                hdr[1].text = "DESPUÉS"
                row = tbl.add_row().cells

                if before_files:
                    insert_image_into_cell(row[0], before_files[0])
                else:
                    row[0].text = "(No hubo fotografía ANTES)"

                if after_files:
                    insert_image_into_cell(row[1], after_files[0])
                else:
                    row[1].text = "(No hubo fotografía DESPUÉS)"

                doc.add_paragraph()
    else:
        doc.add_paragraph("No se registraron actividades dinámicas en el payload.")

    # Guardar y enviar
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return send_file(
        bio,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="informe_generado.docx"
    )


if __name__ == "__main__":
    # debug True para desarrollo, cambia a False en producción
    app.run(host="0.0.0.0", port=5000, debug=True)
