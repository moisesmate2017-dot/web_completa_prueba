#!/usr/bin/env python3
# servidor.py (versión corregida)
# Acepta FormData (payload JSON + imágenes) y genera un único .docx con imágenes
# Requiere: Flask, flask-cors, python-docx, pandas

import os
import tempfile
import json
import traceback
from datetime import datetime
from flask import Flask, request, send_file, jsonify, render_template
from flask_cors import CORS
import pandas as pd
import unicodedata
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# -----------------------
# Configuración básica
# -----------------------
app = Flask(__name__, template_folder="templates")
CORS(app)

# -----------------------
# Utilidades
# -----------------------
def normalizar(texto):
    if texto is None:
        return ""
    return (
        unicodedata.normalize("NFKD", str(texto).lower())
        .encode("ascii", "ignore")
        .decode("ascii")
        .replace(" ", "_")
    )

def valOrDash(v):
    return v if (v is not None and str(v).strip() != "") else "-"

# -----------------------
# Estilos y helpers docx
# -----------------------
def set_cell_style(cell, text, font_size=10, bold=False, align_center=True):
    """
    Pone texto en una celda y aplica estilo.
    """
    # limpiar y asignar texto
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
    """
    Crea una tabla con estilo básico y aplica tamaño/estilos a celdas.
    """
    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    if indent:
        # left indent for table (approx)
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

def make_placeholder_cell(cell, ancho_cm=15, alto_cm=10, text="ESPACIO PARA IMAGEN"):
    """
    Configura una celda como recuadro con borde y texto centrado.
    Notar: docx no expone setting de ancho/alto de celda de forma portable,
    pero se setea la altura de la fila en twips y se coloca texto placeholder.
    """
    # tamaño
    alto_twips = int(alto_cm * 567)  # aprox
    tr = cell._tc.getparent()  # <w:tr> owner row
    # ajustar altura fila (si es posible)
    try:
        trPr = tr.get_or_add_trPr()
        trHeight = OxmlElement("w:trHeight")
        trHeight.set(qn("w:val"), str(alto_twips))
        trHeight.set(qn("w:hRule"), "exact")
        trPr.append(trHeight)
    except Exception:
        pass

    # texto centrado
    p = cell.paragraphs[0]
    p.clear()  # limpiar
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10)
    run.bold = True
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # borde (XML)
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

def insertar_recuadro_foto(doc, ancho_cm=15, alto_cm=10):
    """
    Inserta un recuadro placeholder (tabla 1x1) con tamaño aproximado.
    """
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # attempt to set column width: create cell and set placeholder formatting
    cell = table.cell(0, 0)
    make_placeholder_cell(cell, ancho_cm, alto_cm)
    return table

def insertar_imagen_en_celda(cell, fileobj, ancho_cm=15, alto_cm=10):
    """
    Inserta la imagen en la celda, escalada para caber en el ancho máximo del recuadro.
    Se utiliza run.add_picture para que la imagen quede dentro del párrafo de la celda.
    """
    try:
        # calcular ancho en pulgadas
        ancho_in = ancho_cm / 2.54
        # limpiar contenido previo
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run()
        # run.add_picture acepta file-like y width en Inches
        run.add_picture(fileobj, width=Inches(ancho_in))
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        return True
    except Exception as e:
        print("Error insertar_imagen_en_celda:", e)
        return False

def insertar_imagen_si_existe(doc, clave, imagenes, ancho_cm=15, alto_cm=10, num_recuadros=1):
    """
    Inserta N recuadros horizontales (tabla 1xN). Si la imagen existe para la clave
    (clave o clave__1, clave__2 ...), inserta la(s) imágenes escaladas dentro de cada celda.
    Si no existen, inserta placeholders.
    """
    # construir tabla con N columnas
    cols = num_recuadros
    table = doc.add_table(rows=1, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # establecer celdas
    for i in range(cols):
        cell = table.cell(0, i)
        # clave posible
        if num_recuadros == 1:
            k = clave
        else:
            k = f"{clave}__{i+1}"
        if imagenes and k in imagenes:
            success = insertar_imagen_en_celda(cell, imagenes[k], ancho_cm, alto_cm)
            if not success:
                make_placeholder_cell(cell, ancho_cm, alto_cm)
        else:
            # si no existe imagen, placeholder
            make_placeholder_cell(cell, ancho_cm, alto_cm)
    # espacio después
    doc.add_paragraph("")
    return table

# -----------------------
# Construir lista de bloques / subtítulos (API)
# -----------------------
def construir_bloques_fotos(df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs):
    """
    Construye una lista ordenada de bloques fotográficos (diccionarios):
    { 'titulo', 'clave', 'fotos', 'aplica', 'order' }
    - Tiene en cuenta TODOS los subtítulos posibles según datos de entrada.
    - Para accesorios en red con múltiples entradas, genera uno por cada item encontrado.
    """
    bloques = []
    contador = 1

    # 8. evidencia general
    bloques.append({'titulo': "8. EVIDENCIA FOTOGRÁFICA (del establecimiento)", 'clave':'foto_8_1', 'fotos':1, 'aplica':True, 'order':contador})
    contador += 1

    # 9.1 panorámica general
    bloques.append({'titulo': f"9.{contador-1}. FOTO PANORÁMICA DE LA ZONA", 'clave':'foto_9_panoramica', 'fotos':1, 'aplica':True, 'order':contador})
    contador += 1

    # tanques: placa + panorámicas alrededor + items por tanque
    tanques_list = df_tanques.to_dict(orient='records') if (df_tanques is not None and not df_tanques.empty) else []
    for i, t in enumerate(tanques_list):
        serie = t.get('N° de serie') or t.get('serie') or '-'
        # placa
        bloques.append({'titulo': f"9.{contador-1}. PLACA DE TANQUE {i+1} DE SERIE: {serie}", 'clave':f'foto_9_placa_tank_{i+1}', 'fotos':1, 'aplica':True, 'order':contador})
        contador += 1
        # panorámica de alrededores (4 recuadros)
        bloques.append({'titulo': f"9.{contador-1}. FOTO PANORÁMICA DE ALREDEDORES DE TANQUE {i+1} DE SERIE: {serie}", 'clave':f'foto_9_panoramica_tank_{i+1}', 'fotos':4, 'aplica':True, 'order':contador})
        contador += 1
        # items diversos por tanque (lista)
        titles = [
            "FOTO DE BASES DE CONCRETO",
            "FOTO DE MANÓMETROS 0-60 PSI",
            "FOTO DE MANÓMETROS 0-300 PSI",
            "FOTO DE CONEXIÓN DE CHICOTE A LA MULTIVÁLVULA",
            "STICKERS Y PINTADO DEL TANQUE",
            "FOTO DE LOS 04 ANCLAJES, PERNOS, TORNILLOS",
            "FOTO DE VÁLVULA DE LLENADO",
            "FOTO DE VÁLVULA DE SEGURIDAD",
            "FOTO DE VÁLVULA DE DRENAJE",
            "FOTO DE VÁLVULA DE MULTIVÁLVULA",
            "FOTO DE VÁLVULA DE MEDIDOR DE PORCENTAJE",
        ]
        for idx_title, txt in enumerate(titles):
            clave = f"foto_9_tank{ i+1 }_{ idx_title+1 }"
            bloques.append({'titulo': f"9.{contador-1}. {txt} DE TANQUE {i+1} DE SERIE: {serie}", 'clave':clave, 'fotos':1, 'aplica':True, 'order':contador})
            contador += 1

    # Equipos específicos: por tipo y por cada equipo existente (o placeholder)
    tipos = ["estabilizador", "quemador", "vaporizador", "tablero", "bomba", "dispensador_de_gas", "decantador", "detector", "extintor"]
    equipos_for_block = df_equipos.to_dict(orient='records') if (df_equipos is not None and not df_equipos.empty) else []
    for tipo in tipos:
        lista_eq = [eq for eq in equipos_for_block if (eq.get('Tipo de equipo') or eq.get('tipo') or '').lower() == tipo]
        if lista_eq:
            for idx_eq, eq in enumerate(lista_eq):
                serie = eq.get('Serie') or eq.get('serie') or '-'
                bloques.append({'titulo': f"9.{contador-1}. FOTO DE PLACA DE {tipo.upper()} DE SERIE: {serie}", 'clave':f"foto_9_{tipo}_placa_{idx_eq+1}", 'fotos':1, 'aplica':True, 'order':contador})
                contador += 1
                bloques.append({'titulo': f"9.{contador-1}. FOTO DE {tipo.upper()} (GENERAL)", 'clave':f"foto_9_{tipo}_general_{idx_eq+1}", 'fotos':1, 'aplica':True, 'order':contador})
                contador += 1
        else:
            # añadir placeholders no aplican
            bloques.append({'titulo': f"9.{contador-1}. FOTO DE PLACA DE {tipo.upper()} DE SERIE: -", 'clave':f"foto_9_{tipo}_placa_1", 'fotos':1, 'aplica':False, 'order':contador})
            contador += 1
            bloques.append({'titulo': f"9.{contador-1}. FOTO DE {tipo.upper()}", 'clave':f"foto_9_{tipo}_general_1", 'fotos':1, 'aplica':False, 'order':contador})
            contador += 1

    # Accesorios en red: generar por cada elemento en df_red
    accesorios_map = {
        "llenado_toma_desplazada": "VÁLVULA DE LLENADO (TOMA DESPLAZADA)",
        "retorno_toma_desplazada": "VÁLVULA DE RETORNO (TOMA DESPLAZADA)",
        "alivio_hidrostatico": "VÁLVULA DE ALIVIO HIDROSTÁTICO",
        "regulador_primera_etapa": "REGULADOR DE PRIMERA ETAPA",
        "alivio": "VÁLVULA DE ALIVIO",
        "regulador_2da": "REGULADOR DE SEGUNDA ETAPA",
        "pull_away": "VÁLVULA PULL AWAY",
    }
    # if df_red has rows, create entries for each record and type
    try:
        if df_red is not None and not df_red.empty:
            for idx, row in df_red.reset_index(drop=True).iterrows():
                tipo = (row.get("Tipo") or row.get("tipo") or "").strip()
                keynorm = normalizar(tipo)
                titulo = accesorios_map.get(keynorm, f"ACCESORIO: {tipo}")
                # cada accesorio tendrá su propia clave indexada
                clave = f"foto_9_acc_{keynorm}_{idx+1}"
                bloques.append({'titulo': f"9.{contador-1}. FOTO DE {titulo} (CÓDIGO: {valOrDash(row.get('Código'))})", 'clave':clave, 'fotos':1, 'aplica':True, 'order':contador})
                contador += 1
    except Exception:
        # fallback: si df_red vacio -> nada
        pass

    # Zona de medidores
    tiene_zona = False
    try:
        if df_red is not None and 'Tipo' in df_red.columns:
            tiene_zona = df_red['Tipo'].astype(str).str.lower().str.contains('zona_medidores').any()
    except Exception:
        tiene_zona = False
    bloques.append({'titulo': f"9.{contador-1}. FOTO DE ZONA MEDIDORES", 'clave':'foto_9_zona_medidores', 'fotos':1, 'aplica': bool(tiene_zona), 'order':contador})
    contador += 1

    # Punto de transferencia desplazado (si existe)
    tiene_toma = False
    try:
        if df_red is not None and 'Tipo' in df_red.columns:
            tiene_toma = df_red['Tipo'].astype(str).str.lower().str.contains('llenado_toma_desplazada').any()
    except Exception:
        tiene_toma = False
    bloques.append({'titulo': f"9.{contador-1}. FOTO DEL PUNTO DE TRANSFERENCIA DESPLAZADO", 'clave':'foto_9_toma_transferencia', 'fotos':1, 'aplica': bool(tiene_toma), 'order':contador})
    contador += 1

    return bloques

# -----------------------
# Generador DOCX central
# -----------------------
def generar_docx_desde_dfs(df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs, imagenes):
    """
    Genera el documento a partir de los dataframes y el dict 'imagenes' (con keys -> file objects).
    Devuelve la ruta del archivo temporal .docx creado.
    """
    if imagenes is None:
        imagenes = {}

    doc = Document()

    # --- Header / título
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
    try:
        tabla2.cell(0, 0).merge(tabla2.cell(0, 1)).text = "DOMESTICO"
        tabla2.cell(0, 2).merge(tabla2.cell(0, 5)).text = "INDUSTRIAL"
        tabla2.cell(0, 6).merge(tabla2.cell(0, 7)).text = "CANALIZADO"
    except Exception:
        pass
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
    tabla3 = create_table(doc, max(2, num_tanques + 1), len(headers3))
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
    unique_tanques = sorted(df_accesorios["Tanque"].unique()) if not df_accesorios.empty else []
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
        # merge number column (para agrupar el tanque)
        try:
            tabla4.cell(row_idx - len(atributos), 0).merge(tabla4.cell(row_idx - 1, 0))
            set_cell_style(tabla4.cell(row_idx - len(atributos), 0), str(tanque), font_size=7)
        except Exception:
            pass

    # === 5. Accesorios en redes ===
    add_subtitle(doc, "5. ACCESORIOS EN REDES")
    df_red_local = df_red.copy() if df_red is not None else pd.DataFrame(columns=["Tipo","Marca","Serie","Código","Mes/Año de fabricación"])
    if "Tipo" in df_red_local.columns:
        df_red_local["Tipo"] = df_red_local["Tipo"].astype(str).str.lower().fillna("")
    else:
        df_red_local["Tipo"] = ""
    # mapa para subtítulos estandarizados
    mapa_accesorios = {
        "llenado_toma_desplazada": "5.1. Válvula de llenado (toma desplazada)",
        "retorno_toma_desplazada": "5.2. Válvula de retorno (toma desplazada)",
        "alivio_hidrostatico": "5.3. Válvula de alivio hidrostático",
        "regulador_primera_etapa": "5.4. Regulador de primera etapa",
        "alivio": "5.5. Válvula de alivio",
        "regulador_2da": "5.6. Regulador de segunda etapa",
        "pull_away": "5.7. Válvula Pull Away",
    }
    accesorios_red_dict = {}
    # generar tablas por cada tipo si existen
    for clave, titulo in mapa_accesorios.items():
        add_subtitle(doc, titulo, indent=True)
        try:
            lista = df_red_local[df_red_local["Tipo"] == clave].to_dict(orient="records") if not df_red_local.empty else []
        except Exception:
            lista = []
        filas = max(2, len(lista) + 1)
        tabla = create_table(doc, filas, 5, indent=True)
        headers = ["Válvula", "Marca", "Serie", "Código", "Mes/Año de fabricación"]
        for j, h in enumerate(headers):
            set_cell_style(tabla.cell(0, j), h, bold=True)
        if lista:
            for idx, acc in enumerate(lista):
                set_cell_style(tabla.cell(idx + 1, 0), str(idx + 1))
                set_cell_style(tabla.cell(idx + 1, 1), valOrDash(acc.get("Marca")))
                set_cell_style(tabla.cell(idx + 1, 2), valOrDash(acc.get("Serie")))
                set_cell_style(tabla.cell(idx + 1, 3), valOrDash(acc.get("Código")))
                set_cell_style(tabla.cell(idx + 1, 4), valOrDash(acc.get("Mes/Año de fabricación")))
        else:
            for j in range(5):
                set_cell_style(tabla.cell(1, j), "-")
        accesorios_red_dict[clave] = lista

    # zona medidores (booleano si existe en df_red)
    zona_medidores_bool = False
    try:
        zona_medidores_bool = df_red_local[df_red_local["Tipo"] == "zona_medidores"]["Código"].astype(str).str.lower().str.contains("true").any()
    except Exception:
        zona_medidores_bool = False
    accesorios_red_dict["zona_medidores"] = zona_medidores_bool

    # === 6. Equipos de la instalación ===
    add_subtitle(doc, "6. EQUIPOS DE LA INSTALACIÓN")
    df_equipos_local = df_equipos.copy() if df_equipos is not None else pd.DataFrame()
    if "Tipo de equipo" in df_equipos_local.columns:
        df_equipos_local["Tipo de equipo"] = df_equipos_local["Tipo de equipo"].astype(str).str.lower().fillna("")
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
        datos = grupos_e.get_group(tipo_equipo) if (hasattr(grupos_e, "groups") and tipo_equipo in grupos_e.groups) else pd.DataFrame(columns=columnas)
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
    equipos_instalacion = {k: grupos_e.get_group(k).to_dict(orient="records") if (hasattr(grupos_e, "groups") and k in grupos_e.groups) else [] for k in estructura_equipos.keys()}

    # === 7. Observaciones generales ===
    add_subtitle(doc, "7. OBSERVACIONES GENERALES")
    if df_obs is None:
        df_obs_local = pd.DataFrame(columns=["Subpunto", "Observación"])
    elif isinstance(df_obs, dict):
        rows = []
        mapping = {"obs_71": "7.1", "obs_72": "7.2", "obs_73": "7.3", "obs_74": "7.4", "obs_75": "7.5"}
        for k, v in df_obs.items():
            if k in mapping:
                rows.append({"Subpunto": mapping[k], "Observación": v})
        df_obs_local = pd.DataFrame(rows)
    else:
        df_obs_local = df_obs.copy()

    subtitulos_7 = {
        "7.1": "7.1. Observaciones al cliente",
        "7.2": "7.2. Observaciones en red de llenado y retorno",
        "7.3": "7.3. Observaciones en zona de tanque",
        "7.4": "7.4. Observaciones en red de consumo",
    }
    for clave, titulo in subtitulos_7.items():
        add_subtitle(doc, titulo, indent=True)
        try:
            texto = df_obs_local[df_obs_local["Subpunto"] == clave]["Observación"].values
            if len(texto) and str(texto[0]).strip() != "":
                doc.add_paragraph(str(texto[0]).strip())
            else:
                doc.add_paragraph("-")
        except Exception:
            doc.add_paragraph("-")

    # 7.5 equipos varios
    add_subtitle(doc, "7.5. Observaciones en equipos varios (Vaporizador, Quemador, Decantador, etc)", indent=True)
    equipos_obs = ["Vaporizador","Quemador","Decantador","Dispensador de gas","Bomba de abastecimiento","Tablero eléctrico","Estabilizador","Detector de gases","Extintor"]
    tabla_obs = create_table(doc, len(equipos_obs) + 1, 2, indent=True)
    set_cell_style(tabla_obs.cell(0, 0), "Equipo", bold=True)
    set_cell_style(tabla_obs.cell(0, 1), "Observación", bold=True)
    texto_75 = []
    try:
        texto_75 = df_obs_local[df_obs_local["Subpunto"] == "7.5"]["Observación"].values
    except Exception:
        pass
    observaciones_75 = []
    if len(texto_75) and str(texto_75[0]).strip():
        observaciones_75 = [x.strip() for x in str(texto_75[0]).split(".") if x.strip()]
    for i, equipo in enumerate(equipos_obs):
        set_cell_style(tabla_obs.cell(i + 1, 0), equipo)
        set_cell_style(tabla_obs.cell(i + 1, 1), observaciones_75[i] if i < len(observaciones_75) else "-")

    # === 8. Evidencia general (establecimiento) ===
    add_subtitle(doc, "8. EVIDENCIA FOTOGRÁFICA (del establecimiento)")
    insertar_imagen_si_existe(doc, "foto_8_1", imagenes, ancho_cm=15, alto_cm=10, num_recuadros=1)

    # === 9. Evidencia fotográfica de elementos de la instalación ===
    add_subtitle(doc, "9. Evidencia fotográfica de elementos de la instalación")
    # Reconstruir los mismos bloques que construir_bloques_fotos genera
    bloques9 = []
    # Panoramica general
    bloques9.append(("9.1 FOTO PANORÁMICA DE LA ZONA", True, 1, "foto_9_panoramica"))
    # placas y tanques
    for i, t in enumerate(df_tanques.to_dict(orient='records') if (df_tanques is not None and not df_tanques.empty) else []):
        serie = valOrDash(t.get("N° de serie") or t.get("serie"))
        bloques9.append((f"PLACA DE TANQUE {i+1} DE SERIE: {serie}", True, 1, f"foto_9_placa_tank_{i+1}"))
        bloques9.append((f"PANORÁMICA DE ALREDEDORES DE TANQUE {i+1} DE SERIE: {serie}", True, 4, f"foto_9_panoramica_tank_{i+1}"))
        titles = [
            "FOTO DE BASES DE CONCRETO",
            "FOTO DE MANÓMETROS 0-60 PSI",
            "FOTO DE MANÓMETROS 0-300 PSI",
            "FOTO DE CONEXIÓN DE CHICOTE A LA MULTIVÁLVULA",
            "STICKERS Y PINTADO DEL TANQUE",
            "FOTO DE LOS 04 ANCLAJES, PERNOS, TORNILLOS",
            "FOTO DE VÁLVULA DE LLENADO",
            "FOTO DE VÁLVULA DE SEGURIDAD",
            "FOTO DE VÁLVULA DE DRENAJE",
            "FOTO DE VÁLVULA DE MULTIVÁLVULA",
            "FOTO DE VÁLVULA DE MEDIDOR DE PORCENTAJE",
        ]
        for idx_title, ttxt in enumerate(titles):
            bloques9.append((ttxt + f" (Tanque {i+1})", True, 1, f"foto_9_tank{i+1}_{idx_title+1}"))

    # equipos
    for tipo in ["estabilizador", "quemador", "vaporizador", "tablero", "bomba", "dispensador_de_gas", "decantador", "detector", "extintor"]:
        lista_eq = equipos_instalacion.get(tipo, [])
        if lista_eq:
            for idx_eq, eq in enumerate(lista_eq):
                serie = valOrDash(eq.get("Serie"))
                bloques9.append((f"FOTO DE PLACA DE {tipo.upper()} DE SERIE: {serie}", True, 1, f"foto_9_{tipo}_placa_{idx_eq+1}"))
                bloques9.append((f"FOTO DE {tipo.upper()}", True, 1, f"foto_9_{tipo}_general_{idx_eq+1}"))
        else:
            bloques9.append((f"FOTO DE PLACA DE {tipo.upper()} (NO DISPONIBLE)", False, 1, f"foto_9_{tipo}_placa_1"))
            bloques9.append((f"FOTO DE {tipo.upper()} (NO DISPONIBLE)", False, 1, f"foto_9_{tipo}_general_1"))

    # accesorios df_red: crear para cada fila
    if df_red is not None and not df_red.empty:
        for idx, r in df_red.reset_index(drop=True).iterrows():
            tipo = (r.get("Tipo") or r.get("tipo") or "")
            clave_safe = normalizar(tipo)
            titulo = f"FOTO DE {tipo} (CÓDIGO: {valOrDash(r.get('Código'))})"
            bloques9.append((titulo, True, 1, f"foto_9_acc_{clave_safe}_{idx+1}"))

    # zona medidores y toma desplazada
    bloques9.append(("FOTO DE ZONA MEDIDORES", bool(accesorios_red_dict.get("zona_medidores")), 1, "foto_9_zona_medidores"))
    bloques9.append(("FOTO DEL PUNTO DE TRANSFERENCIA DESPLAZADO", bool(accesorios_red_dict.get("llenado_toma_desplazada")), 1, "foto_9_toma_transferencia"))
    bloques9.append(("FOTO DE LA CAJA DE LA TOMA DESPLAZADA", bool(accesorios_red_dict.get("llenado_toma_desplazada")), 1, "foto_9_toma_caja"))
    bloques9.append(("FOTO DEL RECORRIDO DESDE TOMA DESPLAZADA HASTA TANQUE", bool(accesorios_red_dict.get("llenado_toma_desplazada")), 1, "foto_9_toma_recorrido"))

    # función interna para insertar con subtítulo
    def add_foto_con_subtitulo(doc, texto, incluir_imagen=True, num_recuadros=1, clave_base=None):
        add_subtitle(doc, texto, indent=True)
        if incluir_imagen:
            if clave_base is None:
                # insertar placeholders si no hay clave
                insertar_recuadro_foto(doc)
            else:
                insertar_imagen_si_existe(doc, clave_base, imagenes, ancho_cm=15, alto_cm=10, num_recuadros=num_recuadros)
        else:
            add_note(doc, "*NO CUENTA CON DICHO ELEMENTO")

    # recorrer bloques9 y generar
    i = 0
    while i < len(bloques9):
        texto, incluir, num_recuadros, clave = bloques9[i]
        # intentar agrupar dos subtítulos con 1 imagen cada uno para ahorrar espacio
        if incluir and num_recuadros == 1 and i + 1 < len(bloques9):
            texto2, incluir2, numr2, clave2 = bloques9[i+1]
            if incluir2 and numr2 == 1:
                # agrupar en una tabla 2 columnas: creamos subtítulo para ambos y luego tabla con 2 celdas de imagen
                add_subtitle(doc, texto + "  |  " + texto2, indent=True)
                # crear tabla 1x2 y llenar con imagenes o placeholders
                t = doc.add_table(rows=1, cols=2)
                t.alignment = WD_TABLE_ALIGNMENT.CENTER
                # celda 1
                cell1 = t.cell(0,0)
                if clave in imagenes:
                    insertar_imagen_en_celda(cell1, imagenes[clave], ancho_cm=15, alto_cm=10)
                else:
                    make_placeholder_cell(cell1, 15, 10)
                # celda 2
                cell2 = t.cell(0,1)
                if clave2 in imagenes:
                    insertar_imagen_en_celda(cell2, imagenes[clave2], ancho_cm=15, alto_cm=10)
                else:
                    make_placeholder_cell(cell2, 15, 10)
                doc.add_paragraph("")
                i += 2
                continue
        # si no agrupado, insertar normal
        add_foto_con_subtitulo(doc, texto, incluir_imagen=incluir, num_recuadros=num_recuadros, clave_base=clave)
        i += 1

    # === 10. Evidencia fotográfica (mantenimiento realizado) ===
    add_subtitle(doc, "10. EVIDENCIA FOTOGRÁFICA (MANTENIMIENTO REALIZADO)")
    add_note(doc, "NOTA 1: SE DEBERÁ MENCIONAR LOS TRABAJOS EJECUTADOS POR TANQUE (INCLUIR LAS INSPECCIONES QUE SE REALICEN)")
    add_note(doc, "NOTA 2: LAS IMÁGENES DEBEN TENER UN TAMAÑO DE 15CM X 10CM MÁXIMO")

    series_tanques = [valOrDash(row.get("N° de serie") or row.get("serie")) for row in df_tanques.to_dict(orient='records')] if (df_tanques is not None and not df_tanques.empty) else []
    for idx, serie in enumerate(series_tanques, start=1):
        clave_base = f"foto_10_tank{idx}"
        add_subtitle(doc, f"10.{idx}. TRABAJOS REALIZADOS EN EL TANQUE {idx} DE SERIE: {serie}", indent=True)
        # actividad 1 (antes y después)
        add_foto_con_subtitulo(doc, "(ACTIVIDAD 1: FOTO ANTES Y FOTO DESPUÉS)", incluir_imagen=True, num_recuadros=2, clave_base=clave_base)
        # actividad 2 (antes y después)
        add_foto_con_subtitulo(doc, "(ACTIVIDAD 2: FOTO ANTES Y FOTO DESPUÉS)", incluir_imagen=True, num_recuadros=2, clave_base=clave_base)

    # trabajos en redes
    add_subtitle(doc, f"10.{len(series_tanques)+1}. TRABAJOS REALIZADOS EN REDES DE LLENADO Y RETORNO", indent=True)
    add_foto_con_subtitulo(doc, "(ACTIVIDAD 1: FOTO ANTES Y FOTO DESPUÉS)", incluir_imagen=True, num_recuadros=2, clave_base="foto_10_red_llenado_retorno")
    add_foto_con_subtitulo(doc, "(ACTIVIDAD 2: FOTO ANTES Y FOTO DESPUÉS)", incluir_imagen=True, num_recuadros=2, clave_base="foto_10_red_llenado_retorno")

    add_subtitle(doc, f"10.{len(series_tanques)+2}. TRABAJOS REALIZADOS EN REDES DE CONSUMO", indent=True)
    add_foto_con_subtitulo(doc, "(ACTIVIDAD 1: FOTO ANTES Y FOTO DESPUÉS)", incluir_imagen=True, num_recuadros=2, clave_base="foto_10_red_consumo")
    add_foto_con_subtitulo(doc, "(ACTIVIDAD 2: FOTO ANTES Y FOTO DESPUÉS)", incluir_imagen=True, num_recuadros=2, clave_base="foto_10_red_consumo")

    # === 11,12,13 ===
    add_subtitle(doc, "11. EVIDENCIA FOTOGRÁFICA DE LA INSTALACIÓN")
    insertar_imagen_si_existe(doc, "foto_11_1", imagenes, ancho_cm=15, alto_cm=10, num_recuadros=1)
    add_subtitle(doc, "12. Conclusiones")
    doc.add_paragraph("-")
    add_subtitle(doc, "13. Recomendaciones")
    doc.add_paragraph("-")

    # Guardar en archivo temporal
    fd, path = tempfile.mkstemp(prefix="Informe_Mantenimiento_", suffix=".docx")
    os.close(fd)
    doc.save(path)
    return path

# -----------------------
# Construir DataFrames desde JSON
# -----------------------
def build_dfs_from_json(payload):
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

    # df_accesorios (convertir dict de accesorios por tanque a filas)
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

    # df_red: lista de dicts
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

    # df_equipos
    df_equipos = pd.DataFrame(equipos)

    # df_obs
    obs_rows = []
    for sp in ["7.1","7.2","7.3","7.4","7.5"]:
        obs_rows.append({"Subpunto": sp, "Observación": observaciones.get(sp, "")})
    df_obs = pd.DataFrame(obs_rows)

    return df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs

# -----------------------
# Rutas API
# -----------------------
@app.route("/generar", methods=["POST"])
def generar_informe():
    """
    Endpoint principal: recibe multipart/form-data (payload JSON + files) o JSON puro.
    Devuelve UN SOLO .docx (el que contiene las imágenes).
    """
    try:
        content_type = request.content_type or ""
        imagenes = {}
        if content_type.startswith("multipart/form-data"):
            payload_raw = request.form.get("payload")
            try:
                payload = json.loads(payload_raw) if payload_raw else {}
            except Exception as e:
                return jsonify({"error": "Payload JSON inválido en form-data", "detail": str(e)}), 400
            # collect files
            for key in request.files:
                imagenes[key] = request.files.get(key)
        else:
            payload = request.get_json(silent=True)
            imagenes = {}

        if not payload:
            return jsonify({"error": "No JSON recibido o body vacío"}), 400

        # validaciones mínimas
        general = payload.get("general", {}) or {}
        tanques = payload.get("tanques", []) or []

        required_general = [
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
        missing = [k for k in required_general if not (general.get(k) or "").strip()]
        if missing:
            return jsonify({"error": "Faltan campos obligatorios en 'general'", "missing": missing}), 400

        if len(tanques) == 0:
            return jsonify({"error": "Se requiere al menos un tanque en 'tanques'"}), 400

        # validaciones accesorios de tanque
        accesorios_tanque = payload.get("accesoriosTanque", {}) or {}
        for tk, accs in accesorios_tanque.items():
            for acc_name, fields in (accs or {}).items():
                if any((fields.get(f) or "").strip() for f in ["Marca", "Código", "Serie", "Mes/Año de fabricación"]):
                    missingf = [f for f in ["Marca","Código","Serie","Mes/Año de fabricación"] if not (fields.get(f) or "").strip()]
                    if missingf:
                        return jsonify({"error": f"En accesoriosTanque.{tk}.{acc_name} faltan campos: {missingf}"}), 400

        # validaciones accesorios en red
        accesorios_red = payload.get("accesoriosRed", []) or []
        for i, r in enumerate(accesorios_red):
            if any((r.get(k) or "").strip() for k in ["Marca", "Serie", "Código", "Mes/Año de fabricación"]):
                if not (r.get("Tipo") or "").strip():
                    return jsonify({"error": f"AccesoriosRed[{i}] tiene campos pero falta 'Tipo'"}), 400

        # validaciones equipos
        equipos = payload.get("equipos", []) or []
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
        for i, eq in enumerate(equipos):
            tipo = (eq.get("Tipo de equipo") or eq.get("tipo") or "").strip().lower()
            if tipo and tipo in estructura_equipos:
                required_cols = estructura_equipos[tipo]
                missing_eq = [c for c in required_cols if not (eq.get(c) or "").strip()]
                if missing_eq:
                    return jsonify({"error": f"Equipo[{i}] de tipo '{tipo}' faltan campos: {missing_eq}"}), 400

        # construir dfs
        df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs = build_dfs_from_json(payload)

        # generar docx (devuelve ruta)
        ruta = generar_docx_desde_dfs(df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs, imagenes)

        # enviar archivo al cliente (única descarga)
        try:
            # Flask >= 2.0 uses download_name
            response = send_file(ruta, as_attachment=True, download_name=os.path.basename(ruta))
        except TypeError:
            response = send_file(ruta, as_attachment=True)

        # intentar borrar archivo temporal (no bloquear envío)
        try:
            os.remove(ruta)
        except Exception:
            pass

        return response

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": "Error interno del servidor", "detail": str(e)}), 500

# --- NOTA: He dejado la ruta /generate_docx comentada/DEPRECATED porque en muchos frontends
# se usaba junto con /generar y eso provocaba descargas dobles. Si tu frontend llama a /generate_docx,
# cambia a /generar. Si necesitas que reactive /generate_docx, dime y la re-habilito.
#
# @app.route('/generate_docx', methods=['POST'])
# def api_generate_docx():
#     return jsonify({"error":"use /generar instead"}), 400

# API: devolver bloques fotográficos (POST JSON) - para frontend que pide bloques
@app.route('/bloques_fotos', methods=['POST'])
def api_bloques_fotos():
    try:
        data = request.get_json(force=True)
        df_tanques = pd.DataFrame(data.get('tanks', []))
        df_accesorios = pd.DataFrame()  # frontend doesn't send detailed table by default
        df_red = pd.DataFrame(data.get('redes', []))
        df_equipos = pd.DataFrame(data.get('equipos', []))
        df_info = pd.DataFrame([data.get('generalInfo', {})]) if data.get('generalInfo') else pd.DataFrame()
        df_obs = pd.DataFrame([data.get('obs', {})]) if data.get('obs') else pd.DataFrame()
        bloques = construir_bloques_fotos(df_info, df_tanques, df_accesorios, df_red, df_equipos, df_obs)
        return jsonify(bloques)
    except Exception as e:
        print('Error /bloques_fotos:', e)
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

# Ruta UI
@app.route('/')
def index():
    try:
        return render_template('pagina.html')
    except Exception as e:
        return f"<h3>Error al cargar la página: {str(e)}</h3>"

# Main
if __name__ == '__main__':
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
