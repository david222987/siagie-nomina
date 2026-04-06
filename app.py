import os, re, io, unicodedata
import pandas as pd
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename

try:
    import pdfplumber
    PDFPLUMBER_OK = True
except ImportError:
    PDFPLUMBER_OK = False

try:
    from pypdf import PdfReader
    PYPDF_OK = True
except ImportError:
    PYPDF_OK = False

try:
    import pytesseract
    from PIL import Image
    from pdf2image import convert_from_bytes
    OCR_OK = True
except ImportError:
    OCR_OK = False

# ─────────────────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  LIMPIEZA DE TEXTO
# ═══════════════════════════════════════════════════════════════════════════════

def limpiar(texto: str) -> str:
    if not texto:
        return ""
    texto = unicodedata.normalize("NFKC", texto)
    texto = re.sub(r'[ \t]+', ' ', texto)
    texto = re.sub(r'\n{3,}', '\n\n', texto)
    return texto.strip()

def tiene_texto(texto: str) -> bool:
    return len(re.findall(r'\b\w{3,}\b', texto or "")) >= 8

# ═══════════════════════════════════════════════════════════════════════════════
#  PATRONES ESPECÍFICOS DEL FORMATO SIAGIE
# ═══════════════════════════════════════════════════════════════════════════════

RE_DNI_SIAGIE = re.compile(
    r'D[\.\s]*N[\.\s]*I[\.\s·:]*[\s·]*'
    r'([\d\s·.\-]{8,30})',
    re.IGNORECASE
)
RE_DNI_COMPACTO = re.compile(r'\b(\d{8})\b')

RE_COD_MOD = re.compile(
    r'[Cc][oó]digo\s+[Mm]odular[:\s]*([0-9\s·.]{6,16})'
)
RE_COD_MOD_SOLO = re.compile(r'\b(0\d{6,7})\b')

RE_IE_NOMBRE = re.compile(
    r'[Nn][uú]mero\s+y/o\s+[Nn]ombre[\s:]*([^\n\r]{3,80})'
)
RE_IE_ALT = re.compile(
    r'(?:I\.?\s*E\.?|[Ii]nstituci[oó]n\s+[Ee]ducativa)[:\s]+([^\n\r]{5,80})'
)
IE_CORTE = re.compile(
    r'\s+(?:Gesti[oó]n|PGD|Privada|P[uú]blica|Inicio|Fin|Inicio|'
    r'Caracter[ií]stica|Programa|Forma|EBR|EBA|EBE|CEBA|CETPRO)',
    re.IGNORECASE
)

# ── Nombre: acepta tanto "APELLIDO APELLIDO, Nombre" como solo mayúsculas ──
RE_NOMBRE_SIAGIE = re.compile(
    r'([A-ZÁÉÍÓÚÑ]{2,}(?:\s+[A-ZÁÉÍÓÚÑ]{2,})+),\s*'
    r'([A-Za-záéíóúñÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-Za-záéíóúñÁÉÍÓÚÑ][a-záéíóúñ]+)*)'
)

STOP_WORDS = {
    'DNI', 'CODIGO', 'MODULAR', 'ALUMNO', 'ESTUDIANTE', 'NOMBRE', 'APELLIDO',
    'NUMERO', 'ORDEN', 'GRADO', 'SECCION', 'TURNO', 'MODALIDAD', 'NIVEL',
    'FECHA', 'NACIMIENTO', 'SEXO', 'ESTADO', 'MATRICULA', 'NOMINA', 'TOTAL',
    'INSTITUCION', 'EDUCATIVA', 'MINISTERIO', 'EDUCACION', 'SIAGIE', 'PERU',
    'INICIAL', 'PRIMARIA', 'SECUNDARIA', 'UGEL', 'PIURA', 'LIMA', 'DRE',
    'INICIO', 'FIN', 'GESTION', 'PROGRAMA', 'PERIODO', 'LECTIVO', 'FORMA',
    'RESOLUCION', 'CREACION', 'CARACTERISTICA', 'ALFABETICO', 'ORDEN',
}


# ═══════════════════════════════════════════════════════════════════════════════
#  NORMALIZAR DNI
# ═══════════════════════════════════════════════════════════════════════════════

def normalizar_digitos(raw: str) -> str:
    return re.sub(r'\D', '', raw)

def normalizar_dni(raw: str) -> str:
    digitos = normalizar_digitos(raw)
    if len(digitos) == 8:
        return digitos
    if len(digitos) == 9 and digitos.startswith('1'):
        return digitos[1:]
    if len(digitos) > 8:
        return digitos[-8:]
    return ""

def dni_valido(dni: str) -> bool:
    return len(dni) == 8 and len(set(dni)) > 2


# ═══════════════════════════════════════════════════════════════════════════════
#  EXTRACCIÓN DE TEXTO DEL PDF
# ═══════════════════════════════════════════════════════════════════════════════

def extraer_con_pdfplumber(pdf_bytes: bytes):
    if not PDFPLUMBER_OK:
        return "", []
    texto_total = []
    filas_tabla = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                t = page.extract_text(x_tolerance=3, y_tolerance=3)
                if t:
                    texto_total.append(t)
                tablas = page.extract_tables({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 5,
                    "join_tolerance": 3,
                })
                for tabla in (tablas or []):
                    for fila in tabla:
                        if fila:
                            celdas = [str(c or "").strip() for c in fila]
                            filas_tabla.append(celdas)
    except Exception as e:
        print(f"[pdfplumber] {e}")
    return "\n".join(texto_total), filas_tabla


def extraer_con_pypdf(pdf_bytes: bytes) -> str:
    if not PYPDF_OK:
        return ""
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        return "\n".join(p.extract_text() or "" for p in reader.pages)
    except Exception as e:
        print(f"[pypdf] {e}")
        return ""


def extraer_con_ocr(pdf_bytes: bytes) -> str:
    if not OCR_OK:
        return ""
    try:
        imagenes = convert_from_bytes(pdf_bytes, dpi=300)
        partes = []
        for img in imagenes:
            cfg = r'--oem 3 --psm 6 -l spa+eng'
            partes.append(pytesseract.image_to_string(img, config=cfg))
        return "\n".join(partes)
    except Exception as e:
        print(f"[OCR] {e}")
        return ""


def obtener_texto_y_tablas(pdf_bytes: bytes):
    texto, filas = extraer_con_pdfplumber(pdf_bytes)
    texto = limpiar(texto)
    if tiene_texto(texto):
        return texto, filas, "pdfplumber"

    texto2 = limpiar(extraer_con_pypdf(pdf_bytes))
    if tiene_texto(texto2):
        return texto2, [], "pypdf"

    if OCR_OK:
        texto3 = limpiar(extraer_con_ocr(pdf_bytes))
        return texto3, [], "ocr"

    return texto, filas, "sin_texto"


# ═══════════════════════════════════════════════════════════════════════════════
#  EXTRACCIÓN DE CABECERA
# ═══════════════════════════════════════════════════════════════════════════════

def extraer_cabecera(texto: str, filas: list) -> dict:
    ie = ""
    cod_modular = ""

    m = RE_IE_NOMBRE.search(texto)
    if m:
        ie = m.group(1).strip()
        corte = IE_CORTE.search(ie)
        if corte:
            ie = ie[:corte.start()].strip()
        ie = re.sub(r'\s+\d{4,}.*$', '', ie).strip()

    if not ie:
        m = RE_IE_ALT.search(texto)
        if m:
            ie = m.group(1).strip()
            corte = IE_CORTE.search(ie)
            if corte:
                ie = ie[:corte.start()].strip()

    m_cod = RE_COD_MOD.search(texto)
    if m_cod:
        digitos = normalizar_digitos(m_cod.group(1))
        if 6 <= len(digitos) <= 8:
            cod_modular = digitos

    if not cod_modular:
        m2 = RE_COD_MOD_SOLO.search(texto)
        if m2:
            cod_modular = m2.group(1)

    for fila in filas:
        texto_fila = " ".join(fila)
        if not ie and re.search(r'[Nn][uú]mero\s+y/o\s+[Nn]ombre', texto_fila):
            for celda in fila:
                if celda and not re.search(r'[Nn][uú]mero\s+y/o\s+[Nn]ombre', celda):
                    ie = celda.strip()
                    break

        if not cod_modular and re.search(r'[Cc][oó]digo\s+[Mm]odular', texto_fila):
            for celda in fila:
                d = normalizar_digitos(celda)
                if 6 <= len(d) <= 8:
                    cod_modular = d
                    break

    return {
        "ie": limpiar(ie)[:100] if ie else "—",
        "codigo_modular": cod_modular if cod_modular else "—"
    }


# ═══════════════════════════════════════════════════════════════════════════════
#  HELPERS PARA EXTRACCIÓN DE NOMBRE DESDE CELDA
# ═══════════════════════════════════════════════════════════════════════════════

def extraer_nombre_de_celda(celda: str) -> str:
    """
    Extrae nombre en formato 'APELLIDO APELLIDO, Nombre Nombre' de una celda.
    Ignora celdas que sean encabezados o palabras reservadas.
    """
    if not celda or len(celda) < 5:
        return ""
    m = RE_NOMBRE_SIAGIE.search(celda)
    if m:
        aps = m.group(1).strip()
        nms = m.group(2).strip()
        # Verificar que ningún apellido sea una stop word
        if not any(w in STOP_WORDS for w in aps.split()):
            return f"{aps}, {nms}"
    return ""


def es_celda_dni(celda: str) -> str:
    """
    Retorna el DNI (8 dígitos) si la celda contiene un DNI válido, sino "".
    Soporta formato SIAGIE con separadores y formato compacto.
    """
    # Formato SIAGIE: "D·N·I·· ·9·3·2·8·2·7·7·8"
    m = RE_DNI_SIAGIE.search(celda)
    if m:
        return normalizar_dni(m.group(1))
    # Celda que es SOLO dígitos separados por puntos/espacios (columna DNI pura)
    solo_dig = re.sub(r'[\s·.\-]', '', celda)
    if re.match(r'^\d{8}$', solo_dig) and len(set(solo_dig)) > 2:
        return solo_dig
    return ""


# ═══════════════════════════════════════════════════════════════════════════════
#  EXTRACCIÓN DE ALUMNOS  ← LÓGICA CORREGIDA
# ═══════════════════════════════════════════════════════════════════════════════

def extraer_alumnos(texto: str, filas: list, cabecera: dict) -> list:
    ie  = cabecera["ie"]
    cod = cabecera["codigo_modular"]
    alumnos = []
    dnis_vistos = set()

    # ── ESTRATEGIA 1: Filas de tabla ─────────────────────────────────────────
    # Cada fila de la tabla SIAGIE contiene: [N°, DNI_celda, Nombre_celda, ...]
    # El nombre y el DNI están en celdas DISTINTAS de la misma fila.
    # IMPORTANTE: buscamos DNI y nombre INDEPENDIENTEMENTE dentro de la fila,
    # sin mezclar datos de otras filas.
    for fila in filas:
        dni = ""
        nombre = ""

        for celda in fila:
            # 1a) Intentar extraer DNI de esta celda
            if not dni:
                dni_candidato = es_celda_dni(celda)
                if dni_candidato and dni_valido(dni_candidato):
                    dni = dni_candidato
                    continue  # Esa celda era el DNI, no buscar nombre en ella

            # 1b) Intentar extraer nombre de esta celda (si ya tenemos o no DNI)
            if not nombre:
                nombre_candidato = extraer_nombre_de_celda(celda)
                if nombre_candidato:
                    nombre = nombre_candidato

        # Solo registrar si el DNI es válido y no fue visto antes
        if dni and dni_valido(dni) and dni not in dnis_vistos:
            dnis_vistos.add(dni)
            alumnos.append({
                "dni": dni,
                "nombre": nombre if nombre else "—",
                "ie": ie,
                "codigo_modular": cod
            })

    # ── ESTRATEGIA 2: Texto libre línea por línea ─────────────────────────────
    # Solo se usa si la estrategia de tabla no encontró nada.
    # Empareja cada línea con su DNI y busca el nombre SOLO en esa misma línea.
    if not alumnos:
        lineas = texto.split('\n')
        for linea in lineas:
            # Buscar DNI en la línea
            m_dni = RE_DNI_SIAGIE.search(linea)
            if m_dni:
                dni = normalizar_dni(m_dni.group(1))
            else:
                m2 = RE_DNI_COMPACTO.search(linea)
                dni = m2.group(1) if m2 else ""

            if not dni or not dni_valido(dni) or dni in dnis_vistos:
                continue

            # Buscar nombre SOLO en la misma línea (no en ventana de líneas cercanas)
            nombre = extraer_nombre_de_celda(linea)

            # Si no encontró nombre en la línea del DNI, buscar en la línea INMEDIATA siguiente
            # (en algunos PDFs el nombre aparece en la línea siguiente al DNI)
            if not nombre:
                idx = lineas.index(linea)
                if idx + 1 < len(lineas):
                    nombre = extraer_nombre_de_celda(lineas[idx + 1])

            dnis_vistos.add(dni)
            alumnos.append({
                "dni": dni,
                "nombre": nombre if nombre else "—",
                "ie": ie,
                "codigo_modular": cod
            })

    # ── ESTRATEGIA 3: Regex sobre texto completo ──────────────────────────────
    # Último recurso. Busca DNIs y el nombre en el fragmento cercano INMEDIATO
    # (solo 100 chars antes/después para evitar capturar nombres de otras filas).
    if not alumnos:
        for m in RE_DNI_COMPACTO.finditer(texto):
            dni = m.group(1)
            if not dni_valido(dni) or dni in dnis_vistos:
                continue

            # Ventana reducida: solo 100 chars alrededor del DNI encontrado
            inicio = max(0, m.start() - 50)
            fin    = min(len(texto), m.end() + 150)
            fragmento = texto[inicio:fin]

            nombre = extraer_nombre_de_celda(fragmento)
            dnis_vistos.add(dni)
            alumnos.append({
                "dni": dni,
                "nombre": nombre if nombre else "—",
                "ie": ie,
                "codigo_modular": cod
            })

    return alumnos


# ═══════════════════════════════════════════════════════════════════════════════
#  RUTAS FLASK
# ═══════════════════════════════════════════════════════════════════════════════

def allowed_file(fn):
    return '.' in fn and fn.rsplit('.', 1)[1].lower() == 'pdf'


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/procesar', methods=['POST'])
def procesar():
    if 'archivos' not in request.files:
        return jsonify({"error": "No se recibieron archivos"}), 400

    archivos = request.files.getlist('archivos')
    resultados = []
    errores = []

    for archivo in archivos:
        if not allowed_file(archivo.filename):
            errores.append(f"{archivo.filename}: no es PDF")
            continue
        try:
            pdf_bytes = archivo.read()
            texto, filas, metodo = obtener_texto_y_tablas(pdf_bytes)

            if not texto and not filas:
                errores.append(f"{archivo.filename}: no se pudo extraer texto")
                continue

            cabecera = extraer_cabecera(texto, filas)
            alumnos  = extraer_alumnos(texto, filas, cabecera)

            if not alumnos:
                errores.append(f"{archivo.filename}: no se encontraron alumnos")
                continue

            nombre_arch = secure_filename(archivo.filename)
            for a in alumnos:
                a['archivo'] = nombre_arch
                a['metodo']  = metodo
                resultados.append(a)

        except Exception as e:
            errores.append(f"{archivo.filename}: error → {e}")

    # Deduplicar por DNI
    vistos, unicos = set(), []
    for r in resultados:
        if r['dni'] not in vistos:
            vistos.add(r['dni'])
            unicos.append(r)

    return jsonify({"total": len(unicos), "alumnos": unicos, "errores": errores})


@app.route('/exportar', methods=['POST'])
def exportar():
    datos = request.get_json()
    if not datos or not datos.get('alumnos'):
        return jsonify({"error": "Sin datos"}), 400

    df = pd.DataFrame(datos['alumnos']).rename(columns={
        'dni': 'DNI',
        'nombre': 'Apellidos y Nombres',
        'ie': 'Institución Educativa',
        'codigo_modular': 'Código Modular',
        'archivo': 'Archivo PDF',
        'metodo': 'Método'
    })

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Nómina')
        ws = writer.sheets['Nómina']
        for col in ws.columns:
            w = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(w + 4, 55)
    out.seek(0)

    return send_file(
        out,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='nomina_matricula_2026.xlsx'
    )


@app.route('/estado')
def estado():
    return jsonify({"pdfplumber": PDFPLUMBER_OK, "pypdf": PYPDF_OK, "ocr": OCR_OK})


if __name__ == '__main__':
    print("=" * 55)
    print("  SIAGIE Extractor — Nómina de Matrícula 2026")
    print("  http://localhost:5000")
    print(f"  pdfplumber: {'✓' if PDFPLUMBER_OK else '✗'} | pypdf: {'✓' if PYPDF_OK else '✗'} | OCR: {'✓' if OCR_OK else '✗'}")
    print("=" * 55)
    app.run(debug=True, port=5000)