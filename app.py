import csv
import io
import os
import re
import zipfile
from datetime import datetime

import docx
import pandas as pd
import pypdf
import streamlit as st

try:
	from docxtpl import DocxTemplate
	DOCX_DISPONIBLE = True
except ImportError:
	DOCX_DISPONIBLE = False

st.set_page_config(
	page_title="LexFlow Studio",
	page_icon="⚖️",
	layout="centered",
)

DB_FILE = "usuarios.csv"
DB_REGISTROS = "registros.csv"
REGISTRO_COLUMNAS = [
	"Fecha", "Año", "Usuario", "Pais", "Tema", "Cliente", "Identidad", "Juzgado"
]
LISTA_PAISES = [
	"Honduras", "España", "Argentina", "Bolivia", "Brasil", "Canadá",
	"Chile", "Colombia", "Costa Rica", "Cuba", "Ecuador", "El Salvador",
	"Estados Unidos", "Guatemala", "México", "Nicaragua", "Panamá",
	"Paraguay", "Perú", "República Dominicana", "Uruguay", "Venezuela",
	"Antigua y Barbuda", "Bahamas", "Barbados", "Belice", "Dominica",
	"Granada", "Guyana", "Haití", "Jamaica", "San Cristóbal y Nieves",
	"San Vicente y las Granadinas", "Santa Lucía", "Surinam",
	"Trinidad y Tobago",
]


def inicializar_archivos():
	if not os.path.exists(DB_FILE):
		with open(DB_FILE, "w", newline="", encoding="utf-8") as archivo:
			csv.writer(archivo).writerow(["email", "password", "status"])

	if not os.path.exists(DB_REGISTROS) or os.path.getsize(DB_REGISTROS) == 0:
		with open(DB_REGISTROS, "w", newline="", encoding="utf-8") as archivo:
			csv.writer(archivo).writerow(REGISTRO_COLUMNAS)
		return

	with open(DB_REGISTROS, "r", newline="", encoding="utf-8") as archivo:
		lector = csv.DictReader(archivo)
		campos_existentes = lector.fieldnames or []
		if all(campo in campos_existentes for campo in REGISTRO_COLUMNAS):
			return
		registros_antiguos = list(lector)

	with open(DB_REGISTROS, "w", newline="", encoding="utf-8") as archivo:
		writer = csv.DictWriter(archivo, fieldnames=REGISTRO_COLUMNAS)
		writer.writeheader()
		for registro in registros_antiguos:
			writer.writerow({
				"Fecha": registro.get("Fecha", ""),
				"Año": registro.get("Año", str(datetime.now().year)),
				"Usuario": registro.get("Usuario", ""),
				"Pais": registro.get("Pais", "Honduras"),
				"Tema": registro.get("Tema", "Materia General / Otro"),
				"Cliente": registro.get("Cliente", ""),
				"Identidad": registro.get("Identidad", ""),
				"Juzgado": registro.get("Juzgado", ""),
			})


inicializar_archivos()

st.session_state.setdefault("autenticado", False)
st.session_state.setdefault("usuario_actual", None)


def guardar_usuario(email, password, status):
	with open(DB_FILE, "a", newline="", encoding="utf-8") as archivo:
		csv.writer(archivo).writerow([email, password, status])


def obtener_usuarios():
	with open(DB_FILE, "r", newline="", encoding="utf-8") as archivo:
		return list(csv.DictReader(archivo))


def actualizar_status(email_objetivo, nuevo_status):
	usuarios = obtener_usuarios()
	with open(DB_FILE, "w", newline="", encoding="utf-8") as archivo:
		campos = ["email", "password", "status"]
		writer = csv.DictWriter(archivo, fieldnames=campos)
		writer.writeheader()
		for usuario in usuarios:
			if usuario["email"] == email_objetivo:
				usuario["status"] = nuevo_status
			writer.writerow(usuario)


def eliminar_usuario(email_objetivo):
	usuarios = obtener_usuarios()
	with open(DB_FILE, "w", newline="", encoding="utf-8") as archivo:
		campos = ["email", "password", "status"]
		writer = csv.DictWriter(archivo, fieldnames=campos)
		writer.writeheader()
		for usuario in usuarios:
			if usuario["email"] != email_objetivo:
				writer.writerow(usuario)


def guardar_registro(email, anio, pais, tema, nombre, identidad, juzgado):
	with open(DB_REGISTROS, "a", newline="", encoding="utf-8") as archivo:
		csv.writer(archivo).writerow([
			datetime.now().strftime("%Y-%m-%d %H:%M"),
			anio,
			email,
			pais,
			tema,
			nombre,
			identidad,
			juzgado,
		])


def obtener_registros():
	with open(DB_REGISTROS, "r", newline="", encoding="utf-8") as archivo:
		return list(csv.DictReader(archivo))


def obtener_registros_filtrados(email_usuario):
	todos_los_registros = obtener_registros()
	if email_usuario in ("Admin", "admin@lexflow.com"):
		return todos_los_registros
	return [
		registro
		for registro in todos_los_registros
		if registro["Usuario"] == email_usuario
	]


def obtener_registros_dataframe():
	columnas = REGISTRO_COLUMNAS
	if not os.path.exists(DB_REGISTROS):
		return pd.DataFrame(columns=columnas)
	datos = pd.read_csv(DB_REGISTROS, dtype=str).fillna("")
	if "Tema" not in datos.columns:
		datos["Tema"] = "Materia General / Otro"
	if "Año" not in datos.columns:
		datos["Año"] = str(datetime.now().year)
	if "Pais" not in datos.columns:
		datos["Pais"] = "Honduras"
	return datos.reindex(columns=columnas, fill_value="")


def extraer_datos_de_archivo(archivo):
	nombre = ""
	identidad = ""
	juzgado = ""
	tema = ""
	anio = ""
	pais = "Honduras"
	texto = ""
	try:
		nombre_archivo_original = archivo.name
		nombre_archivo = nombre_archivo_original.lower()
		if nombre_archivo.endswith(".docx"):
			documento = docx.Document(archivo)
			texto = " ".join(
				parrafo.text for parrafo in documento.paragraphs
			)
		elif nombre_archivo.endswith(".pdf"):
			lector = pypdf.PdfReader(archivo)
			texto = " ".join(
				pagina.extract_text() or "" for pagina in lector.pages
			)
		elif nombre_archivo.endswith((".xlsx", ".csv")):
			if nombre_archivo.endswith(".csv"):
				datos = pd.read_csv(archivo, dtype=str).fillna("")
			else:
				datos = pd.read_excel(archivo, dtype=str).fillna("")
			datos.columns = [str(columna).strip().lower() for columna in datos.columns]
			if not datos.empty:
				if "nombre" in datos.columns:
					nombre = str(datos.iloc[0]["nombre"])
				if "identidad" in datos.columns:
					identidad = str(datos.iloc[0]["identidad"])
				if "juzgado" in datos.columns:
					juzgado = str(datos.iloc[0]["juzgado"])
				if "tema" in datos.columns:
					tema = str(datos.iloc[0]["tema"])
				if "pais" in datos.columns:
					pais = str(datos.iloc[0]["pais"])
				if "año" in datos.columns:
					anio = str(datos.iloc[0]["año"])
				elif "anio" in datos.columns:
					anio = str(datos.iloc[0]["anio"])
			texto = datos.to_string()

		if not identidad:
			coincidencia_identidad = re.search(
				r"\b([01]\d{3}-?\d{4}-?\d{5})\b",
				texto,
			)
			if coincidencia_identidad:
				identidad = coincidencia_identidad.group(1).replace("-", "")

		if not nombre:
			coincidencia_nombre = re.search(
				r"(?i)(?:nombre|cliente|compareciente|promovente|imputado)"
				r"\s*[:\-]\s*([A-ZÁÉÍÓÚÑa-záéíóúñ\s]{5,40})",
				texto,
			)
			if coincidencia_nombre:
				nombre = coincidencia_nombre.group(1).strip().title()
			else:
				coincidencia_nombre = re.search(
					r"(?i)(?:yo,|el señor|la señora|comparece)\s+"
					r"([A-ZÁÉÍÓÚÑa-záéíóúñ\s]{5,40}?),\s+"
					r"(?:mayor de edad|de\s+este|con\s+identidad|"
					r"hondureñ|abogado)",
					texto,
				)
				if coincidencia_nombre:
					nombre = coincidencia_nombre.group(1).strip().title()

		if not juzgado:
			coincidencia_juzgado = re.search(
				r"(?i)((?:Juzgado|Tribunal|Corte|Sala)\s+"
				r"(?:de\s+letras\s+de\s+lo\s+|de\s+lo\s+|"
				r"civil|penal|familia|laboral|contra\s+la\s+corrupción)?"
				r"[a-zA-ZáéíóúÁÉÍÓÚñÑ\s]{3,30})",
				texto,
			)
			if coincidencia_juzgado:
				juzgado = coincidencia_juzgado.group(1).strip().title()

		if not anio:
			coincidencia_anio = re.search(
				r"\b(20[0-3][0-9]|19[8-9][0-9])\b",
				texto,
			)
			if coincidencia_anio:
				anio = coincidencia_anio.group(1)

		for pais_candidato in LISTA_PAISES:
			if re.search(re.escape(pais_candidato), texto, re.IGNORECASE) or re.search(
				re.escape(pais_candidato), nombre_archivo, re.IGNORECASE
			):
				pais = pais_candidato
				break

		if not identidad:
			coincidencia_archivo_id = re.search(
				r"\b([01]\d{12})\b",
				nombre_archivo_original,
			)
			if coincidencia_archivo_id:
				identidad = coincidencia_archivo_id.group(1)

		if not nombre:
			nombre_base = os.path.splitext(nombre_archivo_original)[0]
			palabras = [
				palabra for palabra in re.split(r"[_\-\s]+", nombre_base)
				if len(palabra) > 2 and not palabra.isdigit()
			]
			nombre = (
				" ".join(palabra.capitalize() for palabra in palabras[:2])
				if len(palabras) >= 2
				else "Cliente_General"
			)

		anio = anio or str(datetime.now().year)
		identidad = identidad or "0000000000000"
		juzgado = juzgado or "Juzgado General"

		temas_diccionario = {
			"Derecho Constitucional": (
				r"(?i)(recurso de amparo|inconstitucionalidad|derechos fundamentales|"
				r"garantías constitucionales|corte suprema|sala de lo constitucional)"
			),
			"Derecho Civil": (
				r"(?i)(demanda ordinaria|título ejecutivo|embargo|bienes raíces|"
				r"deuda|daños y perjuicios|contrato civil|sucesiones|herencia|"
				r"testamento|propiedad|juzgado de letras civil)"
			),
			"Derecho Penal": (
				r"(?i)(querella|denuncia|imputado|delito|requerimiento fiscal|"
				r"juzgado de letras de lo penal|habeas corpus|prisión preventiva|"
				r"audiencia inicial|sentencia absolutoria|sentencia condenatoria)"
			),
			"Derecho de Familia": (
				r"(?i)(pensión alimenticia|alimentos|demanda de alimentos|divorcio|"
				r"vínculo matrimonial|guarda y custodia|patria potestad|"
				r"reconocimiento de paternidad|adopción)"
			),
			"Derecho Laboral": (
				r"(?i)(prestaciones|despido injustificado|derechos laborales|"
				r"código del trabajo|juzgado de letras del trabajo|sindicato|"
				r"contrato individual de trabajo|riesgos profesionales)"
			),
			"Derecho Mercantil": (
				r"(?i)(sociedad anónima|comerciante individual|título valor|"
				r"letra de cambio|pagaré|cheque|quiebra|competencia desleal|"
				r"registro mercantil|acta constitutiva)"
			),
			"Derecho Administrativo": (
				r"(?i)(acto administrativo|impugnación|resolución administrativa|"
				r"contratación del estado|licitación|servidor público|ministerio|"
				r"secretaría de estado|lo pja)"
			),
			"Derecho Tributario y Fiscal": (
				r"(?i)(impuestos|sar|servicio de administración de rentas|"
				r"defraudación fiscal|reparo|tributo|arancel|código tributario)"
			),
			"Derecho Notarial y Registral": (
				r"(?i)(acta notarial|escritura pública|instrumento público|"
				r"protocolo notarial|testimonio|carta poder|registro de la propiedad|"
				r"auténtica|traspaso)"
			),
			"Derecho Agrario": (
				r"(?i)(reforma agraria|título de propiedad|ina|instituto nacional agrario|"
				r"expropiación|tierras|campesino|juzgado de letras de inquilinato)"
			),
			"Derecho Ambiental": (
				r"(?i)(licencia ambiental|estudio de impacto|contaminación|"
				r"recursos naturales|miambiente|delito ambiental)"
			),
			"Derecho Internacional": (
				r"(?i)(extradición|tratado internacional|exhorto|apostilla|asilo|"
				r"migración|derecho internacional público|privado)"
			),
			"Derecho de Propiedad Intelectual": (
				r"(?i)(marca|patente|derechos de autor|propiedad industrial|"
				r"registro de marcas|invención)"
			),
			"Derecho de Niñez y Adolescencia": (
				r"(?i)(dinnaf|menor infractor|interés superior del niño|"
				r"medidas de protección|juzgado de la niñez)"
			),
			"Derecho Informático": (
				r"(?i)(delito cibernético|firma electrónica|protección de datos|"
				r"comercio electrónico|phishing)"
			),
		}
		if not tema:
			for categoria, patron in temas_diccionario.items():
				if re.search(patron, texto) or re.search(patron, nombre_archivo):
					tema = categoria
					break
			tema = tema or "Materia General / Otro"

		return anio, pais, tema, nombre, identidad, juzgado
	except Exception:
		return (
			str(datetime.now().year),
			"Honduras",
			"Materia General / Otro",
			"Archivo_Revisar",
			"0000000000000",
			"Juzgado General",
		)


def generar_documento_honduras(tipo_documento, datos):
	documento = docx.Document()
	documento.add_heading("DOCUMENTO LEGAL - HONDURAS", level=1)
	documento.add_paragraph(f"TEMA: {tipo_documento.upper()}")
	documento.add_paragraph(
		f"FECHA: {datetime.now().strftime('%d de %B de %Y')}"
	)
	documento.add_paragraph("-" * 80)

	if tipo_documento == "Demanda de Alimentos":
		documento.add_paragraph(
			f"EXPEDIENTE: Correspondiente al año {datos['anio']}"
		)
		documento.add_paragraph(
			f"SEÑOR JUEZ DE LETRAS DE FAMILIA DE {datos['juzgado'].upper()}"
		)
		documento.add_paragraph(
			f"Yo, {datos['nombre']}, con número de identidad "
			f"{datos['identidad']}, mayor de edad, hondureño/a, comparezco "
			"respetuosamente interponiendo demanda de pensión alimenticia."
		)
		documento.add_paragraph(
			"HECHOS:\nPRIMERO: Que existen obligaciones alimentarias "
			"pendientes.\nSEGUNDO: Que corresponde fijar la pensión conforme a derecho."
		)
		documento.add_paragraph(
			"PETICIÓN:\nSolicito admitir la presente demanda y darle el trámite legal."
		)
	elif tipo_documento == "Carta Poder Notarial":
		documento.add_paragraph("INSTRUMENTO PÚBLICO - CARTA PODER")
		documento.add_paragraph(
			f"Comparece {datos['nombre']}, con identidad número "
			f"{datos['identidad']}, quien otorga carta poder para ser representado "
			f"ante {datos['juzgado']}."
		)
	elif tipo_documento == "Acta Notarial de Declaración Jurada":
		documento.add_paragraph("ACTA NOTARIAL DE DECLARACIÓN JURADA")
		documento.add_paragraph(
			f"A solicitud de {datos['nombre']}, con identidad "
			f"{datos['identidad']}, se hace constar bajo juramento la declaración "
			"para los fines legales correspondientes."
		)
	elif tipo_documento == "Requerimiento / Citación":
		documento.add_paragraph("AVISO DE CITACIÓN LEGAL")
		documento.add_paragraph(
			f"Dirigido a: {datos['nombre']}\n"
			f"Identidad: {datos['identidad']}\n"
			f"Autoridad: {datos['juzgado']}"
		)
		documento.add_paragraph(
			"Por medio de la presente se le cita para que comparezca dentro "
			"del plazo legal establecido."
		)
	else:
		documento.add_paragraph(f"DOCUMENTO GENERAL - {datos['nombre']}")
		documento.add_paragraph(
			f"Identidad: {datos['identidad']}\nJuzgado: {datos['juzgado']}"
		)

	documento.add_paragraph("\n\n" + "_" * 34)
	documento.add_paragraph("Firma del Profesional del Derecho / Notario")
	buffer = io.BytesIO()
	documento.save(buffer)
	buffer.seek(0)
	return buffer


def generar_documento_internacional(pais, tipo_documento, datos):
	documento = docx.Document()
	documento.add_heading(f"REPÚBLICA DE {pais.upper()}", level=0)
	documento.add_paragraph(f"JURISDICCIÓN: {datos['juzgado'].upper()}")
	documento.add_paragraph(f"ASUNTO: {tipo_documento.upper()}")
	documento.add_paragraph(f"AÑO DE EXPEDIENTE: {datos['anio']}")
	documento.add_paragraph(
		f"FECHA DE EMISIÓN: {datetime.now().strftime('%d de %B de %Y')}"
	)
	documento.add_paragraph("=" * 70)
	documento.add_paragraph(
		f"MARCO LEGAL APLICABLE:\nConstitución Política de {pais}, "
		"tratados internacionales ratificados y códigos sustantivos y "
		f"procesales vigentes en {pais}."
	)

	if tipo_documento == "Demanda Ordinaria / Civil":
		documento.add_heading("I. INTERPOSICIÓN DE DEMANDA", level=2)
		documento.add_paragraph(
			f"SEÑOR JUEZ O MAGISTRADO EN MATERIA CIVIL DE {pais.upper()}:\n\n"
			f"Yo, {datos['nombre']}, con documento de identidad número "
			f"{datos['identidad']}, comparezco promoviendo demanda ordinaria "
			"civil por incumplimiento contractual y daños y perjuicios."
		)
		documento.add_heading("II. RELACIÓN DE HECHOS", level=2)
		documento.add_paragraph(
			"PRIMERO: Existe una relación jurídica con obligaciones exigibles.\n"
			"SEGUNDO: La parte demandada incurrió en mora e incumplimiento.\n"
			"TERCERO: Las gestiones extrajudiciales no produjeron resultado."
		)
		documento.add_heading("III. FUNDAMENTOS DE DERECHO", level=2)
		documento.add_paragraph(
			f"La acción se fundamenta en la Constitución y en los códigos Civil "
			f"y Procesal Civil vigentes de {pais}."
		)
		documento.add_heading("IV. PETITORIO", level=2)
		documento.add_paragraph(
			"Solicito admitir la demanda, emplazar a la parte demandada y dictar "
			"sentencia condenando al pago de principal, intereses y costas."
		)
	elif tipo_documento == "Demanda de Pensión Alimenticia (Familia)":
		documento.add_heading("I. INTERPOSICIÓN DE DEMANDA DE ALIMENTOS", level=2)
		documento.add_paragraph(
			f"SEÑOR JUEZ DE FAMILIA DE {pais.upper()}:\n\n"
			f"Yo, {datos['nombre']}, con documento número {datos['identidad']}, "
			"comparezco entablando demanda de pensión alimenticia provisional y definitiva."
		)
		documento.add_heading("II. RELACIÓN DE HECHOS", level=2)
		documento.add_paragraph(
			"PRIMERO: Existen hijos menores que dependen del sustento familiar.\n"
			"SEGUNDO: La persona obligada ha incumplido sus deberes de alimentos, "
			"educación, salud y vestuario."
		)
		documento.add_heading("III. FUNDAMENTACIÓN Y PETITORIO", level=2)
		documento.add_paragraph(
			f"Solicito fijar una pensión provisional y tramitar la demanda conforme "
			f"a las normas de familia de {pais}."
		)
	elif tipo_documento == "Querella / Denuncia Penal":
		documento.add_heading("I. INTERPOSICIÓN DE QUERELLA O DENUNCIA", level=2)
		documento.add_paragraph(
			f"FISCALÍA O TRIBUNAL PENAL DE {pais.upper()}:\n\n"
			f"Yo, {datos['nombre']}, titular del documento {datos['identidad']}, "
			"comparezco denunciando hechos que podrían constituir ilícitos penales."
		)
		documento.add_heading("II. HECHOS Y DERECHO APLICABLE", level=2)
		documento.add_paragraph(
			f"Los hechos deberán investigarse conforme al Código Penal, el Código "
			f"Procesal Penal y la Constitución de {pais}."
		)
		documento.add_heading("III. PETICIÓN", level=2)
		documento.add_paragraph(
			"Solicito iniciar la investigación, citar a las partes y dictar las "
			"medidas cautelares que correspondan."
		)
	elif tipo_documento == "Recurso de Amparo Constitucional":
		documento.add_heading("INTERPOSICIÓN DE RECURSO DE AMPARO", level=2)
		documento.add_paragraph(
			f"SALA CONSTITUCIONAL O TRIBUNAL DE GARANTÍAS DE {pais.upper()}:\n\n"
			f"Yo, {datos['nombre']}, con documento {datos['identidad']}, interpongo "
			"acción de amparo contra un acto de autoridad que vulnera derechos fundamentales."
		)
		documento.add_paragraph(
			"PETITORIO:\nSolicito admitir el recurso, suspender provisionalmente el "
			"acto reclamado y restituir los derechos afectados."
		)
	elif tipo_documento == "Hábeas Corpus (Exhibición Personal)":
		documento.add_heading("SOLICITUD DE HÁBEAS CORPUS", level=2)
		documento.add_paragraph(
			f"AUTORIDAD COMPETENTE DE {pais.upper()}:\n\n"
			f"Comparece {datos['nombre']}, con identidad {datos['identidad']}, "
			"solicitando la exhibición personal ante una detención ilegal."
		)
		documento.add_paragraph(
			"Solicito verificar inmediatamente la legalidad de la privación de "
			"libertad y ordenar el cese de cualquier restricción ilegítima."
		)
	elif tipo_documento == "Carta Poder Notarial":
		documento.add_heading("INSTRUMENTO PÚBLICO: CARTA PODER", level=2)
		documento.add_paragraph(
			f"En {pais}, ante Notario Público, comparece {datos['nombre']}, con "
			f"documento {datos['identidad']}, quien otorga carta poder amplia y "
			f"suficiente para actuar ante {datos['juzgado']}."
		)
	elif tipo_documento == "Acta Notarial de Declaración Jurada":
		documento.add_heading("ACTA NOTARIAL DE DECLARACIÓN", level=2)
		documento.add_paragraph(
			f"A solicitud de {datos['nombre']}, documento {datos['identidad']}, "
			f"se hace constar bajo juramento la declaración para surtir efectos "
			f"legales en {pais}."
		)
	elif tipo_documento == "Escritura Pública / Contrato":
		documento.add_heading("ESCRITURA PÚBLICA DE CONTRATO", level=2)
		documento.add_paragraph(
			f"En {pais}, comparece {datos['nombre']}, con identidad {datos['identidad']}, "
			"y conviene celebrar el contrato sujeto a las normas civiles y mercantiles vigentes."
		)
	elif tipo_documento == "Contestación de Demanda":
		documento.add_heading("CONTESTACIÓN DE DEMANDA", level=2)
		documento.add_paragraph(
			f"Ante el tribunal de {pais}, {datos['nombre']}, con documento "
			f"{datos['identidad']}, comparece dando contestación formal a la demanda."
		)
		documento.add_paragraph(
			"Se oponen las excepciones y defensas que correspondan conforme a la ley."
		)
	else:
		documento.add_heading(f"ESCRITO LEGAL GENERAL - {pais.upper()}", level=2)
		documento.add_paragraph(
			f"Promovente: {datos['nombre']}\nIdentidad: {datos['identidad']}\n"
			f"Instancia: {datos['juzgado']}"
		)
		documento.add_paragraph(
			"Documento redactado bajo los parámetros del ordenamiento jurídico "
			"aplicable, sujeto a revisión profesional."
		)
	documento.add_paragraph("\n\n" + "_" * 40)
	documento.add_paragraph(f"Firma, Sello y Colegiación Profesional ({pais})")
	buffer = io.BytesIO()
	documento.save(buffer)
	buffer.seek(0)
	return buffer


def mostrar_generador(pais_jurisdiccion="Honduras"):
	st.subheader("Generador de Documentos Legales")
	if not DOCX_DISPONIBLE:
		st.error("Atención: la librería de Word no está cargada correctamente.")
		return

	modo_generacion = st.radio(
		"Modo de trabajo:",
		["Generador Individual", "Generación Masiva (Lotes)"],
		horizontal=True,
	)
	st.write("---")

	if modo_generacion == "Generador Individual":
		st.write("Redactar un solo documento")
		archivo_plantilla = st.file_uploader(
			"Sube tu plantilla (.docx)",
			type=["docx"],
		)
		tema_manual = st.selectbox(
			"Clasificación del Documento:",
			[
				"Familia_y_Alimentos",
				"Actas_Notariales",
				"Materia_Penal",
				"Materia_Laboral",
				"Materia_Civil",
				"Derecho Constitucional",
				"Derecho Civil",
				"Derecho Penal",
				"Derecho de Familia",
				"Derecho Laboral",
				"Derecho Mercantil",
				"Derecho Administrativo",
				"Derecho Tributario y Fiscal",
				"Derecho Notarial y Registral",
				"Derecho Agrario",
				"Derecho Ambiental",
				"Derecho Internacional",
				"Derecho de Propiedad Intelectual",
				"Derecho de Niñez y Adolescencia",
				"Derecho Informático",
				"Materia General / Otro",
			],
		)
		anio_manual = st.text_input(
			"Año del expediente",
			value=str(datetime.now().year),
		)
		nombre_cliente = st.text_input("Nombre del cliente")
		identidad = st.text_input("Identidad")
		juzgado = st.text_input("Juzgado")

		if st.button("Generar y Guardar"):
			if not archivo_plantilla or not nombre_cliente:
				st.warning(
					"Asegúrate de subir la plantilla y escribir el nombre del cliente."
				)
				return
			try:
				doc = DocxTemplate(archivo_plantilla)
				doc.render({
					"pais": pais_jurisdiccion,
					"nombre": nombre_cliente,
					"identidad": identidad,
					"juzgado": juzgado,
					"tema": tema_manual,
					"año": anio_manual,
				})
				doc.save("documento_final.docx")
				guardar_registro(
					st.session_state.usuario_actual,
					anio_manual,
					pais_jurisdiccion,
					tema_manual,
					nombre_cliente,
					identidad,
					juzgado,
				)
				st.success("¡Documento generado con éxito!")
				with open("documento_final.docx", "rb") as archivo:
					st.download_button(
						label="Descargar Documento",
						data=archivo,
						file_name="documento_final.docx",
					)
			except Exception as error:
				st.error(f"Ocurrió un error al procesar la plantilla: {error}")
	else:
		st.write("Generación automatizada múltiple")
		st.info(
			"Sube una plantilla Word y un Excel o CSV con las columnas "
			"nombre, identidad y juzgado."
		)
		plantilla_masiva = st.file_uploader(
			"1. Sube tu plantilla (.docx)",
			type=["docx"],
			key="plantilla_masiva",
		)
		excel_datos = st.file_uploader(
			"2. Sube tu Excel (.xlsx) o CSV",
			type=["xlsx", "csv"],
			key="excel_datos",
		)

		if st.button("Analizar y Generar Carpetas"):
			if not plantilla_masiva or not excel_datos:
				st.warning("Debes subir tanto la plantilla como el archivo Excel o CSV.")
				return
			try:
				if excel_datos.name.lower().endswith(".csv"):
					datos_lote = pd.read_csv(excel_datos, dtype=str).fillna("")
				else:
					datos_lote = pd.read_excel(
						excel_datos,
						dtype=str,
					).fillna("")

				datos_lote.columns = [
					str(columna).strip().lower()
					for columna in datos_lote.columns
				]
				columnas_requeridas = {"nombre", "identidad", "juzgado"}
				if not columnas_requeridas.issubset(datos_lote.columns):
					st.error(
						"El archivo debe tener las columnas: nombre, identidad y juzgado."
					)
					return
				if datos_lote.empty:
					st.warning("El archivo de datos no contiene registros.")
					return

				st.write(f"Se detectaron **{len(datos_lote)} documentos** en total.")
				conteo_juzgados = datos_lote["juzgado"].value_counts()
				if not conteo_juzgados.empty:
					st.write("**Documentos por juzgado:**")
					st.write(conteo_juzgados.to_dict())

				zip_buffer = io.BytesIO()
				with zipfile.ZipFile(
					zip_buffer,
					"w",
					compression=zipfile.ZIP_DEFLATED,
				) as archivo_zip:
					for indice, fila in datos_lote.iterrows():
						doc = DocxTemplate(plantilla_masiva)
						doc.render({
							"nombre": fila["nombre"],
							"identidad": fila["identidad"],
							"juzgado": fila["juzgado"],
						})
						documento_buffer = io.BytesIO()
						doc.save(documento_buffer)
						nombre_seguro = str(fila["nombre"]).strip()
						for caracter in "\\/:*?\"<>|":
							nombre_seguro = nombre_seguro.replace(caracter, "_")
						nombre_seguro = nombre_seguro.replace(" ", "_") or "cliente"
						juzgado_seguro = str(fila["juzgado"]).strip()
						for caracter in "\\/:*?\"<>|":
							juzgado_seguro = juzgado_seguro.replace(caracter, "_")
						juzgado_seguro = juzgado_seguro.replace(" ", "_") or "sin_juzgado"
						ruta = (
							f"{juzgado_seguro}/Expediente_"
							f"{nombre_seguro}_{indice + 1}.docx"
						)
						archivo_zip.writestr(ruta, documento_buffer.getvalue())
						guardar_registro(
							st.session_state.usuario_actual,
							str(datetime.now().year),
							pais_jurisdiccion,
							"General",
							str(fila["nombre"]),
							str(fila["identidad"]),
							str(fila["juzgado"]),
						)

				st.success("¡Lote generado y guardado en el historial!")
				st.download_button(
					label="Descargar Carpetas Organizadoras (.zip)",
					data=zip_buffer.getvalue(),
					file_name="Paquete_Expedientes_Clasificados.zip",
					mime="application/zip",
				)
			except Exception as error:
				st.error(f"Error procesando los archivos masivos: {error}")


if not st.session_state.autenticado:
	st.image(
		"https://images.unsplash.com/photo-1589829085413-56de8ae18c73?"
		"auto=format&fit=crop&w=1200&q=80",
		width="stretch",
	)
	st.title("LexFlow Studio")
	st.subheader("Sistema de Automatización Legal")
	st.write("---")
	menu = st.radio(
		"Acceso al sistema:",
		["Iniciar Sesión", "Registrarse"],
		horizontal=True,
	)
	st.write("---")

	if menu == "Registrarse":
		st.write("Solicitar Acceso")
		email = st.text_input("Correo Electrónico")
		password = st.text_input("Contraseña", type="password")
		if st.button("Enviar Solicitud"):
			usuarios = obtener_usuarios()
			if email and password and any(
				usuario["email"] == email for usuario in usuarios
			):
				st.error("Este correo ya está registrado.")
			elif not email or not password:
				st.warning("Por favor, llena todos los campos.")
			else:
				guardar_usuario(email, password, "pendiente")
				st.success("¡Solicitud enviada! El administrador la revisará pronto.")

	elif menu == "Iniciar Sesión":
		st.write("Entrar a mi cuenta")
		email = st.text_input("Correo")
		password = st.text_input("Contraseña", type="password")
		if st.button("Entrar"):
			if email == "admin@lexflow.com" and password == "Comida@1":
				st.session_state.autenticado = True
				st.session_state.usuario_actual = "Admin"
				st.rerun()
			usuarios = obtener_usuarios()
			usuario = next(
				(
					item for item in usuarios
					if item["email"] == email and item["password"] == password
				),
				None,
			)
			if usuario and usuario["status"] == "aprobado":
				st.session_state.autenticado = True
				st.session_state.usuario_actual = email
				st.rerun()
			elif usuario and usuario["status"] == "pendiente":
				st.warning(
					"Tu cuenta aún está pendiente de aprobación por el administrador."
				)
			else:
				st.error("Credenciales incorrectas.")
else:
	st.image("https://cdn-icons-png.flaticon.com/512/3011/3011984.png", width=80)
	columna_titulo, columna_logout = st.columns([3, 1])
	with columna_titulo:
		st.title("LexFlow Studio - Panel de Trabajo")
	with columna_logout:
		if st.button("Cerrar Sesión"):
			st.session_state.autenticado = False
			st.session_state.usuario_actual = None
			st.rerun()

	tab_generador, tab_redaccion, tab_clasificador, tab_historial, tab_admin = st.tabs([
		"Generador",
		"Redacción y Escritura",
		"Clasificador de Múltiples Archivos",
		"Búsqueda y Datos",
		"Administración",
	])

	with tab_generador:
		st.write(f"Conectado como: {st.session_state.usuario_actual}")
		pais_plantilla = st.selectbox(
			"País de Jurisdicción:",
			LISTA_PAISES,
			key="pais_plantilla",
		)
		mostrar_generador(pais_plantilla)

		st.write("---")
		st.subheader("Acceso Rápido")
		url_de_tu_app = "https://lexflow-studio.streamlit.app"
		contenido_url = f"[InternetShortcut]\nURL={url_de_tu_app}\nIconIndex=0"
		st.download_button(
			"Descargar Acceso Directo (PC)",
			data=contenido_url,
			file_name="LexFlow_Studio.url",
			mime="application/octet-stream",
		)
		st.info("Para celular: usa 'Agregar a pantalla de inicio' en tu navegador.")

	with tab_redaccion:
		st.subheader("Redacción y Escritura Jurídica y Notarial")
		st.info(
			"Selecciona el tipo de documento y completa los datos necesarios "
			"para generar un archivo Word."
		)
		pais_redaccion = st.selectbox(
			"Seleccione País:",
			LISTA_PAISES,
			key="pais_redaccion",
		)
		tipo_redaccion = st.selectbox(
			"Tipo de Documento a Redactar:",
			[
				"Demanda Ordinaria / Civil",
				"Demanda de Pensión Alimenticia (Familia)",
				"Querella / Denuncia Penal",
				"Recurso de Amparo Constitucional",
				"Hábeas Corpus (Exhibición Personal)",
				"Carta Poder Notarial",
				"Acta Notarial de Declaración Jurada",
				"Escritura Pública / Contrato",
				"Contestación de Demanda",
				"Escrito de Petición General",
			],
		)
		anio_redaccion = st.text_input(
			"Año",
			value=str(datetime.now().year),
			key="redaccion_anio",
		)
		nombre_redaccion = st.text_input(
			"Nombre completo de la parte / cliente",
			key="redaccion_nombre",
		)
		identidad_redaccion = st.text_input(
			"Número de Identidad (13 dígitos)",
			key="redaccion_identidad",
		)
		juzgado_redaccion = st.text_input(
			"Juzgado o Institución asignada",
			key="redaccion_juzgado",
		)

		if st.button("Generar Documento Jurídico"):
			if nombre_redaccion and identidad_redaccion:
				datos_documento = {
					"anio": anio_redaccion,
					"nombre": nombre_redaccion,
					"identidad": identidad_redaccion,
					"juzgado": juzgado_redaccion,
				}
				archivo_generado = generar_documento_internacional(
					pais_redaccion,
					tipo_redaccion,
					datos_documento,
				)
				guardar_registro(
					st.session_state.usuario_actual,
					anio_redaccion,
					pais_redaccion,
					"Redacción Jurídica",
					nombre_redaccion,
					identidad_redaccion,
					juzgado_redaccion,
				)
				st.success(f"¡Documento de {tipo_redaccion} redactado con éxito!")
				st.download_button(
					label=f"Descargar {tipo_redaccion}.docx",
					data=archivo_generado,
					file_name=(
						f"{tipo_redaccion.replace(' ', '_')}_"
						f"{nombre_redaccion.replace(' ', '_')}.docx"
					),
					mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
				)
			else:
				st.warning("Por favor completa al menos el nombre y número de identidad.")

	with tab_clasificador:
		st.subheader("Subir y Clasificar Varios Archivos")
		st.info(
			"Puedes subir múltiples archivos Word, PDF, Excel o CSV al mismo tiempo. "
			"Se organizarán por rama, año e identidad."
		)
		archivos_subidos = st.file_uploader(
			"Selecciona tus documentos en bloque",
			type=["docx", "pdf", "xlsx", "csv"],
			accept_multiple_files=True,
			key="clasificador_docx",
		)

		if archivos_subidos and st.button("Analizar y Organizar"):
			resultados = []
			zip_buffer = io.BytesIO()
			with zipfile.ZipFile(
				zip_buffer,
				"w",
				compression=zipfile.ZIP_DEFLATED,
			) as archivo_zip:
				for archivo in archivos_subidos:
					anio, pais, tema, nombre, identidad, juzgado = extraer_datos_de_archivo(archivo)
					resultados.append({
						"Archivo": archivo.name,
						"Año": anio,
						"Pais": pais,
						"Tema": tema,
						"Nombre": nombre,
						"Identidad": identidad,
						"Juzgado": juzgado,
					})

					identidad_limpia = re.sub(r"[^A-Za-z0-9]", "", identidad)
					nombre_limpio = re.sub(r"[\\/:*?\"<>|]", "-", nombre)
					nombre_limpio = nombre_limpio.replace(" ", "_") or "No_Detectado"
					archivo.seek(0)
					tema_limpio = re.sub(r"[\\/:*?\"<>|]", "-", tema)
					ruta_carpeta = (
						f"Expedientes_Internacionales/{re.sub(r'[\\/:*?\"<>|]', '-', pais)}/"
						f"{tema_limpio}/"
						f"{anio}/"
						f"{identidad_limpia or 'Sin_Identidad'}_{nombre_limpio}"
					)
					ruta_archivo = f"{ruta_carpeta}/{archivo.name}"
					archivo_zip.writestr(ruta_archivo, archivo.read())
					guardar_registro(
						st.session_state.usuario_actual,
						anio,
						pais,
						tema,
						nombre,
						identidad,
						juzgado,
					)

			st.write("Análisis de Coincidencias")
			datos_resultados = pd.DataFrame(resultados)
			st.write(f"Se procesaron **{len(resultados)}** archivos.")
			st.write("Documentos por materia legal:")
			st.dataframe(
				datos_resultados["Tema"].value_counts(),
				width="stretch",
			)
			archivos_por_persona = (
				datos_resultados.groupby(["Identidad", "Nombre"])
				.size()
				.reset_index(name="Cantidad de Archivos")
				.sort_values("Cantidad de Archivos", ascending=False)
			)
			st.write("Resumen de Archivos Apilados por Persona:")
			st.dataframe(archivos_por_persona, width="stretch")
			columna_juzgado, columna_identidad = st.columns(2)
			with columna_juzgado:
				st.write("Documentos por Juzgado")
				st.dataframe(
					datos_resultados["Juzgado"].value_counts(),
					width="stretch",
				)
			with columna_identidad:
				st.write("Documentos por Identidad")
				st.dataframe(
					datos_resultados["Identidad"].value_counts(),
					width="stretch",
				)
			st.success("¡Archivos organizados exitosamente!")
			st.download_button(
				label="Descargar Expedientes Ordenados (.zip)",
				data=zip_buffer.getvalue(),
				file_name="Expedientes_Clasificados.zip",
				mime="application/zip",
			)

	with tab_historial:
		st.subheader("Buscar documentos")
		st.write("Aquí solo aparecen los documentos permitidos para tu cuenta.")
		datos = obtener_registros_dataframe()
		usuario_actual = st.session_state.usuario_actual
		if usuario_actual not in ("Admin", "admin@lexflow.com"):
			datos = datos[datos["Usuario"] == usuario_actual]

		busqueda = st.text_input(
			"Buscar por Año, Nombre, Identidad, Juzgado o Tema:",
			placeholder="Ej. Mario Roberto Suazo",
		)
		if busqueda:
			coincidencias = datos.apply(
				lambda fila: fila.astype(str).str.contains(
					busqueda,
					case=False,
					regex=False,
				).any(),
				axis=1,
			)
			datos = datos[coincidencias]

		if not datos.empty:
			if busqueda:
				st.write(f"Resultados encontrados: {len(datos)}")
			st.dataframe(datos, hide_index=True, width="stretch")
		else:
			st.info("Aún no hay documentos que coincidan con la búsqueda.")

	with tab_admin:
		if st.session_state.usuario_actual == "Admin":
			st.subheader("Panel de Control Maestro (Acceso Exclusivo)")
			st.write("Solo tú puedes ver esto. Acepta o rechaza solicitudes aquí:")
			usuarios = obtener_usuarios()
			pendientes = [
				usuario for usuario in usuarios if usuario["status"] == "pendiente"
			]
			if pendientes:
				for usuario in pendientes:
					st.markdown(f"Solicitud de acceso: **{usuario['email']}**")
					columna_aceptar, columna_rechazar, columna_info = st.columns([
						1,
						1,
						3,
					])
					with columna_aceptar:
						if st.button(
							"Aceptar",
							key=f"ok_{usuario['email']}",
						):
							actualizar_status(usuario["email"], "aprobado")
							st.rerun()
					with columna_rechazar:
						if st.button(
							"Rechazar",
							key=f"no_{usuario['email']}",
						):
							eliminar_usuario(usuario["email"])
							st.rerun()
					st.write("---")
			else:
				st.info("No tienes solicitudes pendientes en este momento.")
		else:
			st.error("ACCESO DENEGADO. Solo el Administrador puede ver esta sección.")

st.write("---")
st.markdown(
	"LexFlow Studio - Automatización Profesional"
)
