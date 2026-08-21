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
	"Fecha", "Usuario", "Tema", "Cliente", "Identidad", "Juzgado"
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
		if "Tema" in (lector.fieldnames or []):
			return
		registros_antiguos = list(lector)

	with open(DB_REGISTROS, "w", newline="", encoding="utf-8") as archivo:
		writer = csv.DictWriter(archivo, fieldnames=REGISTRO_COLUMNAS)
		writer.writeheader()
		for registro in registros_antiguos:
			writer.writerow({
				"Fecha": registro.get("Fecha", ""),
				"Usuario": registro.get("Usuario", ""),
				"Tema": "General",
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


def guardar_registro(email, tema, nombre, identidad, juzgado):
	with open(DB_REGISTROS, "a", newline="", encoding="utf-8") as archivo:
		csv.writer(archivo).writerow([
			datetime.now().strftime("%Y-%m-%d %H:%M"),
			email,
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
		datos["Tema"] = "General"
	return datos.reindex(columns=columnas, fill_value="")


def extraer_datos_de_archivo(archivo):
	nombre = "No_Detectado"
	identidad = "Desconocida"
	juzgado = "Sin_Clasificar"
	tema = "General"
	texto = ""
	try:
		nombre_archivo = archivo.name.lower()
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
			texto = datos.to_string()

		if identidad == "Desconocida":
			coincidencia_identidad = re.search(
				r"\b\d{4}-?\d{4}-?\d{5}\b",
				texto,
			)
			if coincidencia_identidad:
				identidad = coincidencia_identidad.group()

		if juzgado == "Sin_Clasificar":
			coincidencia_juzgado = re.search(
				r"(?i)((?:Juzgado|Tribunal|Corte)\s+de\s+"
				r"[a-zA-ZáéíóúÁÉÍÓÚñÑ]+"
				r"(?:\s+[a-zA-ZáéíóúÁÉÍÓÚñÑ]+){0,3})",
				texto,
			)
			if coincidencia_juzgado:
				juzgado = coincidencia_juzgado.group(1).strip().title()

		if nombre == "No_Detectado":
			coincidencia_nombre = re.search(
				r"(?i)(?:yo,|el señor|la señora|comparece)\s+"
				r"([A-ZÁÉÍÓÚÑ\s]{5,50}?),\s+"
				r"(?:mayor de edad|de\s+este|con\s+identidad|hondureñ)",
				texto,
			)
			if coincidencia_nombre:
				nombre = coincidencia_nombre.group(1).strip().title()

		temas_diccionario = {
			"Familia_y_Alimentos": (
				r"(?i)(pensión alimenticia|alimentos|demanda de alimentos|"
				r"divorcio|vínculo matrimonial|guarda y custodia|patria potestad)"
			),
			"Actas_Notariales": (
				r"(?i)(acta notarial|escritura pública|instrumento público|"
				r"protocolo notarial|testimonio|carta poder)"
			),
			"Materia_Penal": (
				r"(?i)(querella|denuncia|imputado|delito|requerimiento fiscal|"
				r"juzgado de letras de lo penal)"
			),
			"Materia_Laboral": (
				r"(?i)(prestaciones|despido injustificado|derechos laborales|"
				r"código del trabajo|juzgado de letras del trabajo)"
			),
			"Materia_Civil": (
				r"(?i)(demanda ordinaria|título ejecutivo|embargo|bienes raíces|"
				r"deuda|juzgado de letras de lo civil)"
			),
		}
		if tema == "General":
			for categoria, patron in temas_diccionario.items():
				if re.search(patron, texto):
					tema = categoria
					break

		return tema, nombre, identidad, juzgado
	except Exception:
		return "General", "Error de lectura", "Error", "Error"


def mostrar_generador():
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
				"Materia_General",
			],
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
					"nombre": nombre_cliente,
					"identidad": identidad,
					"juzgado": juzgado,
					"tema": tema_manual,
				})
				doc.save("documento_final.docx")
				guardar_registro(
					st.session_state.usuario_actual,
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

	tab_generador, tab_clasificador, tab_historial, tab_admin = st.tabs([
		"Generador",
		"Clasificador Inteligente",
		"Búsqueda y Datos",
		"Administración",
	])

	with tab_generador:
		st.write(f"Conectado como: {st.session_state.usuario_actual}")
		mostrar_generador()

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

	with tab_clasificador:
		st.subheader("Apilar y Ordenar Archivos (Word, PDF, Excel)")
		st.info(
			"Sube archivos de distintos tipos para extraer sus datos y agruparlos "
			"por identidad y nombre."
		)
		archivos_subidos = st.file_uploader(
			"Sube tus archivos aquí",
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
					tema, nombre, identidad, juzgado = extraer_datos_de_archivo(archivo)
					resultados.append({
						"Archivo": archivo.name,
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
						f"Expedientes_Organizados/{tema_limpio}/"
						f"{identidad_limpia or 'Sin_Identidad'}_{nombre_limpio}"
					)
					ruta_archivo = f"{ruta_carpeta}/{archivo.name}"
					archivo_zip.writestr(ruta_archivo, archivo.read())
					guardar_registro(
						st.session_state.usuario_actual,
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
			"Escribe nombre, identidad o juzgado para buscar:",
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
