import csv
import io
import os
import zipfile
from datetime import datetime

import docx
import pandas as pd
import streamlit as st

try:
	from docxtpl import DocxTemplate
	DOCX_TEMPLATE_AVAILABLE = True
except ImportError:
	DOCX_TEMPLATE_AVAILABLE = False

st.set_page_config(
	page_title="LexFlow Studio",
	layout="wide",
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
	"San Vicente y las Granadinas", "Santa Lucía", "Surinam", "Trinidad y Tobago",
]


def inicializar_archivos():
	if not os.path.exists(DB_FILE):
		with open(DB_FILE, "w", newline="", encoding="utf-8") as archivo:
			csv.writer(archivo).writerow(["email", "password", "status"])
	if not os.path.exists(DB_REGISTROS) or os.path.getsize(DB_REGISTROS) == 0:
		with open(DB_REGISTROS, "w", newline="", encoding="utf-8") as archivo:
			csv.writer(archivo).writerow(REGISTRO_COLUMNAS)


def obtener_usuarios():
	with open(DB_FILE, "r", newline="", encoding="utf-8") as archivo:
		return list(csv.DictReader(archivo))


def guardar_usuario(email, password, status="pendiente"):
	with open(DB_FILE, "a", newline="", encoding="utf-8") as archivo:
		csv.writer(archivo).writerow([email, password, status])


def guardar_registro(usuario, pais, tema, cliente, identidad, juzgado):
	with open(DB_REGISTROS, "a", newline="", encoding="utf-8") as archivo:
		csv.writer(archivo).writerow([
			datetime.now().strftime("%Y-%m-%d %H:%M"),
			str(datetime.now().year), usuario, pais, tema, cliente, identidad, juzgado,
		])


def obtener_registros_dataframe(usuario):
	if not os.path.exists(DB_REGISTROS):
		return pd.DataFrame(columns=REGISTRO_COLUMNAS)
	datos = pd.read_csv(DB_REGISTROS, dtype=str).fillna("")
	if usuario != "Admin":
		datos = datos[datos["Usuario"] == usuario]
	return datos.reindex(columns=REGISTRO_COLUMNAS, fill_value="")


inicializar_archivos()
st.session_state.setdefault("autenticado", False)
st.session_state.setdefault("usuario_actual", None)


def generar_escrito_word(tipo_documento, pais, demandante, demandado, hechos):
	documento = docx.Document()
	documento.add_heading(tipo_documento.upper(), level=1)
	documento.add_paragraph(f"JURISDICCIÓN: {pais}")
	documento.add_paragraph(f"DEMANDANTE: {demandante}")
	documento.add_paragraph(f"DEMANDADO: {demandado}")
	documento.add_heading("I. RELACIÓN DE HECHOS", level=2)
	documento.add_paragraph(hechos)
	documento.add_heading("II. FUNDAMENTOS Y PETITORIO", level=2)
	documento.add_paragraph(
		"Complete este documento con las normas vigentes, pruebas y pretensiones "
		"concretas antes de su presentación profesional."
	)
	buffer = io.BytesIO()
	documento.save(buffer)
	buffer.seek(0)
	return buffer


def modulo_generador_escritos(usuario):
	st.header("Generador automatizado de documentos")
	tipo_documento = st.selectbox(
		"Tipo de escrito",
		["Demanda Civil", "Demanda Penal", "Demanda de Alimentos", "Recurso de Apelación"],
		key="app_tipo_documento",
	)
	pais = st.selectbox("Jurisdicción", LISTA_PAISES, key="app_pais_generador")
	columna_izquierda, columna_derecha = st.columns(2)
	with columna_izquierda:
		demandante = st.text_input("Parte demandante", key="app_demandante")
	with columna_derecha:
		demandado = st.text_input("Parte demandada", key="app_demandado")
	hechos = st.text_area("Relación de hechos", key="app_hechos")
	if st.button("Generar escrito", key="app_generar_escrito"):
		if not demandante or not demandado or not hechos:
			st.warning("Completa las partes y los hechos.")
			return
		archivo = generar_escrito_word(tipo_documento, pais, demandante, demandado, hechos)
		guardar_registro(usuario, pais, tipo_documento, demandante, "", "")
		st.success("Borrador generado. Revisa el contenido antes de usarlo.")
		st.download_button(
			"Descargar escrito Word",
			data=archivo,
			file_name=f"{tipo_documento.replace(' ', '_')}.docx",
			mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
			key="app_descargar_escrito",
		)


def modulo_clasificador(usuario):
	st.header("Clasificador masivo de expedientes")
	archivos = st.file_uploader(
		"Sube los documentos del expediente",
		type=["pdf", "docx", "xlsx", "csv", "txt"],
		accept_multiple_files=True,
		key="app_archivos_clasificador",
	)
	pais = st.selectbox("País del expediente", LISTA_PAISES, key="app_pais_clasificador")
	tema = st.selectbox(
		"Rama del derecho",
		["Civil", "Penal", "Familia", "Laboral", "Administrativo"],
		key="app_tema_clasificador",
	)
	if archivos:
		st.dataframe(
			{"Archivo": [archivo.name for archivo in archivos], "País": [pais] * len(archivos)},
			hide_index=True,
			width="stretch",
		)
	if st.button("Organizar lote", key="app_organizar_lote"):
		if not archivos:
			st.warning("Sube al menos un archivo.")
			return
		zip_buffer = io.BytesIO()
		with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as archivo_zip:
			for archivo in archivos:
				archivo.seek(0)
				archivo_zip.writestr(f"{pais}/{tema}/{archivo.name}", archivo.read())
				guardar_registro(usuario, pais, tema, archivo.name, "", "")
		st.success(f"{len(archivos)} archivos organizados.")
		st.download_button(
			"Descargar expediente organizado",
			data=zip_buffer.getvalue(),
			file_name="Expediente_Clasificado.zip",
			mime="application/zip",
			key="app_descargar_lote",
		)


def mostrar_app_principal():
	usuario = st.session_state.usuario_actual
	st.title("LexFlow Studio")
	st.caption(f"Sesión activa: {usuario}")
	if st.button("Cerrar sesión", key="app_logout"):
		st.session_state.autenticado = False
		st.session_state.usuario_actual = None
		st.rerun()
	tab_generador, tab_clasificador, tab_historial = st.tabs([
		"Generador de escritos",
		"Clasificador masivo",
		"Historial",
	])
	with tab_generador:
		modulo_generador_escritos(usuario)
	with tab_clasificador:
		modulo_clasificador(usuario)
	with tab_historial:
		st.header("Búsqueda y datos")
		datos = obtener_registros_dataframe(usuario)
		busqueda = st.text_input("Buscar en el historial", key="app_busqueda_historial")
		if busqueda:
			datos = datos[
				datos.apply(
					lambda fila: fila.astype(str).str.contains(busqueda, case=False, regex=False).any(),
					axis=1,
				)
			]
		st.dataframe(datos, hide_index=True, width="stretch")


def mostrar_login():
	st.title("LexFlow Studio")
	st.subheader("Sistema de automatización legal")
	modo = st.radio("Acceso", ["Iniciar sesión", "Registrarse"], horizontal=True, key="app_modo_acceso")
	if modo == "Registrarse":
		email = st.text_input("Correo", key="app_registro_email")
		password = st.text_input("Contraseña", type="password", key="app_registro_password")
		if st.button("Solicitar acceso", key="app_solicitar_acceso"):
			if not email or not password:
				st.warning("Completa correo y contraseña.")
			elif any(usuario["email"] == email for usuario in obtener_usuarios()):
				st.error("El correo ya está registrado.")
			else:
				guardar_usuario(email, password)
				st.success("Solicitud enviada para aprobación.")
	else:
		email = st.text_input("Correo", key="app_login_email")
		password = st.text_input("Contraseña", type="password", key="app_login_password")
		if st.button("Entrar", key="app_entrar"):
			if email == "admin@lexflow.com" and password == "Comida@1":
				st.session_state.autenticado = True
				st.session_state.usuario_actual = "Admin"
				st.rerun()
			usuario = next(
				(
					item for item in obtener_usuarios()
					if item["email"] == email
					and item["password"] == password
					and item["status"] == "aprobado"
				),
				None,
			)
			if usuario:
				st.session_state.autenticado = True
				st.session_state.usuario_actual = email
				st.rerun()
			else:
				st.error("Credenciales incorrectas o cuenta pendiente.")


if st.session_state.autenticado:
	mostrar_app_principal()
else:
	mostrar_login()
