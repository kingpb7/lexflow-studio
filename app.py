import csv
import os

import streamlit as st

try:
	from docxtpl import DocxTemplate
	DOCX_DISPONIBLE = True
except ImportError:
	DOCX_DISPONIBLE = False

st.set_page_config(page_title="LexFlow Studio", page_icon="⚖️")

DB_FILE = "usuarios.csv"
if not os.path.exists(DB_FILE):
	with open(DB_FILE, "w", newline="", encoding="utf-8") as archivo:
		csv.writer(archivo).writerow(["email", "password", "status"])

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


def mostrar_generador():
	st.subheader("Generador de Documentos Legales")
	if not DOCX_DISPONIBLE:
		st.error("Atención: la librería de Word no está cargada correctamente.")
		return

	archivo_plantilla = st.file_uploader(
		"Sube tu archivo plantilla.docx",
		type=["docx"],
	)
	nombre_cliente = st.text_input("Nombre del cliente")
	identidad = st.text_input("Número de identidad")
	juzgado = st.text_input("Juzgado")

	if st.button("Generar Documento"):
		if not archivo_plantilla:
			st.warning("Por favor, sube primero la plantilla de Word.")
			return
		try:
			doc = DocxTemplate(archivo_plantilla)
			doc.render({
				"nombre": nombre_cliente,
				"identidad": identidad,
				"juzgado": juzgado,
			})
			doc.save("documento_final.docx")
			st.success("¡Documento generado con éxito!")
			with open("documento_final.docx", "rb") as archivo:
				st.download_button(
					label="Descargar Documento",
					data=archivo,
					file_name="documento_final.docx",
				)
		except Exception as error:
			st.error(f"Ocurrió un error al procesar la plantilla: {error}")


if not st.session_state.autenticado:
	st.image("https://cdn-icons-png.flaticon.com/512/3011/3011984.png", width=100)
	st.title("LexFlow Studio")
	st.subheader("Sistema de Automatización Legal")
	menu = st.radio("Selecciona una opción:", ["Iniciar Sesión", "Registrarse"])

	if menu == "Registrarse":
		st.write("Crea tu cuenta para solicitar acceso.")
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
		st.subheader("Iniciar Sesión")
		email = st.text_input("Correo")
		password = st.text_input("Contraseña", type="password")
		if st.button("Entrar"):
			if email == "admin@lexflow.com" and password == "honduras2026":
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

		st.write("---")
		if st.session_state.usuario_actual == "Admin":
			st.subheader("Panel de Administración - Solicitudes de Acceso")
			usuarios = obtener_usuarios()
			pendientes = [
				usuario for usuario in usuarios if usuario["status"] == "pendiente"
			]
			if pendientes:
				st.write("Usuarios esperando aprobación:")
				for usuario in pendientes:
					columna_usuario, columna_accion = st.columns([2, 1])
					columna_usuario.write(f"{usuario['email']}")
					if columna_accion.button(
						"Aprobar acceso",
						key=f"aprobar_{usuario['email']}",
					):
						actualizar_status(usuario["email"], "aprobado")
						st.success(
							f"¡Usuario {usuario['email']} aprobado con éxito!"
						)
						st.rerun()
			else:
				st.info("No hay nuevas solicitudes de registro pendientes.")
			st.write("---")

		st.subheader("Generador de Documentos Legales")
		st.write(f"Conectado como: {st.session_state.usuario_actual}")
		mostrar_generador()

		st.write("---")
		st.subheader("Acceso Rápido")
		url_de_tu_app = "https://tu-enlace.streamlit.app"
		contenido_url = f"[InternetShortcut]\nURL={url_de_tu_app}\nIconIndex=0"
		st.download_button(
			"Descargar Acceso Directo (PC)",
			data=contenido_url,
			file_name="LexFlow_Studio.url",
			mime="application/octet-stream",
		)
		st.info("Para celular: usa 'Agregar a pantalla de inicio' en tu navegador.")

st.write("---")
st.markdown(
	"LexFlow Studio - Automatización Profesional"
)
