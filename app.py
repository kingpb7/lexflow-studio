import os

import streamlit as st
from docxtpl import DocxTemplate

st.set_page_config(page_title="LexFlow Studio", page_icon="⚖️")

if "autenticado" not in st.session_state:
	st.session_state.autenticado = False


def leer_solicitudes():
	if not os.path.exists("solicitudes.txt"):
		return []
	with open("solicitudes.txt", "r", encoding="utf-8") as archivo:
		return [linea.strip().split(",") for linea in archivo.readlines()]


if not st.session_state.autenticado:
	st.image("https://cdn-icons-png.flaticon.com/512/3011/3011984.png", width=100)
	st.title("LexFlow Studio")
	st.subheader("Automatización de Documentos Legales")

	menu = st.radio("¿Qué deseas hacer?", ["Iniciar Sesión", "Registrarse"])

	if menu == "Registrarse":
		st.subheader("Solicitar Acceso")
		nombre = st.text_input("Tu Nombre Completo")
		email = st.text_input("Tu Correo Electrónico")
		if st.button("Enviar solicitud"):
			with open("solicitudes.txt", "a", encoding="utf-8") as archivo:
				archivo.write(f"{nombre},{email},pendiente\n")
			st.success("¡Solicitud enviada! El administrador revisará tu acceso pronto.")

	elif menu == "Iniciar Sesión":
		st.subheader("Iniciar Sesión")
		opcion_login = st.radio(
			"Opciones:",
			["Entrar con Contraseña", "¿Has olvidado tu contraseña?"],
			label_visibility="collapsed",
		)

		if opcion_login == "Entrar con Contraseña":
			contrasena = st.text_input("Contraseña:", type="password")

			if contrasena == "honduras2026":
				st.session_state.autenticado = True
				st.rerun()
			elif contrasena != "":
				st.error("Contraseña incorrecta.")

		elif opcion_login == "¿Has olvidado tu contraseña?":
			st.write("Ingresa el correo electrónico con el que te registraste:")
			correo_recuperacion = st.text_input("Correo electrónico:")
			if st.button("Enviar instrucciones de recuperación"):
				if correo_recuperacion != "":
					st.success(
						f"Se han enviado las instrucciones de recuperación al correo: "
						f"{correo_recuperacion}"
					)
					st.info(
						"Simulación de correo: Tu contraseña de acceso es: "
						"honduras2026"
					)
				else:
					st.warning("Por favor, ingresa un correo electrónico.")

else:
	st.image("https://cdn-icons-png.flaticon.com/512/3011/3011984.png", width=80)
	st.title("LexFlow Studio - Panel de Trabajo")

	if st.button("Cerrar Sesión"):
		st.session_state.autenticado = False
		st.rerun()

	st.write("---")
	st.subheader("Panel de Administración de Solicitudes")
	solicitudes = leer_solicitudes()
	if solicitudes:
		st.write("Usuarios pendientes de aprobación:")
		for indice, solicitud in enumerate(solicitudes):
			st.write(f"Usuario: {solicitud[0]} | Email: {solicitud[1]}")
			if st.button(
				f"Aprobar a {solicitud[0]}",
				key=f"aprobar_{indice}",
			):
				st.success(f"Usuario {solicitud[0]} aprobado por el administrador.")
	else:
		st.write("No hay nuevas solicitudes de registro pendientes.")

	st.write("---")
	st.subheader("Generador de Documentos Legales")
	archivo_plantilla = st.file_uploader(
		"Sube tu archivo plantilla.docx",
		type=["docx"],
	)
	nombre_cliente = st.text_input("Nombre del cliente")
	identidad = st.text_input("Número de identidad")
	juzgado = st.text_input("Juzgado")

	if st.button("Generar Documento"):
		if archivo_plantilla:
			try:
				doc = DocxTemplate(archivo_plantilla)
				contexto = {
					"nombre": nombre_cliente,
					"identidad": identidad,
					"juzgado": juzgado,
				}
				doc.render(contexto)
				doc.save("documento_final.docx")

				st.success("¡Documento generado con éxito!")
				with open("documento_final.docx", "rb") as archivo:
					st.download_button(
						label="Descargar Documento",
						data=archivo,
						file_name="documento_final.docx",
					)
			except Exception:
				st.error(
					"Ocurrió un error al procesar la plantilla: "
					"asegúrate de usar las etiquetas correctas "
					"{{nombre}}, {{identidad}} y {{juzgado}}."
				)
		else:
			st.warning("Por favor, sube primero la plantilla de Word.")

st.write("---")
st.markdown(
	"¿Necesitas ayuda o soporte técnico? "
	"[Haz clic aquí para ir a Facebook](https://www.facebook.com)"
)
