import streamlit as st
import io
from pypdf import PdfReader
import pandas as pd
import datetime
import zipfile
from docx import Document

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="LexFlow Studio - Sistema Jurídico", layout="wide")

# --- BASE DE DATOS EN MEMORIA (ESTADOS DE SESIÓN) ---
if 'users_db' not in st.session_state:
    st.session_state.users_db = {
        "admin@lexflow.com": {"password": "Comida@1", "status": "Aprobado", "role": "admin"}
    }
if 'historial_usuario' not in st.session_state:
    st.session_state.historial_usuario = {}
if 'usuario_actual' not in st.session_state:
    st.session_state.usuario_actual = None
if 'rol_actual' not in st.session_state:
    st.session_state.rol_actual = None

# --- REGISTRO EN HISTORIAL PRIVADO ---
def guardar_en_historial(correo, tipo, detalle):
    if correo not in st.session_state.historial_usuario:
        st.session_state.historial_usuario[correo] = []
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.historial_usuario[correo].append({
        "fecha": fecha_actual,
        "tipo": tipo,
        "detalle": detalle
    })

# --- GENERADOR DE ARCHIVO ZIP ORDENADO ---
def crear_paquete_zip(texto_extraido, lista_nombres_archivos):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        # 1. Archivo maestro con el texto general extraído
        zip_file.writestr("Reporte_Extraccion_General.txt", texto_extraido)
        
        # 2. Carpeta virtual organizada con las fichas
        for nombre in lista_nombres_archivos:
            nombre_limpio = nombre.replace('.pdf', '').replace('.xlsx', '').replace('.xls', '')
            contenido_individual = f"""EXPEDIENTE PROCESADO - LEXFLOW STUDIO
----------------------------------------
Archivo Original: {nombre}
Estado: Extraído Correctamente
Sistema Jurídico: Honduras
"""
            zip_file.writestr(f"expedientes_ordenados/{nombre_limpio}_ficha.txt", contenido_individual)
    zip_buffer.seek(0)
    return zip_buffer

# --- GENERADOR DE WORD (.DOCX) ---
def generar_word(texto):
    doc = Document()
    doc.add_paragraph(texto)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- PANEL DE ADMINISTRACIÓN ---
def panel_administracion():
    st.sidebar.markdown("---")
    st.sidebar.subheader("⚙️ Panel de Administrador")
    menu_admin = st.sidebar.selectbox("Acciones Admin", ["Gestionar Solicitudes", "Ver / Eliminar Usuarios"])

    if menu_admin == "Gestionar Solicitudes":
        st.header("👥 Gestión de Usuarios Pendientes")
        pendientes = [correo for correo, datos in st.session_state.users_db.items() if datos["status"] == "Pendiente"]
        if not pendientes:
            st.info("No hay solicitudes pendientes en este momento.")
        else:
            for correo in pendientes:
                col_a, col_b, col_c = st.columns([3, 1, 1])
                col_a.write(f"**Usuario:** {correo}")
                if col_b.button("✅ Aceptar", key=f"aceptar_{correo}"):
                    st.session_state.users_db[correo]["status"] = "Aprobado"
                    st.success(f"Usuario {correo} aprobado con éxito.")
                    st.rerun()
                if col_c.button("❌ Rechazar", key=f"rechazar_{correo}"):
                    del st.session_state.users_db[correo]
                    st.warning(f"Solicitud de {correo} rechazada.")
                    st.rerun()

    elif menu_admin == "Ver / Eliminar Usuarios":
        st.header("📋 Base de Datos de Usuarios")
        for correo, datos in list(st.session_state.users_db.items()):
            if correo != "admin@lexflow.com":
                col_1, col_2, col_3 = st.columns([3, 2, 1])
                col_1.write(f"**{correo}**")
                col_2.write(f"Estado: {datos['status']}")
                if col_3.button("🗑️ Eliminar", key=f"del_{correo}"):
                    del st.session_state.users_db[correo]
                    st.error(f"Usuario {correo} eliminado.")
                    st.rerun()

# --- CENTRO DE COMANDO PRINCIPAL ---
def centro_de_comando_principal(correo_usuario):
    st.title("⚖️ LexFlow Studio - Centro Jurídico (Modo Offline)")
    st.write(f"Bienvenido, **{correo_usuario}**")

    pestana_lote, pestana_redactor, pestana_historial = st.tabs([
        "📁 Análisis Masivo (PDFs y Excel)",
        "✍️ Redactor de Demandas (Honduras)",
        "🔒 Historial Privado"
    ])

    # --- PESTAÑA 1: CARGA MASIVA (PDFs Y EXCEL) ---
    with pestana_lote:
        st.subheader("Extracción y Organización Masiva")
        st.write("Sube archivos PDF o Excel. El sistema extraerá el texto y preparará tu carpeta organizada.")
        
        archivos_subidos = st.file_uploader(
            "Sube tus expedientes", type=["pdf", "xlsx", "xls"], accept_multiple_files=True
        )
        
        if archivos_subidos:
            st.success(f"Cargados {len(archivos_subidos)} archivos.")
            if st.button("⚡ Procesar y Generar Carpeta (ZIP)"):
                with st.spinner("Procesando documentos..."):
                    texto_consolidado = ""
                    nombres_archivos = []
                    
                    for idx, archivo in enumerate(archivos_subidos):
                        nombres_archivos.append(archivo.name)
                        
                        # Si es PDF
                        if archivo.name.endswith('.pdf'):
                            lector = PdfReader(archivo)
                            texto_doc = f"\n=== ARCHIVO PDF: {archivo.name} ===\n"
                            for i, pagina in enumerate(lector.pages):
                                if i < 4: # Lee solo las primeras 4 páginas
                                    t = pagina.extract_text()
                                    if t: texto_doc += t + "\n"
                            texto_consolidado += texto_doc
                            
                        # Si es Excel
                        elif archivo.name.endswith(('.xlsx', '.xls')):
                            df = pd.read_excel(archivo)
                            texto_doc = f"\n=== ARCHIVO EXCEL: {archivo.name} ===\n"
                            texto_doc += df.to_string() + "\n"
                            texto_consolidado += texto_doc

                    paquete_zip = crear_paquete_zip(texto_consolidado, nombres_archivos)
                    
                    st.markdown("---")
                    st.success("¡Carpeta generada con éxito!")
                    st.download_button(
                        label="📦 Descargar Carpeta Ordenada (ZIP)",
                        data=paquete_zip,
                        file_name="Expedientes_LexFlow.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                    guardar_en_historial(correo_usuario, "Extracción Masiva", f"{len(archivos_subidos)} archivos procesados.")

    # --- PESTAÑA 2: REDACTOR DE DEMANDAS ---
    with pestana_redactor:
        st.subheader("Redacción Automatizada mediante Plantillas (Honduras)")
        col1, col2 = st.columns(2)
        with col1:
            tipo = st.selectbox("Tipo de Escrito", ["Demanda Civil", "Demanda Laboral", "Medidas Cautelares"])
            actor = st.text_input("Nombre del Demandante / Actor")
            identidad_actor = st.text_input("DNI del Demandante")
        with col2:
            juzgado = st.text_input("Juzgado Competente", placeholder="Ej: Juzgado de Letras de lo Civil")
            demandado = st.text_input("Nombre del Demandado")
            cuantia = st.text_input("Cuantía (Lempiras)")
            
        hechos = st.text_area("Relación de Hechos Detallados", placeholder="Redacta los hechos...")
        
        if st.button("⚡ Generar Borrador"):
            if actor and demandado and hechos and juzgado:
                # Motor de plantillas predefinido
                borrador = f"""SE INTERPONE {tipo.upper()}.

Señor Juez del {juzgado}.

Yo, {actor}, mayor de edad, con número de identidad {identidad_actor}, actuando en mi propio nombre / a través de mi apoderado legal, ante usted con el debido respeto comparezco a interponer la presente acción en contra de {demandado}.

La cuantía de la presente demanda se estima en {cuantia} Lempiras.

HECHOS:
{hechos}

FUNDAMENTOS DE DERECHO:
Fundo la presente acción en las disposiciones aplicables del Código Procesal Civil y demás normativa vigente en la República de Honduras.

PETITORIA:
Al Juzgado respetuosamente PIDO:
1. Admitir el presente escrito junto con los documentos acompañados.
2. Darle el trámite correspondiente conforme a ley.
3. En su momento procesal oportuno, dictar sentencia favorable.

Tegucigalpa, M.D.C., {datetime.datetime.now().strftime("%d de %B de %Y")}.


_________________________________
Firma del Demandante / Apoderado Legal
"""
                st.text_area("Borrador Procesal:", value=borrador, height=400)
                
                # Descargar en Word
                archivo_word = generar_word(borrador)
                st.download_button(
                    label="📄 Descargar en Word (.docx)",
                    data=archivo_word,
                    file_name=f"{tipo.replace(' ', '_')}_{actor}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                guardar_en_historial(correo_usuario, f"Redacción {tipo}", f"Actor: {actor}")
            else:
                st.warning("Completa los campos de Actor, Demandado, Juzgado y Hechos.")

    # --- PESTAÑA 3: HISTORIAL PRIVADO ---
    with pestana_historial:
        st.subheader("🔒 Tu Historial de Actividad")
        historial_actual = st.session_state.historial_usuario.get(correo_usuario, [])
        if not historial_actual:
            st.info("No tienes registros.")
        else:
            for item in reversed(historial_actual):
                st.markdown(f"**{item['fecha']}** | `{item['tipo']}` - {item['detalle']}")
                st.markdown("---")

# --- CONTROL DE ACCESO ---
def main():
    if 'conectado' not in st.session_state:
        st.session_state['conectado'] = False
        st.session_state['user_email'] = None

    if not st.session_state['conectado']:
        st.sidebar.title("🔐 Acceso a LexFlow")
        opcion = st.sidebar.selectbox("Selecciona:", ["Iniciar Sesión", "Registro"])
        
        if opcion == "Iniciar Sesión":
            correo = st.text_input("Correo electrónico")
            clave = st.text_input("Contraseña", type="password")
            if st.button("Ingresar"):
                if correo in st.session_state.users_db:
                    datos_usr = st.session_state.users_db[correo]
                    if datos_usr["password"] == clave:
                        if datos_usr["status"] == "Aprobado":
                            st.session_state['conectado'] = True
                            st.session_state['user_email'] = correo
                            st.session_state['rol_actual'] = datos_usr["role"]
                            st.rerun()
                        else:
                            st.warning("Cuenta pendiente de aprobación.")
                    else:
                        st.error("Contraseña incorrecta.")
                else:
                    st.error("Correo no registrado.")
        else:
            nuevo_correo = st.text_input("Correo nuevo")
            nueva_clave = st.text_input("Contraseña", type="password")
            if st.button("Registrarse"):
                if nuevo_correo in st.session_state.users_db:
                    st.error("Correo ya registrado.")
                elif not nuevo_correo or not nueva_clave:
                    st.warning("Completa los campos.")
                else:
                    st.session_state.users_db[nuevo_correo] = {"password": nueva_clave, "status": "Pendiente", "role": "user"}
                    st.success("Registro exitoso. Pendiente de aprobación.")
                    
    else:
        if st.sidebar.button("Cerrar Sesión"):
            st.session_state['conectado'] = False
            st.session_state['user_email'] = None
            st.session_state['rol_actual'] = None
            st.rerun()

        if st.session_state['rol_actual'] == 'admin':
            panel_administracion()
            
        centro_de_comando_principal(st.session_state['user_email'])

if __name__ == "__main__":
    main()
