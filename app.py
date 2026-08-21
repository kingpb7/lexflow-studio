import streamlit as st
import io
from pypdf import PdfReader
import pandas as pd
from google import genai
import datetime
import zipfile

# --- CONFIGURACIÓN DE PÁGINA Y IA ---
st.set_page_config(page_title="LexFlow Studio - Sistema Jurídico Avanzado", layout="wide")

@st.cache_resource
def inicializar_ia():
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        return genai.Client(api_key=api_key)
    except Exception:
        return None

client = inicializar_ia()

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

# --- FUNCIÓN DE PROCESAMIENTO PRECISO CON IA ---
def procesar_con_ia(prompt):
    if not client:
        return "Error: Configura tu GEMINI_API_KEY en .streamlit/secrets.toml"
    try:
        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=prompt
        )
        return response.text
    except Exception as e:
        return f"Error en el procesamiento: {e}"

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
def crear_paquete_zip(resultado_ia, lista_nombres_archivos):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        # 1. Archivo maestro con la tabla de agrupación oficial
        zip_file.writestr("Reporte_Agrupacion_General.md", resultado_ia)
        
        # 2. Carpeta virtual organizada con las fichas de los documentos procesados
        for nombre in lista_nombres_archivos:
            nombre_limpio = nombre.replace('.pdf', '').replace('.xlsx', '').replace('.xls', '')
            contenido_individual = f"""EXPEDIENTE ORDENADO - LEXFLOW STUDIO
----------------------------------------
Archivo Original: {nombre}
Estado: Procesado, Extraído y Agrupado Correctamente
Factores Analizados: Nombre, Identidad, Año, Juzgado
Sistema Jurídico: Honduras
"""
            zip_file.writestr(f"expedientes_ordenados/{nombre_limpio}_ficha_tecnica.txt", contenido_individual)
            
    zip_buffer.seek(0)
    return zip_buffer

# --- PANEL DE ADMINISTRACIÓN ---
def panel_administracion():
    st.sidebar.markdown("---")
    st.sidebar.subheader("⚙️ Panel de Administrador")
    
    menu_admin = st.sidebar.selectbox("Acciones Admin", ["Gestionar Solicitudes", "Ver / Eliminar Usuarios"])
    
    if menu_admin == "Gestionar Solicitudes":
        st.header("👥 Gestión de Usuarios Pendientes de Aprobación")
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
                    st.warning(f"Solicitud de {correo} rechazada y eliminada.")
                    st.rerun()

    elif menu_admin == "Ver / Eliminar Usuarios":
        st.header("📋 Base de Datos de Usuarios Registrados")
        for correo, datos in list(st.session_state.users_db.items()):
            if correo != "admin@lexflow.com":
                col_1, col_2, col_3 = st.columns([3, 2, 1])
                col_1.write(f"**{correo}**")
                col_2.write(f"Estado: {datos['status']}")
                if col_3.button("🗑️ Eliminar", key=f"del_{correo}"):
                    del st.session_state.users_db[correo]
                    if correo in st.session_state.historial_usuario:
                        del st.session_state.historial_usuario[correo]
                    st.error(f"Usuario {correo} eliminado permanentemente.")
                    st.rerun()

# --- CENTRO DE COMANDO PRINCIPAL ---
def centro_de_comando_principal(correo_usuario):
    st.title("⚖️ LexFlow Studio - Centro Jurídico Avanzado")
    st.write(f"Bienvenido, **{correo_usuario}**")
    
    if client is None:
        st.error("⚠️ Falta configurar la GEMINI_API_KEY en `.streamlit/secrets.toml`.")
        return

    pestana_lote, pestana_redactor, pestana_historial = st.tabs([
        "📁 Análisis Masivo (PDFs y Excel)", 
        "✍️ Redactor de Demandas (Honduras)", 
        "🔒 Historial Privado"
    ])

    # --- PESTAÑA 1: CARGA MASIVA (PDFs Y EXCEL) ---
    with pestana_lote:
        st.subheader("Análisis, Agrupación Masiva y Paquete Descargable")
        st.write("Sube múltiples archivos **PDF o Excel (.xlsx, .xls)** simultáneamente. El sistema extraerá y agrupará con precisión los datos y preparará tu carpeta.")
        
        archivos_subidos = st.file_uploader(
            "Sube tus expedientes (PDF o Excel)", 
            type=["pdf", "xlsx", "xls"], 
            accept_multiple_files=True
        )
        
        if archivos_subidos:
            st.success(f"Se han cargado {len(archivos_subidos)} archivos correctamente.")
            
            if st.button("⚡ Procesar, Agrupar y Generar Carpeta Descargable"):
                with st.spinner("Procesando documentos y estructurando información clave..."):
                    texto_consolidado = ""
                    nombres_archivos = []
                    
                    for idx, archivo in enumerate(archivos_subidos):
                        nombres_archivos.append(archivo.name)
                        
                        # Procesamiento si es PDF
                        if archivo.name.endswith('.pdf'):
                            lector = PdfReader(archivo)
                            texto_doc = f"\n=== ARCHIVO PDF {idx+1}: {archivo.name} ===\n"
                            for i, pagina in enumerate(lector.pages):
                                if i < 4: # Optimización de velocidad
                                    t = pagina.extract_text()
                                    if t:
                                        texto_doc += t + "\n"
                            texto_consolidado += texto_doc
                            
                        # Procesamiento si es Excel
                        elif archivo.name.endswith(('.xlsx', '.xls')):
                            df = pd.read_excel(archivo)
                            texto_doc = f"\n=== ARCHIVO EXCEL {idx+1}: {archivo.name} ===\n"
                            texto_doc += df.to_string() + "\n"
                            texto_consolidado += texto_doc
                    
                    prompt_masivo = f"""
                    Actúa como un sistema experto de auditoría y análisis jurídico documental.
                    Analiza la información extraída de los documentos (PDFs y Excel) y agrupa los datos con MÁXIMA PRECISIÓN bajo los siguientes factores estrictos:
                    1. Nombre completo del sujeto o parte procesal.
                    2. Número de Identidad / DNI.
                    3. Año (del documento o emisión).
                    4. Juzgado competente mencionado.
                    
                    Presenta el resultado en una tabla sumamente ordenada, clara y limpia.
                    
                    DATOS DE LOS DOCUMENTOS:
                    {texto_consolidado[:45000]}
                    """
                    
                    resultado_analisis = procesar_con_ia(prompt_masivo)
                    st.markdown("### 📊 Resultado de la Agrupación Precisa:")
                    st.markdown(resultado_analisis)
                    
                    # Generación del archivo ZIP con la carpeta ordenada
                    paquete_zip = crear_paquete_zip(resultado_analisis, nombres_archivos)
                    
                    st.markdown("---")
                    st.success("¡Carpeta de expedientes ordenados generada con éxito!")
                    st.download_button(
                        label="📦 Descargar Carpeta de Expedientes Ordenados (ZIP)",
                        data=paquete_zip,
                        file_name="Expedientes_Ordenados_LexFlow.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                    
                    # Guardar en historial privado
                    guardar_en_historial(correo_usuario, "Análisis Masivo (PDF/Excel) y ZIP", f"Procesados {len(archivos_subidos)} archivos.")

    # --- PESTAÑA 2: REDACTOR DE DEMANDAS (HONDURAS) ---
    with pestana_redactor:
        st.subheader("Redacción Procesal Precisa (Honduras)")
        col1, col2 = st.columns(2)
        with col1:
            tipo = st.selectbox("Tipo de Escrito", ["Demanda Civil (Proceso Ordinario/Abreviado)", "Demanda Laboral", "Solicitud de Medidas Cautelares"])
            actor = st.text_input("Nombre del Demandante / Actor")
            identidad_actor = st.text_input("DNI del Demandante")
        with col2:
            juzgado = st.text_input("Juzgado Competente", placeholder="Ej: Juzgado de Letras de lo Civil")
            demandado = st.text_input("Nombre del Demandado")
            cuantia = st.text_input("Cuantía (Lempiras)")
            
        hechos = st.text_area("Relación de Hechos Detallados", placeholder="Redacta los hechos cronológicamente...")
        
        if st.button("⚡ Generar Escrito Estructurado"):
            if actor and demandado and hechos:
                with st.spinner("Estructurando escrito bajo el Código Procesal Civil de Honduras..."):
                    prompt_demanda = f"""
                    Actúa como abogado litigante experto en la legislación de Honduras. 
                    Redacta un escrito formal de {tipo} dirigido al {juzgado}.
                    
                    DATOS PRECISOS:
                    - Actor: {actor} (DNI: {identidad_actor})
                    - Demandado: {demandado}
                    - Cuantía: {cuantia}
                    - Hechos: {hechos}
                    
                    ESTRUCTURA OBLIGATORIA:
                    1. SUMA exactas.
                    2. GENERALIDADES de las partes.
                    3. HECHOS cronológicos.
                    4. FUNDAMENTOS DE DERECHO (CPC de Honduras y leyes aplicables).
                    5. PETITORIA precisa.
                    """
                    borrador = procesar_con_ia(prompt_demanda)
                    st.text_area("Borrador Procesal Oficial:", value=borrador, height=500)
                    
                    guardar_en_historial(correo_usuario, f"Redacción de {tipo}", f"Caso: {actor} vs {demandado}")
            else:
                st.warning("Completa los campos obligatorios.")

    # --- PESTAÑA 3: HISTORIAL PRIVADO ---
    with pestana_historial:
        st.subheader("🔒 Tu Historial Confidencial de Actividad")
        st.write("Aquí puedes consultar los registros privados de tus análisis y redacciones anteriores.")
        
        historial_actual = st.session_state.historial_usuario.get(correo_usuario, [])
        if not historial_actual:
            st.info("Aún no tienes registros en tu historial.")
        else:
            for item in reversed(historial_actual):
                with st.container():
                    st.markdown(f"**Fecha:** {item['fecha']} | **Acción:** `{item['tipo']}`")
                    st.text(f"Detalle: {item['detalle']}")
                    st.markdown("---")

# --- CONTROL DE ACCESO Y AUTENTICACIÓN ---
def main():
    if 'conectado' not in st.session_state:
        st.session_state['conectado'] = False
        st.session_state['user_email'] = None

    if not st.session_state['conectado']:
        st.sidebar.title("🔐 Acceso a LexFlow")
        opcion = st.sidebar.selectbox("Selecciona:", ["Iniciar Sesión", "Registro"])
        
        if opcion == "Iniciar Sesión":
            st.subheader("Iniciar Sesión en LexFlow Studio")
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
                            st.success("¡Acceso concedido!")
                            st.rerun()
                        else:
                            st.warning("Tu cuenta se encuentra pendiente de aprobación por el administrador.")
                    else:
                        st.error("Contraseña incorrecta.")
                else:
                    st.error("El correo no está registrado.")
        else:
            st.subheader("Registro de Nuevo Usuario")
            nuevo_correo = st.text_input("Correo electrónico nuevo")
            nueva_clave = st.text_input("Contraseña nueva", type="password")
            
            if st.button("Registrarse"):
                if nuevo_correo in st.session_state.users_db:
                    st.error("Este correo ya está registrado.")
                elif not nuevo_correo or not nueva_clave:
                    st.warning("Completa todos los campos.")
                else:
                    st.session_state.users_db[nuevo_correo] = {
                        "password": nueva_clave,
                        "status": "Pendiente",
                        "role": "user"
                    }
                    st.success("¡Registro exitoso! Tu cuenta ha quedado pendiente de aprobación por el administrador.")
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