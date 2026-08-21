import streamlit as st
import io
from pypdf import PdfReader
from google import genai
import datetime

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
    st.session_state.historial_usuario = {} # {email: [registros]}

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

    # Pestañas de trabajo
    pestana_lote, pestana_redactor, pestana_historial = st.tabs([
        "📁 Análisis y Agrupación Masiva (PDFs)", 
        "✍️ Redactor de Demandas (Honduras)", 
        "🔒 Historial Privado"
    ])

    # --- PESTAÑA 1: CARGA MASIVA Y EXTRACCIÓN PRECISA ---
    with pestana_lote:
        st.subheader("Extracción y Agrupación Masiva de Expedientes")
        st.write("Sube múltiples documentos en PDF al mismo tiempo. El sistema extraerá y ordenará con máxima precisión los factores clave.")
        
        archivos_subidos = st.file_uploader("Sube los archivos PDF (puedes seleccionar varios)", type=["pdf"], accept_multiple_files=True)
        
        if archivos_subidos:
            st.success(f"Se han cargado {len(archivos_subidos)} archivos correctamente.")
            
            if st.button("⚡ Procesar y Agrupar Datos Masivos"):
                with st.spinner("Extrayendo texto y estructurando información con precisión jurídica..."):
                    texto_consolidado = ""
                    for idx, archivo in enumerate(archivos_subidos):
                        lector = PdfReader(archivo)
                        texto_doc = f"\n--- DOCUMENTO {idx+1}: {archivo.name} ---\n"
                        for pagina in lector.pages:
                            t = pagina.extract_text()
                            if t:
                                texto_doc += t + "\n"
                        texto_consolidado += texto_doc[:15000] # Limita por archivo para evitar desbordes
                    
                    prompt_masivo = f"""
                    Actúa como un sistema experto de auditoría y análisis jurídico documental.
                    Analiza los siguientes documentos y extrae con MÁXIMA PRECISIÓN los datos para cada caso.
                    
                    Debes estructurar la salida estrictamente en formato de tabla o viñetas organizadas agrupando por los siguientes factores obligatorios:
                    1. Nombre completo del sujeto o parte procesal.
                    2. Número de Identidad / DNI.
                    3. Año (del documento o emisión).
                    4. Juzgado competente mencionado.
                    
                    DOCUMENTOS A ANALIZAR:
                    {texto_consolidado[:70000]}
                    """
                    
                    resultado_analisis = procesar_con_ia(prompt_masivo)
                    st.markdown("### 📊 Resultado de la Agrupación y Extracción:")
                    st.markdown(resultado_analisis)
                    
                    # Guardar en historial privado
                    guardar_en_historial(correo_usuario, "Análisis Masivo de PDFs", f"Procesados {len(archivos_subidos)} documentos con éxito.")

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
        st.write("Aquí puedes consultar los registros de tus análisis y redacciones anteriores.")
        
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
        # Botón de cerrar sesión
        if st.sidebar.button("Cerrar Sesión"):
            st.session_state['conectado'] = False
            st.session_state['user_email'] = None
            st.session_state['rol_actual'] = None
            st.rerun()
            
        # Si es admin, mostrar panel administrativo en la barra lateral
        if st.session_state['rol_actual'] == 'admin':
            panel_administracion()
            
        # Ejecutar módulo principal
        centro_de_comando_principal(st.session_state['user_email'])

if __name__ == "__main__":
    main()