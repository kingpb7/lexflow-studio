import streamlit as st
import io
import re
from pypdf import PdfReader
import pandas as pd
import docx
import datetime
import zipfile

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="LexFlow Studio - Sistema Jurídico Autónomo", layout="wide")

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

# --- MOTOR LOCAL DE EXTRACCIÓN (ALGORITMO PARA FORMATO HONDUREÑO) ---
def extraer_datos_precisos(texto):
    # Usamos el texto original y una versión limpia para diferentes búsquedas
    texto_limpio = texto.replace('\n', ' ')
    
    # 1. Identidad Hondureña (13 dígitos)
    id_match = re.search(r'\b\d{4}[-\s]?\d{4}[-\s]?\d{5}\b', texto_limpio)
    identidad = id_match.group(0).strip() if id_match else "No detectada"
    
    # 2. Nombre (Busca la estructura clásica: "Yo, [NOMBRE], mayor de edad")
    nombre = "No detectado"
    nombre_match = re.search(r'Yo,\s*([A-ZÁÉÍÓÚÑa-záéíóúñ\s]+?),\s*mayor\s+de\s+edad', texto_limpio, re.IGNORECASE)
    if nombre_match:
        nombre = nombre_match.group(1).strip()
    else:
        # Respaldo por si el nombre está al final (Firma)
        nombre_firma = re.search(r'Firma[:\s_]+([A-ZÁÉÍÓÚÑ\s]+)', texto_limpio, re.IGNORECASE)
        if nombre_firma:
            nombre = nombre_firma.group(1).strip()

    # 3. Juzgado (Busca "SEÑOR JUEZ DE..." o "JUZGADO...")
    juzgado = "No detectado"
    juzgado_match = re.search(r'(SEÑOR JUEZ[A-ZÁÉÍÓÚÑa-záéíóúñ\s]+|JUZGADO[A-ZÁÉÍÓÚÑa-záéíóúñ\s]+)(?=\.|,)', texto_limpio, re.IGNORECASE)
    if juzgado_match:
        juzgado = juzgado_match.group(0).strip()
        # Cortamos si captura demasiado texto por error
        if len(juzgado) > 90:
            juzgado = juzgado[:90] + "..."
            
    # 4. Año (Buscamos el último año mencionado, suele ser el de la firma/presentación)
    anio = "No detectado"
    anios = re.findall(r'\b(20\d{2})\b', texto_limpio)
    if anios:
        anio = anios[-1] 
        
    return {"Nombre": nombre, "Identidad": identidad, "Año": anio, "Juzgado": juzgado}

# --- REDACTOR DE DEMANDAS LOCAL (PLANTILLAS CPC HONDURAS) ---
def generar_plantilla_demanda(tipo, actor, identidad_actor, juzgado, demandado, cuantia, hechos):
    fecha_hoy = datetime.datetime.now().strftime("%d de %B de %Y")
    plantilla = f"""SUMA: SE INTERPONE {tipo.upper()}. SE ACOMPAÑAN DOCUMENTOS. PETICIÓN.

SEÑOR JUEZ DEL {juzgado.upper()}

Yo, {actor.upper()}, mayor de edad, hondureño(a), con Documento Nacional de Identificación (DNI) No. {identidad_actor}, actuando en mi propia condición; con el debido respeto comparezco ante usted interponiendo {tipo.upper()} en contra del señor(a) {demandado.upper()}.

I. GENERALIDADES DE LAS PARTES
Parte Demandante: {actor}, con DNI {identidad_actor}.
Parte Demandada: {demandado}.
Cuantía de la pretensión: {cuantia} Lempiras.

II. RELACIÓN DE HECHOS (CRONOLÓGICOS)
{hechos}

III. FUNDAMENTOS DE DERECHO
Fundo la presente acción en lo establecido en el Código Procesal Civil de Honduras, garantizando el derecho a la tutela judicial efectiva y el debido proceso.

IV. PETICIÓN
Al señor Juez, respetuosamente PIDO:
1. Admitir el presente escrito.
2. Darle el trámite correspondiente conforme a ley.
3. En definitiva, dictar sentencia favorable a mis intereses.

Tegucigalpa, M.D.C., a los {fecha_hoy}.

___________________________
Firma: {actor}
DNI: {identidad_actor}
"""
    return plantilla

# --- REGISTRO EN HISTORIAL PRIVADO ---
def guardar_en_historial(correo, tipo, detalle):
    if correo not in st.session_state.historial_usuario:
        st.session_state.historial_usuario[correo] = []
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.historial_usuario[correo].append({"fecha": fecha_actual, "tipo": tipo, "detalle": detalle})

# --- GENERADOR DE ARCHIVO ZIP ORDENADO ---
def crear_paquete_zip(dataframe_resultados, lista_fichas):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        # 1. Archivo maestro Excel con la tabla de agrupación
        excel_buffer = io.BytesIO()
        dataframe_resultados.to_excel(excel_buffer, index=False)
        zip_file.writestr("Reporte_General_Agrupacion.xlsx", excel_buffer.getvalue())
        
        # 2. Carpeta virtual organizada con las fichas
        for ficha in lista_fichas:
            nombre_limpio = ficha['archivo'].replace('.pdf', '').replace('.xlsx', '').replace('.docx', '')
            contenido = f"""EXPEDIENTE ORDENADO - LEXFLOW STUDIO
----------------------------------------
Archivo Original: {ficha['archivo']}
Estado: Procesado Localmente

--- DATOS EXTRAÍDOS ---
Juzgado Competente: {ficha['datos']['Juzgado']}
Nombre Detectado: {ficha['datos']['Nombre']}
Identidad / DNI: {ficha['datos']['Identidad']}
Año: {ficha['datos']['Año']}
"""
            zip_file.writestr(f"expedientes_ordenados/{nombre_limpio}_ficha.txt", contenido)
            
    zip_buffer.seek(0)
    return zip_buffer

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
                    st.success(f"Usuario aprobado.")
                    st.rerun()
                if col_c.button("❌ Rechazar", key=f"rechazar_{correo}"):
                    del st.session_state.users_db[correo]
                    st.warning(f"Solicitud rechazada.")
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
                    if correo in st.session_state.historial_usuario:
                        del st.session_state.historial_usuario[correo]
                    st.error(f"Usuario eliminado.")
                    st.rerun()

# --- CENTRO DE COMANDO PRINCIPAL ---
def centro_de_comando_principal(correo_usuario):
    st.title("⚖️ LexFlow Studio - Centro Jurídico (Modo Rápido)")
    st.write(f"Bienvenido, **{correo_usuario}**")

    pestana_lote, pestana_redactor, pestana_historial = st.tabs([
        "📁 Agrupación Masiva (Word/PDF/Excel)", 
        "✍️ Redactor de Demandas (Honduras)", 
        "🔒 Historial Privado"
    ])

    # --- PESTAÑA 1: CARGA MASIVA ---
    with pestana_lote:
        st.subheader("Análisis y Extracción Inmediata")
        archivos_subidos = st.file_uploader(
            "Sube tus expedientes (Word, PDF o Excel)", 
            type=["pdf", "xlsx", "xls", "docx"], 
            accept_multiple_files=True
        )
        
        if archivos_subidos:
            if st.button("⚡ Agrupar y Generar Carpeta Descargable"):
                with st.spinner("Procesando documentos..."):
                    resultados_tabla = []
                    lista_fichas = []
                    
                    for archivo in archivos_subidos:
                        texto_doc = ""
                        
                        # Lectura de PDF
                        if archivo.name.endswith('.pdf'):
                            lector = PdfReader(archivo)
                            for i, pagina in enumerate(lector.pages):
                                if i < 3: 
                                    t = pagina.extract_text()
                                    if t: texto_doc += t + " "
                                    
                        # Lectura de Word
                        elif archivo.name.endswith('.docx'):
                            doc = docx.Document(archivo)
                            texto_doc = " ".join([para.text for para in doc.paragraphs])
                            
                        # Lectura de Excel
                        elif archivo.name.endswith(('.xlsx', '.xls')):
                            df = pd.read_excel(archivo)
                            texto_doc = " ".join(df.astype(str).values.flatten())
                        
                        # Extraer datos exactos
                        datos_extraidos = extraer_datos_precisos(texto_doc)
                        
                        fila = {"Archivo": archivo.name}
                        fila.update(datos_extraidos)
                        resultados_tabla.append(fila)
                        
                        lista_fichas.append({"archivo": archivo.name, "datos": datos_extraidos})
                    
                    # Mostrar tabla
                    df_resultados = pd.DataFrame(resultados_tabla)
                    st.markdown("### 📊 Resultado de la Agrupación:")
                    st.dataframe(df_resultados, use_container_width=True)
                    
                    # Crear ZIP
                    paquete_zip = crear_paquete_zip(df_resultados, lista_fichas)
                    
                    st.markdown("---")
                    st.success("¡Carpeta de expedientes ordenados lista!")
                    st.download_button(
                        label="📦 Descargar Carpeta ZIP con Excel y Expedientes",
                        data=paquete_zip,
                        file_name="Expedientes_Ordenados_LexFlow.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                    guardar_en_historial(correo_usuario, "Agrupación Masiva", f"Procesados {len(archivos_subidos)} archivos.")

    # --- PESTAÑA 2: REDACTOR DE DEMANDAS ---
    with pestana_redactor:
        st.subheader("Redacción Procesal Estructurada")
        col1, col2 = st.columns(2)
        with col1:
            tipo = st.selectbox("Tipo de Escrito", ["Demanda Civil (Proceso Ordinario)", "Demanda Laboral", "Solicitud de Medidas"])
            actor = st.text_input("Nombre del Demandante / Actor")
            identidad_actor = st.text_input("DNI del Demandante")
        with col2:
            juzgado = st.text_input("Juzgado Competente", placeholder="Ej: Juzgado de Letras de lo Civil...")
            demandado = st.text_input("Nombre del Demandado")
            cuantia = st.text_input("Cuantía (Lempiras)")
            
        hechos = st.text_area("Relación de Hechos", placeholder="Redacta cronológicamente...")
        
        if st.button("⚡ Generar Escrito"):
            if actor and demandado and hechos and juzgado:
                borrador = generar_plantilla_demanda(tipo, actor, identidad_actor, juzgado, demandado, cuantia, hechos)
                st.text_area("Borrador Procesal Oficial:", value=borrador, height=550)
                guardar_en_historial(correo_usuario, f"Redactó {tipo}", f"{actor} vs {demandado}")
            else:
                st.warning("Completa Demandante, Demandado, Juzgado y Hechos.")

    # --- PESTAÑA 3: HISTORIAL ---
    with pestana_historial:
        st.subheader("🔒 Tu Historial Confidencial")
        historial_actual = st.session_state.historial_usuario.get(correo_usuario, [])
        if not historial_actual:
            st.info("Aún no tienes registros.")
        else:
            for item in reversed(historial_actual):
                st.markdown(f"**{item['fecha']}** | Acción: `{item['tipo']}` - {item['detalle']}")
                st.markdown("---")

# --- CONTROL DE ACCESO ---
def main():
    if 'conectado' not in st.session_state:
        st.session_state['conectado'] = False
        st.session_state['user_email'] = None

    if not st.session_state['conectado']:
        st.sidebar.title("🔐 Acceso")
        opcion = st.sidebar.selectbox("Selecciona:", ["Iniciar Sesión", "Registro"])
        
        if opcion == "Iniciar Sesión":
            st.subheader("Iniciar Sesión")
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
            st.subheader("Registro")
            nuevo_correo = st.text_input("Correo nuevo")
            nueva_clave = st.text_input("Contraseña nueva", type="password")
            if st.button("Registrarse"):
                if nuevo_correo in st.session_state.users_db:
                    st.error("El correo ya existe.")
                elif not nuevo_correo or not nueva_clave:
                    st.warning("Completa los campos.")
                else:
                    st.session_state.users_db[nuevo_correo] = {"password": nueva_clave, "status": "Pendiente", "role": "user"}
                    st.success("¡Registro exitoso! Cuenta pendiente de aprobación.")
    else:
        if st.sidebar.button("Cerrar Sesión"):
            st.session_state['conectado'] = False
            st.rerun()
        if st.session_state['rol_actual'] == 'admin':
            panel_administracion()
            
        centro_de_comando_principal(st.session_state['user_email'])

if __name__ == "__main__":
    main()
