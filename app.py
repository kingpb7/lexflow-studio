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

# --- MOTOR LOCAL DE EXTRACCIÓN (BASADO EN TU FORMATO DE WORD) ---
def extraer_datos_precisos(texto):
    texto_limpio = texto.replace('\n', ' ')
    
    # 1. Identidad (13 dígitos)
    id_match = re.search(r'\b\d{4}[-\s]?\d{4}[-\s]?\d{5}\b', texto_limpio)
    identidad = id_match.group(0).strip() if id_match else "Sin_Identidad"
    
    # 2. Nombre (Busca "Yo, [NOMBRE], mayor de edad")
    nombre = "Sin_Nombre"
    nombre_match = re.search(r'Yo,\s*([A-ZÁÉÍÓÚÑa-záéíóúñ\s]+?),\s*mayor\s+de\s+edad', texto_limpio, re.IGNORECASE)
    if nombre_match:
        nombre = nombre_match.group(1).strip()
    else:
        nombre_firma = re.search(r'Firma[:\s_]+([A-ZÁÉÍÓÚÑ\s]+)', texto_limpio, re.IGNORECASE)
        if nombre_firma:
            nombre = nombre_firma.group(1).strip()

    # 3. Juzgado (Busca "SEÑOR JUEZ..." o "JUZGADO...")
    juzgado = "Juzgado_General"
    juzgado_match = re.search(r'(SEÑOR JUEZ[A-ZÁÉÍÓÚÑa-záéíóúñ\s]+|JUZGADO[A-ZÁÉÍÓÚÑa-záéíóúñ\s]+)(?=\.|,)', texto_limpio, re.IGNORECASE)
    if juzgado_match:
        juzgado = juzgado_match.group(0).strip()
        if len(juzgado) > 60:
            juzgado = juzgado[:60]
            
    # 4. Año
    anio = "Sin_Anio"
    anios = re.findall(r'\b(20\d{2})\b', texto_limpio)
    if anios:
        anio = anios[-1] 
        
    # Limpiar caracteres especiales para nombres de archivos seguros en Windows/Linux
    nombre_seguro = re.sub(r'[\\/*?:"<>|]', "", nombre)
    juzgado_seguro = re.sub(r'[\\/*?:"<>|]', "", juzgado)
    
    return {
        "Nombre": nombre, 
        "Identidad": identidad, 
        "Año": anio, 
        "Juzgado": juzgado,
        "Nombre_Archivo": f"{nombre_seguro} - {juzgado_seguro} - {identidad} - {anio}"
    }

# --- REDACTOR BASADO EN LA MATRIZ DE TU WORD ---
def generar_documento_word_demanda(tipo, actor, identidad_actor, juzgado, demandado, cuantia, hechos, abogado):
    doc = docx.Document()
    
    # Encabezado / Suma
    doc.add_paragraph(f"SEÑOR JUEZ DE {juzgado.upper()}.")
    doc.add_paragraph(f"SE INTERPONE {tipo.upper()}. - SE ACOMPAÑAN DOCUMENTOS.- SE OTORGA PODER.")
    doc.add_paragraph()
    
    # Cuerpo Principal con tu misma estructura formal
    cuerpo = (
        f"Yo, {actor}, mayor de edad, con Documento Nacional de Identificación (DNI) No. {identidad_actor}, "
        f"actuando en mi propia condición; con el debido respeto comparezco ante usted, Señor Juez, interponiendo "
        f"{tipo.upper()} en contra del señor(a) {demandado}. La presente demanda se basa en los siguientes hechos y fundamentos de derecho:"
    )
    doc.add_paragraph(cuerpo)
    
    # Hechos
    doc.add_heading("HECHOS", level=2)
    doc.add_paragraph(f"PRIMERO: {hechos}")
    doc.add_paragraph("SEGUNDO: Es el caso, Señor Juez, que la parte demandada ha incumplido con sus obligaciones, motivando la presente acción judicial.")
    doc.add_paragraph("TERCERO: Ante la negativa de honrar la obligación contraída, me veo en la necesidad de recurrir a este órgano jurisdiccional.")
    
    # Medios de Prueba
    doc.add_heading("MEDIOS DE PRUEBA", level=2)
    doc.add_paragraph("1. DOCUMENTAL PRIVADA: Consistente en los documentos acompañados al presente escrito.")
    doc.add_paragraph("2. INTERROGATORIO DE LAS PARTES.")
    
    # Fundamentos de Derecho
    doc.add_heading("FUNDAMENTOS DE DERECHO", level=2)
    doc.add_paragraph("Fundo la presente demanda en los Artículos constitucionales aplicables y las disposiciones civiles y procesales pertinentes sobre obligaciones y contratos.")
    
    # Otorgamiento de Poder
    if abogado:
        doc.add_heading("OTORGAMIENTO DE PODER", level=2)
        doc.add_paragraph(f"Para que me represente, otorgo Poder Especial al Abogado(a) {abogado}, con inscripción profesional vigente, investido(a) de las facultades generales y especiales del mandato judicial.")
    
    # Petición
    doc.add_heading("PETICIÓN", level=2)
    doc.add_paragraph("A usted, Señor Juez, respetuosamente PIDO:\n1. Admitir la presente Demanda junto con los documentos acompañados.\n2. Tener por acreditado el Poder otorgado en su caso.\n3. Emplazar a la parte demandada.\n4. En sentencia definitiva, declarar CON LUGAR la demanda.")
    
    fecha_hoy = datetime.datetime.now().strftime("Tegucigalpa, M.D.C., a los %d días del mes de %B de %Y.")
    doc.add_paragraph(fecha_hoy)
    doc.add_paragraph(f"___________________________________________\n{actor}\nDNI: {identidad_actor}")
    
    # Guardar en memoria BytesIO
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- HISTORIAL Y ZIP ---
def guardar_en_historial(correo, tipo, detalle):
    if correo not in st.session_state.historial_usuario:
        st.session_state.historial_usuario[correo] = []
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.historial_usuario[correo].append({"fecha": fecha_actual, "tipo": tipo, "detalle": detalle})

def crear_paquete_zip_con_words(dataframe_resultados, lista_fichas_datos):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        # 1. Excel Maestro
        excel_buffer = io.BytesIO()
        dataframe_resultados.to_excel(excel_buffer, index=False)
        zip_file.writestr("Reporte_General_Agrupacion.xlsx", excel_buffer.getvalue())
        
        # 2. Generar un documento Word ordenado por cada expediente analizado
        for item in lista_fichas_datos:
            doc_ind = docx.Document()
            doc_ind.add_heading("EXPEDIENTE ORDENADO - LEXFLOW STUDIO", level=1)
            doc_ind.add_paragraph(f"Archivo Original Analizado: {item['archivo']}")
            doc_ind.add_heading("Datos Extraídos:", level=2)
            doc_ind.add_paragraph(f"• Nombre: {item['datos']['Nombre']}")
            doc_ind.add_paragraph(f"• Juzgado: {item['datos']['Juzgado']}")
            doc_ind.add_paragraph(f"• Número de Identidad: {item['datos']['Identidad']}")
            doc_ind.add_paragraph(f"• Año: {item['datos']['Año']}")
            
            doc_bytes = io.BytesIO()
            doc_ind.save(doc_bytes)
            doc_bytes.seek(0)
            
            nombre_archivo_zip = f"expedientes_ordenados/{item['datos']['Nombre_Archivo']}.docx"
            zip_file.writestr(nombre_archivo_zip, doc_bytes.getvalue())
            
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
            st.info("No hay solicitudes pendientes.")
        else:
            for correo in pendientes:
                col_a, col_b, col_c = st.columns([3, 1, 1])
                col_a.write(f"**Usuario:** {correo}")
                if col_b.button("✅ Aceptar", key=f"aceptar_{correo}"):
                    st.session_state.users_db[correo]["status"] = "Aprobado"
                    st.success("Aprobado.")
                    st.rerun()
                if col_c.button("❌ Rechazar", key=f"rechazar_{correo}"):
                    del st.session_state.users_db[correo]
                    st.warning("Rechazado.")
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
                    st.error("Eliminado.")
                    st.rerun()

# --- CENTRO DE COMANDO PRINCIPAL ---
def centro_de_comando_principal(correo_usuario):
    st.title("⚖️ LexFlow Studio - Sistema Jurídico")
    st.write(f"Bienvenido, **{correo_usuario}**")

    pestana_lote, pestana_redactor, pestana_historial = st.tabs([
        "📁 Agrupación y Carpeta Word Ordenada", 
        "✍️ Redactor Procesal Matriz", 
        "🔒 Historial Privado"
    ])

    # --- PESTAÑA 1: CARGA MASIVA Y WORD ORDENADOS ---
    with pestana_lote:
        st.subheader("Agrupación Masiva y Generación de Documentos Word Ordenados")
        st.info("Sube tus expedientes. El sistema generará una carpeta ZIP con un reporte Excel y un archivo Word por cada caso, nombrados bajo el formato: [Nombre] - [Juzgado] - [Identidad] - [Año].")
        
        archivos_subidos = st.file_uploader(
            "Sube tus archivos (Word, PDF o Excel)", 
            type=["pdf", "xlsx", "xls", "docx"], 
            accept_multiple_files=True
        )
        
        if archivos_subidos:
            if st.button("⚡ Procesar, Ordenar y Crear Carpeta ZIP"):
                with st.spinner("Generando documentos ordenados a alta velocidad..."):
                    resultados_tabla = []
                    lista_fichas = []
                    
                    for archivo in archivos_subidos:
                        texto_doc = ""
                        if archivo.name.endswith('.pdf'):
                            lector = PdfReader(archivo)
                            for i, pagina in enumerate(lector.pages):
                                if i < 3:
                                    t = pagina.extract_text()
                                    if t: texto_doc += t + " "
                        elif archivo.name.endswith('.docx'):
                            doc = docx.Document(archivo)
                            texto_doc = " ".join([para.text for para in doc.paragraphs])
                        elif archivo.name.endswith(('.xlsx', '.xls')):
                            df = pd.read_excel(archivo)
                            texto_doc = " ".join(df.astype(str).values.flatten())
                        
                        datos_extraidos = extraer_datos_precisos(texto_doc)
                        
                        fila = {"Archivo Original": archivo.name}
                        fila.update({
                            "Nombre": datos_extraidos["Nombre"],
                            "Juzgado": datos_extraidos["Juzgado"],
                            "Identidad": datos_extraidos["Identidad"],
                            "Año": datos_extraidos["Año"]
                        })
                        resultados_tabla.append(fila)
                        lista_fichas.append({"archivo": archivo.name, "datos": datos_extraidos})
                    
                    df_resultados = pd.DataFrame(resultados_tabla)
                    st.markdown("### 📊 Tabla de Datos Agrupados:")
                    st.dataframe(df_resultados, use_container_width=True)
                    
                    # Paquete ZIP con archivos Word formateados
                    paquete_zip = crear_paquete_zip_con_words(df_resultados, lista_fichas)
                    
                    st.markdown("---")
                    st.success("¡Carpeta comprimida generada con éxito!")
                    st.download_button(
                        label="📦 Descargar Carpeta ZIP con Word y Excel Ordenados",
                        data=paquete_zip,
                        file_name="Expedientes_Ordenados_LexFlow.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                    guardar_en_historial(correo_usuario, "Agrupación Masiva con Word", f"Procesados {len(archivos_subidos)} archivos.")

    # --- PESTAÑA 2: REDACTOR MATRIZ ---
    with pestana_redactor:
        st.subheader("Redacción de Escritos basada en Estructura Matriz")
        st.info("Crea un documento formal respetando exactamente la estructura institucional del formato base.")
        
        col1, col2 = st.columns(2)
        with col1:
            tipo = st.selectbox("Tipo de Escrito", ["Demanda Ordinaria Civil", "Demanda Laboral", "Solicitud de Medidas Cautelares"])
            actor = st.text_input("Nombre del Demandante / Actor")
            identidad_actor = st.text_input("DNI del Demandante")
            abogado = st.text_input("Nombre del Abogado(a) Apoderado")
        with col2:
            juzgado = st.text_input("Juzgado Competente", placeholder="Ej: de Letras Civil de Francisco Morazán")
            demandado = st.text_input("Nombre del Demandado")
            cuantia = st.text_input("Cuantía Pretendida")
            
        hechos = st.text_area("Relación de Hechos Principal", placeholder="Redacta el hecho primero o descripción central...")
        
        if st.button("⚡ Generar y Descargar Documento Word"):
            if actor and demandado and hechos and juzgado:
                word_buffer = generar_documento_word_demanda(tipo, actor, identidad_actor, juzgado, demandado, cuantia, hechos, abogado)
                
                st.success("¡Escrito generado correctamente bajo el formato matriz!")
                st.download_button(
                    label="📥 Descargar Escrito en Formato Word (.docx)",
                    data=word_buffer,
                    file_name=f"{tipo.replace(' ', '_')}_{actor}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                guardar_en_historial(correo_usuario, f"Redacción Matriz {tipo}", f"{actor} vs {demandado}")
            else:
                st.warning("Completa los campos principales (Demandante, Demandado, Juzgado y Hechos).")

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
