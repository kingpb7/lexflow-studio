import streamlit as st
import io
import re
from pypdf import PdfReader
import pandas as pd
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

# --- MOTOR LOCAL DE EXTRACCIÓN Y AGRUPACIÓN (SIN IA, SÚPER RÁPIDO) ---
def extraer_datos_precisos(texto):
    texto_limpio = texto.replace('\n', ' ')
    
    # 1. Identidad Hondureña (4 dígitos - 4 dígitos - 5 dígitos, con o sin guiones)
    id_match = re.search(r'\b\d{4}[-]?\d{4}[-]?\d{5}\b', texto_limpio)
    identidad = id_match.group(0) if id_match else "No detectada"
    
    # 2. Año (Buscando 19xx o 20xx)
    anio_match = re.search(r'\b(19|20)\d{2}\b', texto_limpio)
    anio = anio_match.group(0) if anio_match else "No detectado"
    
    # 3. Juzgado (Busca frases que comiencen con Juzgado o Tribunal)
    juzgado_match = re.search(r'(Juzgado\s+de\s+[a-zA-Z\sáéíóúÁÉÍÓÚ]+|Tribunal\s+[a-zA-Z\sáéíóúÁÉÍÓÚ]+)', texto_limpio, re.IGNORECASE)
    juzgado = juzgado_match.group(0).strip() if juzgado_match else "No detectado"
    
    # 4. Nombre (Intenta buscar nombres propios después de palabras clave comunes en demandas)
    nombre = "No detectado"
    patrones_nombre = [
        r'(?:Demandante|Actor|Señor[a]?|Abogado[a]?|Nombre)[:\s]+([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,3})'
    ]
    for patron in patrones_nombre:
        match = re.search(patron, texto_limpio)
        if match:
            nombre = match.group(1).strip()
            break
            
    return {"Nombre": nombre, "Identidad": identidad, "Año": anio, "Juzgado": juzgado}

# --- REDACTOR DE DEMANDAS LOCAL (PLANTILLAS CPC HONDURAS) ---
def generar_plantilla_demanda(tipo, actor, identidad_actor, juzgado, demandado, cuantia, hechos):
    fecha_hoy = datetime.datetime.now().strftime("%d de %B de %Y")
    plantilla = f"""SUMA: SE INTERPONE {tipo.upper()}. SE ACOMPAÑAN DOCUMENTOS. PETICIÓN.

SEÑOR JUEZ DEL {juzgado.upper()}

Yo, {actor.upper()}, mayor de edad, con Documento Nacional de Identificación (DNI) No. {identidad_actor}, actuando en mi propia condición; con el debido respeto comparezco ante usted interponiendo {tipo.upper()} en contra del (la) señor(a) {demandado.upper()}.

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
1. Admitir el presente escrito junto con los documentos acompañados.
2. Darle el trámite correspondiente conforme a ley.
3. En definitiva, dictar sentencia favorable a mis intereses.

Tegucigalpa, M.D.C., a {fecha_hoy}.

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
    st.session_state.historial_usuario[correo].append({
        "fecha": fecha_actual,
        "tipo": tipo,
        "detalle": detalle
    })

# --- GENERADOR DE ARCHIVO ZIP ORDENADO ---
def crear_paquete_zip(dataframe_resultados, lista_fichas):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        # 1. Archivo maestro Excel con la tabla de agrupación
        excel_buffer = io.BytesIO()
        dataframe_resultados.to_excel(excel_buffer, index=False)
        zip_file.writestr("Reporte_General_Agrupacion.xlsx", excel_buffer.getvalue())
        
        # 2. Carpeta virtual organizada con las fichas de los documentos
        for ficha in lista_fichas:
            nombre_limpio = ficha['archivo'].replace('.pdf', '').replace('.xlsx', '').replace('.xls', '')
            contenido_individual = f"""EXPEDIENTE ORDENADO - LEXFLOW STUDIO
----------------------------------------
Archivo Original: {ficha['archivo']}
Estado: Procesado Localmente (Sin IA)

--- DATOS EXTRAÍDOS ---
Nombre Detectado: {ficha['datos']['Nombre']}
Identidad / DNI: {ficha['datos']['Identidad']}
Año: {ficha['datos']['Año']}
Juzgado Competente: {ficha['datos']['Juzgado']}
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
                    st.success(f"Usuario {correo} aprobado.")
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
    st.title("⚖️ LexFlow Studio - Centro Jurídico (Modo Local Súper Rápido)")
    st.write(f"Bienvenido, **{correo_usuario}**")

    pestana_lote, pestana_redactor, pestana_historial = st.tabs([
        "📁 Agrupación Masiva Rápida (PDF/Excel)", 
        "✍️ Redactor de Demandas (Honduras)", 
        "🔒 Historial Privado"
    ])

    # --- PESTAÑA 1: CARGA MASIVA Y EXTRACCIÓN LOCAL ---
    with pestana_lote:
        st.subheader("Análisis, Agrupación Masiva y Paquete Descargable")
        st.info("⚡ Este proceso se ejecuta localmente en milisegundos. No usa internet ni cuotas de IA.")
        
        archivos_subidos = st.file_uploader(
            "Sube tus expedientes (PDF o Excel)", 
            type=["pdf", "xlsx", "xls"], 
            accept_multiple_files=True
        )
        
        if archivos_subidos:
            if st.button("⚡ Procesar, Agrupar y Generar Carpeta Descargable"):
                with st.spinner("Procesando documentos a alta velocidad..."):
                    resultados_tabla = []
                    lista_fichas = []
                    
                    for archivo in archivos_subidos:
                        texto_doc = ""
                        
                        # Extraer texto según formato
                        if archivo.name.endswith('.pdf'):
                            lector = PdfReader(archivo)
                            # Solo leemos las primeras 3 páginas para máxima velocidad
                            for i, pagina in enumerate(lector.pages):
                                if i < 3: 
                                    t = pagina.extract_text()
                                    if t:
                                        texto_doc += t + " "
                        elif archivo.name.endswith(('.xlsx', '.xls')):
                            df = pd.read_excel(archivo)
                            texto_doc = " ".join(df.astype(str).values.flatten())
                        
                        # Extraer datos exactos con el algoritmo local
                        datos_extraidos = extraer_datos_precisos(texto_doc)
                        
                        # Guardar para la tabla y la carpeta ZIP
                        fila = {"Archivo": archivo.name}
                        fila.update(datos_extraidos)
                        resultados_tabla.append(fila)
                        
                        lista_fichas.append({
                            "archivo": archivo.name,
                            "datos": datos_extraidos
                        })
                    
                    # Mostrar resultados en pantalla
                    df_resultados = pd.DataFrame(resultados_tabla)
                    st.markdown("### 📊 Resultado de la Agrupación:")
                    st.dataframe(df_resultados, use_container_width=True)
                    
                    # Crear el archivo ZIP para descargar
                    paquete_zip = crear_paquete_zip(df_resultados, lista_fichas)
                    
                    st.markdown("---")
                    st.success("¡Carpeta de expedientes ordenados generada con éxito!")
                    st.download_button(
                        label="📦 Descargar Carpeta de Expedientes Ordenados (ZIP)",
                        data=paquete_zip,
                        file_name="Expedientes_Ordenados_LexFlow.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                    
                    guardar_en_historial(correo_usuario, "Agrupación Masiva (Local)", f"Procesados {len(archivos_subidos)} archivos.")

    # --- PESTAÑA 2: REDACTOR DE DEMANDAS ---
    with pestana_redactor:
        st.subheader("Redacción Procesal Estructurada (Honduras)")
        st.info("Generación instantánea utilizando formato estándar del Código Procesal Civil.")
        
        col1, col2 = st.columns(2)
        with col1:
            tipo = st.selectbox("Tipo de Escrito", ["Demanda Civil (Proceso Ordinario)", "Demanda Laboral", "Solicitud de Medidas Cautelares"])
            actor = st.text_input("Nombre del Demandante / Actor")
            identidad_actor = st.text_input("DNI del Demandante")
        with col2:
            juzgado = st.text_input("Juzgado Competente", placeholder="Ej: Juzgado de Letras de lo Civil de Francisco Morazán")
            demandado = st.text_input("Nombre del Demandado")
            cuantia = st.text_input("Cuantía (Lempiras)")
            
        hechos = st.text_area("Relación de Hechos Detallados", placeholder="Redacta los hechos cronológicamente...")
        
        if st.button("⚡ Generar Escrito Estructurado"):
            if actor and demandado and hechos and juzgado:
                borrador = generar_plantilla_demanda(tipo, actor, identidad_actor, juzgado, demandado, cuantia, hechos)
                st.text_area("Borrador Procesal Oficial:", value=borrador, height=550)
                guardar_en_historial(correo_usuario, f"Redacción de {tipo}", f"Caso: {actor} vs {demandado}")
            else:
                st.warning("Completa los campos principales (Demandante, Demandado, Juzgado y Hechos).")

    # --- PESTAÑA 3: HISTORIAL PRIVADO ---
    with pestana_historial:
        st.subheader("🔒 Tu Historial Confidencial")
        historial_actual = st.session_state.historial_usuario.get(correo_usuario, [])
        if not historial_actual:
            st.info("Aún no tienes registros en tu historial.")
        else:
            for item in reversed(historial_actual):
                with st.container():
                    st.markdown(f"**Fecha:** {item['fecha']} | **Acción:** `{item['tipo']}`")
                    st.text(f"Detalle: {item['detalle']}")
                    st.markdown("---")

# --- CONTROL DE ACCESO (LOGIN) ---
def main():
    if 'conectado' not in st.session_state:
        st.session_state['conectado'] = False
        st.session_state['user_email'] = None

    if not st.session_state['conectado']:
        st.sidebar.title("🔐 Acceso a LexFlow")
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
                            st.success("¡Acceso concedido!")
                            st.rerun()
                        else:
                            st.warning("Tu cuenta está pendiente de aprobación por el administrador.")
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
                    st.session_state.users_db[nuevo_correo] = {"password": nueva_clave, "status": "Pendiente", "role": "user"}
                    st.success("¡Registro exitoso! Cuenta pendiente de aprobación.")
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
