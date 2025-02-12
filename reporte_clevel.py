import streamlit as st
import pandas as pd
import io

# Función que procesa el archivo subido y retorna el Excel resultado en un BytesIO
def process_file(uploaded_file):
    # Leer el archivo Excel desde el objeto subido, utilizando la hoja 'BASE'
    df = pd.read_excel(uploaded_file, sheet_name='BASE')
    
    # 1. Procesar la columna de microcertificación
    df['Microcertificación que se encuentra terminando al momento de diligenciar este formulario'] = df.apply(
        lambda row: row['Microcertificación que se encuentra terminando al momento de diligenciar este formulario']
                    if pd.notna(row['Microcertificación que se encuentra terminando al momento de diligenciar este formulario'])
                    else (
                        row['Microcertificación que se encuentra terminando al momento de diligenciar este formulario2']
                        if pd.notna(row['Microcertificación que se encuentra terminando al momento de diligenciar este formulario2'])
                        else (
                            row['Microcertificación que se encuentra terminando al momento de diligenciar este formulario3']
                            if pd.notna(row['Microcertificación que se encuentra terminando al momento de diligenciar este formulario3'])
                            else row['Microcertificación que se encuentra terminando al momento de diligenciar este formulario4']
                        )
                    ),
        axis=1
    )
    # Renombrar la columna a 'nombre_micro'
    df.rename(columns={'Microcertificación que se encuentra terminando al momento de diligenciar este formulario': 'nombre_micro'}, inplace=True)
    
    # 2. Procesar la columna del formador
    df['Formador con el que tomó la microcertificación'] = df.apply(
        lambda row: row['Formador con el que tomó la microcertificación']
                    if pd.notna(row['Formador con el que tomó la microcertificación'])
                    else (
                        row['Formador con el que tomó la microcertificación2']
                        if pd.notna(row['Formador con el que tomó la microcertificación2'])
                        else (
                            row['Formador con el que tomó la microcertificación3']
                            if pd.notna(row['Formador con el que tomó la microcertificación3'])
                            else row['Formador con el que tomó la microcertificación4']
                        )
                    ),
        axis=1
    )
    # Renombrar a 'nombre_docente'
    df.rename(columns={'Formador con el que tomó la microcertificación': 'nombre_docente'}, inplace=True)
    
    # 3. Mapear respuestas textuales a valores numéricos
    mapa_respuestas = {
        "1. En Total Desacuerdo": 1,
        "2. En Desacuerdo": 2,
        "3. De Acuerdo": 3,
        "4. Totalmente de Acuerdo": 4
    }
    # Seleccionamos las columnas de las respuestas (suponiendo que son las columnas 20 a 35)
    columnas_a_usar = df.iloc[:, 19:35]
    columnas_a_usar = columnas_a_usar.replace(mapa_respuestas).astype(float)
    
    # 4. Definir las categorías y la cantidad de preguntas por cada una
    preguntas_por_categoria = {
        "Calidad": 2,
        "Pertinencia": 2,
        "Desempeño": 4,
        "Acompañamiento": 3,
        "Tiempos": 2,
        "Plataforma": 1,
        "Satisfacción": 2
    }
    # Calcular el promedio por categoría para cada encuestado
    promedios_por_categoria = {}
    inicio = 0
    for categoria, num_preguntas in preguntas_por_categoria.items():
        fin = inicio + num_preguntas
        cols = columnas_a_usar.iloc[:, inicio:fin]
        promedios_por_categoria[categoria] = cols.mean(axis=1)
        inicio = fin
    df_promedios = pd.DataFrame(promedios_por_categoria)
    
    # 5. Calcular promedios globales (conversión de escala de 4 a 5)
    promedios_finales = ((df_promedios.mean() / 4) * 5).round(1).to_frame(name='Resultado')
    promedio_general = (promedios_finales.mean()).round(1)
    promedios_finales.loc['Promedio general'] = promedio_general
    
    # 6. Extraer columnas categóricas y crear reporte de docentes
    df_categoricas = df[['nombre_micro', 'FORMADOR', 'Grupo', 'GP']]
    df_docentes = columnas_a_usar.iloc[:, [5, 6, 7, 8]]
    df_resultado_docentes = pd.concat([df_categoricas, df_docentes], axis=1)
    df_resultado_docentes['Promedio'] = df_docentes.mean(axis=1)
    
    # 7. Agrupar para evaluar desempeño por grupos
    df_desempeno_grupos = df_resultado_docentes.groupby(['nombre_micro', 'FORMADOR', 'Grupo']).agg({
        'El formador respondió y aclaró todas las preguntas, teniendo en cuenta las necesidades particulares.': 'mean',
        'La capacidad del formador para explicar los temas facilitó el aprendizaje durante las sesiones.': 'mean',
        'El formador promovió la participación activa y las   actividades prácticas, evitando que las sesiones se volvieran demasiado   teóricas.': 'mean',
        'El formador cumplió con los horarios establecidos,   iniciando y finalizando puntualmente cada sesión.': 'mean',
        'Promedio': ['mean', 'size']
    }).reset_index()
    df_desempeno_grupos.columns = ['Nombre del Microcurso',
                                   'Nombre del Formador',
                                   'Grupo',
                                   'Media de Atención al Estudiante',
                                   'Media de Capacidd de Facilitación',
                                   'Media de Participación Activa',
                                   'Media de Puntualidad',
                                   'Promedio General',
                                   'Número de Participantes']
    df_desempeno_grupos = df_desempeno_grupos.round(1)
    
    # 8. Evaluación de respuestas individuales
    df_respuestas_individuales = pd.concat([df_categoricas, columnas_a_usar], axis=1)
    df_agrupado = df_respuestas_individuales.groupby(['nombre_micro', 'FORMADOR', 'Grupo', 'GP']).agg(
        PARTICIPANTES = ('Grupo', 'size'),
        promedio_temas = ('Los temas tratados en la formación fueron presentados claramente para facilitar la comprensión.', 'mean'),
        promedio_material = ('El material del curso fue coherente, actualizado y útil para lograr el aprendizaje.', 'mean'),
        promedio_mercado = ('Lo aprendido durante la formación corresponde a los requerimientos del mercado laboral actual.', 'mean'),
        promedio_carrera = ('El aprendizaje alcanzado se puede aplicar directamente al trabajo o la carrera.', 'mean'),
        promedio_respuestas = ('El formador respondió y aclaró todas las preguntas, teniendo en cuenta las necesidades particulares.', 'mean'),
        promedio_explicacion = ('La capacidad del formador para explicar los temas facilitó el aprendizaje durante las sesiones.', 'mean'),
        promedio_participacion = ('El formador promovió la participación activa y las   actividades prácticas, evitando que las sesiones se volvieran demasiado   teóricas.', 'mean'),
        promedio_puntualidad = ('El formador cumplió con los horarios establecidos,   iniciando y finalizando puntualmente cada sesión.', 'mean'),
        promedio_atencion = ('El equipo de acompañamiento brindó apoyo y   atención oportuna durante el proceso de formación.', 'mean'),
        promedio_servicio = ('El personal de apoyo mostró disposición y actitud   de servicio durante toda la formación.', 'mean'),
        promedio_comunicacion = ('Las instrucciones y actualizaciones fueron claras   y oportunas, y los canales de comunicación favorecieron un desarrollo   eficiente.', 'mean'),
        promedio_duracion = ('La duración del proceso fue adecuada para cubrir   los temas previstos sin sobrecargar el tiempo de los participantes.', 'mean'),
        promedio_horarios = ('Los horarios de las sesiones fueron convenientes y   permitieron una participación efectiva.', 'mean'),
        promedio_plataforma = ('La plataforma tecnológica utilizada fue intuitiva   y permitió un fácil acceso a los contenidos del curso.', 'mean'),
        promedio_satisfaccion = ('La experiencia general con la formación fue   satisfactoria y se cumplieron las expectativas iniciales.', 'mean'),
        promedio_recomendacion = ('Recomendaría este programa de formación a otros   profesionales interesados en desarrollar sus habilidades ejecutivas.', 'mean'),
    ).reset_index()
    
    # 9. Ajustar escalas (convertir de 4 a 5)
    df_agrupado['promedio_respuestas'] = ((df_agrupado['promedio_respuestas'] / 4) * 5).round(1)
    df_agrupado['promedio_explicacion'] = ((df_agrupado['promedio_explicacion'] / 4) * 5).round(1)
    df_agrupado['promedio_temas'] = ((df_agrupado['promedio_temas'] / 4) * 5).round(1)
    df_agrupado['promedio_material'] = ((df_agrupado['promedio_material'] / 4) * 5).round(1)
    df_agrupado['promedio_mercado'] = ((df_agrupado['promedio_mercado'] / 4) * 5).round(1)
    df_agrupado['promedio_carrera'] = ((df_agrupado['promedio_carrera'] / 4) * 5).round(1)
    df_agrupado['promedio_participacion'] = ((df_agrupado['promedio_participacion'] / 4) * 5).round(1)
    df_agrupado['promedio_puntualidad'] = ((df_agrupado['promedio_puntualidad'] / 4) * 5).round(1)
    df_agrupado['promedio_atencion'] = ((df_agrupado['promedio_atencion'] / 4) * 5).round(1)
    df_agrupado['promedio_servicio'] = ((df_agrupado['promedio_servicio'] / 4) * 5).round(1)
    df_agrupado['promedio_comunicacion'] = ((df_agrupado['promedio_comunicacion'] / 4) * 5).round(1)
    df_agrupado['promedio_duracion'] = ((df_agrupado['promedio_duracion'] / 4) * 5).round(1)
    df_agrupado['promedio_horarios'] = ((df_agrupado['promedio_horarios'] / 4) * 5).round(1)
    df_agrupado['promedio_plataforma'] = ((df_agrupado['promedio_plataforma'] / 4) * 5).round(1)
    df_agrupado['promedio_satisfaccion'] = ((df_agrupado['promedio_satisfaccion'] / 4) * 5).round(1)
    df_agrupado['promedio_recomendacion'] = ((df_agrupado['promedio_recomendacion'] / 4) * 5).round(1)
    
    # 10. Calcular resultados finales por categoría
    df_agrupado['resultado_calidad_final'] = df_agrupado[['promedio_temas', 'promedio_material']].mean(axis=1).round(1)
    df_agrupado['resultado_pertinencia_final'] = df_agrupado[['promedio_mercado', 'promedio_carrera']].mean(axis=1).round(1)
    df_agrupado['resultado_desempeño_final'] = df_agrupado[['promedio_respuestas', 'promedio_explicacion', 'promedio_participacion', 'promedio_puntualidad']].mean(axis=1).round(1)
    df_agrupado['resultado_acompanamiento_final'] = df_agrupado[['promedio_atencion', 'promedio_servicio', 'promedio_comunicacion']].mean(axis=1).round(1)
    df_agrupado['resultado_tiempos_final'] = df_agrupado[['promedio_duracion', 'promedio_horarios']].mean(axis=1).round(1)
    df_agrupado['resultado_plataforma_final'] = df_agrupado['promedio_plataforma']
    df_agrupado['resultado_satisfaccion_final'] = df_agrupado[['promedio_satisfaccion', 'promedio_recomendacion']].mean(axis=1).round(1)
    

    # 10.1 Combinar el DataFrame de comentarios (por ejemplo, df_categoricas) con la columna específica de df
    columnas_comentarios = pd.concat([df_categoricas, df.iloc[:, 35]], axis=1)

    # 10.2. Agrupar y concatenar los comentarios en una sola columna por grupo
    df_comentarios = columnas_comentarios.groupby(['nombre_micro', 'FORMADOR', 'Grupo']).agg(
        comentarios_grupo=pd.NamedAgg(
            column="Según su experiencia, ¿qué recomendaciones o sugerencias considera importantes para mejorar la experiencia del programa  CLevel Propulsor?",
            aggfunc=lambda x: "; ".join(x.dropna().astype(str))
        )
    ).reset_index()

    # 10.3 Combinar los resultados de desempeño (df_agrupado) con los comentarios y resúmenes
    df_results = pd.concat([df_agrupado, df_comentarios[['comentarios_grupo']]], axis=1)

    # 11. Generar reporte consolidado por curso
    conteo_cursos = df_agrupado['nombre_micro'].value_counts()
    porcentaje_cursos = ((conteo_cursos / conteo_cursos.sum()) * 100)
    conteo_porcentaje_cursos = pd.DataFrame({'Conteo ciclo': conteo_cursos, 'Porcentaje': porcentaje_cursos}).round(0)
    
    def concatenar_valores(series):
        return '; '.join(series.dropna().unique())
    
    funciones_agregacion = {
        'PARTICIPANTES': 'sum',
        'FORMADOR': concatenar_valores,
        'Grupo': concatenar_valores
    }
    funciones_agregacion.update({
        col: 'mean'
        for col in df_agrupado.select_dtypes(include='number').columns
        if col not in ['PARTICIPANTES']
    })
    promedio_por_curso = df_agrupado.groupby('nombre_micro').agg(funciones_agregacion).reset_index()
    columnas_numericas = promedio_por_curso.select_dtypes(include='number').columns
    promedio_por_curso[columnas_numericas] = promedio_por_curso[columnas_numericas].round(1)
    promedio_por_curso.set_index('nombre_micro', inplace=True)
    
    resultados_curso = pd.concat([promedio_por_curso, conteo_porcentaje_cursos], axis=1)
    
    # 12. Escribir el resultado en un archivo Excel en memoria
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        promedios_finales.to_excel(writer, sheet_name='Resultado Ciclo')
        resultados_curso.to_excel(writer, sheet_name='Ev Cursos')
        df_agrupado.to_excel(writer, sheet_name='Ev Grupos')
        df_desempeno_grupos.to_excel(writer, sheet_name='Ev Formadores')
        df_results.to_excel(writer, sheet_name='Ev + Comentarios')
    output.seek(0)
    
    return output

# ---------------------------
# Interfaz de la aplicación
# ---------------------------
st.title("ETL Reporte Ciclos C-Level Propulsor")
st.markdown("""
Esta aplicación permite cargar el archivo **Propulsor Compensar.xlsx**, realizar las transformaciones definidas y generar el reporte **reporte_ciclo.xlsx** para su descarga.
""")

# Subida del archivo (se admite solo archivos Excel)
uploaded_file = st.file_uploader("Carga el archivo Propulsor Compensar.xlsx", type=["xlsx"])

if uploaded_file is not None:
    st.success("Archivo subido correctamente.")
    
    # Botón para iniciar el procesamiento
    if st.button("Procesar archivo"):
        try:
            # Procesar el archivo y obtener el Excel en memoria
            output_excel = process_file(uploaded_file)
            
            st.success("Archivo procesado correctamente.")
            st.download_button(
                label="Descargar reporte_ciclo.xlsx",
                data=output_excel,
                file_name="reporte_ciclo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error durante el procesamiento: {e}")
