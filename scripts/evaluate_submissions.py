# scripts/evaluate_submissions.py

import os
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# --- Configuración de rutas ---
import os

# --- Configuración de rutas ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(SCRIPT_DIR)
DATA_DIR = os.path.join(ROOT_DIR, 'data')
SUBMISSIONS_DIR = os.path.join(ROOT_DIR, 'user_submissions')
ORIGINAL_FILE = os.path.join(DATA_DIR, 'base_datos_original.xlsx')
EXPECTED_FILE = os.path.join(DATA_DIR, 'respuestas_esperadas.xlsx')
RESULTS_FILE = 'evaluation_results.xlsx'

# --- Variables globales para el informe ---
results = []

# --- Configuración de las preguntas ---
QUESTIONS = [
    {"num": 1, "topic": "Edición y formato", "question": "Cuantos ID tiene la base de datos", "expected": 30, "type": "numeric"},
    {"num": 2, "topic": "Edición y formato", "question": "Cambia el nombre de la columna 'Seguimiento' por 'Sentimiento'", "expected": "Sentimiento", "type": "text"},
    {"num": 3, "topic": "Edición y formato", "question": "En la tabla, proceda a centrar el contenido de todas las celdas", "expected": "Ajustar en la Base de Datos", "type": "format"},
    {"num": 4, "topic": "Edición y formato", "question": "En la tabla, ajusta el ancho de las columnas para que el texto completo sea visible", "expected": "Ajustar en la Base de Datos", "type": "format"},
    {"num": 5, "topic": "Cálculos", "question": "Calcula el número total de llamadas registradas en la tabla", "expected": 30, "type": "numeric"},
    {"num": 6, "topic": "Edición y formato", "question": "Utiliza la función 'Dar formato como tabla' en la tabla.", "expected": "Ajustar en la Base de Datos", "type": "format"},
    {"num": 7, "topic": "Cálculos", "question": "Cuántas llamadas tuvieron un Sentimiento 'Very Positive'", "expected": 3, "type": "numeric"},
    {"num": 8, "topic": "Cálculos", "question": "Calcula la duración promedio de las llamadas", "expected": 27, "type": "numeric"},
    {"num": 9, "topic": "Edición y formato", "question": "En la tabla, ajusta el formato fecha a 'dd/mm/yyyy'", "expected": "Ajustar en la Base de Datos", "type": "format"},
    {"num": 10, "topic": "Cálculos", "question": "Ordena la tabla por 'Puntuación' de mayor a menor y dime ¿Cuál es el puntaje máximo?", "expected": 10, "type": "numeric"},
    {"num": 11, "topic": "Cálculos", "question": "Cuantas llamadas hay con ese puntaje Máximo", "expected": 2, "type": "numeric"},
    {"num": 12, "topic": "Búsqueda", "question": "Si el 'ID' de un cliente es PJL-11752230. Dime cual es el nombre y apellido al que corresponde", "expected": "Linda Lopez", "type": "text"},
    {"num": 13, "topic": "Edición y formato", "question": "En la tabla, resalta en Rojo las celdas de la columna 'Puntuación' que sean inferiores a 5", "expected": "Ajustar en la Base de Datos", "type": "format"}
]

# --- Configuración de columnas ---
COLUMN_MAPPING = {
    "ID": "C",
    "Nombre del Cliente": "D",
    "Sentimiento": "E",  # Originalmente "Seguimiento", ahora "Sentimiento"
    "Puntuación": "F",
    "Fecha": "G",
    "Motivo": "H",
    "Ciudad": "I",
    "Canal": "J",
    "Duración Llamada (Minutos)": "K"
}

# --- Configuración de filas ---
HEADER_ROW = 5  # La fila donde están los encabezados
DATA_START_ROW = 6  # La fila donde empiezan los datos

# --- Configuración de hojas ---
SHEET_NAME = 'Datos'  # Nombre de la hoja con los datos
ANSWERS_SHEET = 'Respuestas'  # Nombre de la hoja donde el usuario escribe sus respuestas


def add_result(question_num, topic, question_text, status, observations=""):
    """Añade un resultado a la lista global."""
    results.append({
        "No.": question_num,
        "Tema": topic,
        "Pregunta": question_text,
        "Estado": status,
        "Observaciones": observations
    })


def evaluate_submission(submission_path):
    """Evalúa un archivo de envío de usuario."""
    user_filename = os.path.basename(submission_path)
    add_result("", "", f"--- Evaluando: {user_filename} ---", "")

    try:
        # Verificar que el archivo existe
        if not os.path.exists(submission_path):
            add_result("", "", f"Error al procesar {user_filename}", "Error", "Archivo no encontrado")
            return

        # Cargar el archivo
        wb_user = load_workbook(submission_path)
        
        # Verificar que las hojas necesarias existen
        required_sheets = [SHEET_NAME, ANSWERS_SHEET]
        for sheet in required_sheets:
            if sheet not in wb_user.sheetnames:
                add_result("", "", f"Error al procesar {user_filename}", "Error", f"Hoja '{sheet}' no encontrada en el archivo")
                return

        # Obtener las hojas
        ws_data = wb_user[SHEET_NAME]
        ws_answers = wb_user[ANSWERS_SHEET]

        # Verificar y limpiar los nombres de las columnas
        # Obtener los nombres de las columnas desde la hoja
        column_names = []
        for col_letter in COLUMN_MAPPING.values():
            cell_value = ws_user[f'{col_letter}{HEADER_ROW}'].value
            if cell_value:
                # Limpiar espacios y caracteres especiales
                clean_value = str(cell_value).strip()
                column_names.append(clean_value)
            else:
                column_names.append(None)

        # Verificar que todas las columnas requeridas existen
        missing_columns = []
        for col_name, col_letter in COLUMN_MAPPING.items():
            cell_value = column_names[list(COLUMN_MAPPING.values()).index(col_letter)]
            if cell_value != col_name:
                missing_columns.append(col_name)
                
        if missing_columns:
            # Mostrar los nombres reales de las columnas para depuración
            actual_cols = {k: v for k, v in zip(COLUMN_MAPPING.keys(), column_names)}
            add_result("", "", f"Error al procesar {user_filename}", "Error", 
                      f"Columnas faltantes o con nombres incorrectos: {', '.join(missing_columns)}.\n"
                      f"Columnas encontradas: {actual_cols}")
            return

        # Cargar el archivo esperado como referencia
        wb_expected = load_workbook(EXPECTED_FILE)
        
        # Verificar que la hoja 'Datos' existe en el archivo de respuestas esperadas
        sheet_names_expected = [name.strip() for name in wb_expected.sheetnames]
        if SHEET_NAME not in sheet_names_expected:
            add_result("", "", f"Error al procesar {user_filename}", "Error", 
                      f"Hoja '{SHEET_NAME}' no encontrada en el archivo de respuestas esperadas. Hojas disponibles: {', '.join(wb_expected.sheetnames)}")
            return
            
        # Encontrar la hoja correcta en el archivo de respuestas esperadas
        ws_expected = next(ws for ws in wb_expected.worksheets if ws.title.strip() == SHEET_NAME)

        # Cargar los datos usando pandas
        try:
            # Obtener el nombre exacto de la hoja para pandas
            sheet_name_user = wb_user.sheetnames[0]  # Usar la primera hoja encontrada
            sheet_name_expected = wb_expected.sheetnames[0]  # Usar la primera hoja encontrada

            # Generar el rango de columnas
            first_col_letter = COLUMN_MAPPING["ID"]
            last_col_letter = COLUMN_MAPPING["Duración Llamada (Minutos)"]
            column_range = f"{first_col_letter}:{last_col_letter}"

            # Cargar DataFrames usando el nombre exacto de la hoja
            df_user = pd.read_excel(submission_path, sheet_name=sheet_name_user, 
                                   header=HEADER_ROW - 1, usecols=column_range)
            
            # Limpiar los nombres de las columnas
            df_user.columns = df_user.columns.str.strip()
            
            df_expected = pd.read_excel(EXPECTED_FILE, sheet_name=sheet_name_expected, 
                                      header=HEADER_ROW - 1, usecols=column_range)
        except Exception as e:
            add_result("", "", f"Error al procesar {user_filename}", "Error", 
                      f"Error al cargar los datos: {str(e)}")
            return

        # Verificar que los DataFrames tienen las columnas esperadas
        expected_columns = list(COLUMN_MAPPING.keys())
        missing_df_columns = [col for col in expected_columns if col not in df_user.columns]
        
        if missing_df_columns:
            missing_cols_str = ", ".join(missing_df_columns)
            add_result("", "", f"Error al procesar {user_filename}", "Error", 
                      f"Columnas faltantes en DataFrame: {missing_cols_str}")
            return

        # --- PREGUNTAS DE EVALUACIÓN ---
        
        # Verificar las respuestas del usuario
        for question in QUESTIONS:
            question_num = question["num"]
            topic = question["topic"]
            question_text = question["question"]
            expected = question["expected"]
            q_type = question["type"]
            
            try:
                # Buscar la respuesta del usuario en la hoja de respuestas
                # Suponemos que las respuestas están en la columna B (B2 para pregunta 1, B3 para pregunta 2, etc.)
                answer_cell = f"B{question_num + 1}"  # +1 porque empezamos en B2
                user_answer = ws_answers[answer_cell].value
                
                if user_answer is None:
                    add_result(question_num, topic, question_text, "Error", "Respuesta no encontrada")
                    continue
                    
                if q_type == "numeric":
                    # Convertir a número si es necesario
                    try:
                        user_answer = float(user_answer)
                        expected = float(expected)
                    except ValueError:
                        add_result(question_num, topic, question_text, "Error", f"Respuesta no es un número: {user_answer}")
                        continue
                        
                    if user_answer == expected:
                        add_result(question_num, topic, question_text, "Correcto")
                    else:
                        add_result(question_num, topic, question_text, "Incorrecto", 
                                  f"Esperado: {expected}, Obtenido: {user_answer}")
                        
                elif q_type == "text":
                    # Comparar texto (case insensitive)
                    if str(user_answer).strip().lower() == str(expected).strip().lower():
                        add_result(question_num, topic, question_text, "Correcto")
                    else:
                        add_result(question_num, topic, question_text, "Incorrecto", 
                                  f"Esperado: '{expected}', Obtenido: '{user_answer}'")
                        
                elif q_type == "format":
                    # Para preguntas de formato, solo verificar que la respuesta existe
                    add_result(question_num, topic, question_text, "Correcto")
                    
            except Exception as e:
                add_result(question_num, topic, question_text, "Error", f"Error al evaluar: {str(e)}")
        # Usar pandas para contar los IDs ya que es más robusto
        expected_ids = 30
        try:
            # Contar IDs usando pandas (más robusto)
            actual_ids = len(df_user["ID"])  # Usar la columna ID del DataFrame
            if isinstance(actual_ids, int) and isinstance(expected_ids, int):
                if actual_ids == expected_ids:
                    add_result(1, "Edición y formato", "Cuantos ID tiene la base de datos", "Correcto")
                else:
                    add_result(1, "Edición y formato", "Cuantos ID tiene la base de datos", "Incorrecto",
                              f"Esperado: {expected_ids}, Obtenido: {actual_ids}")
            else:
                add_result(1, "Edición y formato", "Cuantos ID tiene la base de datos", "Error",
                          f"Error en tipos de datos: expected_ids={type(expected_ids)}, actual_ids={type(actual_ids)}")
        except Exception as e:
            add_result(1, "Edición y formato", "Cuantos ID tiene la base de datos", "Error",
                      f"Error al contar IDs: {str(e)}")

        # Pregunta 2: Cambia el nombre de la columna "Seguimiento" por "Sentimiento"
        # ### AJUSTE POR C5 ###
        expected_col_name = "Sentimiento"
        actual_col_name = ws_user[f'{COLUMN_MAPPING["Sentimiento"]}{HEADER_ROW}'].value
        if actual_col_name == expected_col_name:
            add_result(2, "Edición y formato", "Cambia el nombre de la columna 'Seguimiento' por 'Sentimiento'",
                       "Correcto")
        else:
            add_result(2, "Edición y formato", "Cambia el nombre de la columna 'Seguimiento' por 'Sentimiento'",
                       "Incorrecto",
                       f"Nombre de columna en usuario: '{actual_col_name}', Esperado: '{expected_col_name}'")

        # Pregunta 3: Centra el contenido de todas las celdas
        # ### AJUSTE POR C5 ###
        # Ajusta el rango de muestreo para que empiece en C5 y cubra las columnas relevantes
        all_cells_centered = True
        # Rango de la tabla de datos: desde C5 hasta K_ultima_fila
        # Para evitar recorrer celdas vacías fuera de la tabla, usaremos un rango aproximado
        max_row_data = ws_user.cell(ws_user.max_row, COLUMN_MAPPING["ID"]).row  # Última fila con datos

        # Muestreamos solo el área donde están los datos y encabezados de la tabla
        for r_idx in range(HEADER_ROW, max_row_data + 1):
            for c_idx in range(ws_user.cell(HEADER_ROW, COLUMN_MAPPING["ID"]).column,
                               ws_user.cell(HEADER_ROW, COLUMN_MAPPING["Duración Llamada (Minutos)"]).column + 1):
                cell = ws_user.cell(row=r_idx, column=c_idx)
                if cell.alignment.horizontal != 'center':
                    all_cells_centered = False
                    break
            if not all_cells_centered:
                break

        if all_cells_centered:
            add_result(3, "Edición y formato", "Centrar el contenido de todas las celdas", "Correcto",
                       "Verificación muestral sobre el rango de la tabla.")
        else:
            add_result(3, "Edición y formato", "Centrar el contenido de todas las celdas", "Incorrecto",
                       "No todas las celdas muestreadas dentro del rango de la tabla están centradas.")

        # Pregunta 4: Ajusta el ancho de las columnas
        # ### AJUSTE POR C5 ###
        # Comparamos el ancho de las columnas relevantes (C a K)
        width_adjusted = True
        for col_letter in COLUMN_MAPPING.values():
            user_col_width = ws_user.column_dimensions[col_letter].width
            expected_col_width = ws_expected.column_dimensions[col_letter].width

            # Si el usuario no ha tocado el ancho, openpyxl puede devolver None
            # y el ancho por defecto es 8.43. Si es menor que un mínimo razonable, es incorrecto.
            if user_col_width is None or user_col_width < 8:  # Un ancho mínimo para que sea legible
                width_adjusted = False
                break
            # Además, podemos comparar con el ancho del archivo esperado, si lo definimos como 'óptimo'
            if expected_col_width is not None and user_col_width < expected_col_width * 0.9:  # 10% de tolerancia
                width_adjusted = False
                break

        if width_adjusted:
            add_result(4, "Edición y formato", "Ajusta el ancho de las columnas", "Correcto",
                       "Comparado con anchos de columnas esperados (con tolerancia) en el rango de la tabla.")
        else:
            add_result(4, "Edición y formato", "Ajusta el ancho de las columnas", "Incorrecto",
                       "El ancho de algunas columnas no parece ajustado correctamente en el rango de la tabla.")

        # Pregunta 5: Calcula el número total de llamadas registradas
        # Es la misma verificación que P1, usando el conteo de filas de pandas
        if df_user.shape[0] == expected_ids:  # df_user.shape[0] ya es el número de filas de datos
            add_result(5, "Fórmulas", "Calcula el número total de llamadas registradas", "Correcto")
        else:
            add_result(5, "Fórmulas", "Calcula el número total de llamadas registradas", "Incorrecto",
                       f"Esperado: {expected_ids}, Obtenido: {df_user.shape[0]}")

        # Pregunta 6: Utiliza la función "Dar formato como tabla"
        # ### AJUSTE POR C5 ###
        # Obtener el rango de la tabla de datos, asumiendo C5:K<ultima_fila>
        table_range_str = f"{COLUMN_MAPPING['ID']}{HEADER_ROW}:{COLUMN_MAPPING['Duración Llamada (Minutos)']}{ws_user.max_row}"
        table_found = False
        if ws_user.tables:
            for table_name, table_obj in ws_user.tables.items():
                if table_obj.ref == table_range_str:  # Compara si el rango de la tabla coincide con el esperado
                    table_found = True
                    break
        if table_found:
            add_result(6, "Fórmulas", "Utiliza la función 'Dar formato como tabla'", "Correcto")
        else:
            add_result(6, "Fórmulas", "Utiliza la función 'Dar formato como tabla'", "Incorrecto",
                       f"No se encontró un formato de tabla aplicado correctamente en la hoja 'Datos' cubriendo el rango {table_range_str}.")

        # Pregunta 7: Cuántas llamadas tuvieron un Sentimiento "Very Positive"
        expected_very_positive = 4
        # Asegúrate de que la columna "Sentimiento" exista después del posible cambio de nombre
        if "Sentimiento" in df_user.columns:
            actual_very_positive = df_user['Sentimiento'].astype(str).str.strip().eq("Very Positive").sum()
            if actual_very_positive == expected_very_positive:
                add_result(7, "Fórmulas", "Cuántas llamadas tuvieron un Sentimiento 'Very Positive'", "Correcto")
            else:
                add_result(7, "Fórmulas", "Cuántas llamadas tuvieron un Sentimiento 'Very Positive'", "Incorrecto",
                           f"Esperado: {expected_very_positive}, Obtenido: {actual_very_positive}")
        else:
            add_result(7, "Fórmulas", "Cuántas llamadas tuvieron un Sentimiento 'Very Positive'", "Incorrecto",
                       "La columna 'Sentimiento' no se encontró en el archivo del usuario.")

        # Pregunta 8: Calcula la duración promedio de las llamadas
        # Suma de todas las duraciones que proporcionaste (14+26+...+36 = 780).
        # Total de filas de datos que proporcionaste = 30.
        expected_avg_duration = 780 / 30  # 26.0 (Calcula esto con tus 30 datos completos)

        if "Duración Llamada (Minutos)" in df_user.columns:
            df_user['Duración Llamada (Minutos)'] = pd.to_numeric(df_user['Duración Llamada (Minutos)'],
                                                                  errors='coerce')
            actual_avg_duration = df_user['Duración Llamada (Minutos)'].mean()
            if abs(actual_avg_duration - expected_avg_duration) < 0.01:  # Tolerancia para floats
                add_result(8, "Fórmulas", "Calcula la duración promedio de las llamadas", "Correcto")
            else:
                add_result(8, "Fórmulas", "Calcula la duración promedio de las llamadas", "Incorrecto",
                           f"Esperado: {expected_avg_duration:.2f}, Obtenido: {actual_avg_duration:.2f}")
        else:
            add_result(8, "Fórmulas", "Calcula la duración promedio de las llamadas", "Incorrecto",
                       "La columna 'Duración Llamada (Minutos)' no se encontró en el archivo del usuario.")

        # Pregunta 9: Ajusta el formato fecha a "dd/mm/yyyy"
        # ### AJUSTE POR C5 ###
        date_format_correct = True
        date_col_letter = COLUMN_MAPPING["Fecha"]
        # Recorre las celdas de datos de la columna de fecha (desde DATA_START_ROW)
        for row_idx in range(DATA_START_ROW, ws_user.max_row + 1):
            cell = ws_user[f'{date_col_letter}{row_idx}']
            if cell.data_type != 'd' or cell.number_format != 'dd/mm/yyyy':
                date_format_correct = False
                break
        if date_format_correct:
            add_result(9, "Filtrado", "Ajusta el formato fecha a 'dd/mm/yyyy'", "Correcto")
        else:
            add_result(9, "Filtrado", "Ajusta el formato fecha a 'dd/mm/yyyy'", "Incorrecto",
                       "El formato de fecha no es 'dd/mm/yyyy' o el tipo de dato no es fecha en todas las celdas de la columna 'Fecha'.")

        # Pregunta 10: Ordena la tabla por "Puntuación" de mayor a menor y dime ¿Cuál es el puntaje máximo?
        if "Puntuación" in df_user.columns:
            df_user_sorted_by_score = df_user.sort_values(by="Puntuación", ascending=False).reset_index(drop=True)
            df_expected_sorted_by_score = df_expected.sort_values(by="Puntuación", ascending=False).reset_index(
                drop=True)

            # Compara si los DataFrames son idénticos después de ordenar
            order_correct = df_user_sorted_by_score.equals(df_expected_sorted_by_score)

            max_score_expected = 10
            max_score_actual = df_user['Puntuación'].max()

            if order_correct and max_score_actual == max_score_expected:
                add_result(10, "Filtrado", "Ordena la tabla por 'Puntuación' de mayor a menor y puntaje máximo",
                           "Correcto")
            else:
                obs = []
                if not order_correct: obs.append("Tabla no ordenada correctamente por Puntuación (Mayor a Menor).")
                if max_score_actual != max_score_expected: obs.append(
                    f"Puntaje máximo esperado: {max_score_expected}, Obtenido: {max_score_actual}.")
                add_result(10, "Filtrado", "Ordena la tabla por 'Puntuación' de mayor a menor y puntaje máximo",
                           "Incorrecto", " ".join(obs))
        else:
            add_result(10, "Filtrado", "Ordena la tabla por 'Puntuación' de mayor a menor y puntaje máximo",
                       "Incorrecto", "La columna 'Puntuación' no se encontró.")

        # Pregunta 11: Cuantas llamadas hay con ese puntaje Máximo
        expected_count_max_score = 5  # 5 llamadas con puntuación 10 en los 30 datos
        if "Puntuación" in df_user.columns:
            max_score = df_user['Puntuación'].max()
            actual_count_max_score = df_user[df_user['Puntuación'] == max_score].shape[0]
            if actual_count_max_score == expected_count_max_score:
                add_result(11, "Filtrado", "Cuantas llamadas hay con ese puntaje Máximo", "Correcto")
            else:
                add_result(11, "Filtrado", "Cuantas llamadas hay con ese puntaje Máximo", "Incorrecto",
                           f"Esperado: {expected_count_max_score}, Obtenido: {actual_count_max_score}.")
        else:
            add_result(11, "Filtrado", "Cuantas llamadas hay con ese puntaje Máximo", "Incorrecto",
                       "La columna 'Puntuación' no se encontró.")

        # Pregunta 12: Si el "ID" de un cliente es PIL-11752230. Dime cual es el nombre y apellido al que corresponde
        search_id = "PJL-11752230"  # Corregí el ID a PJL-11752230, según tu lista de IDs
        expected_name = "Linda Lopez"
        if "ID" in df_user.columns and "Nombre del Cliente" in df_user.columns:
            found_row = df_user[df_user['ID'].astype(str).str.strip() == search_id]
            if not found_row.empty:
                actual_name = found_row['Nombre del Cliente'].iloc[0]
                if actual_name == expected_name:
                    add_result(12, "Filtrado", f"Nombre del cliente con ID {search_id}", "Correcto")
                else:
                    add_result(12, "Filtrado", f"Nombre del cliente con ID {search_id}", "Incorrecto",
                               f"Esperado: '{expected_name}', Obtenido: '{actual_name}'.")
            else:
                add_result(12, "Filtrado", f"Nombre del cliente con ID {search_id}", "Incorrecto",
                           f"ID '{search_id}' no encontrado en la base de datos del usuario.")
        else:
            add_result(12, "Filtrado", f"Nombre del cliente con ID {search_id}", "Incorrecto",
                       "Las columnas 'ID' o 'Nombre del Cliente' no se encontraron.")

        # Pregunta 13: Resalta en Rojo las celdas de la columna "Puntuación" que sean inferiores a 5, usando Formato Condicional.
        # ### AJUSTE POR C5 ###
        conditional_formatting_correct = False
        score_col_letter = COLUMN_MAPPING["Puntuación"]
        score_col_idx = ws_user[f'{score_col_letter}{HEADER_ROW}'].column  # Obtener índice numérico de la columna

        # Verificar si hay alguna regla de formato condicional que coincida con la condición y el rango.
        expected_rgb_color = 'FFFF0000'  # Rojo puro para la fuente

        for cf_rule in ws_user.conditional_formatting.cf_rules:
            applies_to_score_column = False
            for cf_range in cf_rule.ranges:
                # Check if the range (e.g., 'F6:F35') contains 'F' and the data rows
                if score_col_letter in cf_range.coord and cf_range.coord.endswith(
                        f'{DATA_START_ROW}:{score_col_letter}{ws_user.max_row}'):
                    applies_to_score_column = True
                    break

            if applies_to_score_column and cf_rule.type == 'cellIs' and cf_rule.operator == 'lessThan':
                if cf_rule.formula and cf_rule.formula[0] == "5":
                    if cf_rule.dxf and cf_rule.dxf.font and cf_rule.dxf.font.color and cf_rule.dxf.font.color.rgb == expected_rgb_color:
                        conditional_formatting_correct = True
                        break

        if conditional_formatting_correct:
            add_result(13, "Condicional", "Resalta en Rojo las celdas de la columna 'Puntuación' (< 5) con FC",
                       "Correcto", "Se detectó una regla de formato condicional adecuada (menos de 5, rojo).")
        else:
            add_result(13, "Condicional", "Resalta en Rojo las celdas de la columna 'Puntuación' (< 5) con FC",
                       "Incorrecto",
                       "No se detectó una regla de formato condicional correcta para la columna Puntuación (menos de 5, color rojo) o no aplica al rango esperado.")

        # Pregunta 14: Crea un gráfico de barras para mostrar la relación entre "Nombre del cliente" y "Puntuación".
        chart1_found = False
        for sheet_name in wb_user.sheetnames:
            ws_current = wb_user[sheet_name]
            for chart_obj in ws_current._charts:
                if chart_obj.type == 'bar':
                    chart1_found = True
                    break
            if chart1_found: break
        if chart1_found:
            add_result(14, "Gráficos", "Crea un gráfico de barras (Nombre del cliente vs Puntuación)", "Correcto",
                       "Gráfico de barras detectado (verificación superficial de tipo).")
        else:
            add_result(14, "Gráficos", "Crea un gráfico de barras (Nombre del cliente vs Puntuación)", "Incorrecto",
                       "No se encontró un gráfico de barras adecuado.")

        # Pregunta 15: Crea un gráfico que permita mostrar la relación entre la "Duración de la Llamada" y la "Puntuación".
        chart2_found = False
        for sheet_name in wb_user.sheetnames:
            ws_current = wb_user[sheet_name]
            for chart_obj in ws_current._charts:
                if chart_obj.type in ['scatter', 'bar', 'line']:
                    chart2_found = True
                    break
            if chart2_found: break
        if chart2_found:
            add_result(15, "Gráficos", "Crea un gráfico (Duración de la Llamada vs Puntuación)", "Correcto",
                       "Gráfico de relación Duración/Puntuación detectado (verificación superficial de tipo).")
        else:
            add_result(15, "Gráficos", "Crea un gráfico (Duración de la Llamada vs Puntuación)", "Incorrecto",
                       "No se encontró un gráfico adecuado para la relación Duración/Puntuación.")

        wb_user.close()
        wb_expected.close()

    except FileNotFoundError:
        add_result("", "", f"Error: Archivo de usuario no encontrado: {user_filename}", "Error",
                   "Asegúrate de que el archivo exista en la carpeta 'user_submissions'.")
    except KeyError as e:
        add_result("", "", f"Error en la evaluación de {user_filename}", "Error",
                   f"Hoja 'Datos' no encontrada o columna faltante: {e}. Asegúrate que la hoja se llama 'Datos' y las columnas son correctas.")
    except Exception as e:
        add_result("", "", f"Error inesperado al evaluar {user_filename}", "Error", f"Detalles: {e}")

    add_result("", "", "--- Fin de evaluación ---", "")


def generate_report():
    """Genera el archivo Excel con los resultados de la evaluación."""
    df_results = pd.DataFrame(results)

    # Crear un nuevo workbook y sheet para los resultados
    wb_results = Workbook()
    ws_report = wb_results.active
    ws_report.title = "Resultados_Evaluacion"

    # Escribir el DataFrame en la hoja
    for r_idx, row in enumerate(dataframe_to_rows(df_results, index=False, header=True), 1):
        ws_report.append(row)

    # Aplicar formato a los encabezados
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    for cell in ws_report[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Ajustar ancho de columnas
    for col in ws_report.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value is not None and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        if adjusted_width > 100: adjusted_width = 100
        ws_report.column_dimensions[column].width = adjusted_width

    # Formato condicional para el estado
    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

    for row_idx in range(2, ws_report.max_row + 1):
        status_cell = ws_report.cell(row=row_idx, column=4)
        if status_cell.value == "Incorrecto":
            for col_idx in range(1, ws_report.max_column + 1):
                ws_report.cell(row=row_idx, column=col_idx).fill = red_fill
        elif status_cell.value == "Correcto":
            for col_idx in range(1, ws_report.max_column + 1):
                ws_report.cell(row=row_idx, column=col_idx).fill = green_fill
        elif status_cell.value and str(status_cell.value).startswith("---"):
            for col_idx in range(1, ws_report.max_column + 1):
                ws_report.cell(row=row_idx, column=col_idx).fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0",
                                                                               fill_type="solid")
                ws_report.cell(row=row_idx, column=col_idx).font = Font(bold=True)

    wb_results.save(RESULTS_FILE)
    print(f"\nInforme de evaluación generado en '{RESULTS_FILE}'")


def main():
    """Función principal para ejecutar la evaluación."""
    print("Iniciando la evaluación de archivos de usuario...")
    if not os.path.exists(SUBMISSIONS_DIR):
        print(
            f"Error: No se encontró la carpeta de envíos '{SUBMISSIONS_DIR}'. Crea esta carpeta y coloca los archivos de los usuarios aquí.")
        return

    submission_files = [f for f in os.listdir(SUBMISSIONS_DIR) if f.endswith('.xlsx') or f.endswith('.xlsm')]

    if not submission_files:
        print(f"No se encontraron archivos .xlsx o .xlsm en '{SUBMISSIONS_DIR}'.")
        return

    for filename in submission_files:
        submission_path = os.path.join(SUBMISSIONS_DIR, filename)
        print(f"Procesando: {filename}")
        evaluate_submission(submission_path)

    generate_report()
    print("Evaluación completada.")


if __name__ == "__main__":
    main()