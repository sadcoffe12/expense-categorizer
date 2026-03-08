import pandas as pd
import openpyxl
import os
import unicodedata
import re

RULES_FILE = "categorization_rules.txt"
TEMPLATES_FILE = "templates.txt"

def load_rules():
    #Carga las reglas de categorización desde el archivo de texto en orden.
    if not os.path.exists(RULES_FILE):
        return []
    
    rules = []
    try:
        with open(RULES_FILE, 'r', encoding="utf-8") as f:
            for line in f:
                parts = line.strip().split(':', 3)
                if len(parts) == 4:
                    keyword, type, category, new_description = parts
                    rules.append((keyword.lower(), type, category, None if new_description == 'None' else new_description if new_description else None))
    except Exception as e:
        print(f"Error al cargar las reglas: {e}")
        return []
    return rules

def normalize_text(text):
    if not isinstance(text, str):
        return text
    # Convertir a minúsculas y quitar espacios en los extremos
    text = text.lower().strip()
    # Eliminar acentos
    text = ''.join(c for c in unicodedata.normalize('NFD', text)
                if unicodedata.category(c) != 'Mn')
    # Quitar símbolos y puntuación (deja solo letras, números y espacios)
    text = re.sub(r'[^\w\s]', '', text)
    # Reemplazar múltiples espacios por uno solo
    text = re.sub(r'\s+', ' ', text)
    return text

def select_column(df, prompt, default_name, template_suggestion=None):
    # Si el template ya trae una columna válida, la usamos sin preguntar
    if template_suggestion and template_suggestion in df.columns:
        print(f"[Template] Usando columna '{template_suggestion}' para {prompt}.")
        return template_suggestion

    # Si no hay sugerencia o no es válida, procedemos con la selección manual
    print(f"\n--- Selección de Columna: {prompt} ---")
    for i, col_name in enumerate(df.columns):
        print(f"  {i+1}. {col_name}")
    
    user_input = input(f"Selecciona número o nombre (Enter para '{default_name}'): ").strip()

    if not user_input:
        return default_name
    
    if user_input.isdigit():
        idx = int(user_input) - 1
        return df, None.columns[idx] if 0 <= idx < len(df.columns) else default_name
    
    return user_input

def categorize(df, selected_template=None):
    #Categoriza los registros basándose en reglas y palabras clave en orden secuencial.
    print("\n--- Categorizar Registros ---")
    rules = load_rules()
    
    if not rules:
        print("No hay reglas de categorización definidas. Crea algunas primero.")
        return df

    # Extraemos las sugerencias del template si existen
    t_source = selected_template.get('SOURCE_COL') if selected_template else None
    t_type = selected_template.get('TYPE_COL') if selected_template else None
    t_cat = selected_template.get('CAT_COL') if selected_template else None

    # Llamadas a la función de selección
    source_col = select_column(df, "Descripción", "Descripcion", t_source)
    
    type_col = select_column(df, "Tipo de Gasto", "Tipo", t_type)
    if type_col not in df.columns:
        df[type_col] = ""
            
    category_col = select_column(df, "Categoría", "Categoria", t_cat)
    if category_col not in df.columns:
        df[category_col] = ""
        
    categorized_count = 0
    df[source_col] = df[source_col].astype(str)
    df[category_col] = df[category_col].astype(str)
    df[type_col] = df[type_col].astype(str)

    def assign_category(description):
        nonlocal categorized_count
        desc_value = description
        cat_value = ""
        type_value = ""

        desc_lower = desc_value.lower()

        # Recorre las reglas en orden
        for keyword, type, category, new_desc in rules:
                        
            if keyword in desc_lower:
                categorized_count += 1
                cat_value = category
                type_value = type
                if new_desc:
                    desc_value = new_desc
                    desc_lower = desc_value.lower()  # actualiza el texto para reglas posteriores

        return type_value, cat_value, desc_value

    # Aplicar reglas
    results = df[source_col].apply(assign_category)
    results_list = results.tolist()
    results_df = pd.DataFrame(results_list, columns=[type_col, category_col, source_col])
    # Ensure columns are assigned in the correct order
    df[type_col] = results_df[type_col]
    df[category_col] = results_df[category_col]
    df[source_col] = results_df[source_col]

    df.columns = [col.replace('_', ' ').capitalize() for col in df.columns]
    print(f"\n Proceso de categorización completado. Se categorizaron {categorized_count} coincidencias.")
    return df

def apply_format(df, file_path):
    #Aplica un template para limpiar un archivo, y luego elimina/añade columnas/filas 
    # en el orden correcto. Devuelve solo el DataFrame en memoria (no guarda archivo).

    print("--- Aplicación de un template de formato de limpieza ---")
    
    if not os.path.exists(TEMPLATES_FILE):
        print("No se encontraron templates de formato. Por favor, crea uno primero.")
        return df, None

    # --- (La sección de lectura y selección de templates no cambia) ---
    templates = {}
    try:
        with open(TEMPLATES_FILE, "r") as f:
            content = f.read().strip()
            if not content:
                print("El archivo de templates está vacío.")
                return df, None, None
            
            lines = content.split("---")
            for line_block in lines:
                line_block = line_block.strip()
                if not line_block: 
                    continue
                
                current_template = {}
                for line in line_block.split('\n'):
                    line = line.strip()
                    if ':' in line:
                        try:
                            key, value = line.split(": ", 1)
                            key = key.strip()
                            value = value.strip()
                            
                            keys_to_normalize = ['COLS_TO_ADD', 'ORDERED_COLS', 'SOURCE_COL', 'TYPE_COL', 'CAT_COL']
                            if key in keys_to_normalize:
                                if ',' in value:
                                    value = ", ".join([normalize_text(v) for v in value.split(',')])
                                else:
                                    value = normalize_text(value)
                            
                            current_template[key] = value
                        except ValueError:
                            continue
                
                if 'TEMPLATE_NAME' in current_template:
                    templates[current_template['TEMPLATE_NAME']] = current_template
                    
    except Exception as e:
        print(f"Error al leer los templates: {e}")
        return df, None

    if not templates:
        print("No se encontraron templates de formato.")
        return df, None

    print("Templates de formato disponibles:")
    for i, name in enumerate(templates.keys()):
        print(f"  {i+1}. {name}")
    
    try:
        choice = int(input("Selecciona el número del template a aplicar: "))
        selected_name = list(templates.keys())[choice - 1]
        selected_template = templates[selected_name]
    except (ValueError, IndexError):
        print("Selección no válida.")
        return df, None

    # --- Comienza la lógica de aplicación ---
    try:
        # 1. Leer los datos del rango especificado
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active

        header_row = int(selected_template['HEADER_ROW'])
        start_row = int(selected_template['START_ROW'])
        start_col_letter = selected_template['START_COL']
        end_col_letter = selected_template['END_COL']
        
        start_col_idx = openpyxl.utils.column_index_from_string(start_col_letter)
        end_col_idx = openpyxl.utils.column_index_from_string(end_col_letter)

        data = []
        for r in range(header_row, ws.max_row + 1):
            row_data = [ws.cell(row=r, column=c).value for c in range(start_col_idx, end_col_idx + 1)]
            data.append(row_data)

        # 2. Crear el DataFrame en memoria
        if not data:
            print("No se encontraron datos en el rango especificado.")
            return df, None
            
        df_new = pd.DataFrame(data[1:], columns=data[0])
        
        # Limpieza de encabezados
        df_new.columns = [normalize_text(col) for col in df_new.columns]
        print("Encabezados normalizados (sin tildes ni símbolos).")

        df_new.dropna(how='all', inplace=True)

        # 3. Eliminar columnas especificadas
        if 'COLS_TO_DROP' in selected_template and selected_template['COLS_TO_DROP']:
            cols_to_drop = [col.strip() for col in selected_template['COLS_TO_DROP'].split(',') if col.strip()]
            headers_map = {openpyxl.utils.get_column_letter(start_col_idx + i): header for i, header in enumerate(data[0])}
            cols_to_drop_names = [headers_map[letter] for letter in cols_to_drop if letter in headers_map]
            cols_to_drop_names_existing = [name for name in cols_to_drop_names if name in df_new.columns]
            
            df_new.drop(columns=cols_to_drop_names_existing, inplace=True, errors='ignore')
            if cols_to_drop_names_existing:
                print(f"Columnas eliminadas: {', '.join(cols_to_drop_names_existing)}")

        # 4. Eliminar filas especificadas
        if 'ROWS_TO_DROP' in selected_template and selected_template['ROWS_TO_DROP']:
            rows_to_drop_str = [r.strip() for r in selected_template['ROWS_TO_DROP'].split(',') if r.strip()]
            rows_to_drop_indices = [int(r_str) - start_row for r_str in rows_to_drop_str]
            rows_to_drop_indices = [idx for idx in rows_to_drop_indices if 0 <= idx < len(df_new)]
            
            df_new.drop(index=rows_to_drop_indices, inplace=True, errors='ignore')
            if rows_to_drop_indices:
                print(f"Se eliminaron {len(rows_to_drop_indices)} filas especificadas.")
        
        # 5. Añadir nuevas columnas (Ahora más simple, sin necesidad de letras)
        if 'COLS_TO_ADD' in selected_template and selected_template['COLS_TO_ADD']:
            items_to_add = [item.strip() for item in selected_template['COLS_TO_ADD'].split(',')]
            
            for item in items_to_add:
                # Extraemos el nombre antes del paréntesis si existe, o el nombre directo
                name = item.split('(')[0].strip()
                if name not in df_new.columns:
                    df_new[name] = "" # Crea la columna al final
                    print(f"Columna creada: '{name}'")

        # 6. Reordenamiento Final Dinámico
        if 'ORDERED_COLS' in selected_template and selected_template['ORDERED_COLS']:
            # Como ya normalizamos al leer el archivo, solo limpiamos comillas si existieran
            raw_ordered_cols = selected_template['ORDERED_COLS'].replace("'", "").replace('"', "")
            ordered_cols = [col.strip() for col in raw_ordered_cols.split(',')]
            
            # Caso especial: Si el Excel usa "concepto" pero tu orden pide "descripcion"
            # (Ambos ya están en minúsculas por la normalización previa)
            if 'concepto' in df_new.columns and 'descripcion' in ordered_cols:
                df_new.rename(columns={'concepto': 'descripcion'}, inplace=True)
            
            # Ahora el cruce de datos es exacto: minúscula vs minúscula
            existing_cols = [col for col in ordered_cols if col in df_new.columns]
            
            df_new = df_new[existing_cols]
            print(f"Columnas reordenadas exitosamente.")

        print("\n¡Éxito! Archivo limpio generado en memoria.")
        return df_new, selected_template

    except Exception as e:
        print(f"Ocurrió un error al aplicar el formato: {e}")
        return df, None
    
def main():
    #Función principal que ejecuta el programa.
    
    df = None
    file_path = ""
    while True:
        file_path = input("Por favor, ingresa la ruta completa de tu archivo Excel: ").strip()

        # Quita comillas simples o dobles si están al inicio y fin
        if (file_path.startswith(("'", '"')) and file_path.endswith(("'", '"')) 
                and file_path[0] == file_path[-1]):
            file_path = file_path[1:-1]
        try:
            df = pd.read_excel(file_path)
            print("\n¡Archivo Excel cargado con éxito!")
            break
        except FileNotFoundError:
            print("Error: El archivo no se encontró en la ruta especificada. Inténtalo de nuevo.")
        except Exception as e:
            print(f"Ocurrió un error inesperado al leer el archivo: {e}")
            return

    while True:
        print("\n--- Menú Principal ---")
        print("1. Usar un template (Categorizar + Limpiar)")
        print("2. Categorizar registros")
        print("S. Guardar y salir")
        print("F. Salir sin guardar")
        
        choice = input("\nElige una opción (1-6): ")
        
        if choice == '1':
            # DESEMPAQUETAMOS: df recibe el DataFrame, template recibe el diccionario
            df, template = apply_format(df, file_path)
            
            if df is not None:
                # Ahora le pasamos ambos a categorize
                df = categorize(df, template)
        elif choice == '2':
            df = categorize(df, None)
        elif choice == 's' or choice == 'S':
            try:
                new_file_path = os.path.splitext(file_path)[0] + "_modificado.xlsx"
                df.to_excel(new_file_path, index=False)
                print(f"\n Archivo guardado exitosamente como '{new_file_path}'!")
                break
            except Exception as e:
                print(f"Ocurrió un error al guardar el archivo: {e}")
        elif choice == 'f' or choice == 'F':
            print("Saliendo sin guardar cambios. ¡Hasta luego!")
            break
        else:
            print("Opción no válida.")

if __name__ == "__main__":
    main()