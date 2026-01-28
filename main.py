import pandas as pd
import openpyxl
import os

RULES_FILE = "categorization_rules.txt"
TEMPLATES_FILE = "templates.txt"

def load_rules():
    """Carga las reglas de categorización desde el archivo de texto en orden."""
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
        print(f"❌ Error al cargar las reglas: {e}")
        return []
    return rules

def save_rule(keyword, type, category, new_description):
    """Guarda una nueva regla en el archivo de texto."""
    try:
        with open(RULES_FILE, 'a', encoding="utf-8") as f:
            f.write(f"{keyword}:{type}:{category}:{new_description if new_description else 'None'}\n")
        print(f"✅ Regla guardada: Si se encuentra '{keyword}', se asignará el tipo '{type}', la categoría '{category}' y la descripcion '{new_description}'.")
    except Exception as e:
        print(f"❌ Error al guardar la regla: {e}")

def create_rule():
    """Crea una nueva regla de categorización y la guarda."""
    print("\n--- Crear Nueva Regla de Categorización ---")
    keyword = input("Ingresa la palabra clave (ej: emova, .com): ").lower().strip()
    type = input(f"Ingresa el tipo de gasto cuando se encuentre '{keyword}' (ej: Fijo, Variable, Ingreso): ").strip()
    category = input(f"Ingresa la categoría a asignar cuando se encuentre '{keyword}' (ej: Transporte): ").strip()
    new_description = input("Ingrese la nueva descripcion a reemplazar (Dejar en blanco si la quiere mantener):").strip()
    
    if keyword and type and category:
        save_rule(keyword, type, category, new_description)
    else:
        print("❌ Error: La palabra clave, el tipo y la categoría no pueden estar vacías.")

def categorize(df):
    """Categoriza los registros basándose en reglas y palabras clave en orden secuencial."""
    print("\n--- Categorizar Registros ---")
    rules = load_rules()
    
    if not rules:
        print("⚠️ No hay reglas de categorización definidas. Crea algunas primero.")
        return df

    print("Columnas disponibles:", ", ".join(df.columns))
    source_col = input("Ingresa el nombre de la columna para leer la descripción (ej: Descripcion): ")

    # Default to "Descripcion" if left blank
    if not source_col:
        source_col = "Descripcion"
    
    if source_col not in df.columns:
        print(f"❌ Error: La columna '{source_col}' no existe.")
        return df
    
    type_col = input("Ingresa el nombre de la columna para asignar el Tipo de Gasto (ej: Tipo): ")
    
    # Default to "Tipo" if left blank
    if not type_col:
        type_col = "Tipo"

    if type_col not in df.columns:
        print(f"⚠️ La columna '{type_col}' no existe. Se creará automáticamente.")
        df[type_col] = ""
            
    category_col = input("Ingresa el nombre de la columna para asignar la categoría (ej: Categoria): ")
    
    # Default to "Categoria" if left blank
    if not category_col:
        category_col = "Categoria"
    
    if category_col not in df.columns:
        print(f"⚠️ La columna '{category_col}' no existe. Se creará automáticamente.")
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

    
    print(f"\n✅ Proceso de categorización completado. Se categorizaron {categorized_count} coincidencias.")
    return df

def create_format():
    """Crea y guarda un template de formato en un archivo de texto."""
    print("--- Creación de un nuevo template de formato ---")
    
    template_name = input("Introduce el nombre del template de formato: ")
    header_row = int(input("Número de la fila donde están los encabezados (ej. 1): "))
    start_row = int(input("Número de la fila donde inician los datos (ej. 2): "))
    start_col = input("Letra de la columna donde inician los datos (ej. B): ")
    end_col = input("Letra de la columna donde finalizan los datos (ej. G): ")

    # Preguntas para borrar y AÑADIR
    cols_to_drop = input("¿Quiere borrar columnas? Especifique las letras separadas por coma (ej. H, J) o deje en blanco: ").strip()
    rows_to_drop = input("¿Quiere borrar filas? Especifique los números de fila separados por coma (ej. 5, 10) o deje en blanco: ").strip()
    
    # --- NUEVA PREGUNTA ---
    cols_to_add = input("¿Quiere añadir columnas? Especifique en formato nombre(Letra),... (ej. Categoria(B), Sueldo(I)) o deje en blanco: ").strip()
    
    try:
        with open(TEMPLATES_FILE, "a") as f:
            f.write(f"TEMPLATE_NAME: {template_name}\n")
            f.write(f"HEADER_ROW: {header_row}\n")
            f.write(f"START_ROW: {start_row}\n")
            f.write(f"START_COL: {start_col.upper()}\n")
            f.write(f"END_COL: {end_col.upper()}\n")
            f.write(f"COLS_TO_DROP: {cols_to_drop.upper()}\n")
            f.write(f"ROWS_TO_DROP: {rows_to_drop}\n")
            # --- NUEVA LÍNEA GUARDADA ---
            f.write(f"COLS_TO_ADD: {cols_to_add}\n")
            f.write("---\n")
        print(f"✅ Template '{template_name}' guardado correctamente.")
    except Exception as e:
        print(f"❌ Error al guardar el template: {e}")

def apply_format(df, file_path):
    """
    Aplica un template para limpiar un archivo, y luego elimina/añade columnas/filas
    en el orden correcto. Devuelve solo el DataFrame en memoria (no guarda archivo).
    """
    print("--- Aplicación de un template de formato de limpieza ---")
    
    if not os.path.exists(TEMPLATES_FILE):
        print("❌ No se encontraron templates de formato. Por favor, crea uno primero.")
        return df

    # --- (La sección de lectura y selección de templates no cambia) ---
    templates = {}
    try:
        with open(TEMPLATES_FILE, "r") as f:
            content = f.read().strip()
            if not content:
                print("❌ El archivo de templates está vacío.")
                return df
            
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
                            current_template[key] = value
                        except ValueError:
                            continue
                if 'TEMPLATE_NAME' in current_template:
                    templates[current_template['TEMPLATE_NAME']] = current_template
    except Exception as e:
        print(f"❌ Error al leer los templates: {e}")
        return df

    if not templates:
        print("❌ No se encontraron templates de formato.")
        return df

    print("Templates de formato disponibles:")
    for i, name in enumerate(templates.keys()):
        print(f"  {i+1}. {name}")
    
    try:
        choice = int(input("Selecciona el número del template a aplicar: "))
        selected_name = list(templates.keys())[choice - 1]
        selected_template = templates[selected_name]
    except (ValueError, IndexError):
        print("❌ Selección no válida.")
        return df

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
            print("❌ No se encontraron datos en el rango especificado.")
            return df
            
        df_new = pd.DataFrame(data[1:], columns=data[0])
        df_new.dropna(how='all', inplace=True)

        # 3. Eliminar columnas especificadas
        if 'COLS_TO_DROP' in selected_template and selected_template['COLS_TO_DROP']:
            cols_to_drop = [col.strip() for col in selected_template['COLS_TO_DROP'].split(',') if col.strip()]
            headers_map = {openpyxl.utils.get_column_letter(start_col_idx + i): header for i, header in enumerate(data[0])}
            cols_to_drop_names = [headers_map[letter] for letter in cols_to_drop if letter in headers_map]
            cols_to_drop_names_existing = [name for name in cols_to_drop_names if name in df_new.columns]
            
            df_new.drop(columns=cols_to_drop_names_existing, inplace=True, errors='ignore')
            if cols_to_drop_names_existing:
                print(f"[OK] Columnas eliminadas: {', '.join(cols_to_drop_names_existing)}")

        # 4. Eliminar filas especificadas
        if 'ROWS_TO_DROP' in selected_template and selected_template['ROWS_TO_DROP']:
            rows_to_drop_str = [r.strip() for r in selected_template['ROWS_TO_DROP'].split(',') if r.strip()]
            rows_to_drop_indices = [int(r_str) - start_row for r_str in rows_to_drop_str]
            rows_to_drop_indices = [idx for idx in rows_to_drop_indices if 0 <= idx < len(df_new)]
            
            df_new.drop(index=rows_to_drop_indices, inplace=True, errors='ignore')
            if rows_to_drop_indices:
                print(f"[OK] Se eliminaron {len(rows_to_drop_indices)} filas especificadas.")
        
        # 5. Insertar nuevas columnas (ÚLTIMO PASO DE MODIFICACIÓN)
        if 'COLS_TO_ADD' in selected_template and selected_template['COLS_TO_ADD']:
            cols_to_add_str = selected_template['COLS_TO_ADD']
            items_to_add = [item.strip() for item in cols_to_add_str.split(',')]
            
            for item in reversed(items_to_add):
                try:
                    name, col_letter_with_paren = item.split('(')
                    col_letter = col_letter_with_paren.replace(')', '').strip().upper()
                    col_idx = openpyxl.utils.column_index_from_string(col_letter) - 1

                    df_new.insert(loc=col_idx, column=name.strip(), value='')
                    print(f"[OK] Columna '{name.strip()}' insertada en la posición {col_letter}.")
                except ValueError:
                    print(f"⚠️ Formato incorrecto para añadir columna: '{item}'. Se omitirá.")
                    continue
        
        # 🚫 Ya no guardamos a archivo, solo devolvemos el DataFrame
        print("\n[OK] ¡Éxito! Archivo limpio generado en memoria.")
        return df_new

    except Exception as e:
        print(f"❌ Ocurrió un error al aplicar el formato: {e}")
        return df

def main():
    """Función principal que ejecuta el programa."""
    
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
            print("\n✅ ¡Archivo Excel cargado con éxito!")
            break
        except FileNotFoundError:
            print("❌ Error: El archivo no se encontró en la ruta especificada. Inténtalo de nuevo.")
        except Exception as e:
            print(f"❌ Ocurrió un error inesperado al leer el archivo: {e}")
            return

    while True:
        print("\n--- Menú Principal ---")
        print("1. Aplicar formato")
        print("2. Crear formato personalizado")
        print("3. Categorizar registros")
        print("4. Crear nueva regla de categorización")
        print("5. Guardar y salir")
        print("6. Salir sin guardar")
        
        choice = input("\nElige una opción (1-6): ")
        
        if choice == '1':
            df = apply_format(df, file_path)
        elif choice == '2':
            create_format()
        elif choice == '3':
            df = categorize(df)
        elif choice == '4':
            create_rule() 
        elif choice == '5':
            try:
                new_file_path = os.path.splitext(file_path)[0] + "_modificado.xlsx"
                df.to_excel(new_file_path, index=False)
                print(f"\n✅ ¡Archivo guardado exitosamente como '{new_file_path}'!")
                break
            except Exception as e:
                print(f"❌ Ocurrió un error al guardar el archivo: {e}")
        elif choice == '6':
            print("Saliendo sin guardar cambios. ¡Hasta luego!")
            break
        else:
            print("Opción no válida. Por favor, elige un número del 1 al 6.")

if __name__ == "__main__":
    main()