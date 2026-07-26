from fastapi import APIRouter, UploadFile, File, HTTPException, Depends
from sqlalchemy.orm import Session
import shutil
import tempfile
import sqlite3
from pathlib import Path
import pandas as pd
from typing import List, Dict

from ..database import get_db, SessionLocal, init_db, Base, engine
from ..schemas import (
    ValidateSQLResponse, ParseFileResponse,
    CreateDatabaseResponse, ColumnMapping
)
from ..config import Config
from ..models import Category, Expense
from ..utils.file_handler import FileHandler, TextUtils
from ..utils.column_mapper import ColumnMapper
from ..logger import get_logger

logger = get_logger("setup")

router = APIRouter(prefix="/api/setup", tags=["setup"])

@router.post("/validate-sql", response_model=ValidateSQLResponse)
async def validate_sql(file: UploadFile = File(...)):
    """Valida que un archivo SQL sea válido"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".db") as tmp:
            contents = await file.read()
            tmp.write(contents)
            tmp_path = tmp.name
        
        result = FileHandler.validate_sql(tmp_path)
        Path(tmp_path).unlink()
        
        return ValidateSQLResponse(**result)
    
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error validando SQL: {str(e)}")

@router.post("/parse-file", response_model=ParseFileResponse)
async def parse_file(file: UploadFile = File(...)):
    """Parsea CSV/XLSX y devuelve headers + preview + validaciones detalladas"""
    try:
        logger.info(f"Parseando archivo: {file.filename}")
        
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            contents = await file.read()
            tmp.write(contents)
            tmp_path = tmp.name
        
        filename = file.filename.lower()
        
        # Parsear el archivo
        if filename.endswith('.csv'):
            headers, preview = FileHandler.parse_csv(tmp_path)
            df = FileHandler.read_csv_full(tmp_path)
            logger.info(f"CSV parseado: {len(df)} filas, {len(headers)} columnas")
        elif filename.endswith(('.xlsx', '.xls')):
            headers, preview = FileHandler.parse_xlsx(tmp_path)
            df = FileHandler.read_xlsx_full(tmp_path)
            logger.info(f"XLSX parseado: {len(df)} filas, {len(headers)} columnas")
        else:
            logger.error(f"Formato no soportado: {filename}")
            raise ValueError("Formato de archivo no soportado. Use CSV o XLSX")
        
        # Sugerir mapeo automático
        suggested_mapping = ColumnMapper.get_suggested_mapping(headers)
        logger.info(f"Mapeo sugerido: {suggested_mapping}")
        
        # Validación completa con diagnósticos
        validation_result = ColumnMapper.validate_with_diagnostics(df, suggested_mapping)
        logger.info(f"Validación completada - válido: {validation_result.is_valid}, issues: {len(validation_result.issues)}")
        
        Path(tmp_path).unlink()
        
        return ParseFileResponse(
            headers=headers,
            preview=preview,
            row_count=len(df),
            suggested_mapping=suggested_mapping,
            validation_result=validation_result.to_dict()
        )
    
    except Exception as e:
        logger.error(f"Error parseando archivo: {str(e)}", exc_info=True)
        raise HTTPException(status_code=400, detail=f"Error parseando archivo: {str(e)}")

@router.get("/check-database")
async def check_database():
    """Verifica estado de la BD según config.json"""
    try:
        from ..database import validate_database
        
        # Leer config
        config = Config.load()
        db_path = config.get('database_path', '')
        
        # Si no hay BD configurada
        if not db_path:
            logger.info("No hay BD configurada en config.json")
            return {
                "status": "no_database",
                "message": "No hay BD configurada. Por favor, importa datos.",
                "database_path": "",
                "records_count": 0,
                "categories_count": 0
            }
        
        # Validar que BD existe y es válida
        is_valid, error_msg = validate_database(db_path)
        
        if is_valid:
            logger.info(f"BD válida: {db_path}")
            return {
                "status": "database_exists",
                "message": "BD encontrada y válida",
                "database_path": db_path,
                "records_count": config.get('records_count', 0),
                "categories_count": config.get('categories_count', 0)
            }
        else:
            logger.error(f"BD inválida: {error_msg}")
            return {
                "status": "database_invalid",
                "message": error_msg,
                "database_path": db_path,
                "records_count": 0,
                "categories_count": 0
            }
    
    except Exception as e:
        logger.error(f"Error checkeando BD: {str(e)}")
        return {
            "status": "error",
            "message": f"Error: {str(e)}",
            "database_path": "",
            "records_count": 0,
            "categories_count": 0
        }

@router.get("/check-existing-database")
async def check_existing_database():
    """Verifica si ya existe una BD expense.db (para el diálogo de sobrescritura)"""
    try:
        db_path = Path("data/expense.db")
        
        if db_path.exists():
            from ..database import validate_database
            is_valid, error_msg = validate_database(str(db_path))
            
            if is_valid:
                # Contar cuántos registros hay
                import sqlite3
                conn = sqlite3.connect(str(db_path))
                cursor = conn.cursor()
                cursor.execute("SELECT COUNT(*) FROM expenses")
                count = cursor.fetchone()[0]
                conn.close()
                
                return {
                    "exists": True,
                    "valid": True,
                    "record_count": count
                }
            else:
                return {
                    "exists": True,
                    "valid": False,
                    "error": error_msg
                }
        else:
            return {
                "exists": False,
                "valid": False
            }
    
    except Exception as e:
        logger.error(f"Error checking existing database: {str(e)}")
        return {
            "exists": False,
            "valid": False,
            "error": str(e)
        }

@router.post("/create-database", response_model=CreateDatabaseResponse)
async def create_database(
    file: UploadFile = File(...), 
    mapping_json: str = None, 
    recreate: bool = True,
    db: Session = Depends(get_db)
):
    """Crea BD a partir de CSV/XLSX con mapeo de columnas
    
    Parámetros:
    - file: Archivo CSV/XLSX a importar
    - mapping_json: Mapeo de columnas (JSON)
    - recreate: Si True, borra BD vieja y crea nueva. Si False, agrega a BD existente
    """
    import json
    
    errors = []
    records_imported = 0
    db_path = "data/expense.db"
    new_db = None
    
    try:
        logger.info(f"Iniciando importación de archivo: {file.filename} (recreate={recreate})")
        
        # PASO 1: MANEJAR BD ANTERIOR
        try:
            # Cerrar conexiones
            db.close()
            engine.dispose()
            logger.info("Conexiones cerradas")
            
            db_file = Path(db_path)
            
            # Si recreate=True, hacer backup y borrar
            if recreate and db_file.exists():
                backup_path = Path(str(db_file) + ".backup")
                shutil.copy2(db_file, backup_path)
                logger.info(f"Backup creado: {backup_path}")
                
                db_file.unlink()
                logger.info(f"BD anterior eliminada para recrear")
            
            # Crear directorio data
            db_file.parent.mkdir(parents=True, exist_ok=True)
            
        except Exception as cleanup_err:
            logger.warning(f"Error durante limpieza: {str(cleanup_err)}")
        
        # PASO 2: PREPARAR BD
        if recreate:
            logger.info("Recreando BD limpia...")
            Base.metadata.drop_all(bind=engine)
            Base.metadata.create_all(bind=engine)
        else:
            logger.info("Usando BD existente...")
            Base.metadata.create_all(bind=engine)
        
        logger.info("BD lista para inserciones")
        
        # PASO 3: CREAR NUEVA SESIÓN
        new_db = SessionLocal()
        logger.info("Nueva sesión creada")
        
        # PASO 4: PARSEAR ARCHIVO
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            contents = await file.read()
            tmp.write(contents)
            tmp_path = tmp.name
        
        filename = file.filename.lower()
        if filename.endswith('.csv'):
            df = FileHandler.read_csv_full(tmp_path)
        elif filename.endswith(('.xlsx', '.xls')):
            df = FileHandler.read_xlsx_full(tmp_path)
        else:
            return CreateDatabaseResponse(
                success=False,
                records_imported=0,
                database_path="",
                errors=["Formato de archivo no soportado. Use CSV o XLSX"]
            )
        
        headers = list(df.columns)
        logger.info(f"Archivo parseado: {len(df)} filas")
        
        # PASO 5: MAPEO DE COLUMNAS
        mapping_dict = {}
        if mapping_json:
            try:
                mapping_dict = json.loads(mapping_json)
                logger.info(f"Mapeo recibido del usuario")
            except json.JSONDecodeError:
                logger.warning("No se pudo parsear mapeo, usando sugerido")
        
        if not mapping_dict or not all(k in mapping_dict for k in ['fecha', 'concepto', 'monto', 'categoria']):
            suggested = ColumnMapper.get_suggested_mapping(headers)
            logger.info(f"Usando mapeo sugerido automáticamente")
            if suggested:
                mapping_dict = suggested if not mapping_dict else {**suggested, **mapping_dict}
        
        # Validar mapeo
        is_valid, validation_errors = ColumnMapper.validate_mapping(df, mapping_dict)
        if not is_valid:
            return CreateDatabaseResponse(
                success=False,
                records_imported=0,
                database_path="",
                errors=validation_errors
            )
        
        logger.info("Mapeo validado")
        
        # PASO 6: IMPORTAR DATOS
        categories_seen = set()
        
        for idx, row in df.iterrows():
            transformed = ColumnMapper.transform_row(row, mapping_dict)
            
            if transformed is None:
                errors.append(f"Fila {idx+1}: datos inválidos")
                logger.debug(f"Fila {idx+1} rechazada")
                continue
            
            try:
                category_name = transformed['category']
                
                # Crear categoría si no existe
                if category_name not in categories_seen:
                    existing_cat = new_db.query(Category).filter_by(name=category_name).first()
                    if not existing_cat:
                        new_cat = Category(
                            name=category_name,
                            type=transformed['type'],
                            color_hex="#666666"
                        )
                        new_db.add(new_cat)
                        new_db.flush()
                        logger.debug(f"Categoría creada: {category_name}")
                    categories_seen.add(category_name)
                
                # Agregar gasto
                category = new_db.query(Category).filter_by(name=category_name).first()
                expense = Expense(
                    date=transformed['date'],
                    description=transformed['description'],
                    description_cleaned=transformed['description_cleaned'],
                    amount=transformed['amount'],
                    category_id=category.id,
                    type=transformed['type'],
                    location=transformed['location'],
                    notes=transformed['notes'],
                    source=transformed['source']
                )
                new_db.add(expense)
                records_imported += 1
                
            except Exception as e:
                errors.append(f"Fila {idx+1}: {str(e)}")
                logger.debug(f"Error fila {idx+1}: {str(e)}")
                continue
        
        logger.info(f"Importación: {records_imported} registros importados, {len(categories_seen)} categorías")
        
        # PASO 7: GUARDAR CAMBIOS
        new_db.commit()
        logger.info("Cambios guardados en BD")
        
        # PASO 8: CREAR CONFIG.JSON
        Config.set_database(
            database_path=db_path,
            mapping=mapping_dict,
            records_count=records_imported,
            categories_count=len(categories_seen)
        )
        logger.info("Config.json creado/actualizado")
        
        # Limpiar archivo temporal
        try:
            Path(tmp_path).unlink()
        except:
            pass
        
        return CreateDatabaseResponse(
            success=records_imported > 0,
            records_imported=records_imported,
            database_path=db_path,
            errors=errors
        )
    
    except Exception as e:
        logger.error(f"Error fatal: {str(e)}", exc_info=True)
        return CreateDatabaseResponse(
            success=False,
            records_imported=0,
            database_path="",
            errors=[f"Error: {str(e)}"]
        )
    
    finally:
        if new_db:
            try:
                new_db.close()
            except:
                pass

@router.get("/load-database")
async def load_database():
    """Carga BD existente desde config.json"""
    try:
        if not Config.exists():
            return {
                "loaded": False,
                "database_info": {}
            }
        
        config = Config.load()
        db_info = {
            'path': config.get('database_path'),
            'records': config.get('records_count', 0),
            'categories': config.get('categories_count', 0),
            'created_at': config.get('created_at')
        }
        
        return {
            "loaded": True,
            "database_info": db_info
        }
    
    except Exception as e:
        return {
            "loaded": False,
            "database_info": {"error": str(e)}
        }
