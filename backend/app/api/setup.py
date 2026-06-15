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

@router.post("/create-database", response_model=CreateDatabaseResponse)
async def create_database(file: UploadFile = File(...), mapping_json: str = None, db: Session = Depends(get_db)):
    """Crea BD a partir de CSV/XLSX con mapeo de columnas"""
    import json
    
    errors = []
    records_imported = 0
    
    try:
        logger.info(f"Iniciando importación de archivo: {file.filename}")
        
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            contents = await file.read()
            tmp.write(contents)
            tmp_path = tmp.name
        
        logger.info(f"Archivo guardado temporalmente en: {tmp_path}")
        
        filename = file.filename.lower()
        if filename.endswith('.csv'):
            df = FileHandler.read_csv_full(tmp_path)
            headers = list(df.columns)
        elif filename.endswith(('.xlsx', '.xls')):
            df = FileHandler.read_xlsx_full(tmp_path)
            headers = list(df.columns)
        else:
            logger.error(f"Formato de archivo no soportado: {filename}")
            return CreateDatabaseResponse(
                success=False,
                records_imported=0,
                database_path="",
                errors=["Formato de archivo no soportado"]
            )
        
        logger.info(f"Archivo parseado: {len(df)} filas, columnas: {list(df.columns)}")
        
        # Parsear mapping - usar suggested si no se envía
        mapping_dict = {}
        if mapping_json:
            try:
                mapping_dict = json.loads(mapping_json)
                logger.info(f"Mapeo recibido desde usuario: {mapping_dict}")
            except json.JSONDecodeError:
                logger.warning("No se pudo parsear mapping_json, usando sugerido")
        
        # Si no hay mapeo o está incompleto, usar el sugerido automáticamente
        if not mapping_dict or not all(k in mapping_dict for k in ['fecha', 'concepto', 'monto', 'categoria']):
            suggested = ColumnMapper.get_suggested_mapping(headers)
            logger.info(f"Usando mapeo sugerido automáticamente: {suggested}")
            if suggested:
                mapping_dict = suggested if not mapping_dict else {**suggested, **mapping_dict}
        
        logger.info(f"Mapeo final a usar: {mapping_dict}")
        
        # Validar mapeo
        is_valid, validation_errors = ColumnMapper.validate_mapping(df, mapping_dict)
        if not is_valid:
            logger.error(f"Errores de validación del mapeo: {validation_errors}")
            return CreateDatabaseResponse(
                success=False,
                records_imported=0,
                database_path="",
                errors=validation_errors
            )
        
        logger.info("Mapeo validado correctamente")
        
        # Inicializar BD
        Base.metadata.create_all(bind=engine)
        logger.info("Base de datos inicializada")
        
        # Procesar cada fila
        categories_seen = set()
        rows_processed = 0
        
        for idx, row in df.iterrows():
            rows_processed += 1
            logger.debug(f"Procesando fila {idx+1}: {row.to_dict()}")
            
            transformed = ColumnMapper.transform_row(row, mapping_dict)
            
            if transformed is None:
                logger.warning(f"Fila {idx+1}: No se pudo transformar (probablemente fecha inválida)")
                errors.append(f"Fila {idx+1}: fecha inválida o datos incompletos")
                continue
            
            logger.debug(f"Fila {idx+1} transformada: {transformed}")
            
            try:
                category_name = transformed['category']
                if category_name not in categories_seen:
                    existing_cat = db.query(Category).filter_by(name=category_name).first()
                    if not existing_cat:
                        new_cat = Category(
                            name=category_name,
                            type=transformed['type'],
                            color_hex="#666666"
                        )
                        db.add(new_cat)
                        db.flush()
                        logger.info(f"Categoría creada: {category_name}")
                    categories_seen.add(category_name)
                
                category = db.query(Category).filter_by(name=category_name).first()
                
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
                db.add(expense)
                records_imported += 1
                logger.debug(f"Expense añadido: {transformed['description']} - ${transformed['amount']}")
            
            except Exception as e:
                logger.error(f"Error procesando fila {idx+1}: {str(e)}", exc_info=True)
                errors.append(f"Fila {idx+1}: {str(e)}")
                continue
        
        logger.info(f"Filas procesadas: {rows_processed}, registros importados: {records_imported}")
        
        db.commit()
        logger.info("Cambios commitados a la BD")
        
        Path(tmp_path).unlink()
        logger.info("Archivo temporal eliminado")
        
        db_path = "data/expense.db"
        Config.set_database(
            database_path=db_path,
            source_file=file.filename,
            mapping=mapping_dict,
            records_count=records_imported,
            categories_count=len(categories_seen)
        )
        
        logger.info(f"Importación completada: {records_imported} registros, {len(categories_seen)} categorías")
        
        return CreateDatabaseResponse(
            success=True,
            records_imported=records_imported,
            database_path=db_path,
            errors=errors
        )
    
    except Exception as e:
        logger.error(f"Error fatal en create_database: {str(e)}", exc_info=True)
        return CreateDatabaseResponse(
            success=False,
            records_imported=0,
            database_path="",
            errors=[f"Error creando BD: {str(e)}"]
        )

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
