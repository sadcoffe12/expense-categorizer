from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker
from sqlalchemy.pool import StaticPool
from pathlib import Path
import sqlite3

DATABASE_URL = "sqlite:///./data/expense.db"

engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False},
    poolclass=StaticPool,
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def init_db():
    """Crea todas las tablas si no existen"""
    Base.metadata.create_all(bind=engine)

def validate_database(db_path: str) -> tuple[bool, str]:
    """
    Valida que un archivo sea una BD SQLite válida
    Retorna: (is_valid, error_message)
    """
    try:
        db_file = Path(db_path)
        
        # Verificar que archivo existe
        if not db_file.exists():
            return False, f"Archivo no encontrado: {db_path}"
        
        # Intentar conectar
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Verificar estructura básica
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = cursor.fetchall()
        
        conn.close()
        
        # Si hay tablas, probablemente es válida
        if tables:
            return True, ""
        
        # BD vacía pero válida
        return True, ""
    
    except sqlite3.DatabaseError:
        return False, f"BD corrupta o inválida: {db_path}"
    except Exception as e:
        return False, f"Error validando BD: {str(e)}"
