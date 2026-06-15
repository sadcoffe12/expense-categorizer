import json
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

CONFIG_FILE = Path(__file__).parent.parent.parent / "config.json"

class Config:
    """Gestiona configuración de la aplicación"""
    
    @staticmethod
    def load() -> Dict[str, Any]:
        """Carga config.json si existe"""
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}
    
    @staticmethod
    def save(config: Dict[str, Any]) -> None:
        """Guarda config.json"""
        config['last_modified'] = datetime.utcnow().isoformat()
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    
    @staticmethod
    def exists() -> bool:
        """Verifica si config existe"""
        return CONFIG_FILE.exists()
    
    @staticmethod
    def get_database_path() -> Optional[str]:
        """Obtiene ruta guardada de BD"""
        config = Config.load()
        return config.get('database_path')
    
    @staticmethod
    def set_database(database_path: str, source_file: str, mapping: Dict[str, str],
                     records_count: int, categories_count: int) -> None:
        """Guarda configuración de BD"""
        config = {
            'database_path': database_path,
            'source_file': source_file,
            'column_mapping': mapping,
            'records_count': records_count,
            'categories_count': categories_count,
            'created_at': datetime.utcnow().isoformat()
        }
        Config.save(config)
    
    @staticmethod
    def clear() -> None:
        """Elimina config.json"""
        if CONFIG_FILE.exists():
            CONFIG_FILE.unlink()
