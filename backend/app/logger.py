"""
Configuración de logging para la aplicación
"""

import logging
import logging.handlers
from pathlib import Path

# Crear directorio de logs si no existe
LOG_DIR = Path(__file__).parent.parent.parent / "logs"
LOG_DIR.mkdir(exist_ok=True)

# Crear logger principal
logger = logging.getLogger("expense_categorizer")
logger.setLevel(logging.DEBUG)

# Formato de logs
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Handler para archivo (rotación cada 5MB, mantiene 5 archivos)
file_handler = logging.handlers.RotatingFileHandler(
    LOG_DIR / "app.log",
    maxBytes=5*1024*1024,  # 5MB
    backupCount=5
)
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)

# Handler para consola
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(formatter)

# Agregar handlers al logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Logger específico para imports
import_logger = logging.getLogger("expense_categorizer.import")
import_logger.setLevel(logging.DEBUG)

def get_logger(name):
    """Obtiene un logger con el nombre especificado"""
    return logging.getLogger(f"expense_categorizer.{name}")
