#!/usr/bin/env python
"""Script para ejecutar el backend FastAPI"""

import uvicorn
import sys
from pathlib import Path

if __name__ == "__main__":
    backend_dir = Path(__file__).parent
    sys.path.insert(0, str(backend_dir))
    
    print("🚀 Iniciando Expense Categorizer API...")
    print("📍 http://localhost:8000")
    print("📖 Docs: http://localhost:8000/docs")
    print("")
    print("Presiona Ctrl+C para detener")
    
    uvicorn.run(
        "app.main:app",
        host="0.0.0.0",
        port=8000,
        reload=True
    )
