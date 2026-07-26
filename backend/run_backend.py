#!/usr/bin/env python
"""
🚀 Backend Standalone Runner
Script autónomo para iniciar solo el Backend (FastAPI)

Características:
- Crea venv si no existe
- Instala dependencias si faltan
- Inicia Backend en puerto 8000
- Uso: python run_backend.py
"""

import subprocess
import sys
import os
import time
import socket
from pathlib import Path

class Colors:
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    RESET = '\033[0m'
    BOLD = '\033[1m'

def print_success(message):
    print(f"{Colors.GREEN}✅ {message}{Colors.RESET}")

def print_error(message):
    print(f"{Colors.RED}❌ {message}{Colors.RESET}")

def print_warning(message):
    print(f"{Colors.YELLOW}⚠️  {message}{Colors.RESET}")

def print_info(message):
    print(f"{Colors.BLUE}ℹ️  {message}{Colors.RESET}")

def run_command(cmd, cwd=None, description=""):
    """Ejecuta un comando y retorna el resultado"""
    try:
        result = subprocess.run(
            cmd,
            cwd=cwd,
            shell=True,
            capture_output=True,
            text=True,
            timeout=120
        )
        if result.returncode == 0:
            return True, result.stdout
        else:
            return False, result.stderr or result.stdout
    except subprocess.TimeoutExpired:
        return False, f"Timeout (>120s) ejecutando: {description}"
    except Exception as e:
        return False, str(e)

def setup_python_environment(backend_dir):
    """Verifica y configura el entorno virtual de Python"""
    print_info("Verificando entorno Python...")
    
    # Directorio del proyecto raíz (padre del backend)
    project_root = backend_dir.parent
    venv_path = project_root / "venv"
    
    # Verificar si venv existe
    if not venv_path.exists():
        print_warning("Entorno virtual no encontrado, creando...")
        try:
            subprocess.run(
                [sys.executable, "-m", "venv", str(venv_path)],
                check=True,
                capture_output=True
            )
            print_success("Entorno virtual creado")
        except Exception as e:
            print_error(f"Error creando venv: {str(e)}")
            return False, None
    else:
        print_success("Entorno virtual encontrado")
    
    # Determinar pip executable
    if sys.platform == "win32":
        pip_exe = venv_path / "Scripts" / "pip.exe"
        python_exe = venv_path / "Scripts" / "python.exe"
    else:
        pip_exe = venv_path / "bin" / "pip"
        python_exe = venv_path / "bin" / "python"
    
    if not pip_exe.exists():
        print_error(f"pip no encontrado en {pip_exe}")
        return False, None
    
    # Instalar dependencias del backend
    print_info("Verificando dependencias de Python...")
    requirements_file = backend_dir / "requirements.txt"
    
    if requirements_file.exists():
        success, output = run_command(
            f'"{pip_exe}" install -r "{requirements_file}"',
            description="pip install"
        )
        if success:
            print_success("Dependencias Python instaladas")
        else:
            print_error(f"Error instalando dependencias: {output[:200]}")
            return False, None
    else:
        print_warning(f"requirements.txt no encontrado en {requirements_file}")
    
    return True, str(python_exe)

def main():
    backend_dir = Path(__file__).parent
    project_root = backend_dir.parent
    
    print(f"\n{Colors.BOLD}{Colors.BLUE}{'='*60}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.BLUE}{'🚀 Backend - Standalone Runner':^60}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.BLUE}{'='*60}{Colors.RESET}\n")
    
    # Verificar estructura
    print_info("Verificando estructura del proyecto...")
    if not backend_dir.exists():
        print_error(f"Directorio backend no encontrado: {backend_dir}")
        sys.exit(1)
    print_success("Backend encontrado")
    
    # Setup Python
    success, python_exe = setup_python_environment(backend_dir)
    if not success:
        print_error("No se pudo configurar el entorno Python")
        sys.exit(1)
    
    # Iniciar Backend
    print_info("Iniciando Backend (FastAPI)...")
    print_info("Esperando a que el servidor esté listo...\n")
    
    try:
        if sys.platform == "win32":
            cmd = f'"{python_exe}" run_backend.py'
        else:
            cmd = f'"{python_exe}" run_backend.py'
        
        # Crear proceso uvicorn
        # Usar un wrapper que ejecute uvicorn directamente
        # Intentar puerto 8000, si está ocupado usar 8001
        for port in [8000, 8001, 8002, 8003]:
            test_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            result = test_sock.connect_ex(('127.0.0.1', port))
            test_sock.close()
            if result != 0:  # Puerto disponible
                uvicorn_cmd = f'"{python_exe}" -m uvicorn app.main:app --host 0.0.0.0 --port {port} --reload'
                print_success(f"Puerto {port} disponible")
                break
        else:
            uvicorn_cmd = f'"{python_exe}" -m uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload'
        
        subprocess.run(
            uvicorn_cmd,
            shell=True,
            cwd=str(backend_dir),
            check=False
        )
        
    except KeyboardInterrupt:
        print(f"\n{Colors.YELLOW}Backend detenido{Colors.RESET}")
    except Exception as e:
        print_error(f"Error iniciando Backend: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()