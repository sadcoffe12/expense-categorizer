#!/usr/bin/env python3
"""
🚀 Expense Categorizer - One-Click Starter
Inicializa dependencias e inicia todos los servicios (Backend + Frontend)

Características:
- Verifica e instala dependencias automáticamente
- Crea entorno virtual de Python si no existe
- Instala paquetes de npm y Python
- Inicia Backend (FastAPI) en puerto 8000
- Inicia Frontend (Vite) en puerto 5173
- Manejo multiplataforma (Windows/Linux/Mac)
- Limpieza automática de procesos con Ctrl+C
"""

import subprocess
import sys
import os
import time
import socket
import atexit
import signal
import shutil
import glob
from pathlib import Path

# Colores para terminal
class Colors:
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    RESET = '\033[0m'
    BOLD = '\033[1m'

# Variables globales para cleanup
RUNNING_PROCESSES = []

def print_header(message):
    """Imprime encabezado"""
    print(f"\n{Colors.BOLD}{Colors.BLUE}{'='*60}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.BLUE}{message:^60}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.BLUE}{'='*60}{Colors.RESET}\n")

def print_success(message):
    """Imprime mensaje de éxito"""
    print(f"{Colors.GREEN}✅ {message}{Colors.RESET}")

def print_error(message):
    """Imprime mensaje de error"""
    print(f"{Colors.RED}❌ {message}{Colors.RESET}")

def print_warning(message):
    """Imprime mensaje de advertencia"""
    print(f"{Colors.YELLOW}⚠️  {message}{Colors.RESET}")

def print_info(message):
    """Imprime mensaje de información"""
    print(f"{Colors.BLUE}ℹ️  {message}{Colors.RESET}")

def cleanup_processes():
    """Limpia todos los procesos hijo en el exit"""
    global RUNNING_PROCESSES
    if RUNNING_PROCESSES:
        print_info("\nDeteniendo procesos...")
        for proc in RUNNING_PROCESSES:
            try:
                if proc.poll() is None:  # Si aún está corriendo
                    if sys.platform == "win32":
                        proc.terminate()
                    else:
                        os.killpg(os.getpgid(proc.pid), signal.SIGTERM)
                    time.sleep(0.5)
                    if proc.poll() is None:
                        proc.kill()
            except:
                pass

def signal_handler(sig, frame):
    """Manejador para Ctrl+C"""
    cleanup_processes()
    print_warning("\n\n¡Servicios detenidos!")
    sys.exit(0)

# Registrar cleanup handlers
atexit.register(cleanup_processes)
signal.signal(signal.SIGINT, signal_handler)

def find_npm():
    """Busca npm en el sistema de forma confiable"""
    
    # 1. Primero intenta con shutil (más confiable)
    npm_path = shutil.which("npm")
    if npm_path:
        print_info(f"npm encontrado (shutil): {npm_path}")
        return npm_path
    
    # 2. Intenta con comandos del sistema
    try:
        if sys.platform == "win32":
            result = subprocess.run(
                "where npm",
                shell=True, capture_output=True, text=True, timeout=2
            )
        else:
            result = subprocess.run(
                "which npm",
                shell=True, capture_output=True, text=True, timeout=2
            )
        if result.stdout.strip():
            path = result.stdout.strip().split('\n')[0]
            if os.path.exists(path):
                print_info(f"npm encontrado (which): {path}")
                return path
    except:
        pass
    
    # 3. Busca en rutas comunes (especialmente útil para NVM, Homebrew, etc)
    common_paths = [
        # macOS Homebrew (Intel)
        "/usr/local/bin/npm",
        # macOS Homebrew (Apple Silicon)
        "/opt/homebrew/bin/npm",
        # Linux estándar
        "/usr/bin/npm",
        "/bin/npm",
        # NVM - primero expandir ~ a la ruta del usuario
        os.path.expanduser("~/.nvm/versions/node/*/bin/npm"),
        # Local node_modules (último recurso)
        os.path.expanduser("./node_modules/.bin/npm"),
    ]
    
    for path_pattern in common_paths:
        try:
            if "*" in path_pattern:
                # Expandir glob patterns
                matches = glob.glob(path_pattern)
                for match in sorted(matches, reverse=True):  # Usar última versión de Node
                    if os.path.exists(match) and os.path.isfile(match):
                        print_info(f"npm encontrado (glob): {match}")
                        return match
            else:
                # Ruta exacta
                if os.path.exists(path_pattern) and os.path.isfile(path_pattern):
                    print_info(f"npm encontrado (path): {path_pattern}")
                    return path_pattern
        except Exception as e:
            # Ignorar errores en la búsqueda
            continue
    
    # 4. Último intento: si npm no está en PATH pero existe Node.js
    # Retornar "npm" y dejar que el shell lo resuelva
    try:
        result = subprocess.run(
            "npm --version",
            shell=True, capture_output=True, text=True, timeout=2
        )
        if result.returncode == 0:
            print_info("npm encontrado a través del shell")
            return "npm"
    except:
        pass
    
    return None

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

def check_port_in_use(port):
    """Verifica si un puerto está en uso"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        result = sock.connect_ex(('127.0.0.1', port))
        return result == 0
    finally:
        sock.close()

def setup_python_environment(project_root):
    """Verifica y configura el entorno virtual de Python"""
    print_info("Verificando entorno Python...")
    
    venv_path = project_root / "venv"
    backend_dir = project_root / "backend"
    
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
            return False
    else:
        print_success("Entorno virtual encontrado")
    
    # Determinar pip executable
    if sys.platform == "win32":
        pip_exe = venv_path / "Scripts" / "pip.exe"
    else:
        pip_exe = venv_path / "bin" / "pip"
    
    if not pip_exe.exists():
        print_error(f"pip no encontrado en {pip_exe}")
        return False
    
    # Instalar dependencias del backend
    print_info("Instalando dependencias de Python...")
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
            return False
    else:
        print_warning(f"requirements.txt no encontrado en {requirements_file}")
    
    return True

def setup_frontend_environment(project_root):
    """Verifica e instala dependencias del frontend"""
    print_info("Verificando entorno Node.js...")
    
    frontend_dir = project_root / "frontend"
    npm = find_npm()
    
    if not npm:
        print_error("npm no encontrado")
        print_warning("Rutas buscadas:")
        print_warning("  - /usr/local/bin/npm")
        print_warning("  - /opt/homebrew/bin/npm")
        print_warning("  - /usr/bin/npm")
        print_warning("  - ~/.nvm/versions/node/*/bin/npm")
        print_error("Por favor instala Node.js desde https://nodejs.org")
        print_error("O si usas NVM: nvm install node && nvm use node")
        return False
    
    print_success(f"npm listo")
    
    # Verificar si node_modules existe
    node_modules = frontend_dir / "node_modules"
    if not node_modules.exists():
        print_warning("node_modules no encontrado, instalando dependencias...")
        success, output = run_command(
            f'"{npm}" install',
            cwd=str(frontend_dir),
            description="npm install"
        )
        if success:
            print_success("Dependencias Node.js instaladas")
        else:
            print_error(f"Error instalando npm packages: {output[:200]}")
            return False
    else:
        print_success("node_modules encontrado")
    
    return True

def start_backend(project_root):
    """Inicia el backend"""
    print_info("Iniciando Backend (FastAPI)...")
    backend_dir = project_root / "backend"
    
    # Preparar comando para ejecutar con el venv
    if sys.platform == "win32":
        python_exe = project_root / "venv" / "Scripts" / "python.exe"
        cmd = f'"{python_exe}" run_backend.py'
    else:
        python_exe = project_root / "venv" / "bin" / "python"
        cmd = f'"{python_exe}" run_backend.py'
    
    try:
        if sys.platform == "win32":
            process = subprocess.Popen(
                cmd,
                cwd=str(backend_dir),
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
        else:
            # Unix: usar preexec_fn para grupo de procesos
            process = subprocess.Popen(
                cmd,
                cwd=str(backend_dir),
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                preexec_fn=os.setsid
            )
        
        RUNNING_PROCESSES.append(process)
        
        # Esperar a que el puerto esté disponible
        print_info("Esperando Backend (max 20s)...")
        for attempt in range(20):
            time.sleep(1)
            if check_port_in_use(8000):
                print_success(f"Backend iniciado (PID: {process.pid})")
                print_info("  📍 http://localhost:8000")
                print_info("  📖 Docs: http://localhost:8000/docs")
                return True
        
        print_error("Backend no respondió después de 20 segundos")
        return False
        
    except Exception as e:
        print_error(f"Error iniciando Backend: {str(e)}")
        return False

def start_frontend(project_root):
    """Inicia el frontend"""
    print_info("Iniciando Frontend (Vite)...")
    frontend_dir = project_root / "frontend"
    npm = find_npm()
    
    if not npm:
        print_error("npm no encontrado")
        return False
    
    try:
        # Preparar el comando
        npm_dev_cmd = f'"{npm}" run dev'
        
        if sys.platform == "win32":
            process = subprocess.Popen(
                npm_dev_cmd,
                cwd=str(frontend_dir),
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1
            )
        else:
            # Unix/Linux/Mac: usar bash -i para cargar NVM y otros profiles
            # -i = interactive, carga ~/.bashrc, ~/.zshrc, etc
            process = subprocess.Popen(
                f'bash -i -c "{npm_dev_cmd}"',
                cwd=str(frontend_dir),
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                preexec_fn=os.setsid
            )
        
        RUNNING_PROCESSES.append(process)
        
        # Esperar a que el puerto esté disponible
        print_info("Esperando Frontend (max 30s)...")
        
        port_ready = False
        start_time = time.time()
        
        for attempt in range(30):
            time.sleep(1)
            
            # Revisar si el proceso terminó con error
            poll_result = process.poll()
            if poll_result is not None:
                # El proceso terminó
                print_error(f"Proceso Vite terminó con código: {poll_result}")
                # Leer el output del error
                try:
                    output = process.stdout.read()
                    if output:
                        print_error("Output del proceso:")
                        print(output[:500])
                except:
                    pass
                return False
            
            # Verificar si el puerto está listo
            if check_port_in_use(5173):
                elapsed = time.time() - start_time
                print_success(f"Frontend iniciado (PID: {process.pid}, {elapsed:.1f}s)")
                print_info("  🌐 http://localhost:5173")
                port_ready = True
                break
            
            # Mostrar progreso cada 5 segundos
            if (attempt + 1) % 5 == 0:
                elapsed = time.time() - start_time
                print_info(f"  Esperando... {elapsed:.0f}s")
        
        if port_ready:
            return True
        
        # Si llegamos aquí, el puerto no se abrió
        print_error("Frontend no respondió después de 30 segundos")
        
        # Intentar capturar el output para debugging
        if not process.poll():
            print_warning("\nIntentando capturar salida del proceso...")
            try:
                # Non-blocking read del stdout
                if sys.platform != "win32":
                    import fcntl
                    import os as os_module
                    
                    flags = fcntl.fcntl(process.stdout, fcntl.F_GETFL)
                    fcntl.fcntl(process.stdout, fcntl.F_SETFL, flags | os_module.O_NONBLOCK)
                    
                    try:
                        output = process.stdout.read(1000)
                        if output:
                            print_error("Output del proceso:")
                            print(output)
                    except BlockingIOError:
                        print_warning("No hay output disponible")
            except:
                print_warning("No se pudo leer output (Windows/diferente SO)")
        
        return False
        
    except Exception as e:
        print_error(f"Error iniciando Frontend: {str(e)}")
        return False

def main():
    """Función principal - One-click starter"""
    project_root = Path(__file__).parent
    
    print_header("🚀 Expense Categorizer - Setup & Start")
    
    # 1. Verificar estructura del proyecto
    print_info("Verificando estructura del proyecto...")
    backend_dir = project_root / "backend"
    frontend_dir = project_root / "frontend"
    
    if not backend_dir.exists() or not frontend_dir.exists():
        print_error("Estructura del proyecto incompleta")
        print_error(f"Backend: {'✓' if backend_dir.exists() else '✗'}")
        print_error(f"Frontend: {'✓' if frontend_dir.exists() else '✗'}")
        sys.exit(1)
    
    print_success("Estructura del proyecto OK")
    
    # 2. Configurar entorno Python
    print_header("📦 Configurar Python")
    if not setup_python_environment(project_root):
        print_error("No se pudo configurar el entorno Python")
        sys.exit(1)
    
    # 3. Configurar entorno Node.js
    print_header("📦 Configurar Node.js")
    if not setup_frontend_environment(project_root):
        print_error("No se pudo configurar el entorno Node.js")
        sys.exit(1)
    
    # 4. Iniciar servicios
    print_header("🚀 Iniciando Servicios")
    
    if not start_backend(project_root):
        print_error("No se pudo iniciar el Backend")
        cleanup_processes()
        sys.exit(1)
    
    time.sleep(1)  # Pequeña pausa entre servicios
    
    if not start_frontend(project_root):
        print_error("No se pudo iniciar el Frontend")
        cleanup_processes()
        sys.exit(1)
    
    # 5. Éxito - servicios corriendo
    print_header("✨ ¡Servicios Listos!")
    print_success("Backend (FastAPI)  → http://localhost:8000")
    print_success("Frontend (Vite)    → http://localhost:5173")
    print_success("API Docs           → http://localhost:8000/docs")
    print_info("\n💡 Presiona Ctrl+C para detener todos los servicios\n")
    
    # Mantener el script corriendo
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass

if __name__ == "__main__":
    main()
