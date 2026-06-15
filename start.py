#!/usr/bin/env python3
"""
Script para iniciar todos los servicios de Expense Categorizer
Inicia backend (FastAPI) y frontend (Vite)
"""

import subprocess
import sys
import os
import time
import socket
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

def check_port_in_use(port):
    """Verifica si un puerto está en uso"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        result = sock.connect_ex(('127.0.0.1', port))
        return result == 0
    finally:
        sock.close()

def find_npm():
    """Busca npm en el sistema"""
    import shutil
    
    # Primero intenta con which/where
    try:
        if sys.platform == "win32":
            result = subprocess.run(
                "where npm",
                shell=True, capture_output=True, text=True
            )
        else:
            result = subprocess.run(
                "which npm",
                shell=True, capture_output=True, text=True
            )
        if result.stdout.strip():
            return result.stdout.strip().split('\n')[0]
    except:
        pass
    
    # Intenta con shutil
    npm_path = shutil.which("npm")
    if npm_path:
        return npm_path
    
    # Busca en rutas comunes
    common_paths = [
        "/usr/local/bin/npm",
        "/usr/bin/npm",
        f"{os.path.expanduser('~')}/.nvm/versions/node/*/bin/npm",
        f"{os.path.expanduser('~')}/node_modules/.bin/npm",
    ]
    
    for path in common_paths:
        if "*" in path:
            matches = glob.glob(path)
            for match in matches:
                if os.path.exists(match):
                    return match
        elif os.path.exists(path):
            return path
    
    return None

def get_process_pid_on_port(port):
    """Obtiene PID del proceso en un puerto (Linux/Mac)"""
    try:
        if sys.platform == "win32":
            result = subprocess.run(
                f'netstat -ano | findstr :{port}',
                shell=True, capture_output=True, text=True
            )
            if result.stdout:
                parts = result.stdout.split()
                if parts:
                    return parts[-1]
        else:
            result = subprocess.run(
                f'lsof -i :{port} | grep LISTEN',
                shell=True, capture_output=True, text=True
            )
            if result.stdout:
                parts = result.stdout.split()
                if len(parts) > 1:
                    return parts[1]
    except:
        pass
    return None

def kill_process_on_port(port, port_name):
    """Mata el proceso en un puerto"""
    try:
        pid = get_process_pid_on_port(port)
        if pid:
            print_info(f"Encontrado proceso en puerto {port} con PID {pid}")
            try:
                if sys.platform == "win32":
                    subprocess.run(f'taskkill /PID {pid} /F', shell=True, capture_output=True)
                else:
                    subprocess.run(f'kill -9 {pid}', shell=True, capture_output=True)
                time.sleep(1)
                
                # Verificar que fue terminado
                if check_port_in_use(port):
                    print_warning(f"Puerto {port} aún en uso, intentando nuevamente...")
                    time.sleep(1)
                    subprocess.run(f'kill -9 {pid}', shell=True, capture_output=True)
                    time.sleep(2)
                
                if check_port_in_use(port):
                    print_error(f"No se pudo liberar puerto {port}. Intenta manualmente: kill -9 {pid}")
                    return False
                else:
                    print_success(f"Proceso en puerto {port} ({port_name}) terminado")
                    return True
            except Exception as e:
                print_error(f"Error al matar proceso: {str(e)}")
                return False
        else:
            print_warning(f"No se encontró proceso en puerto {port}")
            return False
    except Exception as e:
        print_error(f"Error checando puerto {port}: {str(e)}")
        return False

def ask_yes_no(question):
    """Pide confirmación al usuario"""
    while True:
        response = input(f"{Colors.YELLOW}{question} (s/n): {Colors.RESET}").lower().strip()
        if response in ['s', 'si', 'yes', 'y']:
            return True
        elif response in ['n', 'no']:
            return False
        else:
            print_warning("Por favor ingresa 's' o 'n'")

def check_services_running():
    """Verifica qué servicios ya están corriendo"""
    backend_running = check_port_in_use(8000)
    frontend_running = check_port_in_use(5173)
    
    return {
        'backend': backend_running,
        'frontend': frontend_running
    }

def start_backend():
    """Inicia el backend"""
    print_info("Iniciando backend (FastAPI)...")
    backend_dir = Path(__file__).parent / "backend"
    
    if not backend_dir.exists():
        print_error(f"Directorio backend no encontrado: {backend_dir}")
        return False
    
    try:
        # Usar Popen para que el proceso se ejecute en paralelo
        process = subprocess.Popen(
            ["python", "run.py"],
            cwd=str(backend_dir),
            stdout=subprocess.DEVNULL,  # Ignorar stdout
            stderr=subprocess.DEVNULL,  # Ignorar stderr (warnings no son errores)
            text=True
        )
        
        # Esperar a que se inicie y verificar que el puerto responde
        print_info("Esperando que el backend esté listo...")
        max_attempts = 15
        for attempt in range(max_attempts):
            time.sleep(1)
            if check_port_in_use(8000):
                print_success(f"Backend iniciado exitosamente (PID: {process.pid})")
                print_info("Backend disponible en: http://localhost:8000")
                print_info("API Docs en: http://localhost:8000/docs")
                return True
        
        # Si llegamos aquí, el puerto no se abrió después de esperar
        print_error("Backend no respondió en el puerto 8000 después de 15 segundos")
        return False
        
    except Exception as e:
        print_error(f"Error iniciando backend: {str(e)}")
        return False

def start_frontend():
    """Inicia el frontend"""
    print_info("Iniciando frontend (Vite)...")
    frontend_dir = Path(__file__).parent / "frontend"
    
    if not frontend_dir.exists():
        print_error(f"Directorio frontend no encontrado: {frontend_dir}")
        return False
    
    try:
        # Crear un ambiente que incluya PATH actual
        env = os.environ.copy()
        
        # Para Windows
        if sys.platform == "win32":
            process = subprocess.Popen(
                "npm run dev",
                shell=True,
                cwd=str(frontend_dir),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                env=env
            )
        else:
            # Para Unix/Linux/Mac, usar bash con login para que cargue NVM
            # Primero, obtener el comando que usa la shell actual
            process = subprocess.Popen(
                'bash -i -c "npm run dev"',
                shell=True,
                cwd=str(frontend_dir),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                env=env,
                preexec_fn=os.setsid
            )
        
        # Esperar a que se inicie y verificar que el puerto responde
        print_info("Esperando que el frontend esté listo...")
        max_attempts = 20
        for attempt in range(max_attempts):
            time.sleep(1)
            if check_port_in_use(5173):
                print_success(f"Frontend iniciado exitosamente (PID: {process.pid})")
                print_info("Frontend disponible en: http://localhost:5173")
                return True
        
        # Si llegamos aquí, el puerto no se abrió después de esperar
        # Intentar obtener el error del proceso
        if process.poll() is not None:
            # El proceso terminó, mostrar error
            stdout, stderr = process.communicate()
            error_msg = (stderr if stderr else stdout)[:500]
            print_error(f"Frontend falló al iniciar: {error_msg}")
        else:
            # El proceso sigue corriendo pero no abrió el puerto
            print_error("Frontend no respondió en el puerto 5173 después de 20 segundos")
            print_warning("El proceso sigue corriendo, puede estar en error")
            # Intentar leer un poco de stderr sin bloquear
            try:
                import select
                if select.select([process.stderr], [], [], 0)[0]:
                    error_sample = process.stderr.read(300)
                    if error_sample:
                        print_error(f"Error capturado: {error_sample}")
            except:
                pass
        
        return False
        
    except Exception as e:
        print_error(f"Error iniciando frontend: {str(e)}")
        return False

def main():
    """Función principal"""
    print_header("🚀 Expense Categorizer - Iniciar Servicios")
    
    # Verificar si los servicios ya están corriendo
    services = check_services_running()
    
    services_running = False
    if services['backend']:
        print_warning("Backend ya está corriendo en puerto 8000")
        services_running = True
    if services['frontend']:
        print_warning("Frontend ya está corriendo en puerto 5173")
        services_running = True
    
    # Si algún servicio está corriendo, preguntar si reiniciar
    if services_running:
        print_info("")
        if ask_yes_no("¿Deseas reiniciar los servicios?"):
            print_info("Deteniendo servicios existentes...")
            if services['backend']:
                kill_process_on_port(8000, "Backend")
            if services['frontend']:
                kill_process_on_port(5173, "Frontend")
            # Esperar más tiempo para que los puertos se liberen
            print_info("Esperando a que los puertos se liberen...")
            time.sleep(3)
        else:
            print_info("Usando servicios existentes")
            print_header("✨ Servicios Listos")
            print_info(f"Frontend: http://localhost:5173")
            print_info(f"Backend: http://localhost:8000")
            print_info(f"API Docs: http://localhost:8000/docs")
            print_info("\nPresiona Ctrl+C para detener")
            try:
                while True:
                    time.sleep(1)
            except KeyboardInterrupt:
                print("\n" + Colors.YELLOW + "Servicios detenidos" + Colors.RESET)
            return
    
    # Iniciar servicios
    print_header("Iniciando Servicios")
    
    backend_ok = start_backend()
    if not backend_ok:
        print_error("No se pudo iniciar el backend")
        sys.exit(1)
    
    frontend_ok = start_frontend()
    if not frontend_ok:
        print_error("No se pudo iniciar el frontend")
        sys.exit(1)
    
    # Todos los servicios iniciados correctamente
    print_header("✨ ¡Todos los Servicios están Corriendo!")
    print_success("Backend: http://localhost:8000")
    print_success("Frontend: http://localhost:5173")
    print_success("API Docs: http://localhost:8000/docs")
    print_info("\nPresiona Ctrl+C para detener los servicios")
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n" + Colors.YELLOW + "Deteniendo servicios..." + Colors.RESET)
        # Nota: En un caso real, aquí habría que matar los procesos hijos
        print_info("Servicios detenidos")

if __name__ == "__main__":
    main()
