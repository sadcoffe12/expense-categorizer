#!/usr/bin/env python3
"""
Script para ver los logs de la aplicación
"""

import os
import sys
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
    print(f"\n{Colors.BOLD}{Colors.BLUE}{'='*60}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.BLUE}{message:^60}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.BLUE}{'='*60}{Colors.RESET}\n")

def print_info(message):
    print(f"{Colors.BLUE}ℹ️  {message}{Colors.RESET}")

def main():
    log_dir = Path(__file__).parent / "logs"
    log_file = log_dir / "app.log"
    
    print_header("📋 Expense Categorizer - Ver Logs")
    
    if not log_dir.exists():
        print(f"{Colors.RED}❌ Directorio de logs no encontrado: {log_dir}{Colors.RESET}")
        print(f"{Colors.YELLOW}Ejecuta la aplicación primero para generar logs.{Colors.RESET}")
        return
    
    if not log_file.exists():
        print(f"{Colors.RED}❌ Archivo de logs no encontrado: {log_file}{Colors.RESET}")
        print(f"{Colors.YELLOW}Ejecuta la aplicación primero para generar logs.{Colors.RESET}")
        return
    
    print_info(f"Ubicación del archivo de logs:")
    print(f"{Colors.GREEN}{log_file}{Colors.RESET}\n")
    
    # Mostrar últimas líneas del log
    print_header("📄 Últimas 50 líneas del log")
    
    with open(log_file, 'r') as f:
        lines = f.readlines()
        # Mostrar las últimas 50 líneas
        start = max(0, len(lines) - 50)
        for line in lines[start:]:
            # Colorear según el nivel
            if "ERROR" in line:
                print(f"{Colors.RED}{line.rstrip()}{Colors.RESET}")
            elif "WARNING" in line:
                print(f"{Colors.YELLOW}{line.rstrip()}{Colors.RESET}")
            elif "INFO" in line:
                print(f"{Colors.GREEN}{line.rstrip()}{Colors.RESET}")
            else:
                print(line.rstrip())
    
    print(f"\n{Colors.BLUE}💡 Tip: Puedes abrir el archivo de logs en tu editor:{Colors.RESET}")
    print(f"   code {log_file}")
    print(f"{Colors.BLUE}O verlo en tiempo real con:{Colors.RESET}")
    print(f"   tail -f {log_file}")

if __name__ == "__main__":
    main()
