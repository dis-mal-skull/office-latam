#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import subprocess
import urllib.request
import ctypes
import time
import winreg
from pathlib import Path

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN CORREGIDA
# ══════════════════════════════════════════════════════════════════════════════

VERSIONS = {
    # IMPORTANTE: Para activación KMS se requieren versiones VOLUME
    "2019": {
        "name": "Office 2019 LTSC",
        "product_id": "ProPlus2019Volume",  # Cambiado de Retail a Volume para compatibilidad KMS
        "channel": "PerpetualVL2019",
        "gvlk": "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP",
    },
    "2021": {
        "name": "Office 2021 LTSC",
        "product_id": "ProPlus2021Volume",
        "channel": "PerpetualVL2021",
        "gvlk": "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH",
    }
}

# Lista de programas disponibles para instalación
AVAILABLE_APPS = {
    "Excel": "Excel",
    "Word": "Word", 
    "PowerPoint": "PowerPoint",
    "Outlook": "Outlook",
    "Access": "Access",
    "Publisher": "Publisher",
    "OneNote": "OneNote",
    "Groove": "Groove",
    "Lync": "Lync",
    "OneDrive": "OneDrive",
    "Teams": "Teams"
}

# Programas seleccionados por el usuario
SELECTED_APPS = {}

LANGUAGE = "es-es"
KMS_SERVER = "kms.digiboy.ir"
KMS_PORT = "1688"
WORK_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
ODT_URL = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
SELECTED_VERSION = None

# Colores y Utilidades de Consola
class Colors:
    HEADER = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    WHITE = '\033[97m'

def enable_ansi_colors():
    try:
        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
    except:
        pass

def print_header():
    os.system('cls' if os.name == 'nt' else 'clear')
    print(f"""{Colors.CYAN}
╔══════════════════════════════════════════════════════════════════════════════╗
║                    INSTALADOR AUTOMÁTICO DE OFFICE                           ║
║                         Español - Selección Personalizada                     ║
╚══════════════════════════════════════════════════════════════════════════════╝
{Colors.ENDC}""")

def print_step(step_num, total, message):
    print(f"\n{Colors.BOLD}{Colors.BLUE}[{step_num}/{total}]{Colors.ENDC} {Colors.CYAN}{message}{Colors.ENDC}")
    print("─" * 70)

def print_success(message):
    print(f"  {Colors.GREEN}✓ {message}{Colors.ENDC}")

def print_error(message):
    print(f"  {Colors.RED}✗ {message}{Colors.ENDC}")

def print_warning(message):
    print(f"  {Colors.YELLOW}⚠ {message}{Colors.ENDC}")

def print_info(message):
    print(f"  {Colors.CYAN}→ {message}{Colors.ENDC}")

# Barra de progreso mejorada
def print_progress_bar(current, total, prefix='', suffix='', length=40):
    if total <= 0:
        percent = 0
    else:
        percent = min(100, (current / total) * 100)
    
    filled = int(length * percent / 100)
    bar = '█' * filled + '░' * (length - filled)
    
    current_mb = current / (1024 * 1024)
    total_mb = total / (1024 * 1024)
    
    print(f'\r  {Colors.CYAN}{prefix} │{Colors.GREEN}{bar}{Colors.CYAN}│ {percent:5.1f}% {Colors.WHITE}({current_mb:.1f}/{total_mb:.1f} MB){Colors.ENDC} {suffix}', end='', flush=True)

# Lógica Principal
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if sys.platform == 'win32':
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit(0)

def select_version():
    global SELECTED_VERSION
    print(f"\n  {Colors.GREEN}[1]{Colors.WHITE} Office 2019 (LTSC Volumen)")
    print(f"  {Colors.GREEN}[2]{Colors.WHITE} Office 2021 (LTSC Volumen)")
    print()
    
    while True:
        try:
            choice = input(f"  {Colors.CYAN}Seleccione versión: {Colors.ENDC}").strip()
            if choice == "1":
                SELECTED_VERSION = VERSIONS["2019"]
                return True
            elif choice == "2":
                SELECTED_VERSION = VERSIONS["2021"]
                return True
        except KeyboardInterrupt:
            return False

def select_apps():
    global SELECTED_APPS
    print(f"\n{Colors.BOLD}{Colors.CYAN}SELECCIÓN DE PROGRAMAS{Colors.ENDC}")
    print("─" * 50)
    print(f"{Colors.WHITE}Marque los programas que desea instalar:{Colors.ENDC}\n")
    
    # Mostrar lista de programas con números
    app_list = list(AVAILABLE_APPS.keys())
    for i, app in enumerate(app_list, 1):
        print(f"  {Colors.GREEN}[{i}]{Colors.WHITE} {app}")
    
    print(f"\n  {Colors.YELLOW}[T]{Colors.WHITE} Todos los programas")
    print(f"  {Colors.YELLOW}[S]{Colors.WHITE} Solo Excel (configuración original)")
    print(f"  {Colors.YELLOW}[N]{Colors.WHITE} Ninguno (cancelar)")
    print()
    
    while True:
        try:
            choice = input(f"  {Colors.CYAN}Seleccione programas (ej: 1,3,5 o T o S): {Colors.ENDC}").strip().upper()
            
            if choice == "T":
                # Seleccionar todos
                SELECTED_APPS = {app: True for app in AVAILABLE_APPS.keys()}
                return True
            elif choice == "S":
                # Solo Excel (configuración original)
                SELECTED_APPS = {"Excel": True}
                return True
            elif choice == "N":
                return False
            else:
                # Procesar selección personalizada
                selected_numbers = [num.strip() for num in choice.split(',')]
                valid_selection = True
                SELECTED_APPS = {}
                
                for num_str in selected_numbers:
                    try:
                        num = int(num_str)
                        if 1 <= num <= len(app_list):
                            app_name = app_list[num - 1]
                            SELECTED_APPS[app_name] = True
                        else:
                            valid_selection = False
                            break
                    except ValueError:
                        valid_selection = False
                        break
                
                if valid_selection and SELECTED_APPS:
                    return True
                else:
                    print_error("Selección no válida. Intente de nuevo.")
                    
        except KeyboardInterrupt:
            return False

def uninstall_office():
    print_info("Iniciando desinstalación rápida...")
    
    setup_path = WORK_DIR / "setup.exe"
    config_path = WORK_DIR / "remove_config.xml"
    
    remove_config = """<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>"""
    
    with open(config_path, 'w', encoding='utf-8') as f:
        f.write(remove_config)
        
    if setup_path.exists():
        subprocess.run(f'"{setup_path}" /configure "{config_path}"', shell=True)
        print_success("Comandos de limpieza ejecutados")

    if config_path.exists(): config_path.unlink()
    return True

def download_with_progress(url, destination, description="archivo"):
    print_info(f"Descargando {description}...")
    try:
        # Intentar borrar si existe y está corrupto
        if os.path.exists(destination):
            try: os.remove(destination)
            except: pass

        req = urllib.request.Request(url, method='HEAD')
        with urllib.request.urlopen(req) as response:
            total_size = int(response.headers.get('content-length', 0))
        
        downloaded = 0
        block_size = 8192
        with urllib.request.urlopen(url) as response:
            with open(destination, 'wb') as out_file:
                while True:
                    buffer = response.read(block_size)
                    if not buffer: break
                    downloaded += len(buffer)
                    out_file.write(buffer)
                    print_progress_bar(downloaded, total_size, prefix='⬇ ')
        print()
        return True
    except Exception as e:
        print_error(f"Error: {e}")
        return False

def get_folder_size(path):
    total = 0
    if path.exists():
        for entry in path.rglob('*'):
            if entry.is_file():
                try: total += entry.stat().st_size
                except: pass
    return total

def download_office_with_progress():
    print_info(f"Descargando archivos para {SELECTED_VERSION['name']}...")
    print_warning("Limpiando descargas anteriores conflictivas...")
    
    import shutil
    office_folder = WORK_DIR / "Office"
    if office_folder.exists():
        try: shutil.rmtree(office_folder)
        except: pass

    setup_path = WORK_DIR / "setup.exe"
    config_path = WORK_DIR / "Config.xml"
    
    # Iniciar descarga
    process = subprocess.Popen(
        f'"{setup_path}" /download "{config_path}"',
        shell=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    
    # Esperar conexión inicial
    timeout = 0
    while not office_folder.exists() and timeout < 30:
        time.sleep(1)
        timeout += 1
        print(f"    Iniciando conexión ({timeout}s)...", end='\r')
    
    estimated_size = 350 * 1024 * 1024 
    last_size = 0
    stagnant_count = 0
    
    while process.poll() is None:
        current_size = get_folder_size(office_folder)
        
        if current_size == last_size and current_size > 0:
            stagnant_count += 1
        else:
            stagnant_count = 0 
            
        status = ""
        if stagnant_count > 10: status = "(Procesando...)"
        
        print_progress_bar(current_size, estimated_size, prefix='⬇ ', suffix=status)
        last_size = current_size
        time.sleep(1)
        
    print()
    if process.returncode == 0:
        print_success("Archivos descargados correctamente")
        return True
    else:
        print_error("Error durante la descarga. Revise su conexión.")
        return False

def install_office():
    print_info(f"Instalando {SELECTED_VERSION['name']}...")
    print_warning("Por favor espere y NO cierre la ventana.")
    
    setup_path = WORK_DIR / "setup.exe"
    config_path = WORK_DIR / "Config.xml"
    
    result = subprocess.run(f'"{setup_path}" /configure "{config_path}"', shell=True)
    if result.returncode == 0:
        print_success("Instalación completada")
        return True
    else:
        print_error("La instalación falló")
        return False

def activate_office():
    print_info("Activando licencia KMS...")
    
    ospp_path = None
    possible_paths = [
        Path(os.environ.get('ProgramFiles', 'C:\\Program Files')) / "Microsoft Office" / "Office16" / "ospp.vbs",
        Path(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)')) / "Microsoft Office" / "Office16" / "ospp.vbs",
    ]
    
    for path in possible_paths:
        if path.exists(): ospp_path = path; break
        
    if not ospp_path:
        print_warning("No se encontró ospp.vbs (¿Instalación fallida?)")
        return False

    commands = [
        f'/inpkey:{SELECTED_VERSION["gvlk"]}',
        f'/sethst:{KMS_SERVER}',
        f'/setprt:{KMS_PORT}',
        '/act'
    ]
    
    for cmd_arg in commands:
        subprocess.run(f'cscript //nologo "{ospp_path}" {cmd_arg}', shell=True, stdout=subprocess.DEVNULL)
    
    result = subprocess.run(f'cscript //nologo "{ospp_path}" /dstatus', shell=True, capture_output=True, text=True)
    if "LICENSED" in result.stdout.upper():
        print_success("¡Activación Exitosa!")
        return True
    return False

def create_config_xml():
    # Generar configuración basada en programas seleccionados
    selected_apps_list = list(SELECTED_APPS.keys())
    
    # Crear lista de exclusiones (programas NO seleccionados)
    exclude_apps = []
    for app in AVAILABLE_APPS.keys():
        if app not in SELECTED_APPS:
            exclude_apps.append(f'      <ExcludeApp ID="{app}" />')
    
    exclude_section = "\n".join(exclude_apps) if exclude_apps else ""
    
    config = f"""<Configuration>
  <Add OfficeClientEdition="64" Channel="{SELECTED_VERSION['channel']}">
    <Product ID="{SELECTED_VERSION['product_id']}">
      <Language ID="{LANGUAGE}" />
{exclude_section}
    </Product>
  </Add>
  <RemoveMSI />
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>"""
    
    with open(WORK_DIR / "Config.xml", 'w') as f: 
        f.write(config)
    
    # Mostrar resumen de programas seleccionados
    print_info(f"Programas a instalar: {', '.join(selected_apps_list)}")
    return True

def main():
    enable_ansi_colors()
    print_header()
    
    if not is_admin():
        print_warning("Solicitando permisos de administrador...")
        run_as_admin()
        return

    if not select_version(): return
    if not select_apps(): return

    # Limpia descargas anteriores para evitar conflictos de canal
    try:
        import shutil
        shutil.rmtree(WORK_DIR / "Office", ignore_errors=True)
        shutil.rmtree(WORK_DIR / "office", ignore_errors=True)
    except: pass

    # Detectar y limpiar versiones previas
    print_step(1, 6, "LIMPIEZA INICIAL")
    uninstall_office()

    # Descargar ODT
    print_step(2, 6, "HERRAMIENTAS")
    setup_path = WORK_DIR / "setup.exe"
    if not setup_path.exists():
        if not download_with_progress(ODT_URL, str(setup_path), "Office Deployment Tool"): return

    # Descargar Excel
    print_step(3, 6, "DESCARGA DE ARCHIVOS")
    create_config_xml()
    
    print_info(f"Canal seleccionado: {SELECTED_VERSION['channel']}")
    print_info(f"Producto ID: {SELECTED_VERSION['product_id']}")
    
    if not download_office_with_progress():
        print_error("Fallo en descarga")
        # No continuamos si la descarga falla gravemente
        input("Presione ENTER para salir...")
        return

    # Instalar
    print_step(4, 6, "INSTALACIÓN")
    install_office()

    # Activar
    print_step(5, 6, "ACTIVACIÓN")
    activate_office()
    
    # Limpieza final
    print_step(6, 6, "FINALIZANDO")
    try:
        (WORK_DIR / "Config.xml").unlink(missing_ok=True)
        (WORK_DIR / "remove_config.xml").unlink(missing_ok=True)
        import shutil
        shutil.rmtree(WORK_DIR / "Office", ignore_errors=True)
    except: pass
    
    print_success("Proceso Terminado")
    input("\nPresione ENTER para salir...")

if __name__ == "__main__":
    main()
