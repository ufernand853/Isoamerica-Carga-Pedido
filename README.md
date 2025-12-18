# Isoamerica Carga Pedido

Script para completar la planilla de pedido a partir de un listado general de productos usando pandas y openpyxl. La interfaz para seleccionar archivos usa Tkinter.

## Requisitos
- Python 3.10 o superior.
- Dependencias Python: ver `requirements.txt`.
- **Tkinter**: es parte estándar de Python pero en algunas instalaciones no viene incluido por defecto. Si ves un error `ModuleNotFoundError: No module named 'tkinter'`, instala el paquete del sistema:
  - **Windows**: usa el instalador oficial de Python desde [python.org](https://www.python.org/downloads/) asegurándote de incluir la opción "tcl/tk and IDLE". Si ya tienes Python, repara la instalación desde el instalador y marca esa casilla.
  - **Ubuntu/Debian**: `sudo apt-get update && sudo apt-get install -y python3-tk`
  - **Fedora**: `sudo dnf install python3-tkinter`
  - **macOS**: instala Python de python.org o con `brew install python-tk@3` según tu gestor.

## Instalación
```bash
python -m venv .venv
source .venv/bin/activate  # En Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Uso
```bash
python completar_planilla.py
```
El programa abrirá dos diálogos para seleccionar primero la planilla general y luego la planilla de pedido. Se generará un archivo nuevo junto a la planilla de pedido con sufijo `procesada_YYYY-MM-DD`.

## Generar ejecutable con PyInstaller
1. Asegúrate de tener Tkinter instalado (ver arriba) y las dependencias de `requirements.txt`.
2. Instala PyInstaller:
   ```bash
   pip install pyinstaller
   ```
3. Desde el directorio del proyecto ejecuta:
   ```bash
   pyinstaller --onefile --noconsole --name "cargar_planilla" completar_planilla.py
   ```
4. El ejecutable quedará en `dist/`. Si necesitas incluir archivos adicionales, agrégalos con `--add-data` siguiendo las reglas de PyInstaller para tu sistema operativo.
