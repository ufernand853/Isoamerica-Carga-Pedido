# Isoamerica Carga Pedido

Script para completar la planilla de pedido a partir de un listado general de productos usando pandas y openpyxl. Por defecto utiliza Tkinter para seleccionar los archivos con cuadros de diálogo, pero también puedes forzar el flujo por consola con `--cli` si lo prefieres.

## Requisitos
- Python 3.10 o superior.
- Dependencias Python: ver `requirements.txt`.
- Tkinter instalado en tu intérprete de Python (viene incluido en la mayoría de instalaciones). Si te falta:
  - **Windows**: usa el instalador oficial de python.org y marca la casilla **tcl/tk and IDLE**. Si ya tienes Python, vuelve a ejecutarlo y elige **Modify → tcl/tk and IDLE**.
  - **Debian/Ubuntu**: `sudo apt-get update && sudo apt-get install -y python3-tk`.
  - **Fedora**: `sudo dnf install python3-tkinter`.
  - **macOS (Homebrew)**: `brew install python-tk`.
  - Comprueba que quedó instalado ejecutando `python - <<'PY'
import tkinter
print('Tkinter OK', tkinter.TkVersion)
PY`.
  - Si ves `ModuleNotFoundError: No module named 'tkinter'`, instala Tkinter como arriba o lanza el script con `--cli` para usar el flujo por consola.

## Instalación
```bash
python -m venv .venv
source .venv/bin/activate  # En Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Uso
```bash
python completar_planilla.py [--pedido RUTA_PEDIDO] [--listado RUTA_LISTADO] [--output RUTA_SALIDA] [--cli]
```

- Si no proporcionas rutas, el script abrirá diálogos de selección usando Tkinter. Con `--cli`, solicitará las rutas por consola.
- Si defines `salida`/`output` en `config_archivos.txt`, ese valor se usa como archivo de salida por defecto. Si no, se genera junto a la planilla de pedido con sufijo `procesada_YYYY-MM-DD`, salvo que indiques `--output`.
- Puedes definir los nombres o rutas por defecto en `config_archivos.txt` usando formato `clave=valor` (claves: `pedido`, `listado`, `salida`/`output`). Si falta el archivo o alguna clave, se usan los valores por defecto incluidos.
  
Ejemplo especificando rutas:
```bash
python completar_planilla.py --pedido "/ruta/Planilla pedido.xlsx" --listado "/ruta/Listado general.xlsx"
```

## Generar ejecutable con PyInstaller
1. Instala las dependencias de `requirements.txt` en tu entorno y asegúrate de que Tkinter funcione (ejecuta `python - <<'PY'
import tkinter
print(tkinter.TkVersion)
PY`).
2. Instala PyInstaller:
   ```bash
   pip install pyinstaller
   ```
3. Desde el directorio del proyecto ejecuta (el `--hidden-import=tkinter` ayuda a detectar Tkinter en algunos entornos):
   ```bash
   pyinstaller --onefile --noconsole --hidden-import=tkinter --name "cargar_planilla" completar_planilla.py
   ```
4. El ejecutable quedará en `dist/`. Si necesitas incluir archivos adicionales, agrégalos con `--add-data` siguiendo las reglas de PyInstaller para tu sistema operativo. Cuando ejecutes el `.exe`, podrás usar los cuadros de diálogo de Tkinter o forzar el modo consola añadiendo `--cli`.
