# Catálogo fotográfico de pozos

Aplicación de escritorio para catalogar fotografías arqueológicas por sitio,
pozo y nivel, preparar copias con nombres normalizados y generar el Excel de
entrega al Museo Nacional.

<p align="center">
  <img src="assets/cucharilla.png" alt="Cucharilla arqueológica" width="220">
</p>

## Características

- Conserva intactos los archivos originales.
- Copia cada fotografía sin cambiar su formato ni sus metadatos.
- Tolera caracteres de control defectuosos en metadatos EXIF al escribir Excel,
  sin modificar esos metadatos dentro de la fotografía original o su copia.
- Genera consecutivos por pozo y nivel: `F001`, `F002`, `F003`…
- Crea `Entrega_Museo`, con el Excel y la subcarpeta `Fotos`.
- Visor con zoom hasta 800 %, desplazamiento, ajuste, vista al 100 % y rotación.
- Respeta la orientación EXIF.
- Permite omitir fotografías, recuperarlas y reanudar el trabajo posteriormente.
- Permite corregir un registro ya guardado desde su miniatura. Actualiza de forma
  coordinada la fila de Excel y, si cambia pozo o nivel, el nombre de la copia.
- Conserva en el archivo de estado un historial de las correcciones realizadas.
- Lee JPG/JPEG, PNG, TIFF y WEBP. HEIC/HEIF se habilita con `pillow-heif`.
- Interfaz en español para Windows, Linux y macOS.

## Windows: instalación sencilla

1. Abre la sección **Releases** de este repositorio.
2. Descarga `Catalogador_Pozos_Setup_2.1.1.exe`.
3. Haz doble clic y sigue el asistente: **Siguiente → Instalar → Finalizar**.
4. Abre **Catálogo fotográfico de pozos** desde el menú Inicio o el acceso
   directo del escritorio.

El instalador incluye Python, Tkinter, Pillow, openpyxl y soporte HEIC. Los
usuarios finales no necesitan ejecutar `INSTALAR.bat` ni instalar Python.

> El instalador no está firmado digitalmente. Windows puede mostrar “Editor
> desconocido” o una advertencia de SmartScreen. Para una distribución pública
> sin esa advertencia se necesita firmar el instalador con un certificado de
> firma de código.

### Versión portátil para Windows

Quien prefiera ejecutar el código fuente puede instalar Python 3, ejecutar
`INSTALAR.bat` una vez y después abrir `ABRIR.bat`.

## Uso básico

1. Elige la carpeta con las fotografías descargadas.
2. Elige otra carpeta para la salida.
3. Escribe o confirma el código del sitio y pulsa **Cargar / reanudar**.
4. Revisa la fotografía, escribe manualmente Pozo y Nivel y pulsa
   **Guardar y siguiente**.
5. Usa **Omitir por ahora** si todavía no puedes identificar una foto. Su
   miniatura permite recuperarla después.
6. Si detectas un error, selecciona la miniatura registrada, pulsa **Corregir
   datos**, modifica los campos y guarda la corrección. Nunca edites el Excel y
   el nombre de la copia por separado.
7. Entrega la carpeta `Entrega_Museo` completa. Conserva
   `estado_catalogador.json` fuera de la entrega para poder reanudar.

El manual detallado, incluidos todos los controles de zoom, está en
[`LEEME.txt`](LEEME.txt).

## Linux

La aplicación es compatible a nivel de código. No se distribuye todavía como
paquete `.deb`, `.rpm` o AppImage.

### Ubuntu, Debian y derivados

```bash
sudo apt update
sudo apt install python3 python3-venv python3-tk
git clone URL_DE_ESTE_REPOSITORIO
cd catalogador_pozos
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python catalogador_pozos.py
```

En Fedora, el paquete de Tkinter suele llamarse `python3-tkinter`:

```bash
sudo dnf install python3 python3-tkinter
```

Si `pillow-heif` no se puede instalar en una distribución concreta, instala
solo `Pillow` y `openpyxl`. El programa seguirá funcionando, pero sin HEIC/HEIF.

## macOS

La forma más sencilla es instalar Python 3 desde
[python.org](https://www.python.org/downloads/macos/), cuya distribución incluye
Tkinter, y ejecutar:

```bash
git clone URL_DE_ESTE_REPOSITORIO
cd catalogador_pozos
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python catalogador_pozos.py
```

En macOS también están disponibles los atajos `Command +`, `Command -`,
`Command 0` y `Command 1`. Actualmente no se entrega una aplicación `.app`
firmada o notarizada.

## Compatibilidad

| Sistema | Estado | Distribución disponible |
|---|---|---|
| Windows 10/11 x64 | Probado | Instalador `.exe` y código fuente |
| Linux con Python 3.10+ y Tk | Compatible por código; pendiente de prueba física | Código fuente |
| macOS con Python 3.10+ y Tk | Compatible por código; pendiente de prueba física | Código fuente |

## Desarrollo

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# Linux/macOS: source .venv/bin/activate
python -m pip install -r requirements.txt
python -m unittest -v pruebas_catalogador.py
python catalogador_pozos.py
```

Las pruebas comprueban consecutivos, integridad de originales y copias, Excel,
omisión/recuperación, correcciones y su historial, compatibilidad del estado,
bloqueo real del Excel en Windows, recuperación de transacciones y HEIC cuando
está disponible.

## Construir el instalador de Windows

Requisitos para quien compila, no para el usuario final:

- Python 3.10 o posterior.
- [Inno Setup 6](https://jrsoftware.org/isinfo.php).

Ejecuta `CONSTRUIR_INSTALADOR.bat`. El script instala las dependencias de
construcción, crea la aplicación autónoma con PyInstaller y compila el asistente
con Inno Setup. El resultado queda en `instalador_generado`.

También se incluye el flujo `.github/workflows/compilar-windows.yml`. Al crear
una etiqueta como `v2.1.1`, GitHub Actions compila el instalador, lo guarda como
artefacto y lo adjunta a la Release correspondiente.

## Publicación en GitHub

1. Sustituye `URL_DE_ESTE_REPOSITORIO` en este README por la URL real.
2. Crea el repositorio y sube estos archivos.
3. Elige y añade una licencia antes de hacerlo público. No se incluye una por
   defecto porque esa decisión corresponde a la persona propietaria del código.
4. Crea una etiqueta `v2.1.1` para generar la primera Release automáticamente.

## Privacidad y datos

La aplicación trabaja localmente. No sube fotografías, catálogos ni datos a
Internet. El único acceso de red se utiliza al instalar dependencias desde PyPI
cuando se ejecuta desde el código fuente.
