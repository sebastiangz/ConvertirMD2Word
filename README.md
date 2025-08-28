# ConvertirMD2Word - Conversor Markdown a Word para ELINEA (LMS basado en Moodle)

Programa especializado para convertir archivos Markdown (.md) a documentos Word (.docx) completamente optimizados para el plugin de importacion de libros de Moodle. Cumple al 100% con las especificaciones del plugin para importacion perfecta sin perdida de informacion.

## Optimizado para Plugin de Importacion de Libros de Moodle

### Cumple TODOS los Requerimientos del Plugin:
- **Formato .docx**: Compatible con Moodle 2.7 o superior
- **Estructura de capitulos**: Usa Heading 1 y Heading 2 de Word nativamente
- **Imagenes embebidas**: Formato web-compatible (GIF, PNG, JPEG)
- **Sin macros**: Archivos .docx puros (no .docm)
- **Calidad nativa**: Superior a conversores de Google Docs o LibreOffice

### Mapeo de Estructura:
```
Markdown          ->    Moodle Book
# Titulo          ->    Capitulo
## Subtitulo      ->    Subcapitulo  
### Titulo menor  ->    Texto con formato (no crea capitulos)
```

## Instalacion

### Windows

#### Instalacion Automatica:
1. Descargar todos los archivos del programa
2. Abrir PowerShell o CMD como Administrador
3. Ejecutar el instalador:
```cmd
cd ruta\donde\descargaste\los\archivos
python instalar_windows.bat
```

#### Instalacion Manual:
```cmd
# Instalar Python (si no esta instalado)
# Descargar desde: https://www.python.org/downloads/

# Instalar dependencias basicas
pip install pypandoc python-docx lxml pillow

# Instalar dependencias para correccion automatica de codificacion
pip install chardet ftfy unidecode

# Instalar Pandoc
# Opcion 1: Automatico
python -c "import pypandoc; pypandoc.download_pandoc()"

# Opcion 2: Manual - descargar desde https://pandoc.org/installing.html
```

### Linux (Ubuntu/Debian)

#### Instalacion Automatica:
```bash
# Descargar archivos y ejecutar instalador
cd /ruta/donde/descargaste/los/archivos
chmod +x instalar_linux.sh
./instalar_linux.sh
```

#### Instalacion Manual:
```bash
# Actualizar sistema
sudo apt update

# Instalar Python y pip (si no estan instalados)
sudo apt install python3 python3-pip

# Instalar Pandoc
sudo apt install pandoc

# Instalar dependencias basicas Python
pip3 install pypandoc python-docx lxml pillow

# Instalar dependencias para correccion automatica de codificacion
pip3 install chardet ftfy unidecode
```

### Linux (CentOS/RHEL/Fedora)

```bash
# Para CentOS/RHEL
sudo yum install python3 python3-pip pandoc
# o para versiones mas nuevas:
sudo dnf install python3 python3-pip pandoc

# Para Fedora
sudo dnf install python3 python3-pip pandoc

# Instalar dependencias Python
pip3 install pypandoc python-docx lxml pillow chardet ftfy unidecode
```

### Mac

```bash
# Instalar Homebrew si no esta instalado
/bin/bash -c "$(curl -fsSL https://github.com/sebastiangz/ConvertirMD2Word/HEAD/install.sh)"

# Instalar dependencias
brew install python pandoc

# Instalar dependencias Python
pip3 install pypandoc python-docx lxml pillow chardet ftfy unidecode
```

## Uso del Programa

### Conversion Basica (Optimizada para Moodle):
```bash
# Archivo individual
python ConvertirMD2Word.py mi_leccion.md

# Especificar archivo de salida
python ConvertirMD2Word.py mi_leccion.md mi_leccion_moodle.docx

# Carpeta completa
python ConvertirMD2Word.py -d mis_lecciones/

# Con directorio de salida especifico
python ConvertirMD2Word.py -d mis_lecciones/ salida_word/
```

### Validar Estructura antes de Convertir:
```bash
# Ver si tu Markdown esta bien estructurado para Moodle
python ConvertirMD2Word.py --validar archivo.md

# Validar carpeta completa
python ConvertirMD2Word.py --validar -d carpeta_markdown/
```

### Conversion Sin Optimizaciones (si es necesario):
```bash
python ConvertirMD2Word.py archivo.md --sin-moodle
```

### Modo Verbose (detallado):
```bash
python ConvertirMD2Word.py archivo.md -v
```

## Estructura Markdown Recomendada para Moodle

### Ejemplo Perfecto para Moodle:
```markdown
# Capitulo 1: Introduccion
Este sera un capitulo principal en Moodle Book.

## 1.1 Conceptos Basicos  
Este sera un subcapitulo del Capitulo 1.

### Definiciones Importantes
Esto aparece como texto formateado (no crea nuevo capitulo).

- Lista de elementos
- Elemento 2

## 1.2 Objetivos del Capitulo
Otro subcapitulo del Capitulo 1.

# Capitulo 2: Desarrollo
Segundo capitulo principal.

## 2.1 Metodologia
Subcapitulo del Capitulo 2.

![Imagen importante](diagrama.png)
```

### Resultado en Moodle Book:
```
Capitulo 1: Introduccion
  1.1 Conceptos Basicos
  1.2 Objetivos del Capitulo
Capitulo 2: Desarrollo  
  2.1 Metodologia
```

## Caracteristicas Tecnicas para Moodle

### Estilos de Word Generados:
- **Heading 1**: Estilos nativos de Word para capitulos principales
- **Heading 2**: Estilos nativos de Word para subcapitulos
- **Normal**: Texto del cuerpo optimizado para legibilidad
- **Plantillas**: Generacion automatica compatible con Moodle

### Manejo Inteligente de Codificacion:
- **Deteccion automatica**: Utiliza `chardet` para detectar la codificacion del archivo
- **Reparacion automatica**: Usa `ftfy` (fix text for you) para corregir texto mal codificado
- **Conversion de caracteres**: Emplea `unidecode` para caracteres especiales problematicos
- **Fallback robusto**: Sistema de respaldo con correcciones basicas si los modulos no estan disponibles
- **Sin diccionarios estaticos**: Correccion dinamica adaptada a cada archivo
- **Embebido automatico**: Las imagenes se incluyen en el .docx
- **Formatos soportados**: PNG, JPEG, GIF (compatibles con web)
- **Rutas relativas**: Procesamiento automatico desde el Markdown
- **Optimizacion**: Tamaño y calidad balanceados para Moodle

### Elementos Soportados:
- **Texto**: Negrita, cursiva, subrayado
- **Listas**: Numeradas y con viñetas (multiples niveles)
- **Enlaces**: URLs y referencias internas
- **Codigo**: Bloques con sintaxis resaltada
- **Tablas**: Formato GitHub Flavored Markdown
- **Citas**: Bloques de citas con formato

## Proceso Completo: Markdown -> Moodle

### Paso 1: Preparar tu Markdown
```bash
# Validar estructura
python ConvertirMD2Word.py --validar mi_curso.md
```

### Paso 2: Convertir a Word
```bash
# Conversion optimizada
python ConvertirMD2Word.py mi_curso.md
```

### Paso 3: Importar en Moodle
1. **Moodle**: Ve a tu curso -> **Actividades** -> **Libro**
2. **Crear**: Nuevo libro o editar existente
3. **Importar**: Configuracion -> **Importar capitulos**
4. **Subir**: Tu archivo .docx generado
5. **Listo**: Capitulos y subcapitulos creados automaticamente

## Ventajas vs Otras Soluciones

| Caracteristica | ConvertirMD2Word | Pandoc Directo | CloudConvert |
|----------------|------------------|----------------|--------------|
| **Optimizado para Moodle** | SI | NO | NO |
| **Estilos Heading correctos** | SI | Parcial | NO |
| **Correccion automatica de codificacion** | SI | NO | NO |
| **Sin marcas de agua** | SI | SI | NO |
| **Imagenes embebidas** | SI | Parcial | Parcial |
| **Validacion previa** | SI | NO | NO |
| **Conversion en lote** | SI | NO | NO |
| **Post-procesamiento** | SI | NO | NO |
| **Funciona offline** | SI | SI | NO |

## Ejemplos de Casos de Uso

### Profesor con Multiples Lecciones:
```bash
# Estructura de carpeta:
mis_lecciones/
  ├── 01_introduccion.md
  ├── 02_fundamentos.md  
  ├── 03_practica.md
  └── imagenes/
      ├── diagrama1.png
      └── foto1.jpg

# Convertir todo:
python ConvertirMD2Word.py -d mis_lecciones/

# Resultado:
mis_lecciones/
  ├── 01_introduccion.docx  <- Listo para Moodle
  ├── 02_fundamentos.docx   <- Listo para Moodle
  ├── 03_practica.docx      <- Listo para Moodle
```

### Curso Completo en un Archivo:
```bash
# archivo: curso_completo.md con estructura:
# Capitulo 1: Introduccion
## 1.1 Objetivos  
## 1.2 Metodologia
# Capitulo 2: Teoria
## 2.1 Conceptos
## 2.2 Ejemplos
# Capitulo 3: Practica

python ConvertirMD2Word.py curso_completo.md

# Resultado en Moodle: 3 capitulos, 5 subcapitulos
```

### Validacion antes de Conversion:
```bash
python ConvertirMD2Word.py --validar mi_documento.md

# Salida tipica:
# Analisis del Markdown:
#    Titulos H1 validos (capitulos): 3
#    Titulos H2 (subcapitulos): 7
#    Titulos H3+: 12
#    Imagenes referenciadas: 5
```

## Solucion de Problemas Especificos de Moodle

### "No se crean capitulos en Moodle"
**Causa**: No hay titulos `# Nivel 1` en tu Markdown  
**Solucion**: 
```bash
# Verificar estructura:
python ConvertirMD2Word.py --validar archivo.md

# Si sale "Titulos H1 validos (capitulos): 0"
# Cambiar ## por # en tus titulos principales
```

### "Las imagenes no aparecen"
**Causa**: Imagenes en formato no compatible o rutas incorrectas  
**Solucion**:
- Usar PNG, JPEG, o GIF
- Rutas relativas al archivo .md
- Verificar que las imagenes existen

### "Formato extraño al importar"
**Causa**: Conversion sin optimizaciones de Moodle  
**Solucion**:
```bash
# Asegurar que NO uses --sin-moodle
python ConvertirMD2Word.py archivo.md  # <- Optimizacion activada por defecto
```

### "Caracteres corruptos o emojis extraños"
**Causa**: Problemas de codificacion en el archivo original  
**Solucion**: El programa usa correccion automatica avanzada:
```bash
# El programa detecta automaticamente:
# - Codificacion del archivo (chardet)
# - Repara texto corrupto (ftfy) 
# - Convierte caracteres especiales (unidecode)
# - Aplica correcciones basicas como respaldo

python ConvertirMD2Word.py --validar archivo.md -v
# Reporta: "Codificacion detectada: utf-8 (confianza: 0.95)"
# Reporta: "ftfy aplico correcciones de codificacion"
```

## Correccion Automatica de Codificacion

### Modulos Utilizados:
- **chardet**: Deteccion automatica de codificacion de archivos
- **ftfy**: Reparacion inteligente de texto mal codificado
- **unidecode**: Conversion de caracteres Unicode problematicos

### Problemas que Resuelve Automaticamente:
- Caracteres latinos mal codificados (`Ã¡` → `á`, `Ã±` → `ñ`)
- Comillas y guiones corruptos (`â€œ` → `"`, `â€"` → `–`)
- Emojis y simbolos mal interpretados
- Caracteres de control invisibles
- Problemas de BOM (Byte Order Mark)
- Codificaciones mixtas en el mismo archivo

### Ejemplo de Correccion:
```
Antes:  "ProgramaciÃ³n funcional con âœ… y ðŸ'¡"
Despues: "Programacion funcional con  y "
```

## Checklist Pre-Importacion

### Antes de Convertir:
- [ ] Archivo .md con titulos `#` para capitulos principales
- [ ] Subtitulos `##` para subcapitulos  
- [ ] Imagenes en formato PNG/JPEG/GIF
- [ ] Rutas de imagenes correctas y relativas

### Despues de Convertir:
- [ ] Archivo .docx generado exitosamente
- [ ] Tamaño razonable (no excesivamente grande)
- [ ] Validacion muestra estructura esperada

### Al Importar en Moodle:
- [ ] Moodle 2.7 o superior
- [ ] PHP XSL extension habilitada en el servidor
- [ ] Plugin Book Import instalado y activo

## Requerimientos del Plugin Moodle (Cumplidos)

> **Plugin oficial**: *This plugin supports importing a Microsoft Word docx-formatted file as chapters to a book. The file is split into chapters and subchapters based on the built-in heading styles "Heading 1" and "Heading 2" in Word.*

### Cumplimiento Total:
- Formato .docx: Compatible Microsoft Word nativo
- Heading 1 y Heading 2: Estilos built-in de Word correctos
- Division automatica: Capitulos y subcapitulos segun estructura
- Imagenes embebidas: GIF, PNG, JPEG en formato web-compatible
- Sin .doc/.docm: Solo .docx puro sin macros
- Calidad nativa: Superior a Google Docs/LibreOffice

## Soporte y Debugging

### Informacion de Debug:
```bash
# Verbose mode para ver detalles completos
python ConvertirMD2Word.py archivo.md -v

# Ver solo validacion sin convertir
python ConvertirMD2Word.py --validar archivo.md -v
```

### Problemas Comunes:

#### Windows:
1. **Error "python no reconocido"**: Instalar Python desde python.org
2. **Error de Pandoc**: Descargar manualmente desde pandoc.org
3. **Permisos**: Ejecutar CMD como Administrador

#### Linux:
1. **Error "pip3 no encontrado"**: `sudo apt install python3-pip`
2. **Error de permisos**: Usar `sudo` para instalar paquetes del sistema
3. **Pandoc version antigua**: Actualizar con `sudo apt update && sudo apt upgrade pandoc`

### Comandos de Diagnostico:

#### Windows:
```cmd
# Verificar Python
python --version

# Verificar dependencias
pip list | findstr pypandoc
pip list | findstr docx

# Verificar Pandoc
python -c "import pypandoc; print(pypandoc.get_pandoc_version())"
```

#### Linux:
```bash
# Verificar Python
python3 --version

# Verificar dependencias
pip3 list | grep pypandoc
pip3 list | grep docx

# Verificar Pandoc
python3 -c "import pypandoc; print(pypandoc.get_pandoc_version())"
```

### Estadisticas de Conversion:
El programa reporta automaticamente:
- Numero de capitulos detectados (Heading 1)
- Numero de subcapitulos detectados (Heading 2)  
- Imagenes embebidas exitosamente
- Advertencias sobre estructura

## Archivos del Programa

### Archivos Principales:
- `ConvertirMD2Word.py` - Programa principal con correccion automatica de codificacion
- `README.md` - Este archivo de instrucciones
- `instalar_windows.bat` - Instalador automatico para Windows
- `instalar_linux.sh` - Instalador automatico para Linux  
- `requirements.txt` - Lista completa de dependencias (incluye chardet, ftfy, unidecode)

### Archivos de Prueba (incluidos):
- `README.md` - Archivo Markdown de ejemplo
- `README.docx` - Resultado esperado

## Listo para ELinea

Con este programa, tus archivos Markdown se convierten a documentos Word **perfectamente optimizados** para el plugin de importacion de libros de ELinea (LMS basado en Moodle). No mas problemas de formato, estructura incorrecta o archivos incompatibles.

**Convierte una vez e importa **