@echo off
chcp 65001 > nul
echo.
echo ======================================================
echo   Instalador ConvertirMD2Word para Windows
echo   Conversor Markdown a Word para Moodle
echo ======================================================
echo.

REM Verificar si Python está instalado
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python no esta instalado.
    echo.
    echo Por favor instala Python desde: https://www.python.org/downloads/
    echo Asegurate de marcar "Add Python to PATH" durante la instalacion.
    echo.
    pause
    exit /b 1
)

echo Python encontrado: 
python --version

REM Verificar pip
python -m pip --version > nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: pip no esta disponible.
    echo Reinstala Python con pip incluido.
    pause
    exit /b 1
)

echo pip encontrado: 
python -m pip --version

echo.
echo Instalando dependencias...
echo.

REM Actualizar pip
echo Actualizando pip...
python -m pip install --upgrade pip
if %errorlevel% neq 0 (
    echo ADVERTENCIA: No se pudo actualizar pip, continuando...
)

REM Instalar dependencias principales
echo.
echo Instalando pypandoc...
python -m pip install pypandoc
if %errorlevel% neq 0 (
    echo ERROR: No se pudo instalar pypandoc
    pause
    exit /b 1
)

echo.
echo Instalando python-docx...
python -m pip install python-docx
if %errorlevel% neq 0 (
    echo ERROR: No se pudo instalar python-docx
    pause
    exit /b 1
)

echo.
echo Instalando dependencias para correccion automatica de codificacion...
python -m pip install chardet ftfy unidecode
if %errorlevel% neq 0 (
    echo ADVERTENCIA: Algunas dependencias de correccion de codificacion no se pudieron instalar
    echo El programa funcionara con correcciones basicas
)

echo.
echo Instalando dependencias adicionales...
python -m pip install lxml pillow
if %errorlevel% neq 0 (
    echo ADVERTENCIA: Algunas dependencias adicionales no se pudieron instalar
)

REM Configurar Pandoc
echo.
echo Configurando Pandoc...
python -c "import pypandoc; print('Pandoc version:', pypandoc.get_pandoc_version())" 2>nul
if %errorlevel% neq 0 (
    echo Pandoc no encontrado, instalando automaticamente...
    python -c "import pypandoc; pypandoc.download_pandoc(); print('Pandoc instalado correctamente')"
    if %errorlevel% neq 0 (
        echo.
        echo ADVERTENCIA: No se pudo instalar Pandoc automaticamente.
        echo Por favor descarga e instala Pandoc manualmente desde:
        echo https://pandoc.org/installing.html
        echo.
        echo Despues de instalar Pandoc, el programa funcionara correctamente.
    )
) else (
    echo Pandoc ya esta instalado correctamente.
)

REM Crear requirements.txt
echo.
echo Creando requirements.txt...
echo # Dependencias para ConvertirMD2Word > requirements.txt
echo pypandoc>=1.11 >> requirements.txt
echo python-docx>=0.8.11 >> requirements.txt
echo lxml>=4.6.0 >> requirements.txt
echo pillow>=8.0.0 >> requirements.txt
echo # Dependencias para correccion automatica de codificacion >> requirements.txt
echo chardet>=4.0.0 >> requirements.txt
echo ftfy>=6.0.0 >> requirements.txt
echo unidecode>=1.3.0 >> requirements.txt

REM Crear archivo de prueba
echo.
echo Creando archivo de prueba...
(
echo # Capitulo 1: Introduccion al Curso
echo.
echo Este es el primer capitulo que se importara como un capitulo principal en Moodle Book.
echo.
echo ## 1.1 Objetivos del Capitulo
echo.
echo Este subcapitulo aparecera como un subcapitulo en Moodle.
echo.
echo ### Objetivos Especificos
echo.
echo Los titulos nivel 3 y superiores aparecen como texto con formato pero no crean nuevos capitulos.
echo.
echo - Objetivo 1
echo - Objetivo 2
echo - Objetivo 3
echo.
echo ## 1.2 Metodologia
echo.
echo Otro subcapitulo del Capitulo 1.
echo.
echo # Capitulo 2: Desarrollo del Tema
echo.
echo Este sera el segundo capitulo principal en Moodle Book.
echo.
echo ## 2.1 Fundamentos Teoricos
echo.
echo Subcapitulo del Capitulo 2.
echo.
echo **Texto en negrita** y *texto en cursiva* se preservan correctamente.
echo.
echo ## 2.2 Ejemplos Practicos
echo.
echo Otro subcapitulo con contenido.
echo.
echo ```python
echo # Los bloques de codigo tambien se mantienen
echo def ejemplo():
echo     return "Hola Moodle"
echo ```
echo.
echo # Capitulo 3: Conclusiones
echo.
echo Capitulo final que demuestra la estructura correcta para Moodle.
) > test_moodle.md

echo Archivo de prueba creado: test_moodle.md

REM Probar instalación
echo.
echo ======================================================
echo   PROBANDO LA INSTALACION
echo ======================================================
echo.

echo Validando archivo de prueba...
python ConvertirMD2Word.py --validar test_moodle.md
if %errorlevel% neq 0 (
    echo ERROR: La validacion fallo. Revisa la instalacion.
    pause
    exit /b 1
)

echo.
echo Convirtiendo archivo de prueba...
python ConvertirMD2Word.py test_moodle.md test_output.docx
if %errorlevel% neq 0 (
    echo ERROR: La conversion fallo. Revisa la instalacion.
    pause
    exit /b 1
)

if exist "test_output.docx" (
    echo.
    echo ======================================================
    echo   INSTALACION COMPLETADA EXITOSAMENTE
    echo ======================================================
    echo.
    echo ARCHIVOS CREADOS:
    echo   - test_output.docx - Archivo de prueba convertido
    echo   - test_moodle.md - Archivo Markdown de ejemplo
    echo   - requirements.txt - Lista de dependencias
    echo.
    echo OPTIMIZADO PARA PLUGIN DE IMPORTACION DE LIBROS DE MOODLE:
    echo   - Heading 1 (##^) -^> Capitulos
    echo   - Heading 2 (####^) -^> Subcapitulos
    echo   - Formato .docx compatible con Moodle 2.7+
    echo   - Imagenes en formato web embebidas
    echo.
    echo COMANDOS PRINCIPALES:
    echo   python ConvertirMD2Word.py archivo.md
    echo   python ConvertirMD2Word.py -d carpeta_markdown\
    echo   python ConvertirMD2Word.py --validar archivo.md
    echo   python ConvertirMD2Word.py --help
    echo.
    echo PASOS PARA USAR EN MOODLE:
    echo   1. Convierte tus archivos .md con este programa
    echo   2. En Moodle: Curso -^> Actividades -^> Libro
    echo   3. Crear nuevo libro
    echo   4. Configuracion -^> Importar capitulos
    echo   5. Subir tu archivo .docx convertido
    echo.
    echo Listo para convertir tus documentos para Moodle!
    echo.
) else (
    echo ERROR: No se pudo crear el archivo de prueba.
    echo Revisa que no haya errores en la instalacion.
)

echo Para obtener ayuda, ejecuta:
echo   python ConvertirMD2Word.py --help
echo.
pause