#!/bin/bash
# Instalador ConvertirMD2Word para Linux

echo "======================================================"
echo "  Instalador ConvertirMD2Word para Linux"
echo "  Conversor Markdown a Word para Moodle"
echo "======================================================"
echo

# Detectar distribución
if [ -f /etc/os-release ]; then
    . /etc/os-release
    DISTRO=$ID
    VERSION=$VERSION_ID
elif type lsb_release >/dev/null 2>&1; then
    DISTRO=$(lsb_release -si | tr '[:upper:]' '[:lower:]')
    VERSION=$(lsb_release -sr)
elif [ -f /etc/redhat-release ]; then
    DISTRO="centos"
else
    DISTRO="unknown"
fi

echo "Sistema detectado: $DISTRO $VERSION"
echo

# Función para instalar dependencias del sistema
install_system_deps() {
    case $DISTRO in
        ubuntu|debian|linuxmint)
            echo "Instalando dependencias del sistema para Ubuntu/Debian..."
            sudo apt update
            if ! command -v python3 &> /dev/null; then
                sudo apt install -y python3
            fi
            if ! command -v pip3 &> /dev/null; then
                sudo apt install -y python3-pip
            fi
            if ! command -v pandoc &> /dev/null; then
                sudo apt install -y pandoc
            fi
            ;;
        centos|rhel)
            echo "Instalando dependencias del sistema para CentOS/RHEL..."
            if command -v dnf &> /dev/null; then
                sudo dnf install -y python3 python3-pip pandoc
            else
                sudo yum install -y python3 python3-pip
                # Para pandoc en versiones antiguas, usar EPEL
                if ! command -v pandoc &> /dev/null; then
                    echo "Instalando EPEL para pandoc..."
                    sudo yum install -y epel-release
                    sudo yum install -y pandoc
                fi
            fi
            ;;
        fedora)
            echo "Instalando dependencias del sistema para Fedora..."
            sudo dnf install -y python3 python3-pip pandoc
            ;;
        arch|manjaro)
            echo "Instalando dependencias del sistema para Arch Linux..."
            sudo pacman -S --needed python python-pip pandoc
            ;;
        opensuse*)
            echo "Instalando dependencias del sistema para openSUSE..."
            sudo zypper install -y python3 python3-pip pandoc
            ;;
        *)
            echo "Distribución no reconocida. Instalación manual requerida."
            echo "Por favor instala: python3, python3-pip, pandoc"
            echo "Continuando con la instalación de dependencias Python..."
            ;;
    esac
}

# Verificar Python3
if ! command -v python3 &> /dev/null; then
    echo "Python3 no encontrado. Instalando dependencias del sistema..."
    install_system_deps
fi

if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python3 no está instalado y no se pudo instalar automáticamente."
    echo "Por favor instala Python3 manualmente y vuelve a ejecutar este script."
    exit 1
fi

echo "✓ Python3 encontrado: $(python3 --version)"

# Verificar pip3
if ! command -v pip3 &> /dev/null; then
    echo "pip3 no encontrado. Intentando instalar..."
    case $DISTRO in
        ubuntu|debian|linuxmint)
            sudo apt install -y python3-pip
            ;;
        centos|rhel|fedora)
            if command -v dnf &> /dev/null; then
                sudo dnf install -y python3-pip
            else
                sudo yum install -y python3-pip
            fi
            ;;
        arch|manjaro)
            sudo pacman -S --needed python-pip
            ;;
        opensuse*)
            sudo zypper install -y python3-pip
            ;;
    esac
fi

if ! command -v pip3 &> /dev/null; then
    echo "ERROR: pip3 no está disponible."
    echo "Por favor instala pip3 manualmente."
    exit 1
fi

echo "✓ pip3 encontrado: $(pip3 --version)"

# Verificar Pandoc
if ! command -v pandoc &> /dev/null; then
    echo "Pandoc no encontrado. Instalando..."
    case $DISTRO in
        ubuntu|debian|linuxmint)
            sudo apt install -y pandoc
            ;;
        centos|rhel)
            if command -v dnf &> /dev/null; then
                sudo dnf install -y pandoc
            else
                sudo yum install -y epel-release
                sudo yum install -y pandoc
            fi
            ;;
        fedora)
            sudo dnf install -y pandoc
            ;;
        arch|manjaro)
            sudo pacman -S --needed pandoc
            ;;
        opensuse*)
            sudo zypper install -y pandoc
            ;;
        *)
            echo "No se pudo instalar pandoc automáticamente."
            echo "Puedes instalarlo manualmente o el programa lo instalará automáticamente."
            ;;
    esac
fi

if command -v pandoc &> /dev/null; then
    echo "✓ Pandoc encontrado: $(pandoc --version | head -n1)"
else
    echo "⚠ Pandoc no encontrado, se instalará automáticamente"
fi

# Crear entorno virtual (opcional pero recomendado)
read -p "¿Crear entorno virtual Python? (recomendado) [s/n]: " crear_venv
if [[ $crear_venv == "s" || $crear_venv == "S" || $crear_venv == "y" || $crear_venv == "Y" ]]; then
    echo "Creando entorno virtual..."
    python3 -m venv venv_convertir_md2word
    source venv_convertir_md2word/bin/activate
    echo "✓ Entorno virtual activado"
    PYTHON_CMD="python"
    PIP_CMD="pip"
else
    PYTHON_CMD="python3"
    PIP_CMD="pip3"
fi

echo
echo "Instalando dependencias Python..."
echo

# Actualizar pip
echo "Actualizando pip..."
$PIP_CMD install --upgrade pip

# Instalar dependencias principales
echo
echo "Instalando pypandoc..."
$PIP_CMD install pypandoc
if [ $? -ne 0 ]; then
    echo "ERROR: No se pudo instalar pypandoc"
    exit 1
fi

echo
echo "Instalando python-docx..."
$PIP_CMD install python-docx
if [ $? -ne 0 ]; then
    echo "ERROR: No se pudo instalar python-docx"
    exit 1
fi

echo
echo "Instalando dependencias para correccion automatica de codificacion..."
$PIP_CMD install chardet ftfy unidecode
if [ $? -ne 0 ]; then
    echo "ADVERTENCIA: Algunas dependencias de correccion de codificacion no se pudieron instalar"
    echo "El programa funcionara con correcciones basicas"
fi

echo
echo "Instalando dependencias adicionales..."
$PIP_CMD install lxml pillow
if [ $? -ne 0 ]; then
    echo "ADVERTENCIA: Algunas dependencias adicionales no se pudieron instalar"
fi

# Configurar Pandoc
echo
echo "Configurando Pandoc..."
$PYTHON_CMD -c "import pypandoc; print('Pandoc version:', pypandoc.get_pandoc_version())" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Pandoc no encontrado, instalando automáticamente..."
    $PYTHON_CMD -c "import pypandoc; pypandoc.download_pandoc(); print('Pandoc instalado correctamente')"
    if [ $? -ne 0 ]; then
        echo
        echo "ADVERTENCIA: No se pudo instalar Pandoc automáticamente."
        echo "Instala Pandoc manualmente con:"
        case $DISTRO in
            ubuntu|debian|linuxmint)
                echo "  sudo apt install pandoc"
                ;;
            centos|rhel|fedora)
                echo "  sudo dnf install pandoc"
                ;;
            arch|manjaro)
                echo "  sudo pacman -S pandoc"
                ;;
            *)
                echo "  Consulta https://pandoc.org/installing.html"
                ;;
        esac
        echo
    fi
else
    echo "✓ Pandoc configurado correctamente"
fi

# Crear requirements.txt
echo
echo "Creando requirements.txt..."
cat > requirements.txt << 'EOF'
# Dependencias para ConvertirMD2Word
pypandoc>=1.11
python-docx>=0.8.11
lxml>=4.6.0
pillow>=8.0.0
# Dependencias para correccion automatica de codificacion
chardet>=4.0.0
ftfy>=6.0.0
unidecode>=1.3.0
EOF

echo "✓ requirements.txt creado"

# Hacer ejecutable el script principal
if [ -f "ConvertirMD2Word.py" ]; then
    chmod +x ConvertirMD2Word.py
    echo "✓ ConvertirMD2Word.py configurado como ejecutable"
fi

# Crear archivo de prueba
echo
echo "Creando archivo de prueba..."
cat > test_moodle.md << 'EOF'
# Capitulo 1: Introduccion al Curso

Este es el primer capitulo que se importara como un capitulo principal en Moodle Book.

## 1.1 Objetivos del Capitulo

Este subcapitulo aparecera como un subcapitulo en Moodle.

### Objetivos Especificos

Los titulos nivel 3 y superiores aparecen como texto con formato pero no crean nuevos capitulos.

- Objetivo 1
- Objetivo 2
- Objetivo 3

## 1.2 Metodologia

Otro subcapitulo del Capitulo 1.

# Capitulo 2: Desarrollo del Tema

Este sera el segundo capitulo principal en Moodle Book.

## 2.1 Fundamentos Teoricos

Subcapitulo del Capitulo 2.

**Texto en negrita** y *texto en cursiva* se preservan correctamente.

## 2.2 Ejemplos Practicos

Otro subcapitulo con contenido.

```python
# Los bloques de codigo tambien se mantienen
def ejemplo():
    return "Hola Moodle"
```

# Capitulo 3: Conclusiones

Capitulo final que demuestra la estructura correcta para Moodle.
EOF

echo "✓ Archivo de prueba creado: test_moodle.md"

# Probar instalación
echo
echo "======================================================"
echo "  PROBANDO LA INSTALACION"
echo "======================================================"
echo

echo "Validando archivo de prueba..."
$PYTHON_CMD ConvertirMD2Word.py --validar test_moodle.md
if [ $? -ne 0 ]; then
    echo "ERROR: La validación falló. Revisa la instalación."
    exit 1
fi

echo
echo "Convirtiendo archivo de prueba..."
$PYTHON_CMD ConvertirMD2Word.py test_moodle.md test_output.docx
if [ $? -ne 0 ]; then
    echo "ERROR: La conversión falló. Revisa la instalación."
    exit 1
fi

if [ -f "test_output.docx" ]; then
    echo
    echo "======================================================"
    echo "  INSTALACION COMPLETADA EXITOSAMENTE"
    echo "======================================================"
    echo
    echo "ARCHIVOS CREADOS:"
    echo "  ✓ test_output.docx - Archivo de prueba convertido"
    echo "  ✓ test_moodle.md - Archivo Markdown de ejemplo"
    echo "  ✓ requirements.txt - Lista de dependencias"
    echo
    echo "OPTIMIZADO PARA PLUGIN DE IMPORTACION DE LIBROS DE MOODLE:"
    echo "  ✓ Heading 1 (#) -> Capitulos"
    echo "  ✓ Heading 2 (##) -> Subcapitulos"
    echo "  ✓ Formato .docx compatible con Moodle 2.7+"
    echo "  ✓ Imagenes en formato web embebidas"
    echo
    echo "COMANDOS PRINCIPALES:"
    echo "  $PYTHON_CMD ConvertirMD2Word.py archivo.md"
    echo "  $PYTHON_CMD ConvertirMD2Word.py -d carpeta_markdown/"
    echo "  $PYTHON_CMD ConvertirMD2Word.py --validar archivo.md"
    echo "  $PYTHON_CMD ConvertirMD2Word.py --help"
    echo
    echo "PASOS PARA USAR EN MOODLE:"
    echo "  1. Convierte tus archivos .md con este programa"
    echo "  2. En Moodle: Curso -> Actividades -> Libro"
    echo "  3. Crear nuevo libro"
    echo "  4. Configuracion -> Importar capitulos"
    echo "  5. Subir tu archivo .docx convertido"
    echo

    if [[ $crear_venv == "s" || $crear_venv == "S" || $crear_venv == "y" || $crear_venv == "Y" ]]; then
        echo "PARA ACTIVAR EL ENTORNO VIRTUAL EN EL FUTURO:"
        echo "  source venv_convertir_md2word/bin/activate"
        echo
    fi
    
    echo "¡Listo para convertir tus documentos para Moodle!"
    echo
else
    echo "ERROR: No se pudo crear el archivo de prueba."
    echo "Revisa que no haya errores en la instalación."
fi

echo "Para obtener ayuda, ejecuta:"
echo "  $PYTHON_CMD ConvertirMD2Word.py --help"