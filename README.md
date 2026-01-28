# 🏢 Office LATAM - Instalador Automático 2019/2021

[![GitHub release](https://img.shields.io/github/release/dis-mal-skull/office-latam.svg)](https://github.com/dis-mal-skull/office-latam/releases)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.7+-blue.svg)](https://python.org)

## 📋 Descripción

**Office LATAM** es un instalador automático y personalizable para Microsoft Office 2019/2021 en español, diseñado específicamente para usuarios de América Latina. Permite seleccionar qué componentes instalar con una interfaz amigable y activación automática.

## ✨ Características Principales

- 🎯 **Selección Personalizada**: Elige solo los programas que necesitas
- 🌍 **Español LATAM**: Instalación completamente en español
- 🔄 **Dos Versiones**: Compatible con Office 2019 y 2021 LTSC
- ⚡ **Instalación Automática**: Descarga e instalación sin intervención manual
- 🔐 **Activación KMS**: Activación automática vía servidor KMS
- 📊 **Barra de Progreso**: Visualización del progreso de descarga
- 🧹 **Limpieza Automática**: Elimina archivos temporales al finalizar
- 🎨 **Interfaz Colorida**: Consola con colores y diseño moderno

## 🖥️ Programas Disponibles

1. **Excel** - Hojas de cálculo
2. **Word** - Procesador de texto  
3. **PowerPoint** - Presentaciones
4. **Outlook** - Cliente de correo
5. **Access** - Base de datos
6. **Publisher** - Diseño gráfico
7. **OneNote** - Bloc de notas digital
8. **Groove** - Música (heredado)
9. **Lync** - Mensajería (heredado)
10. **OneDrive** - Almacenamiento en la nube
11. **Teams** - Colaboración

## 🚀 Instalación Rápida

### Método 1: Ejecutable (Recomendado)
```bash
# Descargar y ejecutar
INSTALAR_OFFICE.bat
```

### Método 2: Manual (Avanzado)
```powershell
# Como administrador
python install_office.py
```

## 📋 Opciones de Selección

Durante la instalación puedes elegir:

- **Números separados por coma** (ej: `1,3,5`) - Selección personalizada
- **T** - Todos los programas
- **S** - Solo Excel (configuración original)
- **N** - Cancelar instalación

## ⚙️ Requisitos del Sistema

- **Windows 10/11** (64 bits)
- **Conexión a Internet** estable
- **Permisos de Administrador**
- **Python 3.7+** (se instala automáticamente si no está presente)
- **Espacio en disco**: 3GB mínimo (Excel solo) hasta 15GB (suite completa)

## 🔄 Proceso de Instalación

1. **Verificación de permisos** de administrador
2. **Instalación automática** de Python (si es necesario)
3. **Selección de versión** (2019 o 2021)
4. **Selección de programas** a instalar
5. **Limpieza** de instalaciones previas
6. **Descarga** de Office Deployment Tool
7. **Descarga** de componentes seleccionados
8. **Instalación** silenciosa
9. **Activación** automática KMS
10. **Limpieza final** de archivos temporales

## ⏱️ Tiempos Estimados

| Componentes | Tiempo | Espacio Requerido |
|-------------|--------|-------------------|
| Solo Excel | 15-20 min | ~3GB |
| 3-5 programas | 20-30 min | ~6-8GB |
| Suite completa | 30-45 min | ~12-15GB |

## 🔧 Solución de Problemas

### Problemas Comunes

**❌ "Error de permisos"**
- Ejecutar como administrador
- Desactivar temporalmente el antivirus

**❌ "Error de descarga"**
- Verificar conexión a internet
- Limpiar espacio en disco
- Reiniciar el instalador

**❌ "Error de activación"**
- Esperar 5-10 minutos después de instalación
- Ejecutar el script nuevamente
- Verificar conexión con servidor KMS

### Comandos Útiles

```bash
# Verificar instalación
cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /dstatus

# Reactivar manualmente
cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /act
```

## 🛡️ Seguridad y Privacidad

- ✅ Sin malware o virus
- ✅ Descargas oficiales de Microsoft
- ✅ Sin recolección de datos personales
- ✅ Código abierto y transparente
- ⚠️ Requiere desactivar antivirus temporalmente

## 📝 Notas Importacionales

- Este instalador utiliza **versiones VOLUME** de Office LTSC
- La activación se realiza mediante **servidor KMS** automático
- Compatible con **Windows 10/11** en español
- Puede ejecutarse múltiples veces para agregar componentes
- Los archivos se descargan de servidores **oficiales Microsoft**

## ☕ Donaciones

¿Te ayudó este proyecto? Considera una donación para mantenerlo actualizado:

[![PayPal](https://img.shields.io/badge/PayPal-Donate-blue.svg)](https://paypal.me/sternenfrost)
- **PayPal**: `sternenfrost@gmail.com`

Tu apoyo ayuda a mantener el proyecto activo y con actualizaciones regulares.

---

**⚠️ Descargo de Responsabilidad**: Este software es para uso educativo y personal. El usuario es responsable de cumplir con los términos de licencia de Microsoft.
