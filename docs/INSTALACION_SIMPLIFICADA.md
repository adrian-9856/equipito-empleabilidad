# 📋 Guía de Instalación Simplificada - Sistema de Seguimiento de Graduados

Esta es la guía de instalación simplificada con **solo 3 archivos** para copiar.

## 📌 Requisitos Previos

1. **Cuenta de Google** con acceso a Google Sheets
2. **Cuenta de KoboToolbox** con un formulario de graduados
3. **Token de API de KoboToolbox**
4. **URL de exportación CSV** del formulario

---

## 🚀 Instalación Rápida (3 Archivos)

### Paso 1: Crear la Hoja de Google Sheets

1. Ve a [Google Sheets](https://sheets.google.com)
2. Crea una nueva hoja de cálculo
3. Nómbrala: **"Seguimiento de Graduados - Equipito Empleabilidad"**

---

### Paso 2: Abrir el Editor de Google Apps Script

1. En tu hoja de Google Sheets, ve al menú superior
2. Haz clic en **Extensiones** → **Apps Script**
3. Se abrirá el editor de código
4. Elimina el código predeterminado (`function myFunction() { ... }`)

---

### Paso 3: Copiar los 3 Archivos del Sistema

Solo necesitas copiar **3 archivos**:

#### Archivo 1: Code.gs (Código Principal)

1. En el editor de Apps Script, ya tienes un archivo llamado **Code.gs**
2. Copia todo el contenido de: `google-apps-script-unified/Code.gs`
3. Pégalo en el archivo Code.gs del editor
4. Este archivo contiene **TODO el código del sistema** (más de 1,500 líneas)

#### Archivo 2: FormularioClasificacion.html

1. Haz clic en el botón **"+"** junto a "Archivos"
2. Selecciona **"HTML"**
3. Nómbralo: **FormularioClasificacion**
4. Copia el contenido de: `google-apps-script-unified/FormularioClasificacion.html`
5. Pégalo en el archivo HTML

#### Archivo 3: ConfiguracionKobo.html

1. Haz clic en el botón **"+"** junto a "Archivos"
2. Selecciona **"HTML"**
3. Nómbralo: **ConfiguracionKobo**
4. Copia el contenido de: `google-apps-script-unified/ConfiguracionKobo.html`
5. Pégalo en el archivo HTML

---

### Paso 4: Guardar el Proyecto

1. Haz clic en el icono de **💾 Guardar** (o presiona `Ctrl+S`)
2. Nombra el proyecto: **"Sistema Seguimiento Graduados"**
3. Cierra el editor de Apps Script

---

### Paso 5: Recargar y Autorizar

1. Vuelve a tu hoja de Google Sheets
2. Recarga la página (presiona `F5` o `Ctrl+R`)
3. Deberías ver un nuevo menú: **📊 Seguimiento Graduados**
4. Haz clic en **📊 Seguimiento Graduados** → **📋 Crear Estructura de Hojas**
5. Aparecerá un mensaje de autorización:
   - Haz clic en **"Continuar"**
   - Selecciona tu cuenta de Google
   - Haz clic en **"Avanzado"**
   - Haz clic en **"Ir a Sistema Seguimiento Graduados (no seguro)"**
   - Haz clic en **"Permitir"**

---

### Paso 6: Crear la Estructura de Hojas

1. Ve a **📊 Seguimiento Graduados** → **📋 Crear Estructura de Hojas**
2. Confirma haciendo clic en **"Sí"**
3. El sistema creará automáticamente 10 hojas:
   - Graduados
   - Habilidades
   - GlobalGorros
   - Conexión/Relación
   - Por su Cuenta
   - No Hace Falta
   - Empleados
   - Seguimientos
   - Reportes
   - Configuración

---

### Paso 7: Configurar Credenciales de KoboToolbox

#### 7.1 Obtener el Token de API

1. Inicia sesión en [KoboToolbox](https://kobo.humanitarianresponse.info/)
2. Ve a tu perfil → **"Account Settings"**
3. Busca la sección **"API"**
4. Haz clic en **"Generate Token"** (si no tienes uno)
5. Copia el token generado

#### 7.2 Obtener la URL de Exportación CSV

1. Ve a tu proyecto de KoboToolbox
2. Abre el formulario de graduados
3. Haz clic en **"DATA"** → **"Downloads"**
4. En la sección de exportación, busca **"API"**
5. Copia la URL de exportación CSV

#### 7.3 Configurar en Google Sheets

1. En Google Sheets, ve a **📊 Seguimiento Graduados** → **⚙️ Configurar Credenciales**
2. Completa:
   - **URL de Exportación CSV**: Pega la URL de KoboToolbox
   - **Token de API**: Pega el token
   - **Email para Notificaciones** (opcional): Tu email
3. Haz clic en **🔍 Probar Conexión**
4. Si funciona, haz clic en **Guardar Configuración**

---

### Paso 8: Importar Datos Iniciales

1. Ve a **📊 Seguimiento Graduados** → **🔄 Importar desde KoboToolbox**
2. El sistema importará los datos
3. Verás un mensaje con el número de graduados importados

---

## 🎉 ¡Listo!

Tu sistema está completamente instalado y funcionando. Ahora puedes:

- ✅ Importar graduados desde KoboToolbox
- ✅ Clasificar graduados con menús desplegables
- ✅ Programar seguimientos automáticos
- ✅ Recibir notificaciones por email

---

## 📊 Estructura de Archivos

```
Solo 3 archivos:
└── google-apps-script-unified/
    ├── Code.gs                        (TODO el código en 1 archivo)
    ├── FormularioClasificacion.html   (Interfaz de clasificación)
    └── ConfiguracionKobo.html         (Interfaz de configuración)
```

---

## 🔧 Solución de Problemas

### Error: "No se puede conectar a KoboToolbox"
- Verifica que la URL de exportación sea correcta
- Verifica que el token de API esté activo
- Asegúrate de tener permisos en el proyecto

### El menú no aparece
- Recarga la página de Google Sheets
- Verifica que hayas guardado Code.gs correctamente
- Revisa que hayas autorizado el script

### No se importan datos
- Verifica que tu formulario tenga respuestas
- Revisa los logs: Extensiones → Apps Script → Vista → Logs
- Verifica la configuración de URL y token

---

## 📞 Siguiente Paso

Lee el [Manual de Uso](./MANUAL_USO.md) para aprender a usar el sistema día a día.
