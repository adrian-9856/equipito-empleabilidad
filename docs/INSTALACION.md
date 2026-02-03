# 📋 Guía de Instalación - Sistema de Seguimiento de Graduados

Esta guía te llevará paso a paso para instalar y configurar el sistema de seguimiento de graduados con integración a KoboToolbox.

## 📌 Requisitos Previos

1. **Cuenta de Google** con acceso a Google Sheets
2. **Cuenta de KoboToolbox** con un formulario de graduados
3. **Token de API de KoboToolbox**
4. **URL de exportación CSV** del formulario

---

## 🚀 Paso 1: Crear la Hoja de Google Sheets

1. Ve a [Google Sheets](https://sheets.google.com)
2. Crea una nueva hoja de cálculo
3. Nómbrala: **"Seguimiento de Graduados - Equipito Empleabilidad"**

---

## 💻 Paso 2: Abrir el Editor de Google Apps Script

1. En tu hoja de Google Sheets, ve al menú superior
2. Haz clic en **Extensiones** → **Apps Script**
3. Se abrirá el editor de código
4. Elimina el código predeterminado (`function myFunction() { ... }`)

---

## 📝 Paso 3: Copiar los Archivos del Script

Debes crear los siguientes archivos en el editor de Apps Script:

### 3.1 Archivos de Código (.gs)

Crea cada archivo haciendo clic en el botón **"+"** junto a "Archivos" en el panel izquierdo:

1. **Code.gs** - Archivo principal
2. **KoboToolboxAPI.gs** - Integración con KoboToolbox
3. **SheetManager.gs** - Gestión de hojas
4. **FollowUpTracker.gs** - Sistema de seguimientos
5. **Config.gs** - Configuración
6. **Utils.gs** - Utilidades

### 3.2 Archivos HTML

Para crear archivos HTML:
1. Haz clic en el botón **"+"** junto a "Archivos"
2. Selecciona **"HTML"** en el tipo de archivo
3. Crea:
   - **FormularioClasificacion.html**
   - **ConfiguracionKobo.html**

### 3.3 Copiar el Código

Copia el contenido de cada archivo desde la carpeta `google-apps-script/` de este repositorio a los archivos correspondientes en el editor de Apps Script.

---

## ⚙️ Paso 4: Guardar y Configurar Permisos

1. Haz clic en el icono de **💾 Guardar** (o presiona `Ctrl+S`)
2. Nombra el proyecto: **"Sistema Seguimiento Graduados"**
3. Cierra el editor de Apps Script

---

## 🔄 Paso 5: Recargar la Hoja de Cálculo

1. Vuelve a tu hoja de Google Sheets
2. Recarga la página (presiona `F5` o `Ctrl+R`)
3. Deberías ver un nuevo menú en la barra superior: **📊 Seguimiento Graduados**

---

## 🔐 Paso 6: Autorizar el Script

La primera vez que uses el script:

1. Haz clic en **📊 Seguimiento Graduados** → **📋 Crear Estructura de Hojas**
2. Aparecerá un mensaje de autorización
3. Haz clic en **"Continuar"**
4. Selecciona tu cuenta de Google
5. Haz clic en **"Avanzado"**
6. Haz clic en **"Ir a Sistema Seguimiento Graduados (no seguro)"**
7. Revisa los permisos y haz clic en **"Permitir"**

**Nota:** El mensaje "no seguro" es normal para scripts personalizados. Es seguro autorizarlo.

---

## 🏗️ Paso 7: Crear la Estructura de Hojas

1. Ve a **📊 Seguimiento Graduados** → **📋 Crear Estructura de Hojas**
2. Confirma haciendo clic en **"Sí"**
3. El sistema creará automáticamente todas las hojas necesarias:
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

## 🔗 Paso 8: Obtener Credenciales de KoboToolbox

### 8.1 Obtener el Token de API

1. Inicia sesión en [KoboToolbox](https://kobo.humanitarianresponse.info/)
2. Ve a tu perfil (icono de usuario en la esquina superior derecha)
3. Selecciona **"Account Settings"**
4. Busca la sección **"API"**
5. Haz clic en **"Generate Token"** (si no tienes uno)
6. Copia el token generado (guárdalo en un lugar seguro)

### 8.2 Obtener la URL de Exportación CSV

1. Ve a tu proyecto de KoboToolbox
2. Abre el formulario de graduados
3. Haz clic en **"DATA"** en el menú superior
4. Selecciona **"Downloads"**
5. En la sección de exportación, busca **"API"**
6. Copia la URL de exportación CSV (debe verse algo así):
   ```
   https://kf.kobotoolbox.org/api/v2/assets/[TU_ID_ASSET]/export-settings/[TU_ID_EXPORT]/data.csv
   ```

---

## ⚙️ Paso 9: Configurar el Sistema

1. En Google Sheets, ve a **📊 Seguimiento Graduados** → **⚙️ Configurar Credenciales**
2. Se abrirá un formulario
3. Completa los campos:
   - **URL de Exportación CSV**: Pega la URL que copiaste de KoboToolbox
   - **Token de API**: Pega el token que copiaste
   - **Email para Notificaciones** (opcional): Tu email para recibir alertas
4. (Opcional) Activa las casillas:
   - ✅ **Activar notificaciones automáticas diarias**
   - ✅ **Sincronizar datos automáticamente cada hora**
5. Haz clic en **🔍 Probar Conexión** para verificar que funciona
6. Si la conexión es exitosa, haz clic en **Guardar Configuración**

---

## ✅ Paso 10: Importar Datos Iniciales

1. Ve a **📊 Seguimiento Graduados** → **🔄 Importar desde KoboToolbox**
2. El sistema comenzará a importar los datos
3. Espera a que termine el proceso
4. Verás un mensaje con el número de graduados importados

---

## 🎉 ¡Listo!

Tu sistema de seguimiento de graduados está instalado y configurado. Ahora puedes:

- ✅ Importar graduados desde KoboToolbox
- ✅ Clasificar graduados según el flujo de trabajo
- ✅ Programar seguimientos automáticos
- ✅ Recibir notificaciones de seguimientos pendientes

---

## 🔧 Solución de Problemas

### Error: "No se puede conectar a KoboToolbox"
- Verifica que la URL de exportación sea correcta
- Verifica que el token de API esté activo
- Asegúrate de tener permisos en el proyecto de KoboToolbox

### El menú no aparece
- Recarga la página de Google Sheets
- Verifica que hayas guardado todos los archivos en Apps Script
- Revisa que hayas autorizado el script correctamente

### No se importan datos
- Verifica que tu formulario de KoboToolbox tenga respuestas
- Revisa la consola de errores en Apps Script (Ver → Logs)
- Verifica la configuración de la URL y el token

---

## 📞 Soporte

Si necesitas ayuda adicional, revisa:
- [MANUAL_USO.md](./MANUAL_USO.md) - Guía de uso del sistema
- [FLUJO.md](./FLUJO.md) - Descripción del flujo de trabajo
- Los logs en Apps Script (Extensiones → Apps Script → Vista → Logs)
