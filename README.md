# 📊 Sistema de Seguimiento de Graduados - Equipito Empleabilidad

Sistema automatizado de seguimiento de graduados con integración a **KoboToolbox** y **Google Sheets**, desarrollado con **Google Apps Script**.

## 🎯 Descripción

Este sistema permite gestionar el seguimiento completo de graduados desde la formación hasta el empleo, automatizando:

- ✅ Importación automática de datos desde KoboToolbox
- ✅ Clasificación de graduados según flujo de trabajo
- ✅ Seguimientos programados automáticos (1 semana, 3 meses, 6 meses)
- ✅ Notificaciones por email de seguimientos pendientes
- ✅ Distribución de datos en hojas específicas por clasificación
- ✅ Menús desplegables para facilitar el uso
- ✅ Reportes y estadísticas de empleabilidad

## 🌟 Características Principales

### 📥 Integración con KoboToolbox
- Importación automática o manual de datos
- Sincronización cada hora (opcional)
- Validación automática de datos
- Sin duplicados

### 🗂️ Clasificación de Graduados
Sistema de clasificación con menús desplegables:
- 🎯 **Habilidades**: Desarrollo de competencias específicas
- 🌍 **GlobalGorros**: Participación en proyectos
- 🤝 **Conexión/Relación**: Mentoría con líderes
- 👤 **Por su Cuenta**: Búsqueda independiente
- ⛔ **No Hace Falta**: Sin seguimiento activo
- 💼 **Empleado**: ¡Meta alcanzada!

### 📞 Seguimientos Automáticos
Para graduados empleados:
- **Llamada 1**: A la semana (verificación inicial)
- **Llamada 2**: A los 3 meses (evaluación de adaptación)
- **Llamada 3**: A los 6 meses (permanencia y feedback)

### 🔔 Notificaciones
- Email diario con seguimientos pendientes
- Alertas de nuevos graduados importados
- Recordatorios de llamadas vencidas

### 📊 Estructura de Hojas
- **Graduados**: Registro central de todos los graduados
- **Habilidades, GlobalGorros, Conexión, etc.**: Vistas específicas por clasificación
- **Empleados**: Registro de empleados exitosos
- **Seguimientos**: Gestión de llamadas programadas
- **Reportes**: Estadísticas y métricas

## 🚀 Instalación

### Requisitos Previos
- Cuenta de Google (con acceso a Google Sheets)
- Cuenta de KoboToolbox con formulario de graduados
- Token de API de KoboToolbox
- URL de exportación CSV del formulario

### Pasos de Instalación

1. **Crear hoja de Google Sheets**
   - Crea una nueva hoja en Google Sheets
   - Nómbrala: "Seguimiento de Graduados - Equipito Empleabilidad"

2. **Abrir el editor de Apps Script**
   - Ve a: Extensiones → Apps Script

3. **Copiar los archivos del script**
   - Crea los siguientes archivos en Apps Script:
     - `Code.gs`
     - `KoboToolboxAPI.gs`
     - `SheetManager.gs`
     - `FollowUpTracker.gs`
     - `Config.gs`
     - `Utils.gs`
     - `FormularioClasificacion.html`
     - `ConfiguracionKobo.html`
   - Copia el contenido de cada archivo desde la carpeta `google-apps-script/`

4. **Guardar y autorizar**
   - Guarda el proyecto
   - Recarga la hoja de Google Sheets
   - Autoriza el script cuando se solicite

5. **Crear estructura de hojas**
   - Ve al menú: 📊 Seguimiento Graduados → 📋 Crear Estructura de Hojas

6. **Configurar credenciales**
   - Ve al menú: 📊 Seguimiento Graduados → ⚙️ Configurar Credenciales
   - Ingresa:
     - URL de exportación CSV de KoboToolbox
     - Token de API de KoboToolbox
     - Email para notificaciones (opcional)

7. **Importar datos iniciales**
   - Ve al menú: 📊 Seguimiento Graduados → 🔄 Importar desde KoboToolbox

**📖 Guía detallada:** Ver [docs/INSTALACION.md](docs/INSTALACION.md)

## 📖 Documentación

- **[Guía de Instalación](docs/INSTALACION.md)**: Instalación paso a paso completa
- **[Manual de Uso](docs/MANUAL_USO.md)**: Cómo usar el sistema día a día
- **[Flujo de Trabajo](docs/FLUJO.md)**: Descripción detallada del flujo de seguimiento

## 🎮 Uso Rápido

### Menú Principal
📊 **Seguimiento Graduados**
- 🔄 **Importar desde KoboToolbox**: Obtener nuevos graduados
- 📝 **Clasificar Graduados**: Asignar clasificación con formulario
- 📞 **Ver Seguimientos Pendientes**: Lista de llamadas pendientes
- ⚙️ **Configurar Credenciales**: Configurar KoboToolbox y notificaciones
- 📋 **Crear Estructura de Hojas**: Crear todas las hojas necesarias

### Flujo Diario
1. Revisar seguimientos pendientes
2. Realizar llamadas necesarias
3. Marcar seguimientos como realizados
4. Clasificar nuevos graduados (si los hay)

### Flujo Semanal
1. Importar nuevos graduados (si no está en automático)
2. Clasificar todos los graduados sin clasificar
3. Revisar progreso general

## 🏗️ Estructura del Proyecto

```
equipito-empleabilidad/
├── google-apps-script/
│   ├── Code.gs                           # Archivo principal
│   ├── KoboToolboxAPI.gs                 # Integración con KoboToolbox
│   ├── SheetManager.gs                   # Gestión de hojas
│   ├── FollowUpTracker.gs                # Sistema de seguimientos
│   ├── Config.gs                         # Configuración
│   ├── Utils.gs                          # Utilidades
│   ├── FormularioClasificacion.html      # Interfaz de clasificación
│   └── ConfiguracionKobo.html            # Interfaz de configuración
├── docs/
│   ├── INSTALACION.md                    # Guía de instalación
│   ├── MANUAL_USO.md                     # Manual de uso
│   └── FLUJO.md                          # Descripción del flujo
└── README.md                             # Este archivo
```

## 🔧 Tecnologías Utilizadas

- **Google Apps Script**: Automatización y lógica del sistema
- **Google Sheets**: Base de datos y interfaz de usuario
- **KoboToolbox API**: Fuente de datos de graduados
- **HTML/CSS/JavaScript**: Interfaces de usuario personalizadas
- **MailApp (Gmail)**: Sistema de notificaciones

## 📊 Métricas y KPIs

El sistema permite calcular:
- **Tasa de empleabilidad**: (Empleados / Total Graduados) × 100
- **Tiempo promedio hasta el empleo**
- **Tasa de retención a 6 meses**
- **Seguimientos realizados vs pendientes**
- **Distribución por clasificación**

## 🤝 Contribuciones

Este es un proyecto de código abierto. Las contribuciones son bienvenidas:

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. Commit tus cambios (`git commit -m 'Agregar nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Abre un Pull Request

## 📝 Licencia

Este proyecto está bajo la licencia MIT. Ver el archivo `LICENSE` para más detalles.

## 👥 Autores

- **Equipito Empleabilidad** - Sistema de Seguimiento de Graduados

## 🆘 Soporte

Si tienes problemas o preguntas:

1. Revisa la [documentación completa](docs/)
2. Verifica los logs en Apps Script (Extensiones → Apps Script → Vista → Logs)
3. Abre un issue en GitHub

## 📞 Contacto

Para más información sobre el proyecto o el programa Equipito Empleabilidad, contacta con el equipo.

---

**Hecho con ❤️ para mejorar la empleabilidad de los graduados**