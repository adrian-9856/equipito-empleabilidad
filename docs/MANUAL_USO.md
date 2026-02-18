# 📖 Manual de Uso - Sistema de Seguimiento de Graduados

Esta guía explica cómo usar el sistema día a día para gestionar el seguimiento de graduados.

---

## 📊 Menú Principal

El sistema agrega un menú personalizado en tu hoja de Google Sheets:

**📊 Seguimiento Graduados**
- 🔄 Importar desde KoboToolbox
- 📝 Clasificar Graduados
- 📞 Ver Seguimientos Pendientes
- ⚙️ Configurar Credenciales
- 📋 Crear Estructura de Hojas

---

## 🔄 1. Importar Graduados desde KoboToolbox

### ¿Cuándo usar?
- Al inicio para cargar todos los graduados
- Periódicamente para obtener nuevos graduados (si no está activada la sincronización automática)

### Pasos:
1. Haz clic en **📊 Seguimiento Graduados** → **🔄 Importar desde KoboToolbox**
2. Aparecerá un mensaje indicando que está importando
3. Espera a que termine (puede tardar unos segundos/minutos según la cantidad de datos)
4. Verás un resumen: "Se importaron X graduados nuevos. Total de graduados: Y"

### Resultado:
- Los nuevos graduados se agregan a la hoja **"Graduados"**
- Los graduados existentes se actualizan si hay cambios
- **No se duplican** los graduados

---

## 📝 2. Clasificar Graduados

### ¿Cuándo usar?
- Después de importar nuevos graduados
- Para organizar graduados según el flujo de trabajo

### Pasos:
1. Haz clic en **📊 Seguimiento Graduados** → **📝 Clasificar Graduados**
2. Se abrirá un formulario con:
   - **Graduado**: Desplegable con graduados sin clasificar
   - **Clasificación**: Opciones del flujo (ver más abajo)
   - **Campos adicionales**: Según el tipo de clasificación

3. Selecciona el graduado
4. Haz clic en el tipo de clasificación que corresponde
5. Completa los campos adicionales que aparezcan
6. Agrega notas si es necesario
7. Haz clic en **"Guardar Clasificación"**

### Tipos de Clasificación:

#### 🎯 **Habilidades**
Para graduados que necesitan desarrollar o potenciar habilidades específicas.

**Campos adicionales:**
- Habilidad Principal (ej: Programación, Diseño)
- Nivel (Básico, Intermedio, Avanzado)
- Notas

**Resultado:** Se copia a la hoja "Habilidades"

---

#### 🌍 **GlobalGorros**
Para graduados que participan en proyectos de GlobalGorros.

**Campos adicionales:**
- Proyecto Asignado
- Responsable del proyecto
- Notas

**Resultado:** Se copia a la hoja "GlobalGorros"

---

#### 🤝 **Conexión/Relación**
Para graduados conectados con líderes o mentores.

**Campos adicionales:**
- Líder Asignado
- Tipo de Relación (Mentoría, Referencia, Networking)
- Notas

**Resultado:** Se copia a la hoja "Conexión/Relación"

---

#### 👤 **Por su Cuenta**
Para graduados que buscan empleo de forma independiente.

**Campos adicionales:**
- Actividad (¿Qué está haciendo?)
- Notas

**Resultado:** Se copia a la hoja "Por su Cuenta"

---

#### ⛔ **No Hace Falta**
Para graduados que no requieren seguimiento activo.

**Campos adicionales:**
- Razón (¿Por qué no requiere seguimiento?)
- Notas

**Resultado:** Se copia a la hoja "No Hace Falta"

---

#### 💼 **Empleado** 🎉
Para graduados que ya consiguieron empleo (¡La meta!).

**Campos adicionales:**
- Empresa
- Puesto
- Salario (opcional)
- Notas

**Resultado:**
- Se copia a la hoja "Empleados"
- **Se programan automáticamente 3 seguimientos:**
  - 📞 Llamada 1: A la semana
  - 📞 Llamada 2: A los 3 meses
  - 📞 Llamada 3: A los 6 meses

---

## 📞 3. Ver y Gestionar Seguimientos Pendientes

### ¿Cuándo usar?
- Diariamente para revisar llamadas pendientes
- Después de clasificar graduados como "Empleado"

### Pasos:
1. Haz clic en **📊 Seguimiento Graduados** → **📞 Ver Seguimientos Pendientes**
2. Se abrirá una ventana con la lista de seguimientos que están pendientes
3. Si no hay seguimientos pendientes, verás: "✅ Todo al Día"

### Información mostrada:
- Nombre del graduado
- Tipo de seguimiento (Semanal, 3 meses, 6 meses)
- Fecha programada
- Notas

### Marcar como Realizado:
1. Ve a la hoja **"Seguimientos"**
2. Busca la fila del seguimiento realizado
3. Completa las columnas:
   - **Fecha Realizada**: Fecha en que hiciste la llamada
   - **Estado**: Cambia a "Realizado"
   - **Resultado**: Resumen de la llamada (ej: "Todo bien, sigue en la empresa")
   - **Próximo Paso**: Si hay alguna acción pendiente

**Tip:** También puedes editar directamente en la hoja "Seguimientos"

---

## 📋 4. Trabajar con las Hojas

### Hoja "Graduados"
**Uso:** Contiene todos los graduados importados de KoboToolbox

**Columnas principales:**
- ID, Nombre, Teléfono, Email
- Fecha Entrevista, Formación
- Habilidades, Experiencia
- **Clasificación** (se llena cuando clasificas)
- **Empleado** (Sí/No)
- **Próxima Llamada**
- Notas

**Puedes:**
- Filtrar por clasificación
- Buscar graduados específicos
- Ver toda la información en un solo lugar

---

### Hojas de Clasificación
**Habilidades, GlobalGorros, Conexión/Relación, etc.**

**Uso:** Cada hoja contiene solo los graduados de ese tipo

**Ventajas:**
- Vista enfocada en cada grupo
- Fácil de compartir con responsables específicos
- Permite agregar información adicional específica

---

### Hoja "Seguimientos"
**Uso:** Gestión de llamadas programadas

**Columnas:**
- ID Graduado, Nombre
- Tipo Seguimiento
- Fecha Programada
- Fecha Realizada
- Estado (Pendiente/Realizado)
- Resultado
- Notas
- Próximo Paso

**Acciones:**
- Ordenar por fecha para ver próximas llamadas
- Filtrar por estado (Pendiente/Realizado)
- Marcar como realizado después de cada llamada

---

### Hoja "Reportes"
**Uso:** Crear reportes y estadísticas

**Puedes agregar:**
- Número de graduados por clasificación
- Tasa de empleabilidad
- Estadísticas mensuales
- Gráficos personalizados

---

## ⚙️ 5. Configuración Avanzada

### Sincronización Automática
Si activaste la sincronización automática:
- El sistema importa datos de KoboToolbox **cada hora**
- No necesitas hacer importación manual
- Recibirás notificación si hay graduados nuevos

### Notificaciones por Email
Si activaste las notificaciones:
- Recibirás un email **diario a las 9 AM** si hay seguimientos pendientes
- Incluye la lista de llamadas pendientes
- Te notifica cuando hay nuevos graduados importados

---

## 💡 Mejores Prácticas

### Flujo de Trabajo Recomendado

**Diario:**
1. Revisar seguimientos pendientes
2. Realizar llamadas necesarias
3. Marcar seguimientos como realizados

**Semanal:**
1. Importar nuevos graduados (si no está en automático)
2. Clasificar graduados sin clasificar
3. Revisar progreso general

**Mensual:**
1. Generar reportes de estadísticas
2. Revisar graduados "Por su Cuenta" para ver progreso
3. Actualizar información de empleados si hay cambios

---

## 🔍 Consultas y Filtros Útiles

### En la hoja "Graduados":
- **Filtrar por clasificación:** Habilidades, GlobalGorros, etc.
- **Buscar sin clasificar:** Clasificación = vacío
- **Ver solo empleados:** Columna "Empleado" = Sí

### En la hoja "Seguimientos":
- **Ver pendientes de hoy:** Fecha Programada <= HOY() y Estado = Pendiente
- **Ver seguimientos vencidos:** Fecha Programada < HOY() y Estado = Pendiente
- **Estadísticas por tipo:** Contar por "Tipo Seguimiento"

---

## 🎯 Consejos

1. **Clasifica regularmente:** No dejes acumular graduados sin clasificar
2. **Sé consistente:** Usa las mismas categorías y formatos
3. **Agrega notas:** La información adicional es valiosa
4. **Revisa seguimientos:** No dejes pasar las fechas programadas
5. **Actualiza información:** Si algo cambia, actualízalo en las hojas
6. **Haz respaldos:** Descarga copias periódicas por seguridad

---

## 📱 Acceso Móvil

Puedes usar el sistema desde tu celular:
1. Abre Google Sheets en tu navegador móvil o app
2. Busca tu hoja "Seguimiento de Graduados"
3. Puedes ver y editar datos
4. El menú "📊 Seguimiento Graduados" también funciona en móvil

---

## ❓ Preguntas Frecuentes

**¿Puedo editar manualmente los datos?**
Sí, puedes editar directamente en las hojas si es necesario.

**¿Se pierden los datos si importo de nuevo?**
No, el sistema actualiza sin duplicar.

**¿Puedo cambiar la clasificación de un graduado?**
Sí, puedes reclasificarlo usando el formulario o editando manualmente.

**¿Cómo agrego más campos?**
Puedes agregar columnas adicionales en cualquier hoja. El sistema respeta columnas extras.

**¿Puedo compartir la hoja?**
Sí, usa el botón "Compartir" de Google Sheets normalmente.

---

## 📞 Soporte

Para más información:
- [INSTALACION.md](./INSTALACION.md) - Guía de instalación
- [FLUJO.md](./FLUJO.md) - Descripción del flujo de trabajo
- Revisa los logs en Apps Script si hay errores
