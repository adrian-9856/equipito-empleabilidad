# 🔄 Flujo de Trabajo - Seguimiento de Graduados

Este documento describe el flujo completo del proceso de seguimiento de graduados, desde la entrevista hasta el empleo.

---

## 📊 Diagrama del Flujo

```
┌──────────────────┐
│    GRADUADOS     │ ← Datos desde KoboToolbox
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│   Entrevista /   │
│   Seguimiento    │
│     General      │
└────────┬─────────┘
         │
    ┌────┴────┬─────────────────┬─────────────────┬──────────────────┐
    │         │                 │                 │                  │
    ▼         ▼                 ▼                 ▼                  ▼
┌────────┐ ┌───────────┐ ┌────────────────┐ ┌────────────┐ ┌──────────────┐
│Aliados │ │Plataforma │ │  Conexiones    │ │  Por su    │ │Busca Trabajo │
│        │ │           │ │   Laborales    │ │   Cuenta   │ │    (Fito)    │
└───┬────┘ └─────┬─────┘ └────────┬───────┘ └─────┬──────┘ └──────────────┘
    │            │                 │                │
    └────────────┴────────┬────────┴────────────────┘
                          │
                          ▼
                   ┌─────────────┐
                   │  EMPLEADO   │ 🎉
                   └──────┬──────┘
                          │
                  ┌───────┴───────┐
                  │ Seguimientos  │
                  │  Programados  │
                  └───────┬───────┘
                          │
                  ┌───────┴─────────┐
                  │                 │
                  ▼                 ▼
           ┌─────────────┐   ┌─────────────┐
           │  Llamada 1  │   │  Llamada 2  │
           │  (Semanal)  │   │  (3 meses)  │
           └─────────────┘   └─────────────┘
                  │
                  ▼
           ┌─────────────┐
           │  Llamada 3  │
           │  (6 meses)  │
           └─────────────┘
                  │
                  ▼
           ┌─────────────┐
           │   REPORTE   │
           │   MENSUAL   │
           └─────────────┘
```

---

## 🎯 Etapas del Flujo

### 1️⃣ **Entrada: Graduados**

**Origen:** Formulario de KoboToolbox

**Datos recopilados:**
- ID único
- Nombre completo
- Teléfono
- Email
- Fecha de entrevista
- Formación recibida
- Habilidades
- Experiencia previa

**Proceso:**
1. Los datos se importan automáticamente desde KoboToolbox cada hora (opcional)
2. O manualmente desde el menú: 📊 Seguimiento Graduados → 🔄 Importar desde KoboToolbox
3. Los graduados se almacenan en la hoja "Graduados" sin clasificar

---

### 2️⃣ **Clasificación Inicial: Entrevista/Seguimiento General**

**Descripción:** Punto de entrada para todos los graduados. Evaluación inicial y seguimiento general.

**Datos requeridos:**
- Tipo de entrevista (Inicial, Seguimiento, Evaluación)
- Resultado de la entrevista
- Siguiente paso recomendado

**Acción:**
- Los graduados pasan por esta etapa para determinar su ruta específica
- Se documenta el resultado de la entrevista
- Se decide la siguiente clasificación apropiada

**Hoja:** `Entrevista/Seguimiento General`

---

### 3️⃣ **Rutas de Seguimiento**

Desde la entrevista inicial, los graduados pueden seguir diferentes caminos:

#### 🤝 **Aliados**

**Descripción:** Graduados conectados con organizaciones o empresas aliadas.

**Datos registrados:**
- Aliado asignado
- Contacto del aliado (email/teléfono)
- Fecha de derivación
- Notas

**Objetivo:** Facilitar la inserción laboral a través de aliados estratégicos.

**Hoja:** `Aliados`

---

#### 💻 **Plataforma**

**Descripción:** Graduados registrados en plataformas digitales de búsqueda de empleo o capacitación.

**Datos registrados:**
- Plataforma asignada (LinkedIn, Indeed, Computrabajo, etc.)
- Usuario o email en la plataforma
- Fecha de registro
- Notas

**Objetivo:** Aprovechar plataformas digitales para búsqueda activa de empleo.

**Hoja:** `Plataforma`

---

#### 🌐 **Conexiones Laborales**

**Descripción:** Graduados con contactos directos en empresas o áreas específicas.

**Datos registrados:**
- Contacto (nombre de la persona)
- Empresa/Área
- Tipo de conexión (Referencia, Networking, Contacto directo)
- Notas

**Objetivo:** Utilizar networking y conexiones para acceder a oportunidades laborales.

**Hoja:** `Conexiones Laborales`

---

#### 👤 **Por su Cuenta**

**Descripción:** Graduados que realizan búsqueda laboral de forma independiente.

**Datos registrados:**
- Actividad que está realizando
- Progreso o avances
- Última actualización
- Notas

**Objetivo:** Monitorear y apoyar la búsqueda independiente.

**Hoja:** `Por su Cuenta`

---

#### 🔍 **Busca Trabajo (Fito)**

**Descripción:** Graduados derivados a Fito para seguimiento especializado en búsqueda de empleo.

**Datos registrados:**
- Motivo de derivación a Fito
- Fecha de derivación
- Notas

**Objetivo:** Recibir apoyo especializado de Fito en la búsqueda activa de empleo.

**Hoja:** `Busca Trabajo (Fito)`

**Nota:** Esta hoja funciona como punto de derivación a otro responsable (Fito).

---

### 4️⃣ **Meta Alcanzada: EMPLEADO** 🎉

**Descripción:** Graduado que consiguió empleo formal.

**Datos registrados:**
- Empresa
- Puesto/Cargo
- Fecha de contratación
- Salario (opcional)
- Notas

**Proceso automático:**
1. Al clasificar como "Empleado", el sistema automáticamente programa 3 seguimientos:
   - **Llamada 1**: A la semana (verificación inicial)
   - **Llamada 2**: A los 3 meses (evaluación de adaptación)
   - **Llamada 3**: A los 6 meses (permanencia y feedback)

2. Las llamadas se registran en la hoja "Seguimientos"

3. Se envían notificaciones por email cuando hay llamadas pendientes

**Hoja:** `Empleados`

---

### 5️⃣ **Seguimientos Programados**

**Descripción:** Sistema automático de llamadas de seguimiento para graduados empleados.

**Tipos de seguimiento:**

#### 📞 **Llamada 1 - Semanal (7 días después)**
**Objetivo:** Verificar que el graduado comenzó a trabajar correctamente.

**Preguntas clave:**
- ¿Comenzó a trabajar?
- ¿Cómo fue el primer día/semana?
- ¿Tiene alguna dificultad inicial?
- ¿Necesita algún tipo de apoyo?

---

#### 📞 **Llamada 2 - 3 Meses (90 días después)**
**Objetivo:** Evaluar la adaptación al puesto de trabajo.

**Preguntas clave:**
- ¿Sigue trabajando en la empresa?
- ¿Cómo ha sido la adaptación?
- ¿Qué ha aprendido?
- ¿Está satisfecho con el trabajo?
- ¿Ha tenido algún problema?

---

#### 📞 **Llamada 3 - 6 Meses (180 días después)**
**Objetivo:** Verificar permanencia y obtener feedback.

**Preguntas clave:**
- ¿Continúa en el mismo puesto?
- ¿Ha tenido algún ascenso o cambio?
- ¿Recomendaría la formación?
- ¿Qué mejoraría del programa?
- ¿Considera que la formación le ayudó?

---

### 6️⃣ **Reportes Mensuales**

**Descripción:** Consolidación de estadísticas y métricas del programa.

**Métricas clave:**
- Total de graduados
- Tasa de empleabilidad: (Empleados / Total Graduados) × 100
- Distribución por clasificación
- Seguimientos realizados vs pendientes
- Tasa de retención a 6 meses
- Tiempo promedio hasta el empleo

**Hoja:** `Reportes Mensuales`

---

## 🔄 Flujo Operativo Diario

### Para el Equipo de Seguimiento:

**Cada Día:**
1. **Revisar notificaciones** por email de seguimientos pendientes
2. **Realizar llamadas** programadas
3. **Marcar como realizadas** en la hoja "Seguimientos"
4. **Actualizar estados** si hay cambios

**Cada Semana:**
1. **Importar nuevos graduados** (si no está automático)
2. **Clasificar graduados** sin clasificar
3. **Revisar progreso** de cada categoría

**Cada Mes:**
1. **Generar reportes** de estadísticas
2. **Revisar tasas de empleabilidad**
3. **Analizar efectividad** de cada ruta

---

## 📋 Hojas del Sistema

| Hoja | Descripción |
|------|-------------|
| **Graduados** | Registro central de todos los graduados |
| **Entrevista/Seguimiento General** | Evaluación inicial y seguimiento general |
| **Aliados** | Graduados conectados con aliados |
| **Plataforma** | Graduados en plataformas digitales |
| **Conexiones Laborales** | Graduados con contactos en empresas |
| **Por su Cuenta** | Búsqueda independiente |
| **Busca Trabajo (Fito)** | Derivados a Fito |
| **Empleados** | Graduados con empleo conseguido |
| **Seguimientos** | Llamadas programadas (1 semana, 3 meses, 6 meses) |
| **Reportes Mensuales** | Estadísticas y métricas |
| **Configuración** | Configuración del sistema |

---

## 💡 Mejores Prácticas

### Clasificación
- Clasificar a los graduados lo más pronto posible
- Documentar siempre el resultado de las entrevistas
- Mantener notas actualizadas

### Seguimientos
- No dejar pasar las fechas programadas
- Documentar cada llamada con detalle
- Si no se contacta, reprogramar

### Reportes
- Actualizar estadísticas mensualmente
- Compartir resultados con el equipo
- Usar datos para mejorar el programa

### Comunicación
- Mantener información de contacto actualizada
- Registrar todos los intentos de comunicación
- Documentar cambios de situación laboral

---

## ❓ Preguntas Frecuentes

**¿Qué pasa si un graduado cambia de clasificación?**
Puedes reclasificarlo usando el formulario. La información se actualizará en las hojas correspondientes.

**¿Cómo marco un seguimiento como realizado?**
Ve a la hoja "Seguimientos", busca la fila correspondiente y actualiza los campos: Fecha Realizada, Estado (Realizado), Resultado y Notas.

**¿Qué pasa si un empleado renuncia o es despedido?**
Actualiza su estado en la hoja "Empleados" y documenta la situación en las notas. Puedes agregar seguimientos adicionales si es necesario.

**¿Puedo agregar más tipos de clasificación?**
Sí, pero requiere modificar el código. Contacta al administrador del sistema.

**¿Los seguimientos se reprograman automáticamente?**
No, si no contactas al graduado, debes reprogramar manualmente. Los seguimientos iniciales se programan automáticamente solo cuando clasificas como "Empleado".

---

**📘 Documentación Relacionada:**
- [Guía de Instalación](./INSTALACION_SIMPLIFICADA.md)
- [Manual de Uso](./MANUAL_USO.md)
- [README del Proyecto](../README.md)
