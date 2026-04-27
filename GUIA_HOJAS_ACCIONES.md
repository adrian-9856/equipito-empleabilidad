# 📊 GUÍA DE HOJAS Y ACCIONES — EQUIPITO EMPLEABILIDAD

> **Documento de referencia rápida:** Identifica cada hoja, sus columnas, acciones automáticas y qué menú las controla.

---

## 🎯 NAVEGACIÓN RÁPIDA

¿No sabes qué hoja necesitas cambiar? Usa este mapa:

| 🔍 Si necesitas... | 👉 Mira esta hoja |
|---|---|
| Registrar nuevos graduados | **Graduados** |
| Clasificar perfiles (7 dimensiones) | **Clasificación de Perfiles** |
| Registrar empresas aliadas | **Aliados** |
| Registrar búsquedas en plataforma | **Plataforma** |
| Registrar derivaciones a empresas | **Derivaciones** |
| Registrar búsqueda autónoma | **Activamente busca trabajo** |
| Registrar personas sin empleo | **Paso a paso** |
| Registrar empleo conseguido | **Conexiones Laborales** |
| Ver seguimientos de personas empleadas | **Seguimientos Bot** + **Seguimientos** |
| Encuestas de satisfacción laboral | **Satisfacción Empleo** |
| Sesiones de acompañamiento profesional | **Sesiones Acompañamiento** |
| Ver resumen de estadísticas | **Reporte** |

---

## 📋 DETALLE DE CADA HOJA

### 1. **GRADUADOS** (Azul oscuro 🔵)
- **Color:** #1a73e8
- **¿Qué es?** Registro principal de todos los graduados en el programa
- **Columnas (15):**
  - No. | Fecha de envío | Creamos ID | Nombre completo | Género | Edad |
  - Nivel educativo | Número de teléfono | Formación | Cohorte | Fecha de entrevista |
  - Empleado (Sí/No) | Próxima llamada | Notas | **Etapa** (desplegable)

- **Etapas disponibles:** (automáticamente copia a otras hojas)
  - `Aliados`
  - `Plataforma`
  - `Derivaciones`
  - `Paso a paso`
  - `Activamente busca trabajo`
  - `Conexiones Laborales`

- **Acciones automáticas:**
  - ✅ Al cambiar la columna **Etapa**, el graduado se copia automáticamente a su hoja de clasificación
  - 📥 Se llena automáticamente al importar desde KoboToolbox

- **Menú de control:**
  - 📥 Importar todos los datos → importa nuevos graduados
  - 📥 Importar Graduados → jala de hoja externa cada hora (auto)

---

### 2. **ALIADOS** (Índigo 💜)
- **Color:** #3f51b5
- **¿Qué es?** Graduados registrados con empresas aliadas
- **Columnas comunes (6):** Fecha ingreso | Nombre completo | ID | Género | Edad | Nivel educativo

- **Columnas específicas (7):**
  - Compartió el CV (Sí/No)
  - Área (dropdown: varias opciones)
  - Entrevista (texto)
  - Día de prueba (texto)
  - Confirmación de recepción (Sí/No)
  - Notas (texto)
  - Activo (Sí/No)

- **¿Cómo llega aquí?** Cuando en la hoja **Graduados** cambias la **Etapa** a "Aliados", la fila se copia aquí automáticamente

- **Mantenimiento:** Registra el progreso del graduado con aliados (CV, entrevista, prueba, etc.)

---

### 3. **PLATAFORMA** (Cian 🩵)
- **Color:** #00bcd4
- **¿Qué es?** Graduados registrados en plataformas de empleo
- **Columnas comunes:** Fecha ingreso | Nombre completo | ID | Género | Edad | Nivel educativo

- **Columnas específicas (9):**
  - Cita ☐ (checkbox)
  - Creación de perfil (Sí/No)
  - Contacto ☐ (checkbox)
  - Trámites | Entrevista | Confirmación | Recepción (textos)
  - Nota | Activo (textos/Sí-No)

- **¿Cómo llega?** Cambiar **Etapa** a "Plataforma" en **Graduados**

- **Mantenimiento:** Seguimiento de crear perfil, contacto, entrevistas, etc.

---

### 4. **DERIVACIONES** (Naranja 🟠)
- **Color:** #ff9800
- **¿Qué es?** Graduados derivados a múltiples empresas (hasta 3 intentos)
- **Columnas comunes:** Fecha ingreso | Nombre completo | ID | Género | Edad | Nivel educativo

- **Columnas específicas (8):**
  - Envío 1 | Llamada 1 | Envío 2 | Llamada 2 | Envío 3 | Llamada 3 | Notas | Activo

- **¿Cómo llega?** Cambiar **Etapa** a "Derivaciones" en **Graduados**

- **Mantenimiento:** Registra 3 intentos de derivación a empresas

---

### 5. **PASO A PASO** (Gris 🩶)
- **Color:** #9e9e9e
- **¿Qué es?** Personas que no buscan trabajo pero requieren proceso especial
- **Columnas comunes:** Fecha ingreso | Nombre completo | ID | Género | Edad | Nivel educativo

- **Columnas específicas (6):**
  - DPI | Número de teléfono | Formación | Cohorte | Nota | Activo

- **¿Cómo llega?** Cambiar **Etapa** a "Paso a paso" en **Graduados**

- **Mantenimiento:** Registra datos de personas en proceso especial

---

### 6. **ACTIVAMENTE BUSCA TRABAJO** (Verde 🟢)
- **Color:** #0f9d58
- **¿Qué es?** Graduados que buscan empleo por su cuenta O alguien más les ayuda
- **Columnas comunes:** Fecha ingreso | Nombre completo | ID | Género | Edad | Nivel educativo

- **Columnas específicas (7):**
  - Tipo de búsqueda (dropdown: "Activamente busca trabajo" / "Por su cuenta")
  - Mensaje | Llamada (textos)
  - Nota (texto)
  - Entrevista | Trámites (textos)
  - Activo (Sí/No)

- **¿Cómo llega?** Cambiar **Etapa** a "Activamente busca trabajo" en **Graduados**

- **Mantenimiento:** Registra búsqueda autónoma o con acompañamiento

---

### 7. **CONEXIONES LABORALES** (Naranja oscuro 🟠)
- **Color:** #e65100
- **¿Qué es?** **MÁS IMPORTANTE:** Registro del empleo conseguido (empresa, cargo, salario, fechas)
- **Columnas (18):**
  - Creamos ID | Nombre completo | Teléfono | Género | Edad | Nivel educativo |
  - Tipo | Programa | Proyecto | Especialidad | Empresa | Cargo que desempeña |
  - Tipo de duración de contrato | Tipo de contrato |
  - **Fecha de inicio** | **Fecha de final** | **Duración (meses)** | **Salario mensual**

- **¿Cómo llega?** Desde **Graduados**, seleccionar fila + botón en menú → "Enviar a Conexiones Laborales"

- **Acción automática:** ✅ Al guardar datos de empleo aquí:
  - Se crea automáticamente un registro en **Seguimiento Bot** con 2 fechas de seguimiento programadas
  - Se programa un recordatorio para meses después

- **MÁS IMPORTANTE:** Esta hoja **dispara automáticamente** la creación de datos en **Seguimiento Bot**

---

### 8. **SATISFACCIÓN EMPLEO** (IL-06) (Verde azulado 🟦)
- **Color:** #00897b
- **¿Qué es?** Encuestas de satisfacción laboral (AAPI) de personas empleadas
- **Columnas (15):**
  - Creamos ID | Fecha envío | Nombre | Apellido | Primera vez (Sí/No) |
  - Formación recibida (texto) |
  - P1, P2, P3, P4, P5, P6 (respuestas de encuesta) |
  - Comentarios | Puntaje AAPI | Satisfacción

- **¿Cómo llena?** Automáticamente desde KoboToolbox (formulario IL-06)

- **Menú de control:**
  - 😊 Satisfacción Empleo → importar datos nuevos manualmente

---

### 9. **SESIONES ACOMPAÑAMIENTO** (Púrpura 🟣)
- **Color:** #5e35b1
- **¿Qué es?** Registro de sesiones de acompañamiento profesional (formación, coaching)
- **Columnas (15):**
  - Creamos ID | Fecha envío | Inicio sesión | Proyecto | Nombre | Apellidos |
  - Teléfono | Fecha nacimiento | Edad | Género |
  - Año ingreso Creamos | En qué año | Grado académico |
  - Tipo servicio | Comentario

- **¿Cómo llena?** Automáticamente desde KoboToolbox (formulario IL-08)

- **Menú de control:**
  - 🤝 Sesiones Acompañamiento → importar datos nuevos

---

### 10. **CLASIFICACIÓN DE PERFILES** (Azul índigo 💙)
- **Color:** #5c6bc0
- **¿Qué es?** Evaluación de barreras (7 dimensiones) para cada graduado
- **Dimensiones evaluadas:**
  - D1: Cuidado (responsabilidades familiares)
  - D2: Violencia (seguridad)
  - D3: Movilidad (transporte, distancia)
  - D4: Legal/Salud (documentos, salud)
  - D5: Motivación (deseo de trabajar)
  - D6: Experiencia (historial laboral)
  - D7: Autonomía (independencia)

- **Columnas (21):**
  - Para cada dimensión: Evaluación | Comentario
  - Puntaje Total | Barreras activas | Desmotivación |
  - Perfil Asignado | Notas de observación

- **¿Cómo llena?** Importar desde KoboToolbox formulario "Clasificación de Perfiles"

- **Resultado:** Asigna un **Perfil** según barreras:
  - Rojo (máximas barreras)
  - Amarillo (barreras medias)
  - Verde (mínimas barreras)

- **Menú de control:**
  - 📋 Clasificación de Perfiles → importar evaluaciones nuevas
  - 🔍 Diagnosticar → ver si hay errores

---

### 11. **SEGUIMIENTOS** (Púrpura oscuro 🟣)
- **Color:** #673ab7
- **¿Qué es?** Registro manual de seguimientos después de empleo
- **Columnas (14):**
  - No. | Creamos ID | Nombre completo | Teléfono | Género | Edad | Nivel educativo |
  - Tipo seguimiento | Fecha programada | Fecha realizada |
  - Estado | Resultado | Notas | Próximo paso

- **¿Cómo se usa?** Manual: el usuario crea registros aquí cuando hace seguimientos

- **Acción:** Botón en menú para "Ver Seguimientos Pendientes"

---

### 12. **SEGUIMIENTO BOT** (Verde azulado 🟦)
- **Color:** #00897b
- **¿Qué es?** **AUTOMÁTICO:** Registro de seguimientos programados vía n8n + WhatsApp
- **Columnas (18):**
  - Creamos ID | Nombre | Teléfono | Empresa | Cargo |
  - Fecha Empleo | Fecha Seg1 | Fecha Seg2 | Fecha Recordatorio |
  - Estado | Etapa Actual |
  - Resp S1 P1, Resp S1 P2, Resp S1 P3 (respuestas seguimiento 1)
  - Resp S2 P1, Resp S2 P2, Resp S2 P3 (respuestas seguimiento 2)
  - Email Enviado

- **¿Cómo se llena?** AUTOMÁTICAMENTE cuando se guarda en **Conexiones Laborales**:
  - Se copian datos: Creamos ID, Nombre, Teléfono, Empresa, Cargo, Fecha Empleo
  - Se calculan: Fecha Seg1 (30 días) y Fecha Seg2 (60 días)

- **Integración:** n8n usa estos datos para enviar mensajes WhatsApp automáticos

---

### 13. **REPORTE** (Rosa 🩷)
- **Color:** #e91e63
- **¿Qué es?** **SOLO LECTURA - SE GENERA AUTOMÁTICAMENTE**
- **Contenido:**
  - Tabla de estadísticas generales
  - Conteos por etapa
  - Conteos por perfil (de Clasificación de Perfiles)
  - Seguimientos pendientes

- **¿Cómo se actualiza?** Menú: 📊 Generar Reporte

- **Nota:** No edites esta hoja manualmente; se sobrescribe al actualizar

---

## 🔄 FLUJOS AUTOMÁTICOS CLAVE

### Flujo 1: DE GRADUADOS A CLASIFICACIÓN
```
Graduados (Etapa = columna 15)
    ↓ (cambias Etapa a "Aliados", "Plataforma", etc.)
    ↓ (trigger onEdit automático)
    ↓
Se copia fila → Aliados / Plataforma / Derivaciones / etc.
```

### Flujo 2: EMPLEO CONSEGUIDO
```
Graduados (seleccionas fila)
    ↓ (botón: "Enviar a Conexiones Laborales")
    ↓ (abre formulario)
    ↓ (guardas datos: empresa, cargo, salario, fechas)
    ↓
Conexiones Laborales (se guarda fila)
    ↓ (trigger onEditInstalable)
    ↓ (automático)
    ↓
Seguimiento Bot (se crea fila con fechas programadas)
```

### Flujo 3: IMPORTACIONES AUTOMÁTICAS DESDE KOBOTOOLBOX
```
KoboToolbox
    ↓ (cada hora, trigger horario)
    ↓
Se sincronizan:
- Graduados (nuevos formularios)
- Clasificación de Perfiles (IL-09)
- Satisfacción Empleo (IL-06)
- Sesiones Acompañamiento (IL-08)
```

### Flujo 4: PULL HORARIO DESDE HOJA EXTERNA
```
Hoja externa: "Graduados" (1_596FX6yr8tX93UyIks4emSeE2_vxLJMDyw9Zncsnzs)
    ↓ (cada hora, trigger automático)
    ↓ (dedup por Creamos ID o Nombre)
    ↓
Hoja local: "Graduados Importados"
```

---

## 📱 MENÚ PRINCIPAL: QUÉ HACE CADA OPCIÓN

### 📊 SEGUIMIENTO GRADUADOS (Menú Principal)

| Opción | Función | Qué hace |
|---|---|---|
| 📥 Importar todos los datos | `importarTodosLosDatos()` | Importa Graduados + Clasificación + Satisfacción + Sesiones |
| 📊 Generar Reporte | `generarReporte()` | Actualiza hoja **Reporte** con estadísticas |
| 📝 Clasificar Graduados | `mostrarFormularioClasificacion()` | Abre formulario para clasificar un graduado sin clasificar |
| 📞 Ver Seguimientos Pendientes | `mostrarSeguimientosPendientes()` | Muestra lista de seguimientos por hacer |
| **📂 Importar por separado** | Submenu | ↓ |
| 🔄 Graduados | `importarDatosKobo()` | Importa solo nuevos graduados desde KoboToolbox |
| 📋 Clasificación de Perfiles | `importarClasificacionPerfiles()` | Importa solo clasificaciones nuevas |
| 😊 Satisfacción Empleo | `importarSatisfaccionEmpleo()` | Importa solo encuestas nuevas |
| 🤝 Sesiones Acompañamiento | `importarSesionesAcompanamiento()` | Importa solo sesiones nuevas |
| ⏱ Activar auto-import | `activarAutoImport()` | Pregunta cada cuántos minutos; crea trigger |
| ⏹ Desactivar auto-import | `desactivarAutoImport()` | Borra triggers de importación automática |
| **📥 Importar Graduados** | `importarGraduadosDesdeExterno()` | Jala desde hoja externa (NEW) |
| **⏰ Instalar Trigger Graduados** | `instalarTriggerGraduados()` | Instala trigger horario automático (NEW) |
| 🚀 Instalar Sistema | `instalarSistema()` | Crea las 13 hojas si no existen |
| 🔁 Reinstalar Sistema | `reinstalarSistema()` | ⚠️ Borra TODO y recrea desde cero |

---

## 🎓 DIMENSIONES DE CLASIFICACIÓN (Clasificación de Perfiles)

Cada persona se evalúa en 7 áreas de barrera:

| Dimensión | Evalúa | Niveles |
|---|---|---|
| **D1: Cuidado** | ¿Tiene responsabilidades de cuidado (hijos, ancianos)? | Bajo / Medio / Alto |
| **D2: Violencia** | ¿Ha sufrido violencia o está en riesgo? | Bajo / Medio / Alto |
| **D3: Movilidad** | ¿Tiene acceso a transporte/distancia viable? | Bajo / Medio / Alto |
| **D4: Legal/Salud** | ¿Tiene DPI, vacunas, permisos legales? | Bajo / Medio / Alto |
| **D5: Motivación** | ¿Quiere realmente trabajar? | Bajo / Medio / Alto |
| **D6: Experiencia** | ¿Tiene experiencia laboral previa? | Bajo / Medio / Alto |
| **D7: Autonomía** | ¿Puede decidir/actuar por sí misma? | Bajo / Medio / Alto |

**Resultado:** Se calcula **Puntaje Total** (suma) y se asigna un **Perfil** (Rojo/Amarillo/Verde)

---

## 📌 PREGUNTAS FRECUENTES

### P: ¿Dónde registro un nuevo graduado?
**R:** En la hoja **Graduados**, fila nueva. Completa: Nombre, Teléfono, Formación, Cohorte, etc.

### P: ¿Cómo cambio a un graduado de etapa (p.ej., de Aliados a Plataforma)?
**R:** En la hoja **Graduados**, cambia la columna **Etapa** (columna 15). Se copia automáticamente.

### P: ¿Dónde registro que alguien consiguió empleo?
**R:** En **Graduados**, selecciona su fila → Menú → Enviar a Conexiones Laborales → rellena: Empresa, Cargo, Salario, Fechas.

### P: ¿El sistema hace seguimientos automáticamente?
**R:** Sí, parcialmente. Después de guardar en **Conexiones Laborales**:
  - Se crea un registro en **Seguimiento Bot** con 2 fechas de seguimiento
  - n8n envía mensajes WhatsApp automáticos (si está configurado)

### P: ¿Quién rellena "Clasificación de Perfiles"?
**R:** Se importa automáticamente desde KoboToolbox (formulario IL-09). O manual si lo necesitas.

### P: ¿Por qué hay dos hojas de seguimiento (**Seguimientos** y **Seguimiento Bot**)?
**R:**
  - **Seguimientos:** Registro manual de seguimientos generales
  - **Seguimiento Bot:** Automático, vinculado a empleo conseguido, integrado con n8n/WhatsApp

### P: ¿Cuándo se ejecutan las importaciones automáticas?
**R:** Cada hora, si está activado el trigger "sincronización horaria". Se puede desactivar en el menú.

### P: ¿Puedo editar la hoja "Reporte"?
**R:** No, es automática. Se genera cada vez que haces clic en "Generar Reporte".

---

## 🛠️ COLUMNAS COMUNES (todas menos Conexiones, IL-06, IL-08)

Aparecen en: **Graduados, Aliados, Plataforma, Derivaciones, Paso a paso, Activamente busca trabajo, Seguimientos**

1. **Fecha de ingreso** (fecha)
2. **Nombre completo** (texto)
3. **ID** (texto, ej: "ALI-001")
4. **Género** (dropdown: Femenino / Masculino / Otro)
5. **Edad** (número)
6. **Nivel educativo** (texto)

---

## 📞 SOPORTE

Si tienes dudas sobre qué cambiar o dónde, usa este documento:
1. **¿Qué necesitas registrar?** → Busca en "NAVEGACIÓN RÁPIDA"
2. **¿Qué acción hace cada hoja?** → Lee su sección
3. **¿Está automatizado?** → Mira "FLUJOS AUTOMÁTICOS CLAVE"
4. **¿Qué botón del menú?** → Consulta "MENÚ PRINCIPAL"

---

**Última actualización:** Abril 2026  
**Versión:** 2.0 (con pull horario desde externa)
