# 📋 PLAN ARQUITECTURA: BOT DE SEGUIMIENTO CON N8N
## Equipito Empleabilidad

> **Estado:** Plan estratégico (antes de implementar código)  
> **Fecha:** Abril 2026  
> **Objetivo:** Unificar seguimientos post-empleo y búsqueda activa en un solo bot (n8n + WhatsApp)

---

## 🎯 OBJETIVO GENERAL

Crear un sistema automático que:
1. **Jale datos** de Google Sheets (Seguimiento Bot)
2. **Dispare mensajes WhatsApp** automáticamente en fechas programadas
3. **Registre respuestas** en la misma hoja
4. **Maneje dos flujos:**
   - **Post-empleo** (de Conexiones Laborales): seguimiento laboral
   - **Búsqueda activa** (de Activamente busca trabajo): apoyo en búsqueda

---

## 📊 ARQUITECTURA DE DATOS

### Hoja Unificada: "Seguimiento Bot"

**Cambio clave:** Agregar columna `Tipo Seguimiento`

```
Columnas actuales (18):
  1. Creamos ID
  2. Nombre
  3. Teléfono
  4. Empresa
  5. Cargo
  6. Fecha Empleo
  7. Fecha Seg1  (30 días)
  8. Fecha Seg2  (60 días)
  9. Fecha Recordatorio
 10. Estado
 11. Etapa Actual
 12. Resp S1 P1  (Respuesta Seg1 Pregunta1)
 13. Resp S1 P2
 14. Resp S1 P3
 15. Resp S2 P1  (Respuesta Seg2 Pregunta1)
 16. Resp S2 P2
 17. Resp S2 P3
 18. Email Enviado

AGREGAR:
 19. Tipo Seguimiento    (dropdown: "Post-empleo" | "Búsqueda activa")
 20. Origen              (texto: "Conexiones Laborales" | "Activamente busca trabajo")
```

### ¿Por qué reusar la hoja?

| Opción | Pro | Contra |
|--------|-----|--------|
| **Reusar Seguimiento Bot** | Bot simple, una sola hoja | Algunos campos vacíos según tipo |
| **Crear hoja nueva** | Estructura exacta | Bot duplicado, más mantenimiento |

✅ **Decisión:** Reusar Seguimiento Bot + agregar `Tipo Seguimiento`

---

## 🤖 FLUJO DE N8N

### 1. Trigger: Google Sheets (lectura cada hora)

```
n8n:
  - Trigger: Cron job cada hora
  - Acción: Lee hoja "Seguimiento Bot"
  - Filtro: `Estado = "" OR Estado = "Pendiente"`
  - Busca: Fechas que coinciden con hoy
```

### 2. Lógica según tipo

```javascript
if (tipoSeguimiento === "Post-empleo") {
  // Mensajes sobre empleo
  // Preguntas: cómo va el trabajo, adaptación, salario
  // Acciones: 3 mensajes en 60 días
}

if (tipoSeguimiento === "Búsqueda activa") {
  // Mensajes sobre búsqueda
  // Preguntas: entrevistas, empresas contactadas, ofertas
  // Acciones: 3 mensajes en 60 días
  // + Llamada human en día 21
}
```

### 3. Tres acciones por tipo (7, 14, 21 días)

| Día | Post-empleo | Búsqueda activa | Ejecutor |
|-----|-------------|-----------------|----------|
| 7 | WhatsApp Msg 1 | WhatsApp Msg 1 | 🤖 Bot n8n |
| 14 | WhatsApp Msg 2 | WhatsApp Msg 2 | 🤖 Bot n8n |
| 21 | WhatsApp Msg 3 | **Llamada** | 👤 Colaborador |

### 4. Respuestas → Google Sheets

n8n escucha webhook de WhatsApp y escribe respuestas en:
- `Resp S1 P1, Resp S1 P2, Resp S1 P3` (día 7 y 14)
- `Resp S2 P1, Resp S2 P2, Resp S2 P3` (opcional, para futuro)

---

## 📱 PREGUNTAS WHATSAPP

### POST-EMPLEO (Día 7 y 14)

**Día 7 (Primer seguimiento):**
```
Hola {Nombre}, esperamos vayas bien en {Empresa}!

Nos encantaría saber:
1️⃣ ¿Cómo va tu adaptación a {Cargo}? (Excelente/Bien/Normal/Difícil)
2️⃣ ¿Te sientes apoyado por tu jefe/equipo? (Sí/Parcialmente/No)
3️⃣ ¿Algún desafío específico que enfrentas? (Escribe tu respuesta)
```

**Día 14 (Segundo seguimiento):**
```
¡Hola de nuevo! Queremos seguir tu progreso.

Cuéntanos:
1️⃣ ¿Qué es lo mejor de tu nuevo trabajo? (respuesta libre)
2️⃣ ¿Te sientes más cómodo que hace una semana? (Mucho/Poco/Igual)
3️⃣ ¿Hay algo en lo que podamos ayudarte? (respuesta libre)
```

### BÚSQUEDA ACTIVA (Día 7 y 14)

**Día 7:**
```
Hola {Nombre}, queremos saber cómo va tu búsqueda.

Actualízanos:
1️⃣ ¿Cuántas empresas has contactado? (número)
2️⃣ ¿Has tenido entrevistas? (cuántas)
3️⃣ ¿Necesitas ayuda con algo? (respuesta libre)
```

**Día 14:**
```
¡Seguimos contigo en tu búsqueda!

Dinos:
1️⃣ ¿Has recibido ofertas? (Sí/Pendiente/No)
2️⃣ ¿Qué tipo de trabajo buscas ahora? (respuesta libre)
3️⃣ ¿Cómo te sientes con el progreso? (Optimista/Neutral/Desanimado)
```

---

## 🔧 ESTRUCTURA JSON PARA N8N

### Flujo general

```json
{
  "name": "Seguimiento Automático Equipito",
  "active": true,
  "triggers": [
    {
      "type": "cron",
      "schedule": "0 9 * * *",
      "description": "Ejecutar cada día a las 9:00 AM"
    }
  ],
  "nodes": [
    {
      "id": "trigger_cron",
      "type": "n8n-nodes-base.cron",
      "name": "Cada día a las 9 AM",
      "config": {
        "triggerTimes": {
          "item": [
            {
              "mode": "everyDay",
              "hour": 9,
              "minute": 0
            }
          ]
        }
      }
    },
    {
      "id": "read_sheets",
      "type": "n8n-nodes-base.googleSheets",
      "name": "Leer Seguimiento Bot",
      "config": {
        "resource": "spreadsheet",
        "operation": "read",
        "spreadsheetId": "{{SHEET_ID}}",
        "sheetName": "Seguimiento Bot",
        "readFrom": "A1:T100"
      }
    },
    {
      "id": "filter_hoy",
      "type": "n8n-nodes-base.itemLists",
      "name": "Filtrar por fecha de hoy",
      "config": {
        "operation": "filter",
        "condition": {
          "field": "Fecha Seg1",
          "operator": "equals",
          "value": "{{$now.toDate()}}"
        }
      }
    },
    {
      "id": "split_type",
      "type": "n8n-nodes-base.switchNode",
      "name": "Decidir: ¿Post-empleo o Búsqueda activa?",
      "config": {
        "cases": [
          {
            "condition": "Tipo Seguimiento == 'Post-empleo'",
            "output": 0
          },
          {
            "condition": "Tipo Seguimiento == 'Búsqueda activa'",
            "output": 1
          }
        ]
      }
    },
    {
      "id": "whatsapp_send",
      "type": "n8n-nodes-base.httpRequest",
      "name": "Enviar WhatsApp",
      "config": {
        "method": "POST",
        "url": "https://api.twilio.com/2010-04-01/Accounts/{{TWILIO_ACCOUNT_SID}}/Messages",
        "authentication": "basicAuth",
        "auth": {
          "user": "{{TWILIO_ACCOUNT_SID}}",
          "password": "{{TWILIO_AUTH_TOKEN}}"
        },
        "body": {
          "From": "whatsapp:+{{TWILIO_WHATSAPP_NUMBER}}",
          "To": "whatsapp:+{{$item.json.Telefono}}",
          "Body": "{{MENSAJE_SEGUN_TIPO}}"
        }
      }
    },
    {
      "id": "update_sheets",
      "type": "n8n-nodes-base.googleSheets",
      "name": "Actualizar Estado",
      "config": {
        "resource": "spreadsheet",
        "operation": "update",
        "spreadsheetId": "{{SHEET_ID}}",
        "sheetName": "Seguimiento Bot",
        "range": "J:J",
        "value": "Mensaje Enviado"
      }
    }
  ]
}
```

### Webhook para respuestas WhatsApp

```json
{
  "webhook": {
    "name": "Recibir respuestas WhatsApp",
    "url": "https://n8n.example.com/webhook/whatsapp-response",
    "method": "POST",
    "payload": {
      "from": "+{{NUMERO_PERSONA}}",
      "body": "{{RESPUESTA}}",
      "timestamp": "{{HORA}}"
    }
  },
  "action": "Buscar en Sheets por teléfono y escribir respuesta en columna Resp S1 P1/P2/P3"
}
```

---

## 💰 PRESUPUESTO Y HERRAMIENTAS

### Herramientas necesarias

| Tool | Función | Costo (USD/mes) | Alternativa |
|------|---------|-----------------|-------------|
| **n8n Cloud** | Bot automation | $20-50 | n8n Self-hosted (gratis) |
| **Twilio (WhatsApp)** | Enviar/recibir mensajes | ~$0.01 por msg | Messagebird, Vonage |
| **Google Sheets API** | Lectura/escritura | Gratis (limitado) | - |
| **Zapier** | *(opcional)* alternativa a n8n | $19-99 | - |

### Recomendación de setup

```
Opción A: Producción en la nube (recomendado para empezar)
├─ n8n Cloud ($25/mes)
├─ Twilio WhatsApp ($0.01 por mensaje)
└─ Google Sheets (gratis, preexistente)
Total: ~$25-50/mes + $0.50-2/mes por mensajes

Opción B: Self-hosted (si tienen servidor)
├─ n8n Self-hosted (gratis)
├─ Twilio WhatsApp ($0.01 por mensaje)
└─ Google Sheets (gratis)
Total: $0 + $0.50-2/mes por mensajes

Opción C: Bajo costo (prueba)
├─ Make.com ($0 plan gratuito, limitado)
├─ Twilio ($0.01 por mensaje)
└─ Google Sheets (gratis)
Total: ~$1-2/mes por mensajes
```

**Mi recomendación:** Opción A (n8n Cloud) por ser estable y escalable.

---

## 📌 CAMBIOS NECESARIOS EN GOOGLE SHEETS

### 1. Agregar columnas a "Seguimiento Bot"

```
Nueva columna 19: Tipo Seguimiento
  - Tipo: Dropdown
  - Opciones: ["Post-empleo", "Búsqueda activa"]
  - Valor por defecto: Se establece al crear la fila

Nueva columna 20: Origen
  - Tipo: Texto (solo lectura)
  - Se llena automáticamente según dónde viene el registro
```

### 2. Cambiar columna 10 "Estado"

De: texto libre  
A: Dropdown con opciones:
- "Pendiente" (por enviar mensaje)
- "Mensaje Enviado"
- "Respuesta Recibida"
- "Llamada Programada"
- "Completado"

### 3. NO eliminar hoja "Seguimientos"

Aunque elimines el código de registro automático, mantener la hoja por si el colaborador necesita registrar llamadas manuales.

---

## 🔄 FLUJO DE INTEGRACIÓN

### Cuando llega a "Conexiones Laborales"

```
Usuario: Guarda empleo en Conexiones Laborales
    ↓
Apps Script trigger: onEditInstalable (ya existe)
    ↓
crearFilaSeguimientoBot() (modificar)
    ├─ Llena: Tipo Seguimiento = "Post-empleo"
    ├─ Llena: Origen = "Conexiones Laborales"
    └─ Calcula: Fecha Seg1 (hoy + 7), Fecha Seg2 (hoy + 14)
    ↓
Seguimiento Bot (nueva fila)
    ↓
n8n (diario a las 9 AM)
    ├─ Lee Seguimiento Bot
    ├─ Busca filas con Fecha Seg1 = hoy
    ├─ Envía WhatsApp (mensaje post-empleo)
    └─ Actualiza Estado = "Mensaje Enviado"
```

### Cuando llega a "Activamente busca trabajo"

```
Usuario: Cambia Etapa a "Activamente busca trabajo" en Graduados
    ↓
onEdit trigger (ya existe)
    ↓
Nueva función: crearSeguimientoBusquedaActiva()
    ├─ Busca a la persona en Activamente busca trabajo
    ├─ Crea fila en Seguimiento Bot
    ├─ Llena: Tipo Seguimiento = "Búsqueda activa"
    ├─ Llena: Origen = "Activamente busca trabajo"
    └─ Calcula: Fecha Seg1 (hoy + 7), Fecha Seg2 (hoy + 14)
    ↓
Seguimiento Bot (nueva fila)
    ↓
n8n (diario)
    ├─ Lee Seguimiento Bot
    ├─ Busca filas con Fecha Seg1 = hoy Y Tipo = "Búsqueda activa"
    ├─ Envía WhatsApp (mensaje búsqueda)
    └─ Actualiza Estado = "Mensaje Enviado"
    ↓
Día 21: Marca para llamada humana
    ├─ n8n crea recordatorio
    └─ Colaborador ve "Llamada Pendiente" en un botón del menú
```

---

## ⚙️ CAMBIOS EN CÓDIGO APPS SCRIPT

### 1. Modificar `crearFilaSeguimientoBot()`

```javascript
function crearFilaSeguimientoBot(datos, fechaEmpleo, tipoSeguimiento = "Post-empleo") {
  const hoja = obtenerHoja('Seguimiento Bot');
  const hoy = new Date();
  const seg1 = new Date(hoy.getTime() + 7 * 24 * 60 * 60 * 1000);  // +7 días
  const seg2 = new Date(hoy.getTime() + 14 * 24 * 60 * 60 * 1000); // +14 días

  hoja.appendRow([
    datos.creamosId,
    datos.nombre,
    datos.telefono,
    datos.empresa,
    datos.cargo,
    fechaEmpleo,
    seg1,
    seg2,
    new Date(hoy.getTime() + 21 * 24 * 60 * 60 * 1000),  // Recordatorio +21
    "Pendiente",        // Estado
    "",                 // Etapa Actual
    "",                 // Resp S1 P1
    "",                 // Resp S1 P2
    "",                 // Resp S1 P3
    "",                 // Resp S2 P1
    "",                 // Resp S2 P2
    "",                 // Resp S2 P3
    "No",               // Email Enviado
    tipoSeguimiento,    // ← NUEVA COLUMNA
    "Conexiones Laborales"  // ← NUEVA COLUMNA (Origen)
  ]);
}
```

### 2. Nueva función: `crearSeguimientoBusquedaActiva()`

```javascript
function crearSeguimientoBusquedaActiva(creamosId, nombre, telefono) {
  const hoja = obtenerHoja('Seguimiento Bot');
  const hoy = new Date();
  
  // No tiene "Empresa" ni "Cargo", dejar vacíos
  hoja.appendRow([
    creamosId,
    nombre,
    telefono,
    "",                 // Empresa (vacío)
    "",                 // Cargo (vacío)
    hoy,                // Fecha inicio búsqueda
    new Date(hoy.getTime() + 7 * 24 * 60 * 60 * 1000),
    new Date(hoy.getTime() + 14 * 24 * 60 * 60 * 1000),
    new Date(hoy.getTime() + 21 * 24 * 60 * 60 * 1000),
    "Pendiente",
    "Búsqueda Activa",
    "", "", "", "", "", "",
    "No",
    "Búsqueda activa",    // ← Tipo Seguimiento
    "Activamente busca trabajo"  // ← Origen
  ]);
}
```

### 3. Modificar `onEdit()` para llamar nueva función

En la sección donde detectas cambio de Etapa:

```javascript
if (columnaEtapa === 12) {  // Columna "Etapa"
  const nuevaEtapa = valores[i][14];  // [14] = Etapa (col 15)
  
  if (nuevaEtapa === "Activamente busca trabajo") {
    crearSeguimientoBusquedaActiva(
      valores[i][2],    // Creamos ID
      valores[i][3],    // Nombre
      valores[i][7]     // Teléfono
    );
  }
  // ...resto del código
}
```

---

## 📋 PLAN DE IMPLEMENTACIÓN

### Fase 1: Preparación (Semana 1)
- [ ] Definir preguntas exactas para WhatsApp (post-empleo + búsqueda)
- [ ] Configurar cuenta Twilio (obtener número WhatsApp)
- [ ] Registrarse en n8n Cloud
- [ ] Documentar estructura JSON final de n8n
- [ ] **Código Google Sheets:** Agregar columnas 19-20 a estructura
- [ ] **Código Google Sheets:** Crear funciones nuevas

### Fase 2: Configuración n8n (Semana 2)
- [ ] Crear workflow base en n8n
- [ ] Configurar trigger cron (diario 9 AM)
- [ ] Integrar Google Sheets (lectura)
- [ ] Integrar Twilio (envío WhatsApp)
- [ ] Testar con 2-3 registros de prueba
- [ ] Ajustar mensajes según feedback

### Fase 3: Webhook respuestas (Semana 3)
- [ ] Crear webhook en n8n para recibir respuestas
- [ ] Implementar lógica: buscar persona por teléfono
- [ ] Escribir respuestas en columnas correctas
- [ ] Testar envío + respuesta circular

### Fase 4: Integración Google Sheets (Semana 4)
- [ ] Publicar cambios de código (crearSeguimientoBusquedaActiva)
- [ ] Testar: guardar en Conexiones Laborales → aparece en Seguimiento Bot
- [ ] Testar: cambiar a "Activamente busca trabajo" → aparece en Seguimiento Bot
- [ ] Testar: n8n lee y envía mensajes

### Fase 5: Validación (Semana 5)
- [ ] Revisión completa con equipo
- [ ] Feedback de usuarios
- [ ] Ajustes finales
- [ ] Documentación para colaboradores

---

## ❓ PREGUNTAS ANTES DE PROCEDER

1. **¿Tienen número Twilio?** (O necesitamos comprar)
2. **¿Quién aprueba presupuesto de herramientas?** (n8n Cloud ~$25/mes)
3. **¿Las preguntas WhatsApp están bien?** (Necesito tu visto bueno)
4. **¿Llamada del día 21 (búsqueda activa) cómo se registra?** 
   - ¿En Seguimiento Bot automáticamente?
   - ¿El colaborador marca cuando la hace?
5. **¿Qué pasa si alguien no responde?** (¿Reenviar? ¿Después de cuántos días?)
6. **¿Horarios de envío de mensajes?** (9 AM es bueno o preferis otro)

---

## 📞 PRÓXIMOS PASOS

1. **Aprueba este plan** (o sugiere cambios)
2. **Responde las 6 preguntas arriba**
3. **Aprueba presupuesto** (herramientas)
4. **Comenzamos Fase 1** la próxima semana

---

**¿Qué piensas de esta arquitectura? ¿Hay cambios que necesitas?**
