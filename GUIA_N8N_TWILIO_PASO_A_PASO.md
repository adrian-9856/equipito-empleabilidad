# 🚀 GUÍA PASO A PASO: BOT WHATSAPP CON N8N + TWILIO

> Con $15 de presupuesto Twilio puedes enviar ~1500 mensajes de prueba.
> Esta guía es desde cero, sin necesidad de experiencia previa.

---

## ✅ PASO 0: LO QUE NECESITAS TENER LISTO

Antes de empezar:

```
☑️  Google Sheets con la hoja "Seguimiento Bot" (ya lo tienes)
☑️  Cuenta de Gmail (para n8n)
☑️  Navegador (Chrome, Firefox, Edge)
☑️  15 minutos
☑️  Los $15 USD disponibles para Twilio
```

**NO necesitas:**
- PowerShell (se puede usar desde el navegador)
- Instalar nada en tu computadora
- Conocimiento de programación

---

## 📋 PASO 1: CREAR CUENTA TWILIO (5 minutos)

### 1.1 Ir a Twilio

Abre: https://www.twilio.com

Clic en **"Sign up"** (esquina superior derecha)

### 1.2 Llenar el formulario

```
Email: tu correo
Contraseña: crea una fuerte
Nombre completo: tu nombre real
País: Tu país
Aceptar términos: ✓ Sí
```

Clic en **"Create account"**

### 1.3 Verificar correo

Twilio te envía un email de confirmación.
Clic en el enlace del email para confirmar.

### 1.4 Responder preguntas

Cuando inicies sesión, Twilio preguntará:
```
¿Qué harás? 
→ Responde: "Enviar mensajes de WhatsApp"

¿Para qué proyecto?
→ Responde: "Pruebas de bot para empleabilidad"

¿Cuántos mensajes?
→ Responde: "Menos de 1000 por mes"
```

### 1.5 Panel de Twilio

Al terminar, estarás en el **Dashboard** de Twilio.

**Nota:** Twilio te da $15.50 USD de crédito de prueba (trial).
Este crédito **expira en 30 días**, así que úsalo rápido.

---

## 📱 PASO 2: CONFIGURAR WHATSAPP EN TWILIO (10 minutos)

### 2.1 Ir a WhatsApp Sandbox

En el panel de Twilio (lado izquierdo), busca:

```
Messaging → Explore → Try it out → Sandbox
```

O directo: https://console.twilio.com/us1/develop/sms/try-it-out/whatsapp-sandbox

### 2.2 Tu número Twilio

Verás algo como:

```
Your Sandbox WhatsApp Number: +1 415 523 8886
(Este número es TUYO para pruebas)
```

**Guarda este número** — lo usarás en n8n.

### 2.3 Conectar tu teléfono

Para **recibir y enviar mensajes**, debes conectar TU teléfono.

Sigue las instrucciones en pantalla:

1. Abre WhatsApp en tu teléfono
2. Crea un chat con el número de Twilio
3. Escribe el código que te muestra Twilio (ejemplo: `join cream-flavor`)
4. Envía el mensaje

Cuando confirmes, verás: ✅ **Connected to Sandbox** en verde

### 2.4 Obtener credenciales Twilio

Ahora necesitas dos cosas para n8n:

En el Dashboard de Twilio (panel izquierdo):

```
Account menu (abajo a la izquierda)
↓
Account SID: (cópialo)

Auth Token: (cópialo — haz clic en el ojo para verlo)
```

Guárdalos en un archivo de texto seguro (NO en GitHub).

---

## 🔧 PASO 3: CREAR CUENTA N8N (5 minutos)

### 3.1 Ir a n8n Cloud

Abre: https://n8n.cloud

Clic en **"Get Started for Free"**

### 3.2 Registrarte

```
Email: tu correo (puede ser el mismo de Twilio)
Contraseña: crea una
Nombre: tu nombre
```

Clic en **"Sign Up"**

### 3.3 Verificar correo

n8n te envía un email. Clic en el enlace.

### 3.4 Ya estás dentro

Te llevará al panel de n8n. Verás:

```
"Start building workflows"
```

---

## ⚙️ PASO 4: CREAR EL PRIMER WORKFLOW (10 minutos)

### 4.1 Crear workflow

En el panel de n8n:

Clic en **"+ New workflow"**

Te abre un canvas en blanco.

### 4.2 Agregar trigger (inicio del flujo)

En la esquina superior izquierda, ves:

```
"+ Add first step"
```

Clic ahí.

Busca: **"Webhook"**

Selecciona: **"Webhook"** (la primera opción)

Se agrega un nodo al canvas.

### 4.3 Configurar Webhook

En el panel de la derecha, verás opciones. Busca:

```
HTTP method: GET
↓ (cambia a)
HTTP method: POST
```

Guarda (Cmd+S o Ctrl+S).

Verás que aparece una URL como:

```
https://n8n.example.com/webhook/xxxxx
```

**Esta es tu URL de webhook.** Guárdala.

### 4.4 Agregar nodo de respuesta

Haz clic en el nodo Webhook.

En la esquina inferior derecha, ves dos flechas: **"Execute Workflow"** (play).

Antes de eso, conecta un segundo nodo.

Clic en el punto blanco del lado derecho del nodo Webhook.

Arrastra hasta una zona vacía del canvas.

Busca: **"Respond to Webhook"**

Selecciona ese nodo.

### 4.5 Configurar respuesta

En el panel derecho del nodo "Respond to Webhook":

```
Body: 
{
  "status": "ok"
}
```

Guarda (Cmd+S / Ctrl+S).

### 4.6 Prueba el webhook

Clic en el botón **"Execute Workflow"** (play arriba).

El workflow se ejecuta. Si ves ✅ en ambos nodos = funciona.

---

## 📤 PASO 5: CONECTAR TWILIO + ENVIAR PRIMER MENSAJE (10 minutos)

### 5.1 Estructura del workflow

Ahora agregaremos **entre el Webhook y la respuesta**:
- **Google Sheets**: leer teléfono
- **Twilio**: enviar mensaje

### 5.2 Agregar nodo Google Sheets

Clic en el punto blanco del lado derecho del nodo **Webhook**.

Arrastra a zona vacía.

Busca: **"Google Sheets"**

Selecciona: **"Google Sheets"**

### 5.3 Autenticar Google Sheets

En el panel derecho, verás:

```
Authentication: [Connect]
```

Clic en **"Connect"**.

Se abre una ventana de Google. 

Elige tu cuenta de Gmail.

Da permisos cuando pida.

Vuelve a n8n automáticamente.

### 5.4 Configurar lectura de Sheets

En el panel del nodo Google Sheets:

```
Operation: Read
Spreadsheet ID: [pega tu ID de Sheets]
Sheet Name: Seguimiento Bot
Range: A1:T100
```

¿Dónde está el Spreadsheet ID?

Abre tu Sheet → mira la URL:

```
https://docs.google.com/spreadsheets/d/
                                    ↑ ESTE ID (copia hasta antes de /edit)
```

### 5.5 Agregar nodo Twilio

Clic en el punto blanco del lado derecho del nodo Google Sheets.

Arrastra a zona vacía.

Busca: **"Twilio"**

Selecciona: **"Twilio"**

### 5.6 Autenticar Twilio

En el panel derecho:

```
Authentication: [New Credentials]
```

Clic ahí.

Se abre un formulario:

```
Account SID: [pega el que guardaste en Paso 2]
Auth Token: [pega el que guardaste en Paso 2]
```

Clic en **"Create new credential"**

Vuelve al nodo. Ahora selecciona esa credencial.

### 5.7 Configurar envío de WhatsApp

En el panel del nodo Twilio:

```
Operation: Send Message
Channel: WhatsApp

From: [número Twilio, ejemplo +1 415 523 8886]

To: [número destino — empieza con +502, ejemplo +502 7123456789]

Message: Hola {{$json.Nombre}}! Este es un test del bot.
         
(Reemplaza "Nombre" por la columna de tu Sheet que tiene nombres)
```

### 5.8 Guardar workflow

Clic en el botón **"Save"** (arriba).

Te pedirá nombre:

```
Nombre: Test Bot WhatsApp
```

Clic en **"Save"**.

---

## 🧪 PASO 6: PRUEBA EL BOT (5 minutos)

### 6.1 Ejecutar workflow

Clic en el botón **"Execute Workflow"** (play arriba).

Los nodos se ejecutan de arriba a abajo:
1. Webhook recibe datos
2. Google Sheets lee
3. Twilio envía

### 6.2 ¿Llegó el mensaje?

Abre WhatsApp en tu teléfono.

Busca el chat con el número Twilio (ejemplo: +1 415 523 8886).

¿Ves un mensaje? ✅ **¡FUNCIONA!**

¿No ves nada? Revisa la sección **Errores** abajo.

### 6.3 Responder desde WhatsApp

Envía un mensaje desde tu teléfono al número Twilio.

Ese mensaje se captura en n8n (para la siguiente fase).

---

## 📊 PASO 7: CONECTAR AUTOMÁTICAMENTE (próxima fase)

Una vez que pruebes que funciona:

### 7.1 Cambiar Webhook a Cron (ejecución automática)

En lugar de esperar un webhook, usa:

```
Trigger: Cron
Programación: Cada día a las 9:00 AM
Acción: Leer Seguimiento Bot y enviar mensajes
```

### 7.2 Filtrar filas con fecha de hoy

Agrega un nodo de **Filter** que diga:

```
Si Fecha Seg1 = hoy
↓
Entonces enviar mensaje
```

### 7.3 Guardar respuestas

Agrega un nodo Google Sheets final que:

```
Escriba en la columna "Resp S1 P1" la respuesta recibida
```

---

## ⚠️ ERRORES COMUNES Y SOLUCIONES

### ❌ "No aparece el mensaje en WhatsApp"

**Causa:** Número Twilio incorrecto o no conectado al sandbox

**Solución:**
1. Vuelve a https://console.twilio.com/us1/develop/sms/try-it-out/whatsapp-sandbox
2. Verifica que TU teléfono esté en ✅ verde "Connected to Sandbox"
3. Si no, envía el código de nuevo desde WhatsApp

### ❌ "Error de autenticación en Google Sheets"

**Causa:** Credenciales expiradas o permisos insuficientes

**Solución:**
1. Clic en el nodo Google Sheets
2. Clic en "Authentication" → "Delete"
3. Vuelve a conectar con tu cuenta

### ❌ "Error en Twilio: Invalid phone number"

**Causa:** Número de teléfono mal formateado

**Solución:**
Usa formato: `+[código país][número]`
- Guatemala: `+502 7123456789`
- México: `+52 5512345678`
- El Salvador: `+503 62123456`

### ❌ "Se acabó el dinero Twilio ($15)"

**Causa:** Enviaste demasiados mensajes

**Solución:**
- Cada mensaje WhatsApp = $0.01 USD
- $15 = ~1500 mensajes
- Para más, compra un plan pagado

---

## 💰 PRESUPUESTO DETALLADO

Con los $15 de Twilio:

```
Costo por mensaje WhatsApp: $0.01 USD

Ejemplos:
- 100 mensajes   = $1.00
- 500 mensajes   = $5.00
- 1000 mensajes  = $10.00
- 1500 mensajes  = $15.00

Recomendación:
Usa los primeros $5 para configuración y pruebas.
Deja $10 para pruebas reales con el equipo.
```

---

## ✅ CHECKLIST: YA CASI TERMINAS

Cuando llegues aquí, habrás logrado:

- [ ] Cuenta Twilio creada con número WhatsApp
- [ ] Cuenta n8n creada
- [ ] Primer workflow en n8n
- [ ] Google Sheets conectado a n8n
- [ ] Twilio conectado a n8n
- [ ] Primer mensaje WhatsApp enviado

**Próximo paso:** Agregar trigger automático (cron) y filtros.

---

## 📞 SOPORTE RÁPIDO

Si algo falla:

1. **n8n:** https://docs.n8n.io/
2. **Twilio:** https://www.twilio.com/docs/
3. **Google Sheets API:** https://developers.google.com/sheets/api

O prueba en el **Community Forum** de n8n:
https://community.n8n.io/

---

**Tiempo total:** ~45 minutos (primera vez)  
**Costo:** Gratis (Twilio trial $15)  
**Resultado:** Bot WhatsApp funcional para pruebas

¿Necesitas ayuda en algún paso? Avísame cuál y te lo detallo más.
