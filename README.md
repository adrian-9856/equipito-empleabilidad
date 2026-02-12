# Equipito Empleabilidad

Hojas de apilabilidad para el seguimiento de personas en procesos de empleabilidad.

---

## Cómo usar el script

1. Abre tu **Google Sheet**
2. Ve a **Extensiones > Apps Script**
3. Copia y pega el contenido de `apilabilidad.gs`
4. Ejecuta la función **`configurarHojasApilabilidad()`**
5. Acepta los permisos que solicita Google

El script eliminará todas las hojas existentes y creará las 6 hojas nuevas con su configuración.

---

## Estructura de las hojas

### Columnas comunes en TODAS las hojas (Información principal)

| Campo | Tipo |
|---|---|
| Fecha de ingreso | Fecha (dd/mm/yyyy) |
| Nombre completo | Texto |
| ID | Automático (ej: ALI-001) |
| Género | Texto |
| Edad | Número |
| Nivel educativo | Texto |

---

### 1. Aliados

| Campo | Tipo |
|---|---|
| *(Columna pendiente de nombre)* | Desplegable Sí/No |
| Trámites | Texto |
| Entrevista | Texto |
| Día de prueba | Texto |
| Confirmación de recepción | Texto |
| Notas | Texto |
| Activo | Desplegable Sí/No |

---

### 2. Plataforma

| Campo | Tipo |
|---|---|
| Cita | Checkbox |
| Creación de perfil | Texto |
| Contacto | Checkbox |
| Trámites | Texto |
| Entrevista | Texto |
| Confirmación | Texto |
| Recepción | Texto |
| Nota | Texto |
| Activo | Desplegable Sí/No |

---

### 3. Derivaciones

| Campo | Tipo |
|---|---|
| Envío 1 | Texto |
| Llamada 1 | Texto |
| Envío 2 | Texto |
| Llamada 2 | Texto |
| Envío 3 | Texto |
| Llamada 3 | Texto |
| Notas | Texto |
| Activo | Desplegable Sí/No |

---

### 4. Por su cuenta

| Campo | Tipo |
|---|---|
| Mensaje | Texto |
| Llamada | Texto |
| Nota | Texto |
| Activo | Desplegable Sí/No |

---

### 5. No busca trabajo - Fito

| Campo | Tipo |
|---|---|
| DPI | Texto |
| Número de teléfono | Texto |
| Formación | Texto |
| Cohorte | Texto |
| Nota | Texto |
| Activo | Desplegable Sí/No |

---

### 6. Activamente busca trabajo

| Campo | Tipo |
|---|---|
| Entrevista | Texto |
| Trámites | Texto |
| Activo | Desplegable Sí/No |

---

## Pendientes

- [ ] Definir el nombre de la columna "Pendiente (nombre por definir)" en la hoja **Aliados**
- [ ] Agregar sección de **Seguimiento / Acción** en cada hoja (por definir)
