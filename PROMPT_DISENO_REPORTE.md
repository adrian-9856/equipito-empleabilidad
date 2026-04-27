# PROMPT DE DISEÑO — REPORTE EQUIPITO EMPLEABILIDAD

> Copia y pega este prompt en Canva AI, ChatGPT, Figma AI, o cualquier herramienta de diseño.
> Ajusta los datos con los números reales antes de presentar.

---

## PROMPT PARA CANVA AI / DISEÑO VISUAL

```
Diseña un reporte visual de empleabilidad con las siguientes características:

ORGANIZACIÓN: Equipito Empleabilidad — programa de inserción laboral para jóvenes.
ESTILO: Moderno, profesional, minimalista. Colores corporativos: azul #1a73e8 y verde #0f9d58.
FORMATO: Presentación tipo informe ejecutivo (A4 horizontal o slides 16:9).

SECCIONES QUE DEBE INCLUIR:

1. PORTADA
   - Logo/nombre: Equipito Empleabilidad
   - Título: "Reporte de Seguimiento de Graduados"
   - Período: [MES AÑO]
   - Subtítulo: "Programa de Inserción Laboral"

2. INDICADORES CLAVE (KPIs) — tarjetas grandes con número y etiqueta
   - Total Graduados registrados: [NÚMERO]
   - Empleados activos: [NÚMERO]
   - Tasa de empleo: [%]
   - Seguimientos completados: [NÚMERO]

3. DISTRIBUCIÓN POR ETAPA — gráfico de dona o barras horizontales
   Etapas: Aliados / Plataforma / Derivaciones / Activamente busca trabajo /
           Paso a paso / Conexiones Laborales
   Mostrar: número y % por etapa

4. CLASIFICACIÓN DE PERFILES — gráfico de semáforo o barras apiladas
   - Perfil Verde (bajas barreras): [NÚMERO]
   - Perfil Amarillo (barreras medias): [NÚMERO]
   - Perfil Rojo (altas barreras): [NÚMERO]

5. SEGUIMIENTOS BOT — tabla o timeline
   - Total activos en seguimiento: [NÚMERO]
   - Post-empleo: [NÚMERO] | Búsqueda activa: [NÚMERO]
   - Mensajes enviados este mes: [NÚMERO]
   - Llamadas pendientes: [NÚMERO]

6. CONEXIONES LABORALES — tabla de empleos conseguidos
   Columnas: Nombre | Empresa | Cargo | Salario | Fecha inicio
   (máximo 10 registros recientes)

7. SATISFACCIÓN LABORAL — gauge o estrellas
   - Puntaje promedio AAPI: [X.X / 5]
   - Respuestas recibidas: [NÚMERO]

8. NOTAS Y PRÓXIMOS PASOS
   - Espacio para texto libre del equipo

DISEÑO:
- Cada sección en su propia tarjeta/card con sombra suave
- Tipografía: Google Fonts (Poppins o Inter)
- Paleta: Azul #1a73e8, Verde #0f9d58, Gris claro #f5f5f5, Blanco
- Usar iconos lineales (no rellenos)
- Sin tablas recargadas — preferir visualizaciones
```

---

## PROMPT PARA CHATGPT / CLAUDE (generar texto del reporte)

```
Actúa como analista de datos de un programa de empleabilidad juvenil.
Genera el texto narrativo de un reporte mensual con los siguientes datos:

DATOS DEL MES DE [MES/AÑO]:
- Total graduados: [NÚMERO]
- Nuevos este mes: [NÚMERO]
- Empleados activos: [NÚMERO]
- Tasa de empleo: [%]
- Perfil Verde: [NÚMERO] | Amarillo: [NÚMERO] | Rojo: [NÚMERO]
- Seguimientos enviados: [NÚMERO]
- Respuestas recibidas: [NÚMERO]
- Llamadas completadas: [NÚMERO]

INCLUYE:
1. Resumen ejecutivo (3 oraciones, tono optimista y profesional)
2. Logros del mes (3-5 puntos en viñetas)
3. Desafíos identificados (2-3 puntos)
4. Recomendaciones para el próximo mes (2-3 acciones concretas)
5. Nota motivacional para el equipo (1 párrafo breve)

TONO: Profesional pero cercano, orientado a impacto social.
IDIOMA: Español.
LONGITUD: Máximo 1 página.
```

---

## PROMPT PARA PRESENTACIÓN ANTE DONANTES / DP

```
Crea una presentación ejecutiva de 5 slides para presentar ante donantes
o directores de programa (DP) sobre el avance del programa de empleabilidad.

DATOS A INCLUIR: [PEGAR DATOS DEL MES]

ESTRUCTURA DE SLIDES:

SLIDE 1 — IMPACTO EN NÚMEROS
Tres datos grandes: total graduados / empleados / tasa de éxito
Frase de impacto: "[X] jóvenes con nueva oportunidad laboral este mes"

SLIDE 2 — ¿QUIÉNES SON?
Distribución de perfiles (Verde/Amarillo/Rojo)
Breve descripción de qué significa cada perfil

SLIDE 3 — EL PROCESO
Diagrama de flujo simple:
Graduado → Clasificación → Seguimiento Bot → Empleo → Seguimiento post-empleo

SLIDE 4 — RESULTADOS ESTE MES
Tabla concisa: Empleos conseguidos con empresa/cargo/salario
Estadística de satisfacción laboral (AAPI)

SLIDE 5 — PRÓXIMOS PASOS
3 acciones prioritarias del próximo mes
Necesidades del programa (recursos, aliados, etc.)

ESTILO: Limpio, mucho espacio en blanco, sin texto corrido.
Cada slide: máximo 30 palabras de texto + visualización.
```

---

## ESTRUCTURA JSON PARA GOOGLE SHEETS → REPORTE AUTOMÁTICO

Esta estructura es lo que el script de Apps Script ya genera en la hoja "Reporte".
Puedes usarla también para exportar a otras herramientas:

```json
{
  "reporte": {
    "periodo": "Abril 2026",
    "generado": "2026-04-27T09:00:00",
    "resumen": {
      "total_graduados": 0,
      "empleados_activos": 0,
      "tasa_empleo_pct": 0,
      "nuevos_este_mes": 0
    },
    "por_etapa": {
      "Aliados": 0,
      "Plataforma": 0,
      "Derivaciones": 0,
      "Activamente busca trabajo": 0,
      "Paso a paso": 0,
      "Conexiones Laborales": 0
    },
    "clasificacion_perfiles": {
      "Verde": 0,
      "Amarillo": 0,
      "Rojo": 0
    },
    "seguimiento_bot": {
      "total_activos": 0,
      "post_empleo": 0,
      "busqueda_activa": 0,
      "mensajes_enviados_mes": 0,
      "llamadas_pendientes": 0,
      "respuestas_recibidas": 0
    },
    "satisfaccion_laboral": {
      "puntaje_promedio_aapi": 0,
      "respuestas_totales": 0
    },
    "conexiones_laborales_recientes": [
      {
        "nombre": "",
        "empresa": "",
        "cargo": "",
        "salario": "",
        "fecha_inicio": ""
      }
    ]
  }
}
```

---

## COLORES DE REFERENCIA POR HOJA

Úsalos para mantener consistencia visual en el reporte:

| Hoja | Color | Hex |
|------|-------|-----|
| Graduados | Azul | `#1a73e8` |
| Aliados | Índigo | `#3f51b5` |
| Plataforma | Cian | `#00bcd4` |
| Derivaciones | Naranja | `#ff9800` |
| Activamente busca trabajo | Verde | `#0f9d58` |
| Conexiones Laborales | Naranja oscuro | `#e65100` |
| Seguimiento Bot | Verde azulado | `#00897b` |
| Clasificación de Perfiles | Azul índigo | `#5c6bc0` |
| Satisfacción Empleo | Verde | `#00897b` |
| Sesiones Acompañamiento | Púrpura | `#5e35b1` |
| Reporte | Rosa/Magenta | `#e91e63` |

---

## INSTRUCCIONES DE USO

1. **Para presentar al equipo:** Usa el primer prompt (Canva) y reemplaza los `[NÚMERO]` con datos reales del mes
2. **Para presentar a donantes:** Usa el tercer prompt (5 slides)
3. **Para generar texto narrativo:** Usa el segundo prompt (ChatGPT/Claude)
4. **Para exportar datos a otro sistema:** Usa la estructura JSON

> El reporte automático se genera desde Google Sheets:
> Menú → 📊 Seguimiento Graduados → 📊 Generar Reporte
