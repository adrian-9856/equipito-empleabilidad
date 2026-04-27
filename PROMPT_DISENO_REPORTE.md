# PROMPT DE DISEÑO VISUAL — REPORTES EQUIPITO EMPLEABILIDAD

> Prompts genéricos para diseñar reportes visualmente atractivos.
> Funcionan para cualquier DP o programa. Solo cambia el nombre del programa y los datos.

---

## 1. PROMPT PARA CANVA AI

```
Diseña un reporte visual profesional para un programa social con estas especificaciones:

ESTILO GENERAL:
- Diseño moderno, limpio y minimalista
- Mucho espacio en blanco entre secciones
- Tipografía: una fuente sans-serif para títulos (bold) y una ligera para cuerpo
- Colores: un color primario fuerte + blanco + gris claro (#f5f5f5) como fondo de cards
- Sin tablas con bordes recargados — usar cards con sombra suave
- Iconos lineales (outline), no rellenos
- Formato: A4 vertical u horizontal según el contenido

ESTRUCTURA DE PÁGINAS:

PÁGINA 1 — PORTADA
- Nombre del programa (título grande, color primario)
- Subtítulo: tipo de reporte (mensual / trimestral / anual)
- Período que cubre
- Logo o ícono representativo centrado
- Franja de color en la parte inferior o lateral

PÁGINA 2 — INDICADORES CLAVE (KPIs)
- 4 tarjetas grandes en fila: número grande arriba, etiqueta abajo
- Fondo de cada tarjeta: blanco con borde izquierdo del color del indicador
- Espacio para un párrafo breve de contexto debajo

PÁGINA 3 — DISTRIBUCIÓN / AVANCE
- Gráfico de barras horizontal o dona
- Leyenda clara a la derecha o abajo
- Porcentaje visible en cada segmento
- Título de sección con ícono a la izquierda

PÁGINA 4 — TABLA DE RESULTADOS
- Máximo 10 filas
- Columnas: 3-4 datos clave, no más
- Filas alternadas (una blanca, una gris muy claro)
- Sin bordes excepto línea inferior de cada fila

PÁGINA 5 — NOTAS Y PRÓXIMOS PASOS
- Dos columnas: Logros | Desafíos
- Lista de viñetas con íconos de check y alerta
- Espacio para 3 acciones del próximo período

REGLAS DE DISEÑO:
- Máximo 30 palabras de texto por sección (los datos hablan solos)
- Cada número importante debe ser visualmente grande (mínimo 36pt)
- Paleta de máximo 3 colores + blanco + gris
- Consistencia total entre páginas (mismo header, mismo footer con número de página)
```

---

## 2. PROMPT PARA FIGMA / DISEÑADOR

```
Necesito un sistema de diseño para reportes internos de un programa social.
Debe ser reutilizable para diferentes programas y períodos.

COMPONENTES QUE NECESITO:

1. CARD DE KPI
   - Variantes: pequeña (2x2 cm) / mediana / grande
   - Elementos: número principal, etiqueta, ícono opcional, línea de color lateral
   - Estado: neutro / positivo (verde) / atención (naranja) / crítico (rojo)

2. GRÁFICO DE PROGRESO
   - Barra horizontal con porcentaje al final
   - Variante: dona con número central
   - Etiquetas: nombre de categoría a la izquierda, valor a la derecha

3. TABLA LIMPIA
   - Header con fondo color primario y texto blanco
   - Filas alternas blanco / gris muy claro
   - Sin bordes exteriores (solo separadores internos ligeros)

4. BADGE DE ESTADO
   - Colores: Verde (completado) / Amarillo (en progreso) / Rojo (pendiente)
   - Forma: pill redondeado, texto en 10pt

5. SECCIÓN HEADER
   - Título de sección + ícono + línea separadora

6. LAYOUT DE PÁGINA
   - Margen: 24px en todos los lados
   - Grid: 12 columnas
   - Espaciado entre secciones: 32px

TOKENS DE DISEÑO:
   - Color primario: a definir por programa
   - Texto principal: #212121
   - Texto secundario: #757575
   - Fondo de página: #ffffff
   - Fondo de cards: #f9f9f9
   - Sombra de cards: 0 2px 8px rgba(0,0,0,0.08)
   - Border radius: 8px
```

---

## 3. PROMPT PARA CHATGPT / CLAUDE (texto del reporte)

```
Actúa como redactor de impacto social. Escribe el texto de un reporte mensual
para un programa de [NOMBRE DEL PROGRAMA].

DATOS DEL PERÍODO [MES/AÑO]:
[PEGAR AQUÍ LOS NÚMEROS DEL MES]

ESCRIBE ESTAS SECCIONES:

1. TITULAR DE IMPACTO (1 oración)
   Formato: "[X] personas [logro] gracias a [programa]"

2. RESUMEN EJECUTIVO (3 oraciones)
   - Qué se hizo
   - Cuál fue el resultado más importante
   - Qué sigue

3. LOGROS DEL MES (máximo 5 viñetas)
   Cada viñeta: verbo en pasado + dato concreto + impacto
   Ejemplo: "Acompañamos a 12 personas en búsqueda activa, logrando 4 nuevos empleos"

4. DESAFÍOS (máximo 3 viñetas)
   Tono honesto pero constructivo, no alarmista

5. PRÓXIMAS ACCIONES (3 puntos)
   Específicas, con responsable implícito y fecha si aplica

TONO: Profesional, cercano, orientado a personas (no a procesos).
No usar palabras como "sinergia", "robusto", "implementar" — usar lenguaje simple.
Máximo 200 palabras en total.
```

---

## 4. PROMPT PARA PRESENTACIÓN A DONANTES / DP (5 slides)

```
Crea una presentación ejecutiva de 5 slides para mostrar avance de un programa social
ante [donantes / directores de programa / aliados].

PÚBLICO: personas que toman decisiones, tienen poco tiempo, quieren ver impacto.

SLIDE 1 — EL IMPACTO EN NÚMEROS
Diseño: 3 números grandes centrados, cada uno con su etiqueta
Datos: las 3 métricas más importantes del período
Abajo: 1 frase de impacto humano

SLIDE 2 — ¿A QUIÉN ATENDEMOS?
Diseño: gráfico de dona o barras con categorías de personas
No más de 5 categorías
Cada categoría con color distinto y porcentaje visible

SLIDE 3 — EL PROCESO (cómo funciona)
Diseño: diagrama de flujo horizontal, máximo 5 pasos
Íconos simples en cada paso
Sin texto extenso — solo etiquetas de 2-3 palabras

SLIDE 4 — RESULTADOS CONCRETOS
Diseño: tabla simple con los 5-8 logros más importantes
Columnas: qué / cuánto / cuándo
Una fila destacada (el mayor logro) con fondo de color

SLIDE 5 — QUÉ NECESITAMOS / QUÉ SIGUE
Diseño: dos columnas
Izquierda: próximas 3 acciones (con fecha)
Derecha: recursos o apoyo necesario

REGLAS:
- Máximo 15 palabras por slide (excluyendo tabla/gráfico)
- Fondo blanco o muy claro
- Un solo color de acento por slide
- Sin animaciones complejas
- Fuente grande: mínimo 24pt para texto visible en sala
```

---

## 5. GUÍA DE COLORES POR PROGRAMA

Asigna un color primario diferente a cada programa para diferenciarlos visualmente:

| Programa | Color sugerido | Hex | Por qué |
|----------|---------------|-----|---------|
| Empleabilidad general | Azul confianza | `#1a73e8` | Profesional, estable |
| Empleo juvenil | Verde esperanza | `#0f9d58` | Crecimiento, futuro |
| Inserción mujeres | Violeta | `#7b1fa2` | Fuerza, dignidad |
| Emprendimiento | Naranja energía | `#e65100` | Acción, dinamismo |
| Acompañamiento social | Teal | `#00897b` | Calma, apoyo |
| Formación técnica | Azul índigo | `#3f51b5` | Conocimiento |

**Regla:** Un programa = un color primario.
Úsalo en: header de la portada, borde de KPIs, header de tabla, títulos de sección.

---

## 6. CHECKLIST DE CALIDAD VISUAL

Antes de enviar cualquier reporte, verifica:

- [ ] ¿El número más importante es el más grande visualmente?
- [ ] ¿Puedo entender el reporte en 30 segundos sin leer el texto?
- [ ] ¿Hay consistencia de fuentes? (máximo 2 tipografías)
- [ ] ¿Los colores tienen suficiente contraste? (texto sobre fondo)
- [ ] ¿Las gráficas tienen etiquetas legibles sin necesitar leyenda?
- [ ] ¿El logo / nombre del programa está visible en cada página?
- [ ] ¿El período del reporte está claro desde la portada?
- [ ] ¿Hay número de página o indicador de sección?
- [ ] ¿El archivo pesa menos de 5MB para poder enviarse por correo?

---

## HERRAMIENTAS RECOMENDADAS

| Herramienta | Para qué | Costo |
|-------------|---------|-------|
| **Canva** | Diseño rápido sin diseñador | Gratis / Pro $13/mes |
| **Google Slides** | Colaboración en equipo | Gratis |
| **Figma** | Sistema de diseño reutilizable | Gratis (hasta 3 proyectos) |
| **Looker Studio** | Reportes conectados a Sheets en tiempo real | Gratis |
| **PowerPoint** | Si ya lo tiene el equipo | Licencia Office |

**Recomendación para Equipito:** Canva para reportes mensuales rápidos + Looker Studio para dashboards conectados a Google Sheets en tiempo real (sin copiar datos manualmente).
