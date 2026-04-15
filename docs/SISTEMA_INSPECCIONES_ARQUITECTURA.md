# 📋 SISTEMA DE INSPECCIONES - ARQUITECTURA Y DISEÑO
**Proyecto:** TH-HC-001 EN PROCESO DE VALIDACIÓN  
**Fecha creación:** 12/03/2026  
**Última actualización:** 12/03/2026

---

## 🗂️ UBICACIÓN DE TABLAS (ListObject)

| Tabla | Hoja Excel | Tipo | Propósito |
|-------|------------|------|-----------|
| **tblSecciones** | Configuración | Catálogo | Define tipos de secciones de preguntas |
| **tblOpcionesDeRespuesta** | Configuración | Catálogo | Opciones disponibles por sección y sus valores |
| **tblCriticidad** | Configuración | Catálogo | Niveles de criticidad de preguntas |
| **tblCategoriasRPN** | Configuración | Catálogo | Categorización de resultados RPN |
| **tblConfiguracion** | Configuración | Parámetros | Configuración global del sistema |
| **tblPersonal** | Personal | Maestra | Personal sujeto a inspecciones |
| **tblPlantillas** | Checklist | Maestra | Tipos de checklist por puesto |
| **tblPreguntas** | Checklist | Maestra | Preguntas de checklist por plantilla |
| **tblInspecciones** | Historico | Transaccional | Cabecera de inspecciones ejecutadas |
| **tblRespuestas** | Historico | Transaccional | Detalle de respuestas por inspección |

---

## 📐 ESTRUCTURA DE TABLAS

### **1. tblSecciones** (Hoja: Configuración)
Define los tipos de sección que puede tener un checklist.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| ID Seccion | Texto (UUID) | PK - Identificador único |
| Nombre de sección | Texto | Nombre descriptivo |
| Tipo de respuesta | Texto | "Selección" o "Puntaje" |

**Datos actuales:**
```
ID Seccion                  | Nombre                          | Tipo de respuesta
─────────────────────────────────────────────────────────────────────────────────
6bzXQZjA-EG9wLOz7-GfX01pco  | Auditoría de procesos           | Selección
J0Wjpqx8-fcEPr0g7-n4MIuktw  | Auditoría de técnica aséptica   | Selección
jtUnipGu-U2IUjaAJ-KnEJv7Bz  | RPN                             | Puntaje
```

---

### **2. tblOpcionesDeRespuesta** (Hoja: Configuración)
Opciones de respuesta disponibles por sección con sus valores numéricos.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| ID Opcion | Texto (UUID) | PK - Identificador único |
| ID Seccion | Texto (UUID) | FK → tblSecciones.ID Seccion |
| Opción texto | Texto | Texto de la opción |
| Valor puntaje | Texto/Número | Valor numérico o "-" si no aplica |

**Datos actuales:**
```
AUDITORÍA DE PROCESOS (6bzXQZjA-EG9wLOz7-GfX01pco):
─────────────────────────────────────────────────────────
ID Opcion                   | Opción texto | Valor puntaje
WYKkxXH0-cetasqdd-GLjOwHtB  | No Cumple    | -
wZUEAg0S-E3QWV99w-TjFdnsKL  | Cumple       | -
tQN8gk2D-mzSyvI2y-hg3RBKk0  | No Aplica    | -

AUDITORÍA DE TÉCNICA ASÉPTICA (J0Wjpqx8-fcEPr0g7-n4MIuktw):
─────────────────────────────────────────────────────────────────
ID Opcion                   | Opción texto | Valor puntaje
F7QWv7w6-E20cWC5G-Er9alZKP  | Si           | 0
CYIqbbyR-my2OwcYI-1owOAil4  | No           | 4
oiRJhUYq-OIKmjhNB-89IUHngB  | No Aplica    | -1
```

---

### **3. tblCriticidad** (Hoja: Configuración)
Niveles de criticidad asignados a preguntas.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| ID Criticidad | Texto (UUID) | PK - Identificador único |
| Nombre de criticidad | Texto | Nombre descriptivo |
| Valor | Número | Valor numérico asociado |

**Datos actuales:**
```
ID Criticidad               | Nombre      | Valor
─────────────────────────────────────────────────
ZADo553G-VOuS3n4q-Q3bWUfz5  | Menor       | 1
2xzjgNbw-ENZUmhQF-mKjVM6Iy  | Mayor       | 4
M4dcAe5B-wo5vnoDp-eiEoqu5p  | Crítica     | 0
```

---

### **4. tblCategoriasRPN** (Hoja: Configuración)
Categorización de resultados según RPN calculado.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| ID Categoria | Texto (UUID) | PK - Identificador único |
| Numero categoria | Número | 1, 2, 3, 4, 5 |
| Nombre categoria | Texto | Nombre descriptivo |
| RPN minimo | Número decimal | Límite inferior del rango |
| RPN maximo | Número decimal | Límite superior del rango |
| Requiere historico | Sí/No | Si requiere validación de inspecciones anteriores |
| Cantidad inspecciones | Número | Cantidad de inspecciones consecutivas a validar |
| Color hex | Texto | Color para formato condicional |
| Descripcion | Texto | Descripción detallada |
| Orden | Número | Orden de presentación |

**Datos configurados:**
```
Numero | Nombre       | RPN Min | RPN Max | Req Hist | Cant Insp | Descripción
────────────────────────────────────────────────────────────────────────────────
1      | Categoría 1  | 0       | 14      | No       | 0         | Desempeño óptimo
2      | Categoría 2  | 15      | 19      | No       | 0         | Desempeño aceptable
3      | Categoría 3  | 20      | 40      | No       | 0         | Requiere atención
4      | Categoría 4  | 40.01   | 999     | No       | 0         | Crítico inmediato
5      | Categoría 5  | 20      | 999     | Sí       | 3         | Crítico recurrente (3 inspecciones >20)
```

**Nota especial Categoría 5:**  
Se asigna cuando hay 3 inspecciones consecutivas (mismo Personal + misma Plantilla) con RPN > 20.

---

### **5. tblConfiguracion** (Hoja: Configuración)
Parámetros globales del sistema.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| Clave | Texto | PK - Nombre del parámetro |
| Valor | Texto | Valor del parámetro |
| Categoria | Texto | Agrupación lógica |
| Descripcion | Texto | Para qué sirve |
| Tipo dato | Texto | Numero/Texto/Fecha/Logico |

**Parámetros sugeridos:**
```
Clave                              | Valor | Categoria      | Descripción
───────────────────────────────────────────────────────────────────────────────
PUNTAJE_MAXIMO_TA_BASE            | 57    | Scoring        | Máximo base de puntos en Técnica Aséptica
TA_NO_APLICA_CUENTA_COMO_NO       | Si    | Scoring        | Si No Aplica resta del denominador con valor de No
ALERTAR_DIAS_ANTES_VENCIMIENTO    | 15    | Notificaciones | Días de anticipación para alertar
FRECUENCIA_DEFAULT_MESES          | 3     | Inspecciones   | Frecuencia por defecto si no está en plantilla
VALIDAR_PERSONA_ACTIVA            | Si    | Seguridad      | Solo permitir inspecciones a personal activo
```

---

### **6. tblPersonal** (Hoja: Personal)
Personal sujeto a inspecciones. Estructura con puestos en columnas.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| Iniciales | Texto | PK - Identificador único (ej: ACF) |
| Planta | Texto | Planta principal asignada |
| Quimico | Sí/No | Si puede ser inspeccionado como Químico |
| Digitador | Sí/No | Si puede ser inspeccionado como Digitador |
| Etiquetado | Sí/No | Si puede ser inspeccionado en Etiquetado |
| Ayudante 2 | Sí/No | Si puede ser inspeccionado como Ayudante 2 |
| Ayudante 1 | Sí/No | Si puede ser inspeccionado como Ayudante 1 |
| Ayudante 1 Electrolitos | Sí/No | Si puede ser inspeccionado como Ayudante 1 Electrolitos |
| Operador | Sí/No | Si puede ser inspeccionado como Operador |
| Operador Electrolitos | Sí/No | Si puede ser inspeccionado como Operador Electrolitos |
| Técnico de producción - grado C | Sí/No | Si puede ser inspeccionado como Técnico grado C |
| Técnico de producción - grado D | Sí/No | Si puede ser inspeccionado como Técnico grado D |
| Muestreador | Sí/No | Si puede ser inspeccionado como Muestreador |
| Activo | Sí/No | Si está activo en sistema |

**Nota:** Una persona puede tener múltiples puestos simultáneamente (Sí en varias columnas).

---

### **7. tblPlantillas** (Hoja: Checklist)
Define los tipos de checklist por puesto y etapa.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| ID Plantilla | Texto (UUID) | PK - Identificador único |
| Nombre de plantilla | Texto | Nombre descriptivo |
| Etapa | Texto | Etapa del proceso (Producción/Lavado/Sanitizado/Todas) |
| Puesto | Texto | Puesto objetivo del checklist |
| Frecuencia meses | Número | Frecuencia en meses para re-inspección |

**Ejemplo de datos:**
```
ID Plantilla                | Nombre                          | Etapa       | Puesto                | Freq
───────────────────────────────────────────────────────────────────────────────────────────────────
fxEJV01C-xC6PKG6C-pVOj2dMa  | Operador NPT                    | Producción  | Operador              | [definir]
2hBWFdx9-eaZbiM3u-2KIZiYml  | Operador ONCO                   | Producción  | Operador              | [definir]
6pJsZ9vo-xDXzIcas-FNll3FM6  | Operador Electrolitos NPT       | Producción  | Operador              | [definir]
9yQXj5tE-P7VX9Ciq-iQXHPmbS  | Ayudante 1 NPT                  | Producción  | Ayudante 1            | [definir]
```

---

### **8. tblPreguntas** (Hoja: Checklist)
Preguntas específicas por plantilla (diseño 1:1, no reutilizables).

| Columna | Tipo | Descripción |
|---------|------|-------------|
| ID Plantilla | Texto (UUID) | FK → tblPlantillas.ID Plantilla |
| ID Pregunta | Texto (UUID) | PK - Identificador único |
| ID Sección | Texto (UUID) | FK → tblSecciones.ID Seccion |
| Numero | Número | Número correlativo de pregunta |
| Texto pregunta | Texto | Enunciado de la pregunta |
| ID Criticidad | Texto (UUID) | FK → tblCriticidad.ID Criticidad |
| Orden | Número | Orden de aparición en checklist |
| Activo | Sí/No | Si la pregunta está habilitada |
| Observaciones | Texto | Notas adicionales |
| Fecha creación | Fecha | Fecha de creación de la pregunta |

**Ejemplo de datos:**
```
Plantilla: fxEJV01C... (Operador NPT)
─────────────────────────────────────────────────────────────────────────────────────
ID Pregunta  | ID Sección | N° | Texto pregunta                               | Criticidad
RqL9nt8K...  | 6bzXQZjA.. | 1  | Realiza instalación de válvula sin romper... | Mayor
DOtmRIeT...  | 6bzXQZjA.. | 2  | Desempaca guía indicada en INLET...          | Mayor
K6IovBao...  | 6bzXQZjA.. | 3  | Instalan guía en puerto adecuado...          | Crítica
```

---

### **9. tblInspecciones** (Hoja: Historico) 🔥
Tabla transaccional - Cabecera de inspecciones ejecutadas.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| ID Inspeccion | Texto (UUID) | PK - Identificador único |
| Iniciales personal | Texto | FK → tblPersonal.Iniciales |
| ID Plantilla | Texto (UUID) | FK → tblPlantillas.ID Plantilla |
| Planta ejecutora | Texto | Planta donde se ejecutó (del personal al momento) |
| Fecha inspeccion | Fecha/Hora | Fecha y hora de la inspección |
| Auditor | Texto | Nombre/iniciales del auditor |
| Estado | Texto | "En progreso" / "Completado" / "Cancelado" |
| **--- SCORING TÉCNICA ASÉPTICA ---** |
| TA puntaje obtenido | Número decimal | Suma de valores de respuestas de sección TA |
| TA puntos maximos | Número decimal | Máximo posible (base 57 menos ajustes) |
| TA puntos no aplica | Número decimal | Puntos excluidos del denominador |
| TA porcentaje | Número decimal | Fórmula: (obtenido × 100) / (maximos - no aplica) |
| **--- RPN Y CATEGORIZACIÓN ---** |
| RPN calculado | Número decimal | RPN final calculado |
| Categoria resultado | Texto | "Categoría 1" a "Categoría 5" |
| Requiere accion | Sí/No | Si categoría ≥3 |
| **--- PROGRAMACIÓN ---** |
| Fecha proxima inspeccion | Fecha | Fecha inspeccion + Frecuencia meses |
| Dias para vencimiento | Número | Calculado: Fecha proxima - HOY() |
| Estado programacion | Texto | "Vigente" / "Próximo a vencer" / "Vencido" |
| **--- AUDITORÍA ---** |
| Observaciones generales | Texto largo | Comentarios finales del auditor |
| Fecha calculo | Fecha/Hora | Última vez que se ejecutó el cálculo |
| Usuario calculo | Texto | Usuario que ejecutó el cálculo |
| Fecha completado | Fecha/Hora | Cuándo se completó la inspección |
| Usuario completado | Texto | Usuario que completó |

---

### **10. tblRespuestas** (Hoja: Historico) 🔥
Tabla transaccional - Detalle de respuestas por inspección.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| ID Respuesta | Texto (UUID) | PK - Identificador único |
| ID Inspeccion | Texto (UUID) | FK → tblInspecciones.ID Inspeccion |
| ID Pregunta | Texto (UUID) | FK → tblPreguntas.ID Pregunta |
| ID Opcion | Texto (UUID) | FK → tblOpcionesDeRespuesta.ID Opcion (puede ser NULL) |
| Valor numerico | Número decimal | Copiado de tblOpcionesDeRespuesta.Valor puntaje |
| Observacion | Texto largo | Comentario específico del auditor |
| Fecha respuesta | Fecha/Hora | Timestamp de la respuesta |

---

### **11. tblCronogramaInspecciones** (Hoja: Cronograma) 🔥🆕
Tabla maestra sincronizada - Cronograma de todas las inspecciones programadas por Persona × Plantilla.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| ID Cronograma | Texto (UUID) | PK - Identificador único |
| Iniciales personal | Texto | FK → tblPersonal.Iniciales |
| ID Plantilla | Texto (UUID) | FK → tblPlantillas.ID Plantilla |
| Nombre plantilla | Texto | Copiado para visualización rápida |
| Puesto | Texto | Puesto del checklist |
| Planta personal | Texto | Planta del personal (de tblPersonal) |
| Frecuencia meses | Número | Frecuencia de reinspección (de tblPlantillas) |
| **--- ESTADO DE INSPECCIONES ---** |
| Total inspecciones | Número | Cantidad de inspecciones completadas |
| Fecha primera inspeccion | Fecha | Primera inspección registrada (MIN) |
| Fecha ultima inspeccion | Fecha | Última inspección completada (MAX) |
| ID Ultima inspeccion | Texto (UUID) | Para consulta rápida |
| RPN ultima inspeccion | Número decimal | RPN de última inspección |
| Categoria ultima inspeccion | Texto | Categoría de última inspección |
| **--- PROGRAMACIÓN ---** |
| Fecha proxima inspeccion | Fecha | Fecha ultima + Frecuencia meses |
| Dias para vencimiento | Número | Fecha proxima - HOY() |
| Estado cronograma | Texto | Ver estados posibles abajo |
| Dias alerta | Número | Días de anticipación para alertar |
| **--- VALIDACIÓN ---** |
| Puesto activo en personal | Sí/No | Si el puesto sigue "Si" en tblPersonal |
| Personal activo | Sí/No | Si tblPersonal.Activo = "Si" |
| Plantilla tiene preguntas | Sí/No | Validación de contenido |
| **--- AUDITORÍA ---** |
| Fecha ultima actualizacion | Fecha/Hora | Última sincronización |
| Requiere recalculo | Sí/No | Flag de optimización |

**Estados posibles de cronograma:**
- `Nunca inspeccionado` - No hay inspecciones completadas
- `Vigente` - Días para vencimiento > días de alerta
- `Por vencer` - 0 < Días para vencimiento ≤ días de alerta
- `Vencido` - Días para vencimiento < 0
- `Puesto inactivo` - El puesto cambió a "No" en tblPersonal (pintado especial)

**Regla de formato condicional:**
- 🔴 Rojo: Estado = "Vencido"
- 🟠 Naranja: Estado = "Por vencer"
- 🟢 Verde: Estado = "Vigente"
- ⚪ Gris: Estado = "Nunca inspeccionado"
- 🟣 Morado: Estado = "Puesto inactivo" (mantiene histórico pero marca inactividad)

---

## 🔗 DIAGRAMA DE RELACIONES

```
┌────────────────┐
│  tblPersonal   │
│  (Iniciales)   │
└────────┬───────┘
         │ 1:N
         ▼
┌──────────────────────┐       ┌─────────────────┐
│  tblInspecciones     │◄──N:1─┤ tblPlantillas   │
│  (ID Inspeccion)     │       │ (ID Plantilla)  │
└──────────┬───────────┘       └────────┬────────┘
           │ 1:N                        │ 1:N
           ▼                            ▼
┌──────────────────────┐       ┌─────────────────┐
│  tblRespuestas       │──N:1─►│ tblPreguntas    │
│  (ID Respuesta)      │       │ (ID Pregunta)   │
└──────────┬───────────┘       └────────┬────────┘
           │ N:1                        │ N:1
           │                            │
           │                   ┌────────┴────────┐
           │                   │                 │
           │                   ▼                 ▼
           │          ┌─────────────────┐ ┌──────────────┐
           │          │  tblSecciones   │ │tblCriticidad │
           │          │  (ID Seccion)   │ │(ID Critic.)  │
           │          └────────┬────────┘ └──────────────┘
           │                   │ 1:N
           │                   ▼
           │          ┌─────────────────────────┐
           └─────N:1─►│ tblOpcionesDeRespuesta  │
                      │  (ID Opcion)            │
                      └─────────────────────────┘

┌──────────────────────┐
│ tblCategoriasRPN     │◄──── Usado para categorizar resultado
│ (Numero categoria)   │
└──────────────────────┘

┌──────────────────────┐
│ tblConfiguracion     │◄──── Parámetros globales del sistema
│ (Clave)              │
└──────────────────────┘
```

---

## 🧮 LÓGICA DE CÁLCULOS

### **1. SCORING DE TÉCNICA ASÉPTICA (TA)**

**Fórmula:**
```
TA porcentaje = (TA puntaje obtenido × 100) / (TA puntos maximos - TA puntos no aplica)

Donde:
- TA puntaje obtenido = Σ(Valor numerico de respuestas de sección TA)
- TA puntos maximos = 57 (base configurado en tblConfiguracion)
- TA puntos no aplica = Σ(Valor de opciones "No Aplica" seleccionadas)
```

**Ejemplo:**
```
Checklist con 20 preguntas de TA:
- 15 respondidas "Si" (valor 0) = 15 × 0 = 0 puntos
- 3 respondidas "No" (valor 4) = 3 × 4 = 12 puntos
- 2 respondidas "No Aplica" (valor -1) = 2 × (-1) = -2 puntos

TA puntaje obtenido = 0 + 12 + (-2) = 10
TA puntos maximos = 57
TA puntos no aplica = 2 × 4 = 8 (se usa el valor de "No", no el -1)

TA porcentaje = (10 × 100) / (57 - 8) = 1000 / 49 = 20.41%
```

**Regla especial:** Cuando se selecciona "No Aplica", el denominador se reduce usando el valor que tendría la opción "No" de esa sección.

---

### **2. CÁLCULO DE RPN**

El RPN se calcula utilizando el TA porcentaje:

```
RPN calculado = TA porcentaje (calculado en paso anterior)
```

**Nota:** Para versiones futuras, el cliente puede definir fórmulas adicionales que involucren otras secciones.

---

### **3. CATEGORIZACIÓN DE RESULTADO**

**Proceso:**

1. Calcular RPN de la inspección actual
2. Determinar categoría base según rangos de tblCategoriasRPN
3. **Validación especial Categoría 5:**
   - Si RPN > 20, buscar las últimas 2 inspecciones del mismo Personal + Plantilla
   - Si las 3 inspecciones consecutivas tienen RPN > 20 → asignar Categoría 5
   - Si no cumple condición → asignar categoría según rango normal

**Pseudocódigo:**
```vba
Function DeterminarCategoria(RPN As Double, Iniciales As String, IDPlantilla As String) As String
    ' Buscar últimas 2 inspecciones anteriores
    Dim inspecciones As Collection
    Set inspecciones = ObtenerUltimas2Inspecciones(Iniciales, IDPlantilla)
    
    ' Validar Categoría 5
    If RPN > 20 And inspecciones.Count >= 2 Then
        If inspecciones(1).RPN > 20 And inspecciones(2).RPN > 20 Then
            Return "Categoría 5"
        End If
    End If
    
    ' Categorización normal por rangos
    If RPN >= 0 And RPN <= 14 Then Return "Categoría 1"
    If RPN >= 15 And RPN <= 19 Then Return "Categoría 2"
    If RPN >= 20 And RPN <= 40 Then Return "Categoría 3"
    If RPN > 40 Then Return "Categoría 4"
End Function
```

---

### **4. PROGRAMACIÓN DE PRÓXIMA INSPECCIÓN**

**Fórmula:**
```
Fecha proxima inspeccion = Fecha inspeccion + (Frecuencia meses × 30.44 días)

Dias para vencimiento = Fecha proxima inspeccion - HOY()

Estado programacion = 
    Si Dias para vencimiento > ALERTAR_DIAS_ANTES_VENCIMIENTO → "Vigente"
    Si 0 < Dias para vencimiento ≤ ALERTAR_DIAS_ANTES_VENCIMIENTO → "Próximo a vencer"
    Si Dias para vencimiento ≤ 0 → "Vencido"
```

**Notas:**
- La frecuencia se obtiene de tblPlantillas.Frecuencia meses
- El parámetro ALERTAR_DIAS_ANTES_VENCIMIENTO está en tblConfiguracion (sugerido: 15 días)

---

## ⚙️ REGLAS DE NEGOCIO

### **Validaciones para crear inspección:**

1. **Personal activo:** Verificar que `tblPersonal.Activo = "Si"`
2. **Puesto válido:** Verificar que el puesto de la plantilla coincida con algún puesto activo del personal
   - Ejemplo: Si plantilla es "Operador NPT", verificar que `tblPersonal.Operador = "Si"`
3. **Plantilla activa:** Verificar que existan preguntas activas en tblPreguntas para esa plantilla
4. **No duplicados:** Opcional - verificar que no exista una inspección "En progreso" para mismo Personal+Plantilla

### **Validaciones durante inspección:**

1. **Respuestas obligatorias:** Todas las preguntas con `Activo = "Si"` deben tener respuesta
2. **Opciones válidas:** La opción seleccionada debe pertenecer a la sección de la pregunta
3. **Guardar progreso:** Permitir guardar respuestas parciales (Estado = "En progreso")

### **Validaciones para completar inspección:**

1. **Todas las preguntas respondidas:** Verificar que no haya preguntas activas sin respuesta
2. **Cálculo ejecutado:** Ejecutar el proceso de cálculo automáticamente al completar
3. **Cambiar estado:** De "En progreso" a "Completado"
4. **Timestamp:** Registrar fecha y usuario de completado

### **Validaciones para recalcular:**

1. **Solo si está completada:** No permitir recalcular inspecciones en progreso
2. **Actualizar timestamp:** Actualizar Fecha calculo y Usuario calculo
3. **Mantener respuestas:** El recálculo NO modifica respuestas, solo resultados

---

## 🔐 INTEGRIDAD REFERENCIAL

### **Relaciones obligatorias:**

- tblRespuestas.ID Inspeccion → debe existir en tblInspecciones
- tblRespuestas.ID Pregunta → debe existir en tblPreguntas
- tblRespuestas.ID Opcion → debe existir en tblOpcionesDeRespuesta (excepto RPN manual)
- tblPreguntas.ID Plantilla → debe existir en tblPlantillas
- tblPreguntas.ID Seccion → debe existir en tblSecciones
- tblPreguntas.ID Criticidad → debe existir en tblCriticidad
- tblInspecciones.Iniciales personal → debe existir en tblPersonal.Iniciales
- tblInspecciones.ID Plantilla → debe existir en tblPlantillas

### **Eliminaciones en cascada:**

- Al eliminar tblInspecciones → eliminar todas las tblRespuestas asociadas
- **NO PERMITIR** eliminar tblPlantillas, tblSecciones, tblOpcionesDeRespuesta, tblCriticidad si tienen dependencias

---

## 📊 CONSULTAS PRINCIPALES

### **Dashboard de Inspecciones Pendientes:**
```sql
SELECT 
    p.Iniciales,
    p.Planta,
    pl.Nombre de plantilla,
    MAX(i.Fecha inspeccion) AS Ultima_Inspeccion,
    MAX(i.Fecha proxima inspeccion) AS Proxima_Inspeccion,
    MAX(i.Dias para vencimiento) AS Dias_Restantes,
    MAX(i.Estado programacion) AS Estado
FROM tblPersonal p
CROSS JOIN tblPlantillas pl
LEFT JOIN tblInspecciones i ON p.Iniciales = i.Iniciales personal 
    AND pl.ID Plantilla = i.ID Plantilla
    AND i.Estado = 'Completado'
WHERE p.Activo = 'Si'
    AND [validar que puesto de plantilla = Si en tblPersonal]
GROUP BY p.Iniciales, p.Planta, pl.Nombre de plantilla
ORDER BY Dias_Restantes ASC
```

### **Historial de Inspecciones por Persona:**
```sql
SELECT 
    i.Fecha inspeccion,
    pl.Nombre de plantilla,
    i.Auditor,
    i.RPN calculado,
    i.Categoria resultado,
    i.Estado,
    i.TA porcentaje
FROM tblInspecciones i
INNER JOIN tblPlantillas pl ON i.ID Plantilla = pl.ID Plantilla
WHERE i.Iniciales personal = [Iniciales]
ORDER BY i.Fecha inspeccion DESC
```

### **Últimas 3 Inspecciones (para Categoría 5):**
```sql
SELECT TOP 3
    Fecha inspeccion,
    RPN calculado
FROM tblInspecciones
WHERE Iniciales personal = [Iniciales]
    AND ID Plantilla = [ID Plantilla]
    AND Estado = 'Completado'
ORDER BY Fecha inspeccion DESC
```

---

## 🎯 MÓDULOS VBA A DESARROLLAR

### **1. InspectionScheduler.bas** (programación) ✅ COMPLETADO
**Estado:** Implementado el 12/03/2026  
**Ubicación:** InspectionScheduler.bas

**Funciones públicas:**
- `InicializarCronograma()` - Crea cronograma desde cero (Persona × Plantilla)
- `RecalcularCronograma()` - Actualiza todo el cronograma con última info
- `ActualizarRegistroCronograma(iniciales, idPlantilla)` - Actualiza un registro específico

**Funciones privadas:**
- `ActualizarRegistroCronogramaInterno()` - Lógica de actualización compartida
- `ValidarPuestoActivo()` - Verifica si puesto sigue activo en tblPersonal
- `ObtenerValorPersonal()` - Obtiene valor de columna de tblPersonal
- `ObtenerParametroNumerico()` - Lee parámetros de tblConfiguracion
- `CrearRegistroCronograma()` - Crea registro si no existe
- `GenerarUUID()` - Genera identificadores únicos
- `GenerarCadenaAleatoria()` - Helper para UUID

**Integración con Configuration2:**
- Usa constantes centralizadas para nombres de hojas/tablas
- Usa estados de cronograma e inspección definidos
- Lee parámetros dinámicos de tblConfiguracion

**Características especiales:**
- Maneja estructura peculiar de tblPersonal (puestos en columnas)
- Audita operaciones con AuditLogger2
- Registra errores con ErrorLogger2
- Optimizado para performance (Application.ScreenUpdating/Calculation)
- Detecta puestos inactivos y marca con estado especial

---

### **2. InspectionCore.bas** (módulo principal) 📋 PENDIENTE
- CrearInspeccion()
- AbrirInspeccion()
- GuardarRespuesta()
- CompletarInspeccion() → debe llamar ActualizarRegistroCronograma()
- CancelarInspeccion()

### **3. InspectionCalculator.bas** (cálculos) 📋 PENDIENTE
- CalcularScoringTA()
- CalcularRPN()
- DeterminarCategoria()
- RecalcularInspeccion()
- ObtenerUltimas2Inspecciones()

### **4. InspectionValidation.bas** (validaciones) 📋 PENDIENTE
- ValidarPuestoPersonal()
- ValidarPersonalActivo()
- ValidarTodasRespuestasCompletas()
- ValidarOpcionValida()

### **5. InspectionUI.bas** (interfaz) 📋 PENDIENTE
- MostrarFormularioInspeccion()
- CargarPreguntas()
- MostrarDashboard()
- GenerarReporte()

---

## 📝 NOTAS IMPORTANTES

1. **UUIDs:** Todos los IDs usan formato UUID con guiones (ej: `fxEJV01C-xC6PKG6C-pVOj2dMa`)

2. **Fechas:** Usar formato Date/Time de Excel, zona horaria local

3. **Estado "En progreso":** Permite guardar respuestas parciales y continuar después

4. **Cálculos bajo demanda:** Los campos calculados (RPN, categoría, scoring TA) se recalculan:
   - Automáticamente al completar inspección
   - Manualmente con botón "Recalcular"
   - No se recalculan en tiempo real durante respuesta

5. **Histórico inmutable:** Una vez completada, no se pueden modificar respuestas (solo recalcular con nueva lógica)

6. **Múltiples puestos:** Una misma persona puede ser inspeccionada con diferentes plantillas según sus puestos activos

7. **Frecuencia por plantilla:** Cada plantilla define su propia frecuencia de re-inspección

8. **No Aplica especial:** En sección TA, el valor -1 de "No Aplica" se trata especialmente:
   - Se suma al puntaje obtenido como -1
   - Se resta del denominador con el valor de "No" (4)

---

## 🚀 ROADMAP DE IMPLEMENTACIÓN

### **Fase 1: Estructura de datos** ✅
- [x] Definir tablas
- [x] Crear ListObjects en Excel
- [x] Documentar arquitectura
- [x] Agregar constantes a Configuration2
- [x] Crear tblCronogramaInspecciones
- [x] Implementar InspectionScheduler.bas

### **Fase 2: Módulos core** 
- [x] InspectionScheduler.bas (sincronización de cronograma) ✅
- [ ] InspectionCore.bas (CRUD básico)
- [ ] InspectionCalculator.bas (cálculos)
- [ ] InspectionValidation.bas (validaciones)

### **Fase 3: UI y formularios**
- [ ] Formulario de inspección
- [ ] Dashboard de programación
- [ ] Reportes

### **Fase 4: Testing y ajustes**
- [ ] Pruebas con datos reales
- [ ] Validación de fórmulas
- [ ] Ajuste de categorización

---

## 🔧 GUÍA DE USO - MÓDULO InspectionScheduler

### **Configuración inicial (primera vez):**

**1. Verificar que existen las tablas:**
- ✅ tblPersonal (hoja Personal) con columnas de puestos
- ✅ tblPlantillas (hoja Checklist) con columna "Frecuencia meses"
- ✅ tblInspecciones (hoja Historico) - puede estar vacía
- ✅ tblCronogramaInspecciones (hoja Cronograma) - debe estar vacía
- ✅ tblConfiguracion (hoja Configuración) con parámetros

**2. Configurar parámetros en tblConfiguracion:**
```
Clave: ALERTAR_DIAS_ANTES_VENCIMIENTO
Valor: 15 (o días que prefieras)
Categoria: Notificaciones
Descripcion: Días de anticipación para alertar vencimiento
Tipo dato: Numero
```

**3. Ejecutar inicialización:**
- Ir a hoja "Cronograma"
- Hacer clic en botón "INICIALIZAR CRONOGRAMA" (vincular a `InspectionScheduler.InicializarCronograma`)
- Esperar mensaje de confirmación con cantidad de registros creados

**Ejemplo de mensaje:**
```
Cronograma inicializado exitosamente.

Registros creados: 87
Tiempo: 1.23 segundos
```

---

### **Uso normal (operación diaria):**

**Botón "RECALCULAR CRONOGRAMA":**
- Ejecuta: `InspectionScheduler.RecalcularCronograma()`
- Cuándo usar:
  - Al inicio del día para ver estados actualizados
  - Después de completar múltiples inspecciones
  - Si hay cambios en tblPersonal (puestos activados/desactivados)
- Actualiza: Todos los registros del cronograma

**Botón "NUEVA INSPECCIÓN":**
- Ejecuta: (próximo módulo - InspectionCore.CrearInspeccion)
- Cuándo usar: Para iniciar una nueva inspección
- Automáticamente llamará a `ActualizarRegistroCronograma()` al completar

---

### **Integración con otros módulos:**

**Al completar una inspección (desde InspectionCore):**
```vba
Sub CompletarInspeccion(IDInspeccion As String)
    ' ... código de completar inspección ...
    
    ' Actualizar cronograma automáticamente
    Dim iniciales As String
    Dim idPlantilla As String
    
    ' Obtener datos de la inspección
    iniciales = [obtener de tblInspecciones]
    idPlantilla = [obtener de tblInspecciones]
    
    ' Sincronizar cronograma
    Call InspectionScheduler.ActualizarRegistroCronograma(iniciales, idPlantilla)
End Sub
```

---

### **Reglas de formato condicional (aplicar en hoja Cronograma):**

**Columna "Estado cronograma":**
```vba
' Formato rojo (vencido)
= $[Columna Estado] = "Vencido"
Color de relleno: RGB(255, 199, 206)
Color de texto: RGB(156, 0, 6)

' Formato naranja (por vencer)
= $[Columna Estado] = "Por vencer"
Color de relleno: RGB(255, 235, 156)
Color de texto: RGB(156, 101, 0)

' Formato verde (vigente)
= $[Columna Estado] = "Vigente"
Color de relleno: RGB(198, 239, 206)
Color de texto: RGB(0, 97, 0)

' Formato gris (nunca inspeccionado)
= $[Columna Estado] = "Nunca inspeccionado"
Color de relleno: RGB(217, 217, 217)
Color de texto: RGB(89, 89, 89)

' Formato morado (puesto inactivo) - REGLA ESPECIAL
= $[Columna Estado] = "Puesto inactivo"
Color de relleno: RGB(230, 204, 255)
Color de texto: RGB(112, 48, 160)
```

**Columna "Puesto activo en personal":**
```vba
' Resaltar "No" en amarillo
= $[Columna Puesto activo] = "No"
Color de relleno: RGB(255, 242, 204)
```

---

### **Casos especiales:**

**1. Personal cambia de puesto (ej: Operador → Ayudante 1):**
- En tblPersonal: Cambiar columna "Operador" de "Si" a "No"
- En tblPersonal: Cambiar columna "Ayudante 1" de "No" a "Si"
- Ejecutar "RECALCULAR CRONOGRAMA"
- Resultado:
  - Registro de Operador queda con estado "Puesto inactivo" (histórico preservado)
  - Si no existe cronograma para Ayudante 1, crear manualmente o ejecutar "INICIALIZAR CRONOGRAMA"

**2. Nueva plantilla agregada:**
- Agregar plantilla a tblPlantillas
- Ejecutar "INICIALIZAR CRONOGRAMA" (solo agregará los faltantes, no duplica)

**3. Inspección antigua completada retroactivamente:**
- Completar inspección con fecha pasada en tblInspecciones
- Ejecutar "RECALCULAR CRONOGRAMA" o actualizar registro específico

---

### **Optimización y performance:**

**Para 50 personas con 5 puestos promedio = 250 registros de cronograma:**
- Inicialización: ~1-3 segundos
- Recálculo completo: ~2-5 segundos
- Actualización individual: <0.1 segundos

**Campo "Requiere recalculo":**
- Se marca "Si" cuando se detectan cambios pendientes
- Se marca "No" después de actualizar
- Usar para optimizar recálculos futuros (filtrar solo registros con "Si")

---

**FIN DEL DOCUMENTO**  
Última revisión: 12/03/2026
