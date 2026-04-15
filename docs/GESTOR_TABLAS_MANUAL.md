# 📚 Manual del Gestor de Tablas Maestras

**Sistema de Inspecciones - TH-HC-001**  
**Versión:** 1.0  
**Fecha:** Abril 2026

---

## 📋 Tabla de Contenidos

1. [Introducción](#introducción)
2. [Acceso al Gestor](#acceso-al-gestor)
3. [Descripción de las 5 Tablas](#descripción-de-las-5-tablas)
4. [Guía Paso a Paso: Crear Elementos](#guía-paso-a-paso-crear-elementos)
5. [Validaciones y Reglas](#validaciones-y-reglas)
6. [Operaciones CRUD](#operaciones-crud)
7. [Troubleshooting](#troubleshooting)

---

## Introducción

El **Gestor de Tablas Maestras** es una interfaz única y centralizada para administrar las 5 tablas fundamentales del sistema de inspecciones:

- ✅ **CRITICIDAD** - Niveles de riesgo (Alto, Medio, Bajo)
- ✅ **SECCIONES** - Áreas de inspección (Auditoría, Seguridad, Calidad, etc.)
- ✅ **PLANTILLAS** - Modelos de inspección con etapas y frecuencias
- ✅ **OPCIONES** - Respuestas predefinidas (Si, No, N/A, etc.)
- ✅ **PREGUNTAS** - Ítems específicos de cada plantilla

### ¿Por qué es importante?

Todas las tablas están **interconectadas**. Cambios en una tabla afectan a las demás. Por eso:
- El sistema **valida dependencias** antes de eliminar datos
- Los **IDs se protegen** automáticamente
- Se **registra todo** en el audit trail para auditoría

---

## Acceso al Gestor

### 1. Desde el Menú Principal

1. Abre el archivo: `TH-HC-001 INSPECCIONES.xlsm`
2. Ve a la hoja: **"Menú Principal"**
3. Busca el botón: **"🔧 Gestor de Tablas"**
4. Haz clic para abrir el formulario

### 2. Interfaz del Gestor

```
┌─────────────────────────────────────────────────────────┐
│          GESTOR DE TABLAS MAESTRAS                      │
├─────────────────────────────────────────────────────────┤
│ Tabla: [Seleccione tabla...▼]                           │
├─────────────────────────────────────────────────────────┤
│ ID │ Nombre         │ Valor │ Descripción               │
│ ─────────────────────────────────────────────────────── │
│ 1  │ Alto Riesgo    │ 10    │ Impacto crítico...       │
│ 2  │ Riesgo Medio   │ 5     │ Impacto moderado...      │
│ 3  │ Bajo Riesgo    │ 1     │ Impacto mínimo...        │
│ ─────────────────────────────────────────────────────── │
├─────────────────────────────────────────────────────────┤
│ ID: ________          Nombre: __________________        │
│ Valor: ______         Descripción: ______________       │
├─────────────────────────────────────────────────────────┤
│ [Nuevo] [Guardar] [Eliminar]  [Validar Todo] [Cerrar] │
│ Estado: Seleccione una tabla para comenzar...          │
└─────────────────────────────────────────────────────────┘
```

**Componentes principales:**
- **Selector de Tabla:** Dropdown para elegir qué tabla editar
- **Lista de Datos:** Muestra todos los registros existentes
- **Área de Edición:** Campos para crear/editar registros
- **Botones CRUD:** Nuevo, Guardar, Eliminar
- **Barra de Estado:** Mensajes informativos y de error

---

## Descripción de las 5 Tablas

### 1️⃣ CRITICIDAD (Niveles de Riesgo)

**Ubicación:** Hoja "Checklist"  
**Propósito:** Define los niveles de severidad para evaluaciones

| Campo | Tipo | Requerido | Restricciones |
|-------|------|-----------|---------------|
| **ID** | Texto | Sí ✓ | Único, generado automáticamente |
| **Nombre de criticidad** | Texto | Sí ✓ | Máx 100 caracteres, único |
| **Valor de criticidad** | Número | No | Positivo, recomendado 1-10 |

**Ejemplos de Criticidad:**
```
- Alto (10)
- Crítico (9)
- Medio-Alto (7)
- Medio (5)
- Bajo (2)
- Mínimo (1)
```

**Reglas:**
- ✓ Mínimo 1 criticidad
- ✓ Nombres únicos (no puede haber dos "Alto")
- ✓ Los valores deben ser números positivos
- ⚠️ No puede eliminarse si hay PREGUNTAS que la usan

---

### 2️⃣ SECCIONES (Áreas de Auditoría)

**Ubicación:** Hoja "Checklist"  
**Propósito:** Agrupa temas dentro de un proceso de auditoría

| Campo | Tipo | Requerido | Restricciones |
|-------|------|-----------|---------------|
| **ID** | Texto | Sí ✓ | Único, generado automáticamente |
| **Nombre de sección** | Texto | Sí ✓ | Máx 150 caracteres, único |
| **Tipo de respuesta** | Texto | Sí ✓ | "Selección" o "Puntaje" |

**Ejemplos de Sección:**
```
🔹 Auditoría de procesos (Selección)
🔹 Seguridad e higiene (Puntaje)
🔹 Cumplimiento regulatorio (Selección)
🔹 Gestión de recursos (Puntaje)
🔹 Documentación (Selección)
```

**Tipo de Respuesta:**
- **Selección:** El auditor elige de opciones predefinidas (Sí/No/Parcial)
- **Puntaje:** Se ingresa puntuación numérica

**Reglas:**
- ✓ Mínimo 1 sección
- ✓ Nombres únicos
- ✓ Tipo debe ser exactamente "Selección" o "Puntaje"
- ⚠️ No puede eliminarse si hay OPCIONES o PREGUNTAS que la usan

---

### 3️⃣ PLANTILLAS (Modelos de Inspección)

**Ubicación:** Hoja "Checklist"  
**Propósito:** Define estructuras reutilizables de auditoría con etapas y frecuencia

| Campo | Tipo | Requerido | Restricciones |
|-------|------|-----------|---------------|
| **ID** | Texto | Sí ✓ | Único, generado automáticamente |
| **Nombre de plantilla** | Texto | Sí ✓ | Máx 150 caracteres, único |
| **Etapa** | Texto | No | Ej: "Inicial", "Seguimiento", "Final" |
| **Puesto responsable** | Texto | No | Ej: "Auditor Senior", "Inspector" |
| **Frecuencia (días)** | Número | No | Positivo, ej: 30, 90, 365 |

**Ejemplos de Plantilla:**
```
1. Auditoría Inicial
   - Etapa: Inicial
   - Puesto: Auditor Senior
   - Frecuencia: 1

2. Auditoría de Seguimiento
   - Etapa: Seguimiento
   - Puesto: Auditor Junior
   - Frecuencia: 90

3. Auditoría Anual
   - Etapa: Final
   - Puesto: Auditor Senior
   - Frecuencia: 365
```

**Reglas:**
- ✓ Nombres únicos
- ✓ Frecuencia debe ser número positivo
- ✓ La frecuencia se usa para programar auditorías automáticas
- ⚠️ No puede eliminarse si hay PREGUNTAS que la usan

---

### 4️⃣ OPCIONES (Respuestas Predefinidas)

**Ubicación:** Hoja "Checklist"  
**Propósito:** Define respuestas disponibles para cada sección (operacionaliza Secciones)

| Campo | Tipo | Requerido | Restricciones |
|-------|------|-----------|---------------|
| **ID** | Texto | Sí ✓ | Único, generado automáticamente |
| **Sección** | Referencia | Sí ✓ | Debe existir en SECCIONES |
| **Criticidad** | Referencia | No | Debe existir en CRITICIDAD |
| **Texto de opción** | Texto | Sí ✓ | Máx 200 caracteres, única por sección |
| **Valor de puntaje** | Número | No | Ej: 0, 1, 2, 5 (según tipo sección) |

**Red de Dependencias:**
```
SECCIONES
    ↓
OPCIONES (depende de SECCIONES)
    ↓
PREGUNTAS (puede usar OPCIONES)
```

**Ejemplos de OPCIONES:**

Para Sección "Auditoría de procesos" (tipo Selección):
```
- Opción: Conforme
  Criticidad: (vacío)
  Valor: 0

- Opción: No conforme
  Criticidad: Alto
  Valor: -10

- Opción: Parcialmente conforme
  Criticidad: Medio
  Valor: -5
```

Para Sección "Seguridad" (tipo Puntaje):
```
- Opción: Excelente
  Criticidad: (vacío)
  Valor: 5

- Opción: Bueno
  Criticidad: (vacío)
  Valor: 3

- Opción: Deficiente
  Criticidad: Crítico
  Valor: 1
```

**Reglas:**
- ✓ La SECCIÓN debe existir
- ✓ Si especifica CRITICIDAD, debe existir
- ✓ Textos únicos dentro de la misma sección
- ✓ Los valores deben ser números
- ⚠️ No puede eliminarse si hay PREGUNTAS que la referencian

---

### 5️⃣ PREGUNTAS (Ítems de Auditoría)

**Ubicación:** Hoja "Checklist"  
**Propósito:** Preguntas específicas de cada plantilla + sección + criticidad

| Campo | Tipo | Requerido | Restricciones |
|-------|------|-----------|---------------|
| **ID** | Texto | Sí ✓ | Único, generado automáticamente |
| **Plantilla** | Referencia | Sí ✓ | Debe existir en PLANTILLAS |
| **Sección** | Referencia | Sí ✓ | Debe existir en SECCIONES |
| **Criticidad** | Referencia | No | Debe existir en CRITICIDAD |
| **Texto de pregunta** | Texto | Sí ✓ | Máx 500 caracteres |
| **Orden** | Número | No | Posición dentro de la sección (1, 2, 3...) |
| **Activo** | Verdadero/Falso | No | Permite desactivar sin eliminar |
| **Observaciones** | Texto | No | Guías para el auditor |
| **Fecha de creación** | Fecha | No | Se registra automáticamente |

**Red de Dependencias:**
```
PLANTILLAS ────┐
               ├─→ PREGUNTAS
SECCIONES ─────┤
               ├─→ (puede ser) CRITICIDAD
CRITICIDAD ────┘
```

**Ejemplos de PREGUNTA:**

```
Pregunta 1:
- Plantilla: Auditoría Inicial
- Sección: Auditoría de procesos
- Criticidad: Alto
- Texto: ¿Existen procedimientos documentados para todos los procesos críticos?
- Orden: 1
- Activo: Sí
- Observaciones: Verificar que los procedimientos estén vigentes y disponibles

Pregunta 2:
- Plantilla: Auditoría Inicial
- Sección: Seguridad e Higiene
- Criticidad: Crítico
- Texto: ¿Se han completado todas las capacitaciones de seguridad?
- Orden: 1
- Activo: Sí
- Observaciones: Revisar registros de asistencia de los últimos 6 meses
```

**Reglas:**
- ✓ PLANTILLA debe existir
- ✓ SECCIÓN debe existir
- ✓ Si especifica CRITICIDAD, debe existir
- ✓ ORDEN debe ser número positivo
- ✓ Puede desactivarse sin eliminar (marcar Activo = No)
- ⚠️ Es la tabla final, puede eliminarse sin dependencias

---

## Guía Paso a Paso: Crear Elementos

### 🔴 PASO 1: Crear una CRITICIDAD

1. **Abre el Gestor** → Haz clic en "Gestor de Tablas" desde Menú Principal

2. **Selecciona tabla:**
   ```
   Tabla: [1. Criticidades ▼]
   ```

3. **Haz clic en [Nuevo]**
   - El sistema genera automáticamente un ID único
   - Los campos quedan listos para editar

4. **Completa los campos:**
   ```
   ID: (automático, no editable)
   Nombre de criticidad: Alto Riesgo
   Valor de criticidad: 10
   ```

5. **Validaciones automáticas:**
   - ✓ Nombre: obligatorio, máx 100 caracteres
   - ✓ Valor: debe ser número
   - ✓ Unicidad: no puede haber otro "Alto Riesgo"

6. **Haz clic en [Guardar]**
   - Si hay errores, aparece el mensaje de validación
   - Si es exitoso: "Registro guardado exitosamente"

7. **Resultado:**
   - LA CRITICIDAD aparece en la lista
   - Se registra en el AUDIT TRAIL automáticamente
   - Está disponible para usar en OPCIONES y PREGUNTAS

---

### 🟠 PASO 2: Crear una SECCIÓN

1. **Abre el Gestor** → Selecciona tabla

2. **Cambia a secciones:**
   ```
   Tabla: [2. Secciones ▼]
   ```

3. **Haz clic en [Nuevo]**

4. **Completa:**
   ```
   ID: (automático)
   Nombre de sección: Auditoría de procesos
   Tipo de respuesta: Selección  (o "Puntaje")
   ```

5. **Validaciones:**
   - ✓ Nombre: obligatorio, único
   - ✓ Tipo: debe ser exactamente "Selección" O "Puntaje"

6. **Haz clic en [Guardar]**

7. **Resultado:**
   - SECCIÓN disponible para crear OPCIONES
   - Se puede asignar a PREGUNTAS

---

### 🟡 PASO 3: Crear una PLANTILLA

1. **Tabla:** `[3. Plantillas ▼]`

2. **[Nuevo]**

3. **Completa:**
   ```
   ID: (automático)
   Nombre de plantilla: Auditoría Integral Q1
   Etapa: Inicial
   Puesto responsable: Auditor Senior
   Frecuencia (días): 90
   ```

4. **Validaciones:**
   - ✓ Nombre: obligatorio, único
   - ✓ Frecuencia: número positivo

5. **[Guardar]**

6. **Resultado:**
   - PLANTILLA lista para asignar PREGUNTAS

---

### 🟢 PASO 4: Crear OPCIONES (Respuestas)

⚠️ **IMPORTANTE:** Las OPCIONES dependen de SECCIONES, así que:
1. **Primero** crea la SECCIÓN ("Auditoría de procesos")
2. **Luego** crea las OPCIONES para esa sección

**Procedimiento:**

1. **Tabla:** `[4. Opciones de Respuesta ▼]`

2. **[Nuevo]**

3. **Completa:**
   ```
   ID: (automático)
   Sección: [Auditoría de procesos ▼]  ← Selecciona de la lista
   Criticidad: [Alto Riesgo ▼]         ← O déjalo vacío
   Texto de opción: No conforme
   Valor de puntaje: -10
   ```

4. **Validaciones:**
   - ✓ Sección: **OBLIGATORIA** (debe existir)
   - ✓ Criticidad: si se ingresa, debe existir
   - ✓ Texto: único dentro de esa sección
   - ✓ Valor: número

5. **[Guardar]**

6. **Repeat:** Crea más opciones para la misma sección
   ```
   - Conforme (Valor: 0)
   - Parcialmente conforme (Valor: -5)
   - No conforme (Valor: -10)
   ```

7. **Resultado:**
   - OPCIONES están disponibles para PREGUNTAS
   - El auditor verá estas opciones al responder

---

### 🔵 PASO 5: Crear PREGUNTAS (Ítems de Auditoría)

⚠️ **IMPORTANTE:** Las PREGUNTAS dependen de:
1. ✓ PLANTILLA (debe existir)
2. ✓ SECCIÓN (debe existir)
3. ✓ CRITICIDAD (opcional pero si se especifica, debe existir)

**Procedimiento:**

1. **Tabla:** `[5. Preguntas ▼]`

2. **[Nuevo]**

3. **Completa TODOS los datos:**
   ```
   ID: (automático)
   Plantilla: [Auditoría Integral Q1 ▼]
   Sección: [Auditoría de procesos ▼]
   Criticidad: [Alto Riesgo ▼]  (opcional)
   Texto de pregunta: ¿Existen procedimientos documentados para todos los procesos?
   Orden: 1
   Activo: ☑ (checked)
   Observaciones: Verificar vigencia en últimos 6 meses
   ```

4. **Validaciones:**
   - ✓ Plantilla: **OBLIGATORIA** (debe existir en PLANTILLAS)
   - ✓ Sección: **OBLIGATORIA** (debe existir en SECCIONES)
   - ✓ Criticidad: si se ingresa, debe existir
   - ✓ Texto: obligatorio, máx 500 caracteres
   - ✓ Orden: número positivo
   - ✓ Activo: Sí o No

5. **[Guardar]**

6. **Repeat:** Crea más preguntas
   ```
   Pregunta 2 (Orden: 2)
   Pregunta 3 (Orden: 3)
   etc.
   ```

7. **Resultado:**
   - PREGUNTAS aparecen en las auditorías
   - Los auditores verán estas preguntas en orden
   - Se pueden marcar activas/inactivas

---

## Validaciones y Reglas

### 📋 Matriz de Dependencias

```
CRITICIDAD                   SECCIONES
(Independiente)              (Independiente)
        │                           │
        │                           │
        └─────────────┬─────────────┘
                      ↓
              PLANTILLAS
              (Independiente)
              
    ┌──────────────┬──────────────┐
    ↓              ↓              ↓
OPCIONES      PREGUNTAS      (otros)
(depende de)  (depende de)
                │
        ┌───────┼───────┬────────┐
        ↓       ↓       ↓        ↓
    PLANTILLA SECCIÓN CRITICIDAD (OPCIONES)
```

### ✅ Validaciones por Tabla

#### CRITICIDAD
- ✓ Nombre: obligatorio, único, máx 100 caracteres
- ✓ Valor: número positivo (recomendado 1-10)
- ✓ No puede eliminarse si hay preguntas que la usan

#### SECCIONES
- ✓ Nombre: obligatorio, único, máx 150 caracteres
- ✓ Tipo: "Selección" o "Puntaje" (exacto)
- ✓ No puede eliminarse si hay opciones o preguntas que la usan

#### PLANTILLAS
- ✓ Nombre: obligatorio, único, máx 150 caracteres
- ✓ Frecuencia: número positivo (recomendado: 1, 30, 90, 365)
- ✓ Etapa y Puesto: campos de texto libre
- ✓ No puede eliminarse si hay preguntas que la usan

#### OPCIONES
- ✓ Sección: **OBLIGATORIA** (debe existir en SECCIONES)
- ✓ Criticidad: opcional (si se ingresa, debe existir en CRITICIDAD)
- ✓ Texto: obligatorio, único dentro de la sección
- ✓ Valor: número
- ✓ No puede eliminarse si hay preguntas que la referencian

#### PREGUNTAS
- ✓ Plantilla: **OBLIGATORIA** (debe existir)
- ✓ Sección: **OBLIGATORIA** (debe existir)
- ✓ Criticidad: opcional (si se especifica, debe existir)
- ✓ Texto: obligatorio, máx 500 caracteres
- ✓ Orden: número positivo
- ✓ Activo: Sí/No
- ✓ NO tiene dependencias descendentes

### 🚫 Errores Comunes

| Error | Causa | Solución |
|-------|-------|----------|
| "La Sección seleccionada no existe" | Referencia a SECCIÓN inexistente | Asegúrate de crear la SECCIÓN primero |
| "La Plantilla seleccionada no existe" | Referencia a PLANTILLA inexistente | Crea la PLANTILLA antes de la PREGUNTA |
| "Ya existe una criticidad con ese nombre" | Nombre duplicado | Usa otro nombre o edita el existente |
| "El valor debe ser numérico" | Ingresaste texto en campo numérico | Ingresa solo números |
| "Tipo de respuesta debe ser 'Selección' o 'Puntaje'" | Valor inválido | Selecciona exactamente uno de estos dos |
| "Este registro tiene dependencias" | Intentas eliminar algo que se usa | Marca como inactivo en lugar de eliminar |

---

## Operaciones CRUD

### CREATE (Crear - Botón [Nuevo])

1. **Selecciona tabla**
2. **Haz clic [Nuevo]**
   - Sistema genera ID automático
   - Campos se limpian para nueva entrada
   
3. **Completa datos**
   - Campos obligatorios: marcados con asterisco
   - Campos opcionales: pueden dejarse en blanco
   
4. **Haz clic [Guardar]**
   - Sistema valida todos los datos
   - Si hay error: muestra mensaje de validación
   - Si es exitoso: registro aparece en lista

**Ejemplo:**
```
[Nuevo] → Autocompleta ID → Ingresa "Alto Riesgo" en Nombre →
Ingresa "10" en Valor → [Guardar] → Aparece en lista
```

---

### READ (Leer - Seleccionar de La Lista)

1. **Selecciona tabla**
2. **Haz clic en cualquier registro en la lista**
   - Datos aparecen automáticamente en los campos
   - Puedes ver toda la información del registro

**Nota:** Los campos aparecen como "solo lectura" en modo visualización

---

### UPDATE (Editar - Seleccionar y Modificar)

1. **Selecciona tabla**
2. **Haz clic un registro en la lista**
   - Estado: "Registro seleccionado. Puede editar o eliminar."
   - Botones [Guardar] y [Eliminar] se activan

3. **Modifica cualquier campo** (excepto ID)
4. **Haz clic [Guardar]**
   - Sistema valida cambios
   - Si es válido: actualiza en tabla
   - Se registra en AUDIT TRAIL

**Ejemplo:**
```
Selecciona "Alto Riesgo" → Cambias Valor de 10 a 12 →
[Guardar] → Valor actualizado a 12
```

---

### DELETE (Eliminar - Botón [Eliminar])

**Dos opciones según dependencias:**

#### Opción A: Sin Dependencias
```
[Eliminar] → Confirmap pre-eliminación →
Se pregunta: "¿Está seguro de eliminar este registro?
Esta acción no se puede deshacer."
→ [Sí] → Registro eliminado permanentemente
```

#### Opción B: Con Dependencias (Soft-Delete)
```
[Eliminar] → Sistema detecta referencias →
Advierte: "Este registro tiene dependencias.
Se marcará como INACTIVO (soft-delete).
¿Desea continuar?"
→ [Sí] → Marca como inactivo (NO se elimina)
```

**Diferencia:**
- **Eliminación física:** Registro desaparece completamente (solo si nadie lo referencia)
- **Soft-delete (marcar inactivo):** Registro queda en tabla pero marcado como inactivo(aparecerá en el audit trail)

---

## Troubleshooting

### 🔧 Problemas y Soluciones

#### Problema: "Seleccione una tabla primero"
**Causa:** Hiciste clic en [Nuevo] sin seleccionar tabla  
**Solución:** Selecciona una tabla del dropdown antes de hacer clic en [Nuevo]

---

#### Problema: "La Sección seleccionada no existe en tblSecciones"
**Causa:** Intentas crear OPCIÓN o PREGUNTA referenciando SECCIÓN que no existe  
**Pasos:**
1. Haz clic en [Validar Todo] para detectar inconsistencias
2. Ve a SECCIONES y crea la sección faltante
3. Reinténtalo en las OPCIONES/PREGUNTAS

---

#### Problema: "Ya existe una criticidad con ese nombre"
**Causa:** El nombre ya está en uso en otra criticidad  
**Solución:**
- Opción 1: Usa otro nombre (ej: "Alto Riesgo 2")
- Opción 2: Edita la criticidad existente
- Opción 3: Marca la antigua como inactiva

---

#### Problema: Los ComboBox (desplegables) están vacíos
**Causa:** Las tablas de referencia no tienen datos  
**Ejemplo:** Intentas crear PREGUNTA pero no hay PLANTILLAS
**Solución:**
1. Crea datos en tablas base primero:
   - CRITICIDAD (indep.)
   - SECCIONES (indep.)
   - PLANTILLAS (indep.)
2. Luego crea dependientes:
   - OPCIONES (depende de SECCIONES)
   - PREGUNTAS (depende de todo)

---

#### Problema: "Este registro tiene dependencias"
**Causa:** Intentas eliminar un registro que se usa en otras tablas  
**Ejemplo:** SECCIÓN usada en OPCIONES o PREGUNTAS
```
┌─ SECCIÓN "Auditoría"
└─→ OPCIÓN 1: "Conforme"
    OPCIÓN 2: "No conforme"
    PREGUNTA 1: "¿Existen..."
    
Si intentas eliminar SECCIÓN → ERROR: Tiene dependencias
```

**Soluciones:**
1. **Opción A (Recomendado):** Marca como inactivo (soft-delete)
   - Registro queda en tabla para históricos
   - Ya no aparece en nuevas auditorías
   
2. **Opción B:** Elimina dependencias primero
   - Elimina todas las OPCIONES/PREGUNTAS que la usan
   - Luego elimina la SECCIÓN
   - Riesgo: Pierdes datos

---

#### Problema: "La frecuencia debe ser un número"
**Causa:** Ingresaste texto en campo de frecuencia  
**Solución:** Ingresa solo números (ej: 30, 90, 365)

---

#### Problema: Tipo de respuesta debe ser 'Selección' o 'Puntaje'
**Causa:** Ingresaste algo diferente (ej: "Opción", "Escala")  
**Solución:** Escribe exactamente: `Selección` o `Puntaje` (con mayúscula inicial)

---

### 🔍 Usar "Validar Todo"

El botón **[Validar Todo]** ejecuta un análisis completo:

1. **Verifica integridad referencial:**
   - ¿Todas las OPCIONES apuntan a SECCIONES que existen?
   - ¿Todas las PREGUNTAS apuntan a PLANTILLAS que existen?

2. **Detecta huérfanos:**
   - IDs en una tabla que no están en la tabla referenciada

3. **Reporta inconsistencias:**
   - Duplicados de nombres (si no permitidos)
   - Valores inválidos

4. **Generapalabra de reporte:**
   ```
   ERROR en PREGUNTAS (fila 45): ID PLANTILLA inexistente
   ERROR en OPCIONES (fila 12): SECCIÓN referencia invalida
   ...
   ```

**Cuándo usar:**
- ✓ Después de importar datos nuevos
- ✓ Ante sospecha de inconsistencias
- ✓ Antes de realizar auditorías importantes
- ✓ Regularmente (semanal/mensual)

---

### 📊 Verificar Audit Trail

Todos los cambios se registran automáticamente:

1. Ve a la hoja: **"Historico"**
2. Busca el registro de cambios
3. Aparecerá:
   ```
   Fecha | Hora | Usuario | Tabla | Acción | ID | Cambio | Antes | Después
   ```

**Ejemplo:**
```
2026-04-14 | 14:30 | Juan | CRITICIDAD | NUEVO | abc123 | ... | ...
2026-04-14 | 14:35 | Juan | SECCIONES | EDITAR | def456 | Nombre | "Auditoría" → "Auditoría de Procesos"
2026-04-14 | 14:40 | Juan | OPCIONES | DELETE (soft) | ghi789 | Activo | "Sí" → "No"
```

---

## 🎯 Flujo Completo: Ejemplo Real

**Objetivo:** Crear una auditoría completa llamada "Auditoría Q1 2026"

### Paso 1: Crear Criticidades
```
Tabla → [1. Criticidades]
[Nuevo] → Nombre: "Crítico" → Valor: 10 → [Guardar]
[Nuevo] → Nombre: "Alto" → Valor: 7 → [Guardar]
[Nuevo] → Nombre: "Medio" → Valor: 5 → [Guardar]
[Nuevo] → Nombre: "Bajo" → Valor: 2 → [Guardar]
```

### Paso 2: Crear Secciones
```
Tabla → [2. Secciones]
[Nuevo] → Nombre: "Procesos" → Tipo: "Selección" → [Guardar]
[Nuevo] → Nombre: "Seguridad" → Tipo: "Puntaje" → [Guardar]
[Nuevo] → Nombre: "Documentación" → Tipo: "Selección" → [Guardar]
```

### Paso 3: Crear Plantilla
```
Tabla → [3. Plantillas]
[Nuevo] → Nombre: "Auditoría Q1 2026" → Etapa: "Inicial" → Puesto: "Auditor Senior" → Frecuencia: 90 → [Guardar]
```

### Paso 4: Crear Opciones
```
Tabla → [4. Opciones]

Para "Procesos":
[Nuevo] → Sección: "Procesos" → Criticidad: "Alto" → Texto: "Conforme" → Valor: 0 → [Guardar]
[Nuevo] → Sección: "Procesos" → Criticidad: "Crítico" → Texto: "No conforme" → Valor: -10 → [Guardar]

Para "Seguridad":
[Nuevo] → Sección:"Seguridad" → Texto: "Excelente" → Valor: 5 → [Guardar]
[Nuevo] → Sección: "Seguridad" → Criticidad: "Crítico" → Texto: "Deficiente" → Valor: 1 → [Guardar]
```

### Paso 5: Crear Preguntas
```
Tabla → [5. Preguntas]

[Nuevo]
  Plantilla: "Auditoría Q1 2026"
  Sección: "Procesos"
  Criticidad: "Alto"
  Texto: ¿Existen procedimientos documentados?
  Orden: 1
  Activo: ☑
  [Guardar]

[Nuevo]
  Plantilla: "Auditoría Q1 2026"
  Sección: "Procesos"
  Criticidad: "Crítico"
  Texto: ¿Se cumple la documentación vigente?
  Orden: 2
  Activo: ☑
  [Guardar]

[Nuevo]
  Plantilla: "Auditoría Q1 2026"
  Sección: "Seguridad"
  Criticidad: "Crítico"
  Texto: ¿Están todas las capacitaciones de seguridad vigentes?
  Orden: 1
  Activo: ☑
  [Guardar]
```

### Resultado
**Auditoría "Q1 2026" lista con:**
- ✓ 3 Secciones
- ✓ 5 Opciones de respuesta
- ✓ 3 Preguntas
- ✓ Criticidades definidas
- ✓ Todos los cambios registrados en Audit Trail

---

## 📞 Soporte

- **Error en validación:** Consulta la sección "Validaciones y Reglas"
- **Comportamiento inesperado:** Haz clic en [Validar Todo]
- **Datos inconsistentes:** Revisa el Audit Trail en hoja "Historico"
- **Necesitas restaurar datos:** Usa copias de seguridad automáticas (hablar con TI)

---

**Versión:** 1.0  
**Creado:** Abril 2026  
**Última actualización:** Abril 2026  
**Próxima revisión:** Junio 2026
