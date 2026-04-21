# Instrucciones: Agregar Columna ID Criticidad a tblRespuestas

## Propósito
Para que el sistema de Auditoría de Procesos funcione correctamente, necesitamos guardar el **ID de Criticidad** de cada respuesta en la tabla `tblRespuestas`. Esto permite:
- Persistir el nivel de criticidad de cada pregunta respondida
- Recalcular evaluaciones de Auditoría de Procesos basándose en datos históricos
- Generar reportes y análisis por criticidad

---

## Pasos para Agregar la Columna

### 1. Abrir el Archivo Excel
1. Abrir `TH-HC-001 INSPECCIONES.xlsm`
2. Ir a la hoja **"Histórico"**

### 2. Localizar la Tabla tblRespuestas
1. La tabla `tblRespuestas` debe estar en la hoja "Histórico"
2. Verificar que tenga estas columnas (en orden):
   - [1] **ID Respuesta** (texto)
   - [2] **ID Inspeccion** (texto, FK a tblInspecciones)
   - [3] **ID Pregunta** (texto, FK a tblPreguntas)
   - [4] **ID Opcion** (texto, FK a tblOpcionesDeRespuesta)
   - [5] **Valor numerico** (número: 0, 1, etc.)
   - [6] **Observacion** (texto)
   - [7] **Fecha respuesta** (fecha/hora)

### 3. Agregar Nueva Columna
1. Hacer clic en la **última columna** de la tabla (actualmente "Fecha respuesta")
2. Hacer clic derecho sobre el encabezado de esa columna
3. Seleccionar **"Insertar" → "Columnas de tabla a la derecha"**
4. Cambiar el nombre de la nueva columna a: **`ID Criticidad`**
   - Debe ser exactamente este nombre (sin acentos, con espacio)
   - Tipo de dato: **Texto**

### 4. Estructura Final Esperada
La tabla `tblRespuestas` debe tener estas 8 columnas:
```
[1] ID Respuesta
[2] ID Inspeccion
[3] ID Pregunta
[4] ID Opcion
[5] Valor numerico
[6] Observacion
[7] Fecha respuesta
[8] ID Criticidad          ← NUEVA COLUMNA
```

---

## Validación

Para verificar que se agregó correctamente:
1. Ejecutar una inspección de prueba de tipo "Auditoría de Procesos"
2. Responder al menos 2 preguntas con criticidad "Crítica" como "No Cumple"
3. Guardar la inspección
4. Abrir la hoja "Histórico" y verificar que:
   - Las filas en `tblRespuestas` tienen valores en la columna "ID Criticidad"
   - Los valores son IDs válidos (ej: `M4dcAe5B-wo5vnoDp-eiEoqu5p`)

---

## Notas Técnicas

### Relación con tblPreguntas
- Cada pregunta en `tblPreguntas` tiene un campo "ID Criticidad"
- `tblRespuestas[ID Criticidad]` es una **copia** del valor de `tblPreguntas[ID Criticidad]` para la pregunta respondida
- Esto permite denormalización: aunque la criticidad se puede obtener haciendo JOIN con tblPreguntas, guardarla en tblRespuestas:
  - Mejora el rendimiento (evita JOINs complejos)
  - Preserva el valor histórico (si la criticidad de una pregunta cambia en el futuro)
  - Simplifica los cálculos de Auditoría de Procesos

### Valores Posibles
Los IDs de criticidad válidos están en `tblCriticidad` (hoja "Configuración"):
- **Crítica**: ID específico (ej: `M4dcAe5B-wo5vnoDp-eiEoqu5p`)
- **Mayor**: ID específico
- **Menor**: ID específico

### Compatibilidad con Datos Existentes
- Las respuestas guardadas ANTES de agregar esta columna tendrán valores vacíos en "ID Criticidad"
- Para recárculos de inspecciones antiguas, se puede hacer JOIN con tblPreguntas
- El código actual maneja valores vacíos clasificándolos como criticidad "Menor" (fallback)

---

## Problemas Comunes

### Error: "No se encuentra la columna 'ID Criticidad'"
**Causa**: El nombre de la columna no coincide exactamente.
**Solución**: Verificar que el nombre sea `ID Criticidad` (sin acentos, con espacio, capitalización exacta)

### Error al guardar inspección
**Causa**: La columna no se agregó después de "Fecha respuesta"
**Solución**: Verificar que el orden de columnas coincida con la estructura esperada (tabla arriba)

---

## Fecha de Creación
2025-01-XX

## Autor
Sistema de Inspecciones - Módulo InspectionRepository  
