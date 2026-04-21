# Guía de Debug - Auditoría de Procesos

## Propósito
Este documento explica cómo usar el debug logging extensivo para diagnosticar problemas con el cálculo de Auditoría de Procesos.

---

## Preparación

### 1. Verificar Columna ID Criticidad en tblRespuestas

**CRÍTICO**: La tabla `tblRespuestas` DEBE tener una columna 8 llamada `ID Criticidad`.

1. Abrir Excel → Hoja "Histórico"
2. Localizar tabla `tblRespuestas`
3. Verificar que tenga **8 columnas** en este orden:
   ```
   [1] ID Respuesta
   [2] ID Inspeccion
   [3] ID Pregunta
   [4] ID Opcion
   [5] Valor numerico
   [6] Observacion
   [7] Fecha respuesta
   [8] ID Criticidad  ← DEBE EXISTIR
   ```

**Si no existe**: Ver [INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md](INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md)

### 2. Importar Código Actualizado

1. Abrir VBA Editor (Alt+F11)
2. Importar/reemplazar estos módulos con las versiones actualizadas:
   - `InspectionCalculator.bas`
   - `InspectionRepository.bas`
   - `frmChecklistVirtual.frm`
3. Guardar el archivo (Ctrl+S)
4. Cerrar y reabrir Excel (para asegurar que se cargue el código nuevo)

### 3. Abrir Ventana Inmediato

1. En VBA Editor: Ver → Ventana Inmediato (Ctrl+G)
2. Esta ventana mostrará todos los mensajes de debug

---

## Procedimiento de Prueba

### Paso 1: Limpiar Ventana Inmediato
1. En la Ventana Inmediato, seleccionar todo (Ctrl+A) y eliminar (Del)
2. Esto asegura que solo veas los logs de la prueba actual

### Paso 2: Crear Inspección de Prueba
1. Ir al menú principal del sistema
2. Crear nueva inspección de tipo **"Auditoría de Procesos"**
3. Llenar los datos básicos (evaluado, fecha, etc.)

### Paso 3: Responder Preguntas con Criticidad "Crítica"
Para probar la **Regla 1** (≥2 "No Cumple" de criticidad Crítica → "No Cumple"):

1. Responder **al menos 2 preguntas** de criticidad **"Crítica"** con **"No Cumple"**
2. Las demás preguntas se pueden responder como "Cumple" o "No Aplica"

### Paso 4: Guardar Inspección
1. Hacer clic en **Guardar**
2. Esperar confirmación
3. **NO CERRAR EXCEL** todavía

### Paso 5: Revisar Ventana Inmediato

Ir a VBA Editor → Ventana Inmediato y buscar estos mensajes clave:

---

## Qué Buscar en los Logs

### A. Logs del Formulario (frmChecklistVirtual)

#### A1. Inicialización de Preguntas
```vba
[frmChecklistVirtual.CrearControlesPreguntas] Pregunta XXX inicializada con IDCriticidad: M4dcAe5B-wo5vnoDp-eiEoqu5p
```
✅ **Verificar**: Cada pregunta muestra su IDCriticidad correctamente.

#### A2. Captura de Respuestas
```vba
[frmChecklistVirtual.RecopilarRespuestas] Pregunta XXX: IDOpcion=YYY, ValorNum=0, IDCriticidad=M4dcAe5B-wo5vnoDp-eiEoqu5p
```
✅ **Verificar**: 
- Las respuestas "No Cumple" tienen `ValorNum=0`
- Cada respuesta conserva su `IDCriticidad`

#### A3. Obtener Respuestas con Sección
```vba
[frmChecklistVirtual.ObtenerRespuestasConSeccion] Iniciando - Total respuestas en mRespuestas: 10
  [ObtenerRespuestasConSeccion] Respuesta #1 - IDPregunta: XXX, IDSeccion: YYY, IDCriticidad: M4dcAe5B-wo5vnoDp-eiEoqu5p, IDOpcion: ZZZ
  [ObtenerRespuestasConSeccion] Respuesta #2 - ...
[frmChecklistVirtual.ObtenerRespuestasConSeccion] Completado - Total items en Collection: 10
```
✅ **Verificar**: Todas las respuestas tienen `IDCriticidad` poblado.

---

### B. Logs de Guardado (InspectionRepository)

#### B1. Verificación de Columna ID Criticidad
```vba
[GuardarRespuestas] Total columnas en tblRespuestas: 8
[GuardarRespuestas] Nombre columna 8: 'ID Criticidad'
[GuardarRespuestas] ✓ Columna ID Criticidad encontrada: ID Criticidad
```
✅ **Verificar**: Se encuentra la columna 8 y contiene "Criticidad" en el nombre.

❌ **SI VES ESTO**:
```vba
[GuardarRespuestas] Total columnas en tblRespuestas: 7
[GuardarRespuestas] ✗ ADVERTENCIA: No hay columna 8 (ID Criticidad).
```
**ACCIÓN**: Debes agregar la columna 8 manualmente. Ver [INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md](INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md)

#### B2. Guardado de Cada Respuesta
```vba
[GuardarRespuestas] Procesando respuesta 1
  IDPregunta: XXX
  IDOpcion: YYY
  ValorNumerico: 0
  IDCriticidad: M4dcAe5B-wo5vnoDp-eiEoqu5p
  Estableciendo Columna 8 (ID Criticidad)...
  Respuesta 1 completada
```
✅ **Verificar**: Cada respuesta muestra "Estableciendo Columna 8" y el `IDCriticidad` está presente.

---

### C. Logs de Cálculo (InspectionCalculator)

#### C1. Inicio del Conteo
```vba
  [ContarRespuestasPorCriticidad] Total respuestas recibidas: 10
  [ContarRespuestasPorCriticidad] ID Sección Procesos buscada: XXXXXXX
```

#### C2. Procesamiento de Cada Respuesta
```vba
    Respuesta - IDSeccion: XXXXXXX | IDPregunta: YYYYYYY
      >>> PROCESANDO RESPUESTA #1
          IDCriticidad: M4dcAe5B-wo5vnoDp-eiEoqu5p
          NombreCriticidad: Crítica
          IDOpcion: ZZZZZZZ
          TextoOpcion: No Cumple
          Incrementado: Crítica_NoCumple = 1
```
✅ **Verificar**:
- `IDCriticidad` tiene un valor UUID válido
- `NombreCriticidad` es "Crítica", "Mayor" o "Menor" (NO debe ser "Menor" como fallback si es Crítica)
- `TextoOpcion` es "Cumple", "No Cumple" o "No Aplica"
- Los contadores se incrementan correctamente

#### C3. Búsqueda en tblCriticidad
```vba
        [ObtenerNombreCriticidad] Buscando ID: 'M4dcAe5B-wo5vnoDp-eiEoqu5p'
        [ObtenerNombreCriticidad] Filas en tblCriticidad: 3
        [ObtenerNombreCriticidad] Fila 1 - ID: 'M4dcAe5B-wo5vnoDp-eiEoqu5p'
        [ObtenerNombreCriticidad] ¡MATCH! Retornando: 'Crítica'
```
✅ **Verificar**: Encuentra match exacto y retorna "Crítica".

❌ **SI VES ESTO**:
```vba
        [ObtenerNombreCriticidad] NO ENCONTRADO. Retornando 'Menor' como fallback
```
**PROBLEMA**: El ID de criticidad en tblPreguntas NO coincide con los IDs en tblCriticidad.
**ACCIÓN**: Revisar manualmente:
1. Hoja "Configuración" → tblPreguntas → Columna "ID Criticidad"
2. Hoja "Configuración" → tblCriticidad → Columna "ID Criticidad"
3. Verificar que los IDs coincidan exactamente (copiar/pegar para asegurar)

#### C4. Búsqueda en tblOpcionesDeRespuesta
```vba
        [ObtenerTextoOpcionPorID] Buscando ID: 'XXXXXXX'
        [ObtenerTextoOpcionPorID] Filas en tblOpciones: 15
        [ObtenerTextoOpcionPorID] Fila 1 - ID: 'YYYYYYY'
        [ObtenerTextoOpcionPorID] ¡MATCH! Retornando: 'No Cumple'
```
✅ **Verificar**: Encuentra la opción y retorna el texto correcto.

#### C5. Resumen de Conteo (ChecklistOrchestrator)
```vba
  Auditoría de Procesos - Conteo por criticidad:
    Crítica - Cumple: 0
    Crítica - No Cumple: 2
    Crítica - No Aplica: 0
    Mayor - Cumple: 5
    Mayor - No Cumple: 0
    Mayor - No Aplica: 1
    Menor - Cumple: 2
    Menor - No Cumple: 0
    Menor - No Aplica: 0
  Resultado Auditoría de Procesos: No Cumple
```
✅ **Verificar**:
- `Crítica - No Cumple: 2` (o el número que respondiste)
- `Resultado Auditoría de Procesos: No Cumple` (debe ser "No Cumple" si hay ≥2 Crítica)

---

## Problemas Comunes y Soluciones

### Problema 1: "Columna 8 no existe"
**Log:**
```
[GuardarRespuestas] ✗ ADVERTENCIA: No hay columna 8 (ID Criticidad).
```
**Solución**: Agregar columna "ID Criticidad" como columna 8 en tblRespuestas. Ver [INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md](INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md)

---

### Problema 2: "IDCriticidad vacío o undefined"
**Log:**
```
  IDCriticidad: 
```
o
```
  [ObtenerNombreCriticidad] Buscando ID: ''
```
**Causas posibles**:
1. tblPreguntas no tiene la columna "ID Criticidad"
2. Las preguntas no tienen asignado un ID de criticidad
3. El código no se importó correctamente

**Solución**:
1. Ir a Hoja "Configuración" → tblPreguntas
2. Verificar que existe columna "ID Criticidad" (columna 5)
3. Verificar que cada pregunta tiene un valor en esa columna
4. Reimportar el código de `ChecklistRepository.bas` (línea 131 debe leer `preg(4)`)

---

### Problema 3: "Criticidad siempre es 'Menor'"
**Log:**
```
        [ObtenerNombreCriticidad] Buscando ID: 'M4dcAe5B-wo5vnoDp-eiEoqu5p'
        [ObtenerNombreCriticidad] NO ENCONTRADO. Retornando 'Menor' como fallback
```
**Causa**: El ID en tblPreguntas NO existe en tblCriticidad.

**Solución**:
1. Ir a Hoja "Configuración"
2. Copiar IDs de tblCriticidad (columna "ID Criticidad"):
   - Crítica: [copiar ID exacto]
   - Mayor: [copiar ID exacto]
   - Menor: [copiar ID exacto]
3. En tblPreguntas, PEGAR esos IDs exactos (no escribir manualmente)

---

### Problema 4: "Resultado siempre 'Cumple'"
**Log:**
```
  Resultado Auditoría de Procesos: Cumple
```
pero tenías 2 Crítica No Cumple.

**Causas posibles**:
1. Las respuestas no se están contando (todos los contadores en 0)
2. El IDSeccion de las preguntas no coincide con `idSeccionProcesos`
3. La función `EvaluarAuditoriaProcesos` no está usando los valores correctos

**Solución**:
1. Revisar el log de conteo:
   ```
     Crítica - No Cumple: ?
   ```
   Si es 0, las respuestas no se están procesando.

2. Verificar en el log:
   ```
   [ContarRespuestasPorCriticidad] ID Sección Procesos buscada: XXXXXXX
   Respuesta - IDSeccion: YYYYYYY | ...
   ```
   Si XXXXXXX ≠ YYYYYYY, las respuestas no se están filtrando.

3. Verificar que las preguntas de Auditoría de Procesos están en la sección correcta en tblPreguntas.

---

### Problema 5: "No guarda en columna ID Criticidad"
**Verificación manual**:
1. Después de guardar, ir a Hoja "Histórico" → tblRespuestas
2. Ver las últimas filas agregadas
3. Columna 8 ("ID Criticidad") debe tener valores UUID

**Si está vacía**:

**Log a buscar**:
```
  Estableciendo Columna 8 (ID Criticidad)...
```

**Si NO aparece**: La columna no se detectó como existente.

**Si SÍ aparece pero la celda está vacía**: 
```
  Columna 8 dejada vacía (IDCriticidad no presente en respuesta)
```
El problema está en el formulario, no en el guardado.

---

## Resultado Esperado (Caso de Prueba Exitoso)

### Respuestas de Prueba:
- 2 preguntas con criticidad "Crítica" → "No Cumple"
- 5 preguntas con criticidad "Mayor" → "Cumple"
- 3 preguntas con criticidad "Menor" → "Cumple"

### Debug Output Esperado (Resumen):
```
[PASO 8A] Procesando Auditoría de Procesos...
  Auditoría de Procesos - Conteo por criticidad:
    Crítica - Cumple: 0
    Crítica - No Cumple: 2      ← 2 respuestas Crítica No Cumple
    Crítica - No Aplica: 0
    Mayor - Cumple: 5
    Mayor - No Cumple: 0
    Mayor - No Aplica: 0
    Menor - Cumple: 3
    Menor - No Cumple: 0
    Menor - No Aplica: 0
  Resultado Auditoría de Procesos: No Cumple  ← ✓ CORRECTO (Regla 1 aplicada)
[PASO 8A] Auditoría de Procesos procesada OK
```

### Verificación en Excel:
1. Hoja "Inspecciones" → tblInspecciones → Columna "Auditoria Procesos Resultado" = **"No Cumple"**
2. Hoja "Histórico" → tblRespuestas → Columna 8 "ID Criticidad" tiene valores UUID válidos

---

## Reportar Problemas

Si después de seguir esta guía el problema persiste:

1. **Copiar TODO el contenido de la Ventana Inmediato**
2. **Tomar captura de pantalla de**:
   - Hoja "Histórico" → tblRespuestas (mostrar columnas 1-8)
   - Hoja "Configuración" → tblCriticidad
   - Hoja "Configuración" → tblPreguntas (columnas: ID Pregunta, Texto, ID Criticidad)
3. **Compartir**:
   - Log completo de Ventana Inmediato
   - Capturas de pantalla
   - Descripción de las respuestas que ingresaste (cuántas Crítica No Cumple, etc.)

---

## Fecha
Abril 20, 2026

## Autor
Sistema de Inspecciones - Debug Module
