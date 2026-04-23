# Guía de Depuración: Búsqueda de Historial de Inspecciones

## 📋 **Problema Reportado**

**Síntoma**: Al hacer clic en "Buscar historial" aparece "error 0" sin mostrar información en pantalla.

**Contexto**: 
- Usuario completó una primera inspección exitosamente
- Al intentar hacer una segunda inspección, marcó "Esta NO es la primera inspección"
- Hizo clic en "Buscar historial"
- Sistema mostró "error 0" (sin descripción)

---

## 🔍 **Posibles Causas**

### **1. Datos de entrada vacíos o incorrectos**

El servicio requiere 3 datos obligatorios:
- **Iniciales del personal** (mEvaluado)
- **Puesto** (txtPuesto.Value)
- **IDPlantilla** (opcional pero recomendado)

**Cómo verificar:**
```
1. Abre VBE (Alt+F11)
2. Ventana Inmediato (Ctrl+G)
3. Busca las líneas:
   [BUSCAR] Evaluado: [...]
   [BUSCAR] Puesto: [...]
   [BUSCAR] IDPlantilla: [...]
```

**Problemas comunes:**
- ✅ Iniciales vacías: `[BUSCAR] Evaluado: []`
- ✅ Puesto vacío: `[BUSCAR] Puesto: []`
- ⚠️ IDPlantilla vacía: No es crítico, pero puede causar resultados inesperados

### **2. Base de datos vacía o sin datos coincidentes**

Si la tabla `tblInspecciones` no tiene registros previos para ese personal/puesto:

**Comportamiento esperado:**
```
[BUSCAR] Inspecciones encontradas: 0
MsgBox: "No se encontraron inspecciones anteriores para..."
```

**Cómo verificar:**
```
1. Ir a hoja "Historico"
2. Buscar en tblInspecciones
3. Filtrar por:
   - Columna "Iniciales personal" = [tu personal]
   - Columna "Puesto Evaluado" = [tu puesto]
4. Si no hay filas: Es normal, es la primera inspección
```

### **3. Error en el servicio InspectionHistoryService**

El servicio puede lanzar errores si:
- La estructura de la tabla ha cambiado
- Las columnas esperadas no existen
- Hay datos corruptos en tblInspecciones

**Cómo verificar:**
```
Busca en la Ventana Inmediato:
[BUSCAR] ERROR en BuscarInspeccionesPrevias: [número] - [descripción]
```

**Errores comunes:**
- `Error 5001`: "Las iniciales del personal son requeridas"
- `Error 5002`: "El puesto es requerido cuando filtroPorPuesto=True"
- `Error 9`: "El índice está fuera del intervalo" (columna no existe)

### **4. Error "0" (sin código)**

Si aparece "Error 0" específicamente, indica:
- Un error fue capturado y limpiado (`On Error Resume Next`)
- Pero el `Err.Number` quedó en 0
- Esto es un bug en el manejo de errores

**Nueva implementación (22/04/2026):**
```vba
' Ahora detecta error 0 y muestra mensaje específico:
If Err.Number = 0 Then
    mensajeError = "Error inesperado sin código de error."
    "Esto puede indicar un problema en el servicio de historial."
    "Revise la ventana Inmediato (Ctrl+G) para ver los logs detallados."
End If
```

---

## 🛠️ **Mejoras Implementadas (22/04/2026)**

### **1. Validación robusta de entrada**

**ANTES:**
```vba
If Len(mEvaluado) = 0 Or Len(Me.txtPuesto.Value) = 0 Then
    MsgBox "Debe seleccionar un personal antes de buscar historial."
    Exit Sub
End If
```

**AHORA:**
```vba
If Len(Trim(mEvaluado)) = 0 Then
    Debug.Print "[BUSCAR] ERROR: mEvaluado está vacío"
    MsgBox "Error: No se ha definido el personal evaluado (Iniciales)." & vbCrLf & _
           "Por favor, regrese al selector y elija el personal a inspeccionar."
    Exit Sub
End If

If Len(Trim(Me.txtPuesto.Value)) = 0 Then
    Debug.Print "[BUSCAR] ERROR: Puesto está vacío"
    MsgBox "Error: No se ha definido el puesto del personal." & vbCrLf & _
           "Por favor, regrese al selector y elija el puesto a inspeccionar."
    Exit Sub
End If
```

### **2. Logging detallado del proceso**

**Logs agregados:**
```
[BUSCAR] ===== INICIO Búsqueda de historial =====
[BUSCAR] Evaluado: [NCE]
[BUSCAR] Puesto: [Analista QC]
[BUSCAR] IDPlantilla: [abc-123-xyz]
[BUSCAR] Llamando a InspectionHistoryService.BuscarInspeccionesPrevias...
[BUSCAR]   - iniciales: [NCE]
[BUSCAR]   - filtroPorPuesto: True
[BUSCAR]   - puesto: [Analista QC]
[BUSCAR]   - plantillaID: [abc-123-xyz]
[BUSCAR] Inspecciones encontradas: 2
[BUSCAR] Obteniendo última inspección...
[BUSCAR] Última inspección obtenida: ZMGgBxVC-7ZYZrOAa-PNe8x7PFr0
[BUSCAR] ===== FIN Búsqueda de historial - ÉXITO =====
```

### **3. Manejo de errores con On Error Resume Next**

**Protección contra errores sin capturar:**
```vba
On Error Resume Next
Set inspecciones = InspectionHistoryService.BuscarInspeccionesPrevias(...)

' Verificar si hubo error
If Err.Number <> 0 Then
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    On Error GoTo ErrorHandler
    
    Debug.Print "[BUSCAR] ERROR en BuscarInspeccionesPrevias: " & errNum & " - " & errDesc
    MsgBox "Error al buscar inspecciones previas:" & vbCrLf & _
           "Número: " & errNum & vbCrLf & _
           "Descripción: " & errDesc
    Exit Sub
End If
On Error GoTo ErrorHandler
```

### **4. Validación de Nothing**

**Protección contra resultados Nothing:**
```vba
' Verificar que el resultado no sea Nothing
If inspecciones Is Nothing Then
    Debug.Print "[BUSCAR] ERROR: BuscarInspeccionesPrevias devolvió Nothing"
    MsgBox "Error: No se pudo realizar la búsqueda de inspecciones previas." & vbCrLf & _
           "El servicio de historial no respondió correctamente."
    Exit Sub
End If
```

### **5. ErrorHandler mejorado**

**ANTES:**
```vba
ErrorHandler:
    Call ErrorLogger2.Log("...", Err.Description, Err.Number)
    MsgBox "Error al buscar historial: " & Err.Description & vbCrLf & _
           "Número de error: " & Err.Number, vbCritical, "Error"
```

**AHORA:**
```vba
ErrorHandler:
    Debug.Print "[BUSCAR] ===== ERROR en btnBuscarHistorico_Click ====="
    Debug.Print "[BUSCAR] Error Number: " & Err.Number
    Debug.Print "[BUSCAR] Error Description: " & Err.Description
    Debug.Print "[BUSCAR] Error Source: " & Err.Source
    
    ' Mensaje específico para error 0
    If Err.Number = 0 Then
        mensajeError = "Error inesperado sin código de error." & vbCrLf & _
                      "Esto puede indicar un problema en el servicio de historial." & vbCrLf & _
                      "Revise la ventana Inmediato (Ctrl+G) para ver los logs detallados."
    Else
        mensajeError = "Número de error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Datos buscados:" & vbCrLf & _
                      "  - Personal: " & mEvaluado & vbCrLf & _
                      "  - Puesto: " & Me.txtPuesto.Value & vbCrLf & _
                      "  - Plantilla: " & mIDPlantilla
    End If
    
    MsgBox mensajeError, vbCritical, "Error de búsqueda de historial"
```

---

## 🧪 **Cómo Depurar el Problema**

### **Paso 1: Activar la Ventana Inmediato**

```
1. Abrir Excel con el archivo TH-HC-001 INSPECCIONES.xlsm
2. Presionar Alt+F11 para abrir VBE
3. Presionar Ctrl+G para abrir Ventana Inmediato
4. Dejar esta ventana visible
```

### **Paso 2: Reproducir el error**

```
1. Volver a Excel (Alt+F11)
2. Iniciar nueva inspección
3. Marcar "Esta NO es la primera inspección"
4. Click [Buscar historial]
5. Observar el mensaje de error
```

### **Paso 3: Revisar logs en Ventana Inmediato**

**Buscar líneas que empiecen con `[BUSCAR]`:**

✅ **Caso exitoso (historial encontrado):**
```
[BUSCAR] ===== INICIO Búsqueda de historial =====
[BUSCAR] Evaluado: [NCE]
[BUSCAR] Puesto: [Analista QC]
[BUSCAR] Llamando a InspectionHistoryService.BuscarInspeccionesPrevias...
[BUSCAR] Inspecciones encontradas: 2
[BUSCAR] Obteniendo última inspección...
[BUSCAR] Última inspección obtenida: ZMGgBxVC-7ZYZrOAa-PNe8x7PFr0
[BUSCAR] ===== FIN Búsqueda de historial - ÉXITO =====
```

✅ **Caso sin historial (primera inspección):**
```
[BUSCAR] ===== INICIO Búsqueda de historial =====
[BUSCAR] Evaluado: [NCE]
[BUSCAR] Puesto: [Analista QC]
[BUSCAR] Llamando a InspectionHistoryService.BuscarInspeccionesPrevias...
[BUSCAR] Inspecciones encontradas: 0
```
→ Muestra MsgBox: "No se encontraron inspecciones anteriores..."

❌ **Caso con error de validación:**
```
[BUSCAR] ===== INICIO Búsqueda de historial =====
[BUSCAR] Evaluado: []
[BUSCAR] ERROR: mEvaluado está vacío
```
→ Muestra MsgBox: "Error: No se ha definido el personal evaluado..."

❌ **Caso con error en servicio:**
```
[BUSCAR] ===== INICIO Búsqueda de historial =====
[BUSCAR] Evaluado: [NCE]
[BUSCAR] Puesto: [Analista QC]
[BUSCAR] Llamando a InspectionHistoryService.BuscarInspeccionesPrevias...
[BUSCAR] ERROR en BuscarInspeccionesPrevias: 5001 - Las iniciales del personal son requeridas
```

### **Paso 4: Interpretar resultados**

**Si ves `[BUSCAR] Inspecciones encontradas: 0`:**
- ✅ **Normal**: Es la primera inspección de ese puesto
- ✅ **Solución**: Desmarcar checkbox "Esta NO es la primera inspección"

**Si ves `[BUSCAR] ERROR: mEvaluado está vacío`:**
- ❌ **Problema**: Los datos del formulario no se cargaron correctamente
- ✅ **Solución**: Cerrar formulario, volver al selector, elegir personal nuevamente

**Si ves `[BUSCAR] ERROR en BuscarInspeccionesPrevias: [número]`:**
- ❌ **Problema**: Error en el servicio de historial
- ✅ **Solución**: Anotar número y descripción del error, contactar soporte

**Si NO ves ningún log con `[BUSCAR]`:**
- ❌ **Problema crítico**: El evento del botón no se está ejecutando
- ✅ **Solución**: 
  1. Verificar que el botón existe: `btnBuscarHistorico`
  2. Verificar que el evento está configurado: `btnBuscarHistorico_Click`
  3. Revisar si hay código que deshabilita el botón

---

## 📊 **Estructura de Datos Esperada**

### **Tabla: tblInspecciones (Hoja "Historico")**

**Columnas requeridas:**
- `ID Inspeccion` (String)
- `Iniciales personal` (String)
- `Puesto Evaluado` (String) ← **NUEVA columna**
- `Fecha inspeccion` (Date)
- `RPN calculado` (Double)
- `Categoria resultado` (String)
- `ID Plantilla` (String)

**Columnas opcionales (nuevas):**
- `Numero Inspeccion` (Long) - 1, 2, 3, etc.
- `Es Inspeccion Recurrente` (String) - "SI" o "NO"
- `RPN Total` (Double) - Para inspecciones recurrentes

**Ejemplo de datos:**
```
ID Inspeccion          | Iniciales | Puesto        | Fecha      | RPN   | Numero | Es Recurrente
----------------------|-----------|---------------|------------|-------|--------|---------------
abc-123-xyz           | NCE       | Analista QC   | 2026-04-15 | 75.50 | 1      | NO
def-456-uvw           | NCE       | Analista QC   | 2026-04-22 | 68.25 | 2      | SI
```

---

## 🚨 **Problemas Conocidos y Soluciones**

### **Problema 1: Error 5 - "Argumento o llamada a procedimiento no válida"**

**Síntoma**: 
```
[BUSCAR] ERROR en ObtenerUltimaInspeccion: 5 - Argumento o llamada a procedimiento no válida
```

**Causa**: **Referencia circular en VBA**. El código original asignaba el resultado a la función y luego intentaba acceder a ella:
```vba
' CÓDIGO INCORRECTO (causa Error 5)
Set ObtenerUltimaInspeccion = inspecciones(1)
Debug.Print ObtenerUltimaInspeccion("IDInspeccion")  ' ❌ Error 5
```

Cuando asignas una variable con el mismo nombre que la función, VBA se confunde al intentar acceder a propiedades del objeto.

**Estado**: ✅ **RESUELTO** - InspectionHistoryService.bas actualizado

---

### **Problema 2: "Error 0" sin descripción**

**Síntoma**: MsgBox muestra "Error al buscar historial: [sin descripción] - Número de error: 0"

**Causa**: El error fue capturado con `On Error Resume Next` pero `Err.Number` se limpió

**Solución implementada**: 
- Capturar `Err.Number` y `Err.Description` en variables locales antes de cambiar el error handler
- Mostrar mensaje específico para error 0

---

### **Problema 3: No encuentra inspecciones cuando sí existen**

**Síntoma**: MsgBox "No se encontraron inspecciones anteriores" pero en Historico hay datos

**Causas posibles:**
1. **Iniciales no coinciden**: 
   - En formulario: "NCE"
   - En tabla: "NCE " (con espacio)
   - Solución: Se agregó `Trim()` en la búsqueda

2. **Puesto no coincide**:
   - En formulario: "Analista QC"
   - En tabla: "ANALISTA QC" (mayúsculas)
   - Solución: Comparación con `UCase()` en InspectionHistoryService

3. **Columna "Puesto Evaluado" no existe**:
   - La tabla no tiene esta columna
   - Solución: InspectionHistoryService maneja columnas opcionales con `On Error Resume Next`

---

### **Problema 4: Checkbox se desmarca pero no resetea variables**

**Síntoma**: Al buscar historial sin resultados, el checkbox se desmarca pero al guardar pide RPN anterior

**Causa**: Cambiar `chkEsRecurrente.Value = False` programáticamente NO dispara el evento `Click()`

**Solución implementada (22/04/2026)**:
```vba
' Usar flag para prevenir recursión
mActualizandoCheckbox = True
chkEsRecurrente.Value = False
mActualizandoCheckbox = False

' Resetear variables MANUALMENTE
mEsInspeccionRecurrente = False
mNumeroInspeccion = 1
mRPNAnteriorManual = 0
mRPNAnteriorAuto = 0
mIDInspeccionAnterior = ""
mModoRPN = "NINGUNO"
```

---

## 📞 **Contacto y Soporte**

Si después de seguir esta guía el problema persiste:

1. **Copiar logs de la Ventana Inmediato:**
   - Ctrl+A en Ventana Inmediato
   - Ctrl+C para copiar
   - Pegar en un archivo de texto

2. **Incluir información del contexto:**
   - Personal seleccionado (iniciales)
   - Puesto seleccionado
   - ¿Es realmente la segunda inspección o la primera?
   - Captura de pantalla del error

3. **Verificar datos en tabla:**
   - Ir a hoja "Historico"
   - Filtrar por iniciales del personal
   - Verificar que existen registros previos

**Fecha de última actualización**: 22/04/2026  
**Versión**: 1.1 (con mejoras de logging y validación)  
**Autor**: Sistema TH-HC-001 - Asurity
