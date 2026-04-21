# 📋 PLAN DE MEJORA CERTIFICADO PDF - MVP

**Proyecto:** TH-HC-001 EN PROCESO DE VALIDACIÓN  
**Fecha creación:** 20/04/2026  
**Estado:** PLANIFICACIÓN  
**Prioridad:** Media  
**Estimado total:** 60-90 minutos

---

## 🎯 OBJETIVO

Mejorar el certificado PDF actual con **5 cambios mínimos de alto impacto** que lo conviertan de un reporte técnico a un certificado profesional de competencia GMP, **sin romper funcionalidad existente**.

---

## ⚠️ RESTRICCIONES Y PRINCIPIOS

### ✅ **LO QUE SÍ HAREMOS**
- Modificar solo el diseño de la plantilla (hoja oculta "Plantilla Certificado")
- Ajustar la lógica de poblado en `CertificadoPDFGenerator.bas`
- Ajustar la inicialización en `PlantillaCertificadoSetup.bas`
- Mejorar el nombre de archivo generado

### ❌ **LO QUE NO TOCAREMOS**
- Cálculos de RPN, categorías, puntajes (ya funcionan)
- Flujo de generación del PDF (ExportAsFixedFormat)
- Tablas de datos (tblInspecciones, tblRespuestas, etc.)
- Formularios de captura (frmChecklistVirtual, frmSelectorInspeccion)
- Repositorios de datos (InspectionRepository, ChecklistRepository)

### 🔐 **VALIDACIÓN ANTES DE EMPEZAR**
```
☐ Hacer backup del archivo .xlsm actual
☐ Verificar que existe carpeta: c:\Propuestas Asurity\hoja-inspecciones\backups\
☐ Copiar archivo a: backups\TH-HC-001_PRE_MEJORA_CERT_20260420.xlsm
☐ Confirmar que puedo generar un certificado actual antes de modificar
```

---

## 📐 ARQUITECTURA DE CAMBIOS

```
┌─────────────────────────────────────────────────────────┐
│  PlantillaCertificadoSetup.bas                          │
│  ¿Qué cambia?                                           │
│  • Ajustar estructura de celdas (agregar bloque categoría)│
│  • Agregar formato condicional para incumplimientos     │
│  • Agregar celdas de validez en footer                  │
└─────────────────────────────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────────┐
│  CertificadoPDFGenerator.bas                            │
│  ¿Qué cambia?                                           │
│  • Poblar nuevas celdas (bloque categoría, validez)     │
│  • Aplicar colores según categoría                      │
│  • Resaltar incumplimientos en rojo                     │
│  • Calcular y mostrar fecha de validez                  │
│  • Mejorar nombre de archivo                            │
└─────────────────────────────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────────┐
│  HOJA: "Plantilla Certificado" (oculta)                │
│  ¿Qué cambia?                                           │
│  • Nuevas filas arriba para bloque categoría (3 filas)  │
│  • Todos los datos existentes se desplazan 3 filas abajo│
│  • Footer extendido con 2 filas más (validez)          │
└─────────────────────────────────────────────────────────┘
```

---

## 🚀 PLAN DE EJECUCIÓN (5 PASOS)

---

### **PASO 1: BLOQUE CATEGORÍA ARRIBA** ⭐⭐⭐⭐⭐
**Prioridad:** CRÍTICA  
**Estimado:** 20 minutos  
**Archivos:** `PlantillaCertificadoSetup.bas`, `CertificadoPDFGenerator.bas`

#### 📍 **Ubicación en plantilla**
```
ANTES (actual):
Fila 1-4: Logo + Título "CERTIFICADO DE INSPECCIÓN"
Fila 6+: Datos de inspección...

DESPUÉS (nuevo):
Fila 1-4: Logo + Título "CERTIFICADO DE INSPECCIÓN"
Fila 6-8: ← NUEVO BLOQUE CATEGORÍA (celdas combinadas A6:G8)
Fila 10+: Datos de inspección... (desplazados +3 filas)
```

#### 🔧 **Cambios en PlantillaCertificadoSetup.bas**

**Sub-tarea 1.1:** Insertar 3 filas nuevas después del título
```vb
' En InicializarPlantillaCertificado(), después de crear encabezado:

' === NUEVO: BLOQUE CATEGORÍA ===
ws.Rows("6:8").Insert Shift:=xlDown  ' Insertar 3 filas

' Combinar celdas A6:G8
With ws.Range("A6:G8")
    .Merge
    .Value = "[CATEGORÍA]"  ' Placeholder, se llenará dinámicamente
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Size = 18
    .Font.Bold = True
    .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
End With

' NOTA: Esto desplaza todo +3 filas, ajustar referencias posteriores
```

**Sub-tarea 1.2:** Ajustar todas las referencias de fila en el código
```
EJEMPLO:
Antes: ws.Cells(6, 1).Value = "DATOS DE INSPECCIÓN"
Después: ws.Cells(9, 1).Value = "DATOS DE INSPECCIÓN"  ' +3 filas

Buscar en PlantillaCertificadoSetup.bas:
- Todas las asignaciones ws.Cells(fila, col)
- Sumar 3 a cada número de fila >= 6
```

#### 🔧 **Cambios en CertificadoPDFGenerator.bas**

**Sub-tarea 1.3:** Poblar bloque categoría con datos reales
```vb
' En PoblarPlantillaCertificado(), después de poblar encabezado:

' === POBLAR BLOQUE CATEGORÍA ===
Dim textoCategoria As String
Dim nombreCategoria As String
Dim colorFondo As Long

' Obtener nombre categoría desde tblCategoriasRPN
nombreCategoria = ObtenerNombreCategoria(datosInspeccion("Categoria"))

' Construir texto
textoCategoria = "CATEGORÍA " & datosInspeccion("Categoria") & _
                 " - " & nombreCategoria & vbCrLf & _
                 "RPN: " & Format(datosInspeccion("RPN"), "0.00")

' Asignar a celda
wsPlantilla.Range("A6:G8").Value = textoCategoria

' Aplicar color según categoría
Select Case CLng(datosInspeccion("Categoria"))
    Case 1, 2
        colorFondo = RGB(212, 244, 230)  ' Verde claro
    Case 3
        colorFondo = RGB(254, 249, 219)  ' Amarillo claro
    Case 4, 5
        colorFondo = RGB(253, 223, 223)  ' Rojo claro
    Case Else
        colorFondo = RGB(255, 255, 255)  ' Blanco
End Select

wsPlantilla.Range("A6:G8").Interior.Color = colorFondo
```

**Sub-tarea 1.4:** Crear función auxiliar
```vb
' Nueva función en CertificadoPDFGenerator.bas

Private Function ObtenerNombreCategoria(ByVal numeroCategoria As Long) As String
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tbl = ws.ListObjects("tblCategoriasRPN")
    
    For Each fila In tbl.ListRows
        If fila.Range.Cells(1, tbl.ListColumns("Numero categoria").Index).Value = numeroCategoria Then
            ObtenerNombreCategoria = fila.Range.Cells(1, tbl.ListColumns("Nombre categoria").Index).Value
            Exit Function
        End If
    Next fila
    
    ObtenerNombreCategoria = "Sin categoría"
End Function
```

#### ✅ **Validación Paso 1**
```
☐ La plantilla tiene el bloque categoría en filas 6-8
☐ El texto se ve centrado y grande
☐ El color de fondo cambia según categoría (probar con Cat 1, 3, 5)
☐ Todos los datos inferiores se desplazaron correctamente (no se perdió nada)
☐ El PDF generado muestra la categoría arriba, visible
```

---

### **PASO 2: FILA "ESTADO" EN RESULTADOS** ⭐⭐⭐⭐
**Prioridad:** ALTA  
**Estimado:** 10 minutos  
**Archivos:** `CertificadoPDFGenerator.bas`

#### 📍 **Ubicación en plantilla**
```
SECCIÓN "RESULTADOS GENERALES":
Fila X: TA Puntaje obtenido
Fila X+1: TA Puntos no aplica
Fila X+2: TA Porcentaje
Fila X+3: RPN
Fila X+4: Categoría resultado
Fila X+5: ← NUEVA FILA "ESTADO"
```

#### 🔧 **Cambios en CertificadoPDFGenerator.bas**

**Sub-tarea 2.1:** Agregar campo "Estado" al poblar resultados
```vb
' En PoblarPlantillaCertificado(), después de poblar RPN y Categoría:

' === NUEVA FILA: ESTADO ===
Dim textoEstado As String
Dim iconoEstado As String

Select Case CLng(datosInspeccion("Categoria"))
    Case 1, 2
        iconoEstado = Chr(10004)  ' ✓
        textoEstado = iconoEstado & " COMPETENTE"
        wsPlantilla.Cells(filaEstado, 3).Font.Color = RGB(39, 174, 96)  ' Verde
    Case 3
        iconoEstado = Chr(9888)   ' ⚠
        textoEstado = iconoEstado & " COMPETENTE CON OBSERVACIONES"
        wsPlantilla.Cells(filaEstado, 3).Font.Color = RGB(243, 156, 18)  ' Naranja
    Case 4, 5
        iconoEstado = Chr(10006)  ' ✗
        textoEstado = "✗ NO CALIFICADO"
        wsPlantilla.Cells(filaEstado, 3).Font.Color = RGB(203, 67, 53)  ' Rojo
    Case Else
        textoEstado = "ESTADO INDETERMINADO"
End Select

wsPlantilla.Cells(filaEstado, 1).Value = "Estado:"  ' Columna A
wsPlantilla.Cells(filaEstado, 3).Value = textoEstado  ' Columna C
wsPlantilla.Cells(filaEstado, 1).Font.Bold = True
wsPlantilla.Cells(filaEstado, 3).Font.Bold = True
wsPlantilla.Cells(filaEstado, 3).Font.Size = 11
```

**NOTA:** `filaEstado` = fila donde está RPN + 2 (ajustar según estructura actual)

#### ✅ **Validación Paso 2**
```
☐ Aparece fila "Estado:" en sección Resultados
☐ Muestra "✓ COMPETENTE" para Cat 1-2 (verde)
☐ Muestra "⚠ COMPETENTE CON OBSERVACIONES" para Cat 3 (naranja)
☐ Muestra "✗ NO CALIFICADO" para Cat 4-5 (rojo)
☐ El texto es legible y destacado
```

---

### **PASO 3: RESALTAR INCUMPLIMIENTOS EN ROJO** ⭐⭐⭐⭐
**Prioridad:** ALTA  
**Estimado:** 20 minutos  
**Archivos:** `CertificadoPDFGenerator.bas`

#### 📍 **Ubicación en plantilla**
```
SECCIONES DE PREGUNTAS:
Cada fila de respuesta que tenga:
- "No" (en Técnica Aséptica)
- "No Cumple" (en Auditoría Procesos)

→ Fondo rojo claro + texto en negrita
```

#### 🔧 **Cambios en CertificadoPDFGenerator.bas**

**Sub-tarea 3.1:** Modificar bucle de poblado de respuestas
```vb
' En PoblarSeccionRespuestas() o equivalente:

For Each respuesta In respuestas
    ' ... código existente para poblar pregunta, respuesta, observación ...
    
    ' === NUEVO: DETECTAR Y RESALTAR INCUMPLIMIENTOS ===
    Dim textoRespuesta As String
    textoRespuesta = Trim(CStr(respuesta("Respuesta")))
    
    If textoRespuesta = "No" Or textoRespuesta = "No Cumple" Then
        ' Aplicar fondo rojo claro a toda la fila
        With wsPlantilla.Rows(filaActual).Interior
            .Color = RGB(255, 200, 200)  ' Rojo suave
            .Pattern = xlSolid
        End With
        
        ' Poner respuesta en negrita
        wsPlantilla.Cells(filaActual, colRespuesta).Font.Bold = True
        
        ' Opcional: agregar ícono
        wsPlantilla.Cells(filaActual, colRespuesta).Value = "❌ " & textoRespuesta
    End If
    
    filaActual = filaActual + 1
Next respuesta
```

**NOTA:** Ajustar nombres de variables según código actual (`filaActual`, `colRespuesta`)

#### ✅ **Validación Paso 3**
```
☐ Respuestas "No" tienen fondo rojo claro
☐ Respuestas "No Cumple" tienen fondo rojo claro
☐ El texto en esas celdas está en negrita
☐ El resto de respuestas ("Sí", "Cumple", "No Aplica") NO tienen color
☐ La tabla sigue siendo legible
```

---

### **PASO 4: FOOTER CON VALIDEZ** ⭐⭐⭐
**Prioridad:** MEDIA  
**Estimado:** 15 minutos  
**Archivos:** `PlantillaCertificadoSetup.bas`, `CertificadoPDFGenerator.bas`

#### 📍 **Ubicación en plantilla**
```
FOOTER (final del certificado):
Fila N: Fecha emisión certificado
Fila N+1: ID Inspección
Fila N+2: ← NUEVO: VÁLIDO HASTA
Fila N+3: ← NUEVO: Próxima inspección requerida
Fila N+4: Generado por: Sistema...
```

#### 🔧 **Cambios en PlantillaCertificadoSetup.bas**

**Sub-tarea 4.1:** Agregar líneas de validez al template
```vb
' En InicializarPlantillaCertificado(), sección footer:

' Líneas existentes...
ws.Cells(filaFooter, 1).Value = "Fecha emisión certificado:"
ws.Cells(filaFooter + 1, 1).Value = "ID Inspección:"

' === NUEVAS LÍNEAS ===
ws.Cells(filaFooter + 2, 1).Value = "VÁLIDO HASTA:"
ws.Cells(filaFooter + 2, 1).Font.Bold = True
ws.Cells(filaFooter + 3, 1).Value = "Próxima inspección requerida antes de esta fecha"
ws.Cells(filaFooter + 3, 1).Font.Size = 8
ws.Cells(filaFooter + 3, 1).Font.Italic = True

ws.Cells(filaFooter + 4, 1).Value = "Generado por: Sistema TH-HC-001 v1.0"
```

#### 🔧 **Cambios en CertificadoPDFGenerator.bas**

**Sub-tarea 4.2:** Calcular y poblar fecha de validez
```vb
' En PoblarPlantillaCertificado(), sección footer:

' === CALCULAR VALIDEZ ===
Dim fechaInspeccion As Date
Dim frecuenciaMeses As Long
Dim fechaValidez As Date

fechaInspeccion = CDate(datosInspeccion("FechaInspeccion"))

' Obtener frecuencia de la plantilla (default: 3 meses)
frecuenciaMeses = ObtenerFrecuenciaPlantilla(datosInspeccion("IDPlantilla"))
If frecuenciaMeses = 0 Then frecuenciaMeses = 3  ' Default

' Calcular fecha de validez
fechaValidez = DateAdd("m", frecuenciaMeses, fechaInspeccion)

' Poblar
wsPlantilla.Cells(filaFooter + 2, 3).Value = Format(fechaValidez, "DD/MM/YYYY") & _
                                             " (" & frecuenciaMeses & " meses)"
wsPlantilla.Cells(filaFooter + 2, 3).Font.Bold = True
wsPlantilla.Cells(filaFooter + 2, 3).Font.Size = 10

' Opcional: alertar si ya venció
If Date > fechaValidez Then
    wsPlantilla.Cells(filaFooter + 2, 3).Font.Color = RGB(203, 67, 53)  ' Rojo
    wsPlantilla.Cells(filaFooter + 2, 3).Value = _
        wsPlantilla.Cells(filaFooter + 2, 3).Value & " ⚠ VENCIDO"
End If
```

**Sub-tarea 4.3:** Crear función auxiliar
```vb
' Nueva función en CertificadoPDFGenerator.bas

Private Function ObtenerFrecuenciaPlantilla(ByVal idPlantilla As String) As Long
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set tbl = ws.ListObjects("tblPlantillas")
    
    For Each fila In tbl.ListRows
        If fila.Range.Cells(1, tbl.ListColumns("ID Plantilla").Index).Value = idPlantilla Then
            ObtenerFrecuenciaPlantilla = CLng(fila.Range.Cells(1, _
                tbl.ListColumns("Frecuencia meses").Index).Value)
            Exit Function
        End If
    Next fila
    
    ObtenerFrecuenciaPlantilla = 0  ' No encontrado
End Function
```

#### ✅ **Validación Paso 4**
```
☐ Aparece "VÁLIDO HASTA:" con fecha calculada
☐ La fecha es correcta (fecha inspección + frecuencia meses)
☐ Muestra el número de meses entre paréntesis
☐ Si el certificado está vencido, aparece "⚠ VENCIDO" en rojo
```

---

### **PASO 5: NOMBRE DE ARCHIVO INTELIGENTE** ⭐⭐⭐
**Prioridad:** MEDIA  
**Estimado:** 10 minutos  
**Archivos:** `CertificadoPDFGenerator.bas`

#### 📍 **Ubicación en código**
```
Función: GenerarNombreArchivoPDF()
Ubicación: CertificadoPDFGenerator.bas
```

#### 🔧 **Cambios en CertificadoPDFGenerator.bas**

**Sub-tarea 5.1:** Modificar formato de nombre
```vb
' ANTES (actual):
' CERTIFICADO_ABC_20260420_x8fL.pdf

' DESPUÉS (mejorado):
' CERTIFICADO_ACF_OperadorNPT_2026-04-20_CAT1_APROBADO.pdf

Private Function GenerarNombreArchivoPDF(ByVal datosInspeccion As Object) As String
    Dim iniciales As String
    Dim puesto As String
    Dim fecha As String
    Dim categoria As String
    Dim estado As String
    Dim uuid4 As String
    
    ' Iniciales
    iniciales = Trim(CStr(datosInspeccion("Personal")))
    
    ' Puesto (sin espacios, max 20 chars)
    puesto = Replace(Trim(CStr(datosInspeccion("Puesto"))), " ", "")
    If Len(puesto) > 20 Then puesto = Left(puesto, 20)
    
    ' Fecha formato ISO
    fecha = Format(datosInspeccion("FechaInspeccion"), "YYYY-MM-DD")
    
    ' Categoría
    categoria = "CAT" & Trim(CStr(datosInspeccion("Categoria")))
    
    ' Estado según categoría
    Select Case CLng(datosInspeccion("Categoria"))
        Case 1, 2
            estado = "APROBADO"
        Case 3
            estado = "OBSERVACION"
        Case 4, 5
            estado = "NOCALIFICADO"
        Case Else
            estado = "INDEFINIDO"
    End Select
    
    ' UUID corto (últimos 4 chars para unicidad)
    uuid4 = Right(Replace(datosInspeccion("IDInspeccion"), "-", ""), 4)
    
    ' Construir nombre
    GenerarNombreArchivoPDF = "CERTIFICADO_" & _
                              iniciales & "_" & _
                              puesto & "_" & _
                              fecha & "_" & _
                              categoria & "_" & _
                              estado & "_" & _
                              uuid4 & ".pdf"
End Function
```

#### ✅ **Validación Paso 5**
```
☐ El archivo PDF se guarda con el nuevo nombre
☐ Incluye: Iniciales, Puesto, Fecha ISO, Categoría, Estado, UUID
☐ Ejemplo: CERTIFICADO_ACF_OperadorNPT_2026-04-20_CAT1_APROBADO_x8fL.pdf
☐ El nombre no tiene caracteres inválidos (espacios, /, \, etc.)
☐ El archivo se puede ubicar fácilmente en Windows Explorer
```

---

## 🧪 PLAN DE PRUEBAS INTEGRAL

### **Casos de prueba por categoría**

| # | Escenario | Categoría | RPN | Qué validar |
|---|-----------|-----------|-----|-------------|
| 1 | Desempeño óptimo | 1 | 10 | Bloque verde, estado "✓ COMPETENTE" |
| 2 | Aceptable | 2 | 17 | Bloque verde, estado "✓ COMPETENTE" |
| 3 | Con observaciones | 3 | 25 | Bloque amarillo, estado "⚠ CON OBSERVACIONES" |
| 4 | Crítico | 4 | 50 | Bloque rojo, estado "✗ NO CALIFICADO" |
| 5 | Recurrente | 5 | 30 | Bloque rojo, estado "✗ NO CALIFICADO" |

### **Validación de incumplimientos**

| # | Sección | Respuestas | Qué validar |
|---|---------|------------|-------------|
| 6 | Auditoría Procesos | 8 Cumple, 2 No Cumple | 2 filas con fondo rojo |
| 7 | Técnica Aséptica | 10 Sí, 3 No, 2 No Aplica | 3 filas con fondo rojo |
| 8 | Todo cumple | 100% Sí/Cumple | Ninguna fila resaltada |

### **Validación de fechas**

| # | Fecha inspección | Frecuencia | Fecha validez esperada |
|---|------------------|------------|------------------------|
| 9 | 20/04/2026 | 3 meses | 20/07/2026 |
| 10 | 15/01/2026 | 6 meses | 15/07/2026 |
| 11 | 01/12/2025 (vencido) | 3 meses | 01/03/2026 ⚠ VENCIDO |

---

## 📦 CHECKLIST FINAL DE ENTREGA

### **Antes de considerar completo:**

```
☐ BACKUP CREADO
  ☐ Archivo respaldado en: backups\TH-HC-001_PRE_MEJORA_CERT_20260420.xlsm

☐ PASO 1: BLOQUE CATEGORÍA
  ☐ Implementado en PlantillaCertificadoSetup.bas
  ☐ Implementado en CertificadoPDFGenerator.bas
  ☐ Probado con Cat 1, 3, 5
  ☐ Colores correctos

☐ PASO 2: FILA ESTADO
  ☐ Implementado en CertificadoPDFGenerator.bas
  ☐ Probado con Cat 1, 3, 5
  ☐ Textos e íconos correctos

☐ PASO 3: RESALTAR INCUMPLIMIENTOS
  ☐ Implementado en CertificadoPDFGenerator.bas
  ☐ Probado con inspección que tiene "No" y "No Cumple"
  ☐ Solo se resaltan incumplimientos, no el resto

☐ PASO 4: FOOTER VALIDEZ
  ☐ Implementado en PlantillaCertificadoSetup.bas
  ☐ Implementado en CertificadoPDFGenerator.bas
  ☐ Fecha calculada correctamente
  ☐ Detecta vencimiento

☐ PASO 5: NOMBRE ARCHIVO
  ☐ Implementado en CertificadoPDFGenerator.bas
  ☐ Nombre incluye todos los elementos
  ☐ Nombre es válido en Windows

☐ REGRESIÓN - NO SE ROMPIÓ NADA
  ☐ El formulario de inspección sigue funcionando
  ☐ Se pueden completar y guardar inspecciones
  ☐ Los cálculos de RPN son correctos
  ☐ Los certificados antiguos no se afectan
  ☐ La plantilla se limpia correctamente después de generar PDF

☐ DOCUMENTACIÓN
  ☐ Actualizar TODO_CERTIFICADO_PDF.md con cambios realizados
  ☐ Marcar este plan como COMPLETADO
```

---

## 🔄 PLAN DE ROLLBACK (SI ALGO SALE MAL)

### **Si un paso falla:**

1. **No entrar en pánico** - tienes backup
2. **Identificar qué paso falló** (usar checklist arriba)
3. **Restaurar desde backup:**
   ```
   1. Cerrar TH-HC-001 INSPECCIONES.xlsm
   2. Ir a: backups\TH-HC-001_PRE_MEJORA_CERT_20260420.xlsm
   3. Copiar a carpeta principal
   4. Renombrar a: TH-HC-001 INSPECCIONES.xlsm
   5. Abrir y verificar que funciona
   ```
4. **Revisar error:**
   - ¿Qué línea de código falló?
   - ¿Hay error de sintaxis?
   - ¿Se olvidó ajustar alguna referencia de fila?
5. **Reintentar el paso específico**

### **Si el PDF no se genera:**

```
Checklist de diagnóstico:
☐ ¿La plantilla existe y está oculta? (Visible = xlSheetVeryHidden)
☐ ¿Las celdas tienen contenido antes de exportar?
☐ ¿La ruta del Desktop existe?
☐ ¿Hay permisos de escritura en Desktop?
☐ ¿El nombre de archivo es válido (sin caracteres raros)?
☐ Revisar Debug.Print en Immediate Window
```

---

## 📊 MÉTRICAS DE ÉXITO

Al finalizar este plan, el certificado debe cumplir:

| Métrica | Objetivo | Cómo medir |
|---------|----------|------------|
| Tiempo de comprensión | < 10 segundos | Mostrar PDF a alguien que no conoce el sistema y preguntar: "¿Aprobó o no?" |
| Identificación visual | 100% | Categoría visible sin necesidad de leer texto |
| Detección de problemas | Inmediata | Incumplimientos resaltados en rojo |
| Trazabilidad | Completa | Nombre de archivo + ID inspección + fecha validez |
| Profesionalismo | ⭐⭐⭐⭐⭐ | Aprobación del equipo de Calidad/QA |

---

## 📅 CRONOGRAMA SUGERIDO

| Día | Actividad | Duración |
|-----|-----------|----------|
| **Día 1** | Backup + Paso 1 (Bloque categoría) | 30 min |
| **Día 1** | Paso 2 (Fila estado) + Pruebas Cat 1-5 | 20 min |
| **Día 2** | Paso 3 (Resaltar incumplimientos) | 20 min |
| **Día 2** | Paso 4 (Footer validez) | 15 min |
| **Día 2** | Paso 5 (Nombre archivo) | 10 min |
| **Día 3** | Pruebas exhaustivas + Ajustes finales | 30 min |
| **Día 3** | Documentación + Cierre | 15 min |

**TOTAL:** ~2.5 horas distribuidas en 3 sesiones

---

## ✅ PRÓXIMOS PASOS DESPUÉS DEL MVP

Una vez completado este MVP, **si se desea mejorar aún más**, considerar para versión 2.0:

1. **QR Code** con link a verificación online
2. **Gráfico de tendencia** del evaluado (últimas 5 inspecciones)
3. **Comparativa con peers** (anónima)
4. **Checklist de cierre CAPA** para Cat 4-5
5. **Firmas digitales** en vez de líneas manuales

Pero **NO HACERLO AHORA** - primero validar que este MVP funciona y aporta valor.

---

## 📞 SOPORTE

**Si algo no funciona:**
1. Revisar Debug.Print en Immediate Window (Ctrl+G)
2. Verificar que todas las referencias de tabla existen
3. Consultar este plan paso a paso
4. Restaurar backup si es necesario

---

**FIN DEL PLAN**

---

**Estado:** ⏳ PENDIENTE DE EJECUCIÓN  
**Aprobado por:** [Pendiente]  
**Fecha inicio:** [Pendiente]  
**Fecha finalización:** [Pendiente]
