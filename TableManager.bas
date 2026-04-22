' ======================================================================
' Módulo: TableManager
' Descripción: CRUD centralizado para las 5 tablas maestras del sistema
'              de inspecciones. Todas las operaciones de lectura/escritura
'              sobre ListObject pasan por este módulo.
'
' Tablas gestionadas:
'   - tblCriticidad (Hoja: Configuración)
'   - tblSecciones (Hoja: Configuración)
'   - tblOpcionesDeRespuesta (Hoja: Configuración)
'   - tblPlantillas (Hoja: Checklist)
'   - tblPreguntas (Hoja: Checklist)
'
' Dependencias: Configuration2.bas, AuditLogger2.bas, SheetProtector2.bas
' ======================================================================
Option Explicit

' ======================================================================
' SECCIÓN 1: RESOLUCIÓN DE TABLA (convierte nombre lógico → ListObject)
' ======================================================================

' ----------------------------------------------------------------------
' ObtenerListObject
' Propósito: Dado un nombre lógico ("CRITICIDAD", "SECCIONES", etc.),
'            devuelve el ListObject correspondiente.
' Retorna: El ListObject, o Nothing si no se encuentra.
' ----------------------------------------------------------------------
Public Function ObtenerListObject(ByVal nombreLogico As String) As ListObject
    Dim ws As Worksheet
    Dim tableName As String
    
    Select Case UCase(nombreLogico)
        Case "CRITICIDAD"
            Set ws = ThisWorkbook.Worksheets(Configuration2.SHEET_CHECKLIST)
            tableName = Configuration2.TABLE_CRITICIDAD
        Case "SECCIONES"
            Set ws = ThisWorkbook.Worksheets(Configuration2.SHEET_CHECKLIST)
            tableName = Configuration2.TABLE_SECCIONES
        Case "OPCIONES"
            Set ws = ThisWorkbook.Worksheets(Configuration2.SHEET_CHECKLIST)
            tableName = Configuration2.TABLE_OPCIONES
        Case "PLANTILLAS"
            Set ws = ThisWorkbook.Worksheets(Configuration2.SHEET_CHECKLIST)
            tableName = Configuration2.TABLE_PLANTILLAS
        Case "PREGUNTAS"
            Set ws = ThisWorkbook.Worksheets(Configuration2.SHEET_CHECKLIST)
            tableName = Configuration2.TABLE_PREGUNTAS
        Case Else
            Set ObtenerListObject = Nothing
            Exit Function
    End Select
    
    On Error Resume Next
    Set ObtenerListObject = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

' ======================================================================
' SECCIÓN 2: LECTURA DE DATOS
' ======================================================================

' ----------------------------------------------------------------------
' ObtenerDatosTabla
' Propósito: Devuelve TODAS las filas de una tabla como array 2D.
'            Cada fila contiene los valores de todas las columnas.
' Parámetro: nombreLogico - "CRITICIDAD", "SECCIONES", etc.
' Retorna: Variant (array 2D base 1), o Empty si la tabla está vacía.
' ----------------------------------------------------------------------
Public Function ObtenerDatosTabla(ByVal nombreLogico As String) As Variant
    Dim tbl As ListObject
    Set tbl = ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Then
        ObtenerDatosTabla = Empty
        Exit Function
    End If
    
    If tbl.ListRows.Count = 0 Then
        ObtenerDatosTabla = Empty
        Exit Function
    End If
    
    ' Retornar el rango de datos como array
    ObtenerDatosTabla = tbl.DataBodyRange.Value
End Function

' ----------------------------------------------------------------------
' ObtenerFilaPorIndice
' Propósito: Devuelve una sola fila (1D) de la tabla, por índice (0-based
'            relativo al ListBox del formulario).
' Retorna: Variant (array 1D base 1), o Empty si no existe.
' ----------------------------------------------------------------------
Public Function ObtenerFilaPorIndice(ByVal nombreLogico As String, _
                                      ByVal indice As Long) As Variant
    Dim tbl As ListObject
    Set tbl = ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Then
        ObtenerFilaPorIndice = Empty
        Exit Function
    End If
    
    ' indice es 0-based (del ListBox), convertir a 1-based del ListObject
    Dim filaListObject As Long
    filaListObject = indice + 1
    
    If filaListObject < 1 Or filaListObject > tbl.ListRows.Count Then
        ObtenerFilaPorIndice = Empty
        Exit Function
    End If
    
    Dim numCols As Long
    numCols = tbl.ListColumns.Count
    
    Dim resultado() As Variant
    ReDim resultado(1 To numCols)
    
    Dim col As Long
    For col = 1 To numCols
        resultado(col) = tbl.ListRows(filaListObject).Range.Cells(1, col).Value
    Next col
    
    ObtenerFilaPorIndice = resultado
End Function

' ----------------------------------------------------------------------
' ObtenerItemsLookup
' Propósito: Para un ComboBox lookup, devuelve array 2D con (ID, Nombre)
'            de la tabla indicada. Solo registros activos.
' Retorna: Variant (array 2D), o Empty si vacía.
' ----------------------------------------------------------------------
Public Function ObtenerItemsLookup(ByVal nombreLogico As String) As Variant
    Dim tbl As ListObject
    Set tbl = ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
        ObtenerItemsLookup = Empty
        Exit Function
    End If
    
    ' Determinar columnas de ID y Nombre según tabla
    Dim colID As Long, colNombre As Long
    colID = 1  ' Todas tienen ID en columna 1
    colNombre = 2  ' Todas tienen nombre en columna 2
    
    ' Para preguntas, filtrar por Activo
    Dim tieneActivo As Boolean
    Dim colActivo As Long
    tieneActivo = False
    
    If UCase(nombreLogico) = "PREGUNTAS" Then
        tieneActivo = True
        colActivo = ObtenerIndiceColumna(tbl, "Activo")
    End If
    
    ' Contar filas activas
    Dim totalActivos As Long
    totalActivos = 0
    Dim fila As Long
    
    For fila = 1 To tbl.ListRows.Count
        If tieneActivo Then
            Dim valActivo As String
            valActivo = CStr(tbl.ListRows(fila).Range.Cells(1, colActivo).Value)
            If valActivo = "Sí" Or valActivo = "Si" Or valActivo = "True" Then
                totalActivos = totalActivos + 1
            End If
        Else
            totalActivos = totalActivos + 1
        End If
    Next fila
    
    If totalActivos = 0 Then
        ObtenerItemsLookup = Empty
        Exit Function
    End If
    
    ' Construir array resultado
    Dim resultado() As Variant
    ReDim resultado(1 To totalActivos, 1 To 2)
    
    Dim idx As Long: idx = 1
    For fila = 1 To tbl.ListRows.Count
        Dim incluir As Boolean: incluir = True
        
        If tieneActivo Then
            Dim val2 As String
            val2 = CStr(tbl.ListRows(fila).Range.Cells(1, colActivo).Value)
            incluir = (val2 = "Sí" Or val2 = "Si" Or val2 = "True")
        End If
        
        If incluir Then
            resultado(idx, 1) = tbl.ListRows(fila).Range.Cells(1, colID).Value
            resultado(idx, 2) = tbl.ListRows(fila).Range.Cells(1, colNombre).Value
            idx = idx + 1
        End If
    Next fila
    
    ObtenerItemsLookup = resultado
End Function

' ----------------------------------------------------------------------
' ObtenerNombrePorID
' Propósito: Busca el nombre (columna 2) basado en el ID (columna 1).
'            Útil para mostrar valores legibles en lookups.
' Parámetros:
'   nombreLogico - tabla origen ("CRITICIDAD", "SECCIONES", etc.)
'   idBuscado - ID a buscar
' Retorna: String con el nombre, o texto "ID no encontrado" si no existe.
' ----------------------------------------------------------------------
Public Function ObtenerNombrePorID(ByVal nombreLogico As String, _
                                   ByVal idBuscado As String) As String
    Dim tbl As ListObject
    Set tbl = ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
        ObtenerNombrePorID = "[Vacío]"
        Exit Function
    End If
    
    Dim fila As Long
    For fila = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(fila).Range.Cells(1, 1).Value) = CStr(idBuscado) Then
            ObtenerNombrePorID = CStr(tbl.ListRows(fila).Range.Cells(1, 2).Value)
            Exit Function
        End If
    Next fila
    
    ObtenerNombrePorID = "[No encontrado]"
End Function

' ======================================================================
' SECCIÓN 3: ESCRITURA DE DATOS (CRUD)
' ======================================================================

' ----------------------------------------------------------------------
' GenerarNuevoID
' Propósito: Genera un ID único para una nueva fila.
'            Formato: Prefijo + timestamp para garantizar unicidad.
' Parámetro: nombreLogico - tabla destino
' Retorna: String con el nuevo ID.
' ----------------------------------------------------------------------
Public Function GenerarNuevoID(ByVal nombreLogico As String) As String
    Dim prefijo As String
    
    Select Case UCase(nombreLogico)
        Case "CRITICIDAD":  prefijo = "CRT"
        Case "SECCIONES":   prefijo = "SEC"
        Case "PLANTILLAS":  prefijo = "PLT"
        Case "OPCIONES":    prefijo = "OPC"
        Case "PREGUNTAS":   prefijo = "PRG"
        Case Else:          prefijo = "GEN"
    End Select
    
    ' Formato: PRG-20260414-153045-123
    GenerarNuevoID = prefijo & "-" & Format(Now, "yyyymmdd-hhnnss") & "-" & _
                     Right("000" & Int(Rnd * 999), 3)
End Function

' ----------------------------------------------------------------------
' InsertarFila
' Propósito: Agrega una nueva fila al final de la tabla con los datos
'            del Dictionary proporcionado.
' Parámetros:
'   nombreLogico - nombre lógico de la tabla
'   datos - Dictionary con pares clave/valor
' Retorna: True si se insertó correctamente.
' ----------------------------------------------------------------------
Public Function InsertarFila(ByVal nombreLogico As String, _
                              ByVal datos As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim tbl As ListObject
    Set tbl = ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Then
        InsertarFila = False
        Exit Function
    End If
    
    ' Desproteger hoja temporalmente
    Dim ws As Worksheet
    Set ws = tbl.Parent
    Call SheetProtector2.UnprotectSheet(ws, Configuration2.APP_PASSWORD)
    
    ' Agregar nueva fila
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    
    ' Mapear datos del Dictionary a columnas de la tabla
    Call MapearDatosAFila(nombreLogico, newRow, datos)
    
    ' Re-proteger hoja
    If Configuration2.ENABLE_SHEET_PROTECTION Then
        Call SheetProtector2.ProtectSheet(ws, Configuration2.APP_PASSWORD)
    End If
    
    InsertarFila = True
    Exit Function
    
ErrorHandler:
    InsertarFila = False
    Call ErrorLogger2.Log("TableManager.InsertarFila", Err.Description, Err.Number)
End Function

' ----------------------------------------------------------------------
' ActualizarFila
' Propósito: Actualiza una fila existente con los datos del Dictionary.
' Parámetros:
'   nombreLogico - nombre lógico de la tabla
'   indice - índice 0-based (del ListBox)
'   datos - Dictionary con pares clave/valor
' Retorna: True si se actualizó correctamente.
' ----------------------------------------------------------------------
Public Function ActualizarFila(ByVal nombreLogico As String, _
                                ByVal indice As Long, _
                                ByVal datos As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim tbl As ListObject
    Set tbl = ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Then
        ActualizarFila = False
        Exit Function
    End If
    
    Dim filaIdx As Long
    filaIdx = indice + 1  ' Convertir 0-based a 1-based
    
    If filaIdx < 1 Or filaIdx > tbl.ListRows.Count Then
        ActualizarFila = False
        Exit Function
    End If
    
    ' Desproteger hoja temporalmente
    Dim ws As Worksheet
    Set ws = tbl.Parent
    Call SheetProtector2.UnprotectSheet(ws, Configuration2.APP_PASSWORD)
    
    ' Mapear datos a fila existente
    Call MapearDatosAFila(nombreLogico, tbl.ListRows(filaIdx), datos)
    
    ' Re-proteger
    If Configuration2.ENABLE_SHEET_PROTECTION Then
        Call SheetProtector2.ProtectSheet(ws, Configuration2.APP_PASSWORD)
    End If
    
    ActualizarFila = True
    Exit Function
    
ErrorHandler:
    ActualizarFila = False
    Call ErrorLogger2.Log("TableManager.ActualizarFila", Err.Description, Err.Number)
End Function

' ----------------------------------------------------------------------
' EliminarFila
' Propósito: Elimina físicamente una fila de la tabla.
'            SOLO se usa cuando NO tiene dependencias.
' Parámetros:
'   nombreLogico - nombre lógico de la tabla
'   indice - índice 0-based
' Retorna: True si se eliminó correctamente.
' REGLA: Preferir MarcarInactivo sobre EliminarFila siempre que sea posible.
' ----------------------------------------------------------------------
Public Function EliminarFila(ByVal nombreLogico As String, _
                              ByVal indice As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim tbl As ListObject
    Set tbl = ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Then
        EliminarFila = False
        Exit Function
    End If
    
    Dim filaIdx As Long
    filaIdx = indice + 1
    
    If filaIdx < 1 Or filaIdx > tbl.ListRows.Count Then
        EliminarFila = False
        Exit Function
    End If
    
    ' Desproteger
    Dim ws As Worksheet
    Set ws = tbl.Parent
    Call SheetProtector2.UnprotectSheet(ws, Configuration2.APP_PASSWORD)
    
    ' Eliminar fila
    tbl.ListRows(filaIdx).Delete
    
    ' Re-proteger
    If Configuration2.ENABLE_SHEET_PROTECTION Then
        Call SheetProtector2.ProtectSheet(ws, Configuration2.APP_PASSWORD)
    End If
    
    EliminarFila = True
    Exit Function
    
ErrorHandler:
    EliminarFila = False
    Call ErrorLogger2.Log("TableManager.EliminarFila", Err.Description, Err.Number)
End Function

' ----------------------------------------------------------------------
' MarcarInactivo
' Propósito: Soft-delete. Cambia la columna "Activo" a "No".
'            Para tablas sin columna Activo, no hace nada.
' Parámetros:
'   nombreLogico - nombre lógico de la tabla
'   indice - índice 0-based
' Retorna: True si se marcó correctamente.
' ----------------------------------------------------------------------
Public Function MarcarInactivo(ByVal nombreLogico As String, _
                                ByVal indice As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim tbl As ListObject
    Set tbl = ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Then
        MarcarInactivo = False
        Exit Function
    End If
    
    ' Buscar columna "Activo"
    Dim colActivo As Long
    colActivo = ObtenerIndiceColumna(tbl, "Activo")
    
    If colActivo = 0 Then
        ' La tabla no tiene columna "Activo", no se puede hacer soft-delete
        MarcarInactivo = False
        Exit Function
    End If
    
    Dim filaIdx As Long
    filaIdx = indice + 1
    
    If filaIdx < 1 Or filaIdx > tbl.ListRows.Count Then
        MarcarInactivo = False
        Exit Function
    End If
    
    ' Desproteger
    Dim ws As Worksheet
    Set ws = tbl.Parent
    Call SheetProtector2.UnprotectSheet(ws, Configuration2.APP_PASSWORD)
    
    ' Marcar como inactivo
    tbl.ListRows(filaIdx).Range.Cells(1, colActivo).Value = "No"
    
    ' Re-proteger
    If Configuration2.ENABLE_SHEET_PROTECTION Then
        Call SheetProtector2.ProtectSheet(ws, Configuration2.APP_PASSWORD)
    End If
    
    MarcarInactivo = True
    Exit Function
    
ErrorHandler:
    MarcarInactivo = False
    Call ErrorLogger2.Log("TableManager.MarcarInactivo", Err.Description, Err.Number)
End Function

' ======================================================================
' SECCIÓN 4: MAPEO DE DATOS (Dictionary → Fila de tabla)
' ======================================================================

' ----------------------------------------------------------------------
' MapearDatosAFila
' Propósito: Escribe los valores del Dictionary en las celdas de la fila,
'            mapeando cada clave del dictionary a la columna correcta.
' ----------------------------------------------------------------------
Private Sub MapearDatosAFila(ByVal nombreLogico As String, _
                              ByVal fila As ListRow, _
                              ByVal datos As Object)
    Dim tbl As ListObject
    Set tbl = fila.Parent
    
    Select Case UCase(nombreLogico)
        Case "CRITICIDAD"
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "ID Criticidad")).Value = datos("ID")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Nombre de criticidad")).Value = datos("Nombre")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Valor")).Value = datos("Valor")
            
        Case "SECCIONES"
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "ID Seccion")).Value = datos("ID")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Nombre de sección")).Value = datos("Nombre")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Tipo de respuesta")).Value = datos("TipoRespuesta")
            
        Case "PLANTILLAS"
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "ID Plantilla")).Value = datos("ID")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Nombre de plantilla")).Value = datos("Nombre")
            ' Columna Área (antes Etapa) - usado para precarga de cboArea en frmChecklistVirtual
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Área")).Value = datos("Etapa")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Puesto")).Value = datos("Puesto")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Frecuencia meses")).Value = datos("Frecuencia")
            
        Case "OPCIONES"
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "ID Opcion")).Value = datos("ID")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "ID Seccion")).Value = datos("IDSeccion")
            If ObtenerIndiceColumna(tbl, "ID Criticidad") > 0 Then
                fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "ID Criticidad")).Value = datos("IDCriticidad")
            End If
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Opción texto")).Value = datos("TextoOpcion")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Valor puntaje")).Value = datos("ValorPuntaje")
            
        Case "PREGUNTAS"
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "ID Plantilla")).Value = datos("IDPlantilla")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "ID Pregunta")).Value = datos("IDPregunta")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "ID Sección")).Value = datos("IDSeccion")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "ID Criticidad")).Value = datos("IDCriticidad")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Texto pregunta")).Value = datos("TextoPregunta")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Orden")).Value = datos("Orden")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Activo")).Value = datos("Activo")
            fila.Range.Cells(1, ObtenerIndiceColumna(tbl, "Observaciones")).Value = datos("Observaciones")
            
            ' Fecha creación solo en modo NUEVO
            Dim colFecha As Long
            colFecha = ObtenerIndiceColumna(tbl, "Fecha creación")
            If colFecha > 0 Then
                If fila.Range.Cells(1, colFecha).Value = "" Or IsEmpty(fila.Range.Cells(1, colFecha).Value) Then
                    fila.Range.Cells(1, colFecha).Value = Now
                End If
            End If
    End Select
End Sub

' ======================================================================
' SECCIÓN 5: UTILIDADES
' ======================================================================

' ----------------------------------------------------------------------
' ObtenerIndiceColumna
' Propósito: Busca una columna por nombre en un ListObject.
' Retorna: Índice de la columna (1-based), o 0 si no existe.
' ----------------------------------------------------------------------
Public Function ObtenerIndiceColumna(ByVal tbl As ListObject, _
                                      ByVal nombreColumna As String) As Long
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If LCase(Trim(col.Name)) = LCase(Trim(nombreColumna)) Then
            ObtenerIndiceColumna = col.Index
            Exit Function
        End If
    Next col
    ObtenerIndiceColumna = 0
End Function

' ----------------------------------------------------------------------
' BuscarFilaPorID
' Propósito: Busca una fila que tenga un ID específico en su primera columna.
' Retorna: Índice de ListRow (1-based), o 0 si no se encuentra.
' ----------------------------------------------------------------------
Public Function BuscarFilaPorID(ByVal nombreLogico As String, _
                                 ByVal idBuscado As String) As Long
    Dim tbl As ListObject
    Set tbl = ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
        BuscarFilaPorID = 0
        Exit Function
    End If
    
    Dim fila As Long
    For fila = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(fila).Range.Cells(1, 1).Value) = idBuscado Then
            BuscarFilaPorID = fila
            Exit Function
        End If
    Next fila
    
    BuscarFilaPorID = 0
End Function

' ----------------------------------------------------------------------
' ContarRegistros
' Propósito: Cuenta las filas de una tabla (útil para estadísticas).
' Retorna: Número de filas con datos.
' ----------------------------------------------------------------------
Public Function ContarRegistros(ByVal nombreLogico As String) As Long
    Dim tbl As ListObject
    Set tbl = ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Then
        ContarRegistros = 0
    Else
        ContarRegistros = tbl.ListRows.Count
    End If
End Function

' ======================================================================
' SECCIÓN 6: INTEGRACIÓN CON AUDIT TRAIL
' ======================================================================

' ----------------------------------------------------------------------
' RegistrarCambioAudit
' Propósito: Registra cada operación CRUD en el sistema de Audit Trail.
' Parámetros:
'   nombreLogico - tabla afectada
'   operacion - "NUEVO", "EDITAR", "ELIMINAR"
'   datos - Dictionary con datos del registro
' ----------------------------------------------------------------------
Public Sub RegistrarCambioAudit(ByVal nombreLogico As String, _
                                 ByVal operacion As String, _
                                 ByVal datos As Object)
    On Error Resume Next
    
    Dim descripcion As String
    Dim idRegistro As String
    
    If datos.Exists("ID") Then
        idRegistro = datos("ID")
    ElseIf datos.Exists("IDPregunta") Then
        idRegistro = datos("IDPregunta")
    Else
        idRegistro = "(sin ID)"
    End If
    
    descripcion = operacion & " en " & nombreLogico & " | ID: " & idRegistro
    
    ' Usar el AuditLogger2 existente
    ' Firma: LogAction(action, sheetName, dataModified, beforeChange, afterChange, moduleAndSubroutine)
    Call AuditLogger2.LogAction( _
        action:=descripcion, _
        sheetName:=IIf(nombreLogico = "PLANTILLAS" Or nombreLogico = "PREGUNTAS", _
                       Configuration2.SHEET_CHECKLIST, Configuration2.SHEET_CONFIGURACION), _
        dataModified:=nombreLogico, _
        beforeChange:="", _
        afterChange:=idRegistro, _
        moduleAndSubroutine:="TableManager.RegistrarCambioAudit")
    
    On Error GoTo 0
End Sub

' ======================================================================
' SECCIÓN 7: LANZADOR PÚBLICO (para botón en Menú Principal)
' ======================================================================

' ----------------------------------------------------------------------
' AbrirGestorTablas
' Propósito: Punto de entrada público. Se asigna al botón en Menú Principal.
' Uso: Call TableManager.AbrirGestorTablas
' ----------------------------------------------------------------------
Public Sub AbrirGestorTablas()
    frmGestorTablas.Show vbModal
End Sub

' ======================================================================
' SECCIÓN 8: MIGRACIÓN DE BASE DE DATOS - INSPECCIONES RECURRENTES
' Fecha: 21/04/2026
' Propósito: Agregar 9 nuevas columnas a tblInspecciones para soportar
'            el sistema de inspecciones recurrentes con RPN histórico.
' ======================================================================

' ----------------------------------------------------------------------
' AgregarColumnasInspeccionesRecurrentes
' Propósito: Ejecuta la migración completa de tblInspecciones:
'            - Agrega columnas 32-40
'            - Backfill de "Puesto Evaluado" desde histórico
'            - Validación de integridad
' ADVERTENCIA: Este procedimiento modifica la estructura de BD.
'              Ejecutar solo UNA VEZ. Requiere backup previo.
' ----------------------------------------------------------------------
Public Sub AgregarColumnasInspeccionesRecurrentes()
    On Error GoTo ErrorHandler
    
    Dim wsHistorico As Worksheet
    Dim tbl As ListObject
    Dim registrosAntes As Long
    Dim registrosDespues As Long
    Dim columnasAntes As Long
    Dim i As Long
    
    ' FASE 1: Pre-validación
    Debug.Print "========================================="
    Debug.Print "MIGRACIÓN: Inspecciones Recurrentes"
    Debug.Print "Fecha: " & Now
    Debug.Print "========================================="
    
    Set wsHistorico = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Set tbl = wsHistorico.ListObjects(Configuration2.TABLE_INSPECCIONES)
    
    If tbl Is Nothing Then
        MsgBox "ERROR: No se encontró tblInspecciones", vbCritical
        Exit Sub
    End If
    
    registrosAntes = tbl.ListRows.Count
    columnasAntes = tbl.ListColumns.Count
    
    Debug.Print "Registros actuales: " & registrosAntes
    Debug.Print "Columnas actuales: " & columnasAntes
    
    ' Validar que tiene 31 columnas
    If columnasAntes <> 31 Then
        MsgBox "ERROR: Se esperaban 31 columnas, se encontraron " & columnasAntes & vbCrLf & _
               "La migración no puede continuar.", vbCritical
        Exit Sub
    End If
    
    ' Confirmar con usuario
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Desea agregar 9 columnas nuevas a tblInspecciones?" & vbCrLf & vbCrLf & _
                       "Registros: " & registrosAntes & vbCrLf & _
                       "Columnas actuales: 31 → Columnas nuevas: 40" & vbCrLf & vbCrLf & _
                       "ADVERTENCIA: Esta operación modificará la estructura de la BD.", _
                       vbYesNo + vbQuestion, "Confirmación de Migración")
    
    If respuesta = vbNo Then
        Debug.Print "Migración cancelada por el usuario"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' FASE 1.5: Desproteger hoja (crítico para agregar columnas)
    Debug.Print vbCrLf & "[FASE 1.5] Desprotegiendo hoja..."
    Dim hojaEstabProtegida As Boolean
    hojaEstabProtegida = wsHistorico.ProtectContents
    
    If hojaEstabProtegida Then
        On Error Resume Next
        Call SheetProtector2.UnprotectSheet(wsHistorico, Configuration2.APP_PASSWORD)
        If Err.Number <> 0 Then
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            MsgBox "ERROR: No se pudo desproteger la hoja. Verifique la contraseña.", vbCritical
            Exit Sub
        End If
        On Error GoTo ErrorHandler
        Debug.Print "  Hoja desprotegida OK"
    Else
        Debug.Print "  Hoja no estaba protegida"
    End If
    
    ' FASE 2: Agregar columnas AL FINAL (evita conflicto de movimiento de celdas)
    Debug.Print vbCrLf & "[FASE 2] Agregando columnas..."
    
    ' ESTRATEGIA: Primero agregar TODAS las columnas, luego llenar valores
    ' Esto evita problemas de referencias mientras Excel actualiza la estructura
    
    Dim colNumeroInsp As Long, colEsRecurrente As Long, colPuestoEval As Long
    
    ' Columna 32: Numero Inspeccion
    Debug.Print "  Agregando columna: Numero Inspeccion..."
    tbl.ListColumns.Add
    colNumeroInsp = tbl.ListColumns.Count
    tbl.ListColumns(colNumeroInsp).Name = "Numero Inspeccion"
    
    ' Columna 33: Es Inspeccion Recurrente
    Debug.Print "  Agregando columna: Es Inspeccion Recurrente..."
    tbl.ListColumns.Add
    colEsRecurrente = tbl.ListColumns.Count
    tbl.ListColumns(colEsRecurrente).Name = "Es Inspeccion Recurrente"
    
    ' Columna 34: Puesto Evaluado (CRÍTICA)
    Debug.Print "  Agregando columna: Puesto Evaluado..."
    tbl.ListColumns.Add
    colPuestoEval = tbl.ListColumns.Count
    tbl.ListColumns(colPuestoEval).Name = "Puesto Evaluado"
    
    ' Columna 35: RPN Anterior Manual
    Debug.Print "  Agregando columna: RPN Anterior Manual..."
    tbl.ListColumns.Add
    tbl.ListColumns(tbl.ListColumns.Count).Name = "RPN Anterior Manual"
    
    ' Columna 36: ID Inspeccion Anterior
    Debug.Print "  Agregando columna: ID Inspeccion Anterior..."
    tbl.ListColumns.Add
    tbl.ListColumns(tbl.ListColumns.Count).Name = "ID Inspeccion Anterior"
    
    ' Columna 37: RPN Promedio
    Debug.Print "  Agregando columna: RPN Promedio..."
    tbl.ListColumns.Add
    tbl.ListColumns(tbl.ListColumns.Count).Name = "RPN Promedio"
    
    ' Columna 38: Porcentaje Recuperacion
    Debug.Print "  Agregando columna: Porcentaje Recuperacion..."
    tbl.ListColumns.Add
    tbl.ListColumns(tbl.ListColumns.Count).Name = "Porcentaje Recuperacion"
    
    ' Columna 39: Porcentaje OOL
    Debug.Print "  Agregando columna: Porcentaje OOL..."
    tbl.ListColumns.Add
    tbl.ListColumns(tbl.ListColumns.Count).Name = "Porcentaje OOL"
    
    ' Columna 40: RPN Total
    Debug.Print "  Agregando columna: RPN Total..."
    tbl.ListColumns.Add
    tbl.ListColumns(tbl.ListColumns.Count).Name = "RPN Total"
    
    Debug.Print "[FASE 2] ✓ 9 columnas agregadas correctamente"
    
    ' FASE 2.5: Llenar valores por defecto
    Debug.Print vbCrLf & "[FASE 2.5] Llenando valores por defecto..."
    
    If Not tbl.DataBodyRange Is Nothing Then
        Debug.Print "  Llenando 'Numero Inspeccion' = 1..."
        tbl.ListColumns(colNumeroInsp).DataBodyRange.Value = 1
        
        Debug.Print "  Llenando 'Es Inspeccion Recurrente' = No..."
        tbl.ListColumns(colEsRecurrente).DataBodyRange.Value = "No"
        
        Debug.Print "  Llenando 'Puesto Evaluado' = (vacío)..."
        tbl.ListColumns(colPuestoEval).DataBodyRange.Value = ""
        
        Debug.Print "[FASE 2.5] ✓ Valores por defecto asignados"
    Else
        Debug.Print "[FASE 2.5] ⚠ Tabla vacía, sin valores que llenar"
    End If
    
    ' FASE 3: Backfill de Puesto Evaluado
    If Not tbl.DataBodyRange Is Nothing Then
        Debug.Print vbCrLf & "[FASE 3] Backfill de 'Puesto Evaluado'..."
        Call InferirPuestoHistorico(tbl)
    End If
    
    ' FASE 4: Validación
    Debug.Print vbCrLf & "[FASE 4] Validación post-migración..."
    
    registrosDespues = tbl.ListRows.Count
    
    If registrosAntes <> registrosDespues Then
        MsgBox "ERROR: Pérdida de datos detectada!" & vbCrLf & _
               "Antes: " & registrosAntes & " | Después: " & registrosDespues, vbCritical
        Exit Sub
    End If
    
    If tbl.ListColumns.Count <> 40 Then
        MsgBox "ERROR: Número de columnas incorrecto: " & tbl.ListColumns.Count, vbCritical
        Exit Sub
    End If
    
    ' Contar registros sin puesto asignado
    Dim sinPuesto As Long
    Dim colPuestoValidacion As Long
    sinPuesto = 0
    
    If Not tbl.DataBodyRange Is Nothing Then
        colPuestoValidacion = tbl.ListColumns("Puesto Evaluado").Index
        Dim fila As ListRow
        For Each fila In tbl.ListRows
            Dim puestoVal As String
            puestoVal = Trim(CStr(fila.Range.Cells(1, colPuestoValidacion).Value))
            If Len(puestoVal) = 0 Or puestoVal = "DESCONOCIDO" Then
                sinPuesto = sinPuesto + 1
            End If
        Next fila
    End If
    
    Debug.Print "  Registros antes: " & registrosAntes
    Debug.Print "  Registros después: " & registrosDespues
    Debug.Print "  Columnas actuales: " & tbl.ListColumns.Count
    Debug.Print "  Registros con Puesto Evaluado: " & (registrosDespues - sinPuesto)
    Debug.Print "  Registros sin puesto/desconocido: " & sinPuesto
    
    ' FASE 5: Reproteger hoja si estaba protegida
    If hojaEstabProtegida Then
        Debug.Print vbCrLf & "[FASE 5] Reprotegiendo hoja..."
        On Error Resume Next
        Call SheetProtector2.ProtectSheet(wsHistorico, Configuration2.APP_PASSWORD)
        On Error GoTo ErrorHandler
        Debug.Print "  Hoja reprotegida OK"
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Mensaje final
    Dim mensaje As String
    mensaje = "✓ MIGRACIÓN COMPLETADA" & vbCrLf & vbCrLf & _
              "Columnas: 31 → 40" & vbCrLf & _
              "Registros: " & registrosDespues & vbCrLf & _
              "Registros con puesto asignado: " & (registrosDespues - sinPuesto)
    
    If sinPuesto > 0 Then
        mensaje = mensaje & vbCrLf & vbCrLf & _
                  "⚠ ADVERTENCIA: " & sinPuesto & " registros requieren revisión manual" & vbCrLf & _
                  "(campo 'Puesto Evaluado' vacío o DESCONOCIDO)"
    End If
    
    MsgBox mensaje, vbInformation, "Migración Completada"
    
    Debug.Print vbCrLf & "========================================="
    Debug.Print "MIGRACIÓN COMPLETADA EXITOSAMENTE"
    Debug.Print "========================================="
    
    Exit Sub
    
ErrorHandler:
    ' Restaurar configuración
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Intentar reproteger hoja si estaba protegida
    On Error Resume Next
    If hojaEstabProtegida Then
        Call SheetProtector2.ProtectSheet(wsHistorico, Configuration2.APP_PASSWORD)
    End If
    On Error GoTo 0
    
    MsgBox "ERROR en migración: " & Err.Description, vbCritical
    Debug.Print "ERROR: " & Err.Description
End Sub

' ----------------------------------------------------------------------
' InferirPuestoHistorico (PRIVADA)
' Propósito: Para cada registro histórico en tblInspecciones, infiere
'            el "Puesto Evaluado" buscando en tblPlantillas.
' Lógica:
'   1. Leer ID Plantilla (col 11)
'   2. Buscar en tblPlantillas → obtener Puesto
'   3. Escribir en Puesto Evaluado (col 34)
'   4. Si no encuentra → "DESCONOCIDO" (requiere corrección manual)
' ----------------------------------------------------------------------
Private Sub InferirPuestoHistorico(ByRef tbl As ListObject)
    On Error GoTo ErrorHandler
    
    Dim wsChecklist As Worksheet
    Dim tblPlantillas As ListObject
    Dim fila As ListRow
    Dim idPlantilla As String
    Dim puestoInferido As String
    Dim contador As Long
    Dim exitos As Long
    Dim fallidos As Long
    Dim colPuestoIndex As Long
    Dim colIDPlantillaIndex As Long
    
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set tblPlantillas = wsChecklist.ListObjects(Configuration2.TABLE_PLANTILLAS)
    
    ' Obtener índices por nombre (robusto ante cambios de posición)
    colPuestoIndex = tbl.ListColumns("Puesto Evaluado").Index
    colIDPlantillaIndex = tbl.ListColumns("ID Plantilla").Index
    
    contador = 0
    exitos = 0
    fallidos = 0
    
    Debug.Print "  Inferencia de puesto desde tblPlantillas..."
    
    For Each fila In tbl.ListRows
        contador = contador + 1
        
        ' Leer ID Plantilla (por nombre de columna)
        idPlantilla = Trim(CStr(fila.Range.Cells(1, colIDPlantillaIndex).Value))
        
        If Len(idPlantilla) = 0 Then
            ' Sin plantilla → marcar como desconocido
            fila.Range.Cells(1, colPuestoIndex).Value = "DESCONOCIDO"
            fallidos = fallidos + 1
        Else
            ' Buscar puesto en tblPlantillas
            puestoInferido = BuscarPuestoPorPlantilla(tblPlantillas, idPlantilla)
            
            If puestoInferido = "" Then
                fila.Range.Cells(1, colPuestoIndex).Value = "DESCONOCIDO"
                fallidos = fallidos + 1
            Else
                fila.Range.Cells(1, colPuestoIndex).Value = puestoInferido
                exitos = exitos + 1
            End If
        End If
        
        ' Progress cada 50 registros
        If contador Mod 50 = 0 Then
            Debug.Print "    Procesados: " & contador & " registros..."
        End If
    Next fila
    
    Debug.Print "[FASE 3] ✓ Backfill completado"
    Debug.Print "  Total procesado: " & contador
    Debug.Print "  Éxitos: " & exitos
    Debug.Print "  Desconocidos: " & fallidos
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en InferirPuestoHistorico: " & Err.Description
End Sub

' ----------------------------------------------------------------------
' BuscarPuestoPorPlantilla (PRIVADA)
' Propósito: Busca en tblPlantillas el puesto correspondiente a un ID Plantilla.
' Retorna: Nombre del puesto, o "" si no se encuentra.
' ----------------------------------------------------------------------
Private Function BuscarPuestoPorPlantilla(ByRef tblPlantillas As ListObject, _
                                           ByVal idPlantilla As String) As String
    On Error Resume Next
    
    Dim filaPlantilla As ListRow
    Dim idActual As String
    Dim puesto As String
    
    BuscarPuestoPorPlantilla = ""
    
    If tblPlantillas.DataBodyRange Is Nothing Then Exit Function
    
    For Each filaPlantilla In tblPlantillas.ListRows
        idActual = Trim(CStr(filaPlantilla.Range.Cells(1, _
                   tblPlantillas.ListColumns("ID Plantilla").Index).Value))
        
        If idActual = Trim(idPlantilla) Then
            puesto = Trim(CStr(filaPlantilla.Range.Cells(1, _
                     tblPlantillas.ListColumns("Puesto").Index).Value))
            
            BuscarPuestoPorPlantilla = puesto
            Exit Function
        End If
    Next filaPlantilla
    
End Function
