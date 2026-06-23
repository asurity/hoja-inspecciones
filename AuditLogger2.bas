Option Explicit

' Bandera para mostrar el mensaje de "Audit Trail lleno" una sola vez por sesión.
Private auditTrailLlenoNotificado As Boolean

' ------------------------------------------------------------------------------
' Módulo: AuditLogger2
' Descripción: Servicio centralizado de auditoría con soporte de rotación
'              automática entre hasta AUDIT_MAX_SHEETS hojas Audit Trail.
'              Cuando una hoja alcanza AUDIT_MAX_ROWS filas, AuditRotation2
'              activa la siguiente hoja automáticamente.
' Dependencias:
'   - AuditRotation2  : selecciona la hoja activa con espacio disponible.
'   - SheetProtector2 : protege/desprotege hojas de auditoría.
'   - Configuration2  : AUDIT_PASSWORD, AUDIT_MAX_ROWS, AUDIT_MAX_SHEETS, etc.
'   - ErrorLogger2    : registro de errores internos.
' Última modificación: 22/02/2026 — integración con AuditRotation2.
' ------------------------------------------------------------------------------

' ------------------------------------------------------------------------------
' Subrutina: LogAction
' Registra una acción en la hoja Audit Trail activa con espacio disponible.
' Si la hoja actual está llena, AuditRotation2 rota automáticamente a la
' siguiente hoja antes de escribir.
'
' Parámetros:
'   action              Tipo de acción (ej. "Cargar detección").
'   sheetName           Hoja donde ocurrió el cambio.
'   dataModified        Campo o dato modificado.
'   beforeChange        Valor previo al cambio.
'   afterChange         Valor posterior al cambio.
'   moduleAndSubroutine Módulo y subrutina de origen.
' ------------------------------------------------------------------------------
Public Sub LogAction(ByVal action As String, ByVal sheetName As String, _
                     ByVal dataModified As String, ByVal beforeChange As String, _
                     ByVal afterChange As String, ByVal moduleAndSubroutine As String)
    On Error GoTo ErrorHandler

    Application.EnableEvents = False

    Dim ws          As Worksheet
    Dim tbl         As ListObject
    Dim newRow      As ListRow
    Dim i           As Long
    Dim nombreTabla As String
    ' FASE 5 (22/02/2026): identificador de auditoría confiable — combina usuario de Windows
    ' con el rol activo del sistema. Application.UserName era configurable manualmente por el usuario
    ' y no estaba vinculado al login ni al rol, comprometiendo la trazabilidad.
    Dim auditUser   As String

    ' -- Paso 1: obtener la hoja con espacio disponible (rotación automática) --
    Set ws = AuditRotation2.ObtenerHojaAuditActiva()

    If ws Is Nothing Then
        ' Audit Trail completamente lleno — notificar UNA sola vez al usuario.
        If Not auditTrailLlenoNotificado Then
            auditTrailLlenoNotificado = True
            MsgBox "ATENCIÓN: Se ha alcanzado el límite máximo de registros de auditoría " & _
                   "(" & Configuration2.AUDIT_MAX_SHEETS & " hojas × " & Format(Configuration2.AUDIT_MAX_ROWS, "#,##0") & " filas)." & vbCrLf & vbCrLf & _
                   "El sistema continuará funcionando normalmente, pero NO se registrarán " & _
                   "nuevos eventos de auditoría hasta que se libere espacio." & vbCrLf & vbCrLf & _
                   "Contacte al administrador para archivar los registros existentes.", _
                   vbExclamation, "Audit Trail Lleno"
        End If
        Debug.Print "[AuditLogger2] Audit Trail lleno — registro NO guardado (acción: " & action & ")."
        Application.EnableEvents = True
        Exit Sub
    End If

    ' -- Paso 2: determinar el nombre de tabla correspondiente a esa hoja ------
    nombreTabla = ""
    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        If ws.Name = AuditRotation2.ObtenerNombreHoja(i) Then
            nombreTabla = AuditRotation2.ObtenerNombreTabla(i)
            Exit For
        End If
    Next i

    If nombreTabla = "" Then
        ' Fallback defensivo: intentar con el prefijo base.
        nombreTabla = Configuration2.AUDIT_TABLE_PREFIX
        Debug.Print "[AuditLogger2] AVISO: no se reconoció '" & ws.Name & "' — usando tabla de fallback '" & nombreTabla & "'."
    End If

    ' -- Paso 2.5: construir identificador de auditoría -------------------------
    ' FASE 5 (22/02/2026): Se usa Environ("USERNAME") (usuario de Windows, no editable por el
    ' usuario) en lugar de Application.UserName, que era configurable manualmente y no confiable.
    auditUser = Environ("USERNAME")

    ' -- Paso 3: desproteger, escribir, proteger -------------------------------
    Call SheetProtector2.UnprotectSheet(ws, Configuration2.AUDIT_PASSWORD)

    On Error Resume Next
    Set tbl = ws.ListObjects(nombreTabla)
    On Error GoTo ErrorHandler

    If tbl Is Nothing Then
        Debug.Print "[AuditLogger2] ERROR: tabla '" & nombreTabla & "' no encontrada en '" & ws.Name & "'."
        GoTo ErrorHandler
    End If

    If tbl.DataBodyRange Is Nothing Then
        ' Tabla vacía: escribir en la primera fila de datos.
        ' La posición B9 es estándar para todas las hojas Audit Trail de este proyecto.
        With ws.Range("B9")
            .Value = Date
            .Offset(0, 1).Value = Time
            .Offset(0, 2).Value = auditUser
            .Offset(0, 3).Value = sheetName
            .Offset(0, 4).Value = action
            .Offset(0, 5).Value = dataModified
            .Offset(0, 6).Value = beforeChange
            .Offset(0, 7).Value = afterChange
            .Offset(0, 8).Value = moduleAndSubroutine
        End With
    Else
        Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)
        With newRow.Range
            .Cells(1, 1).Value = Date
            .Cells(1, 2).Value = Time
            .Cells(1, 3).Value = auditUser
            .Cells(1, 4).Value = sheetName
            .Cells(1, 5).Value = action
            .Cells(1, 6).Value = dataModified
            .Cells(1, 7).Value = beforeChange
            .Cells(1, 8).Value = afterChange
            .Cells(1, 9).Value = moduleAndSubroutine
        End With
    End If

    Call SheetProtector2.ProtectSheet(ws, Configuration2.AUDIT_PASSWORD)
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    If Not ws Is Nothing Then
        Call SheetProtector2.ProtectSheet(ws, Configuration2.AUDIT_PASSWORD)
    End If
    Call ErrorLogger2.Log("AuditLogger2.LogAction", VBA.Err.Description, VBA.Err.Number)
End Sub