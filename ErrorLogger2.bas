'' ----------------------------------------------------------------------
' Módulo: ErrorLogger
' Descripción: Centraliza el registro de errores de la aplicación.
'              Captura, formatea y guarda información sobre los errores
'              en una tabla de Excel, asegurando trazabilidad y diagnóstico.
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Subrutina: Log
' Propósito: Registra un error en la hoja "Registro de errores" y en la tabla "tblErrores".
'            Si la hoja o la tabla no existen, las crea automáticamente.
' Parámetros:
'   - sourceSubroutine: Nombre de la subrutina donde ocurrió el error.
'   - errorDescription: Descripción del error.
'   - errorNumber: Código numérico del error.
' Lógica:
'   1. Busca o crea la hoja y la tabla de errores.
'   2. Desprotege la hoja antes de escribir.
'   3. Agrega una nueva fila con los datos del error.
'   4. Vuelve a proteger la hoja y limpia el error.
'   5. Si ocurre un error en el registro, protege la hoja y notifica al usuario.
' ----------------------------------------------------------------------
Public Sub Log(ByVal sourceSubroutine As String, ByVal errorDescription As String, ByVal errorNumber As Long)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    
    ' 1. Obtener o crear la hoja de registro
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Registro de errores")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Registro de errores"
        
        ' Crear los encabezados de la tabla
        With ws
            .Range("B8:G8").Value = Array("Fecha", "Hora", "Nombre de subrutina del error", "Descripción del error", "Número del error", "Con quien se generó el error")
        End With
        
        ' Crear la tabla ListObject
        Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=ws.Range("B8:G8"), XlListObjectHasHeaders:=xlYes, TableStyleName:="TableStyleMedium2")
        tbl.Name = "tblErrores"
    Else
        Set tbl = ws.ListObjects("tblErrores")
    End If
    
    ' Desproteger la hoja de errores antes de escribir.
    ' NOTA: Se desprotege directamente (sin pasar por SheetProtector2) porque
    ' cuando ENABLE_SHEET_PROTECTION=False, UnprotectSheet es no-op, pero la
    ' hoja puede seguir protegida de una ejecución anterior con el flag en True.
    ' El ErrorLogger SIEMPRE debe poder escribir, independientemente del flag.
    On Error Resume Next
    ws.Unprotect Password:=Configuration2.APP_PASSWORD
    On Error GoTo ErrorHandler
    
    ' 2. Determinar la fila de destino para la nueva entrada
    Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)
    
    ' 3. Escribir los datos del error en la nueva fila
    With newRow.Range
        .Cells(1, 1).Value = Date ' Fecha
        .Cells(1, 2).Value = Time ' Hora
        .Cells(1, 3).Value = sourceSubroutine ' Nombre de la subrutina del error
        .Cells(1, 4).Value = errorDescription ' Descripción del error
        .Cells(1, 5).Value = errorNumber ' Número del error
        .Cells(1, 6).Value = Application.userName ' Usuario que generó el error
    End With
    
    ' Volver a proteger la hoja después de escribir.
    Call SheetProtector2.ProtectSheet(ws, Configuration2.APP_PASSWORD)
    
    ' Limpiar el error para que no se propague
    VBA.Err.Clear
    
    Exit Sub
ErrorHandler:
    ' Asegurar que la hoja quede protegida incluso si ocurre un error.
    On Error Resume Next
    ws.Protect Password:=Configuration2.APP_PASSWORD
    On Error GoTo 0
    ' Mostrar el error ORIGINAL que intentaba registrar, para no perder la pista.
    MsgBox "Error al registrar en log." & vbCrLf & _
           "Error original: [" & errorNumber & "] " & errorDescription & vbCrLf & _
           "Origen: " & sourceSubroutine, vbCritical, "Error de Registro"
End Sub
