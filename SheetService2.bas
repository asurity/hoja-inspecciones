
' ----------------------------------------------------------------------
' Módulo: SheetService
' Descripción: Servicio centralizado para todas las acciones de protección y ocultación de hojas.
'              Garantiza que la lógica de protección/ocultación no se duplique en otros módulos o eventos.
'              Facilita el mantenimiento y la seguridad de la estructura del libro.
' ----------------------------------------------------------------------
Option Explicit


'' ----------------------------------------------------------------------
' Función Privada: IsAuditSheet
' Propósito: Devuelve True si el nombre de hoja pertenece al grupo Audit Trail
'            (cualquiera de las AUDIT_MAX_SHEETS hojas configuradas en Configuration2).
'' ----------------------------------------------------------------------
Private Function IsAuditSheet(ByVal sheetName As String) As Boolean
    Dim i As Long
    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        If sheetName = AuditRotation2.ObtenerNombreHoja(i) Then
            IsAuditSheet = True
            Exit Function
        End If
    Next i
    IsAuditSheet = False
End Function

'' ----------------------------------------------------------------------
' Subrutina: HideAndProtectAllSheetsExcept
' Propósito: Muestra la hoja destino y oculta todas las demás, excepto "Menú principal"
'            que siempre permanece visible. La protección individual de cada hoja
'            se gestiona en su propio evento Worksheet_Activate.
' Argumentos:
'   - sheetName: Nombre de la hoja que debe mostrarse.
'   - hojaAnterior: Parámetro legacy mantenido por compatibilidad.
' Lógica:
'   1. Desprotege estructura del libro.
'   2. Hace visible la hoja destino (solo si NO es hoja Audit Trail).
'   3. Oculta TODAS las hojas Audit Trail (siempre, sin excepción).
'   4. Oculta todas las demás hojas excepto destino y "Menú principal".
'   5. Re-protege estructura del libro.
'
' REGLAS DE VISIBILIDAD (10/03/2026):
'   - Menú Principal: SIEMPRE visible
'   - Hojas Audit Trail: NUNCA visibles mediante navegación normal
'   - Otras hojas: Visible solo la hoja destino
'' ----------------------------------------------------------------------
Public Sub HideAndProtectAllSheetsExcept(ByVal sheetName As String, Optional ByVal hojaAnterior As String = "")
    On Error GoTo ErrorHandler
    Call WorkbookProtector2.UnprotectWorkbook
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Dim i As Long
    Dim auditName As String
    
    ' Paso 1: Ocultar TODAS las hojas Audit Trail (sin excepción)
    ' Las hojas Audit Trail solo se muestran mediante ShowAuditTrailGroup()
    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        auditName = AuditRotation2.ObtenerNombreHoja(i)
        On Error Resume Next ' La hoja puede no existir
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(auditName)
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVeryHidden
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' Paso 2: Hacer visible la hoja destino (solo si NO es Audit Trail)
    If Not IsAuditSheet(sheetName) Then
        ThisWorkbook.Sheets(sheetName).Visible = xlSheetVisible
    End If
    
    ' Paso 3: Ocultar todas las demás hojas excepto destino y "Menú principal"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> sheetName And ws.Name <> "Menú principal" And Not IsAuditSheet(ws.Name) Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    
    ' Paso 4: Asegurar que "Menú principal" siempre esté visible
    ThisWorkbook.Sheets("Menú principal").Visible = xlSheetVisible
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Call WorkbookProtector2.ProtectWorkbook
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not IsAuditSheet(sheetName) Then
        ThisWorkbook.Sheets(sheetName).Visible = xlSheetVisible
    End If
    On Error GoTo 0
    Call WorkbookProtector2.ProtectWorkbook
    Call ErrorLogger2.Log("SheetService.HideAndProtectAllSheetsExcept", VBA.Err.Description, VBA.Err.Number)
    MsgBox "Ocurrió un error al intentar mostrar la hoja.", vbCritical
End Sub

'' ----------------------------------------------------------------------
' Subrutina: ShowAuditTrailGroup
' Propósito: Muestra las 5 hojas Audit Trail y oculta el resto del libro,
'            manteniendo siempre visible "Menú principal".
' Lógica:
'   1. Desprotege estructura del libro.
'   2. Oculta todas las hojas que no son de auditoría ni "Menú principal".
'   3. Hace visible cada hoja Audit Trail que exista en el libro.
'   4. Asegura que "Menú principal" esté visible.
'   5. Selecciona la primera hoja de Audit Trail.
'   6. Re-protege estructura del libro.
'' ----------------------------------------------------------------------
Public Sub ShowAuditTrailGroup()
    On Error GoTo ErrorHandler
    Call WorkbookProtector2.UnprotectWorkbook
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Dim i As Long
    Dim auditName As String
    
    ' Paso 1: ocultar todas las hojas que no son de auditoría ni Menú principal.
    For Each ws In ThisWorkbook.Worksheets
        If Not IsAuditSheet(ws.Name) And ws.Name <> "Menú principal" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    
    ' Paso 2: mostrar cada hoja de Audit Trail que exista en el libro.
    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        auditName = AuditRotation2.ObtenerNombreHoja(i)
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(auditName)
        On Error GoTo ErrorHandler
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVisible
        End If
    Next i
    
    ' Paso 3: asegurar que "Menú principal" esté visible.
    ThisWorkbook.Sheets("Menú principal").Visible = xlSheetVisible
    
    ' Paso 4: activar la primera hoja de Audit Trail.
    ThisWorkbook.Sheets(Configuration2.AUDIT_BASE_NAME).Select
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Call WorkbookProtector2.ProtectWorkbook
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Call WorkbookProtector2.ProtectWorkbook
    Call ErrorLogger2.Log("SheetService.ShowAuditTrailGroup", VBA.Err.Description, VBA.Err.Number)
    MsgBox "Ocurrió un error al mostrar las hojas de Audit trail.", vbCritical, "Error de Navegación"
End Sub


'' ----------------------------------------------------------------------
' Subrutina: LockAllSheets
' Propósito: Protege todas las hojas del libro con la contraseña de la aplicación.
' Lógica:
'   1. Desactiva las alertas para evitar mensajes de Excel.
'   2. Itera sobre todas las hojas y las protege usando SheetProtector.
'   3. Reactiva las alertas al finalizar.
'' ----------------------------------------------------------------------
Public Sub LockAllSheets()
    ' FASE 1 (22/02/2026): Cambiado de ProtectSheet a ApplyRoleBasedProtection para que
    ' incluso en el flujo de error de navegación se respete el rol del usuario.
    ' "Usuario" sigue sin poder editar (ProtectSheetForReading) pero puede copiar datos.
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        Call SheetProtector2.ApplyRoleBasedProtection(ws, Configuration2.APP_PASSWORD)
    Next ws
    Application.DisplayAlerts = True
End Sub


'' ----------------------------------------------------------------------
' Subrutina: UnlockAndShowAllSheets
' Propósito: Desprotege y muestra todas las hojas del libro.
' Lógica:
'   1. Desactiva las alertas para evitar mensajes de Excel.
'   2. Itera sobre todas las hojas, las desprotege y las hace visibles.
'   3. Reactiva las alertas al finalizar.
'' ----------------------------------------------------------------------
Public Sub UnlockAndShowAllSheets()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        Call SheetProtector2.UnprotectSheet(ws, Configuration2.APP_PASSWORD)
        ws.Visible = xlSheetVisible
    Next ws
    Application.DisplayAlerts = True
End Sub


'' ----------------------------------------------------------------------
' Subrutina: HideAndProtectSheet
' Propósito: Oculta y protege una hoja específica del libro.
' Argumentos:
'   - sheetName: Nombre de la hoja a ocultar y proteger.
' Lógica:
'   1. Desactiva las alertas para evitar mensajes de Excel.
'   2. Busca la hoja indicada, la oculta y la protege.
'   3. Reactiva las alertas y maneja errores si ocurren.
'' ----------------------------------------------------------------------
Public Sub HideAndProtectSheet(sheetName As String)
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Visible = xlSheetVeryHidden
    Call SheetProtector2.ProtectSheet(ws, Configuration2.APP_PASSWORD)
    Application.DisplayAlerts = True
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("SheetService.HideAndProtectSheet", VBA.Err.Description, VBA.Err.Number)
    Application.DisplayAlerts = True
End Sub