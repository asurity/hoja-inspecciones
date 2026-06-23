
' ----------------------------------------------------------------------
' Módulo: SheetService
' Descripción: Servicio centralizado para todas las acciones de protección y ocultación de hojas.
'              Garantiza que la lógica de protección/ocultación no se duplique en otros módulos o eventos.
'              Facilita el mantenimiento y la seguridad de la estructura del libro.
' ----------------------------------------------------------------------
Option Explicit


' ----------------------------------------------------------------------
' Función Privada: IsAuditSheet
' Propósito: Devuelve True si el nombre de hoja pertenece al grupo Audit Trail.
'            Usa GetAuditTrailSheetNames() (FASE 0) en lugar de
'            AuditRotation2.ObtenerNombreHoja() para evitar dependencia circular.
' Refactorizado: 08/06/2026 - FASE 1 Refactorización Navegación
' ----------------------------------------------------------------------
Private Function IsAuditSheet(ByVal sheetName As String) As Boolean
    Dim auditNames As Variant
    auditNames = Configuration2.GetAuditTrailSheetNames()
    Dim i As Long
    For i = LBound(auditNames) To UBound(auditNames)
        If StrComp(sheetName, CStr(auditNames(i)), vbTextCompare) = 0 Then
            IsAuditSheet = True
            Exit Function
        End If
    Next i
    IsAuditSheet = False
End Function

' ----------------------------------------------------------------------
' Subrutina: ShowOnly
' Propósito: Muestra SOLO las hojas especificadas (más "Menú principal") y
'            oculta todas las demás hojas de módulo. NO itera sobre todas las
'            hojas del libro — solo sobre las definidas en GetAllModuleSheetNames().
' Argumentos:
'   - applyProtection: Si True, aplica ApplyRoleBasedProtection a las hojas visibles.
'   - sheetNames(): Nombres de las hojas que deben quedar visibles.
' REGLAS:
'   - NUNCA toca ScreenUpdating, EnableEvents, ni DisplayAlerts.
'   - NUNCA toca protección de estructura del libro.
'   - "Menú principal" siempre se fuerza visible.
' Creado: 08/06/2026 - FASE 1 Refactorización Navegación
' ----------------------------------------------------------------------
Public Sub ShowOnly(ByVal applyProtection As Boolean, ParamArray sheetNames() As Variant)
    Dim allModules As Variant
    allModules = Configuration2.GetAllModuleSheetNames()
    
    Dim ws As Worksheet

    ' Asegurar que "Menú principal" siempre esté visible
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Menú principal")
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVisible
    End If
    On Error GoTo 0

    Dim i As Long, j As Long
    Dim moduleName As String
    Dim shouldBeVisible As Boolean
    
    
    For i = LBound(allModules) To UBound(allModules)
        moduleName = CStr(allModules(i))
        
        ' Determinar si esta hoja debe estar visible
        shouldBeVisible = False
        For j = LBound(sheetNames) To UBound(sheetNames)
            If StrComp(moduleName, CStr(sheetNames(j)), vbTextCompare) = 0 Then
                shouldBeVisible = True
                Exit For
            End If
        Next j
        
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(moduleName)
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            If shouldBeVisible Then
                On Error Resume Next
                ws.Visible = xlSheetVisible
                On Error GoTo 0
                If applyProtection Then
                    Dim pwd As String
                    If IsAuditSheet(moduleName) Then
                        pwd = Configuration2.AUDIT_PASSWORD
                    Else
                        pwd = Configuration2.APP_PASSWORD
                    End If
                    Call SheetProtector2.ApplyRoleBasedProtection(ws, pwd)
                End If
            Else
                On Error Resume Next
                ws.Visible = xlSheetVeryHidden
                On Error GoTo 0
            End If
        End If
        Set ws = Nothing
    Next i
    
End Sub


' ----------------------------------------------------------------------
' Subrutina: HideAndProtectAllSheetsExcept
' @deprecated — Usar ShowOnly en su lugar. ShowOnly itera solo sobre
'              ALL_MODULE_SHEETS (no sobre todas las hojas del libro) y
'              no toca ScreenUpdating/EnableEvents/DisplayAlerts ni la
'              protección de estructura. Mantenido por compatibilidad
'              durante la transición (FASE 1, 08/06/2026).
'
' Propósito: Muestra la hoja destino y oculta todas las demás, excepto "Menú principal"
'            que siempre permanece visible. La protección individual de cada hoja
'            se gestiona en su propio evento Worksheet_Activate.
' Argumentos:
'   - sheetName: Nombre de la hoja que debe mostrarse.
'   - hojaAnterior: Parámetro legacy mantenido por compatibilidad.
' ----------------------------------------------------------------------
Public Sub HideAndProtectAllSheetsExcept(ByVal sheetName As String, Optional ByVal hojaAnterior As String = "")
    ' ## NAVEGACIÓN ## @deprecated — Usar ShowOnly en su lugar (FASE 1, 08/06/2026)
    Debug.Print "[SheetService.HideAndProtectAllSheetsExcept] @deprecated → '" & sheetName & "'"
    
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Call WorkbookProtector2.UnprotectWorkbook
    
    Dim ws As Worksheet
    Dim i As Long
    Dim auditName As String
    Dim auditNames As Variant
    auditNames = Configuration2.GetAuditTrailSheetNames()
    
    ' Ocultar todas las hojas Audit Trail
    For i = LBound(auditNames) To UBound(auditNames)
        auditName = CStr(auditNames(i))
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(auditName)
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVeryHidden
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' Hacer visible la hoja destino (solo si NO es Audit Trail)
    If Not IsAuditSheet(sheetName) Then
        On Error Resume Next
        ThisWorkbook.Sheets(sheetName).Visible = xlSheetVisible
        On Error GoTo ErrorHandler
    End If
    
    ' Ocultar todas las demás hojas excepto destino y "Menú principal"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> sheetName And ws.Name <> "Menú principal" And Not IsAuditSheet(ws.Name) Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    
    ' Asegurar que "Menú principal" siempre esté visible
    On Error Resume Next
    ThisWorkbook.Sheets("Menú principal").Visible = xlSheetVisible
    On Error GoTo ErrorHandler
    
    ' Aplicar protección según rol a la hoja destino
    On Error Resume Next
    Call SheetProtector2.ApplyRoleBasedProtection( _
        ThisWorkbook.Sheets(sheetName), _
        Configuration2.APP_PASSWORD _
    )
    On Error GoTo ErrorHandler
    
    Call WorkbookProtector2.ProtectWorkbook
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Debug.Print "[SheetService.HideAndProtectAllSheetsExcept] ERROR: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not IsAuditSheet(sheetName) Then
        ThisWorkbook.Sheets(sheetName).Visible = xlSheetVisible
    End If
    On Error GoTo 0
    Call WorkbookProtector2.ProtectWorkbook
    Call ErrorLogger2.Log("SheetService.HideAndProtectAllSheetsExcept", VBA.Err.Description, VBA.Err.Number)
    MsgBox "Ocurrió un error al intentar mostrar la hoja.", vbCritical
End Sub


' ----------------------------------------------------------------------
' Subrutina: ShowAuditTrailGroup
' Propósito: Muestra las 5 hojas Audit Trail y oculta el resto del libro,
'            manteniendo siempre visible "Menú principal".
' Refactorizado: 08/06/2026 - FASE 1 Refactorización Navegación
' Cambios:
'   - Usa GetAuditTrailSheetNames() en vez de AuditRotation2.ObtenerNombreHoja()
'   - Delega visibilidad a ShowOnly (sin protección, sin tocar estructura)
'   - NO toca ScreenUpdating/EnableEvents/DisplayAlerts (lo maneja el llamador)
'   - NO aplica protección (lo maneja Worksheet_Activate de cada hoja)
'   - NO desprotege/reprotege estructura del libro (lo maneja el llamador)
' ----------------------------------------------------------------------
Public Sub ShowAuditTrailGroup()
    ' ## NAVEGACIÓN ## Refactorizado FASE 1 (08/06/2026)
    Debug.Print "[SheetService.ShowAuditTrailGroup] Mostrando grupo Audit Trail..."
    
    On Error GoTo ErrorHandler
    
    Dim auditNames As Variant
    auditNames = Configuration2.GetAuditTrailSheetNames()
    
    ' Delegar visibilidad a ShowOnly: mostrar las 5 audit + Menú principal
    ' Sin aplicar protección (Worksheet_Activate se encarga)
    Call ShowOnly(False, _
        CStr(auditNames(0)), _
        CStr(auditNames(1)), _
        CStr(auditNames(2)), _
        CStr(auditNames(3)), _
        CStr(auditNames(4)) _
    )
    
    ' Activar la primera hoja de Audit Trail
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CStr(auditNames(0)))
    If Not ws Is Nothing Then
        ws.Select
    Else
        ' Buscar cualquier audit trail visible como fallback
        Dim i As Long
        For i = LBound(auditNames) To UBound(auditNames)
            Set ws = Nothing
            Set ws = ThisWorkbook.Sheets(CStr(auditNames(i)))
            If Not ws Is Nothing Then
                If ws.Visible = xlSheetVisible Then
                    ws.Select
                    Exit For
                End If
            End If
        Next i
    End If
    On Error GoTo ErrorHandler
    
    Debug.Print "[SheetService.ShowAuditTrailGroup] Completado"
    Exit Sub

ErrorHandler:
    Debug.Print "[SheetService.ShowAuditTrailGroup] ERROR: " & Err.Number & " - " & Err.Description
    Call ErrorLogger2.Log("SheetService.ShowAuditTrailGroup", VBA.Err.Description, VBA.Err.Number)
End Sub


' ----------------------------------------------------------------------
' Subrutina: LockAllSheets
' Propósito: Protege todas las hojas del libro con la contraseña de la aplicación.
' Lógica:
'   1. Desactiva las alertas para evitar mensajes de Excel.
'   2. Itera sobre todas las hojas y las protege usando SheetProtector.
'   3. Reactiva las alertas al finalizar.
' ----------------------------------------------------------------------
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


' ----------------------------------------------------------------------
' Subrutina: UnlockAndShowAllSheets
' Propósito: Desprotege y muestra todas las hojas del libro.
' Lógica:
'   1. Desactiva las alertas para evitar mensajes de Excel.
'   2. Itera sobre todas las hojas, las desprotege y las hace visibles.
'   3. Reactiva las alertas al finalizar.
' ----------------------------------------------------------------------
Public Sub UnlockAndShowAllSheets()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        Call SheetProtector2.UnprotectSheet(ws, Configuration2.APP_PASSWORD)
        ws.Visible = xlSheetVisible
    Next ws
    Application.DisplayAlerts = True
End Sub


' ----------------------------------------------------------------------
' Subrutina: HideAndProtectSheet
' Propósito: Oculta y protege una hoja específica del libro.
' Argumentos:
'   - sheetName: Nombre de la hoja a ocultar y proteger.
' Lógica:
'   1. Desactiva las alertas para evitar mensajes de Excel.
'   2. Busca la hoja indicada, la oculta y la protege.
'   3. Reactiva las alertas y maneja errores si ocurren.
' ----------------------------------------------------------------------
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
