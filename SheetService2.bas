
' ----------------------------------------------------------------------
' Módulo: SheetService
' Descripción: Servicio centralizado para todas las acciones de protección y ocultación de hojas.
'              Garantiza que la lógica de protección/ocultación no se duplique en otros módulos o eventos.
'              Facilita el mantenimiento y la seguridad de la estructura del libro.
' ----------------------------------------------------------------------
Option Explicit


' ----------------------------------------------------------------------
' Función Privada: IsAuditSheet
' Propósito: Devuelve True si el nombre de hoja pertenece al grupo Audit Trail
'            (cualquiera de las AUDIT_MAX_SHEETS hojas configuradas en Configuration2).
' ----------------------------------------------------------------------
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

' ----------------------------------------------------------------------
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
' ----------------------------------------------------------------------
Public Sub HideAndProtectAllSheetsExcept(ByVal sheetName As String, Optional ByVal hojaAnterior As String = "")
    Debug.Print "═══════════════════════════════════════════════════════════"
    Debug.Print "[SheetService.HideAndProtectAllSheetsExcept] INICIO"
    Debug.Print "[SheetService] Hoja destino solicitada: '" & sheetName & "'"
    Debug.Print "[SheetService] Rol actual del usuario: '" & m_userRole & "'"
    Debug.Print "[SheetService] APP_PASSWORD en Configuration2: '" & Configuration2.APP_PASSWORD & "' (" & Len(Configuration2.APP_PASSWORD) & " caracteres)"
    Debug.Print "[SheetService] ENABLE_WORKBOOK_PROTECTION = " & Configuration2.ENABLE_WORKBOOK_PROTECTION
    Debug.Print "[SheetService] ENABLE_SHEET_PROTECTION = " & Configuration2.ENABLE_SHEET_PROTECTION
    Debug.Print "───────────────────────────────────────────────────────────"
    
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' PASO 1: Desproteger estructura del libro
    Debug.Print "[SheetService] PASO 1: Desprotegiendo estructura del libro..."
    Call WorkbookProtector2.UnprotectWorkbook
    Debug.Print "[SheetService] PASO 1: Completado"
    
    ' Verificar que el libro se desprotegió correctamente (Structure:=False indica desprotegido)
    ' Si sigue protegido, las operaciones de visibilidad fallarán silenciosamente.
    If ThisWorkbook.ProtectStructure Then
        Debug.Print "[SheetService] ⚠ AVISO: La estructura del libro SIGUE protegida después de UnprotectWorkbook"
        Debug.Print "[SheetService]   Las operaciones de ocultación/mostrado de hojas pueden fallar"
        Debug.Print "[SheetService]   Continuando de todas formas (On Error Resume Next capturará errores)..."
    Else
        Debug.Print "[SheetService] ✔ Estructura del libro desprotegida correctamente"
    End If
    
    Dim ws As Worksheet
    Dim i As Long
    Dim auditName As String
    
    ' Paso 1: Ocultar TODAS las hojas Audit Trail (sin excepción)
    ' Las hojas Audit Trail solo se muestran mediante ShowAuditTrailGroup()
    Debug.Print "[SheetService] PASO 2: Ocultando " & Configuration2.AUDIT_MAX_SHEETS & " hojas Audit Trail..."
    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        auditName = AuditRotation2.ObtenerNombreHoja(i)
        On Error Resume Next ' La hoja puede no existir
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(auditName)
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVeryHidden
            Debug.Print "[SheetService]   ✔ Audit Trail " & i & " ('" & auditName & "'): ocultada"
        Else
            Debug.Print "[SheetService]   - Audit Trail " & i & " ('" & auditName & "'): NO EXISTE (omitida)"
        End If
        On Error GoTo ErrorHandler
    Next i
    Debug.Print "[SheetService] PASO 2: Completado"
    
    ' Paso 2: Hacer visible la hoja destino (solo si NO es Audit Trail)
    Debug.Print "[SheetService] PASO 3: Haciendo visible hoja destino '" & sheetName & "'..."
    If Not IsAuditSheet(sheetName) Then
        On Error Resume Next
        ThisWorkbook.Sheets(sheetName).Visible = xlSheetVisible
        If Err.Number <> 0 Then
            Debug.Print "[SheetService] ⚠ ERROR al hacer visible '" & sheetName & "': N°" & Err.Number & " - " & Err.Description
            Debug.Print "[SheetService]   Posible causa: la hoja '" & sheetName & "' NO EXISTE en el libro"
        Else
            Debug.Print "[SheetService]   ✔ Hoja '" & sheetName & "' visible exitosamente"
        End If
        On Error GoTo ErrorHandler
    Else
        Debug.Print "[SheetService]   - Hoja '" & sheetName & "' es Audit Trail (no se hace visible aquí)"
    End If
    Debug.Print "[SheetService] PASO 3: Completado"
    
    ' Paso 3: Ocultar todas las demás hojas excepto destino y "Menú principal"
    Debug.Print "[SheetService] PASO 4: Ocultando hojas no destino..."
    Dim hojasOcultadas As Long
    hojasOcultadas = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> sheetName And ws.Name <> "Menú principal" And Not IsAuditSheet(ws.Name) Then
            ws.Visible = xlSheetVeryHidden
            hojasOcultadas = hojasOcultadas + 1
        End If
    Next ws
    Debug.Print "[SheetService]   ✔ " & hojasOcultadas & " hojas ocultadas (excluyendo destino y Menú principal)"
    Debug.Print "[SheetService] PASO 4: Completado"
    
    ' Paso 4: Asegurar que "Menú principal" siempre esté visible
    Debug.Print "[SheetService] PASO 5: Asegurando visibilidad de 'Menú principal'..."
    On Error Resume Next
    ThisWorkbook.Sheets("Menú principal").Visible = xlSheetVisible
    If Err.Number <> 0 Then
        Debug.Print "[SheetService] ⚠ ERROR: No se pudo hacer visible 'Menú principal': N°" & Err.Number & " - " & Err.Description
    Else
        Debug.Print "[SheetService]   ✔ 'Menú principal' visible"
    End If
    On Error GoTo ErrorHandler
    Debug.Print "[SheetService] PASO 5: Completado"
    
    ' Paso 5: Aplicar protección según rol a la hoja destino
    ' (C07 - La protección se aplica SIEMPRE, independientemente de los eventos Worksheet_Activate)
    Debug.Print "[SheetService] PASO 6: Aplicando protección según rol a '" & sheetName & "'..."
    On Error Resume Next
    Call SheetProtector2.ApplyRoleBasedProtection( _
        ThisWorkbook.Sheets(sheetName), _
        Configuration2.APP_PASSWORD _
    )
    If Err.Number <> 0 Then
        Debug.Print "[SheetService] ⚠ ERROR al aplicar protección: N°" & Err.Number & " - " & Err.Description
    Else
        Debug.Print "[SheetService]   ✔ Protección aplicada según rol '" & m_userRole & "'"
    End If
    On Error GoTo ErrorHandler
    Debug.Print "[SheetService] PASO 6: Completado"
    
    ' Paso 6: Re-proteger estructura del libro (AÚN con ScreenUpdating=False para evitar parpadeo)
    Debug.Print "[SheetService] PASO 7: Re-protegiendo estructura del libro..."
    Call WorkbookProtector2.ProtectWorkbook
    Debug.Print "[SheetService] PASO 7: Completado"
    
    ' NOTA: ScreenUpdating/DisplayAlerts/EnableEvents NO se restauran aquí
    ' porque el llamador (NavigateToSheet) maneja el bloque completo.
    ' Si se restauraran aquí, causarían parpadeo visual al hacer .Select después.
    
    Debug.Print "[SheetService.HideAndProtectAllSheetsExcept] FIN EXITOSO"
    Debug.Print "═══════════════════════════════════════════════════════════"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Debug.Print "[SheetService.HideAndProtectAllSheetsExcept] *** ERROR CAPTURADO ***"
    Debug.Print "[SheetService]   N° de error: " & Err.Number
    Debug.Print "[SheetService]   Descripción: " & Err.Description
    Debug.Print "[SheetService]   Fuente: " & Err.Source
    Debug.Print "[SheetService]   Hoja destino: '" & sheetName & "'"
    Debug.Print "[SheetService]   Rol: '" & m_userRole & "'"
    Debug.Print "[SheetService]   Contraseña usada: '" & Configuration2.APP_PASSWORD & "'"
    Debug.Print "[SheetService] *******************************"
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
' Lógica:
'   1. Desprotege estructura del libro.
'   2. Oculta todas las hojas que no son de auditoría ni "Menú principal".
'   3. Hace visible cada hoja Audit Trail que exista en el libro.
'   4. Asegura que "Menú principal" esté visible.
'   5. Selecciona la primera hoja de Audit Trail.
'   6. Re-protege estructura del libro.
' ----------------------------------------------------------------------
Public Sub ShowAuditTrailGroup()
    Debug.Print "═══════════════════════════════════════════════════════════"
    Debug.Print "[SheetService.ShowAuditTrailGroup] INICIO"
    Debug.Print "[SheetService] Rol actual del usuario: '" & m_userRole & "'"
    Debug.Print "[SheetService] AUDIT_PASSWORD en Configuration2: '" & Configuration2.AUDIT_PASSWORD & "' (" & Len(Configuration2.AUDIT_PASSWORD) & " caracteres)"
    Debug.Print "───────────────────────────────────────────────────────────"
    
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Call WorkbookProtector2.UnprotectWorkbook
    
    ' Verificar que el libro se desprotegió correctamente
    If ThisWorkbook.ProtectStructure Then
        Debug.Print "[SheetService] ⚠ AVISO: La estructura del libro SIGUE protegida después de UnprotectWorkbook"
    Else
        Debug.Print "[SheetService] ✔ Estructura del libro desprotegida correctamente"
    End If
    
    Dim ws As Worksheet
    Dim i As Long
    Dim auditName As String
    
    ' Paso 1: ocultar todas las hojas que no son de auditoría ni Menú principal.
    Debug.Print "[SheetService] PASO 1: Ocultando hojas no-Audit..."
    Dim hojasOcultadas As Long
    hojasOcultadas = 0
    For Each ws In ThisWorkbook.Worksheets
        If Not IsAuditSheet(ws.Name) And ws.Name <> "Menú principal" Then
            ws.Visible = xlSheetVeryHidden
            hojasOcultadas = hojasOcultadas + 1
        End If
    Next ws
    Debug.Print "[SheetService]   ✔ " & hojasOcultadas & " hojas ocultadas"
    
    ' Paso 2: mostrar cada hoja de Audit Trail que exista en el libro.
    Debug.Print "[SheetService] PASO 2: Mostrando hojas Audit Trail..."
    Dim hojasVisibles As Long
    hojasVisibles = 0
    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        auditName = AuditRotation2.ObtenerNombreHoja(i)
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(auditName)
        On Error GoTo ErrorHandler
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVisible
            Debug.Print "[SheetService]   ✔ " & auditName & ": visible"
            hojasVisibles = hojasVisibles + 1
        Else
            Debug.Print "[SheetService]   - " & auditName & ": NO EXISTE"
        End If
    Next i
    Debug.Print "[SheetService]   Total: " & hojasVisibles & " hojas Audit Trail visibles"
    
    ' Paso 3: asegurar que "Menú principal" esté visible.
    Debug.Print "[SheetService] PASO 3: Asegurando visibilidad de 'Menú principal'..."
    On Error Resume Next
    ThisWorkbook.Sheets("Menú principal").Visible = xlSheetVisible
    If Err.Number <> 0 Then
        Debug.Print "[SheetService] ⚠ ERROR: No se pudo hacer visible 'Menú principal'"
    Else
        Debug.Print "[SheetService]   ✔ 'Menú principal' visible"
    End If
    On Error GoTo ErrorHandler
    
    ' Paso 4: Aplicar protección a todas las hojas Audit Trail según rol
    ' (C07 - Las hojas Audit Trail quedan protegidas contra ediciones no autorizadas)
    Debug.Print "[SheetService] PASO 4: Aplicando protección a hojas Audit Trail..."
    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        auditName = AuditRotation2.ObtenerNombreHoja(i)
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(auditName)
        If Not ws Is Nothing Then
            Call SheetProtector2.ApplyRoleBasedProtection(ws, Configuration2.AUDIT_PASSWORD)
            Debug.Print "[SheetService]   ✔ " & auditName & ": protegida"
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' Paso 5: activar la primera hoja de Audit Trail.
    Debug.Print "[SheetService] PASO 5: Activando primera hoja Audit Trail..."
    Dim primerAudit As String
    primerAudit = AuditRotation2.ObtenerNombreHoja(1) ' Obtiene "Audit trail 1"
    
    On Error Resume Next
    ThisWorkbook.Sheets(primerAudit).Select
    If Err.Number <> 0 Then
        Debug.Print "[SheetService]   - No se pudo seleccionar '" & primerAudit & "', buscando alternativa..."
        ' Si no existe la primera, intentar activar cualquiera que esté visible
        For Each ws In ThisWorkbook.Worksheets
            If IsAuditSheet(ws.Name) And ws.Visible = xlSheetVisible Then
                ws.Select
                Debug.Print "[SheetService]   ✔ Seleccionada: " & ws.Name
                Exit For
            End If
        Next ws
    Else
        Debug.Print "[SheetService]   ✔ Seleccionada: " & primerAudit
    End If
    On Error GoTo ErrorHandler
    
    Call WorkbookProtector2.ProtectWorkbook
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    Debug.Print "[SheetService.ShowAuditTrailGroup] FIN EXITOSO"
    Debug.Print "═══════════════════════════════════════════════════════════"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Debug.Print "[SheetService.ShowAuditTrailGroup] *** ERROR CAPTURADO ***"
    Debug.Print "[SheetService]   N°: " & Err.Number & " - " & Err.Description
    Call WorkbookProtector2.ProtectWorkbook
    Call ErrorLogger2.Log("SheetService.ShowAuditTrailGroup", VBA.Err.Description, VBA.Err.Number)
    MsgBox "Ocurrió un error al mostrar las hojas de Audit trail.", vbCritical, "Error de Navegación"
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
