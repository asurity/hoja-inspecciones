
' ----------------------------------------------------------------------
' Módulo: NavigationService
' Descripción: Servicio centralizado de navegación entre hojas del sistema.
'              Gestiona ScreenUpdating/EnableEvents/DisplayAlerts en UN solo lugar
'              (BeginNavigation/EndNavigation) y delega visibilidad a SheetService2.ShowOnly.
' Refactorizado: 08/06/2026 - FASE 2 Refactorización Navegación
' Dependencias:
'   - SheetService2: ShowOnly, ShowAuditTrailGroup
'   - WorkbookProtector2: ToggleProtection, ProtectWorkbook, UnprotectWorkbook
'   - Configuration2: Constantes de nombres de hojas
'   - AuditLogger2: LogAction (solo para módulos sensibles)
'   - ErrorLogger2: Registro de errores
'   - VariablesGlobales2: g_NavigationInProgress
' ----------------------------------------------------------------------
Option Explicit


' ======================================================================
' ## NAVEGACIÓN ## MOTOR CENTRAL DE SUPRESIÓN DE PARPADEO
' ÚNICOS lugares en todo el proyecto que tocan ScreenUpdating,
' EnableEvents y DisplayAlerts (FASE 2, 08/06/2026).
' ======================================================================

Private Sub BeginNavigation()
    g_NavigationInProgress = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

Private Sub EndNavigation()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    g_NavigationInProgress = False
End Sub


' ======================================================================
' ## NAVEGACIÓN ## NAVEGACIÓN GENÉRICA A HOJA
' ======================================================================

' ----------------------------------------------------------------------
' Subrutina: NavigateToSheet
' Propósito: Navega a una hoja específica usando el nuevo motor centralizado.
'            Flujo: BeginNavigation → Unprotect → ShowOnly → Select → Protect → EndNavigation.
' Refactorizado: 08/06/2026 - FASE 2 (BeginNavigation/EndNavigation + ShowOnly)
' ----------------------------------------------------------------------
Public Sub NavigateToSheet(ByVal targetSheetName As String)
    On Error GoTo ErrorHandler
    
    ' Registrar navegación ANTES de BeginNavigation (que desactiva eventos)
    On Error Resume Next
    Call AuditLogger2.LogAction( _
        action:="Navegación a módulo", _
        sheetName:=targetSheetName, _
        dataModified:="Acceso a módulo", _
        beforeChange:="Hoja anterior: " & g_PreviousSheetName, _
        afterChange:="Hoja destino: " & targetSheetName, _
        moduleAndSubroutine:="NavigationService2.NavigateToSheet" _
    )
    g_PreviousSheetName = targetSheetName
    On Error GoTo 0
    
    BeginNavigation
    
    ' Verificar que la hoja destino existe
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo ErrorHandler
    If ws Is Nothing Then
        EndNavigation
        MsgBox "Error: La hoja '" & targetSheetName & "' no existe.", vbCritical, "Error de Navegación"
        Exit Sub
    End If
    
    ' Desproteger estructura → ocultar/mostrar hojas → seleccionar → reproteger
    Call WorkbookProtector2.UnprotectWorkbook
    Call SheetService2.ShowOnly(False, targetSheetName)
    
    On Error Resume Next
    ws.Select
    On Error GoTo ErrorHandler
    
    ' ====== ESTA ES LA LÍNEA DE SEGURIDAD QUE DEBES AGREGAR ======
    ' Forza a aplicar la protección según el rol antes de cerrar la navegación
    Call SheetProtector2.ApplyRoleBasedProtection(ws, Configuration2.APP_PASSWORD)
    ' =============================================================
    
    Call WorkbookProtector2.ProtectWorkbook
    EndNavigation
    Exit Sub
    
ErrorHandler:
    Dim navErrNum As Long, navErrDesc As String
    navErrNum = Err.Number
    navErrDesc = Err.Description
    EndNavigation
    Call WorkbookProtector2.ProtectWorkbook
    Call ErrorLogger2.Log("NavigationService.NavigateToSheet", "Error navegando a: " & targetSheetName & " | " & navErrDesc, navErrNum)
    MsgBox "Error: No se pudo navegar a la hoja '" & targetSheetName & "'.", vbCritical, "Error de Navegación"
End Sub


' ======================================================================
' ## NAVEGACIÓN ## SUBS INDIVIDUALES MostrarXxx (PATRÓN NUEVO)
' Creados: 08/06/2026 - FASE 2
' Cada uno es la fachada limpia para cada módulo del sistema.
' ======================================================================

Public Sub MostrarConfiguracion()
    ' LogAction ahora se hace dentro de NavigateToSheet (evita duplicados)
    Call NavigateToSheet("Configuración")
End Sub

Public Sub MostrarPersonal()
    Call NavigateToSheet(Configuration2.SHEET_PERSONAL)
End Sub

Public Sub MostrarAseguramientoCalidad()
    Call NavigateToSheet(Configuration2.SHEET_ASEGURAMIENTO)
End Sub

Public Sub MostrarMenu()
    On Error GoTo ErrorHandler
    
    ' Registrar navegación al menú principal
    On Error Resume Next
    Call AuditLogger2.LogAction( _
        action:="Navegación a módulo", _
        sheetName:="Menú principal", _
        dataModified:="Regreso al menú", _
        beforeChange:="Hoja anterior: " & g_PreviousSheetName, _
        afterChange:="Hoja destino: Menú principal", _
        moduleAndSubroutine:="NavigationService2.MostrarMenu" _
    )
    g_PreviousSheetName = "Menú principal"
    On Error GoTo 0
    
    BeginNavigation
    Call WorkbookProtector2.UnprotectWorkbook
    Call SheetService2.ShowOnly(False)
    On Error Resume Next
    ThisWorkbook.Sheets("Menú principal").Select
    On Error GoTo ErrorHandler
    Call WorkbookProtector2.ProtectWorkbook
    EndNavigation
    Exit Sub
ErrorHandler:
    EndNavigation
    Call WorkbookProtector2.ProtectWorkbook
End Sub

Public Sub MostrarChecklistVirtual()
    On Error GoTo ErrorHandler
    Call AuditLogger2.LogAction("Apertura Selector Inspección", "Formulario", _
        "frmSelectorInspeccion", "N/A", "Usuario solicitó nueva inspección", _
        "NavigationService2.MostrarChecklistVirtual")
    Dim frmSelector As frmSelectorInspeccion
    Set frmSelector = New frmSelectorInspeccion
    frmSelector.Show vbModal
    Set frmSelector = Nothing
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("NavigationService.MostrarChecklistVirtual", _
        "Error al abrir selector: " & Err.Description, Err.Number)
    MsgBox "Error: No se pudo abrir el selector de inspección.", vbCritical, "Error"
End Sub

Public Sub MostrarResultados()
    Call NavigateToSheet(Configuration2.SHEET_HISTORICO)
End Sub

Public Sub MostrarConfiguracionChecklist()
    Call NavigateToSheet("Checklist")
End Sub

Public Sub MostrarCronograma()
    Call NavigateToSheet(Configuration2.SHEET_CRONOGRAMA)
End Sub

Public Sub MostrarPlantillaCertificado()
    Call NavigateToSheet(Configuration2.SHEET_PLANTILLA_CERTIFICADO)
End Sub

Public Sub MostrarGraficos()
    Call NavigateToSheet(Configuration2.SHEET_GRAFICOS)
End Sub

' ----------------------------------------------------------------------
' MostrarAuditTrail (NUEVO PATRÓN)
' Propósito: Muestra el grupo completo de 5 hojas Audit Trail.
'            Flujo: BeginNavigation → Unprotect → ocultar todo → ShowAuditTrailGroup → Protect → EndNavigation.
'            NO registra en Audit Trail (evita recursión — ya lo hace Workbook_SheetActivate).
' Refactorizado: 08/06/2026 - FASE 2
' ----------------------------------------------------------------------
Public Sub MostrarAuditTrail()
    On Error GoTo ErrorHandler
    
    ' Registrar navegación al grupo Audit Trail
    On Error Resume Next
    Call AuditLogger2.LogAction( _
        action:="Navegación a módulo", _
        sheetName:="Audit Trail (grupo)", _
        dataModified:="Consulta de auditoría", _
        beforeChange:="Hoja anterior: " & g_PreviousSheetName, _
        afterChange:="Hoja destino: Audit Trail (grupo)", _
        moduleAndSubroutine:="NavigationService2.MostrarAuditTrail" _
    )
    g_PreviousSheetName = "Audit trail 1"
    On Error GoTo 0
    
    BeginNavigation
    Call WorkbookProtector2.UnprotectWorkbook
    
    ' Ocultar todas las hojas de módulo primero (solo Menú principal queda visible)
    Call SheetService2.ShowOnly(False)
    
    ' Mostrar las 5 hojas Audit Trail y seleccionar la primera
    Call SheetService2.ShowAuditTrailGroup
    
    Call WorkbookProtector2.ProtectWorkbook
    EndNavigation
    Exit Sub
    
ErrorHandler:
    EndNavigation
    Call WorkbookProtector2.ProtectWorkbook
    Call ErrorLogger2.Log("NavigationService.MostrarAuditTrail", _
        "Error navegando al grupo Audit Trail: " & Err.Description, Err.Number)
    MsgBox "Error: No se pudo navegar a Audit Trail.", vbCritical, "Error de Navegación"
End Sub


' ======================================================================
' ## NAVEGACIÓN ## WRAPPERS LEGACY (COMPATIBILIDAD CON BOTONES EXISTENTES)
' Cada NavigateToXxx antiguo ahora llama al nuevo MostrarXxx.
' ======================================================================

Public Sub NavigateToConfiguracion()
    Call MostrarConfiguracion
End Sub

Public Sub NavigateToPersonalProduccion()
    Call MostrarPersonal
End Sub

Public Sub NavigateToAseguramientoCalidad()
    Call MostrarAseguramientoCalidad
End Sub

Public Sub NavigateToAuditTrail()
    Call MostrarAuditTrail
End Sub

Public Sub NavigateToMenu()
    Call MostrarMenu
End Sub

Public Sub NavigateToChecklistVirtual()
    Call MostrarChecklistVirtual
End Sub

Public Sub NavigateToResultados()
    Call MostrarResultados
End Sub

Public Sub NavigateToConfiguracionChecklist()
    Call MostrarConfiguracionChecklist
End Sub

Public Sub NavigateToCronograma()
    Call MostrarCronograma
End Sub

Public Sub NavigateToPlantillaCertificado()
    Call MostrarPlantillaCertificado
End Sub

Public Sub NavigateToGraficos()
    Call MostrarGraficos
End Sub

' ----------------------------------------------------------------------
' WRAPPERS LEGACY — Hojas que YA NO EXISTEN en el libro
' Mantenidos por compatibilidad con botones antiguos que puedan romperse.
' Cada uno navega a su nombre original (aunque no exista).
' ----------------------------------------------------------------------

Public Sub NavigateToDetecciones()
    Call NavigateToSheet("Detecciones")
End Sub

Public Sub NavigateToDashboard()
    Call NavigateToSheet("Dashboard")
End Sub

Public Sub NavigateToObservaciones()
    Call NavigateToSheet("Observaciones")
End Sub

Public Sub NavigateToRechazo()
    Call NavigateToSheet("Análisis Rechazo")
End Sub

Public Sub NavigateToDesvio()
    Call NavigateToSheet("Análisis Desvío")
End Sub

Public Sub NavigateToTecControlProceso()
    Call NavigateToSheet("TecControlProceso")
End Sub

Public Sub NavigateToControlDeCambios()
    Call NavigateToSheet("Control de cambios")
End Sub
