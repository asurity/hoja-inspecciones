
'' ----------------------------------------------------------------------
' Módulo: NavigationService
' Descripción: Proporciona subrutinas para navegar de manera segura entre las diferentes
'              hojas de cálculo de la aplicación. Centraliza la lógica de selección de hojas
'              y la integra con el sistema de protección para un flujo de trabajo controlado.
' Dependencias:
'   - SheetProtector2: Para aplicar protección según el rol del usuario.
'   - Configuration2: Para las contraseñas de cada hoja.
'   - ErrorLogger2: Para registrar errores de navegación.
'' ----------------------------------------------------------------------
Option Explicit


'' ----------------------------------------------------------------------
' Subrutina: NavigateToSheet
' Propósito: Navega a una hoja específica y aplica la protección correcta según el rol:
'            - Admin   ? sin protección (puede editar todo).
'            - Usuario ? protección de solo lectura (puede ver y copiar, no editar).
' Parámetros:
'   - targetSheetName: Nombre de la hoja de destino.
'' ----------------------------------------------------------------------
Public Sub NavigateToSheet(ByVal targetSheetName As String)
    On Error GoTo ErrorHandler
    ' HideAndProtectAllSheetsExcept oculta todo lo demás, hace visible la hoja
    ' destino y aplica la protección correcta según el rol activo. También
    ' envuelve el cambio de visibilidad con Unprotect/ProtectWorkbook para
    ' evitar Error 1004 cuando la estructura del libro está protegida.
    Call SheetService2.HideAndProtectAllSheetsExcept(targetSheetName)
    ThisWorkbook.Sheets(targetSheetName).Select
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("NavigationService.NavigateToSheet", "No se pudo navegar a la hoja: " & targetSheetName, VBA.Err.Number)
    MsgBox "Error: No se pudo navegar a la hoja '" & targetSheetName & "'.", vbCritical, "Error de Navegación"
End Sub


'' ----------------------------------------------------------------------
' Subrutinas de Navegación Específicas
' Propósito: Proporcionan puntos de entrada públicos y amigables para la navegación,
'            permitiendo ir a hojas específicas mediante nombres claros.
'            Se pueden agregar validaciones de permisos si es necesario.
'' ----------------------------------------------------------------------

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

Public Sub NavigateToConfiguracion()
    Call NavigateToSheet("Configuración")
End Sub

Public Sub NavigateToPersonalProduccion()
    Call NavigateToSheet("PersonalProduccion")
End Sub

Public Sub NavigateToTecControlProceso()
    Call NavigateToSheet("TecControlProceso")
End Sub

Public Sub NavigateToAuditTrail()
    On Error GoTo ErrorHandler
    ' Muestra las 5 hojas de Audit Trail simultáneamente y oculta el resto.
    Call SheetService2.ShowAuditTrailGroup
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("NavigationService.NavigateToAuditTrail", "No se pudo navegar al grupo Audit trail", VBA.Err.Number)
    MsgBox "Error: No se pudo navegar a Audit trail.", vbCritical, "Error de Navegación"
End Sub

Public Sub NavigateToControlDeCambios()
    Call NavigateToSheet("Control de cambios")
End Sub

Public Sub NavigateToMenu()
    ' Esta navegación está permitida para todos los usuarios, sin necesidad de validación.
    Call NavigateToSheet("Menú principal")
End Sub