
'' ----------------------------------------------------------------------
' Módulo: Resultados (Hoja de Detecciones)
'' ----------------------------------------------------------------------
Option Explicit
Private m_oldValues As Variant ' Almacena los valores previos a un cambio para auditoría


'' ----------------------------------------------------------------------
' Evento: Worksheet_Change
' Propósito: Audita cualquier cambio realizado en la hoja, registrando los valores
'            anteriores y posteriores en la tabla de auditoría.
' Lógica:
'   1. Desactiva el refresco de pantalla para mejorar el rendimiento.
'   2. Define la tabla a auditar (tblResultados).
'   3. Llama a TableAuditor.AuditTableChanges con los valores previos.
'   4. Reactiva el refresco de pantalla y maneja errores.
'' ----------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Dim tablesToAudit As Variant
    tablesToAudit = Array("tblDetecciones")
    Call TableAuditor2.AuditTableChanges(Me, Target, tablesToAudit, m_oldValues)
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Call ErrorLogger2.Log("Worksheet_Change (" & Me.Name & ")", VBA.Err.Description, VBA.Err.Number)
End Sub


'' ----------------------------------------------------------------------
' Evento: Worksheet_SelectionChange
' Propósito: Almacena los valores seleccionados antes de un cambio para
'            permitir la auditoría precisa de modificaciones.
' Lógica:
'   1. Limita la auditoría a selecciones menores a 3000 celdas.
'   2. Si la selección es válida, almacena los valores; si no, limpia m_oldValues.
'   3. Maneja errores y registra en el log si ocurre alguno.
'' ----------------------------------------------------------------------
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo ErrorHandler
    Const SELECTION_LIMIT As Long = 3000
    If Target.Cells.CountLarge < SELECTION_LIMIT Then
        m_oldValues = Target.Value
    Else
        m_oldValues = Empty
    End If
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("Worksheet_SelectionChange (" & Me.Name & ")", VBA.Err.Description, VBA.Err.Number)
End Sub


'' ----------------------------------------------------------------------
' Evento: Worksheet_Activate
' Propósito: Controla el acceso a la hoja según el rol del usuario.
'            Solo los administradores pueden editar; los demás solo pueden ver.
' Lógica:
'   1. Si el usuario es Admin, desprotege la hoja.
'   2. Si no, protege la hoja para evitar edición.
'   3. Maneja errores y registra en el log si ocurre alguno.
'' ----------------------------------------------------------------------
Private Sub Worksheet_Activate()
    On Error GoTo ErrorHandler
    ' FASE 1 (22/02/2026): La lógica de protección por rol está centralizada en
    ' SheetProtector2.ApplyRoleBasedProtection (DRY). Antes era un If/ElseIf inline
    ' aquí que venía siendo ignorado cuando SaveDataToTable llamaba a ProtectSheet
    ' directamente (sin rol), pisando el permiso de copiado del rol "Usuario".
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.APP_PASSWORD)
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("Configuración.Worksheet_Activate", VBA.Err.Description, VBA.Err.Number)
End Sub


'' ----------------------------------------------------------------------
' Evento: Worksheet_Deactivate
' Propósito: (Opcional) Puede usarse para proteger la hoja al salir.
'            Actualmente está comentado para permitir flexibilidad.
'' ----------------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    'Call SheetProtector.ProtectSheet(Me, Configuration.APP_PASSWORD)
End Sub