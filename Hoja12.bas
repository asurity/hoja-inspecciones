'' ----------------------------------------------------------------------
' Módulo: Hoja12 ("Audit trail 3")
' Descripción: Eventos para la hoja Audit trail 3.
'              Protección con contraseña exclusiva de auditoría (AUDIT_PASSWORD).
'              No permite edición bajo ningún rol de usuario.
'              La visibilidad se gestiona centralizadamente por SheetService2.
' ## NAVEGACIÓN ## — Hoja del grupo Audit Trail (5 hojas simultáneas)
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Evento: Worksheet_Activate
' Propósito: Protege la hoja de auditoría con contraseña exclusiva.
'            No verifica rol del usuario — máxima seguridad para audit trail.
' ----------------------------------------------------------------------
Private Sub Worksheet_Activate()
    On Error GoTo ErrorHandler
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.AUDIT_PASSWORD)
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("Audit trail 3.Worksheet_Activate", VBA.Err.Description, VBA.Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_Deactivate
' Propósito: Refuerza la protección al salir de la hoja de auditoría.
' ----------------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    On Error Resume Next
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.AUDIT_PASSWORD)
    On Error GoTo 0
End Sub
