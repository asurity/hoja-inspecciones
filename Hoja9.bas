
' ----------------------------------------------------------------------
' Módulo: AuditTrail (Hoja de auditoría)
' Descripción: Controla la protección de la hoja de auditoría para garantizar que
'              siempre permanezca bloqueada con la contraseña de auditoría, incluso
'              para usuarios administradores. No permite edición, solo visualización.
' Dependencias:
'   - SheetProtector: Para proteger la hoja.
'   - Configuration: Para obtener la contraseña de auditoría.
'   - ErrorLogger: Para registrar errores en caso de fallo.
' ----------------------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------------
' Evento: Worksheet_Activate
' Propósito: Protege la hoja de auditoría cada vez que se activa, usando la contraseña
'            de auditoría. No verifica el rol del usuario para asegurar máxima seguridad.
' ----------------------------------------------------------------------
Private Sub Worksheet_Activate()
    On Error GoTo ErrorHandler
    Call SheetProtector2.ProtectSheet(Me, Configuration2.AUDIT_PASSWORD)
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("Audit trail.Worksheet_Activate", VBA.Err.Description, VBA.Err.Number)
End Sub

' ----------------------------------------------------------------------
' Evento: Worksheet_Deactivate
' Propósito: Vuelve a proteger la hoja de auditoría al salir de ella, reforzando la
'            seguridad ante cualquier posible cambio o desbloqueo accidental.
' ----------------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    Call SheetProtector2.ProtectSheet(Me, Configuration2.AUDIT_PASSWORD)
End Sub