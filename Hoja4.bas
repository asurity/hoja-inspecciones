'' ----------------------------------------------------------------------
' Módulo: Hoja "Aseguramiento de calidad"
' Descripción: Eventos para la hoja Aseguramiento de calidad (Personal QA).
'              Contiene la tabla tblAseguramientoCalidad con personal de QA.
'
' INSTRUCCIONES DE INSTALACIÓN:
' 1. Abre el VBA Editor (Alt+F11)
' 2. Busca en el árbol de la izquierda: Microsoft Excel Objetos ? Hoja "Aseguramiento de calidad"
' 3. Haz doble clic en esa hoja para abrir su módulo
' 4. Copia TODO el código de este archivo
' 5. Pégalo en el módulo de la hoja "Aseguramiento de calidad"
' 6. Guarda el archivo
' ----------------------------------------------------------------------
Option Explicit

Private m_oldValues As Variant ' Almacena los valores previos a un cambio para auditoría

'' ----------------------------------------------------------------------
' Evento: Worksheet_Change
' Propósito: Audita cualquier cambio realizado en tblAseguramientoCalidad,
'            registrando los valores anteriores y posteriores.
' ----------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Dim tablesToAudit As Variant
    tablesToAudit = Array(Configuration2.TABLE_ASEGURAMIENTO)
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
' ----------------------------------------------------------------------
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
' Propósito: Se ejecuta cuando se activa la hoja Aseguramiento de calidad.
'            Aplica protección según el rol del usuario.
' ----------------------------------------------------------------------
Private Sub Worksheet_Activate()
    On Error GoTo ErrorHandler
    
    ' Aplicar protección según el rol
    If ThisWorkbook.GetUserRole() = "Admin" Then
        ' Admin puede editar
        Call SheetProtector2.UnprotectSheet(Me, Configuration2.APP_PASSWORD)
    Else
        ' Usuario solo lectura
        Call SheetProtector2.ProtectSheet(Me, Configuration2.APP_PASSWORD)
    End If
    
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("Aseguramiento de calidad.Worksheet_Activate", VBA.Err.Description, VBA.Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_Deactivate
' Propósito: Se ejecuta al salir de la hoja Aseguramiento de calidad.
'            Protege la hoja al salir.
' ----------------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    On Error Resume Next
    Call SheetProtector2.ProtectSheet(Me, Configuration2.APP_PASSWORD)
    On Error GoTo 0
End Sub