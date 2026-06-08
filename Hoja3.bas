'' ----------------------------------------------------------------------
' Módulo: Hoja3 ("Checklist")
' Descripción: Eventos para la hoja Checklist (configuración de plantillas).
'              Contiene las tablas maestras: tblPlantillas, tblPreguntas,
'              tblSecciones, tblOpcionesDeRespuesta.
'              M02: Protección y auditoría centralizada.
'
' INSTRUCCIONES DE INSTALACIÓN:
' 1. Abre el VBA Editor (Alt+F11)
' 2. Busca en el árbol de la izquierda: Microsoft Excel Objetos → Hoja "Checklist"
' 3. Haz doble clic en esa hoja para abrir su módulo
' 4. Copia TODO el código de este archivo
' 5. Pégalo en el módulo de la hoja "Checklist"
' 6. Guarda el archivo
' ----------------------------------------------------------------------
Option Explicit

Private m_oldValues As Variant ' Almacena los valores previos a un cambio para auditoría

'' ----------------------------------------------------------------------
' Evento: Worksheet_Activate
' Propósito: Se ejecuta cuando se activa la hoja Checklist.
'            Aplica protección centralizada según el rol del usuario (M02).
' ----------------------------------------------------------------------
Private Sub Worksheet_Activate()
    On Error GoTo ErrorHandler
    
    ' Aplicar protección centralizada según el rol del usuario (M02)
    ' Admin → desprotegido (puede editar plantillas, preguntas, secciones, opciones)
    ' Usuario → solo lectura con copiado
    ' Otros → sin selección
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.APP_PASSWORD)
    
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("Checklist.Worksheet_Activate", VBA.Err.Description, VBA.Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_Deactivate
' Propósito: Se ejecuta al salir de la hoja Checklist.
'            Aplica protección centralizada según el rol al salir (M02).
' ----------------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    On Error Resume Next
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.APP_PASSWORD)
    On Error GoTo 0
End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_Change
' Propósito: Audita cualquier cambio realizado en las tablas maestras
'            de Checklist, registrando los valores anteriores y posteriores.
' Tablas auditadas: tblPlantillas, tblPreguntas, tblSecciones, tblOpcionesDeRespuesta
' ----------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Dim tablesToAudit As Variant
    tablesToAudit = Array( _
        Configuration2.TABLE_PLANTILLAS, _
        Configuration2.TABLE_PREGUNTAS, _
        Configuration2.TABLE_SECCIONES, _
        Configuration2.TABLE_OPCIONES _
    )
    Call TableAuditor2.AuditTableChanges(Me, Target, tablesToAudit, m_oldValues)
    
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Call ErrorLogger2.Log("Worksheet_Change (Checklist)", VBA.Err.Description, VBA.Err.Number)
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
    Call ErrorLogger2.Log("Worksheet_SelectionChange (Checklist)", VBA.Err.Description, VBA.Err.Number)
End Sub
