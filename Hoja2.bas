'' ----------------------------------------------------------------------
' Módulo: Hoja2 ("Configuración")
' Descripción: Eventos para la hoja Configuración (tablas maestras del sistema).
'              Incluye protección por rol centralizada vía SheetProtector2.
'
' INSTRUCCIONES DE INSTALACIÓN:
' 1. Abre el VBA Editor (Alt+F11)
' 2. Busca en el árbol de la izquierda: Microsoft Excel Objetos → Hoja "Configuración"
' 3. Haz doble clic en esa hoja para abrir su módulo
' 4. Copia TODO el código de este archivo
' 5. Pégalo en el módulo de la hoja "Configuración"
' 6. Guarda el archivo
' ----------------------------------------------------------------------
Option Explicit

Private m_oldValues As Variant ' Almacena los valores previos a un cambio para auditoría

'' ----------------------------------------------------------------------
' Evento: Worksheet_Activate
' Propósito: Se ejecuta cuando se activa la hoja Configuración.
'            Aplica protección centralizada según el rol del usuario.
' ----------------------------------------------------------------------
Private Sub Worksheet_Activate()
    ' ## NAVEGACIÓN ## Guardia: evitar doble ejecución durante navegación (FASE 4, 08/06/2026)
    If g_NavigationInProgress Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    ' Aplicar protección centralizada según el rol del usuario (C03)
    ' Admin → desprotegido (edición libre)
    ' Usuario → solo lectura con copiado
    ' Otros → sin selección
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.APP_PASSWORD)
    
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("Configuración.Worksheet_Activate", VBA.Err.Description, VBA.Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_Deactivate
' Propósito: Se ejecuta al salir de la hoja Configuración.
'            Aplica protección centralizada según el rol al salir.
' ----------------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    On Error Resume Next
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.APP_PASSWORD)
    On Error GoTo 0
End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_Change
' Propósito: Audita cualquier cambio realizado en las tablas de la hoja,
'            registrando los valores anteriores y posteriores.
' Fecha: 16/06/2026 — Ampliado para auditar TODAS las tablas de Configuración.
' ----------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    ' Auditar TODAS las tablas maestras de la hoja Configuración
    Dim tablesToAudit As Variant
    tablesToAudit = Array( _
        Configuration2.TABLE_CONTROL_DE_CAMBIOS, _
        Configuration2.TABLE_CRITICIDAD, _
        Configuration2.TABLE_CATEGORIAS_RPN, _
        Configuration2.TABLE_RANGOS_RECUPERACION, _
        Configuration2.TABLE_RANGOS_OOL, _
        Configuration2.TABLE_CONFIGURACION, _
        Configuration2.TABLE_EQUIPOS, _
        Configuration2.TABLE_PUESTO, _
        Configuration2.TABLE_PUESTOS_INICIALES, _
        Configuration2.TABLE_PLANTA, _
        Configuration2.TABLE_RESUMEN_CRONOGRAMA _
    )
    Call TableAuditor2.AuditTableChanges(Me, Target, tablesToAudit, m_oldValues)
    
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("Worksheet_Change (Configuración)", VBA.Err.Description, VBA.Err.Number)
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
    Call ErrorLogger2.Log("Worksheet_SelectionChange (Configuración)", VBA.Err.Description, VBA.Err.Number)
End Sub
