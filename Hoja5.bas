'' ----------------------------------------------------------------------
' Módulo: Hoja "Personal"
' Descripción: Eventos para la hoja Personal (Personal de Producción).
'              Contiene la tabla tblPersonal con todo el personal del sistema.
'              Sincroniza en caliente el estado "Activo" con el Cronograma.
'
' INSTRUCCIONES DE INSTALACIÓN:
' 1. Abre el VBA Editor (Alt+F11)
' 2. Busca en el árbol de la izquierda: Microsoft Excel Objetos -> Hoja "Personal"
' 3. Haz doble clic en esa hoja para abrir su módulo
' 4. Borra el código viejo y pega TODO este bloque consolidado
' 5. Guarda el archivo
' ----------------------------------------------------------------------
Option Explicit

Private m_oldValues As Variant ' Almacena los valores previos a un cambio para auditoría

'' ----------------------------------------------------------------------
' Evento: Worksheet_Change
' Propósito: Audita cualquier cambio realizado en tblPersonal y
'            sincroniza el estado Activo/Inactivo con el Cronograma en tiempo real.
' Fecha: 16/06/2026 - Agregada sincronización en caliente con Cronograma.
' ----------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    ' 1. EJECUTAR AUDITORÍA EXISTENTE
    Dim tablesToAudit As Variant
    tablesToAudit = Array(Configuration2.TABLE_PERSONAL)
    Call TableAuditor2.AuditTableChanges(Me, Target, tablesToAudit, m_oldValues)
    
    ' 2. EJECUTAR SINCRONIZACIÓN EN CALIENTE CON EL CRONOGRAMA (NUEVO - 16/06/2026)
    Call DetectarYProcesarCambioActivo(Target)
    
    Exit Sub
ErrorHandler:
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
' Propósito: Se ejecuta cuando se activa la hoja Personal.
'            Aplica protección según el rol del usuario.
' ----------------------------------------------------------------------
Private Sub Worksheet_Activate()
    ' ## NAVEGACIÓN ## Guardia: evitar doble ejecución durante navegación (FASE 4, 08/06/2026)
    If g_NavigationInProgress Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    ' Aplicar protección centralizada según el rol del usuario
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.APP_PASSWORD)
    
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("Personal.Worksheet_Activate", VBA.Err.Description, VBA.Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_Deactivate
' Propósito: Se ejecuta al salir de la hoja Personal.
'            Aplica protección centralizada según el rol al salir.
' ----------------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    On Error Resume Next
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.APP_PASSWORD)
    On Error GoTo 0
End Sub


' ============================================================================
' SUBRUTINAS PRIVADAS INTERNAS: GESTIÓN DE SINCRONIZACIÓN EN CALIENTE
' ============================================================================

'' ----------------------------------------------------------------------
' Subrutina: DetectarYProcesarCambioActivo
' Propósito: Intercepta si el cambio ocurrió en la columna "Activo" de la
'            tabla de personal y gatilla la actualización correspondiente.
' Fecha: 16/06/2026
' ----------------------------------------------------------------------
Private Sub DetectarYProcesarCambioActivo(ByVal Target As Range)
    Dim tblPersonal As ListObject
    Dim colActivo As Long
    Dim colIniciales As Long
    Dim rngInterseccion As Range
    Dim celda As Range
    Dim filaOrigen As Range
    Dim inicialesEmpleado As String
    Dim nuevoEstado As String
    
    On Error Resume Next
    Set tblPersonal = Me.ListObjects(Configuration2.TABLE_PERSONAL)
    colActivo = tblPersonal.ListColumns("Activo").Index
    colIniciales = tblPersonal.ListColumns("Iniciales").Index
    On Error GoTo 0
    
    If tblPersonal Is Nothing Then Exit Sub
    If colActivo = 0 Or colIniciales = 0 Then Exit Sub
    
    ' Verificar si el cambio de celda intersecta la columna "Activo"
    Set rngInterseccion = Intersect(Target, tblPersonal.ListColumns(colActivo).DataBodyRange)
    If Not rngInterseccion Is Nothing Then
        
        ' Calcular columna absoluta de "Iniciales" en la hoja
        ' (ListColumns.Index es relativo a la tabla; EntireRow.Cells necesita
        '  columna absoluta de la hoja, no de la tabla)
        Dim colInicialesAbs As Long
        colInicialesAbs = tblPersonal.DataBodyRange.Columns(colIniciales).Column
        
        ' Desactivar eventos temporalmente para evitar loops infinitos durante escritura
        Application.EnableEvents = False
        
        ' Recorrer las celdas modificadas de la columna "Activo"
        For Each celda In rngInterseccion.Cells
            Set filaOrigen = celda.EntireRow
            inicialesEmpleado = Trim(CStr(filaOrigen.Cells(1, colInicialesAbs).Value))
            nuevoEstado = Trim(CStr(celda.Value))
            
            Debug.Print "[HojaPersonal] Cambio detectado en columna 'Activo': " & _
                         inicialesEmpleado & " -> " & nuevoEstado
            
            If Len(inicialesEmpleado) > 0 Then
                ' Actualizar la base de datos interna del cronograma
                Call SincronizarEstadoActivoEnCronograma(inicialesEmpleado, nuevoEstado)
            End If
        Next celda
        
        ' Reactivar los eventos del sistema
        Application.EnableEvents = True
        
        ' Actualizar el resumen visual en el Menú Principal en tiempo real
        Debug.Print "[HojaPersonal] Refrescando resumen cronograma..."
        On Error Resume Next
        Call CronogramaResumen.RefrescarResumenCronograma
        If Err.Number <> 0 Then
            Debug.Print "[HojaPersonal] ERROR al refrescar resumen: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Sub

'' ----------------------------------------------------------------------
' Subrutina: SincronizarEstadoActivoEnCronograma
' Propósito: Modifica la columna "Personal activo" en tblCronogramaInspecciones
'            para el empleado modificado, manejando la protección de la hoja.
' Fecha: 16/06/2026
' ----------------------------------------------------------------------
Private Sub SincronizarEstadoActivoEnCronograma(ByVal iniciales As String, ByVal estado As String)
    Dim wsCrono As Worksheet
    Dim tblCrono As ListObject
    Dim filaCrono As ListRow
    Dim colInicialesCrono As Long
    Dim colPersonalActivoCrono As Long
    
    On Error Resume Next
    Set wsCrono = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCrono = wsCrono.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    colInicialesCrono = tblCrono.ListColumns("Iniciales personal").Index
    colPersonalActivoCrono = tblCrono.ListColumns("Personal activo").Index
    On Error GoTo 0
    
    If tblCrono Is Nothing Then
        Debug.Print "[HojaPersonal] ERROR: No se encontró tblCronogramaInspecciones"
        Exit Sub
    End If
    If colInicialesCrono = 0 Or colPersonalActivoCrono = 0 Then
        Debug.Print "[HojaPersonal] ERROR: Columna 'Iniciales personal' o 'Personal activo' no encontrada en cronograma"
        Exit Sub
    End If
    
    ' Desproteger la hoja cronograma utilizando el sistema centralizado de contraseñas
    Call SheetProtector2.UnprotectSheet(wsCrono, Configuration2.APP_PASSWORD)
    
    ' Buscar y actualizar todas las instancias asociadas a este personal
    Dim filasActualizadas As Long
    filasActualizadas = 0
    
    If Not tblCrono.DataBodyRange Is Nothing Then
        For Each filaCrono In tblCrono.ListRows
            If Trim(CStr(filaCrono.Range.Cells(1, colInicialesCrono).Value)) = Trim(iniciales) Then
                filaCrono.Range.Cells(1, colPersonalActivoCrono).Value = estado
                filasActualizadas = filasActualizadas + 1
            End If
        Next filaCrono
    End If
    
    Debug.Print "[HojaPersonal] Sincronización completada: " & filasActualizadas & _
                 " registro(s) de '" & iniciales & "' actualizados a 'Personal activo' = " & estado
    
    ' Restablecer la protección basada en roles de usuario de la hoja destino
    Call SheetProtector2.ApplyRoleBasedProtection(wsCrono, Configuration2.APP_PASSWORD)
End Sub
