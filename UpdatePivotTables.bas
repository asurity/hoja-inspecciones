' ----------------------------------------------------------------------
' Módulo: UpdatePivotTables
' Descripción: Refresca las tablas dinámicas y gráficos de la hoja "Graficos".
'              Contiene dos modos:
'                - RefrescarGraficosAuto: sin restricción, para uso automático
'                - ActualizarTablasDinamicas: con restricción Admin, para botón manual
' Fecha creación: 14/04/2026
' Última actualización: 17/06/2026 - Agregado RefrescarGraficosAuto
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Subrutina: RefrescarGraficosAuto
' Propósito: Refresca todas las tablas dinámicas y gráficos de la hoja
'            "Graficos" SIN restricción de rol.
'            Diseñada para ser llamada automáticamente desde:
'              - Worksheet_Activate de la hoja Graficos
' Uso: Call UpdatePivotTables.RefrescarGraficosAuto
' ----------------------------------------------------------------------
Public Sub RefrescarGraficosAuto()
    On Error GoTo ErrorHandler
    
    Dim wsGraficos As Worksheet
    Dim pt As PivotTable
    Dim co As ChartObject
    Dim refreshedCount As Long
    
    refreshedCount = 0
    
    ' Verificar que la hoja existe
    On Error Resume Next
    Set wsGraficos = ThisWorkbook.Sheets(Configuration2.SHEET_GRAFICOS)
    On Error GoTo ErrorHandler
    
    If wsGraficos Is Nothing Then
        Debug.Print "[RefrescarGraficosAuto] Hoja '" & Configuration2.SHEET_GRAFICOS & "' no encontrada"
        Exit Sub
    End If
    
    ' Refrescar todas las tablas dinámicas en la hoja Graficos
    For Each pt In wsGraficos.PivotTables
        On Error Resume Next
        pt.RefreshTable
        If Err.Number = 0 Then
            refreshedCount = refreshedCount + 1
        Else
            Debug.Print "[RefrescarGraficosAuto] Error al refrescar tabla dinámica '" & pt.Name & "': " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next pt
    
    ' Refrescar todos los gráficos (charts) en la hoja Graficos
    ' Esto es necesario porque los gráficos vinculados a tablas dinámicas
    ' no siempre se actualizan solo con RefreshTable
    For Each co In wsGraficos.ChartObjects
        On Error Resume Next
        co.Chart.Refresh
        If Err.Number <> 0 Then
            Debug.Print "[RefrescarGraficosAuto] Error al refrescar gráfico '" & co.Name & "': " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next co
    
    Debug.Print "[RefrescarGraficosAuto] Completado - " & refreshedCount & " tabla(s) dinámica(s) refrescada(s)"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "[RefrescarGraficosAuto] ERROR: " & Err.Number & " - " & Err.Description
    Call ErrorLogger2.Log("UpdatePivotTables.RefrescarGraficosAuto", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Subrutina: ActualizarTablasDinamicas
' Propósito: Refresca TODAS las conexiones y tablas del libro (RefreshAll).
'            Requiere rol Admin. Diseñada para el botón manual en la cinta.
' ----------------------------------------------------------------------
Sub ActualizarTablasDinamicas()

    If m_userRole = "Admin" Then
    
        ActiveWorkbook.RefreshAll
        MsgBox "Gráficos actualizados correctamente", vbApplicationModal, "PROCESO EXITOSO"
        
    Else
    
        MsgBox "No tienes permiso para actualizar los análisis de detecciones", vbCritical, "ACCESO DENEGADO"
        
    End If

End Sub