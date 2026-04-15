' ----------------------------------------------------------------------
' Módulo: TableDiagnostics
' Propósito: Diagnóstico de estructura de tablas para detectar nombres
'            reales de columnas y verificar consistencia del esquema.
' Fecha: 14/04/2026
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Subrutina: DiagnosticarTablaCronograma
' Propósito: Lista todas las columnas de tblCronogramaInspecciones
'            para verificar nombres reales.
' ----------------------------------------------------------------------
Public Sub DiagnosticarTablaCronograma()
    On Error GoTo ErrorHandler
    
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA).ListObjects(Configuration2.TABLE_CRONOGRAMA)
    
    Debug.Print "========================================="
    Debug.Print "DIAGNÓSTICO: " & tbl.Name
    Debug.Print "Hoja: " & Configuration2.SHEET_CRONOGRAMA
    Debug.Print "Total columnas: " & tbl.ListColumns.Count
    Debug.Print "========================================="
    
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        Debug.Print "  [" & col.Index & "] " & col.Name
    Next col
    
    Debug.Print "========================================="
    
    MsgBox "Diagnóstico completado. Revisa la Ventana Inmediato (Ctrl+G)", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'' ----------------------------------------------------------------------
' Subrutina: DiagnosticarTodasLasTablas
' Propósito: Lista columnas de TODAS las tablas críticas del sistema.
' ----------------------------------------------------------------------
Public Sub DiagnosticarTodasLasTablas()
    On Error Resume Next
    
    Debug.Print vbCrLf & "╔═══════════════════════════════════════════════════════════╗"
    Debug.Print "║  DIAGNÓSTICO COMPLETO DEL SISTEMA DE INSPECCIONES     ║"
    Debug.Print "╚═══════════════════════════════════════════════════════════╝" & vbCrLf
    
    ' tblCronogramaInspecciones
    Call DiagnosticarTabla(Configuration2.SHEET_CRONOGRAMA, Configuration2.TABLE_CRONOGRAMA)
    
    ' tblResumenCronograma
    Call DiagnosticarTabla(Configuration2.MAIN_MENU_SHEET, Configuration2.TABLE_RESUMEN_CRONOGRAMA)
    
    ' tblPersonal
    Call DiagnosticarTabla(Configuration2.SHEET_PERSONAL, Configuration2.TABLE_PERSONAL)
    
    ' tblPlantillas
    Call DiagnosticarTabla(Configuration2.SHEET_CHECKLIST, Configuration2.TABLE_PLANTILLAS)
    
    ' tblPreguntas
    Call DiagnosticarTabla(Configuration2.SHEET_CHECKLIST, Configuration2.TABLE_PREGUNTAS)
    
    ' tblSecciones
    Call DiagnosticarTabla(Configuration2.SHEET_CHECKLIST, Configuration2.TABLE_SECCIONES)
    
    ' tblOpcionesDeRespuesta
    Call DiagnosticarTabla(Configuration2.SHEET_CHECKLIST, Configuration2.TABLE_OPCIONES)
    
    ' tblEquipos
    Call DiagnosticarTabla(Configuration2.SHEET_CONFIGURACION, Configuration2.TABLE_EQUIPOS)
    
    ' tblAseguramientoCalidad
    Call DiagnosticarTabla(Configuration2.SHEET_ASEGURAMIENTO, Configuration2.TABLE_ASEGURAMIENTO)
    
    Debug.Print vbCrLf & "═══════════════════════════════════════════════════════════"
    Debug.Print "DIAGNÓSTICO COMPLETADO"
    Debug.Print "═══════════════════════════════════════════════════════════" & vbCrLf
    
    MsgBox "Diagnóstico completado." & vbCrLf & vbCrLf & _
           "Presiona Ctrl+G para ver la Ventana Inmediato" & vbCrLf & _
           "y copiar todos los resultados.", vbInformation, "Diagnóstico del Sistema"
End Sub

Private Sub DiagnosticarTabla(ByVal hoja As String, ByVal nombreTabla As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets(hoja)
    Set tbl = ws.ListObjects(nombreTabla)
    
    Debug.Print vbCrLf & "───────────────────────────────────────────────────────────"
    Debug.Print "TABLA: " & nombreTabla
    Debug.Print "Hoja: " & hoja
    Debug.Print "Columnas: " & tbl.ListColumns.Count
    
    If tbl.DataBodyRange Is Nothing Then
        Debug.Print "Filas de datos: 0 (TABLA VACÍA)"
    Else
        Debug.Print "Filas de datos: " & tbl.ListRows.Count
    End If
    
    Debug.Print "───────────────────────────────────────────────────────────"
    
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        Debug.Print "  [" & Format(col.Index, "00") & "] '" & col.Name & "'"
    Next col
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "  ⚠ ERROR al diagnosticar " & nombreTabla & ": " & Err.Description
End Sub
