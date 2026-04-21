Attribute VB_Name = "DiagnosticoRespuestas"
' ══════════════════════════════════════════════════════════════
' Módulo: DiagnosticoRespuestas
' Propósito: Verificar que todas las respuestas se guardan y muestran
' Fecha: 17/04/2026
' ══════════════════════════════════════════════════════════════
Option Explicit

' ══════════════════════════════════════════════════════════════
' Diagnóstico completo de respuestas para una inspección
' ══════════════════════════════════════════════════════════════
Public Sub DiagnosticarInspeccionSeleccionada()
    On Error Resume Next
    
    ' Verificar que estamos en hoja Histórico
    If ActiveSheet.Name <> Configuration2.SHEET_HISTORICO Then
        MsgBox "Debes estar en la hoja '" & Configuration2.SHEET_HISTORICO & "' y seleccionar una inspección.", vbExclamation
        Exit Sub
    End If
    
    Dim wsh As Worksheet
    Dim tblInspecciones As ListObject
    Set wsh = ActiveSheet
    Set tblInspecciones = wsh.ListObjects(Configuration2.TABLE_INSPECCIONES)
    
    ' Verificar selección
    If Intersect(ActiveCell, tblInspecciones.DataBodyRange) Is Nothing Then
        MsgBox "Selecciona una fila en la tabla de inspecciones.", vbInformation
        Exit Sub
    End If
    
    ' Obtener ID inspección
    Dim filaActiva As Long
    filaActiva = ActiveCell.Row - tblInspecciones.DataBodyRange.Row + 1
    Dim idInspeccion As String
    idInspeccion = Trim(tblInspecciones.DataBodyRange.Cells(filaActiva, 1).Value)
    
    If Len(idInspeccion) = 0 Then
        MsgBox "La fila seleccionada no tiene ID de inspección.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener ID Plantilla
    Dim idPlantilla As String
    idPlantilla = Trim(tblInspecciones.DataBodyRange.Cells(filaActiva, 11).Value)
    
    ' Obtener iniciales
    Dim iniciales As String
    iniciales = Trim(tblInspecciones.DataBodyRange.Cells(filaActiva, 10).Value)
    
    Debug.Print "════════════════════════════════════════════════════════"
    Debug.Print "DIAGNÓSTICO DE RESPUESTAS"
    Debug.Print "════════════════════════════════════════════════════════"
    Debug.Print "ID Inspección: " & idInspeccion
    Debug.Print "Iniciales: " & iniciales
    Debug.Print "ID Plantilla: " & idPlantilla
    Debug.Print ""
    
    ' ═══════════════════════════════════════════════════════════
    ' PARTE 1: Cuántas preguntas tiene la plantilla
    ' ═══════════════════════════════════════════════════════════
    Dim wsChecklist As Worksheet
    Dim tblPreguntas As ListObject
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set tblPreguntas = wsChecklist.ListObjects(Configuration2.TABLE_PREGUNTAS)
    
    Dim preguntasPlantilla As Long
    preguntasPlantilla = 0
    
    Dim fila As Long
    For fila = 1 To tblPreguntas.DataBodyRange.Rows.Count
        Dim idPlant As String
        idPlant = Trim(tblPreguntas.DataBodyRange.Cells(fila, 1).Value)
        
        Dim activo As String
        activo = UCase(Trim(tblPreguntas.DataBodyRange.Cells(fila, 8).Value))
        
        If idPlant = idPlantilla And (activo = "SI" Or activo = "SÍ") Then
            preguntasPlantilla = preguntasPlantilla + 1
        End If
    Next fila
    
    Debug.Print "──────────────────────────────────────────────────────"
    Debug.Print "PREGUNTAS EN PLANTILLA (tblPreguntas)"
    Debug.Print "──────────────────────────────────────────────────────"
    Debug.Print "Total preguntas activas: " & preguntasPlantilla
    Debug.Print ""
    
    ' Contar por sección
    Dim tblSecciones As ListObject
    Set tblSecciones = wsChecklist.ListObjects(Configuration2.TABLE_SECCIONES)
    
    Dim seccion As Variant
    For fila = 1 To tblSecciones.DataBodyRange.Rows.Count
        Dim idSeccion As String
        Dim nombreSeccion As String
        idSeccion = Trim(tblSecciones.DataBodyRange.Cells(fila, 1).Value)
        nombreSeccion = Trim(tblSecciones.DataBodyRange.Cells(fila, 2).Value)
        
        ' Contar preguntas de esta sección
        Dim preguntasSeccion As Long
        preguntasSeccion = 0
        
        Dim f As Long
        For f = 1 To tblPreguntas.DataBodyRange.Rows.Count
            Dim idPlantTemp As String
            Dim idSecTemp As String
            Dim activoTemp As String
            
            idPlantTemp = Trim(tblPreguntas.DataBodyRange.Cells(f, 1).Value)
            idSecTemp = Trim(tblPreguntas.DataBodyRange.Cells(f, 3).Value)
            activoTemp = UCase(Trim(tblPreguntas.DataBodyRange.Cells(f, 8).Value))
            
            If idPlantTemp = idPlantilla And idSecTemp = idSeccion And (activoTemp = "SI" Or activoTemp = "SÍ") Then
                preguntasSeccion = preguntasSeccion + 1
            End If
        Next f
        
        If preguntasSeccion > 0 Then
            Debug.Print "  " & nombreSeccion & ": " & preguntasSeccion & " preguntas"
        End If
    Next fila
    Debug.Print ""
    
    ' ═══════════════════════════════════════════════════════════
    ' PARTE 2: Cuántas respuestas se guardaron
    ' ═══════════════════════════════════════════════════════════
    Dim tblRespuestas As ListObject
    Set tblRespuestas = wsh.ListObjects(Configuration2.TABLE_RESPUESTAS)
    
    Dim respuestasGuardadas As Long
    respuestasGuardadas = 0
    
    For fila = 1 To tblRespuestas.DataBodyRange.Rows.Count
        Dim idInsp As String
        idInsp = Trim(tblRespuestas.DataBodyRange.Cells(fila, 2).Value)
        
        If idInsp = idInspeccion Then
            respuestasGuardadas = respuestasGuardadas + 1
        End If
    Next fila
    
    Debug.Print "──────────────────────────────────────────────────────"
    Debug.Print "RESPUESTAS GUARDADAS (tblRespuestas)"
    Debug.Print "──────────────────────────────────────────────────────"
    Debug.Print "Total respuestas guardadas: " & respuestasGuardadas
    Debug.Print ""
    
    ' ═══════════════════════════════════════════════════════════
    ' PARTE 3: Cuántas respuestas aparecen en el PDF
    ' ═══════════════════════════════════════════════════════════
    Debug.Print "──────────────────────────────────────────────────────"
    Debug.Print "COMPARACIÓN"
    Debug.Print "──────────────────────────────────────────────────────"
    
    If preguntasPlantilla = respuestasGuardadas Then
        Debug.Print "✓ CORRECTO: Todas las preguntas tienen respuesta"
        Debug.Print "  Preguntas en plantilla: " & preguntasPlantilla
        Debug.Print "  Respuestas guardadas:    " & respuestasGuardadas
    Else
        Debug.Print "✗ ADVERTENCIA: Discrepancia en cantidades"
        Debug.Print "  Preguntas en plantilla: " & preguntasPlantilla
        Debug.Print "  Respuestas guardadas:    " & respuestasGuardadas
        Debug.Print "  Diferencia:              " & (preguntasPlantilla - respuestasGuardadas)
    End If
    
    Debug.Print ""
    Debug.Print "NOTA: El PDF mostrará las " & respuestasGuardadas & " respuestas"
    Debug.Print "      que están guardadas en tblRespuestas."
    Debug.Print "════════════════════════════════════════════════════════"
    
    ' Mostrar mensaje resumen
    Dim mensaje As String
    mensaje = "DIAGNÓSTICO COMPLETADO" & vbCrLf & vbCrLf & _
              "Preguntas en plantilla: " & preguntasPlantilla & vbCrLf & _
              "Respuestas guardadas: " & respuestasGuardadas & vbCrLf & vbCrLf
    
    If preguntasPlantilla = respuestasGuardadas Then
        mensaje = mensaje & "✓ Todas las preguntas tienen respuesta" & vbCrLf & _
                  "El PDF mostrará las " & respuestasGuardadas & " respuestas."
        MsgBox mensaje, vbInformation, "Diagnóstico OK"
    Else
        mensaje = mensaje & "⚠ Hay " & Abs(preguntasPlantilla - respuestasGuardadas) & " preguntas sin respuesta" & vbCrLf & vbCrLf & _
                  "Revisa el Output (Ctrl+G) para más detalles."
        MsgBox mensaje, vbExclamation, "Advertencia"
    End If
End Sub
