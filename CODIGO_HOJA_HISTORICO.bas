'' ----------------------------------------------------------------------
' Módulo: Hoja "Historico"
' Descripción: Eventos para la hoja Historico.
'              Permite generar certificados PDF con doble clic en inspecciones.
'
' INSTRUCCIONES DE INSTALACIÓN:
' 1. Abre el VBA Editor (Alt+F11)
' 2. Busca en el árbol de la izquierda: Microsoft Excel Objetos → Hoja "Historico"
' 3. Haz doble clic en esa hoja para abrir su módulo
' 4. Copia TODO el código de este archivo
' 5. Pégalo en el módulo de la hoja "Historico"
' 6. Guarda el archivo
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Evento: Worksheet_BeforeDoubleClick
' Propósito: Detecta doble clic en una fila de tblInspecciones y genera
'            el certificado PDF de esa inspección.
' Parámetros:
'   Target: Celda donde se hizo doble clic
'   Cancel: Si se pone en True, cancela el comportamiento normal de Excel
' ----------------------------------------------------------------------
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    Dim tblInspecciones As ListObject
    Dim fila As ListRow
    Dim idInspeccion As String
    
    ' Verificar que existe la tabla tblInspecciones
    On Error Resume Next
    Set tblInspecciones = Me.ListObjects(Configuration2.TABLE_INSPECCIONES)
    On Error GoTo ErrorHandler
    
    If tblInspecciones Is Nothing Then Exit Sub
    If tblInspecciones.DataBodyRange Is Nothing Then Exit Sub
    
    ' Verificar que el doble clic fue dentro del DataBodyRange
    If Intersect(Target, tblInspecciones.DataBodyRange) Is Nothing Then Exit Sub
    
    ' Obtener la fila donde se hizo clic
    Dim filaNum As Long
    filaNum = Target.Row - tblInspecciones.HeaderRowRange.Row
    
    If filaNum < 1 Or filaNum > tblInspecciones.ListRows.Count Then Exit Sub
    
    Set fila = tblInspecciones.ListRows(filaNum)
    
    ' Obtener ID Inspección de esa fila
    idInspeccion = Trim(CStr(fila.Range.Cells(1, tblInspecciones.ListColumns("ID Inspeccion").Index).Value))
    
    If Len(idInspeccion) = 0 Then
        MsgBox "No se pudo obtener el ID de inspección.", vbExclamation
        Exit Sub
    End If
    
    ' Cancelar el comportamiento por defecto del doble clic
    Cancel = True
    
    ' Confirmar con el usuario
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Desea generar el certificado PDF de esta inspección?", _
                       vbQuestion + vbYesNo, "Generar Certificado")
    
    If respuesta = vbYes Then
        Debug.Print "[Historico] Generando PDF para ID: " & idInspeccion
        
        ' Generar el certificado
        On Error Resume Next
        Call CertificadoPDFGenerator.GenerarCertificadoPDF(idInspeccion)
        
        If Err.Number <> 0 Then
            MsgBox "Error al generar el certificado PDF:" & vbCrLf & Err.Description, vbCritical
            Debug.Print "[Historico] ERROR: " & Err.Description
            Err.Clear
        Else
            Debug.Print "[Historico] PDF generado exitosamente"
        End If
        On Error GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "[Historico.BeforeDoubleClick] ERROR: " & Err.Description
    ' No mostramos mensaje al usuario para no interrumpir si fue un clic accidental
End Sub
