' ══════════════════════════════════════════════════════════════
' Módulo: CertificadoPDFGenerator
' Descripción: Genera PDFs de certificados de inspección a partir
'              de datos completados en tblInspecciones + tblRespuestas
' Fecha creación: 17/04/2026
' Dependencias: Configuration2, PlantillaCertificadoSetup,
'               InspectionRepository, ChecklistRepository
' ══════════════════════════════════════════════════════════════
Option Explicit

' ══════════════════════════════════════════════════════════════
' FUNCIÓN PRINCIPAL: GenerarCertificadoPDF
' Propósito: Genera un PDF de certificado a partir del ID inspección
' ══════════════════════════════════════════════════════════════
Public Sub GenerarCertificadoPDF(ByVal idInspeccion As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "===== INICIO GenerarCertificadoPDF ====="
    Debug.Print "ID Inspección: " & idInspeccion
    
    ' Validar que existe ID inspección
    If Len(Trim(idInspeccion)) = 0 Then
        MsgBox "Debe seleccionar una inspección para generar el certificado.", vbExclamation
        Exit Sub
    End If
    
    Dim wsh As Worksheet
    Dim wsPlantilla As Worksheet
    Dim rutaPDF As String
    Dim nombreArchivo As String
    
    ' Obtener hojas
    Debug.Print "[DEBUG] Obteniendo hoja Histórico..."
    Set wsh = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Debug.Print "[DEBUG] Hoja Histórico: " & wsh.Name
    
    ' Desproteger plantilla si existe
    Debug.Print "[DEBUG] Buscando plantilla certificado..."
    On Error Resume Next
    Set wsPlantilla = ThisWorkbook.Sheets(Configuration2.SHEET_PLANTILLA_CERTIFICADO)
    On Error GoTo ErrorHandler
    
    If wsPlantilla Is Nothing Then
        Debug.Print "[ERROR] Plantilla no encontrada"
        MsgBox "ERROR: No existe la plantilla de certificado." & vbCrLf & _
               "Ejecute primero: PlantillaCertificadoSetup.InicializarPlantillaCertificado()", _
               vbCritical
        Exit Sub
    End If
    
    Debug.Print "[DEBUG] Plantilla encontrada: " & wsPlantilla.Name
    Debug.Print "[DEBUG] Estado plantilla: " & wsPlantilla.Visible
    
    Dim datosInspeccion As Object
    Debug.Print "[DEBUG] Obteniendo datos de inspección..."
    Set datosInspeccion = ObtenerDatosInspeccion(idInspeccion)
    
    If datosInspeccion Is Nothing Then
        Debug.Print "[ERROR] Datos de inspección no encontrados"
        MsgBox "No se encontró la inspección con ID: " & idInspeccion, vbCritical
        Exit Sub
    End If
    
    Debug.Print "[DEBUG] Datos de inspección obtenidos correctamente"
    
    ' Hacer visible la plantilla temporalmente
    Debug.Print "[DEBUG] Haciendo plantilla visible..."
    wsPlantilla.Visible = xlSheetVisible
    Debug.Print "[DEBUG] Plantilla visible: " & (wsPlantilla.Visible = xlSheetVisible)
    
    ' Poblar plantilla con datos
    Debug.Print "[DEBUG] Poblando plantilla con datos..."
    Call PoblarPlantillaCertificado(wsPlantilla, datosInspeccion)
    Debug.Print "[DEBUG] Plantilla poblada"
    
    ' Generar nombre de archivo
    nombreArchivo = GenerarNombreArchivoPDF(datosInspeccion)
    Debug.Print "[DEBUG] Nombre archivo: " & nombreArchivo
    
    ' Obtener ruta correcta (Desktop o Escritorio según idioma Windows)
    Debug.Print "[DEBUG] Detectando carpeta Desktop..."
    Dim desktopPath As String
    desktopPath = ObtenerRutaDesktop()
    Debug.Print "[DEBUG] Ruta Desktop detectada: " & desktopPath
    
    rutaPDF = desktopPath & "\" & nombreArchivo
    Debug.Print "[DEBUG] Ruta PDF final: " & rutaPDF
    
    ' Verificar si la carpeta existe
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Debug.Print "[DEBUG] Verificando que la carpeta destino existe..."
    If Not fso.FolderExists(desktopPath) Then
        Debug.Print "[ERROR] Carpeta no existe: " & desktopPath
        MsgBox "Error: No se encuentra la carpeta: " & desktopPath, vbCritical
        Exit Sub
    End If
    Debug.Print "[DEBUG] Carpeta destino verificada: " & desktopPath
    
    ' Debug: Verificar contenido de la plantilla
    Debug.Print "[DEBUG] Verificando contenido de plantilla..."
    Debug.Print "[DEBUG] Celdas usadas en plantilla: " & wsPlantilla.UsedRange.Address
    Debug.Print "[DEBUG] Filas usadas: " & wsPlantilla.UsedRange.Rows.Count
    Debug.Print "[DEBUG] Columnas usadas: " & wsPlantilla.UsedRange.Columns.Count
    
    ' Procesar solo el rango usado
    Dim rangoExportar As Range
    Set rangoExportar = wsPlantilla.UsedRange
    Debug.Print "[DEBUG] Rango a exportar: " & rangoExportar.Address
    
    ' Exportar como PDF
    Debug.Print "[DEBUG] Iniciando exportación a PDF..."
    Debug.Print "[DEBUG] Método: ExportAsFixedFormat(xlTypePDF, ...)"
    
    On Error Resume Next
    wsPlantilla.ExportAsFixedFormat xlTypePDF, rutaPDF, , True
    Dim errExport As Long
    errExport = Err.Number
    Dim descExport As String
    descExport = Err.Description
    On Error GoTo ErrorHandler
    
    If errExport <> 0 Then
        Debug.Print "[ERROR] Exportación fallida"
        Debug.Print "[ERROR] Error Number: " & errExport
        Debug.Print "[ERROR] Error Description: " & descExport
        Err.Raise errExport, "ExportAsFixedFormat", descExport
    End If
    
    Debug.Print "[DEBUG] PDF generado exitosamente"
    
    ' Verificar que el archivo fue creado
    Debug.Print "[DEBUG] Verificando archivo creado..."
    If fso.FileExists(rutaPDF) Then
        Debug.Print "[DEBUG] Archivo existe: " & rutaPDF
        Debug.Print "[DEBUG] Tamaño archivo: " & fso.GetFile(rutaPDF).Size & " bytes"
    Else
        Debug.Print "[ERROR] Archivo NO fue creado"
        MsgBox "Error: El PDF no fue creado en: " & rutaPDF, vbCritical
        Exit Sub
    End If
    
    ' Ocultar plantilla nuevamente
    Debug.Print "[DEBUG] Ocultando plantilla..."
    wsPlantilla.Visible = xlSheetVeryHidden
    Debug.Print "[DEBUG] Plantilla ocultada"
    
    ' Limpiar plantilla
    Debug.Print "[DEBUG] Limpiando plantilla..."
    Call LimpiarPlantillaCertificado(wsPlantilla)
    Debug.Print "[DEBUG] Plantilla limpia"
    
    ' Notificar éxito
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox( _
        "Certificado generado exitosamente." & vbCrLf & vbCrLf & _
        "Archivo: " & nombreArchivo & vbCrLf & _
        "Ubicación: Escritorio" & vbCrLf & vbCrLf & _
        "¿Desea abrir el PDF ahora?", _
        vbQuestion + vbYesNo, "Certificado PDF Generado")
    
    If respuesta = vbYes Then
        Shell "explorer.exe """ & rutaPDF & """"
    End If
    
    Debug.Print "===== FIN GenerarCertificadoPDF - ÉXITO ====="
    Exit Sub
    
ErrorHandler:
    Debug.Print "===== ERROR EN GenerarCertificadoPDF ====="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    
    Call ErrorLogger2.Log("CertificadoPDFGenerator.GenerarCertificadoPDF", Err.Description, Err.Number)
    
    ' Asegurar que plantilla queda oculta
    On Error Resume Next
    wsPlantilla.Visible = xlSheetVeryHidden
    On Error GoTo 0
    
    MsgBox "Error al generar certificado:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error"
    Debug.Print "===== FIN GenerarCertificadoPDF - ERROR ====="
End Sub

' ══════════════════════════════════════════════════════════════
' Obtener todos los datos de una inspección
' ══════════════════════════════════════════════════════════════
Private Function ObtenerDatosInspeccion(ByVal idInspeccion As String) As Object
    On Error GoTo ErrorHandler
    
    Dim tbl As ListObject
    Dim fila As Range
    Dim encontrada As Boolean
    Dim datos As Object
    Set datos = CreateObject("Scripting.Dictionary")
    
    Dim wsHistorico As Worksheet
    Set wsHistorico = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Set tbl = wsHistorico.ListObjects(Configuration2.TABLE_INSPECCIONES)
    
    ' Buscar registro
    encontrada = False
    Dim row As Long
    For row = 1 To tbl.DataBodyRange.Rows.Count
        If Trim(tbl.DataBodyRange.Cells(row, 1).Value) = idInspeccion Then
            encontrada = True
            Exit For
        End If
    Next row
    
    If Not encontrada Then
        Set ObtenerDatosInspeccion = Nothing
        Exit Function
    End If
    
    ' Extraer datos fila
    Set fila = tbl.DataBodyRange.Rows(row)
    
    datos("ID") = idInspeccion
    datos("Area") = fila.Cells(1, 2).Value
    datos("LineaAuditada") = fila.Cells(1, 3).Value
    datos("HoraInicio") = fila.Cells(1, 4).Value
    datos("HoraTermino") = fila.Cells(1, 5).Value
    datos("AY1") = fila.Cells(1, 6).Value
    datos("AY2") = fila.Cells(1, 7).Value
    datos("OP") = fila.Cells(1, 8).Value
    datos("LugarAuditoria") = fila.Cells(1, 9).Value
    datos("Iniciales") = fila.Cells(1, 10).Value
    datos("IDPlantilla") = fila.Cells(1, 11).Value
    datos("Planta") = fila.Cells(1, 12).Value
    datos("FechaInspeccion") = fila.Cells(1, 13).Value
    datos("Auditor") = fila.Cells(1, 14).Value
    datos("TAPuntaje") = fila.Cells(1, 16).Value
    datos("TAMaximos") = fila.Cells(1, 17).Value
    datos("TANoAplica") = fila.Cells(1, 18).Value
    datos("TAPorcentaje") = fila.Cells(1, 19).Value
    
    ' FIX 21/04/2026: Usar ListColumns en lugar de índices hardcodeados
    ' Razón: tblInspecciones tiene 31 columnas (col 20 = Auditoria Procesos, 21 = RPN, 22 = Categoria)
    datos("RPN") = fila.Cells(1, tbl.ListColumns("RPN calculado").Index).Value
    datos("Categoria") = fila.Cells(1, tbl.ListColumns("Categoria resultado").Index).Value
    
    ' Intentar leer resultado de Auditoría de Procesos (buscar en varias posibles columnas)
    On Error Resume Next
    Dim resultadoProcesos As String
    resultadoProcesos = ""
    
    Dim nombresPosibles As Variant
    nombresPosibles = Array("Auditoria Procesos Resultado", "Auditoría Procesos Resultado", _
                           "Resultado Auditoria Procesos", "AP Resultado")
    
    Dim nombreCol As Variant
    Dim colIdx As Variant
    For Each nombreCol In nombresPosibles
        colIdx = Application.Match(CStr(nombreCol), tbl.HeaderRowRange, 0)
        If Not IsError(colIdx) Then
            resultadoProcesos = Trim(CStr(fila.Cells(1, CLng(colIdx)).Value))
            If Len(resultadoProcesos) > 0 Then
                Debug.Print "[CertificadoPDF] Auditoría Procesos encontrada en columna '" & nombreCol & "': " & resultadoProcesos
                Exit For
            End If
        End If
    Next nombreCol
    On Error GoTo ErrorHandler
    
    datos("AuditoriaProcesosResultado") = resultadoProcesos
    
    ' FIX 21/04/2026: Usar ListColumns en lugar de índice hardcodeado
    datos("ObservacionesGenerales") = fila.Cells(1, tbl.ListColumns("Observaciones generales").Index).Value
    
    ' Obtener respuestas
    Set datos("Respuestas") = ObtenerRespuestasInspeccion(idInspeccion)
    
    Set ObtenerDatosInspeccion = datos
    Exit Function
    
ErrorHandler:
    Set ObtenerDatosInspeccion = Nothing
End Function

' ══════════════════════════════════════════════════════════════
' Obtener respuestas de inspección
' ══════════════════════════════════════════════════════════════
Private Function ObtenerRespuestasInspeccion(ByVal idInspeccion As String) As Collection
    Dim respuestas As New Collection
    Dim tbl As ListObject
    Dim wsh As Worksheet
    Dim fila As Long
    
    Set wsh = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Set tbl = wsh.ListObjects(Configuration2.TABLE_RESPUESTAS)
    
    If tbl.DataBodyRange Is Nothing Then
        Set ObtenerRespuestasInspeccion = respuestas
        Exit Function
    End If
    
    ' Buscar todas las respuestas de esta inspección
    For fila = 1 To tbl.DataBodyRange.Rows.Count
        Dim idInsp As String
        idInsp = Trim(tbl.DataBodyRange.Cells(fila, 2).Value)  ' Columna 2 es ID Inspeccion
        
        If idInsp = idInspeccion Then
            Dim resp As Object
            Set resp = CreateObject("Scripting.Dictionary")
            
            resp("IDPregunta") = tbl.DataBodyRange.Cells(fila, 3).Value
            resp("IDOpcion") = tbl.DataBodyRange.Cells(fila, 4).Value
            resp("ValorNumerico") = tbl.DataBodyRange.Cells(fila, 5).Value
            resp("Observacion") = tbl.DataBodyRange.Cells(fila, 6).Value
            
            respuestas.Add resp
        End If
    Next fila
    
    Set ObtenerRespuestasInspeccion = respuestas
End Function

' ══════════════════════════════════════════════════════════════
' Poblar plantilla con datos de inspección
' ══════════════════════════════════════════════════════════════
Private Sub PoblarPlantillaCertificado(wsPlantilla As Worksheet, datosInspeccion As Object)
    On Error Resume Next
    
    Debug.Print "[POBLACIÓN] Iniciando población de plantilla..."
    
    ' Obtener puesto desde tblPersonal (hoja correcta: SHEET_PERSONAL)
    Dim nombreEvaluado As String
    Dim puesto As String
    nombreEvaluado = ObtenerNombrePersonal(datosInspeccion("Iniciales"))
    puesto = ObtenerPuestoPersonal(datosInspeccion("Iniciales"))
    If Len(nombreEvaluado) = 0 Then nombreEvaluado = datosInspeccion("Iniciales")
    If Len(puesto) = 0 Then puesto = "No especificado"
    
    ' ═══════════════════════════════════════════════════════════
    ' NUEVO: SECCIÓN 0 - BLOQUE CATEGORÍA (Filas 6-8)
    ' ═══════════════════════════════════════════════════════════
    Debug.Print "[POBLACIÓN] Poblando bloque categoría..."
    
    Dim textoCategoria As String
    Dim nombreCategoria As String
    Dim colorFondo As Long
    Dim numCategoria As Long
    
    numCategoria = CLng(datosInspeccion("Categoria"))
    
    ' Obtener nombre categoría desde tblCategoriasRPN
    nombreCategoria = ObtenerNombreCategoria(numCategoria)
    
    ' Construir texto del bloque
    textoCategoria = "CATEGORÍA " & numCategoria & " - " & UCase(nombreCategoria) & vbCrLf & _
                     "RPN: " & Format(datosInspeccion("RPN"), "0.00")
    
    ' Asignar a celda A6:G8 (rango combinado)
    wsPlantilla.Range("A6:G8").Value = textoCategoria
    
    ' Aplicar color según categoría
    Select Case numCategoria
        Case 1, 2
            colorFondo = RGB(212, 244, 230)  ' Verde claro (#D4F4E6)
        Case 3
            colorFondo = RGB(254, 249, 219)  ' Amarillo claro (#FEF9DB)
        Case 4, 5
            colorFondo = RGB(253, 223, 223)  ' Rojo claro (#FDDFDF)
        Case Else
            colorFondo = RGB(255, 255, 255)  ' Blanco por defecto
    End Select
    
    wsPlantilla.Range("A6:G8").Interior.Color = colorFondo
    
    Debug.Print "[POBLACIÓN] Bloque categoría completado: Cat " & numCategoria & " - " & nombreCategoria
    
    ' ═══════════════════════════════════════════════════════════
    ' SECCIÓN 1: DATOS DE INSPECCIÓN (Filas 9-17, desplazadas +3)
    ' Usamos referencias de celda directas (layout definido en PlantillaCertificadoSetup)
    ' ═══════════════════════════════════════════════════════════
    Debug.Print "[POBLACIÓN] Poblando datos de inspección..."
    
    wsPlantilla.Cells(10, 3).Value = datosInspeccion("FechaInspeccion")   ' C10
    wsPlantilla.Cells(10, 5).Value = Format(datosInspeccion("HoraInicio"), "HH:MM")  ' E10
    wsPlantilla.Cells(10, 7).Value = Format(datosInspeccion("HoraTermino"), "HH:MM") ' G10
    wsPlantilla.Cells(11, 3).Value = nombreEvaluado                        ' C11
    wsPlantilla.Cells(11, 7).Value = datosInspeccion("Iniciales")          ' G11
    wsPlantilla.Cells(12, 3).Value = puesto                                ' C12
    wsPlantilla.Cells(12, 6).Value = datosInspeccion("Planta")             ' F12
    wsPlantilla.Cells(13, 3).Value = datosInspeccion("Area")              ' C13
    wsPlantilla.Cells(13, 6).Value = datosInspeccion("LineaAuditada")     ' F13
    wsPlantilla.Cells(14, 3).Value = datosInspeccion("LugarAuditoria")    ' C14
    wsPlantilla.Cells(15, 3).Value = datosInspeccion("Auditor")           ' C15
    wsPlantilla.Cells(16, 4).Value = datosInspeccion("AY1")               ' D16
    wsPlantilla.Cells(16, 6).Value = datosInspeccion("AY2")               ' F16
    wsPlantilla.Cells(16, 7).Value = "OP: " & datosInspeccion("OP")       ' G16
    
    Debug.Print "[POBLACIÓN] Datos básicos completados"
    
    ' ═══════════════════════════════════════════════════════════
    ' SECCIÓN 2: RESULTADOS GENERALES (Filas 19-25, desplazadas +3)
    ' ═══════════════════════════════════════════════════════════════
    Debug.Print "[POBLACIÓN] Poblando resultados generales..."
    
    wsPlantilla.Cells(20, 3).Value = datosInspeccion("TAPuntaje")          ' C20
    wsPlantilla.Cells(20, 5).Value = datosInspeccion("TAMaximos")          ' E20
    wsPlantilla.Cells(21, 3).Value = datosInspeccion("TANoAplica")         ' C21
    wsPlantilla.Cells(22, 3).Value = Format(datosInspeccion("TAPorcentaje"), "0.00") & "%" ' C22
    wsPlantilla.Cells(23, 3).Value = Format(datosInspeccion("RPN"), "0.00") ' C23
    wsPlantilla.Cells(24, 3).Value = datosInspeccion("Categoria")          ' C24
    
    ' Estado de Competencia basado en categoría (C25)
    Dim textoEstado As String
    Dim colorEstado As Long
    
    Select Case numCategoria
        Case 1, 2
            textoEstado = ChrW(10004) & " COMPETENTE"  ' ✓ COMPETENTE
            colorEstado = RGB(39, 174, 96)  ' Verde (#27AE60)
        Case 3
            textoEstado = ChrW(9888) & " COMPETENTE CON OBSERVACIONES"  ' ⚠ COMPETENTE CON OBSERVACIONES
            colorEstado = RGB(243, 156, 18)  ' Naranja (#F39C12)
        Case 4, 5
            textoEstado = ChrW(10006) & " NO CALIFICADO"  ' ✗ NO CALIFICADO
            colorEstado = RGB(203, 67, 53)  ' Rojo (#CB4335)
        Case Else
            textoEstado = "ESTADO DESCONOCIDO"
            colorEstado = RGB(0, 0, 0)  ' Negro por defecto
    End Select
    
    wsPlantilla.Cells(25, 3).Value = textoEstado  ' C25
    wsPlantilla.Cells(25, 3).Font.Color = colorEstado
    wsPlantilla.Cells(25, 3).Font.Bold = True
    
    Debug.Print "[POBLACIÓN] Estado asignado: " & textoEstado & " (Cat " & numCategoria & ")"
    Debug.Print "[POBLACIÓN] Resultados completados"
    
    ' ═══════════════════════════════════════════════════════════
    ' SECCIÓN 3: TABLA DE PREGUNTAS Y RESPUESTAS
    ' ═══════════════════════════════════════════════════════════
    Debug.Print "[POBLACIÓN] Poblando tabla de respuestas..."
    
    On Error GoTo 0
    Call PoblarTablaRespuestas(wsPlantilla, datosInspeccion)
    
    ' ═══════════════════════════════════════════════════════════
    ' SECCIÓN 4: PIE DE VALIDEZ (Paso 4 MVP - Fila 50)
    ' ═══════════════════════════════════════════════════════════
    Debug.Print "[POBLACIÓN] Calculando fecha de validez..."
    
    Dim fechaInspeccion As Date
    Dim frecuenciaMeses As Long
    Dim fechaValidez As Date
    Dim textoPie As String
    
    fechaInspeccion = CDate(datosInspeccion("FechaInspeccion"))
    frecuenciaMeses = ObtenerFrecuenciaPlantilla(CStr(datosInspeccion("IDPlantilla")))
    
    If frecuenciaMeses > 0 Then
        fechaValidez = DateAdd("m", frecuenciaMeses, fechaInspeccion)
        textoPie = "Este certificado es válido hasta: " & Format(fechaValidez, "DD/MM/YYYY")
        Debug.Print "[POBLACIÓN] Fecha validez calculada: " & Format(fechaValidez, "DD/MM/YYYY") & " (" & frecuenciaMeses & " meses)"
    Else
        textoPie = "Este certificado es válido hasta: [FRECUENCIA NO DEFINIDA]"
        Debug.Print "[POBLACIÓN] ⚠ Advertencia: Frecuencia no encontrada para plantilla " & datosInspeccion("IDPlantilla")
    End If
    
    wsPlantilla.Range("A50:G50").Value = textoPie
    
    Debug.Print "[POBLACIÓN] ¡Plantilla completamente poblada!"
End Sub

' ══════════════════════════════════════════════════════════════
' Poblar tabla de preguntas y respuestas
' ══════════════════════════════════════════════════════════════
Private Sub PoblarTablaRespuestas(wsPlantilla As Worksheet, datosInspeccion As Object)
    On Error GoTo ErrorHandler
    
    Dim respuestas As Collection
    Set respuestas = datosInspeccion("Respuestas")
    
    Debug.Print "[RESPUESTAS] Total de respuestas: " & respuestas.Count
    
    If respuestas.Count = 0 Then
        Debug.Print "[RESPUESTAS] Sin respuestas para mostrar"
        Exit Sub
    End If
    
    ' Las tablas de preguntas y opciones están en SHEET_CHECKLIST, no en HISTORICO
    Dim wsChecklist As Worksheet
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    
    Dim tblPreguntas As ListObject
    Dim tblOpciones As ListObject
    Set tblPreguntas = wsChecklist.ListObjects(Configuration2.TABLE_PREGUNTAS)
    Set tblOpciones = wsChecklist.ListObjects(Configuration2.TABLE_OPCIONES)
    
    Debug.Print "[RESPUESTAS] tblPreguntas encontrada: " & (Not tblPreguntas Is Nothing)
    Debug.Print "[RESPUESTAS] tblOpciones encontrada: " & (Not tblOpciones Is Nothing)
    
    ' Fila 27 = inicio tabla de preguntas (desplazado +3 por bloque categoría)
    Dim filaInicio As Long
    filaInicio = 27
    
    Dim i As Long
    Dim resp As Object
    Dim fila As Long
    
    For i = 1 To respuestas.Count
        Set resp = respuestas(i)
        fila = filaInicio + (i - 1)
        
        Debug.Print "[RESPUESTAS] Fila " & fila & " - IDPregunta: " & resp("IDPregunta") & " - IDOpcion: " & resp("IDOpcion")
        
        ' Col B: Número secuencial
        wsPlantilla.Cells(fila, 2).Value = i
        
        ' Col C: Texto de pregunta (buscar por ID string)
        wsPlantilla.Cells(fila, 3).Value = ObtenerTextoPregunta(CStr(resp("IDPregunta")), tblPreguntas)
        
        ' Col E: Texto de opción elegida (buscar por ID string)
        Dim textoOpcion As String
        textoOpcion = ObtenerTextoOpcion(CStr(resp("IDOpcion")), tblOpciones)
        wsPlantilla.Cells(fila, 5).Value = textoOpcion
        
        ' Col F: Observación
        wsPlantilla.Cells(fila, 6).Value = resp("Observacion")
        
        ' Formato de fila
        With wsPlantilla.Rows(fila)
            .Font.Name = "Arial"
            .Font.Size = 9
            .VerticalAlignment = xlTop
            .RowHeight = 30
        End With
        With wsPlantilla.Range(wsPlantilla.Cells(fila, 2), wsPlantilla.Cells(fila, 7))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        ' ═══════════════════════════════════════════════════════════
        ' PASO 3 MVP: Resaltar incumplimientos con fondo rojo
        ' ═══════════════════════════════════════════════════════════
        If EsRespuestaIncumplimiento(textoOpcion) Then
            wsPlantilla.Range(wsPlantilla.Cells(fila, 2), wsPlantilla.Cells(fila, 7)).Interior.Color = RGB(253, 223, 223)  ' Rojo claro
            Debug.Print "[RESPUESTAS] ⚠ INCUMPLIMIENTO detectado: " & textoOpcion
        End If
        
        Debug.Print "[RESPUESTAS] OK -> " & wsPlantilla.Cells(fila, 3).Value
    Next i
    
    Debug.Print "[RESPUESTAS] Todas las respuestas pobladas"
    Exit Sub
    
ErrorHandler:
    Debug.Print "[ERROR RESPUESTAS] Fila " & fila & ": " & Err.Description
End Sub

' ══════════════════════════════════════════════════════════════
' Obtener texto de pregunta por UUID (tablas en SHEET_CHECKLIST)
' Columnas tblPreguntas: [1]ID Plantilla [2]ID Pregunta [5]Texto
' ══════════════════════════════════════════════════════════════
Private Function ObtenerTextoPregunta(idPregunta As String, tbl As ListObject) As String
    On Error GoTo ErrorHandler
    
    If tbl Is Nothing Then
        ObtenerTextoPregunta = "[sin tabla]"
        Exit Function
    End If
    
    Dim colID As Long
    Dim colTexto As Long
    colID = tbl.ListColumns("ID Pregunta").Index
    colTexto = tbl.ListColumns("Texto").Index
    
    Dim fila As Long
    For fila = 1 To tbl.DataBodyRange.Rows.Count
        If Trim(CStr(tbl.DataBodyRange.Cells(fila, colID).Value)) = Trim(idPregunta) Then
            ObtenerTextoPregunta = tbl.DataBodyRange.Cells(fila, colTexto).Value
            Exit Function
        End If
    Next fila
    
    ObtenerTextoPregunta = "[Pregunta no encontrada]"
    Exit Function
    
ErrorHandler:
    ObtenerTextoPregunta = "[Error: " & Err.Description & "]"
End Function

' ══════════════════════════════════════════════════════════════
' Obtener texto de opción por UUID (tablas en SHEET_CHECKLIST)
' Columnas tblOpcionesDeRespuesta: [1]ID Opcion [4]Opción texto
' ══════════════════════════════════════════════════════════════
Private Function ObtenerTextoOpcion(idOpcion As String, tbl As ListObject) As String
    On Error GoTo ErrorHandler
    
    If tbl Is Nothing Then
        ObtenerTextoOpcion = "[sin tabla]"
        Exit Function
    End If
    
    Dim colID As Long
    Dim colTexto As Long
    colID = tbl.ListColumns("ID Opcion").Index
    colTexto = tbl.ListColumns("Opción texto").Index
    
    Dim fila As Long
    For fila = 1 To tbl.DataBodyRange.Rows.Count
        If Trim(CStr(tbl.DataBodyRange.Cells(fila, colID).Value)) = Trim(idOpcion) Then
            ObtenerTextoOpcion = tbl.DataBodyRange.Cells(fila, colTexto).Value
            Exit Function
        End If
    Next fila
    
    ObtenerTextoOpcion = "[Opción no encontrada]"
    Exit Function
    
ErrorHandler:
    ObtenerTextoOpcion = "[Error: " & Err.Description & "]"
End Function

' ══════════════════════════════════════════════════════════════
' Detectar si una respuesta indica incumplimiento (Paso 3 MVP)
' Busca palabras clave que indican no conformidad
' ══════════════════════════════════════════════════════════════
Private Function EsRespuestaIncumplimiento(textoOpcion As String) As Boolean
    On Error Resume Next
    
    EsRespuestaIncumplimiento = False
    
    If Len(Trim(textoOpcion)) = 0 Then Exit Function
    
    ' Normalizar texto (quitar acentos, a minúsculas, espacios extra)
    Dim textoNorm As String
    textoNorm = LCase(Trim(textoOpcion))
    textoNorm = Replace(textoNorm, "á", "a")
    textoNorm = Replace(textoNorm, "é", "e")
    textoNorm = Replace(textoNorm, "í", "i")
    textoNorm = Replace(textoNorm, "ó", "o")
    textoNorm = Replace(textoNorm, "ú", "u")
    
    ' Patrones de incumplimiento (orden de especificidad: más específico primero)
    If InStr(textoNorm, "no cumple") > 0 Then EsRespuestaIncumplimiento = True: Exit Function
    If InStr(textoNorm, "incumplimiento") > 0 Then EsRespuestaIncumplimiento = True: Exit Function
    If InStr(textoNorm, "no conforme") > 0 Then EsRespuestaIncumplimiento = True: Exit Function
    If InStr(textoNorm, "deficiente") > 0 Then EsRespuestaIncumplimiento = True: Exit Function
    If InStr(textoNorm, "inadecuado") > 0 Then EsRespuestaIncumplimiento = True: Exit Function
    If InStr(textoNorm, "rechazado") > 0 Then EsRespuestaIncumplimiento = True: Exit Function
    
    ' Patrón simple "No" solo si es palabra completa (evitar falsos positivos como "Nominal")
    ' Buscar "no" como palabra aislada o al inicio/fin
    If textoNorm = "no" Then EsRespuestaIncumplimiento = True: Exit Function
    If Left(textoNorm, 3) = "no " Then EsRespuestaIncumplimiento = True: Exit Function
    If Right(textoNorm, 3) = " no" Then EsRespuestaIncumplimiento = True: Exit Function
    If InStr(textoNorm, " no ") > 0 Then EsRespuestaIncumplimiento = True: Exit Function
    
    On Error GoTo 0
End Function

' ══════════════════════════════════════════════════════════════
' Obtener nombre por iniciales (tblPersonal en SHEET_PERSONAL)
' tblPersonal: [1]Iniciales [2]Planta [3-13]Puestos Si/No [14]Activo
' No hay columna Nombre → se devuelven las iniciales como identificador
' ══════════════════════════════════════════════════════════════
Private Function ObtenerNombrePersonal(iniciales As String) As String
    ' tblPersonal no tiene columna Nombre Completo — devolver iniciales
    ObtenerNombrePersonal = Trim(iniciales)
End Function

' ══════════════════════════════════════════════════════════════
' Obtener puestos activos por iniciales (tblPersonal en SHEET_PERSONAL)
' Columnas 3-13 son puestos con valor "Si"/"No"
' ══════════════════════════════════════════════════════════════
Private Function ObtenerPuestoPersonal(iniciales As String) As String
    On Error GoTo ErrorHandler
    
    If Len(Trim(iniciales)) = 0 Then
        ObtenerPuestoPersonal = ""
        Exit Function
    End If
    
    Dim wsPersonal As Worksheet
    Dim tbl As ListObject
    Dim fila As Long
    
    Set wsPersonal = ThisWorkbook.Sheets(Configuration2.SHEET_PERSONAL)
    
    On Error Resume Next
    Set tbl = wsPersonal.ListObjects(Configuration2.TABLE_PERSONAL)
    On Error GoTo ErrorHandler
    
    If tbl Is Nothing Then
        ObtenerPuestoPersonal = ""
        Exit Function
    End If
    
    Dim colIniciales As Long
    colIniciales = tbl.ListColumns("Iniciales").Index
    
    ' Nombres de puestos (columnas 3 a 13 según Configuration2)
    Dim puestosNombres As Variant
    puestosNombres = Configuration2.GetPuestosColumns()
    
    ' Buscar la fila de este personal
    For fila = 1 To tbl.DataBodyRange.Rows.Count
        If Trim(CStr(tbl.DataBodyRange.Cells(fila, colIniciales).Value)) = Trim(iniciales) Then
            ' Iterar columnas de puestos y recopilar los que tienen "Si"
            Dim puestosActivos As String
            Dim p As Long
            puestosActivos = ""
            For p = 0 To UBound(puestosNombres)
                Dim colPuesto As Long
                On Error Resume Next
                colPuesto = tbl.ListColumns(puestosNombres(p)).Index
                On Error GoTo ErrorHandler
                If colPuesto > 0 Then
                    Dim valPuesto As String
                    valPuesto = UCase(Trim(CStr(tbl.DataBodyRange.Cells(fila, colPuesto).Value)))
                    If valPuesto = "SI" Or valPuesto = "SÍ" Then
                        If Len(puestosActivos) > 0 Then puestosActivos = puestosActivos & ", "
                        puestosActivos = puestosActivos & puestosNombres(p)
                    End If
                End If
                colPuesto = 0
            Next p
            ObtenerPuestoPersonal = puestosActivos
            Exit Function
        End If
    Next fila
    
    ObtenerPuestoPersonal = ""
    Exit Function
    
ErrorHandler:
    ObtenerPuestoPersonal = ""
End Function

' ══════════════════════════════════════════════════════════════
' Obtener frecuencia de plantilla (en meses) desde tblPlantillas
' Paso 4 MVP - Para calcular fecha de validez del certificado
' ══════════════════════════════════════════════════════════════
Private Function ObtenerFrecuenciaPlantilla(ByVal idPlantilla As String) As Long
    On Error GoTo ErrorHandler
    
    Dim wsChecklist As Worksheet
    Dim tbl As ListObject
    Dim fila As Long
    
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    
    On Error Resume Next
    Set tbl = wsChecklist.ListObjects(Configuration2.TABLE_PLANTILLAS)
    On Error GoTo ErrorHandler
    
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then
        ObtenerFrecuenciaPlantilla = 0
        Exit Function
    End If
    
    Dim colID As Long
    Dim colFrecuencia As Long
    
    colID = tbl.ListColumns("ID Plantilla").Index
    colFrecuencia = tbl.ListColumns("Frecuencia meses").Index
    
    ' Buscar plantilla por ID
    For fila = 1 To tbl.DataBodyRange.Rows.Count
        If Trim(CStr(tbl.DataBodyRange.Cells(fila, colID).Value)) = Trim(idPlantilla) Then
            Dim valorFrecuencia As Variant
            valorFrecuencia = tbl.DataBodyRange.Cells(fila, colFrecuencia).Value
            
            If IsNumeric(valorFrecuencia) Then
                ObtenerFrecuenciaPlantilla = CLng(valorFrecuencia)
            Else
                ObtenerFrecuenciaPlantilla = 3  ' Default: 3 meses
            End If
            
            Exit Function
        End If
    Next fila
    
    ' No encontrada - usar default
    ObtenerFrecuenciaPlantilla = 3
    Exit Function
    
ErrorHandler:
    ObtenerFrecuenciaPlantilla = 3  ' Default en caso de error
End Function

' ══════════════════════════════════════════════════════════════
' NUEVA FUNCIÓN: Obtener nombre de categoría desde tblCategoriasRPN
' ══════════════════════════════════════════════════════════════
Private Function ObtenerNombreCategoria(ByVal numeroCategoria As Long) As String
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim tbl As ListObject
    Dim fila As Long
    
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    
    On Error Resume Next
    Set tbl = wsConfig.ListObjects("tblCategoriasRPN")
    On Error GoTo ErrorHandler
    
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then
        ObtenerNombreCategoria = "Sin categoría"
        Exit Function
    End If
    
    Dim colNumero As Long
    Dim colNombre As Long
    
    colNumero = tbl.ListColumns("Numero categoria").Index
    colNombre = tbl.ListColumns("Nombre categoria").Index
    
    ' Buscar categoría por número
    For fila = 1 To tbl.DataBodyRange.Rows.Count
        If CLng(tbl.DataBodyRange.Cells(fila, colNumero).Value) = numeroCategoria Then
            ObtenerNombreCategoria = Trim(CStr(tbl.DataBodyRange.Cells(fila, colNombre).Value))
            Exit Function
        End If
    Next fila
    
    ' No encontrada
    ObtenerNombreCategoria = "Sin categoría"
    Exit Function
    
ErrorHandler:
    ObtenerNombreCategoria = "Error al obtener categoría"
End Function

' ══════════════════════════════════════════════════════════════
' Limpiar plantilla después de generar PDF
' ══════════════════════════════════════════════════════════════
Private Sub LimpiarPlantillaCertificado(wsPlantilla As Worksheet)
    On Error Resume Next
    
    ' Solo limpiar celdas de DATOS — nunca los labels, títulos ni encabezados estáticos
    ' que fueron definidos por PlantillaCertificadoSetup y deben persistir entre ejecuciones
    
    ' ── NUEVO: Bloque Categoría ─────────────────────────────────
    wsPlantilla.Range("A6:G8").Value = "[CATEGORÍA SE LLENARÁ DINÁMICAMENTE]"
    wsPlantilla.Range("A6:G8").Interior.Color = RGB(255, 255, 255)  ' Blanco
    
    ' ── Sección 1: Datos de inspección (desplazadas +3) ─────────
    wsPlantilla.Cells(10, 3).ClearContents  ' C10 - Fecha
    wsPlantilla.Cells(10, 5).ClearContents  ' E10 - Hora inicio
    wsPlantilla.Cells(10, 7).ClearContents  ' G10 - Hora fin
    wsPlantilla.Cells(11, 3).ClearContents  ' C11 - Nombre evaluado
    wsPlantilla.Cells(11, 7).ClearContents  ' G11 - Iniciales
    wsPlantilla.Cells(12, 3).ClearContents  ' C12 - Puesto
    wsPlantilla.Cells(12, 6).ClearContents  ' F12 - Planta
    wsPlantilla.Cells(13, 3).ClearContents  ' C13 - Área
    wsPlantilla.Cells(13, 6).ClearContents  ' F13 - Línea
    wsPlantilla.Cells(14, 3).ClearContents  ' C14 - Lugar auditoría
    wsPlantilla.Cells(15, 3).ClearContents  ' C15 - Evaluador
    wsPlantilla.Cells(16, 4).ClearContents  ' D16 - AY1
    wsPlantilla.Cells(16, 6).ClearContents  ' F16 - AY2
    wsPlantilla.Cells(16, 7).ClearContents  ' G16 - OP
    
    ' ── Sección 2: Resultados (desplazadas +3) ──────────────────
    wsPlantilla.Cells(20, 3).ClearContents  ' C20 - TA puntaje
    wsPlantilla.Cells(20, 5).ClearContents  ' E20 - TA máximos
    wsPlantilla.Cells(21, 3).ClearContents  ' C21 - TA no aplica
    wsPlantilla.Cells(22, 3).ClearContents  ' C22 - Porcentaje
    wsPlantilla.Cells(23, 3).ClearContents  ' C23 - RPN
    wsPlantilla.Cells(24, 3).ClearContents  ' C24 - Categoría
    wsPlantilla.Cells(25, 3).ClearContents  ' C25 - Estado
    wsPlantilla.Cells(25, 3).Font.Color = RGB(0, 0, 0)  ' Restaurar color negro
    wsPlantilla.Cells(25, 3).Font.Bold = True  ' Mantener negrita por diseño
    
    ' ── Sección 3: Filas de respuestas (fila 27 en adelante, desplazada +3) ─────
    Dim ultimaFila As Long
    ultimaFila = wsPlantilla.UsedRange.Row + wsPlantilla.UsedRange.Rows.Count - 1
    If ultimaFila >= 27 Then
        wsPlantilla.Range( _
            wsPlantilla.Cells(27, 2), _
            wsPlantilla.Cells(ultimaFila, 7) _
        ).ClearContents
    End If
    
    ' ── Sección 4: Pie de validez (Paso 4 MVP - Fila 50) ─────────
    wsPlantilla.Range("A50:G50").Value = "Este certificado es válido hasta: [SE CALCULARÁ DINÁMICAMENTE]"
    
    On Error GoTo 0
End Sub

' ══════════════════════════════════════════════════════════════
' Generar nombre inteligente para archivo PDF (Paso 5 MVP)
' Formato: CERTIFICADO_[PUESTO]_[INICIALES]_[FECHA]_CAT[N]_[ESTADO].pdf
' ══════════════════════════════════════════════════════════════
Private Function GenerarNombreArchivoPDF(datosInspeccion As Object) As String
    Dim nombreArchivo As String
    
    ' Obtener componentes del nombre
    Dim iniciales As String
    Dim fechaStr As String
    Dim puesto As String
    Dim categoria As Long
    Dim estadoCorto As String
    
    iniciales = Trim(datosInspeccion("Iniciales"))
    fechaStr = Format(datosInspeccion("FechaInspeccion"), "YYYY-MM-DD")
    categoria = CLng(datosInspeccion("Categoria"))
    
    ' Obtener puesto (limpiar caracteres especiales para nombre de archivo)
    puesto = ObtenerPuestoPersonal(iniciales)
    If Len(puesto) = 0 Then
        puesto = "SIN_PUESTO"
    Else
        ' Si tiene múltiples puestos separados por coma, tomar solo el primero
        If InStr(puesto, ",") > 0 Then
            puesto = Trim(Left(puesto, InStr(puesto, ",") - 1))
        End If
        ' Limpiar espacios y caracteres no válidos para nombres de archivo
        puesto = Replace(puesto, " ", "_")
        puesto = Replace(puesto, "/", "_")
        puesto = Replace(puesto, "\", "_")
        puesto = Replace(puesto, ":", "_")
        puesto = UCase(puesto)  ' Mayúsculas para consistencia
    End If
    
    ' Determinar estado corto basado en categoría
    Select Case categoria
        Case 1, 2
            estadoCorto = "COMPETENTE"
        Case 3
            estadoCorto = "OBSERVACIONES"
        Case 4, 5
            estadoCorto = "NO_CALIFICADO"
        Case Else
            estadoCorto = "DESCONOCIDO"
    End Select
    
    ' Construir nombre completo
    ' Formato: CERTIFICADO_[PUESTO]_[INICIALES]_[FECHA]_CAT[N]_[ESTADO].pdf
    nombreArchivo = Configuration2.PDF_PREFIJO_NOMBRE & "_" & _
                    puesto & "_" & _
                    iniciales & "_" & _
                    fechaStr & "_" & _
                    "CAT" & categoria & "_" & _
                    estadoCorto & ".pdf"
    
    Debug.Print "[ARCHIVO] Nombre generado: " & nombreArchivo
    
    GenerarNombreArchivoPDF = nombreArchivo
End Function

' ══════════════════════════════════════════════════════════════
' Obtener ruta correcta del Desktop (robusto para múltiples idiomas y OneDrive)
' ══════════════════════════════════════════════════════════════
Private Function ObtenerRutaDesktop() As String
    On Error GoTo ErrorHandler
    
    ' Usar la API de Windows Shell para obtener la ruta correcta del Desktop
    ' Esto funciona incluso con OneDrive y en cualquier idioma
    Dim wshShell As Object
    Set wshShell = CreateObject("WScript.Shell")
    
    Dim desktopPath As String
    desktopPath = wshShell.SpecialFolders("Desktop")
    
    Debug.Print "[DEBUG] ObtenerRutaDesktop - Ruta obtenida desde Shell.SpecialFolders: " & desktopPath
    
    ' Verificar que la ruta existe
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(desktopPath) Then
        Debug.Print "[DEBUG] ObtenerRutaDesktop - Ruta verificada: " & desktopPath
        ObtenerRutaDesktop = desktopPath
        Exit Function
    End If
    
    ' Si falla, intentar rutas alternativas
    Debug.Print "[ERROR] ObtenerRutaDesktop - Shell.SpecialFolders falló, intentando alternativas..."
    
    Dim rutasIntento As Variant
    rutasIntento = Array( _
        Environ("USERPROFILE") & "\Desktop", _
        Environ("USERPROFILE") & "\Escritorio", _
        Environ("USERPROFILE") & "\OneDrive\Desktop", _
        Environ("USERPROFILE") & "\OneDrive\Escritorio", _
        Environ("USERPROFILE") & "\Documentos", _
        Environ("USERPROFILE") & "\Documents", _
        ThisWorkbook.Path _
    )
    
    Dim ruta As Variant
    For Each ruta In rutasIntento
        Debug.Print "[DEBUG] ObtenerRutaDesktop - Probando: " & ruta
        If fso.FolderExists(ruta) Then
            Debug.Print "[DEBUG] ObtenerRutaDesktop - Ruta encontrada: " & ruta
            ObtenerRutaDesktop = ruta
            Exit Function
        End If
    Next ruta
    
    ' Si nada funciona, usar carpeta del workbook
    Debug.Print "[ERROR] ObtenerRutaDesktop - Usando carpeta del libro: " & ThisWorkbook.Path
    ObtenerRutaDesktop = ThisWorkbook.Path
    Exit Function
    
ErrorHandler:
    Debug.Print "[ERROR] ObtenerRutaDesktop - Excepción: " & Err.Description
    ObtenerRutaDesktop = ThisWorkbook.Path
End Function

' ══════════════════════════════════════════════════════════════
' FUNCIÓN SIMPLE: GenerarPDFDesdeSeleccion
' Propósito: Genera PDF automáticamente desde la fila seleccionada
'            en la tabla de inspecciones (útil para botones)
' ══════════════════════════════════════════════════════════════
Public Sub GenerarPDFDesdeSeleccion()
    On Error GoTo ErrorHandler
    
    Dim wsh As Worksheet
    Dim tbl As ListObject
    Dim filaActiva As Long
    Dim idInspeccion As String
    
    ' Verificar que estamos en la hoja de Histórico
    If ActiveSheet.Name <> Configuration2.SHEET_HISTORICO Then
        MsgBox "Debe estar en la hoja '" & Configuration2.SHEET_HISTORICO & _
               "' para generar certificados.", vbInformation
        Exit Sub
    End If
    
    Set wsh = ActiveSheet
    Set tbl = wsh.ListObjects(Configuration2.TABLE_INSPECCIONES)
    
    ' Verificar que hay celda seleccionada dentro de tabla
    If Intersect(ActiveCell, tbl.DataBodyRange) Is Nothing Then
        MsgBox "Seleccione una inspección en la tabla para generar el certificado.", vbInformation
        Exit Sub
    End If
    
    ' Obtener ID de inspección (columna 1)
    filaActiva = ActiveCell.Row - tbl.DataBodyRange.Row + 1
    idInspeccion = Trim(tbl.DataBodyRange.Cells(filaActiva, 1).Value)
    
    If Len(idInspeccion) = 0 Then
        MsgBox "La fila seleccionada no tiene ID de inspección válido.", vbExclamation
        Exit Sub
    End If
    
    ' Generar PDF
    Call GenerarCertificadoPDF(idInspeccion)
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Call ErrorLogger2.Log("CertificadoPDFGenerator.GenerarPDFDesdeSeleccion", Err.Description, Err.Number)
End Sub
