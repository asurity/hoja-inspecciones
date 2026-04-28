' --------------------------------------------------------------
' Módulo: CertificadoPDFGenerator
' Descripción: Genera PDFs de certificados de inspección a partir
'              de datos completados en tblInspecciones + tblRespuestas
' Fecha creación: 17/04/2026
' Dependencias: Configuration2, PlantillaCertificadoSetup,
'               InspectionRepository, ChecklistRepository
' --------------------------------------------------------------
Option Explicit

' --------------------------------------------------------------
' FUNCIÓN PRINCIPAL: GenerarCertificadoPDF
' Propósito: Genera un PDF de certificado a partir del ID inspección
' --------------------------------------------------------------
Public Sub GenerarCertificadoPDF(ByVal idInspeccion As String)
    On Error GoTo ErrorHandler
    
    ' Debug.Print "===== INICIO GenerarCertificadoPDF ====="
    ' Debug.Print "ID Inspección: " & idInspeccion
    
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
    ' Debug.Print "[DEBUG] Obteniendo hoja Histórico..."
    Set wsh = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    ' Debug.Print "[DEBUG] Hoja Histórico: " & wsh.Name
    
    ' Desproteger plantilla si existe
    ' Debug.Print "[DEBUG] Buscando plantilla certificado..."
    On Error Resume Next
    Set wsPlantilla = ThisWorkbook.Sheets(Configuration2.SHEET_PLANTILLA_CERTIFICADO)
    On Error GoTo ErrorHandler
    
    If wsPlantilla Is Nothing Then
        ' Debug.Print "[ERROR] Plantilla no encontrada"
        MsgBox "ERROR: No existe la plantilla de certificado." & vbCrLf & _
               "Ejecute primero: PlantillaCertificadoSetup.InicializarPlantillaCertificado()", _
               vbCritical
        Exit Sub
    End If
    
    ' Debug.Print "[DEBUG] Plantilla encontrada: " & wsPlantilla.Name
    ' Debug.Print "[DEBUG] Estado plantilla: " & wsPlantilla.Visible
    
    Dim datosInspeccion As Object
    ' Debug.Print "[DEBUG] Obteniendo datos de inspección..."
    Set datosInspeccion = ObtenerDatosInspeccion(idInspeccion)
    
    If datosInspeccion Is Nothing Then
        ' Debug.Print "[ERROR] Datos de inspección no encontrados"
        MsgBox "No se encontró la inspección con ID: " & idInspeccion, vbCritical
        Exit Sub
    End If
    
    ' Debug.Print "[DEBUG] Datos de inspección obtenidos correctamente"
    
    ' DESPROTEGER WORKBOOK para permitir cambio de visibilidad (URS-22)
    Dim estabaProtegido As Boolean
    estabaProtegido = ThisWorkbook.ProtectStructure
    
    If estabaProtegido Then
        ' Debug.Print "[DEBUG] Desprotegiendo workbook para cambiar visibilidad..."
        Call WorkbookProtector2.UnprotectWorkbook
    End If
    
    ' Hacer visible la plantilla temporalmente
    ' Debug.Print "[DEBUG] Haciendo plantilla visible..."
    wsPlantilla.Visible = xlSheetVisible
    ' Debug.Print "[DEBUG] Plantilla visible: " & (wsPlantilla.Visible = xlSheetVisible)
    
    ' Poblar plantilla con datos
    ' Debug.Print "[DEBUG] Poblando plantilla con datos..."
    Call PoblarPlantillaCertificado(wsPlantilla, datosInspeccion)
    ' Debug.Print "[DEBUG] Plantilla poblada"
    
    ' Generar nombre de archivo
    nombreArchivo = GenerarNombreArchivoPDF(datosInspeccion)
    ' Debug.Print "[DEBUG] Nombre archivo: " & nombreArchivo
    
    ' Obtener ruta correcta (Desktop o Escritorio según idioma Windows)
    ' Debug.Print "[DEBUG] Detectando carpeta Desktop..."
    Dim desktopPath As String
    desktopPath = ObtenerRutaDesktop()
    ' Debug.Print "[DEBUG] Ruta Desktop detectada: " & desktopPath
    
    rutaPDF = desktopPath & "\" & nombreArchivo
    ' Debug.Print "[DEBUG] Ruta PDF final: " & rutaPDF
    
    ' Verificar si la carpeta existe
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Debug.Print "[DEBUG] Verificando que la carpeta destino existe..."
    If Not fso.FolderExists(desktopPath) Then
        ' Debug.Print "[ERROR] Carpeta no existe: " & desktopPath
        MsgBox "Error: No se encuentra la carpeta: " & desktopPath, vbCritical
        Exit Sub
    End If
    ' Debug.Print "[DEBUG] Carpeta destino verificada: " & desktopPath
    
    ' Debug: Verificar contenido de la plantilla
    ' Debug.Print "[DEBUG] Verificando contenido de plantilla..."
    ' Debug.Print "[DEBUG] Celdas usadas en plantilla: " & wsPlantilla.UsedRange.Address
    ' Debug.Print "[DEBUG] Filas usadas: " & wsPlantilla.UsedRange.Rows.Count
    ' Debug.Print "[DEBUG] Columnas usadas: " & wsPlantilla.UsedRange.Columns.Count
    
    ' Procesar solo el rango usado
    Dim rangoExportar As Range
    Set rangoExportar = wsPlantilla.UsedRange
    ' Debug.Print "[DEBUG] Rango a exportar: " & rangoExportar.Address
    
    ' Configurar números de página en el pie de página
    ' Debug.Print "[DEBUG] Configurando números de página..."
    With wsPlantilla.PageSetup
        ' Pie de página centrado: "Página X de Y"
        .CenterFooter = "Página &P de &N"
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    ' Debug.Print "[DEBUG] Números de página configurados"
    
    ' Exportar como PDF
    ' Debug.Print "[DEBUG] Iniciando exportación a PDF..."
    ' Debug.Print "[DEBUG] Método: ExportAsFixedFormat(xlTypePDF, ...)"
    
    On Error Resume Next
    wsPlantilla.ExportAsFixedFormat xlTypePDF, rutaPDF, , True
    Dim errExport As Long
    errExport = Err.Number
    Dim descExport As String
    descExport = Err.Description
    On Error GoTo ErrorHandler
    
    If errExport <> 0 Then
        ' Debug.Print "[ERROR] Exportación fallida"
        ' Debug.Print "[ERROR] Error Number: " & errExport
        ' Debug.Print "[ERROR] Error Description: " & descExport
        Err.Raise errExport, "ExportAsFixedFormat", descExport
    End If
    
    ' Debug.Print "[DEBUG] PDF generado exitosamente"
    
    ' Verificar que el archivo fue creado
    ' Debug.Print "[DEBUG] Verificando archivo creado..."
    If fso.FileExists(rutaPDF) Then
        ' Debug.Print "[DEBUG] Archivo existe: " & rutaPDF
        ' Debug.Print "[DEBUG] Tamaño archivo: " & fso.GetFile(rutaPDF).Size & " bytes"
    Else
        ' Debug.Print "[ERROR] Archivo NO fue creado"
        MsgBox "Error: El PDF no fue creado en: " & rutaPDF, vbCritical
        Exit Sub
    End If
    
    ' ====================================================================================
    ' REGISTRO AUDIT TRAIL: Emisión de certificado PDF
    ' Fecha: 25/04/2026
    ' Propósito: Trazabilidad completa de emisión de documentos oficiales
    ' ====================================================================================
    Dim userName As String
    userName = Application.UserName  ' Nombre configurado en Excel (consistente con menú principal)
    
    ' Construir detalle completo del certificado emitido
    Dim detallesCertificado As String
    detallesCertificado = "ID Inspección: " & idInspeccion & _
                          " | Evaluado: " & CStr(datosInspeccion("Iniciales")) & _
                          " | Puesto: " & CStr(datosInspeccion("Puesto")) & _
                          " | RPN: " & Format(datosInspeccion("RPN"), "0.00")
    
    ' Agregar categoría si existe
    If datosInspeccion.Exists("Categoria") Then
        detallesCertificado = detallesCertificado & " | Categoría: " & CStr(datosInspeccion("Categoria"))
    End If
    
    ' Agregar nombre de archivo y usuario que emitió
    detallesCertificado = detallesCertificado & _
                          " | Archivo: " & nombreArchivo & _
                          " | Emitido por: " & userName
    
    Call AuditLogger2.LogAction( _
        action:="Emisión de certificado PDF", _
        sheetName:=Configuration2.SHEET_HISTORICO, _
        dataModified:="Certificado generado", _
        beforeChange:="N/A", _
        afterChange:=detallesCertificado, _
        moduleAndSubroutine:="CertificadoPDFGenerator.GenerarCertificadoPDF" _
    )
    ' ====================================================================================
    
    ' Ocultar plantilla nuevamente
    ' Debug.Print "[DEBUG] Ocultando plantilla..."
    wsPlantilla.Visible = xlSheetVeryHidden
    ' Debug.Print "[DEBUG] Plantilla ocultada"
    
    ' REPROTEGER WORKBOOK si estaba protegido (URS-22)
    If estabaProtegido Then
        ' Debug.Print "[DEBUG] Reprotegiendo workbook..."
        Call WorkbookProtector2.ProtectWorkbook
    End If
    
    ' Limpiar plantilla
    ' Debug.Print "[DEBUG] Limpiando plantilla..."
    Call LimpiarPlantillaCertificado(wsPlantilla)
    ' Debug.Print "[DEBUG] Plantilla limpia"
    
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
    
    ' Debug.Print "===== FIN GenerarCertificadoPDF - ÉXITO ====="
    Exit Sub
    
ErrorHandler:
    ' Debug.Print "===== ERROR EN GenerarCertificadoPDF ====="
    ' Debug.Print "Error Number: " & Err.Number
    ' Debug.Print "Error Description: " & Err.Description
    ' Debug.Print "Error Source: " & Err.Source
    
    Call ErrorLogger2.Log("CertificadoPDFGenerator.GenerarCertificadoPDF", Err.Description, Err.Number)
    
    ' Asegurar que plantilla queda oculta
    On Error Resume Next
    If Not wsPlantilla Is Nothing Then
        wsPlantilla.Visible = xlSheetVeryHidden
    End If
    
    ' Reproteger workbook si se desprotegió
    If estabaProtegido Then
        Call WorkbookProtector2.ProtectWorkbook
    End If
    On Error GoTo 0
    On Error GoTo 0
    
    MsgBox "Error al generar certificado:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error"
    ' Debug.Print "===== FIN GenerarCertificadoPDF - ERROR ====="
End Sub

' --------------------------------------------------------------
' Obtener todos los datos de una inspección
' --------------------------------------------------------------
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
    
    ' Fecha Auditada (puede no existir en inspecciones antiguas)
    On Error Resume Next
    Dim colFechaAuditada As Variant
    colFechaAuditada = Application.Match("Fecha Auditada", tbl.HeaderRowRange, 0)
    If Not IsError(colFechaAuditada) Then
        datos("FechaAuditada") = fila.Cells(1, CLng(colFechaAuditada)).Value
    Else
        datos("FechaAuditada") = datos("FechaInspeccion")  ' Fallback a fecha inspección
    End If
    On Error GoTo ErrorHandler
    
    ' Puesto Evaluado (columna 34 - inspecciones recurrentes)
    On Error Resume Next
    Dim colPuestoEval As Variant
    colPuestoEval = Application.Match("Puesto Evaluado", tbl.HeaderRowRange, 0)
    If Not IsError(colPuestoEval) Then
        Dim puestoEval As String
        puestoEval = Trim(CStr(fila.Cells(1, CLng(colPuestoEval)).Value))
        If Len(puestoEval) > 0 Then
            datos("PuestoEvaluado") = puestoEval
        End If
    End If
    On Error GoTo ErrorHandler
    
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
                ' Debug.Print "[CertificadoPDF] Auditoría Procesos encontrada en columna '" & nombreCol & "': " & resultadoProcesos
                Exit For
            End If
        End If
    Next nombreCol
    On Error GoTo ErrorHandler
    
    datos("AuditoriaProcesosResultado") = resultadoProcesos
    
    ' Cargar conteos de Auditoría de Procesos (columnas agregadas 22/04/2026)
    On Error Resume Next
    Dim colAPCritica As Variant
    Dim colAPMayor As Variant
    Dim colAPMenor As Variant
    
    colAPCritica = Application.Match("AP Critica No Cumple", tbl.HeaderRowRange, 0)
    If Not IsError(colAPCritica) Then
        datos("APCriticaNoCumple") = CLng(fila.Cells(1, CLng(colAPCritica)).Value)
    Else
        datos("APCriticaNoCumple") = 0
    End If
    
    colAPMayor = Application.Match("AP Mayor No Cumple", tbl.HeaderRowRange, 0)
    If Not IsError(colAPMayor) Then
        datos("APMayorNoCumple") = CLng(fila.Cells(1, CLng(colAPMayor)).Value)
    Else
        datos("APMayorNoCumple") = 0
    End If
    
    colAPMenor = Application.Match("AP Menor No Cumple", tbl.HeaderRowRange, 0)
    If Not IsError(colAPMenor) Then
        datos("APMenorNoCumple") = CLng(fila.Cells(1, CLng(colAPMenor)).Value)
    Else
        datos("APMenorNoCumple") = 0
    End If
    On Error GoTo ErrorHandler
    
    ' Debug.Print "[CertificadoPDF] Conteos AP - Crítica: " & datos("APCriticaNoCumple") & ", Mayor: " & datos("APMayorNoCumple") & ", Menor: " & datos("APMenorNoCumple")
    
    ' Cargar RPN Total y RPN Anterior Manual (inspecciones recurrentes)
    On Error Resume Next
    Dim colRPNTotal As Variant
    Dim colRPNAnterior As Variant
    Dim colPorcRecuperacion As Variant
    Dim colPorcOOL As Variant
    
    colRPNTotal = Application.Match("RPN Total", tbl.HeaderRowRange, 0)
    If Not IsError(colRPNTotal) Then
        datos("RPNTotal") = CDbl(fila.Cells(1, CLng(colRPNTotal)).Value)
    Else
        datos("RPNTotal") = 0
    End If
    
    colRPNAnterior = Application.Match("RPN Anterior Manual", tbl.HeaderRowRange, 0)
    If Not IsError(colRPNAnterior) Then
        datos("RPNAnterior") = CDbl(fila.Cells(1, CLng(colRPNAnterior)).Value)
    Else
        datos("RPNAnterior") = 0
    End If
    
    ' Cargar Porcentaje Recuperación (columna 38)
    colPorcRecuperacion = Application.Match("Porcentaje Recuperacion", tbl.HeaderRowRange, 0)
    If Not IsError(colPorcRecuperacion) Then
        Dim valorRecup As String
        valorRecup = Trim(CStr(fila.Cells(1, CLng(colPorcRecuperacion)).Value))
        ' Si es Configuration2.VALOR_NO_APLICA o vacío, usar 0
        If Configuration2.EsValorNoAplica(valorRecup) Then
            datos("PorcentajeRecuperacion") = 0
        Else
            datos("PorcentajeRecuperacion") = CDbl(valorRecup)
        End If
    Else
        datos("PorcentajeRecuperacion") = 0
    End If
    
    ' Cargar Porcentaje OOL (columna 39)
    colPorcOOL = Application.Match("Porcentaje OOL", tbl.HeaderRowRange, 0)
    If Not IsError(colPorcOOL) Then
        Dim valorOOL As String
        valorOOL = Trim(CStr(fila.Cells(1, CLng(colPorcOOL)).Value))
        ' Si es Configuration2.VALOR_NO_APLICA o vacío, usar 0
        If Configuration2.EsValorNoAplica(valorOOL) Then
            datos("PorcentajeOOL") = 0
        Else
            datos("PorcentajeOOL") = CDbl(valorOOL)
        End If
    Else
        datos("PorcentajeOOL") = 0
    End If
    On Error GoTo ErrorHandler
    
    ' Debug.Print "[CertificadoPDF] RPN Total: " & datos("RPNTotal") & ", RPN Anterior: " & datos("RPNAnterior") & ", % Recuperación: " & datos("PorcentajeRecuperacion") & ", % OOL: " & datos("PorcentajeOOL")
    
    ' -------------------------------------------------------------------
    ' CARGAR CALIFICACIONES VESTUARIO/OPERADOR (FASE 7 - 23/04/2026)
    ' Columnas 44-47 en tblInspecciones
    ' -------------------------------------------------------------------
    On Error Resume Next
    Dim colCalifVest As Variant
    Dim colFechaVencVest As Variant
    Dim colCalifOper As Variant
    Dim colFechaVencOper As Variant
    
    colCalifVest = Application.Match("Calificacion Vestuario", tbl.HeaderRowRange, 0)
    If Not IsError(colCalifVest) Then
        datos("CalificacionVestuario") = Trim(CStr(fila.Cells(1, CLng(colCalifVest)).Value))
    Else
        datos("CalificacionVestuario") = Configuration2.VALOR_NO_APLICA
    End If
    
    colFechaVencVest = Application.Match("Fecha Venc Vestuario", tbl.HeaderRowRange, 0)
    If Not IsError(colFechaVencVest) Then
        datos("FechaVencVestuario") = Trim(CStr(fila.Cells(1, CLng(colFechaVencVest)).Value))
    Else
        datos("FechaVencVestuario") = Configuration2.VALOR_NO_APLICA
    End If
    
    colCalifOper = Application.Match("Calificacion Operador", tbl.HeaderRowRange, 0)
    If Not IsError(colCalifOper) Then
        datos("CalificacionOperador") = Trim(CStr(fila.Cells(1, CLng(colCalifOper)).Value))
    Else
        datos("CalificacionOperador") = Configuration2.VALOR_NO_APLICA
    End If
    
    colFechaVencOper = Application.Match("Fecha Venc Operador", tbl.HeaderRowRange, 0)
    If Not IsError(colFechaVencOper) Then
        datos("FechaVencOperador") = Trim(CStr(fila.Cells(1, CLng(colFechaVencOper)).Value))
    Else
        datos("FechaVencOperador") = Configuration2.VALOR_NO_APLICA
    End If
    On Error GoTo ErrorHandler
    
    ' Debug.Print "[CertificadoPDF] Calificaciones - Vestuario: " & datos("CalificacionVestuario") & " (Venc: " & datos("FechaVencVestuario") & "), Operador: " & datos("CalificacionOperador") & " (Venc: " & datos("FechaVencOperador") & ")"
    
    ' Cargar indicador de inspección recurrente (columna 33 - agregada Fase 5)
    On Error Resume Next
    Dim colEsRecurrente As Variant
    colEsRecurrente = Application.Match("Es Inspeccion Recurrente", tbl.HeaderRowRange, 0)
    If Not IsError(colEsRecurrente) Then
        Dim valorRecurrente As String
        valorRecurrente = Trim(UCase(CStr(fila.Cells(1, CLng(colEsRecurrente)).Value)))
        ' Convertir "SI"/"NO" a Boolean
        datos("EsInspeccionRecurrente") = (valorRecurrente = "SI" Or valorRecurrente = "SÍ")
        ' Debug.Print "[CertificadoPDF] Es Inspección Recurrente: " & datos("EsInspeccionRecurrente") & " (valor tabla: '" & valorRecurrente & "')"
    Else
        ' Columna no existe (inspecciones antiguas), asumir no recurrente
        datos("EsInspeccionRecurrente") = False
        ' Debug.Print "[CertificadoPDF] Columna 'Es Inspeccion Recurrente' no existe, asumiendo No Recurrente"
    End If
    On Error GoTo ErrorHandler
    
    ' FIX 21/04/2026: Usar ListColumns en lugar de índice hardcodeado
    datos("ObservacionesGenerales") = fila.Cells(1, tbl.ListColumns("Observaciones generales").Index).Value
    
    ' Obtener respuestas
    Set datos("Respuestas") = ObtenerRespuestasInspeccion(idInspeccion)
    
    Set ObtenerDatosInspeccion = datos
    Exit Function
    
ErrorHandler:
    Set ObtenerDatosInspeccion = Nothing
End Function

' --------------------------------------------------------------
' Obtener respuestas de inspección
' --------------------------------------------------------------
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

' --------------------------------------------------------------
' Poblar plantilla con datos de inspección
' --------------------------------------------------------------
Private Sub PoblarPlantillaCertificado(wsPlantilla As Worksheet, datosInspeccion As Object)
    On Error Resume Next
    
    ' Debug.Print "[POBLACIÓN] Iniciando población de plantilla..."
    
    ' Obtener puesto desde tblPersonal (hoja correcta: SHEET_PERSONAL)
    Dim nombreEvaluado As String
    Dim puesto As String
    nombreEvaluado = ObtenerNombrePersonal(datosInspeccion("Iniciales"))
    puesto = ObtenerPuestoPersonal(datosInspeccion("Iniciales"))
    If Len(nombreEvaluado) = 0 Then nombreEvaluado = datosInspeccion("Iniciales")
    If Len(puesto) = 0 Then puesto = "No especificado"
    
    ' -----------------------------------------------------------
    ' NUEVO: SECCIÓN 0 - BLOQUE CATEGORÍA (Filas 6-8)
    ' -----------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Poblando bloque categoría..."
    
    Dim textoCategoria As String
    Dim nombreCategoria As String
    Dim colorFondo As Long
    Dim numCategoria As Long
    
    numCategoria = CLng(datosInspeccion("Categoria"))
    
    ' Obtener nombre categoría desde tblCategoriasRPN
    nombreCategoria = ObtenerNombreCategoria(numCategoria)
    
    ' Construir texto del bloque (solo categoría, sin RPN)
    textoCategoria = "CATEGORÍA " & numCategoria
    
    ' Asignar a celda A6:D8 (rango combinado - 4 columnas)
    wsPlantilla.Range("A6:D8").Value = textoCategoria
    
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
    
    wsPlantilla.Range("A6:D8").Interior.Color = colorFondo
    
    ' Debug.Print "[POBLACIÓN] Bloque categoría completado: Cat " & numCategoria & " - " & nombreCategoria
    
    ' -----------------------------------------------------------
    ' SECCIÓN 1: DATOS DE INSPECCIÓN (Filas 9-17, desplazadas +3)
    ' Usamos referencias de celda directas (layout definido en PlantillaCertificadoSetup)
    ' -----------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Poblando datos de inspección..."
    
    ' Determinar puesto a mostrar (usar PuestoEvaluado si existe, sino usar el puesto del personal)
    Dim puestoAMostrar As String
    If datosInspeccion.Exists("PuestoEvaluado") And Len(Trim(CStr(datosInspeccion("PuestoEvaluado")))) > 0 Then
        puestoAMostrar = Trim(CStr(datosInspeccion("PuestoEvaluado")))
    Else
        puestoAMostrar = puesto
    End If
    
    ' NUEVA ESTRUCTURA - 4 COLUMNAS
    wsPlantilla.Cells(10, 2).Value = Format(datosInspeccion("FechaInspeccion"), "DD-MM-YYYY")   ' B10 - Fecha ejecución
    wsPlantilla.Cells(11, 2).Value = Format(datosInspeccion("HoraInicio"), "HH:MM") & " - " & Format(datosInspeccion("HoraTermino"), "HH:MM")  ' B11 - Rango horas
    wsPlantilla.Cells(12, 2).Value = datosInspeccion("Planta")                                  ' B12 - Planta
    wsPlantilla.Cells(13, 2).Value = datosInspeccion("LineaAuditada")                            ' B13 - Línea auditada
    wsPlantilla.Cells(14, 2).Value = Format(datosInspeccion("FechaAuditada"), "DD-MM-YYYY")     ' B14 - Fecha evaluada
    wsPlantilla.Cells(15, 2).Value = datosInspeccion("Iniciales")                               ' B15 - Iniciales evaluado
    wsPlantilla.Cells(16, 2).Value = puestoAMostrar                                              ' B16 - Puesto evaluado
    wsPlantilla.Cells(17, 2).Value = datosInspeccion("Area")                                     ' B17 - Área
    wsPlantilla.Cells(18, 2).Value = datosInspeccion("LugarAuditoria")                           ' B18 - Lugar auditoría
    wsPlantilla.Cells(19, 2).Value = datosInspeccion("Auditor")                                  ' B19 - Evaluador
    wsPlantilla.Cells(20, 3).Value = datosInspeccion("AY1")                                      ' C20 - AY1
    wsPlantilla.Cells(21, 3).Value = datosInspeccion("AY2")                                      ' C21 - AY2
    wsPlantilla.Cells(22, 3).Value = datosInspeccion("OP")                                       ' C22 - OP
    
    ' Debug.Print "[POBLACIÓN] Datos básicos completados"
    
    ' -------------------------------------------------------------------
    ' NUEVA SECCIÓN: CALIFICACIONES VESTUARIO/OPERADOR (FASE 7 - 23/04/2026)
    ' Filas 23-24
    ' -------------------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Poblando calificaciones vestuario/operador..."
    
    ' Fila 23: Calificación Vestuario
    wsPlantilla.Cells(23, 2).Value = datosInspeccion("CalificacionVestuario")  ' B23 - Calificación Vestuario
    wsPlantilla.Cells(23, 4).Value = datosInspeccion("FechaVencVestuario")     ' D23 - Fecha Venc Vestuario
    
    ' Aplicar formato condicional a D23 (fecha vencimiento vestuario)
    Call AplicarFormatoFechaVencimiento(wsPlantilla, 23, 4, datosInspeccion("FechaVencVestuario"))
    
    ' Fila 24: Calificación Operador
    wsPlantilla.Cells(24, 2).Value = datosInspeccion("CalificacionOperador")   ' B24 - Calificación Operador
    wsPlantilla.Cells(24, 4).Value = datosInspeccion("FechaVencOperador")      ' D24 - Fecha Venc Operador
    
    ' Aplicar formato condicional a D24 (fecha vencimiento operador)
    Call AplicarFormatoFechaVencimiento(wsPlantilla, 24, 4, datosInspeccion("FechaVencOperador"))
    
    ' Debug.Print "[POBLACIÓN] Calificaciones completadas"
    
    ' -----------------------------------------------------------
    ' SECCIÓN 2: AUDITORÍA DE PROCESOS (B26-B28) - DESPLAZADO +2 FILAS
    ' -----------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Poblando conteos Auditoría de Procesos..."
    
    wsPlantilla.Cells(26, 2).Value = datosInspeccion("APCriticaNoCumple")  ' B26 - No Cumple Críticos (era B24)
    wsPlantilla.Cells(27, 2).Value = datosInspeccion("APMayorNoCumple")    ' B27 - No Cumple Mayores (era B25)
    wsPlantilla.Cells(28, 2).Value = datosInspeccion("APMenorNoCumple")    ' B28 - No Cumple Menores (era B26)
    
    ' Debug.Print "[POBLACIÓN] Conteos AP - Crítica: " & datosInspeccion("APCriticaNoCumple") & ", Mayor: " & datosInspeccion("APMayorNoCumple") & ", Menor: " & datosInspeccion("APMenorNoCumple")
    
    ' Evaluar resultado de Auditoría de Procesos (B29:D31) - DESPLAZADO +2 FILAS
    Dim resultadoAP As String
    resultadoAP = EvaluarResultadoAP(datosInspeccion("APCriticaNoCumple"), datosInspeccion("APMayorNoCumple"))
    
    wsPlantilla.Range("B29:D31").Merge
    wsPlantilla.Cells(29, 2).Value = resultadoAP
    
    ' Formato del resultado AP
    With wsPlantilla.Cells(29, 2)
        .Font.Name = "Arial"
        .Font.Size = 12
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' Color según resultado
        If resultadoAP = "Cumple" Then
            .Interior.Color = RGB(212, 244, 230)  ' Verde claro
            .Font.Color = RGB(39, 174, 96)        ' Verde
        Else
            .Interior.Color = RGB(253, 223, 223)  ' Rojo claro
            .Font.Color = RGB(203, 67, 53)        ' Rojo
        End If
    End With
    
    ' Debug.Print "[POBLACIÓN] Resultado AP: " & resultadoAP
    
    ' -----------------------------------------------------------
    ' SECCIÓN 3: RESULTADOS TÉCNICOS (B33-B38, D36-D37) - DESPLAZADO +2 FILAS
    ' -----------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Poblando resultados técnicos..."
    
    wsPlantilla.Cells(33, 2).Value = datosInspeccion("TAMaximos")          ' B33 - Puntos Máximos (era B31)
    wsPlantilla.Cells(34, 2).Value = datosInspeccion("TAPuntaje")          ' B34 - Puntos Obtenidos (era B32)
    wsPlantilla.Cells(35, 2).Value = datosInspeccion("TANoAplica")         ' B35 - Puntos No Aplica (era B33)
    wsPlantilla.Cells(36, 2).Value = Format(datosInspeccion("TAPorcentaje"), "0.00")  ' B36 - % TA (era B34)
    
    ' B37 - % TA Anterior (RPN Total anterior si existe) - DESPLAZADO
    Dim rpnAnterior As Double
    rpnAnterior = 0
    If datosInspeccion.Exists("RPNAnterior") Then
        On Error Resume Next
        rpnAnterior = CDbl(datosInspeccion("RPNAnterior"))
        If Err.Number <> 0 Then rpnAnterior = 0
        Err.Clear
        On Error GoTo 0
    End If
    wsPlantilla.Cells(37, 2).Value = Format(rpnAnterior, "0.00")    ' B37 - % TA Anterior (era B35)
    
    ' B36 - RPN Total
    ' SIMPLIFICADO: Siempre usar RPN Total de la BD (se guarda tanto en primera como recurrente)
    ' - Primera inspección: RPN Total = % TA
    ' - Inspección recurrente: RPN Total = Promedio + factores
    Dim rpnTotal As Double
    rpnTotal = 0
    
    If datosInspeccion.Exists("RPNTotal") Then
        On Error Resume Next
        rpnTotal = CDbl(datosInspeccion("RPNTotal"))
        If Err.Number <> 0 Then
            ' Si hay error, usar RPN calculado como fallback
            rpnTotal = CDbl(datosInspeccion("RPN"))
            ' Debug.Print "[CertificadoPDF] Error al leer RPN Total, usando RPN calculado: " & rpnTotal
        Else
            ' Debug.Print "[CertificadoPDF] RPN Total leído de BD: " & rpnTotal
        End If
        Err.Clear
        On Error GoTo 0
    Else
        ' Columna no existe (inspección antigua), usar RPN calculado
        rpnTotal = CDbl(datosInspeccion("RPN"))
        ' Debug.Print "[CertificadoPDF] RPN Total no existe en BD, usando RPN calculado: " & rpnTotal
    End If
    
    wsPlantilla.Cells(38, 2).Value = Format(rpnTotal, "0.00")              ' B38 - RPN Total (era B36)
    
    ' D36 - % Recuperación microbiológica (leer de BD si existe)
    Dim porcRecuperacion As Double
    porcRecuperacion = 0
    If datosInspeccion.Exists("PorcentajeRecuperacion") Then
        On Error Resume Next
        porcRecuperacion = CDbl(datosInspeccion("PorcentajeRecuperacion"))
        If Err.Number <> 0 Then porcRecuperacion = 0
        Err.Clear
        On Error GoTo 0
    End If
    wsPlantilla.Cells(36, 4).Value = Format(porcRecuperacion, "0.00")  ' D36 - % Recovery
    
    ' D37 - % OOL (leer de BD si existe)
    Dim porcOOL As Double
    porcOOL = 0
    If datosInspeccion.Exists("PorcentajeOOL") Then
        On Error Resume Next
        porcOOL = CDbl(datosInspeccion("PorcentajeOOL"))
        If Err.Number <> 0 Then porcOOL = 0
        Err.Clear
        On Error GoTo 0
    End If
    wsPlantilla.Cells(37, 4).Value = Format(porcOOL, "0.00")            ' D37 - % OOL
    
    ' Debug.Print "[POBLACIÓN] Resultados técnicos completados"
    
    ' -----------------------------------------------------------
    ' SECCIÓN 4: ESTADO DE COMPETENCIA (basado en categoría)
    ' -----------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Determinando estado de competencia..."
    
    ' Determinar estado de competencia (podría mostrarse en otra celda si se necesita)
    Dim textoEstado As String
    Dim colorEstado As Long
    
    Select Case numCategoria
        Case 1, 2
            textoEstado = ChrW(10004) & " COMPETENTE"  ' ? COMPETENTE
            colorEstado = RGB(39, 174, 96)  ' Verde (#27AE60)
        Case 3
            textoEstado = ChrW(9888) & " COMPETENTE CON OBSERVACIONES"  ' ? COMPETENTE CON OBSERVACIONES
            colorEstado = RGB(243, 156, 18)  ' Naranja (#F39C12)
        Case 4, 5
            textoEstado = ChrW(10006) & " NO CALIFICADO"  ' ? NO CALIFICADO
            colorEstado = RGB(203, 67, 53)  ' Rojo (#CB4335)
        Case Else
            textoEstado = "ESTADO DESCONOCIDO"
            colorEstado = RGB(0, 0, 0)  ' Negro por defecto
    End Select
    
    ' Debug.Print "[POBLACIÓN] Estado asignado: " & textoEstado & " (Cat " & numCategoria & ")"
    
    ' -----------------------------------------------------------
    ' SECCIÓN 3: TABLA DE PREGUNTAS Y RESPUESTAS
    ' -----------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Poblando tabla de respuestas..."
    
    On Error GoTo 0
    
    ' Calcular fila de inicio para tabla de preguntas (2 filas después de resultados)
    Dim filaInicioPreguntas As Long
    filaInicioPreguntas = 40  ' Última fila de resultados es 38, entonces 40 para encabezado
    
    ' Poblar tabla de respuestas y obtener última fila
    Dim ultimaFilaRespuestas As Long
    ultimaFilaRespuestas = PoblarTablaRespuestas(wsPlantilla, datosInspeccion, filaInicioPreguntas)
    
    ' -----------------------------------------------------------
    ' SECCIÓN DE FEEDBACK (3 filas después de las preguntas)
    ' -----------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Agregando sección de Feedback..."
    
    Dim filaFeedback As Long
    filaFeedback = ultimaFilaRespuestas + 2  ' 1 fila vacía de separación
    
    ' Título de la sección (fila N)
    wsPlantilla.Range(wsPlantilla.Cells(filaFeedback, 1), wsPlantilla.Cells(filaFeedback, 4)).Merge
    wsPlantilla.Cells(filaFeedback, 1).Value = "SECCIÓN DE FEEDBACK"
    
    With wsPlantilla.Cells(filaFeedback, 1)
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(189, 215, 238)  ' Azul claro (mismo que encabezados)
    End With
    
    ' Fila N+1: Primer párrafo con espacio para fecha
    Dim filaTexto1 As Long
    filaTexto1 = filaFeedback + 1
    
    wsPlantilla.Range(wsPlantilla.Cells(filaTexto1, 1), wsPlantilla.Cells(filaTexto1, 4)).Merge
    wsPlantilla.Cells(filaTexto1, 1).Value = "Con fecha ___________________________ se informa del resultado de la auditoría y se revisan no conformidades encontradas en el comportamiento del personal en áreas clasificadas de Therapia iv."
    
    With wsPlantilla.Cells(filaTexto1, 1)
        .Font.Name = "Arial"
        .Font.Size = 10
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ' Fila N+2: Segundo párrafo
    Dim filaTexto2 As Long
    filaTexto2 = filaFeedback + 2
    
    wsPlantilla.Range(wsPlantilla.Cells(filaTexto2, 1), wsPlantilla.Cells(filaTexto2, 4)).Merge
    wsPlantilla.Cells(filaTexto2, 1).Value = "He tomado conocimiento de los hallazgos encontrados a fin de establecer un plan de trabajo para eliminar y mitigar el riesgo de las brechas detectadas."
    
    With wsPlantilla.Cells(filaTexto2, 1)
        .Font.Name = "Arial"
        .Font.Size = 10
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ' Fila N+3: Vacía (se deja en blanco intencionalmente)
    
    ' Fila N+4: Línea con opciones Sí, No y Firma
    Dim filaOpciones As Long
    filaOpciones = filaFeedback + 4
    
    ' Columna A: Sí
    wsPlantilla.Cells(filaOpciones, 1).Value = "Sí"
    
    With wsPlantilla.Cells(filaOpciones, 1)
        .Font.Name = "Arial"
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Columna B: No
    wsPlantilla.Cells(filaOpciones, 2).Value = "No"
    
    With wsPlantilla.Cells(filaOpciones, 2)
        .Font.Name = "Arial"
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Columnas C-D: Firma (combinadas)
    wsPlantilla.Range(wsPlantilla.Cells(filaOpciones, 3), wsPlantilla.Cells(filaOpciones, 4)).Merge
    wsPlantilla.Cells(filaOpciones, 3).Value = "Firma: ___________________"
    
    With wsPlantilla.Cells(filaOpciones, 3)
        .Font.Name = "Arial"
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Ajustar altura de todas las filas a 25
    wsPlantilla.Rows(filaFeedback).rowHeight = 25
    wsPlantilla.Rows(filaTexto1).rowHeight = 25
    wsPlantilla.Rows(filaTexto2).rowHeight = 25
    wsPlantilla.Rows(filaFeedback + 3).rowHeight = 25  ' Fila vacía
    wsPlantilla.Rows(filaOpciones).rowHeight = 25
    
    ' Aplicar bordes: Borde bold alrededor de toda la sección
    Dim rangoFeedbackCompleto As Range
    Set rangoFeedbackCompleto = wsPlantilla.Range(wsPlantilla.Cells(filaFeedback, 1), wsPlantilla.Cells(filaOpciones, 4))
    
    With rangoFeedbackCompleto
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium  ' Borde grueso alrededor
    End With
    
    ' Bordes inferiores en filas internas
    With wsPlantilla.Range(wsPlantilla.Cells(filaFeedback, 1), wsPlantilla.Cells(filaFeedback, 4)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With wsPlantilla.Range(wsPlantilla.Cells(filaTexto1, 1), wsPlantilla.Cells(filaTexto1, 4)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With wsPlantilla.Range(wsPlantilla.Cells(filaTexto2, 1), wsPlantilla.Cells(filaTexto2, 4)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With wsPlantilla.Range(wsPlantilla.Cells(filaFeedback + 3, 1), wsPlantilla.Cells(filaFeedback + 3, 4)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Bordes verticales internos en la fila de opciones
    With wsPlantilla.Cells(filaOpciones, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With wsPlantilla.Cells(filaOpciones, 2).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Variable para siguiente sección (usar filaOpciones como última fila de feedback)
    Dim filaFirma As Long
    filaFirma = filaOpciones  ' Actualizar para mantener compatibilidad con código posterior
    
    ' Debug.Print "[POBLACIÓN] Sección de Feedback agregada en filas " & filaFeedback & "-" & filaOpciones
    
    ' -----------------------------------------------------------
    ' SECCIÓN DE OBSERVACIONES DEL ÁREA DE PRODUCCIÓN (2 filas después de feedback)
    ' -----------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Agregando sección de Observaciones del Área de Producción..."
    
    Dim filaObservaciones As Long
    filaObservaciones = filaFirma + 2  ' 2 filas después (1 fila en blanco entre secciones)
    
    ' Título de la sección
    wsPlantilla.Range(wsPlantilla.Cells(filaObservaciones, 1), wsPlantilla.Cells(filaObservaciones, 4)).Merge
    wsPlantilla.Cells(filaObservaciones, 1).Value = "OBSERVACIONES DEL ÁREA DE PRODUCCIÓN"
    
    With wsPlantilla.Cells(filaObservaciones, 1)
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(189, 215, 238)  ' Azul claro (mismo que encabezados)
    End With
    
    ' 5 filas vacías con bordes activos debajo del título
    Dim filaInicioTexto As Long
    Dim filaFinTexto As Long
    filaInicioTexto = filaObservaciones + 1
    filaFinTexto = filaInicioTexto + 4  ' 5 filas en total (0, 1, 2, 3, 4)
    
    ' NO combinar - mantener filas separadas para que se vean los bordes internos
    
    ' Formato de cada fila individual
    Dim i As Long
    For i = filaInicioTexto To filaFinTexto
        ' Combinar columnas A-D en cada fila
        wsPlantilla.Range(wsPlantilla.Cells(i, 1), wsPlantilla.Cells(i, 4)).Merge
        
        With wsPlantilla.Cells(i, 1)
            .Font.Name = "Arial"
            .Font.Size = 10
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .WrapText = True
            .Value = ""  ' Dejar vacío
        End With
        
        ' Ajustar altura de fila
        wsPlantilla.Rows(i).rowHeight = 25
        
        ' Aplicar borde inferior a cada fila (excepto la última)
        If i < filaFinTexto Then
            With wsPlantilla.Range(wsPlantilla.Cells(i, 1), wsPlantilla.Cells(i, 4)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End If
    Next i
    
    ' Aplicar borde grueso alrededor de toda la sección de observaciones
    Dim rangoObservacionesCompleto As Range
    Set rangoObservacionesCompleto = wsPlantilla.Range(wsPlantilla.Cells(filaObservaciones, 1), wsPlantilla.Cells(filaFinTexto, 4))
    
    With rangoObservacionesCompleto
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium  ' Borde grueso alrededor
    End With
    
    ' Borde inferior del título
    With wsPlantilla.Range(wsPlantilla.Cells(filaObservaciones, 1), wsPlantilla.Cells(filaObservaciones, 4)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Debug.Print "[POBLACIÓN] Sección de Observaciones agregada en filas " & filaObservaciones & "-" & filaFinTexto
    
    ' -----------------------------------------------------------
    ' SECCIÓN DE PLAN DE TRABAJO (2 filas después de observaciones)
    ' -----------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Agregando sección de Plan de Trabajo..."
    
    Dim filaPlanTrabajo As Long
    filaPlanTrabajo = filaFinTexto + 2  ' 2 filas después (1 fila en blanco entre secciones)
    
    ' Título de la sección
    wsPlantilla.Range(wsPlantilla.Cells(filaPlanTrabajo, 1), wsPlantilla.Cells(filaPlanTrabajo, 4)).Merge
    wsPlantilla.Cells(filaPlanTrabajo, 1).Value = "PLAN DE TRABAJO (CONSIDERAR PERIODO DE SEGUIMIENTO E INCLUIR N° DE DESVÍO SI APLICA)"
    
    With wsPlantilla.Cells(filaPlanTrabajo, 1)
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(189, 215, 238)  ' Azul claro (mismo que encabezados)
    End With
    
    ' 12 filas vacías con bordes activos debajo del título
    Dim filaInicioPlan As Long
    Dim filaFinPlan As Long
    filaInicioPlan = filaPlanTrabajo + 1
    filaFinPlan = filaInicioPlan + 11  ' 12 filas en total (0-11)
    
    ' NO combinar - mantener filas separadas para que se vean los bordes internos
    
    ' Formato de cada fila individual
    Dim j As Long
    For j = filaInicioPlan To filaFinPlan
        ' Combinar columnas A-D en cada fila
        wsPlantilla.Range(wsPlantilla.Cells(j, 1), wsPlantilla.Cells(j, 4)).Merge
        
        With wsPlantilla.Cells(j, 1)
            .Font.Name = "Arial"
            .Font.Size = 10
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .WrapText = True
            .Value = ""  ' Dejar vacío
        End With
        
        ' Ajustar altura de fila
        wsPlantilla.Rows(j).rowHeight = 25
        
        ' Aplicar borde inferior a cada fila (excepto la última)
        If j < filaFinPlan Then
            With wsPlantilla.Range(wsPlantilla.Cells(j, 1), wsPlantilla.Cells(j, 4)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End If
    Next j
    
    ' Aplicar borde grueso alrededor de toda la sección de plan de trabajo
    Dim rangoPlanCompleto As Range
    Set rangoPlanCompleto = wsPlantilla.Range(wsPlantilla.Cells(filaPlanTrabajo, 1), wsPlantilla.Cells(filaFinPlan, 4))
    
    With rangoPlanCompleto
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium  ' Borde grueso alrededor
    End With
    
    ' Borde inferior del título
    With wsPlantilla.Range(wsPlantilla.Cells(filaPlanTrabajo, 1), wsPlantilla.Cells(filaPlanTrabajo, 4)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Debug.Print "[POBLACIÓN] Sección de Plan de Trabajo agregada en filas " & filaPlanTrabajo & "-" & filaFinPlan
    
    ' -----------------------------------------------------------
    ' SECCIÓN DE MUESTREO MICROBIOLÓGICO (2 filas después de plan de trabajo)
    ' AGREGADO 27/04/2026
    ' -----------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Agregando sección de Muestreo Microbiológico..."
    
    Dim filaMuestreoMicro As Long
    filaMuestreoMicro = filaFinPlan + 2  ' 2 filas después (1 fila en blanco)
    
    ' Título de la sección
    wsPlantilla.Range(wsPlantilla.Cells(filaMuestreoMicro, 1), wsPlantilla.Cells(filaMuestreoMicro, 4)).Merge
    wsPlantilla.Cells(filaMuestreoMicro, 1).Value = "MUESTREO MICROBIOLÓGICO"
    
    With wsPlantilla.Cells(filaMuestreoMicro, 1)
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(189, 215, 238)  ' Azul claro (mismo que encabezados)
    End With
    
    ' Fila con el texto de verificación
    Dim filaTextoMuestreo As Long
    filaTextoMuestreo = filaMuestreoMicro + 1
    
    wsPlantilla.Range(wsPlantilla.Cells(filaTextoMuestreo, 1), wsPlantilla.Cells(filaTextoMuestreo, 4)).Merge
    wsPlantilla.Cells(filaTextoMuestreo, 1).Value = "Muestreo microbiológico: Por incumplimiento del punto 17 se requiere seguimiento (2 verificaciones) / No requiere verificación: SI (  ) / NO (  )"
    
    With wsPlantilla.Cells(filaTextoMuestreo, 1)
        .Font.Name = "Arial"
        .Font.Size = 10
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    wsPlantilla.Rows(filaTextoMuestreo).rowHeight = 30
    
    ' Fila para N° desvío
    Dim filaDesvio As Long
    filaDesvio = filaTextoMuestreo + 1
    
    wsPlantilla.Range(wsPlantilla.Cells(filaDesvio, 1), wsPlantilla.Cells(filaDesvio, 4)).Merge
    wsPlantilla.Cells(filaDesvio, 1).Value = "N° desvío: _______________________"
    
    With wsPlantilla.Cells(filaDesvio, 1)
        .Font.Name = "Arial"
        .Font.Size = 10
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    wsPlantilla.Rows(filaDesvio).rowHeight = 25
    
    ' Aplicar borde grueso alrededor de toda la sección de muestreo microbiológico
    Dim rangoMuestreoCompleto As Range
    Set rangoMuestreoCompleto = wsPlantilla.Range(wsPlantilla.Cells(filaMuestreoMicro, 1), wsPlantilla.Cells(filaDesvio, 4))
    
    With rangoMuestreoCompleto
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium  ' Borde grueso alrededor
    End With
    
    ' Borde inferior del título
    With wsPlantilla.Range(wsPlantilla.Cells(filaMuestreoMicro, 1), wsPlantilla.Cells(filaMuestreoMicro, 4)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Debug.Print "[POBLACIÓN] Sección de Muestreo Microbiológico agregada en filas " & filaMuestreoMicro & "-" & filaDesvio
    
    ' -----------------------------------------------------------
    ' SECCIÓN DE FIRMAS DE AUTORIZACIÓN (2 filas después de muestreo microbiológico)
    ' -----------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Agregando sección de Firmas de Autorización..."
    
    Dim filaFirmasAutorizacion As Long
    filaFirmasAutorizacion = filaDesvio + 2  ' 2 filas después (1 fila en blanco)
    
    ' Fila 1: Líneas de firma (distribuidas en 3 secciones)
    ' Firma 1 (Columna A)
    wsPlantilla.Cells(filaFirmasAutorizacion, 1).Value = "______________________"
    wsPlantilla.Cells(filaFirmasAutorizacion, 1).HorizontalAlignment = xlCenter
    
    ' Firma 2 (Columnas B-C combinadas)
    wsPlantilla.Range(wsPlantilla.Cells(filaFirmasAutorizacion, 2), wsPlantilla.Cells(filaFirmasAutorizacion, 3)).Merge
    wsPlantilla.Cells(filaFirmasAutorizacion, 2).Value = "______________________"
    wsPlantilla.Cells(filaFirmasAutorizacion, 2).HorizontalAlignment = xlCenter
    
    ' Firma 3 (Columna D)
    wsPlantilla.Cells(filaFirmasAutorizacion, 4).Value = "______________________"
    wsPlantilla.Cells(filaFirmasAutorizacion, 4).HorizontalAlignment = xlCenter
    
    ' Formato de las líneas de firma
    With wsPlantilla.Range(wsPlantilla.Cells(filaFirmasAutorizacion, 1), wsPlantilla.Cells(filaFirmasAutorizacion, 4))
        .Font.Name = "Arial"
        .Font.Size = 10
        .VerticalAlignment = xlBottom
    End With
    
    ' Fila 2: Nombres de puestos
    Dim filaNombresPuestos As Long
    filaNombresPuestos = filaFirmasAutorizacion + 1
    
    ' Nombre puesto 1 (Columna A)
    wsPlantilla.Cells(filaNombresPuestos, 1).Value = "Jefe de Capacitación/Área"
    wsPlantilla.Cells(filaNombresPuestos, 1).HorizontalAlignment = xlCenter
    
    ' Nombre puesto 2 (Columnas B-C combinadas)
    wsPlantilla.Range(wsPlantilla.Cells(filaNombresPuestos, 2), wsPlantilla.Cells(filaNombresPuestos, 3)).Merge
    wsPlantilla.Cells(filaNombresPuestos, 2).Value = "Subgerente de Producción"
    wsPlantilla.Cells(filaNombresPuestos, 2).HorizontalAlignment = xlCenter
    
    ' Nombre puesto 3 (Columna D)
    wsPlantilla.Cells(filaNombresPuestos, 4).Value = "Jefe de Aseguramiento de Calidad"
    wsPlantilla.Cells(filaNombresPuestos, 4).HorizontalAlignment = xlCenter
    wsPlantilla.Cells(filaNombresPuestos, 4).WrapText = True
    
    ' Formato de los nombres de puestos
    With wsPlantilla.Range(wsPlantilla.Cells(filaNombresPuestos, 1), wsPlantilla.Cells(filaNombresPuestos, 4))
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .VerticalAlignment = xlTop
    End With
    
    ' Fila 3: Texto "Nombre/Firma/Fecha" debajo de cada puesto
    Dim filaTextoFirma As Long
    filaTextoFirma = filaNombresPuestos + 1
    
    ' Texto firma 1 (Columna A)
    wsPlantilla.Cells(filaTextoFirma, 1).Value = "Nombre/Firma/Fecha"
    wsPlantilla.Cells(filaTextoFirma, 1).HorizontalAlignment = xlCenter
    
    ' Texto firma 2 (Columnas B-C combinadas)
    wsPlantilla.Range(wsPlantilla.Cells(filaTextoFirma, 2), wsPlantilla.Cells(filaTextoFirma, 3)).Merge
    wsPlantilla.Cells(filaTextoFirma, 2).Value = "Nombre/Firma/Fecha"
    wsPlantilla.Cells(filaTextoFirma, 2).HorizontalAlignment = xlCenter
    
    ' Texto firma 3 (Columna D)
    wsPlantilla.Cells(filaTextoFirma, 4).Value = "Nombre/Firma/Fecha"
    wsPlantilla.Cells(filaTextoFirma, 4).HorizontalAlignment = xlCenter
    
    ' Formato del texto "Nombre/Firma/Fecha"
    With wsPlantilla.Range(wsPlantilla.Cells(filaTextoFirma, 1), wsPlantilla.Cells(filaTextoFirma, 4))
        .Font.Name = "Arial"
        .Font.Size = 8
        .Font.Italic = True
        .VerticalAlignment = xlTop
    End With
    
    ' Ajustar altura de filas
    wsPlantilla.Rows(filaFirmasAutorizacion).rowHeight = 30
    wsPlantilla.Rows(filaNombresPuestos).rowHeight = 35
    wsPlantilla.Rows(filaTextoFirma).rowHeight = 20
    
    ' Debug.Print "[POBLACIÓN] Sección de Firmas de Autorización agregada en filas " & filaFirmasAutorizacion & "-" & filaTextoFirma
    
    ' -----------------------------------------------------------
    ' SECCIÓN 4: FOOTER DINÁMICO (Después de las firmas de autorización)
    ' -----------------------------------------------------------
    ' Debug.Print "[POBLACIÓN] Agregando footer de emisión..."
    
    ' Calcular fila del footer (3 filas después de las firmas)
    Dim filaFooter As Long
    filaFooter = filaTextoFirma + 3
    
    ' Obtener fecha y hora actual (zona horaria America/Santiago)
    Dim fechaEmision As Date
    Dim horaEmision As String
    fechaEmision = Date
    horaEmision = Format(Now, "HH:MM:SS")
    
    ' Usuario que emitió (usar el nombre del auditor de la inspección)
    Dim usuarioEmision As String
    usuarioEmision = datosInspeccion("Auditor")
    If Len(Trim(usuarioEmision)) = 0 Then
        usuarioEmision = "No especificado"
    End If
    
    ' Construir texto del footer
    Dim textoFooter As String
    textoFooter = "Certificado emitido por: " & usuarioEmision & " | " & _
                  "Fecha: " & Format(fechaEmision, "DD-MM-YYYY") & " | " & _
                  "Hora: " & horaEmision & " (America/Santiago)"
    
    ' Escribir footer en la hoja (A-D para 4 columnas)
    wsPlantilla.Range(wsPlantilla.Cells(filaFooter, 1), wsPlantilla.Cells(filaFooter, 4)).Merge
    wsPlantilla.Cells(filaFooter, 1).Value = textoFooter
    
    ' Formato del footer
    With wsPlantilla.Cells(filaFooter, 1)
        .Font.Name = "Arial"
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = RGB(128, 128, 128)  ' Gris
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Debug.Print "[POBLACIÓN] Footer agregado en fila " & filaFooter
    ' Debug.Print "[POBLACIÓN] ¡Plantilla completamente poblada!"
End Sub

' --------------------------------------------------------------
' Poblar tabla de preguntas y respuestas
' --------------------------------------------------------------
Private Function PoblarTablaRespuestas(wsPlantilla As Worksheet, datosInspeccion As Object, filaEncabezado As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim respuestas As Collection
    Set respuestas = datosInspeccion("Respuestas")
    
    ' Debug.Print "[RESPUESTAS] Total de respuestas: " & respuestas.Count
    ' Debug.Print "[RESPUESTAS] Fila encabezado: " & filaEncabezado
    
    If respuestas.Count = 0 Then
        ' Debug.Print "[RESPUESTAS] Sin respuestas para mostrar"
        PoblarTablaRespuestas = filaEncabezado  ' Retornar solo fila de encabezados si no hay respuestas
        Exit Function
    End If
    
    ' Las tablas de preguntas y opciones están en SHEET_CHECKLIST, no en HISTORICO
    Dim wsChecklist As Worksheet
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    
    Dim tblPreguntas As ListObject
    Dim tblOpciones As ListObject
    Set tblPreguntas = wsChecklist.ListObjects(Configuration2.TABLE_PREGUNTAS)
    Set tblOpciones = wsChecklist.ListObjects(Configuration2.TABLE_OPCIONES)
    
    ' Debug.Print "[RESPUESTAS] tblPreguntas encontrada: " & (Not tblPreguntas Is Nothing)
    ' Debug.Print "[RESPUESTAS] tblOpciones encontrada: " & (Not tblOpciones Is Nothing)
    ' Debug.Print "[RESPUESTAS] Limpiando desde fila " & filaEncabezado & " antes de poblar..."
    Dim ultimaFilaExistente As Long
    ultimaFilaExistente = wsPlantilla.UsedRange.row + wsPlantilla.UsedRange.Rows.Count - 1
    
    If ultimaFilaExistente >= filaEncabezado Then
        ' Limpiar desde fila de encabezado hasta el final (columnas A-G para evitar residuos)
        wsPlantilla.Range( _
            wsPlantilla.Cells(filaEncabezado, 1), _
            wsPlantilla.Cells(ultimaFilaExistente, 7) _
        ).Clear  ' Clear elimina contenido Y formato
        ' Debug.Print "[RESPUESTAS] Limpiadas filas " & filaEncabezado & "-" & ultimaFilaExistente
    Else
        ' Debug.Print "[RESPUESTAS] No hay filas previas para limpiar"
    End If
    
    ' ENCABEZADOS DE TABLA DE PREGUNTAS (dinámica según parámetro)
    wsPlantilla.Cells(filaEncabezado, 1).Value = "N° DE PREGUNTA"   ' Columna A
    wsPlantilla.Cells(filaEncabezado, 2).Value = "PREGUNTA"          ' Columna B
    wsPlantilla.Cells(filaEncabezado, 3).Value = "RESPUESTA"         ' Columna C
    wsPlantilla.Cells(filaEncabezado, 4).Value = "OBSERVACIONES"     ' Columna D
    
    ' Formato encabezados
    With wsPlantilla.Range(wsPlantilla.Cells(filaEncabezado, 1), wsPlantilla.Cells(filaEncabezado, 4))
        .Font.Bold = True
        .Font.Size = 10
        .Font.Name = "Arial"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(189, 215, 238)  ' Azul claro
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With
    
    ' Fila de inicio para datos = encabezado + 1
    Dim filaInicio As Long
    filaInicio = filaEncabezado + 1
    
    Dim i As Long
    Dim resp As Object
    Dim fila As Long
    
    For i = 1 To respuestas.Count
        Set resp = respuestas(i)
        fila = filaInicio + (i - 1)
        
        ' Debug.Print "[RESPUESTAS] Fila " & fila & " - IDPregunta: " & resp("IDPregunta") & " - IDOpcion: " & resp("IDOpcion")
        
        ' Col A: Número secuencial
        wsPlantilla.Cells(fila, 1).Value = i
        
        ' Col B: Texto de pregunta (buscar por ID string)
        wsPlantilla.Cells(fila, 2).Value = ObtenerTextoPregunta(CStr(resp("IDPregunta")), tblPreguntas)
        
        ' Col C: Texto de opción elegida (buscar por ID string)
        Dim textoOpcion As String
        textoOpcion = ObtenerTextoOpcion(CStr(resp("IDOpcion")), tblOpciones)
        wsPlantilla.Cells(fila, 3).Value = textoOpcion
        
        ' Col D: Observación
        wsPlantilla.Cells(fila, 4).Value = resp("Observacion")
        
        ' Formato de fila
        With wsPlantilla.Range(wsPlantilla.Cells(fila, 1), wsPlantilla.Cells(fila, 4))
            .Font.Name = "Arial"
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        ' Auto-ajustar altura de fila según contenido
        wsPlantilla.Rows(fila).AutoFit
        
        ' -----------------------------------------------------------
        ' PASO 3 MVP: Resaltar incumplimientos con fondo rojo
        ' -----------------------------------------------------------
        If EsRespuestaIncumplimiento(textoOpcion) Then
            wsPlantilla.Range(wsPlantilla.Cells(fila, 1), wsPlantilla.Cells(fila, 4)).Interior.Color = RGB(253, 223, 223)  ' Rojo claro
            ' Debug.Print "[RESPUESTAS] ? INCUMPLIMIENTO detectado: " & textoOpcion
        End If
        
        ' Debug.Print "[RESPUESTAS] OK -> " & wsPlantilla.Cells(fila, 2).Value
    Next i
    
    ' Debug.Print "[RESPUESTAS] Todas las respuestas pobladas"
    
    ' Retornar la última fila utilizada
    PoblarTablaRespuestas = fila
    Exit Function
    
ErrorHandler:
    ' Debug.Print "[ERROR RESPUESTAS] Fila " & fila & ": " & Err.Description
    PoblarTablaRespuestas = filaEncabezado  ' Retornar fila de encabezados en caso de error
End Function

' --------------------------------------------------------------
' Obtener texto de pregunta por UUID (tablas en SHEET_CHECKLIST)
' Columnas tblPreguntas: [1]ID Plantilla [2]ID Pregunta [5]Texto
' --------------------------------------------------------------
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

' --------------------------------------------------------------
' Obtener texto de opción por UUID (tablas en SHEET_CHECKLIST)
' Columnas tblOpcionesDeRespuesta: [1]ID Opcion [4]Opción texto
' --------------------------------------------------------------
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

' --------------------------------------------------------------
' Detectar si una respuesta indica incumplimiento (Paso 3 MVP)
' Busca palabras clave que indican no conformidad
' --------------------------------------------------------------
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

' --------------------------------------------------------------
' Obtener nombre por iniciales (tblPersonal en SHEET_PERSONAL)
' tblPersonal: [1]Iniciales [2]Planta [3-13]Puestos Si/No [14]Activo
' No hay columna Nombre ? se devuelven las iniciales como identificador
' --------------------------------------------------------------
' --------------------------------------------------------------
' Evaluar resultado de Auditoría de Procesos según regla de negocio
' Reglas (en orden de prioridad):
'   1. Si hay 2 o más "No Cumple" de criticidad "Crítica" ? "No Cumple"
'   2. Si hay 1 "No Cumple" de criticidad "Crítica" Y 2 o más "No Cumple"
'      de criticidad "Mayor" ? "No Cumple"
'   3. Si hay 4 o más "No Cumple" de criticidad "Mayor" ? "No Cumple"
'   4. En cualquier otro caso ? "Cumple"
' --------------------------------------------------------------
Private Function EvaluarResultadoAP(ByVal criticaNoCumple As Long, ByVal mayorNoCumple As Long) As String
    On Error Resume Next
    
    ' Regla 1: 2 o más "No Cumple" de criticidad "Crítica"
    If criticaNoCumple >= 2 Then
        EvaluarResultadoAP = "No Cumple"
        Exit Function
    End If
    
    ' Regla 2: 1 "No Cumple" de criticidad "Crítica" Y
    '          2 o más "No Cumple" de criticidad "Mayor"
    If criticaNoCumple >= 1 And mayorNoCumple >= 2 Then
        EvaluarResultadoAP = "No Cumple"
        Exit Function
    End If
    
    ' Regla 3: 4 o más "No Cumple" de criticidad "Mayor"
    If mayorNoCumple >= 4 Then
        EvaluarResultadoAP = "No Cumple"
        Exit Function
    End If
    
    ' Regla 4: En cualquier otro caso ? "Cumple"
    EvaluarResultadoAP = "Cumple"
    
    On Error GoTo 0
End Function

' --------------------------------------------------------------
' Aplicar formato condicional a celda de fecha de vencimiento
' FASE 7 - 23/04/2026 | Actualizado 24/04/2026
' Propósito: Aplicar color verde si la fecha está vigente (futura),
'            o color rojo si ya está vencida (pasada o igual a hoy).
' Parámetros:
'   ws: Worksheet de la plantilla
'   fila: Número de fila de la celda
'   columna: Número de columna de la celda
'   fechaStr: Fecha como string (dd-mm-yyyy o dd/mm/yyyy)
' --------------------------------------------------------------
Private Sub AplicarFormatoFechaVencimiento(ws As Worksheet, fila As Long, columna As Long, fechaStr As String)
    On Error Resume Next
    
    ' Limpiar espacios
    fechaStr = Trim(fechaStr)
    
    ' Si la fecha es "-" o Configuration2.VALOR_NO_APLICA, no aplicar formato
    If fechaStr = "-" Or Configuration2.EsValorNoAplica(fechaStr) Then
        Exit Sub
    End If
    
    ' Intentar convertir a fecha
    Dim fechaVenc As Date
    Dim esValida As Boolean
    esValida = False
    
    ' Probar conversión directa
    On Error Resume Next
    fechaVenc = CDate(fechaStr)
    If Err.Number = 0 Then esValida = True
    Err.Clear
    On Error GoTo 0
    
    ' Si no es válida, salir
    If Not esValida Then
        ' Debug.Print "[FormatoFecha] Fecha no válida: " & fechaStr
        Exit Sub
    End If
    
    ' Comparar con la fecha actual
    Dim hoy As Date
    hoy = Date
    
    ' Debug.Print "[FormatoFecha] Fecha: " & fechaStr & " | Hoy: " & Format(hoy, "dd-mm-yyyy") & " | Vigente: " & (fechaVenc > hoy)
    
    ' Aplicar formato según criterio simple: vigente o vencida
    With ws.Cells(fila, columna)
        .Font.Name = "Arial"
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        If fechaVenc > hoy Then
            ' VIGENTE (fecha futura): Fondo verde
            .Interior.Color = RGB(212, 244, 230)  ' Verde claro
            .Font.Color = RGB(39, 174, 96)        ' Verde oscuro
            ' Debug.Print "[FormatoFecha] Aplicado formato VERDE (vigente)"
        Else
            ' VENCIDA (fecha pasada o hoy mismo): Fondo rojo
            .Interior.Color = RGB(253, 223, 223)  ' Rojo claro
            .Font.Color = RGB(203, 67, 53)        ' Rojo oscuro
            .Font.Bold = True
            ' Debug.Print "[FormatoFecha] Aplicado formato ROJO (vencida)"
        End If
    End With
    
    On Error GoTo 0
End Sub

Private Function ObtenerNombrePersonal(iniciales As String) As String
    ' tblPersonal no tiene columna Nombre Completo — devolver iniciales
    ObtenerNombrePersonal = Trim(iniciales)
End Function

' --------------------------------------------------------------
' Obtener puestos activos por iniciales (tblPersonal en SHEET_PERSONAL)
' Columnas 3-13 son puestos con valor "Si"/"No"
' --------------------------------------------------------------
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

' --------------------------------------------------------------
' Obtener frecuencia de plantilla (en meses) desde tblPlantillas
' Paso 4 MVP - Para calcular fecha de validez del certificado
' --------------------------------------------------------------
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

' --------------------------------------------------------------
' NUEVA FUNCIÓN: Obtener nombre de categoría desde tblCategoriasRPN
' --------------------------------------------------------------
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

' --------------------------------------------------------------
' Limpiar plantilla después de generar PDF
' --------------------------------------------------------------
Private Sub LimpiarPlantillaCertificado(wsPlantilla As Worksheet)
    On Error Resume Next
    
    ' Solo limpiar celdas de DATOS — nunca los labels, títulos ni encabezados estáticos
    ' que fueron definidos por PlantillaCertificadoSetup y deben persistir entre ejecuciones
    
    ' -- NUEVO: Bloque Categoría ---------------------------------
    wsPlantilla.Range("A6:D8").ClearContents  ' Limpiar completamente
    wsPlantilla.Range("A6:D8").Interior.Color = RGB(255, 255, 255)  ' Blanco
    
    ' -- Sección 1: Datos de inspección (desplazadas +3) ---------
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
    
    ' -- Sección 2: Resultados generales ------------------
    wsPlantilla.Cells(20, 3).ClearContents  ' C20 - TA puntaje
    wsPlantilla.Cells(20, 5).ClearContents  ' E20 - TA máximos
    wsPlantilla.Cells(21, 3).ClearContents  ' C21 - TA no aplica
    wsPlantilla.Cells(22, 3).ClearContents  ' C22 - Porcentaje
    
    ' -- Sección 3: Calificaciones Vestuario/Operador (Filas 23-24) -----
    ' IMPORTANTE: NO limpiar C23 ni C24 porque contienen el texto estático "Fecha de vencimiento: "
    wsPlantilla.Cells(23, 2).ClearContents  ' B23 - Calificación Vestuario
    wsPlantilla.Cells(23, 4).ClearContents  ' D23 - Fecha Venc Vestuario
    wsPlantilla.Cells(24, 2).ClearContents  ' B24 - Calificación Operador
    wsPlantilla.Cells(24, 4).ClearContents  ' D24 - Fecha Venc Operador
    
    ' -- Sección 4: Auditoría de Procesos (Filas 26-28) -----
    wsPlantilla.Cells(26, 2).ClearContents  ' B26 - AP Crítica No Cumple
    wsPlantilla.Cells(27, 2).ClearContents  ' B27 - AP Mayor No Cumple
    wsPlantilla.Cells(28, 2).ClearContents  ' B28 - AP Menor No Cumple
    wsPlantilla.Range("B29:D31").ClearContents  ' Resultado AP
    
    ' -- Sección 5: Resultados técnicos (Filas 33-38) -----
    wsPlantilla.Cells(33, 2).ClearContents  ' B33 - TA Máximos
    wsPlantilla.Cells(34, 2).ClearContents  ' B34 - TA Puntaje
    wsPlantilla.Cells(35, 2).ClearContents  ' B35 - TA No Aplica
    wsPlantilla.Cells(36, 2).ClearContents  ' B36 - % TA
    wsPlantilla.Cells(37, 2).ClearContents  ' B37 - % TA Anterior
    wsPlantilla.Cells(38, 2).ClearContents  ' B38 - RPN Total
    wsPlantilla.Cells(36, 4).ClearContents  ' D36 - % Recuperación
    wsPlantilla.Cells(37, 4).ClearContents  ' D37 - % OOL
    
    ' -- Sección 6: Filas de respuestas (fila 40 en adelante) -----
    Dim ultimaFila As Long
    ultimaFila = wsPlantilla.UsedRange.row + wsPlantilla.UsedRange.Rows.Count - 1
    If ultimaFila >= 40 Then
        wsPlantilla.Range( _
            wsPlantilla.Cells(40, 2), _
            wsPlantilla.Cells(ultimaFila, 7) _
        ).ClearContents
    End If
    
    ' -- Limpiar configuración de números de página --------------
    With wsPlantilla.PageSetup
        .CenterFooter = ""
    End With
    
    On Error GoTo 0
End Sub

' --------------------------------------------------------------
' Generar nombre inteligente para archivo PDF (Paso 5 MVP)
' Formato: CERTIFICADO_[PUESTO]_[INICIALES]_[FECHA]_CAT[N]_[ESTADO].pdf
' --------------------------------------------------------------
Private Function GenerarNombreArchivoPDF(datosInspeccion As Object) As String
    On Error GoTo ErrorHandler
    
    Dim nombreArchivo As String
    
    ' Obtener componentes del nombre
    Dim iniciales As String
    Dim fechaStr As String
    Dim puesto As String
    Dim categoria As Long
    Dim estadoCorto As String
    
    ' Iniciales (siempre string)
    iniciales = Trim(CStr(datosInspeccion("Iniciales")))
    
    ' Fecha: Convertir a Date antes de formatear (manejo robusto de tipos)
    Dim fechaValue As Variant
    fechaValue = datosInspeccion("FechaInspeccion")
    
    ' Debug.Print "[ARCHIVO] FechaInspeccion - Tipo: " & TypeName(fechaValue) & " | VarType: " & VarType(fechaValue) & " | Valor: [" & fechaValue & "]"
    
    If IsEmpty(fechaValue) Or IsNull(fechaValue) Then
        fechaStr = Format(Date, "YYYY-MM-DD")  ' Usar fecha actual si no hay valor
        ' Debug.Print "[ARCHIVO] Fecha vacía, usando fecha actual"
    ElseIf IsDate(fechaValue) Then
        ' Puede ser String "dd-mm-yyyy", Date, o número serial de Excel
        If VarType(fechaValue) = vbString Then
            ' String ? Convertir a Date
            fechaStr = Format(CDate(fechaValue), "YYYY-MM-DD")
        Else
            ' Date o número serial ? Formatear directamente
            fechaStr = Format(fechaValue, "YYYY-MM-DD")
        End If
    Else
        ' Si no es fecha válida, usar fecha actual
        fechaStr = Format(Date, "YYYY-MM-DD")
        ' Debug.Print "[ARCHIVO] Fecha no válida [" & fechaValue & "], usando fecha actual"
    End If
    
    ' Categoría: Convertir a Long con manejo de errores
    Dim categoriaValue As Variant
    categoriaValue = datosInspeccion("Categoria")
    
    ' Debug.Print "[ARCHIVO] Categoria - Tipo: " & TypeName(categoriaValue) & " | VarType: " & VarType(categoriaValue) & " | Valor: [" & categoriaValue & "]"
    
    If IsEmpty(categoriaValue) Or IsNull(categoriaValue) Then
        categoria = 0
        ' Debug.Print "[ARCHIVO] Categoría vacía, usando 0"
    ElseIf IsNumeric(categoriaValue) Then
        categoria = CLng(categoriaValue)
    Else
        categoria = 0
        ' Debug.Print "[ARCHIVO] Categoría no numérica [" & categoriaValue & "], usando 0"
    End If
    
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
    
    ' Debug.Print "[ARCHIVO] Componentes del nombre:"
    ' Debug.Print "[ARCHIVO]   - Prefijo: " & Configuration2.PDF_PREFIJO_NOMBRE
    ' Debug.Print "[ARCHIVO]   - Puesto: " & puesto
    ' Debug.Print "[ARCHIVO]   - Iniciales: " & iniciales
    ' Debug.Print "[ARCHIVO]   - Fecha: " & fechaStr
    ' Debug.Print "[ARCHIVO]   - Categoría: " & categoria
    ' Debug.Print "[ARCHIVO]   - Estado: " & estadoCorto
    ' Debug.Print "[ARCHIVO] Nombre generado: " & nombreArchivo
    
    GenerarNombreArchivoPDF = nombreArchivo
    Exit Function
    
ErrorHandler:
    ' Debug.Print "[ERROR] GenerarNombreArchivoPDF: " & Err.Description
    ' Generar nombre fallback
    GenerarNombreArchivoPDF = "CERTIFICADO_" & Format(Now, "YYYY-MM-DD_HHMMSS") & ".pdf"
End Function

' --------------------------------------------------------------
' Obtener ruta correcta del Desktop (robusto para múltiples idiomas y OneDrive)
' --------------------------------------------------------------
Private Function ObtenerRutaDesktop() As String
    On Error GoTo ErrorHandler
    
    ' Usar la API de Windows Shell para obtener la ruta correcta del Desktop
    ' Esto funciona incluso con OneDrive y en cualquier idioma
    Dim wshShell As Object
    Set wshShell = CreateObject("WScript.Shell")
    
    Dim desktopPath As String
    desktopPath = wshShell.SpecialFolders("Desktop")
    
    ' Debug.Print "[DEBUG] ObtenerRutaDesktop - Ruta obtenida desde Shell.SpecialFolders: " & desktopPath
    
    ' Verificar que la ruta existe
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(desktopPath) Then
        ' Debug.Print "[DEBUG] ObtenerRutaDesktop - Ruta verificada: " & desktopPath
        ObtenerRutaDesktop = desktopPath
        Exit Function
    End If
    
    ' Si falla, intentar rutas alternativas
    ' Debug.Print "[ERROR] ObtenerRutaDesktop - Shell.SpecialFolders falló, intentando alternativas..."
    
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
        ' Debug.Print "[DEBUG] ObtenerRutaDesktop - Probando: " & ruta
        If fso.FolderExists(ruta) Then
            ' Debug.Print "[DEBUG] ObtenerRutaDesktop - Ruta encontrada: " & ruta
            ObtenerRutaDesktop = ruta
            Exit Function
        End If
    Next ruta
    
    ' Si nada funciona, usar carpeta del workbook
    ' Debug.Print "[ERROR] ObtenerRutaDesktop - Usando carpeta del libro: " & ThisWorkbook.Path
    ObtenerRutaDesktop = ThisWorkbook.Path
    Exit Function
    
ErrorHandler:
    ' Debug.Print "[ERROR] ObtenerRutaDesktop - Excepción: " & Err.Description
    ObtenerRutaDesktop = ThisWorkbook.Path
End Function

' --------------------------------------------------------------
' FUNCIÓN SIMPLE: GenerarPDFDesdeSeleccion
' Propósito: Genera PDF automáticamente desde la fila seleccionada
'            en la tabla de inspecciones (útil para botones)
' --------------------------------------------------------------
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
    filaActiva = ActiveCell.row - tbl.DataBodyRange.row + 1
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