' ══════════════════════════════════════════════════════════════
' Módulo: PlantillaCertificadoSetup
' Descripción: Módulo TEMPORAL para crear y formatear la hoja
'              "Plantilla Certificado" con todo el diseño necesario
'              para generación de certificados PDF.
' Fecha creación: 15/04/2026
' Uso: Ejecutar InicializarPlantillaCertificado() UNA SOLA VEZ
' Dependencias: Configuration2
' NOTA: Este módulo puede eliminarse después de ejecutarlo.
' ══════════════════════════════════════════════════════════════
Option Explicit

' ══════════════════════════════════════════════════════════════
' FUNCIÓN PRINCIPAL - EJECUTAR UNA SOLA VEZ
' ══════════════════════════════════════════════════════════════
Public Sub InicializarPlantillaCertificado()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim respuesta As VbMsgBoxResult
    
    ' PRIMERO: Desproteger libro si está protegido (antes de cualquier operación)
    Dim libroProtegido As Boolean
    libroProtegido = ThisWorkbook.ProtectStructure
    If libroProtegido Then
        ThisWorkbook.Unprotect Configuration2.APP_PASSWORD
    End If
    
    ' Verificar si la hoja ya existe
    Dim hojaExiste As Boolean
    hojaExiste = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Configuration2.SHEET_PLANTILLA_CERTIFICADO)
    If Not ws Is Nothing Then hojaExiste = True
    On Error GoTo ErrorHandler
    
    If hojaExiste Then
        respuesta = MsgBox("La hoja '" & Configuration2.SHEET_PLANTILLA_CERTIFICADO & "' ya existe." & vbCrLf & vbCrLf & _
                          "¿Desea eliminarla y recrearla desde cero?" & vbCrLf & _
                          "(Se perderán todos los cambios manuales)", _
                          vbYesNo + vbQuestion, "Confirmar recreación")
        
        If respuesta = vbNo Then
            ' Reproteger libro antes de salir
            If libroProtegido Then
                ThisWorkbook.Protect Configuration2.APP_PASSWORD, Structure:=True
            End If
            MsgBox "Operación cancelada.", vbInformation
            Exit Sub
        End If
        
        ' Hacer visible antes de eliminar (puede estar muy oculta)
        ws.Visible = xlSheetVisible
        
        ' Eliminar hoja existente
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Crear nueva hoja
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = Configuration2.SHEET_PLANTILLA_CERTIFICADO
    
    Application.ScreenUpdating = False
    
    ' PASO 1: Configuración básica de página
    Call ConfigurarPaginaCertificado(ws)
    
    ' PASO 2: Configurar anchos de columna
    Call ConfigurarAnchoColumnas(ws)
    
    ' PASO 3: Diseñar encabezado
    Call DisenarEncabezado(ws)
    
    ' PASO 4: Diseñar sección Datos de Inspección
    Call DisenarSeccionDatos(ws)
    
    ' PASO 5: Diseñar sección Resultados Generales
    Call DisenarSeccionResultados(ws)
    
    ' PASO 6: DISEÑAR ENCABEZADO DE TABLA DE PREGUNTAS
    ' NOTA: Ya NO se crea aquí - se crea dinámicamente en CertificadoPDFGenerator
    ' Call DisenarEncabezadoTabla(ws)
    
    ' PASO 7: DISEÑAR PLACEHOLDERS OBSERVACIONES Y FIRMAS
    ' NOTA: Ya NO se crean aquí - se crean dinámicamente en CertificadoPDFGenerator
    ' Call DisenarSeccionObservacionesFirmas(ws)
    
    ' PASO 8: Definir rangos nombrados
    Call DefinirRangosNombrados(ws)
    
    ' PASO 9: Ocultar hoja (muy oculta)
    ws.Visible = xlSheetVeryHidden
    
    ' Reproteger libro si estaba protegido
    If libroProtegido Then
        ThisWorkbook.Protect Configuration2.APP_PASSWORD, Structure:=True
    End If
    
    Application.ScreenUpdating = True
    
    MsgBox "Plantilla de certificado creada exitosamente." & vbCrLf & vbCrLf & _
           "Hoja: " & Configuration2.SHEET_PLANTILLA_CERTIFICADO & vbCrLf & _
           "Estado: Muy oculta (xlSheetVeryHidden)" & vbCrLf & vbCrLf & _
           "Rangos nombrados: 20+ definidos", _
           vbInformation, "Setup Completado - Fase 1"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    If libroProtegido Then ThisWorkbook.Protect Configuration2.APP_PASSWORD, Structure:=True
    MsgBox "Error al crear plantilla de certificado:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Error en Setup"
End Sub

' ══════════════════════════════════════════════════════════════
' PASO 1: CONFIGURAR PÁGINA PARA EXPORTACIÓN PDF
' ══════════════════════════════════════════════════════════════
Private Sub ConfigurarPaginaCertificado(ws As Worksheet)
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False  ' Permitir múltiples páginas si necesario
        .PrintGridlines = False
        .PrintHeadings = False
    End With
End Sub

' ══════════════════════════════════════════════════════════════
' PASO 2: CONFIGURAR ANCHOS DE COLUMNA
' ══════════════════════════════════════════════════════════════
Private Sub ConfigurarAnchoColumnas(ws As Worksheet)
    With ws
        .Columns("A:A").ColumnWidth = 3      ' Margen izquierdo
        .Columns("B:B").ColumnWidth = 12     ' Etiquetas
        .Columns("C:C").ColumnWidth = 20     ' Valores
        .Columns("D:D").ColumnWidth = 12     ' Etiquetas
        .Columns("E:E").ColumnWidth = 20     ' Valores
        .Columns("F:F").ColumnWidth = 12     ' Etiquetas extras
        .Columns("G:G").ColumnWidth = 20     ' Valores extras
        .Columns("H:H").ColumnWidth = 3      ' Margen derecho
    End With
End Sub

' ══════════════════════════════════════════════════════════════
' PASO 3: DISEÑAR ENCABEZADO (Filas 1-4)
' ══════════════════════════════════════════════════════════════
Private Sub DisenarEncabezado(ws As Worksheet)
    With ws
        ' Espacio para logo (A1:B4)
        .Range("A1:B4").Merge
        .Range("A1").Value = "LOGO"
        .Range("A1").Font.Name = "Arial"
        .Range("A1").Font.Size = 8
        .Range("A1").Font.Color = RGB(128, 128, 128)
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").VerticalAlignment = xlCenter
        
        ' Título principal (C1:G2)
        .Range("C1:G2").Merge
        .Range("C1").Value = "CERTIFICADO DE INSPECCIÓN"
        .Range("C1").Font.Name = "Arial"
        .Range("C1").Font.Size = 16
        .Range("C1").Font.Bold = True
        .Range("C1").Font.Color = RGB(0, 0, 0)  ' Negro explícito
        .Range("C1").HorizontalAlignment = xlCenter
        .Range("C1").VerticalAlignment = xlCenter
        
        ' Subtítulo (C3:G3)
        .Range("C3:G3").Merge
        .Range("C3").Value = "TÉCNICA ASÉPTICA"
        .Range("C3").Font.Name = "Arial"
        .Range("C3").Font.Size = 14
        .Range("C3").Font.Bold = True
        .Range("C3").Font.Color = RGB(0, 0, 0)  ' Negro explícito
        .Range("C3").HorizontalAlignment = xlCenter
        .Range("C3").VerticalAlignment = xlCenter
        
        ' Proyecto (C4:G4)
        .Range("C4:G4").Merge
        .Range("C4").Value = "PROYECTO: TH-HC-001"
        .Range("C4").Font.Name = "Arial"
        .Range("C4").Font.Size = 10
        .Range("C4").HorizontalAlignment = xlCenter
        .Range("C4").VerticalAlignment = xlCenter
        
        ' Fondo gris claro y bordes
        .Range("A1:G4").Interior.Color = RGB(242, 242, 242)
        .Range("A1:G4").BorderAround LineStyle:=xlContinuous, Weight:=xlThick
        
        ' ═══════════════════════════════════════════════════════════
        ' NUEVO: BLOQUE CATEGORÍA (Filas 6-8)
        ' ═══════════════════════════════════════════════════════════
        ' Combinar celdas A6:D8 para bloque categoría (4 columnas según requisito)
        .Range("A6:D8").Merge
        .Range("A6").Value = ""  ' Dejar vacío, se llenará al generar el certificado
        .Range("A6").Font.Name = "Arial"
        .Range("A6").Font.Size = 18
        .Range("A6").Font.Bold = True
        .Range("A6").Font.Color = RGB(0, 0, 0)
        .Range("A6").HorizontalAlignment = xlCenter
        .Range("A6").VerticalAlignment = xlCenter
        .Range("A6:D8").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        ' Color por defecto (se cambiará dinámicamente)
        .Range("A6:D8").Interior.Color = RGB(255, 255, 255)  ' Blanco por defecto
    End With
End Sub

' ══════════════════════════════════════════════════════════════
' PASO 4: DISEÑAR SECCIÓN DATOS DE INSPECCIÓN (Filas 9-17)
' NOTA: Filas desplazadas +3 por bloque categoría
' ══════════════════════════════════════════════════════════════
Private Sub DisenarSeccionDatos(ws As Worksheet)
    With ws
        ' Título de sección (fila 9)
        .Range("A9:G9").Merge
        .Range("A9").Value = "DATOS DE INSPECCIÓN"
        .Range("A9").Font.Name = "Arial"
        .Range("A9").Font.Size = 11
        .Range("A9").Font.Bold = True
        .Range("A9").Font.Color = RGB(0, 0, 0)  ' Negro explícito
        .Range("A9").HorizontalAlignment = xlCenter
        .Range("A9").Interior.Color = RGB(214, 234, 248)  ' Azul claro
        
        ' Fila 10: Fecha y hora
        .Range("B10").Value = "Fecha inspección:"
        .Range("C10").Value = "[FECHA]"
        .Range("D10").Value = "Hora:"
        .Range("E10").Value = "[INICIO]"
        .Range("F10").Value = "-"
        .Range("G10").Value = "[FIN]"
        
        ' Fila 11: Evaluado
        .Range("B11").Value = "Evaluado:"
        .Range("C11:E11").Merge
        .Range("C11").Value = "[NOMBRE COMPLETO]"
        .Range("F11").Value = "Iniciales:"
        .Range("G11").Value = "[ABC]"
        
        ' Fila 12: Puesto y Planta
        .Range("B12").Value = "Puesto:"
        .Range("C12:D12").Merge
        .Range("C12").Value = "[PUESTO]"
        .Range("E12").Value = "Planta:"
        .Range("F12:G12").Merge
        .Range("F12").Value = "[PLANTA]"
        
        ' Fila 13: Área y Línea
        .Range("B13").Value = "Área:"
        .Range("C13:D13").Merge
        .Range("C13").Value = "[ÁREA]"
        .Range("E13").Value = "Línea:"
        .Range("F13:G13").Merge
        .Range("F13").Value = "[LÍNEA]"
        
        ' Fila 14: Lugar auditoría
        .Range("B14").Value = "Lugar auditoría:"
        .Range("C14:G14").Merge
        .Range("C14").Value = "[LUGAR]"
        
        ' Fila 15: Evaluador
        .Range("B15").Value = "Evaluador:"
        .Range("C15:G15").Merge
        .Range("C15").Value = "[EVALUADOR]"
        
        ' Fila 16: Personal línea
        .Range("B16").Value = "Personal línea:"
        .Range("C16").Value = "AY1:"
        .Range("D16").Value = "[---]"
        .Range("E16").Value = "AY2:"
        .Range("F16").Value = "[---]"
        .Range("G16").Value = "OP: [---]"
        
        ' Fila 17: (espacio en blanco)
        .Range("A17:G17").Value = ""
        
        ' Aplicar formato de fuente a toda la sección
        .Range("B10:G16").Font.Name = "Arial"
        .Range("B10:G16").Font.Size = 9
        .Range("B10:G16").Font.Color = RGB(0, 0, 0)  ' Negro explícito
        
        ' Labels en negrita
        .Range("B10,D10,F10,B11,F11,B12,E12,B13,E13,B14,B15,B16,C16,E16").Font.Bold = True
        
        ' Bordes finos
        .Range("A9:G16").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        .Range("A9:G16").Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range("A9:G16").Borders(xlInsideHorizontal).Weight = xlThin
    End With
End Sub

' ══════════════════════════════════════════════════════════════
' PASO 5: DISEÑAR SECCIÓN RESULTADOS GENERALES (Filas 19-22)
' NOTA: Filas desplazadas +3 por bloque categoría
' ACTUALIZADO: Filas 23-24 ahora son calificaciones (FASE 7)
' ══════════════════════════════════════════════════════════════
Private Sub DisenarSeccionResultados(ws As Worksheet)
    With ws
        ' Título de sección (fila 19)
        .Range("A19:G19").Merge
        .Range("A19").Value = "RESULTADOS GENERALES"
        .Range("A19").Font.Name = "Arial"
        .Range("A19").Font.Size = 11
        .Range("A19").Font.Bold = True
        .Range("A19").Font.Color = RGB(0, 0, 0)  ' Negro explícito
        .Range("A19").HorizontalAlignment = xlCenter
        .Range("A19").Interior.Color = RGB(213, 244, 230)  ' Verde claro
        
        ' Fila 20: TA Puntaje obtenido
        .Range("B20").Value = "TA Puntaje obtenido:"
        .Range("C20").Value = "[X]"
        .Range("D20").Value = "/ máximos:"
        .Range("E20").Value = "[Y]"
        
        ' Fila 21: TA Puntos no aplica
        .Range("B21").Value = "TA Puntos no aplica:"
        .Range("C21").Value = "[Z]"
        
        ' Fila 22: TA Porcentaje
        .Range("B22").Value = "TA Porcentaje:"
        .Range("C22").Value = "[XX.XX]"
        .Range("D22").Value = "%"
        
        ' --- FASE 7 (23/04/2026): CALIFICACIONES VESTUARIO/OPERADOR ---
        ' Fila 23: Calificación Vestuario
        .Range("B23").Value = "Calificación Vestuario:"
        .Range("C23").Value = "Fecha de vencimiento:"
        .Range("D23").Value = "[DD-MM-YYYY]"
        
        ' Fila 24: Calificación Operador
        .Range("B24").Value = "Calificación Operador:"
        .Range("C24").Value = "Fecha de vencimiento:"
        .Range("D24").Value = "[DD-MM-YYYY]"
        
        ' Aplicar formato de fuente
        .Range("B20:G24").Font.Name = "Arial"
        .Range("B20:G24").Font.Size = 9
        .Range("B20:G24").Font.Color = RGB(0, 0, 0)  ' Negro explícito
        
        ' Labels en negrita
        .Range("B20,D20,B21,B22,D22,B23,C23,B24,C24").Font.Bold = True
        
        ' Bordes
        .Range("A19:G24").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        .Range("A19:G24").Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range("A19:G24").Borders(xlInsideHorizontal).Weight = xlThin
    End With
End Sub

' ══════════════════════════════════════════════════════════════
' PASO 6: DISEÑAR ENCABEZADO DE TABLA DE PREGUNTAS (Fila 26)
' NOTA: Fila desplazada +3 por bloque categoría
' ══════════════════════════════════════════════════════════════
Private Sub DisenarEncabezadoTabla(ws As Worksheet)
    With ws
        ' Encabezados de columna
        .Range("B26").Value = "#"
        .Range("C26:D26").Merge
        .Range("C26").Value = "PREGUNTA"
        .Range("E26").Value = "RESPUESTA"
        .Range("F26:G26").Merge
        .Range("F26").Value = "OBSERVACIÓN"
        
        ' Formato de encabezado
        .Range("B26:G26").Font.Name = "Arial"
        .Range("B26:G26").Font.Size = 8
        .Range("B26:G26").Font.Bold = True
        .Range("B26:G26").Font.Color = RGB(0, 0, 0)  ' Negro explícito
        .Range("B26:G26").HorizontalAlignment = xlCenter
        .Range("B26:G26").VerticalAlignment = xlCenter
        .Range("B26:G26").Interior.Color = RGB(204, 204, 204)  ' Gris oscuro
        
        ' Bordes gruesos
        .Range("B26:G26").BorderAround LineStyle:=xlContinuous, Weight:=xlThick
        .Range("B26:G26").Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range("B26:G26").Borders(xlInsideVertical).Weight = xlMedium
        
        ' Nota: Las filas 24+ se llenarán dinámicamente por el generador
    End With
End Sub

' ══════════════════════════════════════════════════════════════
' PASO 7: DISEÑAR PLACEHOLDERS OBSERVACIONES Y FIRMAS
' ══════════════════════════════════════════════════════════════
Private Sub DisenarSeccionObservacionesFirmas(ws As Worksheet)
    With ws
        ' Estas secciones se crearán dinámicamente después de las preguntas
        ' Aquí solo creamos un comentario para referencia
        
        ' Fila 55: Placeholder para observaciones (se moverá dinámicamente si es necesario)
        .Range("B55").Value = "[SECCIÓN OBSERVACIONES - Se creará dinámicamente si es necesario]"
        .Range("B55").Font.Name = "Arial"
        .Range("B55").Font.Size = 8
        .Range("B55").Font.Color = RGB(150, 150, 150)
        .Range("B55").Font.Italic = True
        
        ' Fila 60: Placeholder para firmas (se moverá dinámicamente si es necesario)
        .Range("B60").Value = "[SECCIÓN FIRMAS - Se creará dinámicamente si es necesario]"
        .Range("B60").Font.Name = "Arial"
        .Range("B60").Font.Size = 8
        .Range("B60").Font.Color = RGB(150, 150, 150)
        .Range("B60").Font.Italic = True
    End With
End Sub

' ══════════════════════════════════════════════════════════════
' PASO 8: DEFINIR RANGOS NOMBRADOS
' ══════════════════════════════════════════════════════════════
Private Sub DefinirRangosNombrados(ws As Worksheet)
    ' Eliminar rangos nombrados existentes si existen
    Dim nm As Name
    On Error Resume Next
    For Each nm In ThisWorkbook.Names
        If InStr(nm.Name, "rngCert") > 0 Then
            nm.Delete
        End If
    Next nm
    On Error GoTo 0
    
    Dim nombreHoja As String
    nombreHoja = "'" & ws.Name & "'!"
    
    ' Rango del bloque categoría
    ThisWorkbook.Names.Add Name:="rngCertBloqueCategoria", RefersTo:=nombreHoja & "$A$6:$G$8"
    
    ' Rangos de la sección de datos (desplazados +3 filas)
    ThisWorkbook.Names.Add Name:="rngCertFechaInsp", RefersTo:=nombreHoja & "$C$10"
    ThisWorkbook.Names.Add Name:="rngCertHoraInicio", RefersTo:=nombreHoja & "$E$10"
    ThisWorkbook.Names.Add Name:="rngCertHoraFin", RefersTo:=nombreHoja & "$G$10"
    ThisWorkbook.Names.Add Name:="rngCertEvaluadoNombre", RefersTo:=nombreHoja & "$C$11"
    ThisWorkbook.Names.Add Name:="rngCertEvaluadoIniciales", RefersTo:=nombreHoja & "$G$11"
    ThisWorkbook.Names.Add Name:="rngCertPuesto", RefersTo:=nombreHoja & "$C$12"
    ThisWorkbook.Names.Add Name:="rngCertPlanta", RefersTo:=nombreHoja & "$F$12"
    ThisWorkbook.Names.Add Name:="rngCertArea", RefersTo:=nombreHoja & "$C$13"
    ThisWorkbook.Names.Add Name:="rngCertLinea", RefersTo:=nombreHoja & "$F$13"
    ThisWorkbook.Names.Add Name:="rngCertLugar", RefersTo:=nombreHoja & "$C$14"
    ThisWorkbook.Names.Add Name:="rngCertEvaluador", RefersTo:=nombreHoja & "$C$15"
    ThisWorkbook.Names.Add Name:="rngCertAY1", RefersTo:=nombreHoja & "$D$16"
    ThisWorkbook.Names.Add Name:="rngCertAY2", RefersTo:=nombreHoja & "$F$16"
    ThisWorkbook.Names.Add Name:="rngCertOP", RefersTo:=nombreHoja & "$G$16"
    
    ' Rangos de resultados (desplazados +3 filas)
    ThisWorkbook.Names.Add Name:="rngCertTAPuntaje", RefersTo:=nombreHoja & "$C$20"
    ThisWorkbook.Names.Add Name:="rngCertTAMaximos", RefersTo:=nombreHoja & "$E$20"
    ThisWorkbook.Names.Add Name:="rngCertTANoAplica", RefersTo:=nombreHoja & "$C$21"
    ThisWorkbook.Names.Add Name:="rngCertTAPorcentaje", RefersTo:=nombreHoja & "$C$22"
    ThisWorkbook.Names.Add Name:="rngCertRPN", RefersTo:=nombreHoja & "$C$23"
    ThisWorkbook.Names.Add Name:="rngCertCategoria", RefersTo:=nombreHoja & "$C$24"
    ThisWorkbook.Names.Add Name:="rngCertEstado", RefersTo:=nombreHoja & "$C$25"
    
    ' Rango inicial de tabla de preguntas (primera fila de datos, desplazada +3)
    ThisWorkbook.Names.Add Name:="rngCertTablaPreguntasInicio", RefersTo:=nombreHoja & "$B$27"
    
    ' Rango de fecha de validez (Paso 4 MVP - fila 50)
    ThisWorkbook.Names.Add Name:="rngCertPieValidez", RefersTo:=nombreHoja & "$A$50:$G$50"
End Sub
