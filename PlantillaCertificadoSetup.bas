Attribute VB_Name = "PlantillaCertificadoSetup"
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
            MsgBox "Operación cancelada.", vbInformation
            Exit Sub
        End If
        
        ' Eliminar hoja existente
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Crear nueva hoja
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = Configuration2.SHEET_PLANTILLA_CERTIFICADO
    
    ' Desproteger libro temporalmente si está protegido
    Dim libroProtegido As Boolean
    libroProtegido = ThisWorkbook.ProtectStructure
    If libroProtegido Then
        ThisWorkbook.Unprotect Configuration2.APP_PASSWORD
    End If
    
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
    
    ' PASO 6: Diseñar encabezado de tabla de preguntas
    Call DisenarEncabezadoTabla(ws)
    
    ' PASO 7: Diseñar placeholders para observaciones y firmas
    Call DisenarSeccionObservacionesFirmas(ws)
    
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
        .Range("C1").HorizontalAlignment = xlCenter
        .Range("C1").VerticalAlignment = xlCenter
        
        ' Subtítulo (C3:G3)
        .Range("C3:G3").Merge
        .Range("C3").Value = "TÉCNICA ASÉPTICA"
        .Range("C3").Font.Name = "Arial"
        .Range("C3").Font.Size = 14
        .Range("C3").Font.Bold = True
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
    End With
End Sub

' ══════════════════════════════════════════════════════════════
' PASO 4: DISEÑAR SECCIÓN DATOS DE INSPECCIÓN (Filas 6-14)
' ══════════════════════════════════════════════════════════════
Private Sub DisenarSeccionDatos(ws As Worksheet)
    With ws
        ' Título de sección (fila 6)
        .Range("A6:G6").Merge
        .Range("A6").Value = "DATOS DE INSPECCIÓN"
        .Range("A6").Font.Name = "Arial"
        .Range("A6").Font.Size = 11
        .Range("A6").Font.Bold = True
        .Range("A6").HorizontalAlignment = xlCenter
        .Range("A6").Interior.Color = RGB(214, 234, 248)  ' Azul claro
        
        ' Fila 7: Fecha y hora
        .Range("B7").Value = "Fecha inspección:"
        .Range("C7").Value = "[FECHA]"
        .Range("D7").Value = "Hora:"
        .Range("E7").Value = "[INICIO]"
        .Range("F7").Value = "-"
        .Range("G7").Value = "[FIN]"
        
        ' Fila 8: Evaluado
        .Range("B8").Value = "Evaluado:"
        .Range("C8:E8").Merge
        .Range("C8").Value = "[NOMBRE COMPLETO]"
        .Range("F8").Value = "Iniciales:"
        .Range("G8").Value = "[ABC]"
        
        ' Fila 9: Puesto y Planta
        .Range("B9").Value = "Puesto:"
        .Range("C9:D9").Merge
        .Range("C9").Value = "[PUESTO]"
        .Range("E9").Value = "Planta:"
        .Range("F9:G9").Merge
        .Range("F9").Value = "[PLANTA]"
        
        ' Fila 10: Área y Línea
        .Range("B10").Value = "Área:"
        .Range("C10:D10").Merge
        .Range("C10").Value = "[ÁREA]"
        .Range("E10").Value = "Línea:"
        .Range("F10:G10").Merge
        .Range("F10").Value = "[LÍNEA]"
        
        ' Fila 11: Lugar auditoría
        .Range("B11").Value = "Lugar auditoría:"
        .Range("C11:G11").Merge
        .Range("C11").Value = "[LUGAR]"
        
        ' Fila 12: Evaluador
        .Range("B12").Value = "Evaluador:"
        .Range("C12:G12").Merge
        .Range("C12").Value = "[EVALUADOR]"
        
        ' Fila 13: Personal línea
        .Range("B13").Value = "Personal línea:"
        .Range("C13").Value = "AY1:"
        .Range("D13").Value = "[---]"
        .Range("E13").Value = "AY2:"
        .Range("F13").Value = "[---]"
        .Range("G13").Value = "OP: [---]"
        
        ' Fila 14: (espacio en blanco)
        .Range("A14:G14").Value = ""
        
        ' Aplicar formato de fuente a toda la sección
        .Range("B7:G13").Font.Name = "Arial"
        .Range("B7:G13").Font.Size = 9
        
        ' Labels en negrita
        .Range("B7,D7,F7,B8,F8,B9,E9,B10,E10,B11,B12,B13,C13,E13").Font.Bold = True
        
        ' Bordes finos
        .Range("A6:G13").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        .Range("A6:G13").Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range("A6:G13").Borders(xlInsideHorizontal).Weight = xlThin
    End With
End Sub

' ══════════════════════════════════════════════════════════════
' PASO 5: DISEÑAR SECCIÓN RESULTADOS GENERALES (Filas 16-21)
' ══════════════════════════════════════════════════════════════
Private Sub DisenarSeccionResultados(ws As Worksheet)
    With ws
        ' Título de sección (fila 16)
        .Range("A16:G16").Merge
        .Range("A16").Value = "RESULTADOS GENERALES"
        .Range("A16").Font.Name = "Arial"
        .Range("A16").Font.Size = 11
        .Range("A16").Font.Bold = True
        .Range("A16").HorizontalAlignment = xlCenter
        .Range("A16").Interior.Color = RGB(213, 244, 230)  ' Verde claro
        
        ' Fila 17: TA Puntaje obtenido
        .Range("B17").Value = "TA Puntaje obtenido:"
        .Range("C17").Value = "[X]"
        .Range("D17").Value = "/ máximos:"
        .Range("E17").Value = "[Y]"
        
        ' Fila 18: TA Puntos no aplica
        .Range("B18").Value = "TA Puntos no aplica:"
        .Range("C18").Value = "[Z]"
        
        ' Fila 19: TA Porcentaje
        .Range("B19").Value = "TA Porcentaje:"
        .Range("C19").Value = "[XX.XX]"
        .Range("D19").Value = "%"
        
        ' Fila 20: RPN
        .Range("B20").Value = "RPN:"
        .Range("C20").Value = "[XX.XX]"
        
        ' Fila 21: Categoría
        .Range("B21").Value = "Categoría resultado:"
        .Range("C21:G21").Merge
        .Range("C21").Value = "[N - Descripción]"
        
        ' Aplicar formato de fuente
        .Range("B17:G21").Font.Name = "Arial"
        .Range("B17:G21").Font.Size = 9
        
        ' Labels en negrita
        .Range("B17,D17,B18,B19,D19,B20,B21").Font.Bold = True
        
        ' Bordes
        .Range("A16:G21").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        .Range("A16:G21").Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range("A16:G21").Borders(xlInsideHorizontal).Weight = xlThin
    End With
End Sub

' ══════════════════════════════════════════════════════════════
' PASO 6: DISEÑAR ENCABEZADO DE TABLA DE PREGUNTAS (Fila 23)
' ══════════════════════════════════════════════════════════════
Private Sub DisenarEncabezadoTabla(ws As Worksheet)
    With ws
        ' Encabezados de columna
        .Range("B23").Value = "#"
        .Range("C23:D23").Merge
        .Range("C23").Value = "PREGUNTA"
        .Range("E23").Value = "RESPUESTA"
        .Range("F23:G23").Merge
        .Range("F23").Value = "OBSERVACIÓN"
        
        ' Formato de encabezado
        .Range("B23:G23").Font.Name = "Arial"
        .Range("B23:G23").Font.Size = 8
        .Range("B23:G23").Font.Bold = True
        .Range("B23:G23").HorizontalAlignment = xlCenter
        .Range("B23:G23").VerticalAlignment = xlCenter
        .Range("B23:G23").Interior.Color = RGB(204, 204, 204)  ' Gris oscuro
        
        ' Bordes gruesos
        .Range("B23:G23").BorderAround LineStyle:=xlContinuous, Weight:=xlThick
        .Range("B23:G23").Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range("B23:G23").Borders(xlInsideVertical).Weight = xlMedium
        
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
        
        ' Fila 50: Placeholder para observaciones (se moverá dinámicamente)
        .Range("B50").Value = "[SECCIÓN OBSERVACIONES - Se creará dinámicamente]"
        .Range("B50").Font.Name = "Arial"
        .Range("B50").Font.Size = 8
        .Range("B50").Font.Color = RGB(150, 150, 150)
        .Range("B50").Font.Italic = True
        
        ' Fila 55: Placeholder para firmas (se moverá dinámicamente)
        .Range("B55").Value = "[SECCIÓN FIRMAS - Se creará dinámicamente]"
        .Range("B55").Font.Name = "Arial"
        .Range("B55").Font.Size = 8
        .Range("B55").Font.Color = RGB(150, 150, 150)
        .Range("B55").Font.Italic = True
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
    
    ' Rangos de la sección de datos
    ThisWorkbook.Names.Add Name:="rngCertFechaInsp", RefersTo:=nombreHoja & "$C$7"
    ThisWorkbook.Names.Add Name:="rngCertHoraInicio", RefersTo:=nombreHoja & "$E$7"
    ThisWorkbook.Names.Add Name:="rngCertHoraFin", RefersTo:=nombreHoja & "$G$7"
    ThisWorkbook.Names.Add Name:="rngCertEvaluadoNombre", RefersTo:=nombreHoja & "$C$8"
    ThisWorkbook.Names.Add Name:="rngCertEvaluadoIniciales", RefersTo:=nombreHoja & "$G$8"
    ThisWorkbook.Names.Add Name:="rngCertPuesto", RefersTo:=nombreHoja & "$C$9"
    ThisWorkbook.Names.Add Name:="rngCertPlanta", RefersTo:=nombreHoja & "$F$9"
    ThisWorkbook.Names.Add Name:="rngCertArea", RefersTo:=nombreHoja & "$C$10"
    ThisWorkbook.Names.Add Name:="rngCertLinea", RefersTo:=nombreHoja & "$F$10"
    ThisWorkbook.Names.Add Name:="rngCertLugar", RefersTo:=nombreHoja & "$C$11"
    ThisWorkbook.Names.Add Name:="rngCertEvaluador", RefersTo:=nombreHoja & "$C$12"
    ThisWorkbook.Names.Add Name:="rngCertAY1", RefersTo:=nombreHoja & "$D$13"
    ThisWorkbook.Names.Add Name:="rngCertAY2", RefersTo:=nombreHoja & "$F$13"
    ThisWorkbook.Names.Add Name:="rngCertOP", RefersTo:=nombreHoja & "$G$13"
    
    ' Rangos de resultados
    ThisWorkbook.Names.Add Name:="rngCertTAPuntaje", RefersTo:=nombreHoja & "$C$17"
    ThisWorkbook.Names.Add Name:="rngCertTAMaximos", RefersTo:=nombreHoja & "$E$17"
    ThisWorkbook.Names.Add Name:="rngCertTANoAplica", RefersTo:=nombreHoja & "$C$18"
    ThisWorkbook.Names.Add Name:="rngCertTAPorcentaje", RefersTo:=nombreHoja & "$C$19"
    ThisWorkbook.Names.Add Name:="rngCertRPN", RefersTo:=nombreHoja & "$C$20"
    ThisWorkbook.Names.Add Name:="rngCertCategoria", RefersTo:=nombreHoja & "$C$21"
    
    ' Rango inicial de tabla de preguntas (primera fila de datos)
    ThisWorkbook.Names.Add Name:="rngCertTablaPreguntasInicio", RefersTo:=nombreHoja & "$B$24"
End Sub
