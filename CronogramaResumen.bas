Attribute VB_Name = "CronogramaResumen"
' ----------------------------------------------------------------------
' Módulo: CronogramaResumen
' Descripción: Gestiona la tabla resumen del cronograma en el Menú principal.
'              Lee tblCronogramaInspecciones, filtra por planta, ordena por
'              criticidad de puesto y urgencia, y escribe en tblResumenCronograma.
' Fecha creación: 14/04/2026
' Dependencias: Configuration2, ErrorLogger2, InspectionScheduler
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Subrutina: RefrescarResumenCronograma
' Propósito: Lee tblCronogramaInspecciones, filtra por planta seleccionada,
'            ordena por criticidad de puesto y urgencia, y actualiza
'            tblResumenCronograma en el Menú principal.
' Lógica:
'   1. Leer filtro de planta desde celda RESUMEN_FILTRO_PLANTA_CELDA
'   2. Leer todos los registros activos de tblCronogramaInspecciones
'   3. Filtrar por planta (si no es "Todas")
'   4. Excluir puestos inactivos y personal inactivo
'   5. Ordenar por criticidad de puesto (primario) y días vencimiento (secundario)
'   6. Escribir resultados en tblResumenCronograma
' ----------------------------------------------------------------------
Public Sub RefrescarResumenCronograma()
    On Error GoTo ErrorHandler
    
    Dim wsMenu As Worksheet
    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    Dim tblResumen As ListObject
    
    Dim filtroPlanta As String
    Dim ordenCriticidad As Variant
    Dim datos() As Variant
    Dim cantRegistros As Long
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    ' --- Obtener referencias ---
    Set wsMenu = ThisWorkbook.Sheets(Configuration2.MAIN_MENU_SHEET)
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    Set tblResumen = wsMenu.ListObjects(Configuration2.TABLE_RESUMEN_CRONOGRAMA)
    
    ' --- Leer filtro de planta ---
    filtroPlanta = Trim(wsMenu.Range(Configuration2.RESUMEN_FILTRO_PLANTA_CELDA).Value)
    If filtroPlanta = "" Then filtroPlanta = "Todas"
    
    ' --- Obtener orden de criticidad ---
    ordenCriticidad = Configuration2.GetOrdenCriticidadPuestos()
    
    ' --- Limpiar tabla resumen ---
    If Not tblResumen.DataBodyRange Is Nothing Then
        tblResumen.DataBodyRange.Delete
    End If
    
    ' --- Verificar que hay datos en cronograma ---
    If tblCronograma.DataBodyRange Is Nothing Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' --- Recopilar registros válidos ---
    Dim registros As Collection
    Set registros = New Collection
    
    Dim cronogramaRow As ListRow
    For Each cronogramaRow In tblCronograma.ListRows
        Dim personalActivo As String
        Dim puestoActivo As String
        Dim plantaPersonal As String
        Dim puesto As String
        Dim iniciales As String
        Dim idPlantilla As String
        Dim idCronograma As String
        Dim fechaProxima As Variant
        Dim diasVencimiento As Variant
        Dim estadoCrono As String
        
        personalActivo = Trim(cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Personal activo").Index).Value)
        puestoActivo = Trim(cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Puesto activo en personal").Index).Value)
        plantaPersonal = Trim(cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Planta personal").Index).Value)
        puesto = Trim(cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Puesto").Index).Value)
        
        ' Filtrar: solo personal activo con puesto activo
        If UCase(personalActivo) <> "SI" Then GoTo SiguienteRegistro
        If UCase(puestoActivo) <> "SI" Then GoTo SiguienteRegistro
        
        ' Filtrar por planta (si no es "Todas")
        If filtroPlanta <> "Todas" Then
            If plantaPersonal <> filtroPlanta Then GoTo SiguienteRegistro
        End If
        
        ' Excluir puestos no priorizados (Quimico, Digitador, Etiquetado)
        If Not EsPuestoPriorizado(puesto, ordenCriticidad) Then GoTo SiguienteRegistro
        
        ' Recopilar datos del registro
        iniciales = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Iniciales personal").Index).Value
        idPlantilla = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("ID Plantilla").Index).Value
        idCronograma = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("ID Cronograma").Index).Value
        fechaProxima = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Fecha proxima inspeccion").Index).Value
        diasVencimiento = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Dias para vencimiento").Index).Value
        estadoCrono = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Estado cronograma").Index).Value
        
        ' RECALCULAR días de vencimiento considerando TODO EL MES
        ' Si la fecha próxima es 22-Mayo-2026, tiene hasta 31-Mayo-2026
        If IsDate(fechaProxima) Then
            Dim ultimoDiaMes As Date
            ultimoDiaMes = DateSerial(Year(fechaProxima), Month(fechaProxima) + 1, 0) ' Último día del mes
            diasVencimiento = CLng(ultimoDiaMes - Date) ' Días hasta fin del mes
        Else
            diasVencimiento = -9999 ' Máxima urgencia si no hay fecha
        End If
        
        ' Almacenar como array: (iniciales, puesto, fechaProxima, diasVencimiento, idCronograma, idPlantilla, indiceCriticidad)
        Dim indiceCriticidad As Long
        indiceCriticidad = ObtenerIndiceCriticidad(puesto, ordenCriticidad)
        
        Dim reg(0 To 6) As Variant
        reg(0) = iniciales
        reg(1) = puesto
        reg(2) = fechaProxima
        reg(3) = CLng(diasVencimiento)
        reg(4) = idCronograma
        reg(5) = idPlantilla
        reg(6) = indiceCriticidad
        
        registros.Add reg
        
SiguienteRegistro:
    Next cronogramaRow
    
    ' --- Si no hay registros, salir ---
    If registros.Count = 0 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' --- Convertir a array 2D para ordenar ---
    Dim arrDatos() As Variant
    ReDim arrDatos(1 To registros.Count, 1 To 7)
    For i = 1 To registros.Count
        Dim r As Variant
        r = registros(i)
        arrDatos(i, 1) = r(0) ' iniciales
        arrDatos(i, 2) = r(1) ' puesto
        arrDatos(i, 3) = r(2) ' fechaProxima
        arrDatos(i, 4) = r(3) ' diasVencimiento
        arrDatos(i, 5) = r(4) ' idCronograma
        arrDatos(i, 6) = r(5) ' idPlantilla
        arrDatos(i, 7) = r(6) ' indiceCriticidad
    Next i
    
    ' --- Ordenar: primario por criticidad ASC, secundario por días vencimiento ASC ---
    Call OrdenarArray2D(arrDatos, registros.Count)
    
    ' --- Escribir en tblResumenCronograma ---
    For i = 1 To registros.Count
        Dim newRow As ListRow
        Set newRow = tblResumen.ListRows.Add
        
        With newRow.Range
            .Cells(1, tblResumen.ListColumns("Iniciales").Index).Value = arrDatos(i, 1)
            .Cells(1, tblResumen.ListColumns("Puesto").Index).Value = arrDatos(i, 2)
            .Cells(1, tblResumen.ListColumns("Mes proxima inspeccion").Index).Value = ConvertirFechaAMes(arrDatos(i, 3))
            .Cells(1, tblResumen.ListColumns("Dias vencimiento").Index).Value = arrDatos(i, 4)
            .Cells(1, tblResumen.ListColumns("ID Cronograma").Index).Value = arrDatos(i, 5)
            .Cells(1, tblResumen.ListColumns("ID Plantilla").Index).Value = arrDatos(i, 6)
        End With
    Next i
    
    ' --- Aplicar formato condicional tipo semáforo ---
    Call AplicarFormatoSemaforoVencimiento(tblResumen)
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Call ErrorLogger2.Log("CronogramaResumen.RefrescarResumenCronograma", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Subrutina: AplicarFormatoSemaforoVencimiento
' Propósito: Aplica formato condicional tipo semáforo a la columna
'            "Dias vencimiento" según los rangos de urgencia.
' Parámetros:
'   tblResumen: Tabla tblResumenCronograma
' ----------------------------------------------------------------------
Private Sub AplicarFormatoSemaforoVencimiento(ByRef tblResumen As ListObject)
    On Error GoTo ErrorHandler
    
    ' Si no hay datos, salir
    If tblResumen.DataBodyRange Is Nothing Then Exit Sub
    
    Dim colDiasVencimiento As Range
    Dim colIndex As Long
    Dim celda As Range
    Dim diasValor As Variant
    
    ' Obtener columna "Dias vencimiento"
    colIndex = tblResumen.ListColumns("Dias vencimiento").Index
    Set colDiasVencimiento = tblResumen.ListColumns(colIndex).DataBodyRange
    
    ' Limpiar formato previo
    colDiasVencimiento.Interior.Pattern = xlNone
    colDiasVencimiento.Font.Bold = False
    
    ' Aplicar formato celda por celda
    For Each celda In colDiasVencimiento.Cells
        If IsNumeric(celda.Value) Then
            diasValor = CLng(celda.Value)
            
            ' 🔴 VENCIDO: < 0
            If diasValor < 0 Then
                celda.Interior.Color = RGB(255, 199, 206) ' #FFC7CE - Rojo claro
                celda.Font.Color = RGB(156, 0, 6)          ' #9C0006 - Rojo oscuro
                celda.Font.Bold = True
            
            ' 🟠 URGENTE: 0-10 días
            ElseIf diasValor >= 0 And diasValor <= 10 Then
                celda.Interior.Color = RGB(255, 235, 156) ' #FFEB9C - Naranja claro
                celda.Font.Color = RGB(156, 101, 0)       ' #9C6500 - Naranja oscuro
                celda.Font.Bold = True
            
            ' 🟡 PRÓXIMO: 11-40 días
            ElseIf diasValor > 10 And diasValor <= 40 Then
                celda.Interior.Color = RGB(255, 255, 204) ' #FFFFCC - Amarillo claro
                celda.Font.Color = RGB(100, 100, 0)       ' #646400 - Amarillo oscuro
            
            ' 🟢 OK: > 40 días
            ElseIf diasValor > 40 Then
                celda.Interior.Color = RGB(198, 239, 206) ' #C6EFCE - Verde claro
                celda.Font.Color = RGB(0, 97, 0)          ' #006100 - Verde oscuro
            End If
        End If
    Next celda
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en AplicarFormatoSemaforoVencimiento: " & Err.Description
End Sub

'' ----------------------------------------------------------------------
' Subrutina: FiltrarResumenPorPlanta
' Propósito: Llamado al cambiar el filtro de planta. Refresca la tabla.
' ----------------------------------------------------------------------
Public Sub FiltrarResumenPorPlanta()
    Call RefrescarResumenCronograma
End Sub

'' ----------------------------------------------------------------------
' Función: EsPuestoPriorizado
' Propósito: Verifica si un puesto está en la lista de criticidad
'            (excluye Quimico, Digitador, Etiquetado).
' ----------------------------------------------------------------------
Private Function EsPuestoPriorizado(ByVal puesto As String, ByRef ordenCriticidad As Variant) As Boolean
    Dim i As Long
    For i = LBound(ordenCriticidad) To UBound(ordenCriticidad)
        If Trim(CStr(ordenCriticidad(i))) = Trim(puesto) Then
            EsPuestoPriorizado = True
            Exit Function
        End If
    Next i
    EsPuestoPriorizado = False
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerIndiceCriticidad
' Propósito: Retorna el índice de criticidad de un puesto (0 = más crítico).
' ----------------------------------------------------------------------
Private Function ObtenerIndiceCriticidad(ByVal puesto As String, ByRef ordenCriticidad As Variant) As Long
    Dim i As Long
    For i = LBound(ordenCriticidad) To UBound(ordenCriticidad)
        If Trim(CStr(ordenCriticidad(i))) = Trim(puesto) Then
            ObtenerIndiceCriticidad = i
            Exit Function
        End If
    Next i
    ObtenerIndiceCriticidad = 999 ' No encontrado, baja prioridad
End Function

'' ----------------------------------------------------------------------
' Subrutina: OrdenarArray2D
' Propósito: Ordena array 2D por columna 7 (criticidad ASC) y luego
'            columna 4 (días vencimiento ASC). Bubble sort simple.
' ----------------------------------------------------------------------
Private Sub OrdenarArray2D(ByRef arr() As Variant, ByVal n As Long)
    Dim i As Long, j As Long, k As Long
    Dim temp As Variant
    Dim swapped As Boolean
    
    For i = 1 To n - 1
        swapped = False
        For j = 1 To n - i
            Dim doSwap As Boolean
            doSwap = False
            
            ' Comparar primero por criticidad (col 7)
            If arr(j, 7) > arr(j + 1, 7) Then
                doSwap = True
            ElseIf arr(j, 7) = arr(j + 1, 7) Then
                ' Si misma criticidad, comparar por días vencimiento (col 4) ASC
                If arr(j, 4) > arr(j + 1, 4) Then
                    doSwap = True
                End If
            End If
            
            If doSwap Then
                ' Intercambiar filas
                For k = 1 To 7
                    temp = arr(j, k)
                    arr(j, k) = arr(j + 1, k)
                    arr(j + 1, k) = temp
                Next k
                swapped = True
            End If
        Next j
        If Not swapped Then Exit For
    Next i
End Sub

'' ----------------------------------------------------------------------
' Función: ConvertirFechaAMes
' Propósito: Convierte una fecha a formato de mes con primera letra mayúscula
'            (ej: "Mayo", "Julio"). Solo muestra el mes porque el cliente
'            tiene todo el mes para realizar la inspección.
' Parámetros:
'   fecha: Variant (Date o Empty)
' Retorna: String con el nombre del mes capitalizado o "-" si está vacío
' ----------------------------------------------------------------------
Private Function ConvertirFechaAMes(ByVal fecha As Variant) As String
    On Error GoTo ErrorHandler
    
    ' Si la fecha está vacía o no es válida, retornar "-"
    If IsEmpty(fecha) Or Not IsDate(fecha) Then
        ConvertirFechaAMes = "-"
        Exit Function
    End If
    
    ' Obtener el mes en formato completo (ej: "mayo")
    Dim mesCompleto As String
    mesCompleto = LCase(Format(fecha, "mmmm"))
    
    ' Capitalizar primera letra
    If Len(mesCompleto) > 0 Then
        ConvertirFechaAMes = UCase(Left(mesCompleto, 1)) & Mid(mesCompleto, 2)
    Else
        ConvertirFechaAMes = "-"
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertirFechaAMes = "-"
End Function
