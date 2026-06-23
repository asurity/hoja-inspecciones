' ══════════════════════════════════════════════════════════════════════════
' Módulo: InspectionHistoryService
' Descripción: Gestión de histórico de inspecciones del personal
'              Servicios de consulta para inspecciones recurrentes
' Fecha creación: 21/04/2026
' Autor: Sistema TH-HC-001
' Versión: 1.0.0
' ══════════════════════════════════════════════════════════════════════════
' DEPENDENCIAS:
'   - InspectionRepository.bas (consultas a tblInspecciones)
'   - Configuration2.bas (nombres de tablas y columnas)
'   - ErrorLogger2.bas (logging de errores)
' ══════════════════════════════════════════════════════════════════════════
Option Explicit

' ══════════════════════════════════════════════════════════════════════════
' Buscar todas las inspecciones previas de un personal
' ══════════════════════════════════════════════════════════════════════════
' Parámetros:
'   - iniciales: Iniciales del personal (requerido)
'   - filtroPorPuesto: Si True, filtra por puesto específico (default: True según Q1)
'   - puesto: Puesto a filtrar (requerido si filtroPorPuesto = True)
'   - idPlantilla: ID de plantilla (opcional, para mayor especificidad)
'
' Retorna:
'   Collection de Dictionary con estructura:
'   - "IDInspeccion" (String)
'   - "FechaInspeccion" (Date)
'   - "PuestoEvaluado" (String)
'   - "RPN" (Double) - RPN calculado o RPN Total si es recurrente
'   - "Categoria" (String)
'   - "NumeroInspeccion" (Long)
'   - "EsRecurrente" (Boolean)
'   
'   Ordenados por Fecha DESC (más reciente primero)
'
' Excepciones:
'   - Err.Raise si iniciales vacío
'   - Err.Raise si filtroPorPuesto=True y puesto vacío
' ══════════════════════════════════════════════════════════════════════════
Public Function BuscarInspeccionesPrevias( _
    ByVal iniciales As String, _
    Optional ByVal filtroPorPuesto As Boolean = True, _
    Optional ByVal puesto As String = "", _
    Optional ByVal idPlantilla As String = "" _
) As Collection
    
    On Error GoTo ErrorHandler
    
    ' Validaciones de entrada
    If Trim(iniciales) = "" Then
        Err.Raise vbObjectError + 5001, "InspectionHistoryService.BuscarInspeccionesPrevias", _
                  "Las iniciales del personal son requeridas"
    End If
    
    If filtroPorPuesto And Trim(puesto) = "" Then
        Err.Raise vbObjectError + 5002, "InspectionHistoryService.BuscarInspeccionesPrevias", _
                  "El puesto es requerido cuando filtroPorPuesto=True"
    End If
    
    Debug.Print "[HISTORY] Buscando inspecciones previas: " & iniciales & _
                IIf(filtroPorPuesto, " | Puesto: " & puesto, " | Todos los puestos")
    
    ' Obtener referencia a tabla
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Worksheets(Configuration2.SHEET_HISTORICO)
    Set tbl = ws.ListObjects(Configuration2.TABLE_INSPECCIONES)
    
    If tbl.ListRows.Count = 0 Then
        Debug.Print "[HISTORY] No hay registros en tblInspecciones"
        Set BuscarInspeccionesPrevias = New Collection
        Exit Function
    End If
    
    ' Obtener índices de columnas (usando nombres para evitar hardcoding)
    Dim colIDInsp As Long, colIniciales As Long, colPuesto As Long
    Dim colFecha As Long, colRPN As Long, colCategoria As Long
    Dim colNumInsp As Long, colEsRec As Long, colRPNTotal As Long
    Dim colIDPlantilla As Long, colEstado As Long
    
    colIDInsp = tbl.ListColumns("ID Inspeccion").Index
    colIniciales = tbl.ListColumns("Iniciales personal").Index
    colFecha = tbl.ListColumns("Fecha inspeccion").Index
    colRPN = tbl.ListColumns("TA porcentaje").Index
    colCategoria = tbl.ListColumns("Categoria resultado").Index
    colIDPlantilla = tbl.ListColumns("ID Plantilla").Index
    colEstado = tbl.ListColumns("Estado").Index
    
    ' Columnas nuevas (pueden no existir aún en BD - manejo con IsError)
    On Error Resume Next
    colPuesto = tbl.ListColumns("Puesto Evaluado").Index
    If Err.Number <> 0 Then colPuesto = 0: Err.Clear
    
    colNumInsp = tbl.ListColumns("Numero Inspeccion").Index
    If Err.Number <> 0 Then colNumInsp = 0: Err.Clear
    
    colEsRec = tbl.ListColumns("Es Inspeccion Recurrente").Index
    If Err.Number <> 0 Then colEsRec = 0: Err.Clear
    
    colRPNTotal = tbl.ListColumns("RPN Total").Index
    If Err.Number <> 0 Then colRPNTotal = 0: Err.Clear
    
    On Error GoTo ErrorHandler
    
    ' Crear colección de resultados
    Dim resultados As Collection
    Set resultados = New Collection
    
    ' Recorrer todas las filas y filtrar
    Dim fila As ListRow
    Dim inicialesFila As String
    Dim puestoFila As String
    Dim idPlantillaFila As String
    Dim coincide As Boolean
    Dim inspeccion As Object
    
    For Each fila In tbl.ListRows
        coincide = False
        
        ' NUEVO (05/08/2026): SALTAR registros inhabilitados
        ' Las inspecciones inhabilitadas por Admin NO son inspecciones reales
        ' y NO deben interferir con la búsqueda de inspecciones previas
        ' (bloquearían el cálculo recurrente del mismo puesto+personal)
        If UCase(Trim(CStr(fila.Range.Cells(1, colEstado).Value))) = UCase(Configuration2.INSPECCION_INHABILITADA) Then
            GoTo NextFila
        End If
        
        ' Filtro primario: Iniciales
        inicialesFila = Trim(CStr(fila.Range.Cells(1, colIniciales).Value))
        If UCase(inicialesFila) = UCase(Trim(iniciales)) Then
            
            ' Filtro secundario: Puesto (si aplica)
            If filtroPorPuesto And colPuesto > 0 Then
                puestoFila = Trim(CStr(fila.Range.Cells(1, colPuesto).Value))
                If UCase(puestoFila) = UCase(Trim(puesto)) Then
                    coincide = True
                End If
            ElseIf Not filtroPorPuesto Then
                coincide = True
            End If
            
            ' Filtro terciario: ID Plantilla (si se proporcionó)
            If coincide And Trim(idPlantilla) <> "" Then
                idPlantillaFila = Trim(CStr(fila.Range.Cells(1, colIDPlantilla).Value))
                If UCase(idPlantillaFila) <> UCase(Trim(idPlantilla)) Then
                    coincide = False
                End If
            End If
            
            ' Si coincide todos los filtros, agregar a resultados
            If coincide Then
                Set inspeccion = CreateObject("Scripting.Dictionary")
                inspeccion("IDInspeccion") = CStr(fila.Range.Cells(1, colIDInsp).Value)
                
                ' Leer fecha de forma independiente de locale
                Dim cellFechaValue As Variant
                cellFechaValue = fila.Range.Cells(1, colFecha).Value
                If IsNumeric(cellFechaValue) Then
                    ' Ya es un número serial de Excel (Date)
                    inspeccion("FechaInspeccion") = CDate(cellFechaValue)
                Else
                    ' Es string → usar ParseFechaDMY
                    Dim parsedFechaHist As Variant
                    parsedFechaHist = ChecklistValidator.ParseFechaDMY(CStr(cellFechaValue))
                    If Not IsEmpty(parsedFechaHist) Then
                        inspeccion("FechaInspeccion") = CDate(parsedFechaHist)
                    Else
                        inspeccion("FechaInspeccion") = Date  ' Fallback: hoy
                    End If
                End If
                
                ' Puesto (si columna existe)
                If colPuesto > 0 Then
                    inspeccion("PuestoEvaluado") = CStr(fila.Range.Cells(1, colPuesto).Value)
                Else
                    inspeccion("PuestoEvaluado") = "[No disponible]"
                End If
                
                ' IMPORTANTE (15/06/2026 - CORREGIDO): Leer de columna "TA porcentaje"
                ' en vez de "RPN calculado". "TA porcentaje" SIEMPRE contiene el % TA puro
                ' sin factores adicionales, incluso en inspecciones recurrentes.
                ' "RPN calculado" en cambio guarda el RPN Final que en recurrentes
                ' incluye %Recuperación + %OOL, contaminando el promedio.
                '
                ' Ejemplo correcto:
                '   Inspección 2: Promedio = (TA1 + TA2) / 2 + Factores2
                '   Inspección 3: Promedio = (TA2 + TA3) / 2 + Factores3
                '
                ' Ejemplo incorrecto (lo que hacía antes):
                '   Inspección 3: Promedio = (RPNTotal2 + TA3) / 2 + Factores3
                '                          = ((TA1+TA2)/2 + Factores2 + TA3) / 2 + Factores3
                '
                inspeccion("RPN") = CDbl(fila.Range.Cells(1, colRPN).Value)
                
                inspeccion("Categoria") = CStr(fila.Range.Cells(1, colCategoria).Value)
                
                ' Número de inspección (si columna existe)
                If colNumInsp > 0 Then
                    Dim numInspValue As Variant
                    numInspValue = fila.Range.Cells(1, colNumInsp).Value
                    If IsNumeric(numInspValue) Then
                        inspeccion("NumeroInspeccion") = CLng(numInspValue)
                    Else
                        inspeccion("NumeroInspeccion") = 1
                    End If
                Else
                    inspeccion("NumeroInspeccion") = 1
                End If
                
                ' Es recurrente (si columna existe)
                If colEsRec > 0 Then
                    Dim esRecValue As String
                    esRecValue = UCase(Trim(CStr(fila.Range.Cells(1, colEsRec).Value)))
                    inspeccion("EsRecurrente") = (esRecValue = "SI" Or esRecValue = "SÍ")
                Else
                    inspeccion("EsRecurrente") = False
                End If
                
                resultados.Add inspeccion
            End If
        End If
NextFila:
    Next fila
    
    ' Ordenar resultados por fecha DESC (más reciente primero)
    Dim resultadosOrdenados As Collection
    Set resultadosOrdenados = OrdenarPorFecha(resultados)
    
    Debug.Print "[HISTORY] Encontradas " & resultadosOrdenados.Count & " inspecciones previas"
    
    Set BuscarInspeccionesPrevias = resultadosOrdenados
    Exit Function
    
ErrorHandler:
    Call ErrorLogger2.Log("InspectionHistoryService.BuscarInspeccionesPrevias", _
                          Err.Description, Err.Number)
    Err.Raise Err.Number, "InspectionHistoryService.BuscarInspeccionesPrevias", Err.Description
End Function

' ══════════════════════════════════════════════════════════════════════════
' Obtener última inspección de un personal
' ══════════════════════════════════════════════════════════════════════════
' Parámetros:
'   - iniciales: Iniciales del personal (requerido)
'   - filtroPorPuesto: Si True, filtra por puesto específico (default: True)
'   - puesto: Puesto a filtrar (requerido si filtroPorPuesto = True)
'   - idPlantilla: ID de plantilla (opcional)
'
' Retorna:
'   Dictionary con datos de última inspección (estructura igual a BuscarInspeccionesPrevias)
'   Nothing si no hay inspecciones previas
'
' Excepciones:
'   - Propaga excepciones de BuscarInspeccionesPrevias
' ══════════════════════════════════════════════════════════════════════════
Public Function ObtenerUltimaInspeccion( _
    ByVal iniciales As String, _
    Optional ByVal filtroPorPuesto As Boolean = True, _
    Optional ByVal puesto As String = "", _
    Optional ByVal idPlantilla As String = "" _
) As Object
    
    On Error GoTo ErrorHandler
    
    Dim inspecciones As Collection
    Set inspecciones = BuscarInspeccionesPrevias(iniciales, filtroPorPuesto, puesto, idPlantilla)
    
    If inspecciones.Count = 0 Then
        Debug.Print "[HISTORY] No hay inspecciones previas para " & iniciales
        Set ObtenerUltimaInspeccion = Nothing
        Exit Function
    End If
    
    ' La primera de la colección es la más reciente (ordenadas DESC)
    Dim ultimaInsp As Object
    Set ultimaInsp = inspecciones(1)
    
    ' Log de resultado (usando variable temporal para evitar error de referencia circular)
    Debug.Print "[HISTORY] Última inspección: " & ultimaInsp("IDInspeccion") & _
                " | Fecha: " & ultimaInsp("FechaInspeccion") & _
                " | RPN: " & ultimaInsp("RPN")
    
    ' Retornar el Dictionary
    Set ObtenerUltimaInspeccion = ultimaInsp
    Exit Function
    
ErrorHandler:
    Call ErrorLogger2.Log("InspectionHistoryService.ObtenerUltimaInspeccion", _
                          Err.Description, Err.Number)
    Err.Raise Err.Number, "InspectionHistoryService.ObtenerUltimaInspeccion", Err.Description
End Function

' ══════════════════════════════════════════════════════════════════════════
' Calcular número de inspección siguiente
' ══════════════════════════════════════════════════════════════════════════
' Parámetros:
'   - iniciales: Iniciales del personal (requerido)
'   - filtroPorPuesto: Si True, filtra por puesto específico (default: True)
'   - puesto: Puesto a filtrar (requerido si filtroPorPuesto = True)
'   - idPlantilla: ID de plantilla (opcional)
'
' Retorna:
'   - 1 si no hay inspecciones previas
'   - MAX(Numero Inspeccion) + 1 si hay inspecciones previas
'
' Excepciones:
'   - Propaga excepciones de BuscarInspeccionesPrevias
' ══════════════════════════════════════════════════════════════════════════
Public Function CalcularNumeroInspeccionSiguiente( _
    ByVal iniciales As String, _
    Optional ByVal filtroPorPuesto As Boolean = True, _
    Optional ByVal puesto As String = "", _
    Optional ByVal idPlantilla As String = "" _
) As Long
    
    On Error GoTo ErrorHandler
    
    Dim inspecciones As Collection
    Set inspecciones = BuscarInspeccionesPrevias(iniciales, filtroPorPuesto, puesto, idPlantilla)
    
    If inspecciones.Count = 0 Then
        Debug.Print "[HISTORY] Primera inspección para " & iniciales & _
                    IIf(filtroPorPuesto, " - " & puesto, "")
        CalcularNumeroInspeccionSiguiente = 1
        Exit Function
    End If
    
    ' Buscar el número máximo
    Dim maxNumero As Long
    maxNumero = 0
    
    Dim insp As Object
    Dim i As Long
    For i = 1 To inspecciones.Count
        Set insp = inspecciones(i)
        If insp("NumeroInspeccion") > maxNumero Then
            maxNumero = insp("NumeroInspeccion")
        End If
    Next i
    
    ' Siguiente número = MAX + 1
    CalcularNumeroInspeccionSiguiente = maxNumero + 1
    
    Debug.Print "[HISTORY] Número de inspección siguiente: " & CalcularNumeroInspeccionSiguiente & _
                " (MAX anterior: " & maxNumero & ")"
    
    Exit Function
    
ErrorHandler:
    Call ErrorLogger2.Log("InspectionHistoryService.CalcularNumeroInspeccionSiguiente", _
                          Err.Description, Err.Number)
    Err.Raise Err.Number, "InspectionHistoryService.CalcularNumeroInspeccionSiguiente", Err.Description
End Function

' ══════════════════════════════════════════════════════════════════════════
' Validar coherencia de RPN anterior manual
' ══════════════════════════════════════════════════════════════════════════
' Parámetros:
'   - rpnManual: Valor de RPN ingresado manualmente
'   - iniciales: Iniciales del personal (para validación contextual)
'   - puesto: Puesto del personal (opcional)
'
' Retorna:
'   - True si el RPN es válido
'   - False si el RPN tiene problemas
'
' Validaciones:
'   1. RPN >= 0 (0 es válido - desempeño perfecto)
'   2. RPN dentro de rangos lógicos (entre 0 y 100)
'   3. Opcional: Advertir si difiere mucho del histórico automático
'
' Excepciones:
'   - No genera excepciones, retorna False en caso de error
' ══════════════════════════════════════════════════════════════════════════
Public Function ValidarRPNAnteriorManual( _
    ByVal rpnManual As Double, _
    ByVal iniciales As String, _
    Optional ByVal puesto As String = "" _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    ValidarRPNAnteriorManual = False  ' Asumir inválido por defecto
    
    ' Validación 1: RPN >= 0 (0 = desempeño perfecto)
    If rpnManual < 0 Then
        Debug.Print "[HISTORY VALIDATION] RPN manual negativo: " & rpnManual
        Exit Function
    End If
    
    ' Validación 2: RPN dentro de rangos lógicos (0 a 100)
    If rpnManual > 100 Then
        Debug.Print "[HISTORY VALIDATION] RPN fuera de rango (0-100): " & rpnManual
        Exit Function
    End If
    
    ' Validación 3: Comparar con histórico si existe
    Dim ultimaInsp As Object
    Set ultimaInsp = ObtenerUltimaInspeccion(iniciales, True, puesto)
    
    If Not ultimaInsp Is Nothing Then
        Dim rpnHistorico As Double
        rpnHistorico = ultimaInsp("RPN")
        
        ' Advertir si difiere más de 50% del histórico
        Dim diferenciaPorcentual As Double
        diferenciaPorcentual = Abs((rpnManual - rpnHistorico) / rpnHistorico) * 100
        
        If diferenciaPorcentual > 50 Then
            Debug.Print "[HISTORY VALIDATION] ⚠️ ADVERTENCIA: RPN manual (" & rpnManual & _
                        ") difiere " & Round(diferenciaPorcentual, 1) & "% del histórico (" & _
                        rpnHistorico & ")"
            ' No invalidar, solo advertir
        End If
    End If
    
    ' Si llegó aquí, es válido
    ValidarRPNAnteriorManual = True
    Debug.Print "[HISTORY VALIDATION] ✓ RPN manual válido: " & rpnManual
    
    Exit Function
    
ErrorHandler:
    Debug.Print "[HISTORY VALIDATION] Error en validación: " & Err.Description
    ValidarRPNAnteriorManual = False
    Call ErrorLogger2.Log("InspectionHistoryService.ValidarRPNAnteriorManual", _
                          Err.Description, Err.Number)
End Function

' ══════════════════════════════════════════════════════════════════════════
' FUNCIÓN AUXILIAR: Ordenar colección por fecha (DESC)
' ══════════════════════════════════════════════════════════════════════════
' Implementación de ordenamiento burbuja simple
' Para optimizar en el futuro si hay muchos registros
' ══════════════════════════════════════════════════════════════════════════
Private Function OrdenarPorFecha(ByVal coleccion As Collection) As Collection
    
    If coleccion.Count <= 1 Then
        Set OrdenarPorFecha = coleccion
        Exit Function
    End If
    
    ' Copiar a array para ordenamiento más eficiente
    Dim items() As Object
    ReDim items(1 To coleccion.Count)
    
    Dim i As Long
    For i = 1 To coleccion.Count
        Set items(i) = coleccion(i)
    Next i
    
    ' Ordenamiento burbuja (DESC - más reciente primero)
    Dim j As Long
    Dim temp As Object
    Dim swapped As Boolean
    
    For i = 1 To UBound(items) - 1
        swapped = False
        For j = 1 To UBound(items) - i
            ' ─────────────────────────────────────────────────────
            ' CORREGIDO (15/06/2026): Cuando dos inspecciones tienen
            ' la misma fecha (ej. múltiples inspecciones el mismo día),
            ' usar NumeroInspeccion DESC como desempate para garantizar
            ' que la más reciente quede primero.
            ' ─────────────────────────────────────────────────────
            Dim fechaJ As Date
            Dim fechaJ1 As Date
            fechaJ = items(j)("FechaInspeccion")
            fechaJ1 = items(j + 1)("FechaInspeccion")
            
            Dim necesitaSwap As Boolean
            necesitaSwap = False
            
            If fechaJ < fechaJ1 Then
                ' j es más antigua que j+1 → llevar j+1 al frente
                necesitaSwap = True
            ElseIf fechaJ = fechaJ1 Then
                ' Misma fecha → DESEMPATE: NumeroInspeccion DESC
                If items(j)("NumeroInspeccion") < items(j + 1)("NumeroInspeccion") Then
                    necesitaSwap = True
                End If
            End If
            
            If necesitaSwap Then
                Set temp = items(j)
                Set items(j) = items(j + 1)
                Set items(j + 1) = temp
                swapped = True
            End If
        Next j
        If Not swapped Then Exit For  ' Optimización: salir si ya está ordenado
    Next i
    
    ' Reconstruir colección ordenada
    Dim resultado As Collection
    Set resultado = New Collection
    
    For i = 1 To UBound(items)
        resultado.Add items(i)
    Next i
    
    Set OrdenarPorFecha = resultado
End Function

' ══════════════════════════════════════════════════════════════════════════
' Función de utilidad: Formatear información de inspección para UI
' ══════════════════════════════════════════════════════════════════════════
' Retorna string formateado con información de inspección para mostrar en UI
' Formato: "Inspección #N - DD/MM/YYYY - RPN: XX.XX - Cat: N"
' ══════════════════════════════════════════════════════════════════════════
Public Function FormatearInfoInspeccion(ByVal inspeccion As Object) As String
    
    If inspeccion Is Nothing Then
        FormatearInfoInspeccion = "No hay inspecciones previas"
        Exit Function
    End If
    
    Dim resultado As String
    resultado = "Inspección #" & inspeccion("NumeroInspeccion") & _
                " - " & Format(inspeccion("FechaInspeccion"), "dd/mm/yyyy") & _
                " - RPN: " & Format(inspeccion("RPN"), "0.00") & _
                " - Cat: " & inspeccion("Categoria")
    
    If inspeccion("EsRecurrente") Then
        resultado = resultado & " (Recurrente)"
    End If
    
    FormatearInfoInspeccion = resultado
End Function

' ══════════════════════════════════════════════════════════════════════════
' Función de utilidad: Obtener resumen de historial
' ══════════════════════════════════════════════════════════════════════════
' Retorna string con resumen de todas las inspecciones previas
' Útil para mostrar en lblInfoHistorico del formulario
' ══════════════════════════════════════════════════════════════════════════
Public Function ObtenerResumenHistorial( _
    ByVal iniciales As String, _
    Optional ByVal filtroPorPuesto As Boolean = True, _
    Optional ByVal puesto As String = "", _
    Optional ByVal maxResultados As Long = 5 _
) As String
    
    On Error GoTo ErrorHandler
    
    Dim inspecciones As Collection
    Set inspecciones = BuscarInspeccionesPrevias(iniciales, filtroPorPuesto, puesto)
    
    If inspecciones.Count = 0 Then
        ObtenerResumenHistorial = "No hay inspecciones previas registradas"
        Exit Function
    End If
    
    Dim resultado As String
    resultado = "Encontradas " & inspecciones.Count & " inspecciones previas:" & vbCrLf
    
    Dim i As Long
    Dim maxMostrar As Long
    maxMostrar = IIf(inspecciones.Count < maxResultados, inspecciones.Count, maxResultados)
    
    For i = 1 To maxMostrar
        resultado = resultado & FormatearInfoInspeccion(inspecciones(i)) & vbCrLf
    Next i
    
    If inspecciones.Count > maxResultados Then
        resultado = resultado & "... y " & (inspecciones.Count - maxResultados) & " más"
    End If
    
    ObtenerResumenHistorial = resultado
    Exit Function
    
ErrorHandler:
    ObtenerResumenHistorial = "Error al obtener historial: " & Err.Description
    Call ErrorLogger2.Log("InspectionHistoryService.ObtenerResumenHistorial", _
                          Err.Description, Err.Number)
End Function
