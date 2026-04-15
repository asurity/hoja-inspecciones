' ----------------------------------------------------------------------
' Módulo: InspectionCalculator
' Descripción: Cálculos de Scoring TA, RPN, categorización y programación.
'              Implementa la lógica de negocio descrita en la arquitectura.
' Fecha creación: 14/04/2026
' Dependencias: Configuration2, ErrorLogger2, InspectionRepository,
'               ChecklistRepository
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Función: CalcularScoringTA
' Propósito: Calcula el scoring de Técnica Aséptica a partir de las
'            respuestas del usuario.
' Parámetros:
'   respuestas: Collection de Dictionary con claves:
'     "IDPregunta", "IDOpcion", "ValorNumerico", "IDSeccion"
'   idSeccionTA: ID de la sección de Técnica Aséptica
' Retorna: Dictionary con claves:
'   "puntaje"    - TA puntaje obtenido (suma de valores numéricos)
'   "maximos"    - TA puntos máximos (de tblConfiguracion)
'   "noaplica"   - TA puntos no aplica (ajuste del denominador)
'   "porcentaje" - TA porcentaje calculado
' ----------------------------------------------------------------------
Public Function CalcularScoringTA(ByVal respuestas As Collection, _
                                   ByVal idSeccionTA As String) As Object
    On Error GoTo ErrorHandler
    
    Dim resultado As Object
    Set resultado = CreateObject("Scripting.Dictionary")
    
    Dim puntaje As Double
    Dim maximos As Double
    Dim noAplica As Double
    Dim porcentaje As Double
    
    ' Obtener máximo base de tblConfiguracion
    maximos = ObtenerParametroNumerico("PUNTAJE_MAXIMO_TA_BASE")
    If maximos = 0 Then maximos = 57 ' Valor por defecto
    
    ' Obtener valor de la opción "No" para esta sección (para ajuste No Aplica)
    Dim valorNo As Double
    valorNo = ObtenerValorOpcionNo(idSeccionTA)
    
    ' Recorrer respuestas filtrando solo las de la sección TA
    Dim resp As Variant
    For Each resp In respuestas
        Dim dictResp As Object
        Set dictResp = resp
        
        If dictResp("IDSeccion") = idSeccionTA Then
            Dim valorNumerico As Double
            valorNumerico = CDbl(dictResp("ValorNumerico"))
            
            ' Sumar al puntaje total (incluye valores negativos de No Aplica)
            puntaje = puntaje + valorNumerico
            
            ' Si es "No Aplica" (valor -1), acumular ajuste del denominador
            ' usando el valor de "No" en lugar de -1
            If valorNumerico = -1 Then
                noAplica = noAplica + valorNo
            End If
        End If
    Next resp
    
    ' Calcular porcentaje: (puntaje × 100) / (maximos - noaplica)
    Dim denominador As Double
    denominador = maximos - noAplica
    
    If denominador > 0 Then
        porcentaje = (puntaje * 100) / denominador
    Else
        porcentaje = 0 ' Todas las respuestas son "No Aplica"
    End If
    
    ' Redondear a 2 decimales
    porcentaje = Round(porcentaje, 2)
    
    resultado("puntaje") = puntaje
    resultado("maximos") = maximos
    resultado("noaplica") = noAplica
    resultado("porcentaje") = porcentaje
    
    Set CalcularScoringTA = resultado
    Exit Function
    
ErrorHandler:
    Set resultado = CreateObject("Scripting.Dictionary")
    resultado("puntaje") = 0
    resultado("maximos") = 0
    resultado("noaplica") = 0
    resultado("porcentaje") = 0
    Set CalcularScoringTA = resultado
    Call ErrorLogger2.Log("InspectionCalculator.CalcularScoringTA", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: CalcularRPN
' Propósito: Calcula el RPN (Risk Priority Number) a partir del scoring TA.
'            Según arquitectura: RPN = TA porcentaje.
' Parámetros:
'   taData: Dictionary retornado por CalcularScoringTA
' Retorna: Valor RPN (Double)
' ----------------------------------------------------------------------
Public Function CalcularRPN(ByVal taData As Object) As Double
    On Error GoTo ErrorHandler
    
    ' Según arquitectura documentada: RPN = TA porcentaje
    CalcularRPN = CDbl(taData("porcentaje"))
    Exit Function
    
ErrorHandler:
    CalcularRPN = 0
    Call ErrorLogger2.Log("InspectionCalculator.CalcularRPN", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: DeterminarCategoria
' Propósito: Determina la categoría de resultado según el RPN calculado
'            y el historial de inspecciones anteriores.
' Lógica:
'   1. Si RPN > 20, buscar últimas 2 inspecciones del mismo Personal+Plantilla
'   2. Si las 3 (actual + 2 anteriores) tienen RPN > 20 → Categoría 5
'   3. Si no → buscar en tblCategoriasRPN el rango que contiene el RPN
' Parámetros:
'   rpn: Valor RPN calculado
'   iniciales: Iniciales del personal evaluado
'   idPlantilla: ID de la plantilla usada
' Retorna: Número de categoría (1-5)
' ----------------------------------------------------------------------
Public Function DeterminarCategoria(ByVal rpn As Double, _
                                     ByVal iniciales As String, _
                                     ByVal idPlantilla As String) As Long
    On Error GoTo ErrorHandler
    
    ' Paso 1: Validar Categoría 5 (recurrencia)
    If rpn > 20 Then
        Dim historial As Collection
        Set historial = InspectionRepository.ObtenerUltimasNInspecciones(iniciales, idPlantilla, 2)
        
        If historial.Count >= 2 Then
            If CDbl(historial(1)) > 20 And CDbl(historial(2)) > 20 Then
                DeterminarCategoria = 5
                Exit Function
            End If
        End If
    End If
    
    ' Paso 2: Categorización normal por rangos de tblCategoriasRPN
    Dim wsConfig As Worksheet
    Dim tblCategorias As ListObject
    Dim catRow As ListRow
    
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tblCategorias = wsConfig.ListObjects(Configuration2.TABLE_CATEGORIAS_RPN)
    
    If tblCategorias.DataBodyRange Is Nothing Then
        ' Fallback hardcoded si la tabla está vacía
        DeterminarCategoria = CategorizarPorRangoDefault(rpn)
        Exit Function
    End If
    
    For Each catRow In tblCategorias.ListRows
        Dim numCat As Long
        Dim rpnMin As Double
        Dim rpnMax As Double
        Dim reqHistorico As Boolean
        
        numCat = CLng(catRow.Range.Cells(1, tblCategorias.ListColumns("Numero categoria").Index).Value)
        rpnMin = CDbl(catRow.Range.Cells(1, tblCategorias.ListColumns("RPN minimo").Index).Value)
        rpnMax = CDbl(catRow.Range.Cells(1, tblCategorias.ListColumns("RPN maximo").Index).Value)
        
        On Error Resume Next
        reqHistorico = CBool(catRow.Range.Cells(1, tblCategorias.ListColumns("Requiere historico").Index).Value)
        On Error GoTo ErrorHandler
        
        ' Saltar categoría 5 en búsqueda por rango (ya se evaluó arriba)
        If reqHistorico Then GoTo SiguienteCategoria
        
        If rpn >= rpnMin And rpn <= rpnMax Then
            DeterminarCategoria = numCat
            Exit Function
        End If
        
SiguienteCategoria:
    Next catRow
    
    ' Fallback: si ningún rango coincide, usar rango por defecto
    DeterminarCategoria = CategorizarPorRangoDefault(rpn)
    Exit Function
    
ErrorHandler:
    DeterminarCategoria = 0
    Call ErrorLogger2.Log("InspectionCalculator.DeterminarCategoria", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: CalcularFechaProxima
' Propósito: Calcula la fecha de la próxima inspección.
' Fórmula: Fecha inspección + frecuencia en meses
' Parámetros:
'   fechaInspeccion: Fecha de la inspección actual
'   frecuenciaMeses: Frecuencia en meses (de tblPlantillas)
' Retorna: Fecha próxima inspección
' ----------------------------------------------------------------------
Public Function CalcularFechaProxima(ByVal fechaInspeccion As Date, _
                                      ByVal frecuenciaMeses As Long) As Date
    On Error GoTo ErrorHandler
    
    If frecuenciaMeses <= 0 Then
        frecuenciaMeses = ObtenerParametroNumerico(Configuration2.PARAM_FRECUENCIA_DEFAULT)
        If frecuenciaMeses <= 0 Then frecuenciaMeses = 3
    End If
    
    CalcularFechaProxima = DateAdd("m", frecuenciaMeses, fechaInspeccion)
    Exit Function
    
ErrorHandler:
    CalcularFechaProxima = DateAdd("m", 3, fechaInspeccion)
    Call ErrorLogger2.Log("InspectionCalculator.CalcularFechaProxima", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: CalcularDiasVencimiento
' Propósito: Calcula los días restantes hasta la próxima inspección.
' Retorna: Número de días (positivo = vigente, negativo = vencido)
' ----------------------------------------------------------------------
Public Function CalcularDiasVencimiento(ByVal fechaProxima As Date) As Long
    CalcularDiasVencimiento = CLng(fechaProxima - Date)
End Function

'' ----------------------------------------------------------------------
' Función: DeterminarEstadoProgramacion
' Propósito: Determina el estado de programación según los días de vencimiento.
' Retorna: "Vigente", "Por vencer" o "Vencido"
' ----------------------------------------------------------------------
Public Function DeterminarEstadoProgramacion(ByVal diasVencimiento As Long) As String
    On Error GoTo ErrorHandler
    
    Dim diasAlerta As Long
    diasAlerta = ObtenerParametroNumerico(Configuration2.PARAM_DIAS_ALERTA_VENCIMIENTO)
    If diasAlerta <= 0 Then diasAlerta = 15
    
    If diasVencimiento <= 0 Then
        DeterminarEstadoProgramacion = Configuration2.ESTADO_VENCIDO
    ElseIf diasVencimiento <= diasAlerta Then
        DeterminarEstadoProgramacion = Configuration2.ESTADO_POR_VENCER
    Else
        DeterminarEstadoProgramacion = Configuration2.ESTADO_VIGENTE
    End If
    Exit Function
    
ErrorHandler:
    DeterminarEstadoProgramacion = Configuration2.ESTADO_VIGENTE
    Call ErrorLogger2.Log("InspectionCalculator.DeterminarEstadoProgramacion", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: DeterminarRequiereAccion
' Propósito: Determina si la inspección requiere acción correctiva.
' Regla: Categoría >= 3 requiere acción.
' ----------------------------------------------------------------------
Public Function DeterminarRequiereAccion(ByVal categoria As Long) As Boolean
    DeterminarRequiereAccion = (categoria >= 3)
End Function

' ======================================================================
' FUNCIONES AUXILIARES PRIVADAS
' ======================================================================

'' ----------------------------------------------------------------------
' Función: ObtenerParametroNumerico
' Propósito: Lee un parámetro numérico de tblConfiguracion.
' ----------------------------------------------------------------------
Private Function ObtenerParametroNumerico(ByVal clave As String) As Double
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim tblConfig As ListObject
    Dim configRow As ListRow
    
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tblConfig = wsConfig.ListObjects(Configuration2.TABLE_CONFIGURACION)
    
    If tblConfig.DataBodyRange Is Nothing Then
        ObtenerParametroNumerico = 0
        Exit Function
    End If
    
    For Each configRow In tblConfig.ListRows
        Dim claveActual As String
        claveActual = Trim(configRow.Range.Cells(1, tblConfig.ListColumns("Clave").Index).Value)
        
        If claveActual = clave Then
            ObtenerParametroNumerico = CDbl(configRow.Range.Cells(1, tblConfig.ListColumns("Valor").Index).Value)
            Exit Function
        End If
    Next configRow
    
    ObtenerParametroNumerico = 0
    Exit Function
    
ErrorHandler:
    ObtenerParametroNumerico = 0
    Call ErrorLogger2.Log("InspectionCalculator.ObtenerParametroNumerico", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerValorOpcionNo
' Propósito: Obtiene el valor numérico de la opción "No" para una sección.
'            Se usa para calcular el ajuste del denominador cuando hay
'            respuestas "No Aplica".
' ----------------------------------------------------------------------
Private Function ObtenerValorOpcionNo(ByVal idSeccion As String) As Double
    On Error GoTo ErrorHandler
    
    Dim opciones As Collection
    Set opciones = ChecklistRepository.ObtenerOpcionesRespuesta(idSeccion)
    
    Dim op As Variant
    For Each op In opciones
        Dim arrOp() As Variant
        arrOp = op
        
        Dim textoOpcion As String
        textoOpcion = Trim(CStr(arrOp(1)))
        
        If textoOpcion = "No" Then
            ObtenerValorOpcionNo = CDbl(arrOp(2))
            Exit Function
        End If
    Next op
    
    ' Fallback: valor por defecto de "No" = 4
    ObtenerValorOpcionNo = 4
    Exit Function
    
ErrorHandler:
    ObtenerValorOpcionNo = 4
    Call ErrorLogger2.Log("InspectionCalculator.ObtenerValorOpcionNo", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: CategorizarPorRangoDefault
' Propósito: Fallback de categorización si tblCategoriasRPN está vacía.
'            Usa los rangos documentados en la arquitectura.
' ----------------------------------------------------------------------
Private Function CategorizarPorRangoDefault(ByVal rpn As Double) As Long
    If rpn >= 0 And rpn <= 14 Then
        CategorizarPorRangoDefault = 1
    ElseIf rpn >= 15 And rpn <= 19 Then
        CategorizarPorRangoDefault = 2
    ElseIf rpn >= 20 And rpn <= 40 Then
        CategorizarPorRangoDefault = 3
    ElseIf rpn > 40 Then
        CategorizarPorRangoDefault = 4
    Else
        CategorizarPorRangoDefault = 1
    End If
End Function
