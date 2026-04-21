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
'     "IDPregunta", "IDOpcion", "ValorNumerico", "IDSeccion", "IDCriticidad"
'   idSeccionTA: ID de la sección de Técnica Aséptica
'   idPlantilla: ID de la plantilla para calcular máximos dinámicos
' Retorna: Dictionary con claves:
'   "puntaje"    - TA puntaje obtenido (suma de valores SOLO de respuestas que NO son "No Aplica")
'   "maximos"    - TA puntos máximos (suma de valores de criticidad de todas las preguntas TA)
'   "noaplica"   - TA puntos no aplica (suma de valores de criticidad de respuestas "No Aplica")
'   "porcentaje" - TA porcentaje calculado
' ----------------------------------------------------------------------
Public Function CalcularScoringTA(ByVal respuestas As Collection, _
                                   ByVal idSeccionTA As String, _
                                   ByVal idPlantilla As String) As Object
    On Error GoTo ErrorHandler
    
    Debug.Print "  [CalcularScoringTA] INICIO"
    Debug.Print "    ID Sección TA: " & idSeccionTA
    Debug.Print "    ID Plantilla: " & idPlantilla
    
    Dim resultado As Object
    Set resultado = CreateObject("Scripting.Dictionary")
    
    Dim puntaje As Double
    Dim maximos As Double
    Dim noAplica As Double
    Dim porcentaje As Double
    
    ' ===================================================================
    ' PASO 1: Calcular máximos dinámicos sumando valores de criticidad
    ' de todas las preguntas de TA de esta plantilla
    ' ===================================================================
    Debug.Print "    [PASO 1] Calculando máximos dinámicos..."
    
    Dim preguntasTA As Collection
    Set preguntasTA = ChecklistRepository.ObtenerPreguntasPorPlantillaYSeccion(idPlantilla, idSeccionTA)
    
    Debug.Print "      Total preguntas TA en plantilla: " & preguntasTA.Count
    
    Dim preg As Variant
    For Each preg In preguntasTA
        Dim idCriticidadPreg As String
        idCriticidadPreg = CStr(preg(4)) ' preg(4) = ID Criticidad
        
        Dim valorCriticidad As Double
        valorCriticidad = ObtenerValorCriticidad(idCriticidadPreg)
        
        maximos = maximos + valorCriticidad
        
        Debug.Print "      Pregunta " & preg(1) & " | Criticidad: " & idCriticidadPreg & " | Valor: " & valorCriticidad
    Next preg
    
    Debug.Print "      Máximos calculados: " & maximos
    
    ' ===================================================================
    ' PASO 2: Procesar respuestas para calcular puntaje y ajuste No Aplica
    ' ===================================================================
    Debug.Print "    [PASO 2] Procesando respuestas..."
    Debug.Print "      Total respuestas recibidas: " & respuestas.Count
    
    Dim resp As Variant
    Dim contadorTA As Long
    contadorTA = 0
    
    For Each resp In respuestas
        Dim dictResp As Object
        Set dictResp = resp
        
        ' Filtrar solo respuestas de la sección TA
        If dictResp("IDSeccion") = idSeccionTA Then
            contadorTA = contadorTA + 1
            
            Dim idPregunta As String
            Dim idOpcion As String
            Dim valorNumerico As Double
            Dim idCriticidad As String
            
            idPregunta = CStr(dictResp("IDPregunta"))
            idOpcion = CStr(dictResp("IDOpcion"))
            valorNumerico = CDbl(dictResp("ValorNumerico"))
            idCriticidad = CStr(dictResp("IDCriticidad"))
            
            ' Obtener texto de la opción para identificar "No Aplica"
            Dim textoOpcion As String
            textoOpcion = ObtenerTextoOpcionPorID(idOpcion)
            
            Debug.Print "        Respuesta #" & contadorTA
            Debug.Print "          IDPregunta: " & idPregunta
            Debug.Print "          IDOpcion: " & idOpcion
            Debug.Print "          TextoOpcion: " & textoOpcion
            Debug.Print "          ValorNumerico: " & valorNumerico
            Debug.Print "          IDCriticidad: " & idCriticidad
            
            ' Verificar si es "No Aplica"
            If UCase(Trim(textoOpcion)) = "NO APLICA" Then
                ' Es No Aplica: NO sumar al puntaje, SÍ ajustar denominador
                Dim valorCrit As Double
                valorCrit = ObtenerValorCriticidad(idCriticidad)
                noAplica = noAplica + valorCrit
                
                Debug.Print "          >>> ES NO APLICA -> No suma a puntaje, ajuste denominador: +" & valorCrit
            Else
                ' NO es No Aplica: sumar valor al puntaje
                puntaje = puntaje + valorNumerico
                
                Debug.Print "          >>> NO es No Aplica -> Suma a puntaje: +" & valorNumerico
            End If
        End If
    Next resp
    
    Debug.Print "      Respuestas TA procesadas: " & contadorTA
    Debug.Print "      Puntaje obtenido: " & puntaje
    Debug.Print "      No Aplica (ajuste): " & noAplica
    
    ' ===================================================================
    ' PASO 3: Calcular porcentaje con denominador ajustado
    ' ===================================================================
    Debug.Print "    [PASO 3] Calculando porcentaje..."
    
    Dim denominador As Double
    denominador = maximos - noAplica
    
    Debug.Print "      Denominador (maximos - noaplica): " & maximos & " - " & noAplica & " = " & denominador
    
    If denominador > 0 Then
        porcentaje = (puntaje * 100) / denominador
        Debug.Print "      Porcentaje: (" & puntaje & " × 100) / " & denominador & " = " & porcentaje
    Else
        porcentaje = 0 ' Todas las respuestas son "No Aplica"
        Debug.Print "      Porcentaje: 0 (denominador = 0, todas No Aplica)"
    End If
    
    ' Redondear a 2 decimales
    porcentaje = Round(porcentaje, 2)
    
    resultado("puntaje") = puntaje
    resultado("maximos") = maximos
    resultado("noaplica") = noAplica
    resultado("porcentaje") = porcentaje
    
    Debug.Print "  [CalcularScoringTA] FIN - Porcentaje final: " & Format(porcentaje, "0.00") & "%"
    
    Set CalcularScoringTA = resultado
    Exit Function
    
ErrorHandler:
    Debug.Print "  [CalcularScoringTA] ERROR: " & Err.Description
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
' Función: ObtenerValorCriticidad
' Propósito: Obtiene el valor numérico de una criticidad dado su ID.
' Parámetros:
'   idCriticidad: ID de la criticidad (FK a tblCriticidad)
' Retorna: Valor numérico de la criticidad (ej: Crítica=8, Mayor=4, Menor=1)
' ----------------------------------------------------------------------
Private Function ObtenerValorCriticidad(ByVal idCriticidad As String) As Double
    On Error GoTo ErrorHandler
    
    Dim wsChecklist As Worksheet
    Dim tblCriticidad As ListObject
    Dim critRow As ListRow
    
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set tblCriticidad = wsChecklist.ListObjects(Configuration2.TABLE_CRITICIDAD)
    
    If tblCriticidad.DataBodyRange Is Nothing Then
        ObtenerValorCriticidad = 1 ' Fallback: Menor
        Exit Function
    End If
    
    ' Buscar manualmente sin usar .Index
    Dim idColIdx As Long
    Dim valorColIdx As Long
    idColIdx = 0
    valorColIdx = 0
    
    Dim col As ListColumn
    For Each col In tblCriticidad.ListColumns
        If col.Name = "ID Criticidad" Then idColIdx = col.Index
        If col.Name = "Valor" Then valorColIdx = col.Index
    Next col
    
    If idColIdx = 0 Or valorColIdx = 0 Then
        ObtenerValorCriticidad = 1 ' Fallback
        Exit Function
    End If
    
    For Each critRow In tblCriticidad.ListRows
        Dim idActual As String
        idActual = Trim(CStr(critRow.Range.Cells(1, idColIdx).Value))
        
        If idActual = Trim(idCriticidad) Then
            ObtenerValorCriticidad = CDbl(critRow.Range.Cells(1, valorColIdx).Value)
            Exit Function
        End If
    Next critRow
    
    ' Fallback: si no se encuentra, retornar 1 (Menor)
    ObtenerValorCriticidad = 1
    Exit Function
    
ErrorHandler:
    ObtenerValorCriticidad = 1
    Call ErrorLogger2.Log("InspectionCalculator.ObtenerValorCriticidad", Err.Description, Err.Number)
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

'' ----------------------------------------------------------------------
' Función: ContarRespuestasPorCriticidad
' Propósito: Cuenta las respuestas de la sección Auditoría de Procesos
'            agrupadas por criticidad de la pregunta.
' Parámetros:
'   respuestas: Collection de Dictionary con claves:
'     "IDPregunta", "IDOpcion", "ValorNumerico", "IDSeccion", "IDCriticidad"
'   idSeccionProcesos: ID de la sección de Auditoría de Procesos
' Retorna: Dictionary con claves por nombre de criticidad:
'   "Crítica_Cumple", "Crítica_NoCumple", "Crítica_NoAplica"
'   "Mayor_Cumple", "Mayor_NoCumple", "Mayor_NoAplica"
'   "Menor_Cumple", "Menor_NoCumple", "Menor_NoAplica"
' ----------------------------------------------------------------------
Public Function ContarRespuestasPorCriticidad(ByVal respuestas As Collection, _
                                               ByVal idSeccionProcesos As String) As Object
    On Error GoTo ErrorHandler
    
    Dim resultado As Object
    Set resultado = CreateObject("Scripting.Dictionary")
    
    ' Inicializar contadores
    resultado("Crítica_Cumple") = 0
    resultado("Crítica_NoCumple") = 0
    resultado("Crítica_NoAplica") = 0
    resultado("Mayor_Cumple") = 0
    resultado("Mayor_NoCumple") = 0
    resultado("Mayor_NoAplica") = 0
    resultado("Menor_Cumple") = 0
    resultado("Menor_NoCumple") = 0
    resultado("Menor_NoAplica") = 0
    
    ' Recorrer respuestas filtrando solo las de la sección Auditoría de Procesos
     Debug.Print "  [ContarRespuestasPorCriticidad] Total respuestas recibidas: " & respuestas.Count
    Debug.Print "  [ContarRespuestasPorCriticidad] ID Sección Procesos buscada: " & idSeccionProcesos
    
    Dim resp As Variant
    Dim contadorProcesadas As Long
    contadorProcesadas = 0
    
    For Each resp In respuestas
        Dim dictResp As Object
        Set dictResp = resp
        
        ' Debug: Mostrar IDSeccion de cada respuesta
        Debug.Print "    Respuesta - IDSeccion: " & dictResp("IDSeccion") & " | IDPregunta: " & dictResp("IDPregunta")
        
        ' Solo procesar respuestas de la sección de Auditoría de Procesos
        If dictResp("IDSeccion") = idSeccionProcesos Then
            contadorProcesadas = contadorProcesadas + 1
            
            ' Obtener criticidad de la pregunta
            Dim idCriticidad As String
            Dim nombreCriticidad As String
            idCriticidad = dictResp("IDCriticidad")
            nombreCriticidad = ObtenerNombreCriticidad(idCriticidad)
            
            ' Obtener texto de la opción seleccionada
            Dim idOpcion As String
            Dim textoOpcion As String
            idOpcion = dictResp("IDOpcion")
            textoOpcion = ObtenerTextoOpcionPorID(idOpcion)
            
            ' Debug detallado
            Debug.Print "      >>> PROCESANDO RESPUESTA #" & contadorProcesadas
            Debug.Print "          IDCriticidad: " & idCriticidad
            Debug.Print "          NombreCriticidad: " & nombreCriticidad
            Debug.Print "          IDOpcion: " & idOpcion
            Debug.Print "          TextoOpcion: " & textoOpcion
            
            ' Incrementar contador correspondiente
            Dim clave As String
            Select Case UCase(Trim(textoOpcion))
                Case "CUMPLE"
                    clave = nombreCriticidad & "_Cumple"
                    If resultado.Exists(clave) Then
                        resultado(clave) = resultado(clave) + 1
                        Debug.Print "          Incrementado: " & clave & " = " & resultado(clave)
                    Else
                        Debug.Print "          ERROR: Clave no existe: " & clave
                    End If
                    
                Case "NO CUMPLE"
                    clave = nombreCriticidad & "_NoCumple"
                    If resultado.Exists(clave) Then
                        resultado(clave) = resultado(clave) + 1
                        Debug.Print "          Incrementado: " & clave & " = " & resultado(clave)
                    Else
                        Debug.Print "          ERROR: Clave no existe: " & clave
                    End If
                    
                Case "NO APLICA"
                    clave = nombreCriticidad & "_NoAplica"
                    If resultado.Exists(clave) Then
                        resultado(clave) = resultado(clave) + 1
                        Debug.Print "          Incrementado: " & clave & " = " & resultado(clave)
                    Else
                        Debug.Print "          ERROR: Clave no existe: " & clave
                    End If
                    
                Case Else
                    Debug.Print "          ADVERTENCIA: TextoOpcion no reconocido: " & textoOpcion
            End Select
        End If
    Next resp
    
    Debug.Print "  [ContarRespuestasPorCriticidad] Total respuestas procesadas: " & contadorProcesadas
    
    Set ContarRespuestasPorCriticidad = resultado
    Exit Function
    
ErrorHandler:
    Set resultado = CreateObject("Scripting.Dictionary")
    resultado("Crítica_Cumple") = 0
    resultado("Crítica_NoCumple") = 0
    resultado("Crítica_NoAplica") = 0
    resultado("Mayor_Cumple") = 0
    resultado("Mayor_NoCumple") = 0
    resultado("Mayor_NoAplica") = 0
    resultado("Menor_Cumple") = 0
    resultado("Menor_NoCumple") = 0
    resultado("Menor_NoAplica") = 0
    Set ContarRespuestasPorCriticidad = resultado
    Call ErrorLogger2.Log("InspectionCalculator.ContarRespuestasPorCriticidad", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: EvaluarAuditoriaProcesos
' Propósito: Determina el resultado final de Auditoría de Procesos según
'            las reglas de negocio basadas en criticidad.
' Reglas (en orden de prioridad):
'   1. Si hay 2 o más "No Cumple" de criticidad "Crítica" → "No Cumple"
'   2. Si hay 1 "No Cumple" de criticidad "Crítica" Y 2 o más "No Cumple"
'      de criticidad "Mayor" → "No Cumple"
'   3. Si hay 4 o más "No Cumple" de criticidad "Mayor" → "No Cumple"
'   4. En cualquier otro caso → "Cumple"
' Parámetros:
'   conteo: Dictionary retornado por ContarRespuestasPorCriticidad
' Retorna: "Cumple" o "No Cumple"
' ----------------------------------------------------------------------
Public Function EvaluarAuditoriaProcesos(ByVal conteo As Object) As String
    On Error GoTo ErrorHandler
    
    Dim criticaNoCumple As Long
    Dim mayorNoCumple As Long
    
    ' Obtener contadores
    criticaNoCumple = CLng(conteo("Crítica_NoCumple"))
    mayorNoCumple = CLng(conteo("Mayor_NoCumple"))
    
    ' Aplicar reglas de negocio en orden de prioridad
    
    ' Regla 1: 2 o más "No Cumple" de criticidad "Crítica"
    If criticaNoCumple >= 2 Then
        EvaluarAuditoriaProcesos = "No Cumple"
        Exit Function
    End If
    
    ' Regla 2: 1 "No Cumple" de criticidad "Crítica" Y 
    '          2 o más "No Cumple" de criticidad "Mayor"
    If criticaNoCumple >= 1 And mayorNoCumple >= 2 Then
        EvaluarAuditoriaProcesos = "No Cumple"
        Exit Function
    End If
    
    ' Regla 3: 4 o más "No Cumple" de criticidad "Mayor"
    If mayorNoCumple >= 4 Then
        EvaluarAuditoriaProcesos = "No Cumple"
        Exit Function
    End If
    
    ' Regla 4: En cualquier otro caso → "Cumple"
    EvaluarAuditoriaProcesos = "Cumple"
    Exit Function
    
ErrorHandler:
    EvaluarAuditoriaProcesos = "Cumple" ' Por defecto en caso de error
    Call ErrorLogger2.Log("InspectionCalculator.EvaluarAuditoriaProcesos", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerNombreCriticidad
' Propósito: Obtiene el nombre de criticidad a partir del ID.
' Parámetros:
'   idCriticidad: ID de la criticidad (FK a tblCriticidad)
' Retorna: Nombre de la criticidad ("Crítica", "Mayor", "Menor")
' ----------------------------------------------------------------------
Private Function ObtenerNombreCriticidad(ByVal idCriticidad As String) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "        [ObtenerNombreCriticidad] Buscando ID: '" & idCriticidad & "'"
    
    Dim wsConfig As Worksheet
    Dim tblCriticidad As ListObject
    Dim critRow As ListRow
    
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set tblCriticidad = wsConfig.ListObjects(Configuration2.TABLE_CRITICIDAD)
    
    If tblCriticidad.DataBodyRange Is Nothing Then
        Debug.Print "        [ObtenerNombreCriticidad] ERROR: Tabla vacía, retornando 'Menor'"
        ObtenerNombreCriticidad = "Menor" ' Fallback
        Exit Function
    End If
    
    Debug.Print "        [ObtenerNombreCriticidad] Filas en tblCriticidad: " & tblCriticidad.ListRows.Count
    
    ' Buscar índices de columnas de forma robusta SIN usar .Index
    Dim colIDIndex As Long
    Dim colNombreIndex As Long
    Dim col As ListColumn
    Dim i As Long
    
    ' Buscar columna "ID Criticidad"
    colIDIndex = 0
    On Error Resume Next
    For i = 1 To tblCriticidad.ListColumns.Count
        If InStr(1, tblCriticidad.ListColumns(i).Name, "Criticidad", vbTextCompare) > 0 And _
           InStr(1, tblCriticidad.ListColumns(i).Name, "ID", vbTextCompare) > 0 Then
            colIDIndex = i
            Exit For
        End If
    Next i
    On Error GoTo ErrorHandler
    
    If colIDIndex = 0 Then colIDIndex = 1  ' Fallback a columna 1
    Debug.Print "        [ObtenerNombreCriticidad] Usando columna ID: " & colIDIndex
    
    ' Buscar columna "Nombre de criticidad"
    colNombreIndex = 0
    On Error Resume Next
    For i = 1 To tblCriticidad.ListColumns.Count
        If InStr(1, tblCriticidad.ListColumns(i).Name, "Nombre", vbTextCompare) > 0 And _
           InStr(1, tblCriticidad.ListColumns(i).Name, "criticidad", vbTextCompare) > 0 Then
            colNombreIndex = i
            Exit For
        End If
    Next i
    On Error GoTo ErrorHandler
    
    If colNombreIndex = 0 Then colNombreIndex = 2  ' Fallback a columna 2
    Debug.Print "        [ObtenerNombreCriticidad] Usando columna Nombre: " & colNombreIndex
    
    Dim contador As Long
    contador = 0
    For Each critRow In tblCriticidad.ListRows
        contador = contador + 1
        Dim idActual As String
        On Error Resume Next
        idActual = Trim(CStr(critRow.Range.Cells(1, colIDIndex).Value))
        On Error GoTo ErrorHandler
        
        If contador <= 3 Then  ' Solo mostrar primeras 3 para no saturar
            Debug.Print "        [ObtenerNombreCriticidad] Fila " & contador & " - ID: '" & idActual & "'"
        End If
        
        If idActual = Trim(idCriticidad) Then
            Dim nombreEncontrado As String
            On Error Resume Next
            nombreEncontrado = Trim(CStr(critRow.Range.Cells(1, colNombreIndex).Value))
            On Error GoTo ErrorHandler
            Debug.Print "        [ObtenerNombreCriticidad] ¡MATCH! Retornando: '" & nombreEncontrado & "'"
            ObtenerNombreCriticidad = nombreEncontrado
            Exit Function
        End If
    Next critRow
    
    ' Si no se encuentra, retornar "Menor" como fallback
    Debug.Print "        [ObtenerNombreCriticidad] NO ENCONTRADO. Retornando 'Menor' como fallback"
    ObtenerNombreCriticidad = "Menor"
    Exit Function
    
ErrorHandler:
    Debug.Print "        [ObtenerNombreCriticidad] ERROR: " & Err.Description
    ObtenerNombreCriticidad = "Menor"
    Call ErrorLogger2.Log("InspectionCalculator.ObtenerNombreCriticidad", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerTextoOpcionPorID
' Propósito: Obtiene el texto de una opción a partir de su ID.
' Parámetros:
'   idOpcion: ID de la opción (FK a tblOpcionesDeRespuesta)
' Retorna: Texto de la opción ("Cumple", "No Cumple", "No Aplica", etc.)
' ----------------------------------------------------------------------
Private Function ObtenerTextoOpcionPorID(ByVal idOpcion As String) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "        [ObtenerTextoOpcionPorID] Buscando ID: '" & idOpcion & "'"
    
    Dim wsConfig As Worksheet
    Dim tblOpciones As ListObject
    Dim opcionRow As ListRow
    
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set tblOpciones = wsConfig.ListObjects(Configuration2.TABLE_OPCIONES)
    
    If tblOpciones.DataBodyRange Is Nothing Then
        Debug.Print "        [ObtenerTextoOpcionPorID] ERROR: Tabla vacía"
        ObtenerTextoOpcionPorID = ""
        Exit Function
    End If
    
    Debug.Print "        [ObtenerTextoOpcionPorID] Filas en tblOpciones: " & tblOpciones.ListRows.Count
    
    ' Buscar índices de columnas de forma robusta SIN usar .Index
    Dim colIDIndex As Long
    Dim colTextoIndex As Long
    Dim i As Long
    
    ' Buscar columna "ID Opcion" o "ID Opción"
    colIDIndex = 0
    On Error Resume Next
    For i = 1 To tblOpciones.ListColumns.Count
        Dim nombreCol As String
        nombreCol = tblOpciones.ListColumns(i).Name
        If (InStr(1, nombreCol, "ID", vbTextCompare) > 0 And _
            InStr(1, nombreCol, "Opcion", vbTextCompare) > 0) Or _
           (InStr(1, nombreCol, "ID", vbTextCompare) > 0 And _
            InStr(1, nombreCol, "Opción", vbTextCompare) > 0) Then
            colIDIndex = i
            Exit For
        End If
    Next i
    On Error GoTo ErrorHandler
    
    If colIDIndex = 0 Then colIDIndex = 1  ' Fallback a columna 1
    Debug.Print "        [ObtenerTextoOpcionPorID] Usando columna ID: " & colIDIndex
    
    ' Buscar columna "Opción texto" o "Opcion texto"
    colTextoIndex = 0
    On Error Resume Next
    For i = 1 To tblOpciones.ListColumns.Count
        nombreCol = tblOpciones.ListColumns(i).Name
        If (InStr(1, nombreCol, "texto", vbTextCompare) > 0 And _
            InStr(1, nombreCol, "Opcion", vbTextCompare) > 0) Or _
           (InStr(1, nombreCol, "texto", vbTextCompare) > 0 And _
            InStr(1, nombreCol, "Opción", vbTextCompare) > 0) Then
            colTextoIndex = i
            Exit For
        End If
    Next i
    On Error GoTo ErrorHandler
    
    If colTextoIndex = 0 Then colTextoIndex = 2  ' Fallback a columna 2
    Debug.Print "        [ObtenerTextoOpcionPorID] Usando columna Texto: " & colTextoIndex
    
    Dim contador As Long
    contador = 0
    For Each opcionRow In tblOpciones.ListRows
        contador = contador + 1
        Dim idActual As String
        On Error Resume Next
        idActual = Trim(CStr(opcionRow.Range.Cells(1, colIDIndex).Value))
        On Error GoTo ErrorHandler
        
        If contador <= 3 Then ' Solo mostrar primeros 3 para no saturar
            Debug.Print "        [ObtenerTextoOpcionPorID] Fila " & contador & " - ID: '" & idActual & "'"
        End If
        
        If idActual = Trim(idOpcion) Then
            Dim textoEncontrado As String
            On Error Resume Next
            textoEncontrado = Trim(CStr(opcionRow.Range.Cells(1, colTextoIndex).Value))
            On Error GoTo ErrorHandler
            Debug.Print "        [ObtenerTextoOpcionPorID] ¡MATCH! Retornando: '" & textoEncontrado & "'"
            ObtenerTextoOpcionPorID = textoEncontrado
            Exit Function
        End If
    Next opcionRow
    
    Debug.Print "        [ObtenerTextoOpcionPorID] NO ENCONTRADO. Retornando vacío"
    ObtenerTextoOpcionPorID = ""
    Exit Function
    
ErrorHandler:
    Debug.Print "        [ObtenerTextoOpcionPorID] ERROR: " & Err.Description
    ObtenerTextoOpcionPorID = ""
    Call ErrorLogger2.Log("InspectionCalculator.ObtenerTextoOpcionPorID", Err.Description, Err.Number)
End Function
