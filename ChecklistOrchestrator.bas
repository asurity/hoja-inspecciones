' ----------------------------------------------------------------------
' Módulo: ChecklistOrchestrator
' Descripción: Orquesta el flujo completo de apertura del checklist virtual
'              y el pipeline transaccional de guardado de inspecciones.
' Fecha creación: 14/04/2026
' Dependencias: Configuration2, ErrorLogger2, AuditLogger2,
'               ChecklistRepository, ChecklistValidator,
'               InspectionRepository, InspectionCalculator,
'               InspectionScheduler, CronogramaResumen,
'               frmChecklistVirtual
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Subrutina: AbrirChecklistVirtual
' Propósito: Punto de entrada para abrir el formulario del checklist virtual.
'            Llamado desde doble clic en tblResumenCronograma o desde
'            frmSelectorInspeccion.
' Parámetros:
'   iniciales: Iniciales del personal a evaluar
'   idPlantilla: ID de la plantilla de inspección
'   puesto: Puesto del personal evaluado
'   idCronograma: ID del registro en tblCronogramaInspecciones
'   area: Área de la plantilla (opcional, desde frmSelectorInspeccion)
' ----------------------------------------------------------------------
Public Sub AbrirChecklistVirtual(ByVal iniciales As String, _
                                  ByVal idPlantilla As String, _
                                  ByVal puesto As String, _
                                  ByVal idCronograma As String, _
                                  Optional ByVal area As String = "")
    On Error GoTo ErrorHandler
    
    ' Registrar apertura del formulario en Audit Trail
    Call AuditLogger2.LogAction( _
        action:="Apertura Checklist Virtual", _
        sheetName:="Formulario", _
        dataModified:="frmChecklistVirtual", _
        beforeChange:="N/A", _
        afterChange:="Evaluado: " & iniciales & " | Puesto: " & puesto & " | ID Plantilla: " & idPlantilla, _
        moduleAndSubroutine:="ChecklistOrchestrator.AbrirChecklistVirtual" _
    )
    
    ' Obtener la planta del personal
    Dim planta As String
    planta = ChecklistRepository.ObtenerPlantaPersonal(iniciales)
    
    If Len(planta) = 0 Then
        MsgBox "No se encontró la planta del personal con iniciales '" & iniciales & "'.", _
               vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Crear y configurar el formulario
    Dim frm As frmChecklistVirtual
    Set frm = New frmChecklistVirtual
    
    frm.Evaluado = iniciales
    frm.Puesto = puesto
    frm.IDPlantilla = idPlantilla
    frm.IDCronograma = idCronograma
    frm.Planta = planta
    
    ' Asignar área si se proporciona (desde frmSelectorInspeccion)
    If Len(area) > 0 Then
        frm.Area = area
    End If
    
    ' Mostrar formulario modal
    frm.Show vbModal
    
    ' Liberar referencia explícitamente
    Set frm = Nothing
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("ChecklistOrchestrator.AbrirChecklistVirtual", Err.Description, Err.Number)
    MsgBox "Error al abrir el checklist virtual: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: GuardarInspeccionCompleta
' Propósito: Pipeline transaccional completo para guardar una inspección.
'            Ejecuta validación, persistencia, cálculos, actualización
'            de cronograma y logging de auditoría.
'            Si falla en cualquier paso, hace rollback eliminando registros
'            parciales.
' Parámetros:
'   frm: Referencia al formulario frmChecklistVirtual con todos los datos
' ----------------------------------------------------------------------
Public Sub GuardarInspeccionCompleta(ByRef frm As frmChecklistVirtual)
    On Error GoTo ErrorHandler
    
    Debug.Print "===== INICIO GuardarInspeccionCompleta ====="
    
    Dim idInspeccion As String
    Dim rollbackNecesario As Boolean
    rollbackNecesario = False
    
    ' ===================================================================
    ' PASO 1: Validar todo
    ' ===================================================================
    Debug.Print "[PASO 1] Recopilando respuestas..."
    
    ' Recopilar respuestas del form antes de validar
    frm.RecopilarObservacionesPublic
    
    Debug.Print "[PASO 1] Construyendo datos de validación..."
    Dim datosValidacion As Object
    Set datosValidacion = CreateObject("Scripting.Dictionary")
    datosValidacion("Iniciales") = frm.Evaluado
    datosValidacion("Puesto") = frm.Puesto
    datosValidacion("Planta") = frm.Planta
    datosValidacion("Area") = frm.Area
    datosValidacion("LineaAuditada") = frm.LineaAuditada
    datosValidacion("Evaluador") = frm.Evaluador
    datosValidacion("LugarAuditoria") = frm.LugarAuditoria
    datosValidacion("FechaInspeccion") = frm.FechaInspeccion
    datosValidacion("FechaAuditada") = frm.FechaAuditada
    datosValidacion("HoraInicio") = frm.HoraInicio
    datosValidacion("HoraTermino") = frm.HoraTermino
    
    Debug.Print "[PASO 1] Construyendo respuestas para validación..."
    Debug.Print "  Total preguntas esperadas: " & frm.ObtenerCantidadPreguntas()
    
    ' Construir diccionario simple de respuestas (IDPregunta → IDOpcion)
    Dim respuestasValidacion As Object
    Set respuestasValidacion = CreateObject("Scripting.Dictionary")
    Dim respConSeccion As Collection
    Set respConSeccion = frm.ObtenerRespuestasConSeccion()
    Dim rItem As Variant
    For Each rItem In respConSeccion
        Dim dItem As Object
        Set dItem = rItem
        Dim idOpcionResp As String
        idOpcionResp = Trim(CStr(dItem("IDOpcion")))
        
        If Len(idOpcionResp) > 0 Then
            respuestasValidacion(dItem("IDPregunta")) = idOpcionResp
        End If
    Next rItem
    
    Debug.Print "  Respuestas con opción seleccionada: " & respuestasValidacion.Count
    Debug.Print "[PASO 1] Validando..."
    
    ' Validar cabecera y respuestas
    Dim erroresValidacion As String
    erroresValidacion = ChecklistValidator.ValidarTodo(datosValidacion, respuestasValidacion, frm.ObtenerCantidadPreguntas())
    
    If Len(erroresValidacion) > 0 Then
        MsgBox "No se puede guardar la inspección." & vbCrLf & vbCrLf & _
               "Corrija los siguientes errores:" & vbCrLf & vbCrLf & _
               erroresValidacion, vbExclamation, "Validación Fallida"
        Exit Sub
    End If
    
    Debug.Print "[PASO 1] Validación OK"
    
    ' ===================================================================
    ' PASO 2: Confirmar con usuario
    ' ===================================================================
    Debug.Print "[PASO 2] Solicitando confirmación al usuario..."
    
    Dim respuesta As VbMsgBoxResult
    Dim fechaMostrar As String
    Dim fechaParsedMsg As Variant
    fechaParsedMsg = ChecklistValidator.ParseFechaDMY(frm.FechaInspeccion)
    If Not IsEmpty(fechaParsedMsg) Then
        fechaMostrar = Format(fechaParsedMsg, "dd/mm/yyyy")
    Else
        fechaMostrar = frm.FechaInspeccion  ' fallback: mostrar el string tal cual
    End If
    
    respuesta = MsgBox("¿Está seguro que desea guardar esta inspección?" & vbCrLf & _
                       "Evaluado: " & frm.Evaluado & vbCrLf & _
                       "Puesto: " & frm.Puesto & vbCrLf & _
                       "Fecha: " & fechaMostrar, _
                       vbQuestion + vbYesNo, "Confirmar guardado")
    
    If respuesta <> vbYes Then
        Debug.Print "[PASO 2] Usuario canceló el guardado"
        Exit Sub
    End If
    
    Debug.Print "[PASO 2] Usuario confirmó. Continuando..."
    
    ' ===================================================================
    ' PASO 3-4: Desactivar actualizaciones para rendimiento
    ' ===================================================================
    Debug.Print "[PASO 3-4] Desactivando actualizaciones..."
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Debug.Print "[PASO 3-4] Actualizaciones desactivadas"
    
    ' ===================================================================
    ' PASO 5: Construir objeto datos de inspección
    ' ===================================================================
    Debug.Print "[PASO 5] Construyendo datos de inspección..."
    Dim datos As Object
    Set datos = CreateObject("Scripting.Dictionary")
    
    datos("Iniciales") = frm.Evaluado
    datos("IDPlantilla") = frm.IDPlantilla
    datos("Planta") = frm.Planta
    datos("FechaInspeccion") = ChecklistValidator.ParseFechaDMY(frm.FechaInspeccion)
    datos("FechaAuditada") = ChecklistValidator.ParseFechaDMY(frm.FechaAuditada)
    datos("Evaluador") = frm.Evaluador
    datos("Area") = frm.Area
    datos("LineaAuditada") = frm.LineaAuditada
    datos("HoraInicio") = frm.HoraInicio
    datos("HoraTermino") = frm.HoraTermino
    datos("AY1") = frm.AY1
    datos("AY2") = frm.AY2
    datos("OP") = frm.OP
    datos("LugarAuditoria") = frm.LugarAuditoria
    datos("ObservacionGeneral") = frm.ObservacionGeneral
    
    ' NUEVOS CAMPOS - FASE 7 (23/04/2026): Calificaciones y Vencimientos
    datos("CalificacionVestuario") = frm.CalificacionVestuario
    ' ACTUALIZADO 23/06/2026: Parsear fechas de vencimiento con ParseFechaDMY
    ' para evitar que Excel las interprete según el locale (bug: 02/06/2026 → Feb 6 en vez de Jun 2)
    Dim fechaVencVestParsed As Variant
    Dim fechaVencOperParsed As Variant
    fechaVencVestParsed = ChecklistValidator.ParseFechaDMY(frm.FechaVencVestuario)
    fechaVencOperParsed = ChecklistValidator.ParseFechaDMY(frm.FechaVencOperador)
    datos("FechaVencVestuario") = IIf(Not IsEmpty(fechaVencVestParsed), fechaVencVestParsed, frm.FechaVencVestuario)
    datos("CalificacionOperador") = frm.CalificacionOperador
    datos("FechaVencOperador") = IIf(Not IsEmpty(fechaVencOperParsed), fechaVencOperParsed, frm.FechaVencOperador)
    
    Debug.Print "[PASO 5] Datos preparados OK"
    
    ' ===================================================================
    ' PASO 6: Crear registro en tblInspecciones
    ' ===================================================================
    Debug.Print "[PASO 6] Creando registro en tblInspecciones..."
    idInspeccion = InspectionRepository.CrearInspeccion(datos)
    Debug.Print "[PASO 6] ID Inspección creado: " & idInspeccion
    
    If Len(idInspeccion) = 0 Then
        Debug.Print "[PASO 6] ERROR: No se pudo crear inspección"
        Err.Raise vbObjectError + 1001, "ChecklistOrchestrator", _
                  "No se pudo crear el registro de inspección."
    End If
    
    rollbackNecesario = True
    Debug.Print "[PASO 6] Registro creado OK, rollback habilitado"
    
    ' ===================================================================
    ' PASO 7: Guardar respuestas en tblRespuestas
    ' ===================================================================
    Debug.Print "[PASO 7] Guardando respuestas en tblRespuestas..."
    ' IMPORTANTE: Usar ObtenerRespuestasConSeccion() para incluir IDCriticidad
    Dim respuestas As Collection
    Set respuestas = frm.ObtenerRespuestasConSeccion()  ' Cambiado de ObtenerRespuestas() a ObtenerRespuestasConSeccion()
    Debug.Print "  Total respuestas a guardar: " & respuestas.Count
    
    If respuestas.Count > 0 Then
        Dim primerResp As Variant
        Set primerResp = respuestas(1)
        Debug.Print "  Primer respuesta - IDPregunta: " & CStr(primerResp("IDPregunta"))
        Debug.Print "  Primer respuesta - IDOpcion: " & CStr(primerResp("IDOpcion"))
        If primerResp.Exists("IDCriticidad") Then
            Debug.Print "  Primer respuesta - IDCriticidad: " & CStr(primerResp("IDCriticidad"))
        End If
    End If
    
    Call InspectionRepository.GuardarRespuestas(idInspeccion, respuestas)
    Debug.Print "[PASO 7] Respuestas guardadas OK"
    
    ' ===================================================================
    ' PASO 8: Calcular scoring TA
    ' ===================================================================
    Debug.Print "[PASO 8] Calculando scoring TA..."
    Dim idSeccionTA As String
    idSeccionTA = frm.IDSeccionTA
    
    Dim respuestasConSeccion As Collection
    Set respuestasConSeccion = frm.ObtenerRespuestasConSeccion()
    
    Dim taData As Object
    Set taData = InspectionCalculator.CalcularScoringTA(respuestasConSeccion, idSeccionTA, frm.IDPlantilla)
    Debug.Print "  TA Puntaje: " & taData("puntaje") & "/" & taData("maximos")
    Debug.Print "  TA No Aplica: " & taData("noaplica")
    Debug.Print "  TA Denominador (maximos - noaplica): " & (taData("maximos") - taData("noaplica"))
    Debug.Print "  TA Porcentaje: " & Format(taData("porcentaje"), "0.00") & "%"
    Debug.Print "[PASO 8] Scoring TA calculado OK"
    
    ' ===================================================================
    ' PASO 8A: Contar y evaluar Auditoría de Procesos
    ' ===================================================================
    Debug.Print "[PASO 8A] Procesando Auditoría de Procesos..."
    Dim idSeccionProcesos As String
    idSeccionProcesos = frm.IDSeccionProcesos
    
    Dim resultadoProcesos As String
    resultadoProcesos = "No evaluado" ' Por defecto
    
    ' Declarar variables para conteos (se guardarán en tblInspecciones)
    Dim apCriticaNoCumple As Long
    Dim apMayorNoCumple As Long
    Dim apMenorNoCumple As Long
    apCriticaNoCumple = 0
    apMayorNoCumple = 0
    apMenorNoCumple = 0
    
    If Len(idSeccionProcesos) > 0 Then
        ' Contar respuestas por criticidad
        Dim conteoProcesos As Object
        Set conteoProcesos = InspectionCalculator.ContarRespuestasPorCriticidad(respuestasConSeccion, idSeccionProcesos)
        
        ' Debug.Print "  Auditoría de Procesos - Conteo por criticidad:"
        ' Debug.Print "    Crítica - Cumple: " & conteoProcesos("Crítica_Cumple")
        ' Debug.Print "    Crítica - No Cumple: " & conteoProcesos("Crítica_NoCumple")
        ' Debug.Print "    Crítica - No Aplica: " & conteoProcesos("Crítica_NoAplica")
        ' Debug.Print "    Mayor - Cumple: " & conteoProcesos("Mayor_Cumple")
        ' Debug.Print "    Mayor - No Cumple: " & conteoProcesos("Mayor_NoCumple")
        ' Debug.Print "    Mayor - No Aplica: " & conteoProcesos("Mayor_NoAplica")
        ' Debug.Print "    Menor - Cumple: " & conteoProcesos("Menor_Cumple")
        ' Debug.Print "    Menor - No Cumple: " & conteoProcesos("Menor_NoCumple")
        ' Debug.Print "    Menor - No Aplica: " & conteoProcesos("Menor_NoAplica")
        
        ' Evaluar resultado según reglas de negocio
        resultadoProcesos = InspectionCalculator.EvaluarAuditoriaProcesos(conteoProcesos)
        Debug.Print "  Resultado Auditoría de Procesos: " & resultadoProcesos
        
        ' Guardar conteos para persistir en tblInspecciones
        apCriticaNoCumple = CLng(conteoProcesos("Crítica_NoCumple"))
        apMayorNoCumple = CLng(conteoProcesos("Mayor_NoCumple"))
        apMenorNoCumple = CLng(conteoProcesos("Menor_NoCumple"))
        Debug.Print "  Conteos guardados - Crítica: " & apCriticaNoCumple & ", Mayor: " & apMayorNoCumple & ", Menor: " & apMenorNoCumple
    Else
        Debug.Print "  ADVERTENCIA: No se encontró ID de sección de Auditoría de Procesos"
    End If
    Debug.Print "[PASO 8A] Auditoría de Procesos procesada OK"
    
    ' ===================================================================
    ' PASO 9: Calcular métricas (RPN, Categoría) con soporte recurrente
    ' ===================================================================
    Debug.Print "[PASO 9] Calculando métricas (RPN + Categoría)..."
    
    ' Extraer datos recurrentes del formulario
    Dim esRecurrente As Boolean
    Dim numeroInspeccion As Long
    Dim rpnAnterior As Double
    Dim idInspeccionAnterior As String
    
    esRecurrente = frm.EsInspeccionRecurrente
    numeroInspeccion = frm.NumeroInspeccion
    
    Debug.Print "[CalcularMetricas] Valores del formulario:"
    Debug.Print "  esRecurrente: " & esRecurrente
    Debug.Print "  numeroInspeccion: " & numeroInspeccion
    Debug.Print "  RPNAnteriorAuto: " & frm.RPNAnteriorAuto
    Debug.Print "  RPNAnteriorManual: " & frm.RPNAnteriorManual
    Debug.Print "  IDInspeccionAnterior: '" & frm.IDInspeccionAnterior & "'"
    
    ' Determinar qué RPN anterior usar (auto tiene prioridad si existe ID)
    ' NOTA: RPNAnteriorAuto puede ser 0 (desempeño perfecto), por eso verificamos el ID
    If Len(Trim(frm.IDInspeccionAnterior)) > 0 Then
        ' Hay historial automático (aunque RPN sea 0)
        rpnAnterior = frm.RPNAnteriorAuto
        idInspeccionAnterior = frm.IDInspeccionAnterior
        Debug.Print "  Usando RPNAnteriorAuto: " & rpnAnterior & " (ID: " & idInspeccionAnterior & ")"
    Else
        ' No hay historial automático, usar manual si existe
        rpnAnterior = frm.RPNAnteriorManual
        idInspeccionAnterior = "" ' No hay ID si fue manual
        Debug.Print "  Usando RPNAnteriorManual: " & rpnAnterior
    End If
    
    ' Calcular métricas usando nuevo pipeline
    ' FASE 6 (23/04/2026): Agregar factores adicionales (% Recuperación y % OOL)
    Dim porcRecuperacion As Double
    Dim porcOOL As Double
    porcRecuperacion = frm.PorcRecuperacion
    porcOOL = frm.PorcOOL
    
    Debug.Print "  Factores adicionales - % Recuperación: " & porcRecuperacion & ", % OOL: " & porcOOL
    
    Dim metricas As Object
    ' CATEGORÍA 5 (16/06/2026): Pasar frm.Evaluado (iniciales) y frm.IDPlantilla para regla de historial
    ' BUGFIX (17/06/2026): Pasar frm.ModoRPN para distinguir modo manual de automático
    '   (rpnAnterior puede ser 0 en modo manual si la inspección anterior fue perfecta)
    Set metricas = CalcularMetricasInspeccion(taData, esRecurrente, numeroInspeccion, rpnAnterior, idInspeccionAnterior, frm.Puesto, frm.Evaluado, frm.IDPlantilla, porcRecuperacion, porcOOL, frm.ModoRPN)
    
    ' Extraer valores calculados
    Dim rpn As Double
    Dim categoria As Long
    rpn = metricas("RPN_Final")
    categoria = CLng(metricas("Categoria"))
    
    Debug.Print "  RPN final: " & Format(rpn, "0.00")
    Debug.Print "  Categoría: " & categoria
    If esRecurrente Then
        Debug.Print "  [RECURRENTE] Número inspección: " & numeroInspeccion
        Debug.Print "  [RECURRENTE] RPN Anterior: " & Format(rpnAnterior, "0.00")
        Debug.Print "  [RECURRENTE] RPN Promedio: " & Format(metricas("RPN_Promedio"), "0.00")
        Debug.Print "  [RECURRENTE] RPN Total: " & Format(metricas("RPN_Total"), "0.00")
    End If
    Debug.Print "[PASO 9] Métricas calculadas OK"
    
    ' ===================================================================
    ' PASO 10: Calcular fecha próxima
    ' ===================================================================
    Debug.Print "[PASO 10] Calculando fecha próxima inspección..."
    Dim frecuencia As Long
    frecuencia = frm.FrecuenciaMeses
    
    Dim fechaProxima As Date
    fechaProxima = InspectionCalculator.CalcularFechaProxima(ChecklistValidator.ParseFechaDMY(frm.FechaInspeccion), frecuencia)
    
    Dim diasVencimiento As Long
    diasVencimiento = InspectionCalculator.CalcularDiasVencimiento(fechaProxima)
    
    Dim estadoProgramacion As String
    estadoProgramacion = InspectionCalculator.DeterminarEstadoProgramacion(diasVencimiento)
    Debug.Print "  Fecha próxima: " & Format(fechaProxima, "dd/mm/yyyy")
    Debug.Print "  Días vencimiento: " & diasVencimiento
    Debug.Print "  Estado: " & estadoProgramacion
    Debug.Print "[PASO 10] Fecha próxima calculada OK"
    
    ' ===================================================================
    ' PASO 11: Actualizar tblInspecciones con cálculos
    ' ===================================================================
    Debug.Print "[PASO 11] Actualizando inspección con cálculos..."
    Dim calculos As Object
    Set calculos = CreateObject("Scripting.Dictionary")
    
    calculos("TA_puntaje") = taData("puntaje")
    calculos("TA_maximos") = taData("maximos")
    calculos("TA_noaplica") = taData("noaplica")
    calculos("TA_porcentaje") = taData("porcentaje")
    calculos("Auditoria_Procesos_Resultado") = resultadoProcesos
    calculos("AP_Critica_NoCumple") = apCriticaNoCumple
    calculos("AP_Mayor_NoCumple") = apMayorNoCumple
    calculos("AP_Menor_NoCumple") = apMenorNoCumple
    
    ' IMPORTANTE (23/04/2026): calculos("RPN") SIEMPRE contiene el % TA puro (sin factores)
    ' - Primera inspección: RPN = % TA
    ' - Inspecciones recurrentes: RPN = % TA actual (NO el promedio, NO el total)
    ' Esto es correcto porque:
    '   - Se guarda en columna "RPN calculado" de tblInspecciones
    '   - Se usa para calcular promedios en futuras inspecciones
    '   - El RPN Total (con factores) se guarda por separado en columna "RPN Total"
    calculos("RPN") = rpn
    calculos("Categoria") = categoria
    calculos("RequiereAccion") = InspectionCalculator.DeterminarRequiereAccion(categoria)
    calculos("FechaProxima") = fechaProxima
    calculos("DiasVencimiento") = diasVencimiento
    calculos("EstadoProgramacion") = estadoProgramacion
    
    ' Agregar datos recurrentes a calculos
    calculos("NumeroInspeccion") = metricas("NumeroInspeccion")
    calculos("EsInspeccionRecurrente") = metricas("EsInspeccionRecurrente")
    calculos("PuestoEvaluado") = metricas("PuestoEvaluado")
    
    ' NUEVO: SIEMPRE guardar RPN Total (tanto en primera como recurrente)
    ' - Primera inspección: RPN Total = RPN TA
    ' - Inspección recurrente: RPN Total = Promedio + factores adicionales
    calculos("RPNTotal") = metricas("RPN_Total")
    
    ' Guardar campos adicionales solo si es recurrente
    If metricas("EsInspeccionRecurrente") Then
        calculos("RPNAnterior") = metricas("RPN_Anterior")
        calculos("RPNPromedio") = metricas("RPN_Promedio")
        calculos("IDInspeccionAnterior") = metricas("IDInspeccionAnterior")
        
        ' Factores adicionales (FASE 6 - 23/04/2026)
        ' IMPORTANTE: SIEMPRE guardar, incluso con valor 0 para Técnico C/D y Sanitizador
        ' Esto evita columnas vacías en la base de datos y facilita reportes
        calculos("PorcRecuperacion") = metricas("PorcRecuperacion")
        calculos("PorcOOL") = metricas("PorcOOL")
    End If
    
    Call InspectionRepository.ActualizarCalculosInspeccion(idInspeccion, calculos)
    Debug.Print "[PASO 11] Cálculos actualizados OK"
    
    ' ===================================================================
    ' PASO 12: Actualizar tblCronogramaInspecciones
    ' ===================================================================
    Debug.Print "[PASO 12] Actualizando cronograma..."
    Call InspectionScheduler.ActualizarRegistroCronograma(frm.Evaluado, frm.IDPlantilla)
    Debug.Print "[PASO 12] Cronograma actualizado OK"
    
    ' ===================================================================
    ' PASO 13: Log en Audit Trail
    ' ===================================================================
    Debug.Print "[PASO 13] Registrando en Audit Trail..."
    Dim detallesAudit As String
    detallesAudit = "ID: " & idInspeccion & " | Evaluado: " & frm.Evaluado & _
        " | Puesto: " & frm.Puesto & " | RPN: " & Format(rpn, "0.00") & _
        " | Cat: " & categoria
    
    If esRecurrente Then
        detallesAudit = detallesAudit & " | RECURRENTE #" & numeroInspeccion & _
            " | RPN Ant: " & Format(rpnAnterior, "0.00")
    End If
    
    Call AuditLogger2.LogAction( _
        "Inspección completada", _
        Configuration2.SHEET_HISTORICO, _
        "tblInspecciones / tblRespuestas", _
        "", _
        detallesAudit, _
        "ChecklistOrchestrator.GuardarInspeccionCompleta")
    Debug.Print "[PASO 13] Audit Trail registrado OK"
    
    ' ===================================================================
    ' PASO 14-15: Restaurar configuración de aplicación
    ' ===================================================================
    Debug.Print "[PASO 14-15] Restaurando configuración Excel..."
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Debug.Print "[PASO 14-15] Configuración restaurada OK"
    
    ' ===================================================================
    ' PASO 15A: Guardar libro y crear copia de seguridad automática
    ' ===================================================================
    Debug.Print "[PASO 15A] Guardando libro y creando backup..."
    On Error Resume Next
    Call ThisWorkbook.Save  ' Dispara Workbook_BeforeSave → CrearBackupAutomatico
    If Err.Number <> 0 Then
        Debug.Print "[PASO 15A] AVISO: No se pudo guardar/backup: " & Err.Description
        Err.Clear
    Else
        Debug.Print "[PASO 15A] Libro guardado y backup creado OK"
    End If
    On Error GoTo ErrorHandler
    
    ' ===================================================================
    ' PASO 16: Notificar éxito y ofrecer generar certificado
    ' ===================================================================
    Debug.Print "[PASO 16] Mostrando mensaje de éxito..."
    Dim respuestaPDF As VbMsgBoxResult
    respuestaPDF = MsgBox("Inspección guardada exitosamente." & vbCrLf & _
           "RPN: " & Format(rpn, "0.00") & vbCrLf & _
           "Categoría: " & categoria & vbCrLf & _
           "Próxima inspección: " & Format(fechaProxima, "dd/mm/yyyy") & vbCrLf & vbCrLf & _
           "¿Desea generar el certificado PDF ahora?", _
           vbInformation + vbYesNo, "Inspección completada")
    Debug.Print "[PASO 16] Usuario notificado"
    
    ' ===================================================================
    ' PASO 16A: Generar certificado PDF si el usuario lo solicita
    ' ===================================================================
    If respuestaPDF = vbYes Then
        Debug.Print "[PASO 16A] Usuario solicitó generar certificado PDF..."
        On Error Resume Next
        Call CertificadoPDFGenerator.GenerarCertificadoPDF(idInspeccion)
        If Err.Number <> 0 Then
            Debug.Print "[PASO 16A] ERROR al generar PDF: " & Err.Description
            MsgBox "Error al generar el certificado PDF:" & vbCrLf & Err.Description, vbExclamation
            Err.Clear
        Else
            Debug.Print "[PASO 16A] Certificado PDF generado exitosamente"
        End If
        On Error GoTo ErrorHandler
    Else
        Debug.Print "[PASO 16A] Usuario omitió generación de certificado"
    End If
    
    ' ===================================================================
    ' PASO 17: Descargar formulario
    ' ===================================================================
    Debug.Print "[PASO 17] Descargando formulario..."
    Unload frm
    Debug.Print "[PASO 17] Formulario descargado OK"
    
    ' ===================================================================
    ' PASO 18: Refrescar cronograma resumen
    ' ===================================================================
    Debug.Print "[PASO 18] Refrescando cronograma resumen..."
    Call CronogramaResumen.RefrescarResumenCronograma
    Debug.Print "[PASO 18] Cronograma resumen refrescado OK"
    
    ' ===================================================================
    ' ÉXITO COMPLETO - Desactivar rollback
    ' ===================================================================
    rollbackNecesario = False  ' Ya no necesitamos rollback (todo se guardó exitosamente)
    
    Debug.Print "===== FIN GuardarInspeccionCompleta - ÉXITO ====="
    Exit Sub
    
ErrorHandler:
    Debug.Print "===== ERROR EN GuardarInspeccionCompleta ====="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    
    ' ====================================================================================
    ' FASE 3: ROLLBACK TRANSACCIONAL - Eliminar inspección si se creó parcialmente
    ' Fecha: 25/04/2026
    ' Propósito: Mantener integridad de datos - no dejar registros huérfanos si falla
    '            algún paso después de crear la inspección
    ' ====================================================================================
    If rollbackNecesario And Len(idInspeccion) > 0 Then
        Debug.Print "⚠️ Ejecutando ROLLBACK para inspección: " & idInspeccion
        
        On Error Resume Next  ' No fallar durante rollback
        Call InspectionRepository.EliminarInspeccion(idInspeccion)
        
        ' Registrar rollback en auditoría para trazabilidad
        Call ErrorLogger2.Log("GuardarInspeccionCompleta.Rollback", _
                              "Inspección " & idInspeccion & " eliminada por error en guardado. " & _
                              "Error original: " & Err.Description, 0)
        
        Debug.Print "✓ ROLLBACK completado - Inspección " & idInspeccion & " eliminada"
        On Error GoTo 0
    End If
    
    ' Restaurar configuración de aplicación
    Debug.Print "Restaurando configuración de Excel..."
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Registrar error original en auditoría
    Debug.Print "Registrando error en ErrorLogger..."
    Call ErrorLogger2.Log("ChecklistOrchestrator.GuardarInspeccionCompleta", Err.Description, Err.Number)
    
    ' ====================================================================================
    ' FASE 3: MENSAJES DE ERROR INTELIGENTES
    ' Propósito: Dar soluciones específicas según el tipo de error
    ' ====================================================================================
    Debug.Print "Construyendo mensaje de error al usuario..."
    Dim mensajeError As String
    mensajeError = "ERROR: La inspección NO fue guardada." & vbCrLf & vbCrLf & _
                   "Detalle técnico:" & vbCrLf & _
                   Err.Description & vbCrLf & vbCrLf
    
    ' Detectar tipo de error y agregar instrucciones específicas
    If InStr(Err.Description, "ID Criticidad") > 0 Then
        ' Error relacionado con ID Criticidad (columna faltante o dato vacío)
        mensajeError = mensajeError & _
                      "SOLUCIÓN:" & vbCrLf & _
                      "1. Verifique que la plantilla tiene la columna 'ID Criticidad'" & vbCrLf & _
                      "2. Todas las preguntas deben tener un valor (Crítica, Mayor, Menor, Ninguna)" & vbCrLf & _
                      "3. Consulte docs/INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md" & vbCrLf & _
                      "4. Contacte al administrador si el problema persiste"
    ElseIf InStr(Err.Description, "columna") > 0 Or InStr(Err.Description, "column") > 0 Then
        ' Error relacionado con estructura de tabla
        mensajeError = mensajeError & _
                      "SOLUCIÓN:" & vbCrLf & _
                      "1. La estructura de las tablas puede estar corrupta" & vbCrLf & _
                      "2. Ejecute: PlantillaCertificadoSetup.InicializarTablasRequeridas()" & vbCrLf & _
                      "3. Contacte al administrador si el problema persiste"
    Else
        ' Error genérico
        mensajeError = mensajeError & _
                      "Por favor, intente nuevamente o contacte al administrador."
    End If
    
    Debug.Print "Mostrando mensaje de error al usuario..."
    MsgBox mensajeError, vbCritical, "Error al Guardar Inspección"
    
    Debug.Print "===== FIN GuardarInspeccionCompleta - ERROR ====="
End Sub

' ══════════════════════════════════════════════════════════════════════
' FUNCIÓN: CalcularMetricasInspeccion
' ══════════════════════════════════════════════════════════════════════
' Descripción:
'   Pipeline de decisión para cálculo de RPN y Categoría.
'   Determina si usar flujo actual (1ra inspección) o flujo recurrente (2da+).
'
' Parámetros:
'   taData (Object)              - Dictionary con datos de scoring TA
'   esRecurrente (Boolean)       - Flag si es inspección recurrente
'   numeroInspeccion (Long)      - Número de inspección (1, 2, 3...)
'   rpnAnterior (Double)         - RPN de inspección anterior (0 si no aplica)
'   idInspeccionAnterior (String)- UUID de inspección anterior (vacío si manual)
'   puestoEvaluado (String)      - Puesto específico de esta inspección
'   modoRPN (String)             - Modo de RPN: "AUTO", "MANUAL", "NINGUNO" (BUGFIX 17/06/2026)
'
' Retorna:
'   (Object) Dictionary con métricas calculadas:
'     - NumeroInspeccion (Long)
'     - EsInspeccionRecurrente (Boolean)
'     - PuestoEvaluado (String)
'     - RPN_TA (Double) - RPN base calculado desde TA
'     - RPN_Final (Double) - RPN final para categorización
'     - Categoria (Long) - Número de categoría (1-5)
'     - RPN_Anterior (Double) - Solo si recurrente
'     - RPN_Promedio (Double) - Solo si recurrente
'     - RPN_Total (Double) - Solo si recurrente
'     - IDInspeccionAnterior (String) - Solo si recurrente con auto
'
' Fecha creación: 21/04/2026
' ══════════════════════════════════════════════════════════════════════
Private Function CalcularMetricasInspeccion( _
    ByVal taData As Object, _
    ByVal esRecurrente As Boolean, _
    ByVal numeroInspeccion As Long, _
    ByVal rpnAnterior As Double, _
    ByVal idInspeccionAnterior As String, _
    ByVal puestoEvaluado As String, _
    ByVal iniciales As String, _
    ByVal idPlantilla As String, _
    Optional ByVal porcRecuperacion As Double = 0, _
    Optional ByVal porcOOL As Double = 0, _
    Optional ByVal modoRPN As String = "" _
) As Object
    
    On Error GoTo ErrorHandler
    
    Dim metricas As Object
    Set metricas = CreateObject("Scripting.Dictionary")
    
    ' ─────────────────────────────────────────────────────────────────
    ' PASO 1: Calcular RPN TA (SIEMPRE es la base)
    ' ─────────────────────────────────────────────────────────────────
    Dim rpnTA As Double
    rpnTA = InspectionCalculator.CalcularRPN(taData)
    metricas("RPN_TA") = rpnTA
    
    Debug.Print "[CalcularMetricas] RPN TA calculado: " & Format(rpnTA, "0.00")
    
    ' ─────────────────────────────────────────────────────────────────
    ' PASO 2: Decidir flujo según tipo de inspección
    ' ─────────────────────────────────────────────────────────────────
    If Not esRecurrente Or numeroInspeccion = 1 Then
        ' ═══════════════════════════════════════════════════════════
        ' FLUJO ACTUAL (1ra inspección)
        ' ═══════════════════════════════════════════════════════════
        Debug.Print "[CalcularMetricas] FLUJO: Primera inspección"
        
        metricas("NumeroInspeccion") = 1
        metricas("EsInspeccionRecurrente") = False
        metricas("PuestoEvaluado") = puestoEvaluado
        metricas("RPN_Final") = rpnTA
        metricas("RPN_Total") = rpnTA  ' NUEVO: En primera inspección, RPN Total = RPN TA
        ' CATEGORÍA 5 (16/06/2026): Pasar iniciales reales para evaluación de historial
        metricas("Categoria") = InspectionCalculator.DeterminarCategoria(rpnTA, iniciales, idPlantilla)
        
        Debug.Print "[CalcularMetricas] RPN Final: " & Format(rpnTA, "0.00")
        Debug.Print "[CalcularMetricas] RPN Total: " & Format(rpnTA, "0.00") & " (= RPN TA en primera inspección)"
        Debug.Print "[CalcularMetricas] Categoría: " & metricas("Categoria")
        
    Else
        ' ═══════════════════════════════════════════════════════════
        ' FLUJO NUEVO (2da+ inspección)
        ' ═══════════════════════════════════════════════════════════
        Debug.Print "[CalcularMetricas] FLUJO: Inspección recurrente #" & numeroInspeccion
        
        ' ─────────────────────────────────────────────────────────────
        ' VALIDACIÓN: Verificar que hay datos de inspección anterior
        ' ─────────────────────────────────────────────────────────────
        ' Dos escenarios válidos:
        ' 1. Recurrente automático: IDInspeccionAnterior existe (historial real en BD)
        ' 2. Recurrente manual: modoRPN = "MANUAL" (usuario ingresó %TA anterior)
        ' BUGFIX (17/06/2026): Se usaba rpnAnterior > 0 para detectar modo manual,
        '   pero %TA anterior = 0 es válido (inspección anterior perfecta sin hallazgos)
        ' ─────────────────────────────────────────────────────────────
        Debug.Print "[CalcularMetricas] Validando historial cargado..."
        Debug.Print "[CalcularMetricas]   IDInspeccionAnterior: '" & idInspeccionAnterior & "'"
        Debug.Print "[CalcularMetricas]   rpnAnterior: " & rpnAnterior
        Debug.Print "[CalcularMetricas]   modoRPN: '" & modoRPN & "'"
        
        Dim tieneHistorialAutomatico As Boolean
        Dim tieneRPNManual As Boolean
        
        tieneHistorialAutomatico = (Len(Trim(idInspeccionAnterior)) > 0)
        tieneRPNManual = (StrComp(modoRPN, "MANUAL", vbTextCompare) = 0)
        
        If Not tieneHistorialAutomatico And Not tieneRPNManual Then
            ' Error: No hay historial automático NI RPN manual
            Debug.Print "[CalcularMetricas] ERROR: Sin datos de inspección anterior"
            Debug.Print "[CalcularMetricas] No hay ID de historial ni RPN manual ingresado"
            
            Call ErrorLogger2.Log("ChecklistOrchestrator.CalcularMetricasInspeccion", _
                "Inspección recurrente #" & numeroInspeccion & " sin datos anteriores. " & _
                "IDInspeccionAnterior='' y RPNAnterior=0", 1001)
            
            MsgBox "ERROR: Falta información de la inspección anterior." & vbCrLf & vbCrLf & _
                   "Esta es la inspección #" & numeroInspeccion & " del puesto '" & puestoEvaluado & "'." & vbCrLf & _
                   "El sistema requiere uno de los siguientes:" & vbCrLf & vbCrLf & _
                   "1. Historial automático (si existe inspección anterior en BD), O" & vbCrLf & _
                   "2. RPN anterior ingresado manualmente (si no hay historial)" & vbCrLf & vbCrLf & _
                   "Por favor active el modo recurrente e ingrese el RPN anterior manualmente.", _
                   vbCritical, "Datos Anteriores Faltantes"
            
            Err.Raise 1001, "ChecklistOrchestrator.CalcularMetricasInspeccion", _
                "Inspección recurrente sin datos de inspección anterior"
        End If
        
        If tieneHistorialAutomatico Then
            Debug.Print "[CalcularMetricas] Modo: RECURRENTE AUTOMÁTICO (historial de BD)"
        Else
            Debug.Print "[CalcularMetricas] Modo: RECURRENTE MANUAL (RPN ingresado manualmente)"
        End If
        
        Debug.Print "[CalcularMetricas] Validación OK - historial anterior cargado correctamente"
        Debug.Print "[CalcularMetricas] RPN Anterior: " & Format(rpnAnterior, "0.00") & " (puede ser 0 si desempeño perfecto)"
        metricas("NumeroInspeccion") = numeroInspeccion
        metricas("EsInspeccionRecurrente") = True
        metricas("PuestoEvaluado") = puestoEvaluado
        metricas("RPN_Anterior") = rpnAnterior
        metricas("IDInspeccionAnterior") = idInspeccionAnterior
        
        ' ─────────────────────────────────────────────────────────────
        ' Calcular RPN Promedio
        ' IMPORTANTE (23/04/2026): rpnAnterior es el % TA puro de la
        ' inspección anterior (SIN factores adicionales).
        ' - Inspección 2: Promedio = (TA1 + TA2) / 2
        ' - Inspección 3: Promedio = (TA2 + TA3) / 2
        ' - Inspección 4: Promedio = (TA3 + TA4) / 2
        ' ─────────────────────────────────────────────────────────────
        Dim rpnPromedio As Double
        Debug.Print "[CalcularMetricas] Llamando a CalcularRPNPromedio(" & rpnAnterior & ", " & rpnTA & ")"
        rpnPromedio = RecurrentInspectionCalculator.CalcularRPNPromedio(rpnAnterior, rpnTA)
        metricas("RPN_Promedio") = rpnPromedio
        
        Debug.Print "[CalcularMetricas] RPN Anterior: " & Format(rpnAnterior, "0.00")
        Debug.Print "[CalcularMetricas] RPN Promedio: " & Format(rpnPromedio, "0.00")
        
        ' Calcular RPN Total con factores adicionales (% Recuperación + % OOL)
        ' FASE 6 (23/04/2026): Ahora incluye microbiología
        Debug.Print "[CalcularMetricas] Factores adicionales - % Recuperación: " & porcRecuperacion & ", % OOL: " & porcOOL
        
        Dim rpnTotal As Double
        rpnTotal = RecurrentInspectionCalculator.CalcularRPNTotal(rpnPromedio, porcRecuperacion, porcOOL)
        metricas("RPN_Total") = rpnTotal
        metricas("RPN_Final") = rpnTotal
        metricas("PorcRecuperacion") = porcRecuperacion
        metricas("PorcOOL") = porcOOL
        
        Debug.Print "[CalcularMetricas] RPN Total: " & Format(rpnTotal, "0.00")
        
        ' Categoría basada en RPN Total (con evaluación de Categoría 5 por historial)
        ' CATEGORÍA 5 (16/06/2026): Pasar iniciales e idPlantilla para regla de 3 consecutivas
        metricas("Categoria") = RecurrentInspectionCalculator.DeterminarCategoriaRPNTotal(rpnTotal, iniciales, idPlantilla)
        
        Debug.Print "[CalcularMetricas] Categoría: " & metricas("Categoria")
        
        ' Validar consistencia (advertencia si cambio brusco)
        Call RecurrentInspectionCalculator.ValidarConsistenciaRPN(rpnAnterior, rpnTA)
    End If
    
    Set CalcularMetricasInspeccion = metricas
    Exit Function
    
ErrorHandler:
    Call ErrorLogger2.Log("ChecklistOrchestrator.CalcularMetricasInspeccion", Err.Description, Err.Number)
    Err.Raise Err.Number, "ChecklistOrchestrator.CalcularMetricasInspeccion", Err.Description
End Function
