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
'            CronogramaButtons.
' Parámetros:
'   iniciales: Iniciales del personal a evaluar
'   idPlantilla: ID de la plantilla de inspección
'   puesto: Puesto del personal evaluado
'   idCronograma: ID del registro en tblCronogramaInspecciones
' ----------------------------------------------------------------------
Public Sub AbrirChecklistVirtual(ByVal iniciales As String, _
                                  ByVal idPlantilla As String, _
                                  ByVal puesto As String, _
                                  ByVal idCronograma As String)
    On Error GoTo ErrorHandler
    
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
    respuesta = MsgBox("¿Está seguro que desea guardar esta inspección?" & vbCrLf & _
                       "Evaluado: " & frm.Evaluado & vbCrLf & _
                       "Puesto: " & frm.Puesto & vbCrLf & _
                       "Fecha: " & Format(frm.FechaInspeccion, "dd/mm/yyyy"), _
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
    datos("FechaInspeccion") = CDate(frm.FechaInspeccion)
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
    
    If Len(idSeccionProcesos) > 0 Then
        ' Contar respuestas por criticidad
        Dim conteoProcesos As Object
        Set conteoProcesos = InspectionCalculator.ContarRespuestasPorCriticidad(respuestasConSeccion, idSeccionProcesos)
        
        Debug.Print "  Auditoría de Procesos - Conteo por criticidad:"
        Debug.Print "    Crítica - Cumple: " & conteoProcesos("Crítica_Cumple")
        Debug.Print "    Crítica - No Cumple: " & conteoProcesos("Crítica_NoCumple")
        Debug.Print "    Crítica - No Aplica: " & conteoProcesos("Crítica_NoAplica")
        Debug.Print "    Mayor - Cumple: " & conteoProcesos("Mayor_Cumple")
        Debug.Print "    Mayor - No Cumple: " & conteoProcesos("Mayor_NoCumple")
        Debug.Print "    Mayor - No Aplica: " & conteoProcesos("Mayor_NoAplica")
        Debug.Print "    Menor - Cumple: " & conteoProcesos("Menor_Cumple")
        Debug.Print "    Menor - No Cumple: " & conteoProcesos("Menor_NoCumple")
        Debug.Print "    Menor - No Aplica: " & conteoProcesos("Menor_NoAplica")
        
        ' Evaluar resultado según reglas de negocio
        resultadoProcesos = InspectionCalculator.EvaluarAuditoriaProcesos(conteoProcesos)
        Debug.Print "  Resultado Auditoría de Procesos: " & resultadoProcesos
    Else
        Debug.Print "  ADVERTENCIA: No se encontró ID de sección de Auditoría de Procesos"
    End If
    Debug.Print "[PASO 8A] Auditoría de Procesos procesada OK"
    
    ' ===================================================================
    ' PASO 9: Calcular RPN
    ' ===================================================================
    Debug.Print "[PASO 9] Calculando RPN..."
    Dim rpn As Double
    rpn = InspectionCalculator.CalcularRPN(taData)
    Debug.Print "  RPN calculado: " & Format(rpn, "0.00")
    Debug.Print "[PASO 9] RPN calculado OK"
    
    ' ===================================================================
    ' PASO 10: Determinar categoría
    ' ===================================================================
    Debug.Print "[PASO 10] Determinando categoría..."
    Dim categoria As Long
    categoria = InspectionCalculator.DeterminarCategoria(rpn, frm.Evaluado, frm.IDPlantilla)
    Debug.Print "  Categoría asignada: " & categoria
    Debug.Print "[PASO 10] Categoría determinada OK"
    
    ' ===================================================================
    ' PASO 11: Calcular fecha próxima
    ' ===================================================================
    Debug.Print "[PASO 11] Calculando fecha próxima inspección..."
    Dim frecuencia As Long
    frecuencia = frm.FrecuenciaMeses
    
    Dim fechaProxima As Date
    fechaProxima = InspectionCalculator.CalcularFechaProxima(CDate(frm.FechaInspeccion), frecuencia)
    
    Dim diasVencimiento As Long
    diasVencimiento = InspectionCalculator.CalcularDiasVencimiento(fechaProxima)
    
    Dim estadoProgramacion As String
    estadoProgramacion = InspectionCalculator.DeterminarEstadoProgramacion(diasVencimiento)
    Debug.Print "  Fecha próxima: " & Format(fechaProxima, "dd/mm/yyyy")
    Debug.Print "  Días vencimiento: " & diasVencimiento
    Debug.Print "  Estado: " & estadoProgramacion
    Debug.Print "[PASO 11] Fecha próxima calculada OK"
    
    ' ===================================================================
    ' PASO 12: Actualizar tblInspecciones con cálculos
    ' ===================================================================
    Debug.Print "[PASO 12] Actualizando inspección con cálculos..."
    Dim calculos As Object
    Set calculos = CreateObject("Scripting.Dictionary")
    
    calculos("TA_puntaje") = taData("puntaje")
    calculos("TA_maximos") = taData("maximos")
    calculos("TA_noaplica") = taData("noaplica")
    calculos("TA_porcentaje") = taData("porcentaje")
    calculos("Auditoria_Procesos_Resultado") = resultadoProcesos
    calculos("RPN") = rpn
    calculos("Categoria") = "Categoría " & categoria
    calculos("RequiereAccion") = InspectionCalculator.DeterminarRequiereAccion(categoria)
    calculos("FechaProxima") = fechaProxima
    calculos("DiasVencimiento") = diasVencimiento
    calculos("EstadoProgramacion") = estadoProgramacion
    
    Call InspectionRepository.ActualizarCalculosInspeccion(idInspeccion, calculos)
    Debug.Print "[PASO 12] Cálculos actualizados OK"
    
    ' ===================================================================
    ' PASO 13: Actualizar tblCronogramaInspecciones
    ' ===================================================================
    Debug.Print "[PASO 13] Actualizando cronograma..."
    Call InspectionScheduler.ActualizarRegistroCronograma(frm.Evaluado, frm.IDPlantilla)
    Debug.Print "[PASO 13] Cronograma actualizado OK"
    
    ' ===================================================================
    ' PASO 14: Log en Audit Trail
    ' ===================================================================
    Debug.Print "[PASO 14] Registrando en Audit Trail..."
    Call AuditLogger2.LogAction( _
        "Inspección completada", _
        Configuration2.SHEET_HISTORICO, _
        "tblInspecciones / tblRespuestas", _
        "", _
        "ID: " & idInspeccion & " | Evaluado: " & frm.Evaluado & _
        " | Puesto: " & frm.Puesto & " | RPN: " & Format(rpn, "0.00") & _
        " | Cat: " & categoria, _
        "ChecklistOrchestrator.GuardarInspeccionCompleta")
    Debug.Print "[PASO 14] Audit Trail registrado OK"
    
    ' ===================================================================
    ' PASO 15-16: Restaurar configuración de aplicación
    ' ===================================================================
    Debug.Print "[PASO 15-16] Restaurando configuración Excel..."
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Debug.Print "[PASO 15-16] Configuración restaurada OK"
    
    ' ===================================================================
    ' PASO 17: Notificar éxito y ofrecer generar certificado
    ' ===================================================================
    Debug.Print "[PASO 17] Mostrando mensaje de éxito..."
    Dim respuestaPDF As VbMsgBoxResult
    respuestaPDF = MsgBox("Inspección guardada exitosamente." & vbCrLf & _
           "RPN: " & Format(rpn, "0.00") & vbCrLf & _
           "Categoría: " & categoria & vbCrLf & _
           "Próxima inspección: " & Format(fechaProxima, "dd/mm/yyyy") & vbCrLf & vbCrLf & _
           "¿Desea generar el certificado PDF ahora?", _
           vbInformation + vbYesNo, "Inspección completada")
    Debug.Print "[PASO 17] Usuario notificado"
    
    ' ===================================================================
    ' PASO 17A: Generar certificado PDF si el usuario lo solicita
    ' ===================================================================
    If respuestaPDF = vbYes Then
        Debug.Print "[PASO 17A] Usuario solicitó generar certificado PDF..."
        On Error Resume Next
        Call CertificadoPDFGenerator.GenerarCertificadoPDF(idInspeccion)
        If Err.Number <> 0 Then
            Debug.Print "[PASO 17A] ERROR al generar PDF: " & Err.Description
            MsgBox "Error al generar el certificado PDF:" & vbCrLf & Err.Description, vbExclamation
            Err.Clear
        Else
            Debug.Print "[PASO 17A] Certificado PDF generado exitosamente"
        End If
        On Error GoTo ErrorHandler
    Else
        Debug.Print "[PASO 17A] Usuario omitió generación de certificado"
    End If
    
    ' ===================================================================
    ' PASO 18: Descargar formulario
    ' ===================================================================
    Debug.Print "[PASO 18] Descargando formulario..."
    Unload frm
    Debug.Print "[PASO 18] Formulario descargado OK"
    
    ' ===================================================================
    ' PASO 19: Refrescar cronograma resumen
    ' ===================================================================
    Debug.Print "[PASO 19] Refrescando cronograma resumen..."
    Call CronogramaResumen.RefrescarResumenCronograma
    Debug.Print "[PASO 19] Cronograma resumen refrescado OK"
    
    Debug.Print "===== FIN GuardarInspeccionCompleta - ÉXITO ====="
    Exit Sub
    
ErrorHandler:
    Debug.Print "===== ERROR EN GuardarInspeccionCompleta ====="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    
    ' --- ROLLBACK ---
    If rollbackNecesario And Len(idInspeccion) > 0 Then
        Debug.Print "Ejecutando ROLLBACK para inspección: " & idInspeccion
        On Error Resume Next
        Call InspectionRepository.EliminarInspeccion(idInspeccion)
        Debug.Print "ROLLBACK completado"
        On Error GoTo 0
    End If
    
    ' Restaurar configuración de aplicación
    Debug.Print "Restaurando configuración de Excel..."
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Debug.Print "Registrando error en ErrorLogger..."
    Call ErrorLogger2.Log("ChecklistOrchestrator.GuardarInspeccionCompleta", Err.Description, Err.Number)
    
    Debug.Print "Mostrando mensaje de error al usuario..."
    MsgBox "Error al guardar la inspección: " & Err.Description & vbCrLf & _
           "Se ha realizado rollback de los datos parciales.", _
           vbCritical, "Error"
    
    Debug.Print "===== FIN GuardarInspeccionCompleta - ERROR ====="
End Sub
