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
    
    Dim idInspeccion As String
    Dim rollbackNecesario As Boolean
    rollbackNecesario = False
    
    ' ===================================================================
    ' PASO 1: Validar todo
    ' ===================================================================
    ' Recopilar respuestas del form antes de validar
    frm.RecopilarObservacionesPublic
    
    ' Construir diccionario de datos de cabecera para validación
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
    datosValidacion("HoraInicio") = frm.HoraInicio
    datosValidacion("HoraTermino") = frm.HoraTermino
    
    ' Construir diccionario simple de respuestas (IDPregunta → IDOpcion)
    Dim respuestasValidacion As Object
    Set respuestasValidacion = CreateObject("Scripting.Dictionary")
    Dim respConSeccion As Collection
    Set respConSeccion = frm.ObtenerRespuestasConSeccion()
    Dim rItem As Variant
    For Each rItem In respConSeccion
        Dim dItem As Object
        Set dItem = rItem
        If Len(Trim(CStr(dItem("IDOpcion")))) > 0 Then
            respuestasValidacion(dItem("IDPregunta")) = dItem("IDOpcion")
        End If
    Next rItem
    
    Dim erroresValidacion As String
    erroresValidacion = ChecklistValidator.ValidarTodo(datosValidacion, respuestasValidacion, frm.ObtenerCantidadPreguntas())
    
    If Len(erroresValidacion) > 0 Then
        MsgBox "No se puede guardar. Corrija los siguientes errores:" & vbCrLf & vbCrLf & _
               erroresValidacion, vbExclamation, "Validación"
        Exit Sub
    End If
    
    ' ===================================================================
    ' PASO 2: Confirmar con usuario
    ' ===================================================================
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Está seguro que desea guardar esta inspección?" & vbCrLf & _
                       "Evaluado: " & frm.Evaluado & vbCrLf & _
                       "Puesto: " & frm.Puesto & vbCrLf & _
                       "Fecha: " & frm.FechaInspeccion, _
                       vbQuestion + vbYesNo, "Confirmar guardado")
    
    If respuesta <> vbYes Then Exit Sub
    
    ' ===================================================================
    ' PASO 3-4: Desactivar actualizaciones para rendimiento
    ' ===================================================================
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' ===================================================================
    ' PASO 5: Construir objeto datos de inspección
    ' ===================================================================
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
    
    ' ===================================================================
    ' PASO 6: Crear registro en tblInspecciones
    ' ===================================================================
    idInspeccion = InspectionRepository.CrearInspeccion(datos)
    
    If Len(idInspeccion) = 0 Then
        Err.Raise vbObjectError + 1001, "ChecklistOrchestrator", _
                  "No se pudo crear el registro de inspección."
    End If
    
    rollbackNecesario = True
    
    ' ===================================================================
    ' PASO 7: Guardar respuestas en tblRespuestas
    ' ===================================================================
    Dim respuestas As Collection
    Set respuestas = frm.ObtenerRespuestas()
    
    Call InspectionRepository.GuardarRespuestas(idInspeccion, respuestas)
    
    ' ===================================================================
    ' PASO 8: Calcular scoring TA
    ' ===================================================================
    Dim respuestasConSeccion As Collection
    Set respuestasConSeccion = frm.ObtenerRespuestasConSeccion()
    
    Dim idSeccionTA As String
    idSeccionTA = frm.IDSeccionTA
    
    Dim taData As Object
    Set taData = InspectionCalculator.CalcularScoringTA(respuestasConSeccion, idSeccionTA)
    
    ' ===================================================================
    ' PASO 9: Calcular RPN
    ' ===================================================================
    Dim rpn As Double
    rpn = InspectionCalculator.CalcularRPN(taData)
    
    ' ===================================================================
    ' PASO 10: Determinar categoría
    ' ===================================================================
    Dim categoria As Long
    categoria = InspectionCalculator.DeterminarCategoria(rpn, frm.Evaluado, frm.IDPlantilla)
    
    ' ===================================================================
    ' PASO 11: Calcular fecha próxima
    ' ===================================================================
    Dim frecuencia As Long
    frecuencia = frm.FrecuenciaMeses
    
    Dim fechaProxima As Date
    fechaProxima = InspectionCalculator.CalcularFechaProxima(CDate(frm.FechaInspeccion), frecuencia)
    
    Dim diasVencimiento As Long
    diasVencimiento = InspectionCalculator.CalcularDiasVencimiento(fechaProxima)
    
    Dim estadoProgramacion As String
    estadoProgramacion = InspectionCalculator.DeterminarEstadoProgramacion(diasVencimiento)
    
    ' ===================================================================
    ' PASO 12: Actualizar tblInspecciones con cálculos
    ' ===================================================================
    Dim calculos As Object
    Set calculos = CreateObject("Scripting.Dictionary")
    
    calculos("TA_puntaje") = taData("puntaje")
    calculos("TA_maximos") = taData("maximos")
    calculos("TA_noaplica") = taData("noaplica")
    calculos("TA_porcentaje") = taData("porcentaje")
    calculos("RPN") = rpn
    calculos("Categoria") = "Categoría " & categoria
    calculos("RequiereAccion") = InspectionCalculator.DeterminarRequiereAccion(categoria)
    calculos("FechaProxima") = fechaProxima
    calculos("DiasVencimiento") = diasVencimiento
    calculos("EstadoProgramacion") = estadoProgramacion
    
    Call InspectionRepository.ActualizarCalculosInspeccion(idInspeccion, calculos)
    
    ' ===================================================================
    ' PASO 13: Actualizar tblCronogramaInspecciones
    ' ===================================================================
    Call InspectionScheduler.ActualizarRegistroCronograma(frm.Evaluado, frm.IDPlantilla)
    
    ' ===================================================================
    ' PASO 14: Log en Audit Trail
    ' ===================================================================
    Call AuditLogger2.LogAction( _
        "Inspección completada", _
        Configuration2.SHEET_HISTORICO, _
        "tblInspecciones / tblRespuestas", _
        "", _
        "ID: " & idInspeccion & " | Evaluado: " & frm.Evaluado & _
        " | Puesto: " & frm.Puesto & " | RPN: " & Format(rpn, "0.00") & _
        " | Cat: " & categoria, _
        "ChecklistOrchestrator.GuardarInspeccionCompleta")
    
    ' ===================================================================
    ' PASO 15-16: Restaurar configuración de aplicación
    ' ===================================================================
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' ===================================================================
    ' PASO 17: Notificar éxito
    ' ===================================================================
    MsgBox "Inspección guardada exitosamente." & vbCrLf & vbCrLf & _
           "RPN: " & Format(rpn, "0.00") & vbCrLf & _
           "Categoría: " & categoria & vbCrLf & _
           "Próxima inspección: " & Format(fechaProxima, "dd/mm/yyyy"), _
           vbInformation, "Inspección completada"
    
    ' ===================================================================
    ' PASO 18: Descargar formulario
    ' ===================================================================
    Unload frm
    
    ' ===================================================================
    ' PASO 19: Refrescar cronograma resumen
    ' ===================================================================
    Call CronogramaResumen.RefrescarResumenCronograma
    
    Exit Sub
    
ErrorHandler:
    ' --- ROLLBACK ---
    If rollbackNecesario And Len(idInspeccion) > 0 Then
        On Error Resume Next
        Call InspectionRepository.EliminarInspeccion(idInspeccion)
        On Error GoTo 0
    End If
    
    ' Restaurar configuración de aplicación
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Call ErrorLogger2.Log("ChecklistOrchestrator.GuardarInspeccionCompleta", Err.Description, Err.Number)
    MsgBox "Error al guardar la inspección: " & Err.Description & vbCrLf & _
           "Se ha realizado rollback de los datos parciales.", _
           vbCritical, "Error"
End Sub
