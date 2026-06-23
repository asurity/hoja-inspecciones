' ----------------------------------------------------------------------
' M\u00f3dulo: InspectionScheduler
' Descripci\u00f3n: Gestiona la sincronizaci\u00f3n y actualizaci\u00f3n del cronograma de
'              inspecciones. Mantiene tblCronogramaInspecciones sincronizada
'              con tblPersonal, tblPlantillas y tblInspecciones.
' Fecha creaci\u00f3n: 12/03/2026
' Dependencias: Configuration2, ErrorLogger2, AuditLogger2
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Subrutina: InicializarCronograma
' Prop\u00f3sito: Crea el cronograma completo desde cero, generando un registro
'            por cada combinaci\u00f3n v\u00e1lida de Persona \u00d7 Plantilla (seg\u00fan puestos activos).
' L\u00f3gica:
'   1. Limpia tblCronogramaInspecciones completamente.
'   2. Para cada persona en tblPersonal:
'      - Revisa cada columna de puesto (Quimico, Operador, Ayudante 1, etc.)
'      - Si puesto = "Si", busca plantillas que correspondan a ese puesto
'      - Crea registro en cronograma con estado inicial "Nunca inspeccionado"
'   3. Audita la operaci\u00f3n.
' Uso: Ejecutar manualmente desde bot\u00f3n o al configurar sistema por primera vez.
' ----------------------------------------------------------------------
Public Sub InicializarCronograma()
    On Error GoTo ErrorHandler
    
    Dim wsPersonal As Worksheet
    Dim wsCronograma As Worksheet
    Dim wsPlantillas As Worksheet
    Dim tblPersonal As ListObject
    Dim tblCronograma As ListObject
    Dim tblPlantillas As ListObject
    
    Dim personaRow As ListRow
    Dim plantillaRow As ListRow
    Dim nuevoRow As ListRow
    
    Dim iniciales As String
    Dim planta As String
    Dim activo As String
    Dim puestoColumna As Variant
    Dim puestosArray As Variant
    Dim puestoValor As String
    Dim puestoNombre As String
    
    Dim registrosCreados As Long
    Dim startTime As Double
    
    startTime = Timer
    registrosCreados = 0
    
    ' --- Obtener referencias a hojas y tablas ---
    Set wsPersonal = ThisWorkbook.Sheets(Configuration2.SHEET_PERSONAL)
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set wsPlantillas = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    
    Set tblPersonal = wsPersonal.ListObjects(Configuration2.TABLE_PERSONAL)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    Set tblPlantillas = wsPlantillas.ListObjects(Configuration2.TABLE_PLANTILLAS)
    
    ' --- Limpiar cronograma existente ---
    Application.ScreenUpdating = False
    
    If Not tblCronograma.DataBodyRange Is Nothing Then
        tblCronograma.DataBodyRange.Delete
    End If
    
    ' --- Obtener array de columnas de puestos ---
    puestosArray = Configuration2.GetPuestosColumns()
    
    ' --- Iterar sobre cada persona en tblPersonal ---
    If Not tblPersonal.DataBodyRange Is Nothing Then
        For Each personaRow In tblPersonal.ListRows
            iniciales = personaRow.Range.Cells(1, tblPersonal.ListColumns("Iniciales").Index).Value
            planta = personaRow.Range.Cells(1, tblPersonal.ListColumns("Planta").Index).Value
            activo = personaRow.Range.Cells(1, tblPersonal.ListColumns("Activo").Index).Value
            
            ' Iterar sobre cada puesto (columna en tblPersonal)
            For Each puestoColumna In puestosArray
                puestoNombre = CStr(puestoColumna)
                
                ' Verificar si la columna existe en tblPersonal
                On Error Resume Next
                puestoValor = personaRow.Range.Cells(1, tblPersonal.ListColumns(puestoNombre).Index).Value
                On Error GoTo ErrorHandler
                
                ' Si el puesto est\u00e1 activo (valor "Si")
                If UCase(Trim(puestoValor)) = "SI" Then
                    ' Buscar plantillas que correspondan a este puesto
                    If Not tblPlantillas.DataBodyRange Is Nothing Then
                        For Each plantillaRow In tblPlantillas.ListRows
                            Dim plantillaPuesto As String
                            Dim plantillaID As String
                            Dim plantillaNombre As String
                            Dim frecuencia As Variant
                            
                            plantillaPuesto = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("Puesto").Index).Value
                            
                            ' Si el puesto de la plantilla coincide con el puesto de la persona
                            If Trim(plantillaPuesto) = Trim(puestoNombre) Then
                                plantillaID = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("ID Plantilla").Index).Value
                                plantillaNombre = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("Nombre de plantilla").Index).Value
                                frecuencia = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("Frecuencia meses").Index).Value
                                
                                ' Crear nuevo registro en cronograma
                                Set nuevoRow = tblCronograma.ListRows.Add
                                
                                With nuevoRow.Range
                                    .Cells(1, tblCronograma.ListColumns("ID Cronograma").Index).Value = GenerarUUID()
                                    .Cells(1, tblCronograma.ListColumns("Iniciales personal").Index).Value = iniciales
                                    .Cells(1, tblCronograma.ListColumns("ID Plantilla").Index).Value = plantillaID
                                    .Cells(1, tblCronograma.ListColumns("Nombre plantilla").Index).Value = plantillaNombre
                                    .Cells(1, tblCronograma.ListColumns("Puesto").Index).Value = puestoNombre
                                    .Cells(1, tblCronograma.ListColumns("Planta personal").Index).Value = planta
                                    .Cells(1, tblCronograma.ListColumns("Frecuencia meses").Index).Value = IIf(IsNumeric(frecuencia), frecuencia, 3)
                                    
                                    ' Campos de estado inicial
                                    .Cells(1, tblCronograma.ListColumns("Total inspecciones").Index).Value = 0
                                    .Cells(1, tblCronograma.ListColumns("Estado cronograma").Index).Value = Configuration2.ESTADO_NUNCA_INSPECCIONADO
                                    
                                    ' Validaciones
                                    .Cells(1, tblCronograma.ListColumns("Puesto activo en personal").Index).Value = "Si"
                                    .Cells(1, tblCronograma.ListColumns("Personal activo").Index).Value = activo
                                    
                                    ' Auditoría
                                    .Cells(1, tblCronograma.ListColumns("Fecha ultima actualizacion").Index).Value = Now
                                    .Cells(1, tblCronograma.ListColumns("Requiere recalculo").Index).Value = "No"
                                    
                                    ' Gestión de cronograma: activa por defecto al inicializar
                                    ' (backward compatible: si la columna no existe, On Error Resume Next lo ignora)
                                    On Error Resume Next
                                    .Cells(1, tblCronograma.ListColumns("Activo en cronograma").Index).Value = "Si"
                                    Err.Clear
                                    On Error GoTo ErrorHandler
                                End With
                                
                                registrosCreados = registrosCreados + 1
                            End If
                        Next plantillaRow
                    End If
                End If
            Next puestoColumna
        Next personaRow
    End If
    
    Application.ScreenUpdating = True
    
    ' --- Auditar operaci\u00f3n ---
    Call AuditLogger2.LogAction( _
        action:="Inicializaci\u00f3n de cronograma", _
        sheetName:=Configuration2.SHEET_CRONOGRAMA, _
        dataModified:="tblCronogramaInspecciones", _
        beforeChange:="Tabla vac\u00eda", _
        afterChange:=registrosCreados & " registros creados en " & Format(Timer - startTime, "0.00") & " segundos", _
        moduleAndSubroutine:="InspectionScheduler.InicializarCronograma" _
    )
    
    MsgBox "Cronograma inicializado exitosamente." & vbCrLf & vbCrLf & _
           "Registros creados: " & registrosCreados & vbCrLf & _
           "Tiempo: " & Format(Timer - startTime, "0.00") & " segundos", _
           vbInformation, "Inicializaci\u00f3n Completa"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Call ErrorLogger2.Log("InspectionScheduler.InicializarCronograma", Err.Description, Err.Number)
    MsgBox "Error al inicializar cronograma: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: RecalcularCronograma
' Prop\u00f3sito: Actualiza todos los registros del cronograma con la informaci\u00f3n
'            m\u00e1s reciente de inspecciones completadas.
' L\u00f3gica:
'   1. Para cada registro en tblCronogramaInspecciones:
'      - Busca la \u00faltima inspecci\u00f3n completada (mismo Iniciales + ID Plantilla)
'      - Actualiza: Total inspecciones, fecha \u00faltima, RPN, categor\u00eda
'      - Calcula: Fecha pr\u00f3xima inspecci\u00f3n, d\u00edas para vencimiento
'      - Determina: Estado del cronograma (Vigente/Por vencer/Vencido)
'      - Valida: Si el puesto sigue activo en tblPersonal
'   2. Marca campo "Requiere recalculo" = "No" para optimizar futuras ejecuciones.
' Uso: Ejecutar desde bot\u00f3n "Recalcular Cronograma" o autom\u00e1ticamente al completar inspecci\u00f3n.
' ----------------------------------------------------------------------
Public Sub RecalcularCronograma()
    On Error GoTo ErrorHandler
    
    Dim wsCronograma As Worksheet
    Dim wsInspecciones As Worksheet
    Dim wsPersonal As Worksheet
    Dim tblCronograma As ListObject
    Dim tblInspecciones As ListObject
    Dim tblPersonal As ListObject
    
    Dim cronogramaRow As ListRow
    Dim registrosActualizados As Long
    Dim startTime As Double
    
    startTime = Timer
    registrosActualizados = 0
    
    ' --- Obtener referencias ---
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set wsInspecciones = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Set wsPersonal = ThisWorkbook.Sheets(Configuration2.SHEET_PERSONAL)
    
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    Set tblInspecciones = wsInspecciones.ListObjects(Configuration2.TABLE_INSPECCIONES)
    Set tblPersonal = wsPersonal.ListObjects(Configuration2.TABLE_PERSONAL)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- Iterar sobre cada registro del cronograma ---
    If Not tblCronograma.DataBodyRange Is Nothing Then
        For Each cronogramaRow In tblCronograma.ListRows
            Dim iniciales As String
            Dim idPlantilla As String
            Dim puesto As String
            
            iniciales = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Iniciales personal").Index).Value
            idPlantilla = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("ID Plantilla").Index).Value
            puesto = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Puesto").Index).Value
            
            ' Actualizar registro individual
            Call ActualizarRegistroCronogramaInterno(cronogramaRow, tblCronograma, tblInspecciones, tblPersonal, iniciales, idPlantilla, puesto)
            
            registrosActualizados = registrosActualizados + 1
        Next cronogramaRow
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' --- Auditar operaci\u00f3n ---
    Call AuditLogger2.LogAction( _
        action:="Recalcular cronograma completo", _
        sheetName:=Configuration2.SHEET_CRONOGRAMA, _
        dataModified:="tblCronogramaInspecciones", _
        beforeChange:="N/A", _
        afterChange:=registrosActualizados & " registros actualizados en " & Format(Timer - startTime, "0.00") & " segundos", _
        moduleAndSubroutine:="InspectionScheduler.RecalcularCronograma" _
    )
    
    MsgBox "Cronograma recalculado exitosamente." & vbCrLf & vbCrLf & _
           "Registros actualizados: " & registrosActualizados & vbCrLf & _
           "Tiempo: " & Format(Timer - startTime, "0.00") & " segundos", _
           vbInformation, "C\u00e1lculo Completo"
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call ErrorLogger2.Log("InspectionScheduler.RecalcularCronograma", Err.Description, Err.Number)
    MsgBox "Error al recalcular cronograma: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: ActualizarRegistroCronograma
' Prop\u00f3sito: Actualiza un registro espec\u00edfico del cronograma despu\u00e9s de
'            completar o modificar una inspecci\u00f3n.
' Par\u00e1metros:
'   - iniciales: Iniciales del personal (FK a tblPersonal)
'   - idPlantilla: ID de la plantilla inspeccionada (FK a tblPlantillas)
' L\u00f3gica: Similar a RecalcularCronograma pero solo para un registro espec\u00edfico.
' Uso: Llamar autom\u00e1ticamente al completar una inspecci\u00f3n desde InspectionCore.
' ----------------------------------------------------------------------
Public Sub ActualizarRegistroCronograma(ByVal iniciales As String, ByVal idPlantilla As String)
    On Error GoTo ErrorHandler
    
    Dim wsCronograma As Worksheet
    Dim wsInspecciones As Worksheet
    Dim wsPersonal As Worksheet
    Dim tblCronograma As ListObject
    Dim tblInspecciones As ListObject
    Dim tblPersonal As ListObject
    
    Dim cronogramaRow As ListRow
    Dim encontrado As Boolean
    
    ' --- Obtener referencias ---
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set wsInspecciones = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Set wsPersonal = ThisWorkbook.Sheets(Configuration2.SHEET_PERSONAL)
    
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    Set tblInspecciones = wsInspecciones.ListObjects(Configuration2.TABLE_INSPECCIONES)
    Set tblPersonal = wsPersonal.ListObjects(Configuration2.TABLE_PERSONAL)
    
    ' FASE 9 (09/06/2026): Desproteger hoja Cronograma para permitir escritura VBA
    Call SheetProtector2.UnprotectSheet(wsCronograma, Configuration2.APP_PASSWORD)
    
    encontrado = False
    
    ' --- Buscar el registro espec\u00edfico en cronograma ---
    If Not tblCronograma.DataBodyRange Is Nothing Then
        For Each cronogramaRow In tblCronograma.ListRows
            Dim currentIniciales As String
            Dim currentIDPlantilla As String
            Dim currentPuesto As String
            
            currentIniciales = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Iniciales personal").Index).Value
            currentIDPlantilla = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("ID Plantilla").Index).Value
            
            If Trim(currentIniciales) = Trim(iniciales) And Trim(currentIDPlantilla) = Trim(idPlantilla) Then
                currentPuesto = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Puesto").Index).Value
                
                ' Actualizar este registro
                Call ActualizarRegistroCronogramaInterno(cronogramaRow, tblCronograma, tblInspecciones, tblPersonal, iniciales, idPlantilla, currentPuesto)
                
                encontrado = True
                Exit For
            End If
        Next cronogramaRow
    End If
    
    If Not encontrado Then
        ' Si no existe el registro en cronograma, crear uno nuevo
        Call CrearRegistroCronograma(iniciales, idPlantilla)
    End If
    
    ' Reproteger hoja según el rol del usuario
    Call SheetProtector2.ApplyRoleBasedProtection(wsCronograma, Configuration2.APP_PASSWORD)
    Exit Sub
    
ErrorHandler:
    ' Reproteger hoja incluso si hay error (fail-safe)
    On Error Resume Next
    Call SheetProtector2.ApplyRoleBasedProtection(wsCronograma, Configuration2.APP_PASSWORD)
    On Error GoTo 0
    Call ErrorLogger2.Log("InspectionScheduler.ActualizarRegistroCronograma", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Subrutina INTERNA: ActualizarRegistroCronogramaInterno
' Prop\u00f3sito: L\u00f3gica compartida de actualizaci\u00f3n de un registro de cronograma.
'            Busca la \u00faltima inspecci\u00f3n y calcula todos los campos derivados.
' ----------------------------------------------------------------------
Private Sub ActualizarRegistroCronogramaInterno( _
    ByRef cronogramaRow As ListRow, _
    ByRef tblCronograma As ListObject, _
    ByRef tblInspecciones As ListObject, _
    ByRef tblPersonal As ListObject, _
    ByVal iniciales As String, _
    ByVal idPlantilla As String, _
    ByVal puesto As String)
    
    On Error GoTo ErrorHandler
    
    Dim inspeccionRow As ListRow
    Dim ultimaFecha As Date
    Dim ultimoRPN As Double
    Dim ultimaCategoria As String
    Dim ultimoID As String
    Dim totalInspecciones As Long
    Dim encontradaInspeccion As Boolean
    Dim frecuenciaMeses As Long
    Dim fechaProxima As Date
    Dim diasVencimiento As Long
    Dim estadoCronograma As String
    Dim diasAlerta As Long
    
    encontradaInspeccion = False
    totalInspecciones = 0
    ultimaFecha = DateSerial(1900, 1, 1)
    
    ' --- Buscar todas las inspecciones completadas para este Persona+Plantilla ---
    If Not tblInspecciones.DataBodyRange Is Nothing Then
        For Each inspeccionRow In tblInspecciones.ListRows
            Dim inspIniciales As String
            Dim inspPlantilla As String
            Dim inspEstado As String
            Dim inspFecha As Date
            
            inspIniciales = inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("Iniciales personal").Index).Value
            inspPlantilla = inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("ID Plantilla").Index).Value
            inspEstado = inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("Estado").Index).Value
            
            ' Filtrar por Iniciales + ID Plantilla + Estado Completado
            If Trim(inspIniciales) = Trim(iniciales) And _
               Trim(inspPlantilla) = Trim(idPlantilla) And _
               Trim(inspEstado) = Configuration2.INSPECCION_COMPLETADO Then
                
                totalInspecciones = totalInspecciones + 1
                
                ' Obtener fecha de inspecci\u00f3n
                On Error Resume Next
                inspFecha = inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("Fecha inspeccion").Index).Value
                On Error GoTo ErrorHandler
                
                ' Si esta es la inspecci\u00f3n m\u00e1s reciente
                If inspFecha > ultimaFecha Then
                    ultimaFecha = inspFecha
                    ultimoRPN = inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("RPN calculado").Index).Value
                    ultimaCategoria = inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("Categoria resultado").Index).Value
                    ultimoID = inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("ID Inspeccion").Index).Value
                    encontradaInspeccion = True
                End If
            End If
        Next inspeccionRow
    End If
    
    ' --- Obtener frecuencia y d\u00edas de alerta ---
    frecuenciaMeses = cronogramaRow.Range.Cells(1, tblCronograma.ListColumns("Frecuencia meses").Index).Value
    If frecuenciaMeses = 0 Then frecuenciaMeses = 3 ' Default
    
    diasAlerta = ObtenerParametroNumerico(PARAM_DIAS_ALERTA_VENCIMIENTO, 15)
    
    ' --- Calcular fecha pr\u00f3xima y d\u00edas para vencimiento ---
    If encontradaInspeccion Then
        fechaProxima = DateAdd("m", frecuenciaMeses, ultimaFecha)
        diasVencimiento = DateDiff("d", Date, fechaProxima)
        
        ' Determinar estado del cronograma
        If diasVencimiento < 0 Then
            estadoCronograma = Configuration2.ESTADO_VENCIDO
        ElseIf diasVencimiento <= diasAlerta Then
            estadoCronograma = Configuration2.ESTADO_POR_VENCER
        Else
            estadoCronograma = Configuration2.ESTADO_VIGENTE
        End If
    Else
        ' Nunca inspeccionado
        estadoCronograma = Configuration2.ESTADO_NUNCA_INSPECCIONADO
        diasVencimiento = -9999 ' Indicador de vencimiento cr\u00edtico
    End If
    
    ' --- Validar si el puesto sigue activo en tblPersonal ---
    Dim puestoActivoEnPersonal As String
    puestoActivoEnPersonal = ValidarPuestoActivo(tblPersonal, iniciales, puesto)
    
    ' Si el puesto ya no est\u00e1 activo, cambiar estado
    If puestoActivoEnPersonal = "No" Then
        estadoCronograma = Configuration2.ESTADO_PUESTO_INACTIVO
    End If
    
    ' --- Validar si el personal sigue activo ---
    Dim personalActivo As String
    personalActivo = ObtenerValorPersonal(tblPersonal, iniciales, "Activo")
    
    ' ========================================================================
    ' FASE 10 (17/06/2026): CAPTURAR valores ANTERIORES antes de sobreescribir.
    ' Necesario porque el guardado de inspección deshabilita eventos
    ' (Application.EnableEvents = False) y el Worksheet_Change de la Hoja8
    ' jamás se entera de estas escrituras. Auditamos directamente aquí, igual
    ' que PausarInspecciones/ReactivarInspecciones en CronogramaGestorService.
    ' ========================================================================
    Dim oldTotalInsp As Variant
    Dim oldFechaUltima As Variant
    Dim oldIDUltima As Variant
    Dim oldRPNUltima As Variant
    Dim oldCatUltima As Variant
    Dim oldFechaProx As Variant
    Dim oldDiasVenc As Variant
    Dim oldEstado As Variant
    Dim oldPuestoActivo As Variant
    Dim oldPersActivo As Variant
    Dim oldFechaAct As Variant
    Dim oldReqRecalc As Variant
    
    With cronogramaRow.Range
        oldTotalInsp = .Cells(1, tblCronograma.ListColumns("Total inspecciones").Index).Value
        oldFechaUltima = .Cells(1, tblCronograma.ListColumns("Fecha ultima inspeccion").Index).Value
        oldIDUltima = .Cells(1, tblCronograma.ListColumns("ID Ultima inspeccion").Index).Value
        oldRPNUltima = .Cells(1, tblCronograma.ListColumns("RPN ultima inspeccion").Index).Value
        oldCatUltima = .Cells(1, tblCronograma.ListColumns("Categoria ultima inspeccion").Index).Value
        oldFechaProx = .Cells(1, tblCronograma.ListColumns("Fecha proxima inspeccion").Index).Value
        oldDiasVenc = .Cells(1, tblCronograma.ListColumns("Dias para vencimiento").Index).Value
        oldEstado = .Cells(1, tblCronograma.ListColumns("Estado cronograma").Index).Value
        oldPuestoActivo = .Cells(1, tblCronograma.ListColumns("Puesto activo en personal").Index).Value
        oldPersActivo = .Cells(1, tblCronograma.ListColumns("Personal activo").Index).Value
        oldFechaAct = .Cells(1, tblCronograma.ListColumns("Fecha ultima actualizacion").Index).Value
        oldReqRecalc = .Cells(1, tblCronograma.ListColumns("Requiere recalculo").Index).Value
    End With
    
    ' --- Actualizar campos del cronograma ---
    With cronogramaRow.Range
        .Cells(1, tblCronograma.ListColumns("Total inspecciones").Index).Value = totalInspecciones
        
        If encontradaInspeccion Then
            .Cells(1, tblCronograma.ListColumns("Fecha ultima inspeccion").Index).Value = ultimaFecha
            .Cells(1, tblCronograma.ListColumns("ID Ultima inspeccion").Index).Value = ultimoID
            .Cells(1, tblCronograma.ListColumns("RPN ultima inspeccion").Index).Value = ultimoRPN
            .Cells(1, tblCronograma.ListColumns("Categoria ultima inspeccion").Index).Value = ultimaCategoria
            .Cells(1, tblCronograma.ListColumns("Fecha proxima inspeccion").Index).Value = fechaProxima
            .Cells(1, tblCronograma.ListColumns("Dias para vencimiento").Index).Value = diasVencimiento
        Else
            ' Limpiar campos si nunca ha sido inspeccionado
            .Cells(1, tblCronograma.ListColumns("Fecha ultima inspeccion").Index).Value = ""
            .Cells(1, tblCronograma.ListColumns("ID Ultima inspeccion").Index).Value = ""
            .Cells(1, tblCronograma.ListColumns("RPN ultima inspeccion").Index).Value = ""
            .Cells(1, tblCronograma.ListColumns("Categoria ultima inspeccion").Index).Value = ""
            .Cells(1, tblCronograma.ListColumns("Fecha proxima inspeccion").Index).Value = ""
            .Cells(1, tblCronograma.ListColumns("Dias para vencimiento").Index).Value = diasVencimiento
        End If
        
        .Cells(1, tblCronograma.ListColumns("Estado cronograma").Index).Value = estadoCronograma
        .Cells(1, tblCronograma.ListColumns("Puesto activo en personal").Index).Value = puestoActivoEnPersonal
        .Cells(1, tblCronograma.ListColumns("Personal activo").Index).Value = personalActivo
        .Cells(1, tblCronograma.ListColumns("Fecha ultima actualizacion").Index).Value = Now
        .Cells(1, tblCronograma.ListColumns("Requiere recalculo").Index).Value = "No"
    End With
    
    ' ========================================================================
    ' FASE 10 (17/06/2026): AUDITAR cambios en cronograma.
    ' Se compara cada campo viejo vs nuevo y se registra solo si hubo cambios.
    ' Esto cubre el hueco dejado por Application.EnableEvents = False durante
    ' el guardado programático de inspecciones (ChecklistOrchestrator).
    ' ========================================================================
    Dim changesBefore As String
    Dim changesAfter As String
    Dim changeCount As Long
    changeCount = 0
    
    If CStr(oldTotalInsp) <> CStr(totalInspecciones) Then
        changesBefore = changesBefore & "Total inspecciones: " & AuditValor(oldTotalInsp) & vbCrLf
        changesAfter = changesAfter & "Total inspecciones: " & AuditValor(totalInspecciones) & vbCrLf
        changeCount = changeCount + 1
    End If
    
    If encontradaInspeccion Then
        If CStr(oldFechaUltima) <> CStr(ultimaFecha) Then
            changesBefore = changesBefore & "Fecha ultima: " & AuditValor(oldFechaUltima) & vbCrLf
            changesAfter = changesAfter & "Fecha ultima: " & AuditValor(ultimaFecha) & vbCrLf
            changeCount = changeCount + 1
        End If
        If CStr(oldIDUltima) <> CStr(ultimoID) Then
            changesBefore = changesBefore & "ID Ultima: " & AuditValor(oldIDUltima) & vbCrLf
            changesAfter = changesAfter & "ID Ultima: " & AuditValor(ultimoID) & vbCrLf
            changeCount = changeCount + 1
        End If
        If CStr(oldRPNUltima) <> CStr(ultimoRPN) Then
            changesBefore = changesBefore & "RPN ultima: " & AuditValor(oldRPNUltima) & vbCrLf
            changesAfter = changesAfter & "RPN ultima: " & AuditValor(ultimoRPN) & vbCrLf
            changeCount = changeCount + 1
        End If
        If CStr(oldCatUltima) <> CStr(ultimaCategoria) Then
            changesBefore = changesBefore & "Categoria ultima: " & AuditValor(oldCatUltima) & vbCrLf
            changesAfter = changesAfter & "Categoria ultima: " & AuditValor(ultimaCategoria) & vbCrLf
            changeCount = changeCount + 1
        End If
        If CStr(oldFechaProx) <> CStr(fechaProxima) Then
            changesBefore = changesBefore & "Fecha proxima: " & AuditValor(oldFechaProx) & vbCrLf
            changesAfter = changesAfter & "Fecha proxima: " & AuditValor(fechaProxima) & vbCrLf
            changeCount = changeCount + 1
        End If
        If CStr(oldDiasVenc) <> CStr(diasVencimiento) Then
            changesBefore = changesBefore & "Dias vencimiento: " & AuditValor(oldDiasVenc) & vbCrLf
            changesAfter = changesAfter & "Dias vencimiento: " & AuditValor(diasVencimiento) & vbCrLf
            changeCount = changeCount + 1
        End If
    Else
        ' Nunca inspeccionado: limpiar campos solo si antes tenían valor
        If Len(CStr(oldFechaUltima)) > 0 Then
            changesBefore = changesBefore & "Fecha ultima: " & AuditValor(oldFechaUltima) & vbCrLf
            changesAfter = changesAfter & "Fecha ultima: (limpiado)" & vbCrLf
            changeCount = changeCount + 1
        End If
        If Len(CStr(oldIDUltima)) > 0 Then
            changesBefore = changesBefore & "ID Ultima: " & AuditValor(oldIDUltima) & vbCrLf
            changesAfter = changesAfter & "ID Ultima: (limpiado)" & vbCrLf
            changeCount = changeCount + 1
        End If
        If Len(CStr(oldRPNUltima)) > 0 Then
            changesBefore = changesBefore & "RPN ultima: " & AuditValor(oldRPNUltima) & vbCrLf
            changesAfter = changesAfter & "RPN ultima: (limpiado)" & vbCrLf
            changeCount = changeCount + 1
        End If
        If Len(CStr(oldCatUltima)) > 0 Then
            changesBefore = changesBefore & "Categoria ultima: " & AuditValor(oldCatUltima) & vbCrLf
            changesAfter = changesAfter & "Categoria ultima: (limpiado)" & vbCrLf
            changeCount = changeCount + 1
        End If
        If Len(CStr(oldFechaProx)) > 0 Then
            changesBefore = changesBefore & "Fecha proxima: " & AuditValor(oldFechaProx) & vbCrLf
            changesAfter = changesAfter & "Fecha proxima: (limpiado)" & vbCrLf
            changeCount = changeCount + 1
        End If
        If CStr(oldDiasVenc) <> CStr(diasVencimiento) Then
            changesBefore = changesBefore & "Dias vencimiento: " & AuditValor(oldDiasVenc) & vbCrLf
            changesAfter = changesAfter & "Dias vencimiento: " & AuditValor(diasVencimiento) & vbCrLf
            changeCount = changeCount + 1
        End If
    End If
    
    If CStr(oldEstado) <> CStr(estadoCronograma) Then
        changesBefore = changesBefore & "Estado: " & AuditValor(oldEstado) & vbCrLf
        changesAfter = changesAfter & "Estado: " & AuditValor(estadoCronograma) & vbCrLf
        changeCount = changeCount + 1
    End If
    If CStr(oldPuestoActivo) <> CStr(puestoActivoEnPersonal) Then
        changesBefore = changesBefore & "Puesto activo: " & AuditValor(oldPuestoActivo) & vbCrLf
        changesAfter = changesAfter & "Puesto activo: " & AuditValor(puestoActivoEnPersonal) & vbCrLf
        changeCount = changeCount + 1
    End If
    If CStr(oldPersActivo) <> CStr(personalActivo) Then
        changesBefore = changesBefore & "Personal activo: " & AuditValor(oldPersActivo) & vbCrLf
        changesAfter = changesAfter & "Personal activo: " & AuditValor(personalActivo) & vbCrLf
        changeCount = changeCount + 1
    End If
    
    If changeCount > 0 Then
        Call AuditLogger2.LogAction( _
            "Modificación en cronograma (macro)", _
            Configuration2.SHEET_CRONOGRAMA, _
            "tblCronogramaInspecciones | " & iniciales & " / " & idPlantilla & " (" & puesto & ")", _
            changesBefore, _
            changesAfter, _
            "InspectionScheduler.ActualizarRegistroCronogramaInterno")
    End If
    
    Exit Sub
    
ErrorHandler:
    ' No lanzar error, solo registrar y continuar con siguiente registro
    Call ErrorLogger2.Log("InspectionScheduler.ActualizarRegistroCronogramaInterno", "Error en " & iniciales & " - " & puesto & ": " & Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Funci\u00f3n: ValidarPuestoActivo
' Prop\u00f3sito: Verifica si un puesto espec\u00edfico sigue activo (valor "Si") en
'            tblPersonal para una persona dada.
' Retorna: "Si" si el puesto est\u00e1 activo, "No" si est\u00e1 inactivo o no existe.
' ----------------------------------------------------------------------
Private Function ValidarPuestoActivo(ByRef tblPersonal As ListObject, ByVal iniciales As String, ByVal puesto As String) As String
    On Error GoTo ErrorHandler
    
    Dim personaRow As ListRow
    
    If Not tblPersonal.DataBodyRange Is Nothing Then
        For Each personaRow In tblPersonal.ListRows
            Dim currentIniciales As String
            currentIniciales = personaRow.Range.Cells(1, tblPersonal.ListColumns("Iniciales").Index).Value
            
            If Trim(currentIniciales) = Trim(iniciales) Then
                ' Buscar la columna del puesto
                On Error Resume Next
                Dim puestoValor As String
                puestoValor = personaRow.Range.Cells(1, tblPersonal.ListColumns(puesto).Index).Value
                On Error GoTo ErrorHandler
                
                If UCase(Trim(puestoValor)) = "SI" Then
                    ValidarPuestoActivo = "Si"
                Else
                    ValidarPuestoActivo = "No"
                End If
                
                Exit Function
            End If
        Next personaRow
    End If
    
    ' Si no se encontr\u00f3 la persona o el puesto, devolver "No"
    ValidarPuestoActivo = "No"
    Exit Function
    
ErrorHandler:
    ValidarPuestoActivo = "No"
End Function

'' ----------------------------------------------------------------------
' Funci\u00f3n: ObtenerValorPersonal
' Prop\u00f3sito: Obtiene el valor de una columna espec\u00edfica de tblPersonal
'            para una persona dada (por iniciales).
' ----------------------------------------------------------------------
Private Function ObtenerValorPersonal(ByRef tblPersonal As ListObject, ByVal iniciales As String, ByVal nombreColumna As String) As String
    On Error GoTo ErrorHandler
    
    Dim personaRow As ListRow
    
    If Not tblPersonal.DataBodyRange Is Nothing Then
        For Each personaRow In tblPersonal.ListRows
            Dim currentIniciales As String
            currentIniciales = personaRow.Range.Cells(1, tblPersonal.ListColumns("Iniciales").Index).Value
            
            If Trim(currentIniciales) = Trim(iniciales) Then
                On Error Resume Next
                ObtenerValorPersonal = personaRow.Range.Cells(1, tblPersonal.ListColumns(nombreColumna).Index).Value
                On Error GoTo ErrorHandler
                Exit Function
            End If
        Next personaRow
    End If
    
    ObtenerValorPersonal = ""
    Exit Function
    
ErrorHandler:
    ObtenerValorPersonal = ""
End Function

'' ----------------------------------------------------------------------
' Funci\u00f3n: ObtenerParametroNumerico
' Prop\u00f3sito: Obtiene un par\u00e1metro num\u00e9rico de tblConfiguracion.
' Retorna: El valor configurado o un valor por defecto si no existe.
' ----------------------------------------------------------------------
Private Function ObtenerParametroNumerico(ByVal clave As String, ByVal valorDefault As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim tblConfig As ListObject
    Dim configRow As ListRow
    
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tblConfig = wsConfig.ListObjects(Configuration2.TABLE_CONFIGURACION)
    
    If Not tblConfig.DataBodyRange Is Nothing Then
        For Each configRow In tblConfig.ListRows
            Dim currentClave As String
            currentClave = configRow.Range.Cells(1, tblConfig.ListColumns("Clave").Index).Value
            
            If Trim(currentClave) = Trim(clave) Then
                Dim valor As String
                valor = configRow.Range.Cells(1, tblConfig.ListColumns("Valor").Index).Value
                
                If IsNumeric(valor) Then
                    ObtenerParametroNumerico = CLng(valor)
                Else
                    ObtenerParametroNumerico = valorDefault
                End If
                
                Exit Function
            End If
        Next configRow
    End If
    
    ' Si no se encontr\u00f3, devolver valor por defecto
    ObtenerParametroNumerico = valorDefault
    Exit Function
    
ErrorHandler:
    ObtenerParametroNumerico = valorDefault
End Function

'' ----------------------------------------------------------------------
' Subrutina: CrearRegistroCronograma
' Prop\u00f3sito: Crea un nuevo registro en el cronograma si no existe.
' Uso: Llamado cuando se completa una inspecci\u00f3n para una combinaci\u00f3n
'      Persona+Plantilla que a\u00fan no existe en el cronograma.
' ----------------------------------------------------------------------
Private Sub CrearRegistroCronograma(ByVal iniciales As String, ByVal idPlantilla As String)
    On Error GoTo ErrorHandler
    
    Dim wsCronograma As Worksheet
    Dim wsPlantillas As Worksheet
    Dim wsPersonal As Worksheet
    Dim tblCronograma As ListObject
    Dim tblPlantillas As ListObject
    Dim tblPersonal As ListObject
    
    Dim nuevoRow As ListRow
    Dim plantillaRow As ListRow
    Dim plantillaNombre As String
    Dim plantillaPuesto As String
    Dim frecuencia As Variant
    Dim planta As String
    Dim activo As String
    
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set wsPlantillas = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set wsPersonal = ThisWorkbook.Sheets(Configuration2.SHEET_PERSONAL)
    
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    Set tblPlantillas = wsPlantillas.ListObjects(Configuration2.TABLE_PLANTILLAS)
    Set tblPersonal = wsPersonal.ListObjects(Configuration2.TABLE_PERSONAL)
    
    ' Buscar datos de la plantilla
    If Not tblPlantillas.DataBodyRange Is Nothing Then
        For Each plantillaRow In tblPlantillas.ListRows
            If Trim(plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("ID Plantilla").Index).Value) = Trim(idPlantilla) Then
                plantillaNombre = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("Nombre de plantilla").Index).Value
                plantillaPuesto = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("Puesto").Index).Value
                frecuencia = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("Frecuencia meses").Index).Value
                Exit For
            End If
        Next plantillaRow
    End If
    
    ' Obtener datos del personal
    planta = ObtenerValorPersonal(tblPersonal, iniciales, "Planta")
    activo = ObtenerValorPersonal(tblPersonal, iniciales, "Activo")
    
    ' Crear nuevo registro
    Set nuevoRow = tblCronograma.ListRows.Add
    
    With nuevoRow.Range
        .Cells(1, tblCronograma.ListColumns("ID Cronograma").Index).Value = GenerarUUID()
        .Cells(1, tblCronograma.ListColumns("Iniciales personal").Index).Value = iniciales
        .Cells(1, tblCronograma.ListColumns("ID Plantilla").Index).Value = idPlantilla
        .Cells(1, tblCronograma.ListColumns("Nombre plantilla").Index).Value = plantillaNombre
        .Cells(1, tblCronograma.ListColumns("Puesto").Index).Value = plantillaPuesto
        .Cells(1, tblCronograma.ListColumns("Planta personal").Index).Value = planta
        .Cells(1, tblCronograma.ListColumns("Frecuencia meses").Index).Value = IIf(IsNumeric(frecuencia), frecuencia, 3)
        .Cells(1, tblCronograma.ListColumns("Total inspecciones").Index).Value = 0
        .Cells(1, tblCronograma.ListColumns("Estado cronograma").Index).Value = Configuration2.ESTADO_NUNCA_INSPECCIONADO
        .Cells(1, tblCronograma.ListColumns("Puesto activo en personal").Index).Value = "Si"
        .Cells(1, tblCronograma.ListColumns("Personal activo").Index).Value = activo
        .Cells(1, tblCronograma.ListColumns("Fecha ultima actualizacion").Index).Value = Now
        .Cells(1, tblCronograma.ListColumns("Requiere recalculo").Index).Value = "Si"
    End With
    
    ' Actualizar inmediatamente
    Call ActualizarRegistroCronograma(iniciales, idPlantilla)
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("InspectionScheduler.CrearRegistroCronograma", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Funci\u00f3n: GenerarUUID
' Prop\u00f3sito: Genera un identificador \u00fanico en formato UUID con guiones.
' Formato: xxxxxxxx-xxxxxxxx-xxxxxxxx (ejemplo: fxEJV01C-xC6PKG6C-pVOj2dMa)
' ----------------------------------------------------------------------
Private Function GenerarUUID() As String
    Dim parte1 As String
    Dim parte2 As String
    Dim parte3 As String
    
    parte1 = GenerarCadenaAleatoria(8)
    parte2 = GenerarCadenaAleatoria(8)
    parte3 = GenerarCadenaAleatoria(10)
    
    GenerarUUID = parte1 & "-" & parte2 & "-" & parte3
End Function

'' ----------------------------------------------------------------------
' Funci\u00f3n auxiliar: GenerarCadenaAleatoria
' Prop\u00f3sito: Genera una cadena aleatoria de caracteres alfanum\u00e9ricos.
' ----------------------------------------------------------------------
Private Function GenerarCadenaAleatoria(ByVal longitud As Integer) As String
    Dim caracteres As String
    Dim i As Integer
    Dim resultado As String
    
    caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    resultado = ""
    
    Randomize
    For i = 1 To longitud
        resultado = resultado & Mid(caracteres, Int((Len(caracteres) * Rnd) + 1), 1)
    Next i
    
    GenerarCadenaAleatoria = resultado
End Function

'' ----------------------------------------------------------------------
' Función: AuditValor
' Propósito: Formatea un valor para el Audit Trail. Si está vacío o es
'            nulo, retorna "(vacío)". Si es fecha, la formatea.
'            Usada por ActualizarRegistroCronogramaInterno para construir
'            las cadenas before/after del log de auditoría.
' FASE 10: 17/06/2026 — Agregada para cubrir hueco de auditoría programática.
' ----------------------------------------------------------------------
Private Function AuditValor(ByVal val As Variant) As String
    If IsEmpty(val) Then
        AuditValor = "(vacío)"
        Exit Function
    End If
    
    If IsNull(val) Then
        AuditValor = "(vacío)"
        Exit Function
    End If
    
    If Len(CStr(val)) = 0 Then
        AuditValor = "(vacío)"
        Exit Function
    End If
    
    If IsDate(val) Then
        AuditValor = Format(val, "dd/mm/yyyy hh:nn:ss")
    Else
        AuditValor = CStr(val)
    End If
End Function
