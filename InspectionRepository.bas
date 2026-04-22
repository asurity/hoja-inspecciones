' ----------------------------------------------------------------------
' Módulo: InspectionRepository
' Descripción: CRUD para tblInspecciones y tblRespuestas.
'              Maneja la creación de inspecciones, guardado de respuestas,
'              actualización de cálculos y consultas históricas.
' Fecha creación: 14/04/2026
' Dependencias: Configuration2, ErrorLogger2
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Función: CrearInspeccion
' Propósito: Inserta una nueva fila en tblInspecciones con los datos del formulario.
' Parámetros:
'   datos: Dictionary con claves correspondientes a columnas de tblInspecciones
' Retorna: ID_Inspeccion generado (String) o "" si falla.
' ----------------------------------------------------------------------
Public Function CrearInspeccion(ByVal datos As Object) As String
    On Error GoTo ErrorHandler
    
    Dim wsHistorico As Worksheet
    Dim tblInspecciones As ListObject
    Dim newRow As ListRow
    Dim idInspeccion As String
    
    Set wsHistorico = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Set tblInspecciones = wsHistorico.ListObjects(Configuration2.TABLE_INSPECCIONES)
    
    ' Generar ID único
    idInspeccion = GenerarUUID()
    
    ' Insertar nueva fila
    Set newRow = tblInspecciones.ListRows.Add
    
    With newRow.Range
        .Cells(1, tblInspecciones.ListColumns("ID Inspeccion").Index).Value = idInspeccion
        .Cells(1, tblInspecciones.ListColumns("Iniciales personal").Index).Value = datos("Iniciales")
        .Cells(1, tblInspecciones.ListColumns("ID Plantilla").Index).Value = datos("IDPlantilla")
        .Cells(1, tblInspecciones.ListColumns("Planta").Index).Value = datos("Planta")
        .Cells(1, tblInspecciones.ListColumns("Fecha inspeccion").Index).Value = datos("FechaInspeccion")
        .Cells(1, tblInspecciones.ListColumns("Auditor").Index).Value = datos("Evaluador")
        .Cells(1, tblInspecciones.ListColumns("Estado").Index).Value = Configuration2.INSPECCION_EN_PROGRESO
        
        ' Campos nuevos del checklist virtual
        .Cells(1, tblInspecciones.ListColumns("Area").Index).Value = datos("Area")
        .Cells(1, tblInspecciones.ListColumns("Linea Auditada").Index).Value = datos("LineaAuditada")
        .Cells(1, tblInspecciones.ListColumns("Hora inicio").Index).Value = datos("HoraInicio")
        .Cells(1, tblInspecciones.ListColumns("Hora termino").Index).Value = datos("HoraTermino")
        .Cells(1, tblInspecciones.ListColumns("Iniciales AY1").Index).Value = datos("AY1")
        .Cells(1, tblInspecciones.ListColumns("Iniciales AY2").Index).Value = datos("AY2")
        .Cells(1, tblInspecciones.ListColumns("Iniciales OP").Index).Value = datos("OP")
        .Cells(1, tblInspecciones.ListColumns("Lugar Auditoria").Index).Value = datos("LugarAuditoria")
        .Cells(1, tblInspecciones.ListColumns("Observaciones generales").Index).Value = datos("ObservacionGeneral")
    End With
    
    CrearInspeccion = idInspeccion
    Exit Function
    
ErrorHandler:
    CrearInspeccion = ""
    Call ErrorLogger2.Log("InspectionRepository.CrearInspeccion", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Subrutina: GuardarRespuestas
' Propósito: Inserta múltiples filas en tblRespuestas para una inspección.
' Parámetros:
'   idInspeccion: ID de la inspección padre
'   respuestas: Collection de Dictionary con claves:
'     "IDPregunta", "IDOpcion", "ValorNumerico", "Observacion"
' ----------------------------------------------------------------------
Public Sub GuardarRespuestas(ByVal idInspeccion As String, ByVal respuestas As Collection)
    On Error GoTo ErrorHandler
    
    Dim wsHistorico As Worksheet
    Dim tblRespuestas As ListObject
    Dim newRow As ListRow
    Dim resp As Variant
    
    Debug.Print "[GuardarRespuestas] Inicio. ID Inspeccion: " & idInspeccion
    
    Set wsHistorico = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Debug.Print "[GuardarRespuestas] Hoja Histórico obtenida: " & wsHistorico.Name
    
    Set tblRespuestas = wsHistorico.ListObjects(Configuration2.TABLE_RESPUESTAS)
    Debug.Print "[GuardarRespuestas] Tabla tblRespuestas obtenida con " & tblRespuestas.ListColumns.Count & " columnas"
    
    ' Leer nombres de columnas para verificar estructura
    Dim colIdx As Long
    For colIdx = 1 To tblRespuestas.ListColumns.Count
        Debug.Print "  Columna " & colIdx & ": " & tblRespuestas.ListColumns(colIdx).Name
    Next colIdx
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim contador As Long
    contador = 0
    
    ' Verificar si existe columna ID Criticidad (columna 8)
    Dim tieneColumnaCriticidad As Boolean
    tieneColumnaCriticidad = False
    
    Debug.Print "[GuardarRespuestas] Total columnas en tblRespuestas: " & tblRespuestas.ListColumns.Count
    
    If tblRespuestas.ListColumns.Count >= 8 Then
        Dim nombreCol8 As String
        nombreCol8 = tblRespuestas.ListColumns(8).Name
        Debug.Print "[GuardarRespuestas] Nombre columna 8: '" & nombreCol8 & "'"
        
        If InStr(1, nombreCol8, "Criticidad", vbTextCompare) > 0 Then
            tieneColumnaCriticidad = True
            Debug.Print "[GuardarRespuestas] ✓ Columna ID Criticidad encontrada: " & nombreCol8
        Else
            Debug.Print "[GuardarRespuestas] ✗ ADVERTENCIA: Columna 8 existe pero no es ID Criticidad: " & nombreCol8
            Debug.Print "[GuardarRespuestas] ✗ Ver: docs/INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md"
        End If
    Else
        Debug.Print "[GuardarRespuestas] ✗ ADVERTENCIA: No hay columna 8 (ID Criticidad)."
        Debug.Print "[GuardarRespuestas] ✗ ACCIÓN REQUERIDA: Agregar columna 'ID Criticidad' como columna 8"
        Debug.Print "[GuardarRespuestas] ✗ Ver: docs/INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md"
    End If
    
    For Each resp In respuestas
        contador = contador + 1
        Debug.Print "[GuardarRespuestas] Procesando respuesta " & contador
        
        Dim dictResp As Object
        Set dictResp = resp
        
        Debug.Print "  IDPregunta: " & dictResp("IDPregunta")
        Debug.Print "  IDOpcion: " & dictResp("IDOpcion")
        Debug.Print "  ValorNumerico: " & dictResp("ValorNumerico")
        If dictResp.Exists("IDCriticidad") Then
            Debug.Print "  IDCriticidad: " & dictResp("IDCriticidad")
        Else
            Debug.Print "  IDCriticidad: (no presente en Dictionary)"
        End If
        
        Set newRow = tblRespuestas.ListRows.Add
        Debug.Print "  Nueva fila agregada"
        
        ' Usar índices de columna en lugar de nombres para evitar problemas de acentos
        ' Estructura de tblRespuestas:
        ' [1] ID Respuesta, [2] ID Inspeccion, [3] ID Pregunta, 
        ' [4] ID Opcion, [5] Valor numerico, [6] Observacion, [7] Fecha respuesta
        ' [8] ID Criticidad (NUEVA - ver docs/INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md)
        
        With newRow.Range
            Debug.Print "  Estableciendo Columna 1 (ID Respuesta)..."
            .Cells(1, 1).Value = GenerarUUID()
            
            Debug.Print "  Estableciendo Columna 2 (ID Inspeccion)..."
            .Cells(1, 2).Value = idInspeccion
            
            Debug.Print "  Estableciendo Columna 3 (ID Pregunta)..."
            .Cells(1, 3).Value = dictResp("IDPregunta")
            
            Debug.Print "  Estableciendo Columna 4 (ID Opcion)..."
            .Cells(1, 4).Value = dictResp("IDOpcion")
            
            Debug.Print "  Estableciendo Columna 5 (Valor numerico)..."
            .Cells(1, 5).Value = dictResp("ValorNumerico")
            
            Debug.Print "  Estableciendo Columna 6 (Observacion)..."
            .Cells(1, 6).Value = dictResp("Observacion")
            
            Debug.Print "  Estableciendo Columna 7 (Fecha respuesta)..."
            .Cells(1, 7).Value = Now
            
            ' Guardar ID Criticidad si la columna existe y el dato está presente
            If tieneColumnaCriticidad Then
                If dictResp.Exists("IDCriticidad") Then
                    Debug.Print "  Estableciendo Columna 8 (ID Criticidad)..."
                    .Cells(1, 8).Value = dictResp("IDCriticidad")
                Else
                    Debug.Print "  Columna 8 dejada vacía (IDCriticidad no presente en respuesta)"
                    .Cells(1, 8).Value = ""
                End If
            End If
        End With
        
        Debug.Print "  Respuesta " & contador & " completada"
    Next resp
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Debug.Print "[GuardarRespuestas] ERROR en respuesta " & contador & ": " & Err.Description
    Call ErrorLogger2.Log("InspectionRepository.GuardarRespuestas", Err.Description, Err.Number)
    Err.Raise Err.Number, "InspectionRepository.GuardarRespuestas", Err.Description
End Sub

'' ----------------------------------------------------------------------
' Subrutina: ActualizarCalculosInspeccion
' Propósito: Actualiza los campos de cálculos en tblInspecciones después
'            de completar el scoring y RPN.
' Parámetros:
'   idInspeccion: ID de la inspección a actualizar
'   calculos: Dictionary con claves:
'     "TA_puntaje", "TA_maximos", "TA_noaplica", "TA_porcentaje",
'     "RPN", "Categoria", "RequiereAccion",
'     "FechaProxima", "DiasVencimiento", "EstadoProgramacion"
' ----------------------------------------------------------------------
Public Sub ActualizarCalculosInspeccion(ByVal idInspeccion As String, ByVal calculos As Object)
    On Error GoTo ErrorHandler
    
    Dim wsHistorico As Worksheet
    Dim tblInspecciones As ListObject
    Dim inspeccionRow As ListRow
    
    Set wsHistorico = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Set tblInspecciones = wsHistorico.ListObjects(Configuration2.TABLE_INSPECCIONES)
    
    If tblInspecciones.DataBodyRange Is Nothing Then Exit Sub
    
    For Each inspeccionRow In tblInspecciones.ListRows
        Dim currentID As String
        currentID = Trim(inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("ID Inspeccion").Index).Value)
        
        If currentID = Trim(idInspeccion) Then
            With inspeccionRow.Range
                ' Scoring TA
                .Cells(1, tblInspecciones.ListColumns("TA puntaje obtenido").Index).Value = calculos("TA_puntaje")
                .Cells(1, tblInspecciones.ListColumns("TA puntos maximos").Index).Value = calculos("TA_maximos")
                .Cells(1, tblInspecciones.ListColumns("TA puntos no aplica").Index).Value = calculos("TA_noaplica")
                .Cells(1, tblInspecciones.ListColumns("TA porcentaje").Index).Value = calculos("TA_porcentaje")
                
                ' Auditoría de Procesos (intentar guardar si la columna existe)
                If calculos.Exists("Auditoria_Procesos_Resultado") Then
                    On Error Resume Next
                    Dim colIndexProcesos As Variant
                    colIndexProcesos = Empty
                    
                    ' Intentar encontrar la columna por varios nombres posibles
                    Dim nombresPosibles As Variant
                    nombresPosibles = Array("Auditoria Procesos Resultado", "Auditoría Procesos Resultado", _
                                           "Resultado Auditoria Procesos", "AP Resultado")
                    
                    Dim nombreCol As Variant
                    For Each nombreCol In nombresPosibles
                        colIndexProcesos = Application.Match(CStr(nombreCol), tblInspecciones.HeaderRowRange, 0)
                        If Not IsError(colIndexProcesos) Then
                            .Cells(1, CLng(colIndexProcesos)).Value = calculos("Auditoria_Procesos_Resultado")
                            Debug.Print "[ActualizarCalculos] Auditoría Procesos guardada en columna '" & nombreCol & "' (" & colIndexProcesos & "): " & calculos("Auditoria_Procesos_Resultado")
                            Exit For
                        End If
                    Next nombreCol
                    
                    If IsError(colIndexProcesos) Or IsEmpty(colIndexProcesos) Then
                        Debug.Print "[ActualizarCalculos] ADVERTENCIA: No se encontró columna para Auditoría de Procesos. Dato: " & calculos("Auditoria_Procesos_Resultado")
                        Debug.Print "[ActualizarCalculos] Columnas existentes: " & tblInspecciones.ListColumns.Count
                    End If
                    On Error GoTo ErrorHandler
                End If
                
                ' RPN y categoría
                .Cells(1, tblInspecciones.ListColumns("RPN calculado").Index).Value = calculos("RPN")
                .Cells(1, tblInspecciones.ListColumns("Categoria resultado").Index).Value = calculos("Categoria")
                .Cells(1, tblInspecciones.ListColumns("Requiere accion").Index).Value = calculos("RequiereAccion")
                
                ' Programación
                .Cells(1, tblInspecciones.ListColumns("Fecha proxima inspeccion").Index).Value = calculos("FechaProxima")
                .Cells(1, tblInspecciones.ListColumns("Dias para vencimiento").Index).Value = calculos("DiasVencimiento")
                .Cells(1, tblInspecciones.ListColumns("Estado programacion").Index).Value = calculos("EstadoProgramacion")
                
                ' ----------------------------------------------------------------------
                ' DATOS RECURRENTES (Nuevos - Fase 5)
                ' ----------------------------------------------------------------------
                ' Guardar campos recurrentes si existen en calculos
                If calculos.Exists("NumeroInspeccion") Then
                    On Error Resume Next
                    Dim colIdx As Long
                    
                    ' Numero Inspeccion (columna 32)
                    colIdx = 0
                    colIdx = Application.Match("Numero Inspeccion", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then
                        .Cells(1, colIdx).Value = calculos("NumeroInspeccion")
                        Debug.Print "[ActualizarCalculos] Numero Inspeccion: " & calculos("NumeroInspeccion")
                    End If
                    
                    ' Es Inspeccion Recurrente (columna 33)
                    colIdx = 0
                    colIdx = Application.Match("Es Inspeccion Recurrente", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then
                        Dim valorRecurrente As String
                        valorRecurrente = IIf(calculos("EsInspeccionRecurrente"), "Si", "No")
                        .Cells(1, colIdx).Value = valorRecurrente
                        Debug.Print "[ActualizarCalculos] Es Recurrente: " & valorRecurrente
                    End If
                    
                    ' Puesto Evaluado (columna 34)
                    colIdx = 0
                    colIdx = Application.Match("Puesto Evaluado", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then
                        .Cells(1, colIdx).Value = calculos("PuestoEvaluado")
                        Debug.Print "[ActualizarCalculos] Puesto Evaluado: " & calculos("PuestoEvaluado")
                    End If
                    
                    ' Campos especificos de inspecciones recurrentes (2da+)
                    If calculos("EsInspeccionRecurrente") Then
                        Debug.Print "[ActualizarCalculos] Guardando datos recurrentes..."
                        
                        ' RPN Anterior (columna 35 si manual, vacia si auto)
                        If calculos.Exists("RPNAnterior") Then
                            colIdx = 0
                            colIdx = Application.Match("RPN Anterior Manual", tblInspecciones.HeaderRowRange, 0)
                            If colIdx > 0 And Len(calculos("IDInspeccionAnterior")) = 0 Then
                                ' Solo guardar si fue manual (sin ID anterior)
                                .Cells(1, colIdx).Value = calculos("RPNAnterior")
                                Debug.Print "[ActualizarCalculos] RPN Anterior Manual: " & Format(calculos("RPNAnterior"), "0.00")
                            End If
                        End If
                        
                        ' ID Inspeccion Anterior (columna 36)
                        If calculos.Exists("IDInspeccionAnterior") And Len(calculos("IDInspeccionAnterior")) > 0 Then
                            colIdx = 0
                            colIdx = Application.Match("ID Inspeccion Anterior", tblInspecciones.HeaderRowRange, 0)
                            If colIdx > 0 Then
                                .Cells(1, colIdx).Value = calculos("IDInspeccionAnterior")
                                Debug.Print "[ActualizarCalculos] ID Inspeccion Anterior: " & calculos("IDInspeccionAnterior")
                            End If
                        End If
                        
                        ' RPN Promedio (columna 37)
                        If calculos.Exists("RPNPromedio") Then
                            colIdx = 0
                            colIdx = Application.Match("RPN Promedio", tblInspecciones.HeaderRowRange, 0)
                            If colIdx > 0 Then
                                .Cells(1, colIdx).Value = calculos("RPNPromedio")
                                Debug.Print "[ActualizarCalculos] RPN Promedio: " & Format(calculos("RPNPromedio"), "0.00")
                            End If
                        End If
                        
                        ' RPN Total (columna 40)
                        If calculos.Exists("RPNTotal") Then
                            colIdx = 0
                            colIdx = Application.Match("RPN Total", tblInspecciones.HeaderRowRange, 0)
                            If colIdx > 0 Then
                                .Cells(1, colIdx).Value = calculos("RPNTotal")
                                Debug.Print "[ActualizarCalculos] RPN Total: " & Format(calculos("RPNTotal"), "0.00")
                            End If
                        End If
                        
                        Debug.Print "[ActualizarCalculos] Datos recurrentes guardados OK"
                    End If
                    
                    On Error GoTo ErrorHandler
                Else
                    ' No hay datos recurrentes, asumir 1ra inspeccion
                    On Error Resume Next
                    colIdx = Application.Match("Numero Inspeccion", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then .Cells(1, colIdx).Value = 1
                    
                    colIdx = Application.Match("Es Inspeccion Recurrente", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then .Cells(1, colIdx).Value = "No"
                    On Error GoTo ErrorHandler
                End If
                
                ' Estado final
                .Cells(1, tblInspecciones.ListColumns("Estado").Index).Value = Configuration2.INSPECCION_COMPLETADO
                
                ' Auditoría
                .Cells(1, tblInspecciones.ListColumns("Fecha calculo").Index).Value = Now
                .Cells(1, tblInspecciones.ListColumns("Usuario calculo").Index).Value = Environ("Username")
            End With
            
            Exit For
        End If
    Next inspeccionRow
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("InspectionRepository.ActualizarCalculosInspeccion", Err.Description, Err.Number)
    Err.Raise Err.Number, "InspectionRepository.ActualizarCalculosInspeccion", Err.Description
End Sub

'' ----------------------------------------------------------------------
' Función: ObtenerUltimasNInspecciones
' Propósito: Obtiene las últimas N inspecciones completadas para una
'            combinación persona+plantilla. Para validar Categoría 5.
' Retorna: Collection de RPNs ordenados por fecha DESC.
' ----------------------------------------------------------------------
Public Function ObtenerUltimasNInspecciones(ByVal iniciales As String, ByVal idPlantilla As String, ByVal n As Long) As Collection
    On Error GoTo ErrorHandler
    
    Dim wsHistorico As Worksheet
    Dim tblInspecciones As ListObject
    Dim inspeccionRow As ListRow
    Dim resultado As New Collection
    
    Set wsHistorico = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Set tblInspecciones = wsHistorico.ListObjects(Configuration2.TABLE_INSPECCIONES)
    
    If tblInspecciones.DataBodyRange Is Nothing Then
        Set ObtenerUltimasNInspecciones = resultado
        Exit Function
    End If
    
    ' Recopilar todas las inspecciones completadas para esta persona+plantilla
    Dim tempFechas As New Collection
    Dim tempRPNs As New Collection
    
    For Each inspeccionRow In tblInspecciones.ListRows
        Dim inspIniciales As String
        Dim inspPlantilla As String
        Dim inspEstado As String
        
        inspIniciales = Trim(inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("Iniciales personal").Index).Value)
        inspPlantilla = Trim(inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("ID Plantilla").Index).Value)
        inspEstado = Trim(inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("Estado").Index).Value)
        
        If inspIniciales = Trim(iniciales) And inspPlantilla = Trim(idPlantilla) And _
           inspEstado = Configuration2.INSPECCION_COMPLETADO Then
            
            Dim inspFecha As Date
            Dim inspRPN As Double
            
            On Error Resume Next
            inspFecha = inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("Fecha inspeccion").Index).Value
            inspRPN = CDbl(inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("RPN calculado").Index).Value)
            On Error GoTo ErrorHandler
            
            Dim par(0 To 1) As Variant
            par(0) = inspFecha
            par(1) = inspRPN
            tempFechas.Add par
        End If
    Next inspeccionRow
    
    ' Ordenar por fecha DESC y tomar las últimas N
    ' Usar bubble sort simple sobre la collection
    Dim arrTemp() As Variant
    Dim cnt As Long
    cnt = tempFechas.Count
    
    If cnt = 0 Then
        Set ObtenerUltimasNInspecciones = resultado
        Exit Function
    End If
    
    ReDim arrTemp(1 To cnt, 1 To 2)
    Dim idx As Long
    For idx = 1 To cnt
        Dim p As Variant
        p = tempFechas(idx)
        arrTemp(idx, 1) = p(0) ' fecha
        arrTemp(idx, 2) = p(1) ' RPN
    Next idx
    
    ' Ordenar por fecha DESC
    Dim ii As Long, jj As Long
    Dim tmpVal As Variant
    For ii = 1 To cnt - 1
        For jj = 1 To cnt - ii
            If arrTemp(jj, 1) < arrTemp(jj + 1, 1) Then
                ' Swap fecha
                tmpVal = arrTemp(jj, 1)
                arrTemp(jj, 1) = arrTemp(jj + 1, 1)
                arrTemp(jj + 1, 1) = tmpVal
                ' Swap RPN
                tmpVal = arrTemp(jj, 2)
                arrTemp(jj, 2) = arrTemp(jj + 1, 2)
                arrTemp(jj + 1, 2) = tmpVal
            End If
        Next jj
    Next ii
    
    ' Tomar las primeras N
    Dim limite As Long
    limite = IIf(n < cnt, n, cnt)
    For idx = 1 To limite
        resultado.Add CDbl(arrTemp(idx, 2))
    Next idx
    
    Set ObtenerUltimasNInspecciones = resultado
    Exit Function
    
ErrorHandler:
    Set ObtenerUltimasNInspecciones = New Collection
    Call ErrorLogger2.Log("InspectionRepository.ObtenerUltimasNInspecciones", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Subrutina: EliminarInspeccion
' Propósito: Elimina una inspección y sus respuestas (para rollback).
' ----------------------------------------------------------------------
Public Sub EliminarInspeccion(ByVal idInspeccion As String)
    On Error GoTo ErrorHandler
    
    Dim wsHistorico As Worksheet
    Dim tblInspecciones As ListObject
    Dim tblRespuestas As ListObject
    Dim i As Long
    
    Set wsHistorico = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Set tblInspecciones = wsHistorico.ListObjects(Configuration2.TABLE_INSPECCIONES)
    Set tblRespuestas = wsHistorico.ListObjects(Configuration2.TABLE_RESPUESTAS)
    
    ' Eliminar respuestas asociadas (recorrer de atrás hacia adelante)
    If Not tblRespuestas.DataBodyRange Is Nothing Then
        For i = tblRespuestas.ListRows.Count To 1 Step -1
            Dim respID As String
            respID = Trim(tblRespuestas.ListRows(i).Range.Cells(1, tblRespuestas.ListColumns("ID Inspeccion").Index).Value)
            If respID = Trim(idInspeccion) Then
                tblRespuestas.ListRows(i).Delete
            End If
        Next i
    End If
    
    ' Eliminar inspección
    If Not tblInspecciones.DataBodyRange Is Nothing Then
        For i = tblInspecciones.ListRows.Count To 1 Step -1
            Dim inspID As String
            inspID = Trim(tblInspecciones.ListRows(i).Range.Cells(1, tblInspecciones.ListColumns("ID Inspeccion").Index).Value)
            If inspID = Trim(idInspeccion) Then
                tblInspecciones.ListRows(i).Delete
                Exit For
            End If
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("InspectionRepository.EliminarInspeccion", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Función: GenerarUUID
' Propósito: Genera un identificador único en formato UUID con guiones.
' Nota: Misma implementación que InspectionScheduler para consistencia.
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
