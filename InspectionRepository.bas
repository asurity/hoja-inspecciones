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
        ' Campos obligatorios
        .Cells(1, tblInspecciones.ListColumns("ID Inspeccion").Index).Value = idInspeccion
        .Cells(1, tblInspecciones.ListColumns("Iniciales personal").Index).Value = datos("Iniciales")
        .Cells(1, tblInspecciones.ListColumns("ID Plantilla").Index).Value = datos("IDPlantilla")
        .Cells(1, tblInspecciones.ListColumns("Planta ejecutora").Index).Value = datos("Planta")
        .Cells(1, tblInspecciones.ListColumns("Fecha inspeccion").Index).Value = datos("FechaInspeccion")
        .Cells(1, tblInspecciones.ListColumns("Auditor").Index).Value = datos("Evaluador")
        .Cells(1, tblInspecciones.ListColumns("Estado").Index).Value = Configuration2.INSPECCION_EN_PROGRESO
        
        ' Campos nuevos del checklist virtual
        .Cells(1, tblInspecciones.ListColumns("Area").Index).Value = datos("Area")
        .Cells(1, tblInspecciones.ListColumns("Linea auditada").Index).Value = datos("LineaAuditada")
        .Cells(1, tblInspecciones.ListColumns("Hora inicio").Index).Value = datos("HoraInicio")
        .Cells(1, tblInspecciones.ListColumns("Hora termino").Index).Value = datos("HoraTermino")
        .Cells(1, tblInspecciones.ListColumns("Iniciales AY1").Index).Value = datos("AY1")
        .Cells(1, tblInspecciones.ListColumns("Iniciales AY2").Index).Value = datos("AY2")
        .Cells(1, tblInspecciones.ListColumns("Iniciales OP").Index).Value = datos("OP")
        .Cells(1, tblInspecciones.ListColumns("Lugar auditoria").Index).Value = datos("LugarAuditoria")
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
    
    Set wsHistorico = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Set tblRespuestas = wsHistorico.ListObjects(Configuration2.TABLE_RESPUESTAS)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For Each resp In respuestas
        Dim dictResp As Object
        Set dictResp = resp
        
        Set newRow = tblRespuestas.ListRows.Add
        
        With newRow.Range
            .Cells(1, tblRespuestas.ListColumns("ID Respuesta").Index).Value = GenerarUUID()
            .Cells(1, tblRespuestas.ListColumns("ID Inspeccion").Index).Value = idInspeccion
            .Cells(1, tblRespuestas.ListColumns("ID Pregunta").Index).Value = dictResp("IDPregunta")
            .Cells(1, tblRespuestas.ListColumns("ID Opcion").Index).Value = dictResp("IDOpcion")
            .Cells(1, tblRespuestas.ListColumns("Valor numerico").Index).Value = dictResp("ValorNumerico")
            .Cells(1, tblRespuestas.ListColumns("Observacion").Index).Value = dictResp("Observacion")
            .Cells(1, tblRespuestas.ListColumns("Fecha respuesta").Index).Value = Now
        End With
    Next resp
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
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
                .Cells(1, tblInspecciones.ListColumns("TA puntaje").Index).Value = calculos("TA_puntaje")
                .Cells(1, tblInspecciones.ListColumns("TA maximos").Index).Value = calculos("TA_maximos")
                .Cells(1, tblInspecciones.ListColumns("TA noaplica").Index).Value = calculos("TA_noaplica")
                .Cells(1, tblInspecciones.ListColumns("TA porcentaje").Index).Value = calculos("TA_porcentaje")
                
                ' RPN y categoría
                .Cells(1, tblInspecciones.ListColumns("RPN calculado").Index).Value = calculos("RPN")
                .Cells(1, tblInspecciones.ListColumns("Categoria resultado").Index).Value = calculos("Categoria")
                .Cells(1, tblInspecciones.ListColumns("Requiere accion").Index).Value = calculos("RequiereAccion")
                
                ' Programación
                .Cells(1, tblInspecciones.ListColumns("Fecha proxima").Index).Value = calculos("FechaProxima")
                .Cells(1, tblInspecciones.ListColumns("Dias vencimiento").Index).Value = calculos("DiasVencimiento")
                .Cells(1, tblInspecciones.ListColumns("Estado programacion").Index).Value = calculos("EstadoProgramacion")
                
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
