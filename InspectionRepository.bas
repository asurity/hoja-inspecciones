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
    
    ' FASE 9 (09/06/2026): Desproteger hoja Histórico para permitir escritura VBA
    Call SheetProtector2.UnprotectSheet(wsHistorico, Configuration2.APP_PASSWORD)
    
    ' Generar ID único
    idInspeccion = GenerarUUID()
    
    ' Insertar nueva fila
    Set newRow = tblInspecciones.ListRows.Add
    
    With newRow.Range
        .Cells(1, tblInspecciones.ListColumns("ID Inspeccion").Index).Value = idInspeccion
        .Cells(1, tblInspecciones.ListColumns("Iniciales personal").Index).Value = datos("Iniciales")
        .Cells(1, tblInspecciones.ListColumns("ID Plantilla").Index).Value = datos("IDPlantilla")
        .Cells(1, tblInspecciones.ListColumns("Planta").Index).Value = datos("Planta")
        .Cells(1, tblInspecciones.ListColumns("Auditor").Index).Value = datos("Evaluador")
        .Cells(1, tblInspecciones.ListColumns("Estado").Index).Value = Configuration2.INSPECCION_EN_PROGRESO

        ' --- FECHA INSPECCION (CORRECCIÓN FORMATO EXCEL) ---
        Dim parsedFechaInsp As Variant
        parsedFechaInsp = ChecklistValidator.ParseFechaDMY(CStr(datos("FechaInspeccion")))
        If Not IsEmpty(parsedFechaInsp) Then
            .Cells(1, tblInspecciones.ListColumns("Fecha inspeccion").Index).Value = parsedFechaInsp ' Envía Date puro
        Else
            .Cells(1, tblInspecciones.ListColumns("Fecha inspeccion").Index).Value = datos("FechaInspeccion")
        End If
        .Cells(1, tblInspecciones.ListColumns("Fecha inspeccion").Index).NumberFormat = "dd/mm/yyyy"
        
        ' --- FECHA AUDITADA (CORRECCIÓN FORMATO EXCEL) ---
        If datos.Exists("FechaAuditada") Then
            On Error Resume Next
            Dim colFechaAuditada As Variant
            colFechaAuditada = Application.Match("Fecha Auditada", tblInspecciones.HeaderRowRange, 0)
            If Not IsError(colFechaAuditada) Then
                Dim parsedFechaAud As Variant
                parsedFechaAud = ChecklistValidator.ParseFechaDMY(CStr(datos("FechaAuditada")))
                If Not IsEmpty(parsedFechaAud) Then
                    .Cells(1, CLng(colFechaAuditada)).Value = parsedFechaAud
                Else
                    .Cells(1, CLng(colFechaAuditada)).Value = datos("FechaAuditada")
                End If
                .Cells(1, CLng(colFechaAuditada)).NumberFormat = "dd/mm/yyyy"
            End If
            On Error GoTo ErrorHandler
        End If
        
        ' Campos nuevos del checklist virtual
        .Cells(1, tblInspecciones.ListColumns("Area").Index).Value = datos("Area")
        .Cells(1, tblInspecciones.ListColumns("Linea Auditada").Index).Value = datos("LineaAuditada")
        .Cells(1, tblInspecciones.ListColumns("Hora inicio").Index).Value = datos("HoraInicio")
        .Cells(1, tblInspecciones.ListColumns("Hora termino").Index).Value = datos("HoraTermino")
        
        ' Campos opcionales: AY1, AY2, OP - usar Configuration2.VALOR_NO_APLICA si están vacíos
        Dim ay1Value As String, ay2Value As String, opValue As String
        ay1Value = Trim(CStr(datos("AY1")))
        ay2Value = Trim(CStr(datos("AY2")))
        opValue = Trim(CStr(datos("OP")))
        
        .Cells(1, tblInspecciones.ListColumns("Iniciales AY1").Index).Value = IIf(Len(ay1Value) > 0, ay1Value, Configuration2.VALOR_NO_APLICA)
        .Cells(1, tblInspecciones.ListColumns("Iniciales AY2").Index).Value = IIf(Len(ay2Value) > 0, ay2Value, Configuration2.VALOR_NO_APLICA)
        .Cells(1, tblInspecciones.ListColumns("Iniciales OP").Index).Value = IIf(Len(opValue) > 0, opValue, Configuration2.VALOR_NO_APLICA)
        
        .Cells(1, tblInspecciones.ListColumns("Lugar Auditoria").Index).Value = datos("LugarAuditoria")
        
        ' Observaciones generales: guardar Configuration2.VALOR_NO_APLICA si está vacía
        Dim obsGeneral As String
        obsGeneral = Trim(CStr(datos("ObservacionGeneral")))
        If Len(obsGeneral) = 0 Then
            .Cells(1, tblInspecciones.ListColumns("Observaciones generales").Index).Value = Configuration2.VALOR_NO_APLICA
        Else
            .Cells(1, tblInspecciones.ListColumns("Observaciones generales").Index).Value = obsGeneral
        End If
        
        ' === NUEVOS CAMPOS - FASE 7 (23/04/2026): Calificaciones y Vencimientos ===
        On Error Resume Next
        Dim colIdx As Variant 
        
        ' Calificación Vestuario
        colIdx = Application.Match("Calificacion Vestuario", tblInspecciones.HeaderRowRange, 0)
        If Not IsError(colIdx) Then
            Dim califVest As String
            califVest = Trim(CStr(datos("CalificacionVestuario")))
            .Cells(1, CLng(colIdx)).Value = IIf(Len(califVest) > 0, califVest, "Si")
        End If
        
        ' --- FECHA VENCIMIENTO VESTUARIO (CORRECCIÓN FORMATO EXCEL) ---
        colIdx = Application.Match("Fecha Venc Vestuario", tblInspecciones.HeaderRowRange, 0)
        If Not IsError(colIdx) Then
            Dim parsedVencVest As Variant
            parsedVencVest = ChecklistValidator.ParseFechaDMY(CStr(datos("FechaVencVestuario")))
            
            If Not IsEmpty(parsedVencVest) Then
                .Cells(1, CLng(colIdx)).Value = parsedVencVest
                .Cells(1, CLng(colIdx)).NumberFormat = "dd/mm/yyyy"
            Else
                Dim fechaVencVest As String
                fechaVencVest = Trim(CStr(datos("FechaVencVestuario")))
                .Cells(1, CLng(colIdx)).Value = IIf(Len(fechaVencVest) > 0, fechaVencVest, Configuration2.VALOR_NO_APLICA)
            End If
        End If
        
        ' Calificación Operador
        colIdx = Application.Match("Calificacion Operador", tblInspecciones.HeaderRowRange, 0)
        If Not IsError(colIdx) Then
            Dim califOper As String
            califOper = Trim(CStr(datos("CalificacionOperador")))
            .Cells(1, CLng(colIdx)).Value = IIf(Len(califOper) > 0, califOper, "Si")
        End If
        
        ' --- FECHA VENCIMIENTO OPERADOR (CORRECCIÓN FORMATO EXCEL) ---
        colIdx = Application.Match("Fecha Venc Operador", tblInspecciones.HeaderRowRange, 0)
        If Not IsError(colIdx) Then
            Dim parsedVencOper As Variant
            parsedVencOper = ChecklistValidator.ParseFechaDMY(CStr(datos("FechaVencOperador")))
            
            If Not IsEmpty(parsedVencOper) Then
                .Cells(1, CLng(colIdx)).Value = parsedVencOper
                .Cells(1, CLng(colIdx)).NumberFormat = "dd/mm/yyyy"
            Else
                Dim fechaVencOper As String
                fechaVencOper = Trim(CStr(datos("FechaVencOperador")))
                .Cells(1, CLng(colIdx)).Value = IIf(Len(fechaVencOper) > 0, fechaVencOper, Configuration2.VALOR_NO_APLICA)
            End If
        End If
        
        ' Columnas no usadas actualmente
        colIdx = Application.Match("Fecha completado", tblInspecciones.HeaderRowRange, 0)
        If Not IsError(colIdx) Then
            .Cells(1, CLng(colIdx)).Value = Configuration2.VALOR_NO_APLICA
        End If
        
        colIdx = Application.Match("Usuario completado", tblInspecciones.HeaderRowRange, 0)
        If Not IsError(colIdx) Then
            .Cells(1, CLng(colIdx)).Value = Configuration2.VALOR_NO_APLICA
        End If

        On Error GoTo ErrorHandler
        ' === FIN NUEVOS CAMPOS FASE 7 ===
        
    End With
    
    ' Reproteger hoja según el rol del usuario
    Call SheetProtector2.ApplyRoleBasedProtection(wsHistorico, Configuration2.APP_PASSWORD)
    
    CrearInspeccion = idInspeccion
    Exit Function
    
ErrorHandler:
    ' Reproteger hoja incluso si hay error (fail-safe)
    On Error Resume Next
    Call SheetProtector2.ApplyRoleBasedProtection(wsHistorico, Configuration2.APP_PASSWORD)
    On Error GoTo 0
    
    CrearInspeccion = ""
    Call ErrorLogger2.Log("InspectionRepository.CrearInspeccion", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ExisteEnCronograma (PUBLIC)
' Propósito: Verifica si existe una entrada activa en tblCronogramaInspecciones
'            que coincida con Iniciales + IDPlantilla + Puesto + Planta.
'            Usada por btnInhabilitar_Click para validar antes de inhabilitar.
'
' Parámetros:
'   iniciales:   Iniciales del personal
'   idPlantilla: ID de la plantilla
'   puesto:      Nombre del puesto
'   planta:      Nombre de la planta
'
' Retorna: True si existe al menos una entrada coincidente en el cronograma.
' Fecha creación: 17/06/2026 - Fix: validar existencia en cronograma antes de inhabilitar
' ----------------------------------------------------------------------
Public Function ExisteEnCronograma( _
    ByVal iniciales As String, _
    ByVal idPlantilla As String, _
    ByVal puesto As String, _
    ByVal planta As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    
    If tblCronograma.DataBodyRange Is Nothing Then
        ExisteEnCronograma = False
        Exit Function
    End If
    
    Dim cronoRow As ListRow
    Dim colIniciales As Long, colIDPlantilla As Long
    Dim colPuesto As Long, colPlanta As Long
    
    colIniciales = tblCronograma.ListColumns("Iniciales personal").Index
    colIDPlantilla = tblCronograma.ListColumns("ID Plantilla").Index
    colPuesto = tblCronograma.ListColumns("Puesto").Index
    colPlanta = tblCronograma.ListColumns("Planta personal").Index
    
    For Each cronoRow In tblCronograma.ListRows
        Dim iniCrono As String, idPltCrono As String
        Dim puestoCrono As String, plantaCrono As String
        
        iniCrono = Trim(CStr(cronoRow.Range.Cells(1, colIniciales).Value))
        idPltCrono = Trim(CStr(cronoRow.Range.Cells(1, colIDPlantilla).Value))
        puestoCrono = Trim(CStr(cronoRow.Range.Cells(1, colPuesto).Value))
        plantaCrono = Trim(CStr(cronoRow.Range.Cells(1, colPlanta).Value))
        
        If UCase(iniCrono) = UCase(iniciales) And _
           UCase(idPltCrono) = UCase(idPlantilla) And _
           UCase(puestoCrono) = UCase(puesto) And _
           UCase(plantaCrono) = UCase(planta) Then
            
            ExisteEnCronograma = True
            Exit Function
        End If
    Next cronoRow
    
    ExisteEnCronograma = False
    Exit Function
    
ErrorHandler:
    ExisteEnCronograma = False
    Call ErrorLogger2.Log("InspectionRepository.ExisteEnCronograma", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Subrutina: InhabilitarInspeccionEnCronograma (PUBLIC)
' Propósito: Marca una entrada en tblCronogramaInspecciones como inhabilitada
'            actualizando los campos "Activo en cronograma" y "Estado cronograma"
'            a "Inhabilitada". Esto hace que desaparezca del resumen (tblResumenCronograma).
'            NO escribe en tblInspecciones ni en tblRespuestas.
'            El audit trail se registra vía Worksheet_Change de Hoja8 (Cronograma).
'
' Parámetros:
'   iniciales:   Iniciales del personal (ej. "JGP")
'   idPlantilla: ID de la plantilla (ej. "TPL003")
'   puesto:      Nombre del puesto evaluado (ej. "Químico")
'   planta:      Nombre de la planta (ej. "PTA")
'
' Fecha creación: 05/08/2026 - Feature "Inhabilitar Inspección"
' Última mod: 17/06/2026 - Refactor: solo UPDATE en cronograma, sin INSERT en tblInspecciones
' ----------------------------------------------------------------------
Public Sub InhabilitarInspeccionEnCronograma( _
    ByVal iniciales As String, _
    ByVal idPlantilla As String, _
    ByVal puesto As String, _
    ByVal planta As String)
    
    On Error GoTo ErrorHandler
    
    ' ──────────────────────────────────────────────────────────────────
    ' FIX (17/06/2026): Deshabilitar refresco de pantalla durante
    ' la operación para evitar saltos causados por Protect/Unprotect.
    ' EnableEvents se deja activo para que Worksheet_Change registre
    ' los cambios en el audit trail (tblCronogramaInspecciones).
    ' ──────────────────────────────────────────────────────────────────
    Application.ScreenUpdating = False
    
    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    
    If tblCronograma.DataBodyRange Is Nothing Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Desproteger para escritura VBA
    Call SheetProtector2.UnprotectSheet(wsCronograma, Configuration2.APP_PASSWORD)
    
    Dim cronoRow As ListRow
    Dim colIniciales As Long, colIDPlantilla As Long
    Dim colPuesto As Long, colPlanta As Long
    Dim colActivoCrono As Long, colEstadoCrono As Long
    
    colIniciales = tblCronograma.ListColumns("Iniciales personal").Index
    colIDPlantilla = tblCronograma.ListColumns("ID Plantilla").Index
    colPuesto = tblCronograma.ListColumns("Puesto").Index
    colPlanta = tblCronograma.ListColumns("Planta personal").Index
    
    ' Columnas a actualizar (manejar si no existen aún)
    On Error Resume Next
    colActivoCrono = tblCronograma.ListColumns("Activo en cronograma").Index
    If Err.Number <> 0 Then colActivoCrono = 0: Err.Clear
    
    colEstadoCrono = tblCronograma.ListColumns("Estado cronograma").Index
    If Err.Number <> 0 Then colEstadoCrono = 0: Err.Clear
    On Error GoTo ErrorHandler
    
    For Each cronoRow In tblCronograma.ListRows
        Dim iniCrono As String, idPltCrono As String
        Dim puestoCrono As String, plantaCrono As String
        
        iniCrono = Trim(CStr(cronoRow.Range.Cells(1, colIniciales).Value))
        idPltCrono = Trim(CStr(cronoRow.Range.Cells(1, colIDPlantilla).Value))
        puestoCrono = Trim(CStr(cronoRow.Range.Cells(1, colPuesto).Value))
        plantaCrono = Trim(CStr(cronoRow.Range.Cells(1, colPlanta).Value))
        
        If UCase(iniCrono) = UCase(iniciales) And _
           UCase(idPltCrono) = UCase(idPlantilla) And _
           UCase(puestoCrono) = UCase(puesto) And _
           UCase(plantaCrono) = UCase(planta) Then
            
            ' Marcar como inhabilitada
            If colActivoCrono > 0 Then
                cronoRow.Range.Cells(1, colActivoCrono).Value = "Inhabilitada"
            End If
            If colEstadoCrono > 0 Then
                cronoRow.Range.Cells(1, colEstadoCrono).Value = "Inhabilitada"
            End If
            
            Exit For
        End If
    Next cronoRow
    
    ' Reproteger
    Call SheetProtector2.ApplyRoleBasedProtection(wsCronograma, Configuration2.APP_PASSWORD)
    
    ' Restaurar refresco de pantalla
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    Call SheetProtector2.ApplyRoleBasedProtection(wsCronograma, Configuration2.APP_PASSWORD)
    On Error GoTo 0
    
    ' Restaurar refresco de pantalla incluso en error
    Application.ScreenUpdating = True
    
    Call ErrorLogger2.Log("InspectionRepository.InhabilitarInspeccionEnCronograma", Err.Description, Err.Number)
End Sub

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
    
    ' Debug.Print "[GuardarRespuestas] Inicio. ID Inspeccion: " & idInspeccion
    
    Set wsHistorico = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    ' Debug.Print "[GuardarRespuestas] Hoja Histórico obtenida: " & wsHistorico.Name
    
    ' FASE 9 (09/06/2026): Desproteger hoja Histórico para permitir escritura VBA
    Call SheetProtector2.UnprotectSheet(wsHistorico, Configuration2.APP_PASSWORD)
    
    Set tblRespuestas = wsHistorico.ListObjects(Configuration2.TABLE_RESPUESTAS)
    ' Debug.Print "[GuardarRespuestas] Tabla tblRespuestas obtenida con " & tblRespuestas.ListColumns.Count & " columnas"
    
    ' Leer nombres de columnas para verificar estructura
    Dim colIdx As Long
    For colIdx = 1 To tblRespuestas.ListColumns.Count
        ' Debug.Print "  Columna " & colIdx & ": " & tblRespuestas.ListColumns(colIdx).Name
    Next colIdx
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim contador As Long
    contador = 0
    
    ' ====================================================================================
    ' FASE 2: VALIDACIÓN ESTRICTA - Verificar que columna ID Criticidad existe
    ' Fecha: 25/04/2026
    ' Propósito: Prevenir pérdida silenciosa de datos críticos
    ' Si la columna no existe, FALLAR INMEDIATAMENTE con mensaje claro
    ' ====================================================================================
    Dim tieneColumnaCriticidad As Boolean
    tieneColumnaCriticidad = False
    
    ' Buscar columna "ID Criticidad" en la tabla
    On Error Resume Next
    Dim colIdxCriticidad As Variant
    colIdxCriticidad = Application.Match("ID Criticidad", tblRespuestas.HeaderRowRange, 0)
    If Not IsError(colIdxCriticidad) Then
        tieneColumnaCriticidad = True
    End If
    On Error GoTo ErrorHandler
    
    ' Si la columna NO existe, FALLAR con error claro
    If Not tieneColumnaCriticidad Then
        Err.Raise vbObjectError + 1000, "InspectionRepository.GuardarRespuestas", _
                  "ERROR CRÍTICO: La columna 'ID Criticidad' no existe en tblRespuestas." & vbCrLf & vbCrLf & _
                  "El sistema no puede guardar respuestas sin esta columna, ya que afecta los cálculos de:" & vbCrLf & _
                  "  - Auditoría de Procesos (conteos Crítica/Mayor/Menor)" & vbCrLf & _
                  "  - Certificados PDF (sección de no cumplimientos)" & vbCrLf & vbCrLf & _
                  "SOLUCIÓN:" & vbCrLf & _
                  "  1. Abra el archivo Excel" & vbCrLf & _
                  "  2. Ejecute: PlantillaCertificadoSetup.InicializarTablasRequeridas()" & vbCrLf & _
                  "  3. Esto agregará la columna 'ID Criticidad' a tblRespuestas" & vbCrLf & vbCrLf & _
                  "Consulte: docs/INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md para más detalles."
    End If
    
    ' Debug.Print "[GuardarRespuestas] ✓ Validación OK: Columna ID Criticidad existe"
    
    For Each resp In respuestas
        contador = contador + 1
        ' Debug.Print "[GuardarRespuestas] Procesando respuesta " & contador
        
        Dim dictResp As Object
        Set dictResp = resp
        
        ' Debug.Print "  IDPregunta: " & dictResp("IDPregunta")
        ' Debug.Print "  IDOpcion: " & dictResp("IDOpcion")
        ' Debug.Print "  ValorNumerico: " & dictResp("ValorNumerico")
        If dictResp.Exists("IDCriticidad") Then
            ' Debug.Print "  IDCriticidad: " & dictResp("IDCriticidad")
        Else
            ' Debug.Print "  IDCriticidad: (no presente en Dictionary)"
        End If
        
        Set newRow = tblRespuestas.ListRows.Add
        ' Debug.Print "  Nueva fila agregada"
        
        ' Usar índices de columna en lugar de nombres para evitar problemas de acentos
        ' Estructura de tblRespuestas:
        ' [1] ID Respuesta, [2] ID Inspeccion, [3] ID Pregunta, 
        ' [4] ID Opcion, [5] Valor numerico, [6] Observacion, [7] Fecha respuesta
        ' [8] ID Criticidad (NUEVA - ver docs/INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md)
        
        With newRow.Range
            ' Debug.Print "  Estableciendo Columna 1 (ID Respuesta)..."
            .Cells(1, 1).Value = GenerarUUID()
            
            ' Debug.Print "  Estableciendo Columna 2 (ID Inspeccion)..."
            .Cells(1, 2).Value = idInspeccion
            
            ' Debug.Print "  Estableciendo Columna 3 (ID Pregunta)..."
            .Cells(1, 3).Value = dictResp("IDPregunta")
            
            ' Debug.Print "  Estableciendo Columna 4 (ID Opcion)..."
            .Cells(1, 4).Value = dictResp("IDOpcion")
            
            ' Debug.Print "  Estableciendo Columna 5 (Valor numerico)..."
            .Cells(1, 5).Value = dictResp("ValorNumerico")
            
            ' Debug.Print "  Estableciendo Columna 6 (Observacion)..."
            Dim obsValue As String
            obsValue = Trim(CStr(dictResp("Observacion")))
            If Len(obsValue) = 0 Then
                .Cells(1, 6).Value = Configuration2.VALOR_NO_APLICA
            Else
                .Cells(1, 6).Value = obsValue
            End If
            
            ' Debug.Print "  Estableciendo Columna 7 (Fecha respuesta)..."
            .Cells(1, 7).Value = Now
            
            ' ====================================================================================
            ' FASE 2: VALIDACIÓN ESTRICTA - ID Criticidad es OBLIGATORIO
            ' Fecha: 25/04/2026
            ' No se permite guardar respuestas sin ID Criticidad
            ' Si falta, FALLAR con error claro indicando qué pregunta tiene el problema
            ' ====================================================================================
            
            ' Validar que el dato existe en el Dictionary
            If Not dictResp.Exists("IDCriticidad") Then
                ' FALLAR: La pregunta no tiene ID Criticidad asignado
                Dim textoPregunta As String
                textoPregunta = "(Pregunta sin texto)"
                If dictResp.Exists("TextoPregunta") Then
                    textoPregunta = dictResp("TextoPregunta")
                ElseIf dictResp.Exists("IDPregunta") Then
                    textoPregunta = "ID: " & dictResp("IDPregunta")
                End If
                
                Err.Raise vbObjectError + 1001, "InspectionRepository.GuardarRespuestas", _
                          "ERROR: Una pregunta no tiene ID Criticidad asignado." & vbCrLf & vbCrLf & _
                          "Pregunta: " & textoPregunta & vbCrLf & vbCrLf & _
                          "El sistema requiere que TODAS las preguntas tengan un nivel de criticidad asignado:" & vbCrLf & _
                          "  • Crítica: No cumplimientos graves que requieren acción inmediata" & vbCrLf & _
                          "  • Mayor: No cumplimientos importantes" & vbCrLf & _
                          "  • Menor: No cumplimientos de baja prioridad" & vbCrLf & _
                          "  • Ninguna: Preguntas informativas sin impacto" & vbCrLf & vbCrLf & _
                          "SOLUCIÓN:" & vbCrLf & _
                          "  1. Verifique que la plantilla tiene la columna 'ID Criticidad'" & vbCrLf & _
                          "  2. Asegúrese de que TODAS las preguntas tienen un valor en esta columna" & vbCrLf & _
                          "  3. Revise el módulo ChecklistRepository que carga las preguntas"
            End If
            
            ' Validar que el valor NO sea vacío o "N/A"
            Dim idCrit As String
            idCrit = Trim(CStr(dictResp("IDCriticidad")))
            
            If Configuration2.EsValorNoAplica(idCrit) Or Len(idCrit) = 0 Then
                ' FALLAR: El ID Criticidad está vacío
                Dim textoPregunta2 As String
                textoPregunta2 = "(Pregunta sin texto)"
                If dictResp.Exists("TextoPregunta") Then
                    textoPregunta2 = dictResp("TextoPregunta")
                ElseIf dictResp.Exists("IDPregunta") Then
                    textoPregunta2 = "ID: " & dictResp("IDPregunta")
                End If
                
                Err.Raise vbObjectError + 1002, "InspectionRepository.GuardarRespuestas", _
                          "ERROR: Una pregunta tiene ID Criticidad vacío o 'N/A'." & vbCrLf & vbCrLf & _
                          "Pregunta: " & textoPregunta2 & vbCrLf & _
                          "Valor actual: '" & idCrit & "'" & vbCrLf & vbCrLf & _
                          "Valores válidos:" & vbCrLf & _
                          "  • Crítica" & vbCrLf & _
                          "  • Mayor" & vbCrLf & _
                          "  • Menor" & vbCrLf & _
                          "  • Ninguna" & vbCrLf & vbCrLf & _
                          "SOLUCIÓN:" & vbCrLf & _
                          "  1. Abra la hoja 'Checklist' y vaya a la tabla de preguntas" & vbCrLf & _
                          "  2. Busque la pregunta indicada arriba" & vbCrLf & _
                          "  3. Asigne un valor válido en la columna 'ID Criticidad'"
            End If
            
            ' Guardar el valor validado (columna 8)
            .Cells(1, 8).Value = idCrit
            ' Debug.Print "  Estableciendo Columna 8 (ID Criticidad): " & idCrit
        End With
        
        ' Debug.Print "  Respuesta " & contador & " completada"
    Next resp
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    ' Reproteger hoja incluso si hay error (fail-safe)
    On Error Resume Next
    Call SheetProtector2.ApplyRoleBasedProtection(wsHistorico, Configuration2.APP_PASSWORD)
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    On Error GoTo 0
    
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
    
    ' FASE 9 (09/06/2026): Desproteger hoja Histórico para permitir escritura VBA
    Call SheetProtector2.UnprotectSheet(wsHistorico, Configuration2.APP_PASSWORD)
    
    If tblInspecciones.DataBodyRange Is Nothing Then
        ' Reproteger incluso si no hay datos
        Call SheetProtector2.ApplyRoleBasedProtection(wsHistorico, Configuration2.APP_PASSWORD)
        Exit Sub
    End If
    
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
                            ' Debug.Print "[ActualizarCalculos] Auditoría Procesos guardada en columna '" & nombreCol & "' (" & colIndexProcesos & "): " & calculos("Auditoria_Procesos_Resultado")
                            Exit For
                        End If
                    Next nombreCol
                    
                    If IsError(colIndexProcesos) Or IsEmpty(colIndexProcesos) Then
                        ' Debug.Print "[ActualizarCalculos] ADVERTENCIA: No se encontró columna para Auditoría de Procesos. Dato: " & calculos("Auditoria_Procesos_Resultado")
                        ' Debug.Print "[ActualizarCalculos] Columnas existentes: " & tblInspecciones.ListColumns.Count
                    End If
                    On Error GoTo ErrorHandler
                End If
                
                ' Conteos de Auditoría de Procesos (columnas 41, 42, 43)
                If calculos.Exists("AP_Critica_NoCumple") Then
                    On Error Resume Next
                    Dim colAPCritica As Variant
                    colAPCritica = Application.Match("AP Critica No Cumple", tblInspecciones.HeaderRowRange, 0)
                    If Not IsError(colAPCritica) Then
                        .Cells(1, CLng(colAPCritica)).Value = calculos("AP_Critica_NoCumple")
                        ' Debug.Print "[ActualizarCalculos] AP Crítica No Cumple: " & calculos("AP_Critica_NoCumple")
                    End If
                    On Error GoTo ErrorHandler
                End If
                
                If calculos.Exists("AP_Mayor_NoCumple") Then
                    On Error Resume Next
                    Dim colAPMayor As Variant
                    colAPMayor = Application.Match("AP Mayor No Cumple", tblInspecciones.HeaderRowRange, 0)
                    If Not IsError(colAPMayor) Then
                        .Cells(1, CLng(colAPMayor)).Value = calculos("AP_Mayor_NoCumple")
                        ' Debug.Print "[ActualizarCalculos] AP Mayor No Cumple: " & calculos("AP_Mayor_NoCumple")
                    End If
                    On Error GoTo ErrorHandler
                End If
                
                If calculos.Exists("AP_Menor_NoCumple") Then
                    On Error Resume Next
                    Dim colAPMenor As Variant
                    colAPMenor = Application.Match("AP Menor No Cumple", tblInspecciones.HeaderRowRange, 0)
                    If Not IsError(colAPMenor) Then
                        .Cells(1, CLng(colAPMenor)).Value = calculos("AP_Menor_NoCumple")
                        ' Debug.Print "[ActualizarCalculos] AP Menor No Cumple: " & calculos("AP_Menor_NoCumple")
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
                    Dim colIdx As Variant
                    
                    ' Numero Inspeccion (columna 32)
                    colIdx = 0
                    colIdx = Application.Match("Numero Inspeccion", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then
                        .Cells(1, colIdx).Value = calculos("NumeroInspeccion")
                        ' Debug.Print "[ActualizarCalculos] Numero Inspeccion: " & calculos("NumeroInspeccion")
                    End If
                    
                    ' Es Inspeccion Recurrente (columna 33)
                    colIdx = 0
                    colIdx = Application.Match("Es Inspeccion Recurrente", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then
                        Dim valorRecurrente As String
                        valorRecurrente = IIf(calculos("EsInspeccionRecurrente"), "Si", "No")
                        .Cells(1, colIdx).Value = valorRecurrente
                        ' Debug.Print "[ActualizarCalculos] Es Recurrente: " & valorRecurrente
                    End If
                    
                    ' Puesto Evaluado (columna 34)
                    colIdx = 0
                    colIdx = Application.Match("Puesto Evaluado", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then
                        .Cells(1, colIdx).Value = calculos("PuestoEvaluado")
                        ' Debug.Print "[ActualizarCalculos] Puesto Evaluado: " & calculos("PuestoEvaluado")
                    End If
                    
                    ' RPN Total (columna 40) - SIEMPRE guardar (1ra y recurrentes)
                    ' - Primera inspección: RPN Total = % TA
                    ' - Inspección recurrente: RPN Total = Promedio + factores
                    If calculos.Exists("RPNTotal") Then
                        colIdx = 0
                        colIdx = Application.Match("RPN Total", tblInspecciones.HeaderRowRange, 0)
                        If colIdx > 0 Then
                            .Cells(1, colIdx).Value = calculos("RPNTotal")
                            ' Debug.Print "[ActualizarCalculos] RPN Total: " & Format(calculos("RPNTotal"), "0.00")
                        End If
                    End If
                    
                    ' Campos especificos de inspecciones recurrentes (2da+)
                    If calculos("EsInspeccionRecurrente") Then
                        ' Debug.Print "[ActualizarCalculos] Guardando datos recurrentes..."
                        
                        ' RPN Anterior (columna 35) - SIEMPRE guardar, sea manual o automático
                        If calculos.Exists("RPNAnterior") Then
                            colIdx = 0
                            colIdx = Application.Match("RPN Anterior Manual", tblInspecciones.HeaderRowRange, 0)
                            If colIdx > 0 Then
                                .Cells(1, colIdx).Value = calculos("RPNAnterior")
                                ' Debug.Print "[ActualizarCalculos] RPN Anterior: " & Format(calculos("RPNAnterior"), "0.00")
                            End If
                        End If
                        
                        ' ID Inspeccion Anterior (columna 36)
                        ' - Si existe y tiene valor: guardar el ID (modo automático)
                        ' - Si no existe o está vacío: guardar Configuration2.VALOR_NO_APLICA (modo manual)
                        colIdx = 0
                        colIdx = Application.Match("ID Inspeccion Anterior", tblInspecciones.HeaderRowRange, 0)
                        If colIdx > 0 Then
                            If calculos.Exists("IDInspeccionAnterior") And Len(calculos("IDInspeccionAnterior")) > 0 Then
                                .Cells(1, colIdx).Value = calculos("IDInspeccionAnterior")
                                ' Debug.Print "[ActualizarCalculos] ID Inspeccion Anterior: " & calculos("IDInspeccionAnterior")
                            Else
                                .Cells(1, colIdx).Value = Configuration2.VALOR_NO_APLICA
                                ' Debug.Print "[ActualizarCalculos] ID Inspeccion Anterior: N/A (modo manual)"
                            End If
                        End If
                        
                        ' RPN Promedio (columna 37)
                        If calculos.Exists("RPNPromedio") Then
                            colIdx = 0
                            colIdx = Application.Match("RPN Promedio", tblInspecciones.HeaderRowRange, 0)
                            If colIdx > 0 Then
                                .Cells(1, colIdx).Value = calculos("RPNPromedio")
                                ' Debug.Print "[ActualizarCalculos] RPN Promedio: " & Format(calculos("RPNPromedio"), "0.00")
                            End If
                        End If
                        
                        ' Factores adicionales (FASE 6 - 23/04/2026)
                        ' IMPORTANTE: Siempre se guardan en inspecciones recurrentes
                        ' - Operador, Muestreador, Ayudantes → valores reales (Grado A/B)
                        ' - Técnico C/D, Sanitizador → 0 (no requieren factores)
                        ' Esto evita columnas vacías y facilita análisis de datos
                        
                        ' Porcentaje Recuperación (columna 38)
                        colIdx = 0
                        colIdx = Application.Match("Porcentaje Recuperacion", tblInspecciones.HeaderRowRange, 0)
                        If colIdx > 0 Then
                            .Cells(1, colIdx).Value = calculos("PorcRecuperacion")
                            ' Debug.Print "[ActualizarCalculos] % Recuperación: " & Format(calculos("PorcRecuperacion"), "0.00")
                        End If
                        
                        ' Porcentaje OOL (columna 39)
                        colIdx = 0
                        colIdx = Application.Match("Porcentaje OOL", tblInspecciones.HeaderRowRange, 0)
                        If colIdx > 0 Then
                            .Cells(1, colIdx).Value = calculos("PorcOOL")
                            ' Debug.Print "[ActualizarCalculos] % OOL: " & Format(calculos("PorcOOL"), "0.00")
                        End If
                        
                        ' Debug.Print "[ActualizarCalculos] Datos recurrentes guardados OK"
                    End If
                    
                    On Error GoTo ErrorHandler
                Else
                    ' ═══════════════════════════════════════════════════════════════
                    ' PRIMERA INSPECCIÓN (NO RECURRENTE)
                    ' Completar columnas recurrentes con Configuration2.VALOR_NO_APLICA para evitar vacíos
                    ' ═══════════════════════════════════════════════════════════════
                    ' Debug.Print "[ActualizarCalculos] Primera inspección detectada - completando campos recurrentes con 'N/A'"
                    
                    On Error Resume Next
                    colIdx = Application.Match("Numero Inspeccion", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then .Cells(1, colIdx).Value = 1
                    
                    colIdx = Application.Match("Es Inspeccion Recurrente", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then .Cells(1, colIdx).Value = "No"
                    
                    ' RPN Anterior Manual (columna 35)
                    colIdx = Application.Match("RPN Anterior Manual", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then .Cells(1, colIdx).Value = Configuration2.VALOR_NO_APLICA
                    
                    ' ID Inspeccion Anterior (columna 36)
                    colIdx = Application.Match("ID Inspeccion Anterior", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then .Cells(1, colIdx).Value = Configuration2.VALOR_NO_APLICA
                    
                    ' RPN Promedio (columna 37)
                    colIdx = Application.Match("RPN Promedio", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then .Cells(1, colIdx).Value = Configuration2.VALOR_NO_APLICA
                    
                    ' Porcentaje Recuperacion (columna 38)
                    colIdx = Application.Match("Porcentaje Recuperacion", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then .Cells(1, colIdx).Value = Configuration2.VALOR_NO_APLICA
                    
                    ' Porcentaje OOL (columna 39)
                    colIdx = Application.Match("Porcentaje OOL", tblInspecciones.HeaderRowRange, 0)
                    If colIdx > 0 Then .Cells(1, colIdx).Value = Configuration2.VALOR_NO_APLICA
                    
                    ' Debug.Print "[ActualizarCalculos] Campos recurrentes completados con 'N/A'"
                    On Error GoTo ErrorHandler
                End If
                
                ' Estado final
                .Cells(1, tblInspecciones.ListColumns("Estado").Index).Value = Configuration2.INSPECCION_COMPLETADO
                
                ' Auditoría
                .Cells(1, tblInspecciones.ListColumns("Fecha calculo").Index).Value = Now
                .Cells(1, tblInspecciones.ListColumns("Usuario calculo").Index).Value = _
                    .Cells(1, tblInspecciones.ListColumns("Auditor").Index).Value
                
                ' Columnas no usadas actualmente: completar con Configuration2.VALOR_NO_APLICA para evitar vacíos
                On Error Resume Next
                colIdx = Application.Match("Fecha completado", tblInspecciones.HeaderRowRange, 0)
                If colIdx > 0 Then .Cells(1, colIdx).Value = Configuration2.VALOR_NO_APLICA
                
                colIdx = Application.Match("Usuario completado", tblInspecciones.HeaderRowRange, 0)
                If colIdx > 0 Then .Cells(1, colIdx).Value = Configuration2.VALOR_NO_APLICA
                On Error GoTo ErrorHandler
            End With
            
            Exit For
        End If
    Next inspeccionRow
    
    ' Reproteger hoja según el rol del usuario
    Call SheetProtector2.ApplyRoleBasedProtection(wsHistorico, Configuration2.APP_PASSWORD)
    
    Exit Sub
    
ErrorHandler:
    ' Reproteger hoja incluso si hay error (fail-safe)
    On Error Resume Next
    Call SheetProtector2.ApplyRoleBasedProtection(wsHistorico, Configuration2.APP_PASSWORD)
    On Error GoTo 0
    
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
            ' CATEGORÍA 5 (16/06/2026): Leer RPN Total (con factores) para evaluar regla de 3 consecutivas
            inspRPN = CDbl(inspeccionRow.Range.Cells(1, tblInspecciones.ListColumns("RPN Total").Index).Value)
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
