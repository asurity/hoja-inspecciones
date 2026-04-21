' ----------------------------------------------------------------------
' Módulo: ChecklistRepository
' Descripción: Capa de acceso a datos de solo lectura para el checklist virtual.
'              Provee funciones para obtener plantillas, preguntas, opciones,
'              equipos, personal y evaluadores desde las tablas maestras.
' Fecha creación: 14/04/2026
' Dependencias: Configuration2, ErrorLogger2
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Función: ObtenerPlantillaPorPuesto
' Propósito: Busca la plantilla correspondiente a un puesto dado.
' Retorna: Array(ID_Plantilla, Nombre, Frecuencia_meses) o Empty si no existe.
' Nota: Una plantilla contiene preguntas de ambas secciones (Auditoría + TA).
' ----------------------------------------------------------------------
Public Function ObtenerPlantillaPorPuesto(ByVal puesto As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim wsChecklist As Worksheet
    Dim tblPlantillas As ListObject
    Dim plantillaRow As ListRow
    
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set tblPlantillas = wsChecklist.ListObjects(Configuration2.TABLE_PLANTILLAS)
    
    If tblPlantillas.DataBodyRange Is Nothing Then
        ObtenerPlantillaPorPuesto = Empty
        Exit Function
    End If
    
    For Each plantillaRow In tblPlantillas.ListRows
        Dim plantillaPuesto As String
        plantillaPuesto = Trim(plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("Puesto").Index).Value)
        
        If plantillaPuesto = Trim(puesto) Then
            Dim resultado(0 To 2) As Variant
            resultado(0) = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("ID Plantilla").Index).Value
            resultado(1) = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("Nombre de plantilla").Index).Value
            resultado(2) = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("Frecuencia meses").Index).Value
            ObtenerPlantillaPorPuesto = resultado
            Exit Function
        End If
    Next plantillaRow
    
    ObtenerPlantillaPorPuesto = Empty
    Exit Function
    
ErrorHandler:
    ObtenerPlantillaPorPuesto = Empty
    Call ErrorLogger2.Log("ChecklistRepository.ObtenerPlantillaPorPuesto", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerPreguntasPorPlantillaYSeccion
' Propósito: Obtiene las preguntas activas de una plantilla filtradas por sección.
' Retorna: Collection de arrays (ID, Numero, Texto, ID_Seccion, ID_Criticidad).
'          Ordenadas por columna "Orden".
' Parámetros:
'   idPlantilla: ID de la plantilla (FK a tblPlantillas)
'   idSeccion: ID de la sección para filtrar (FK a tblSecciones). "" = todas.
' ----------------------------------------------------------------------
Public Function ObtenerPreguntasPorPlantillaYSeccion(ByVal idPlantilla As String, Optional ByVal idSeccion As String = "") As Collection
    On Error GoTo ErrorHandler
    
    Dim wsChecklist As Worksheet
    Dim tblPreguntas As ListObject
    Dim preguntaRow As ListRow
    Dim resultado As New Collection
    
    Debug.Print "=== ObtenerPreguntasPorPlantillaYSeccion ==="
    Debug.Print "IDPlantilla recibido: [" & idPlantilla & "]"
    Debug.Print "IDSeccion recibido: [" & idSeccion & "]"
    
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set tblPreguntas = wsChecklist.ListObjects(Configuration2.TABLE_PREGUNTAS)
    Debug.Print "Tabla: " & tblPreguntas.Name
    Debug.Print "Columnas disponibles: " & tblPreguntas.ListColumns.Count
    
    ' Listar columnas
    Dim col As ListColumn
    For Each col In tblPreguntas.ListColumns
        Debug.Print "  - Columna " & col.Index & ": " & col.Name
    Next col
    
    If tblPreguntas.DataBodyRange Is Nothing Then
        Debug.Print "ADVERTENCIA: DataBodyRange vacío (no hay filas)"
        Set ObtenerPreguntasPorPlantillaYSeccion = resultado
        Exit Function
    End If
    
    Debug.Print "Filas totales en tabla: " & tblPreguntas.ListRows.Count
    
    ' Recopilar preguntas que coincidan
    Dim tempCol As New Collection
    Dim contadorFiltradas As Long: contadorFiltradas = 0
    
    For Each preguntaRow In tblPreguntas.ListRows
        Dim pID As String
        Dim pPlantilla As String
        Dim pSeccion As String
        Dim pActivo As String
        
        pPlantilla = Trim(preguntaRow.Range.Cells(1, tblPreguntas.ListColumns("ID Plantilla").Index).Value)
        
        If pPlantilla <> Trim(idPlantilla) Then
            Debug.Print "  Pregunta ignorada: IDPlantilla no coincide [" & pPlantilla & "] <> [" & idPlantilla & "]"
            GoTo SiguientePregunta
        End If
        
        pActivo = Trim(preguntaRow.Range.Cells(1, tblPreguntas.ListColumns("Activo").Index).Value)
        If UCase(pActivo) <> "SI" And UCase(pActivo) <> "SÍ" Then
            Debug.Print "  Pregunta ignorada: No está activa [" & pActivo & "]"
            GoTo SiguientePregunta
        End If
        
        pSeccion = Trim(preguntaRow.Range.Cells(1, tblPreguntas.ListColumns("ID Seccion").Index).Value)
        If idSeccion <> "" And pSeccion <> Trim(idSeccion) Then
            Debug.Print "  Pregunta ignorada: IDSeccion no coincide [" & pSeccion & "] <> [" & idSeccion & "]"
            GoTo SiguientePregunta
        End If
        
        contadorFiltradas = contadorFiltradas + 1
        Debug.Print "  Pregunta ACEPTADA #" & contadorFiltradas & ": " & preguntaRow.Range.Cells(1, tblPreguntas.ListColumns("ID Pregunta").Index).Value
        
        Dim preg(0 To 5) As Variant
        preg(0) = preguntaRow.Range.Cells(1, tblPreguntas.ListColumns("ID Pregunta").Index).Value
        preg(1) = preguntaRow.Range.Cells(1, tblPreguntas.ListColumns("Numero").Index).Value
        preg(2) = preguntaRow.Range.Cells(1, tblPreguntas.ListColumns("Texto").Index).Value
        preg(3) = pSeccion
        preg(4) = preguntaRow.Range.Cells(1, tblPreguntas.ListColumns("ID Criticidad").Index).Value
        preg(5) = preguntaRow.Range.Cells(1, tblPreguntas.ListColumns("Orden").Index).Value
        
        tempCol.Add preg
        
SiguientePregunta:
    Next preguntaRow
    
    ' Ordenar por columna Orden (índice 5) usando insertion sort
    Dim sorted As New Collection
    Dim item As Variant
    Dim inserted As Boolean
    
    For Each item In tempCol
        inserted = False
        Dim pos As Long
        For pos = 1 To sorted.Count
            Dim existing As Variant
            existing = sorted(pos)
            If CLng(Val(CStr(item(5)))) < CLng(Val(CStr(existing(5)))) Then
                If pos = 1 Then
                    sorted.Add item, Before:=1
                Else
                    sorted.Add item, Before:=pos
                End If
                inserted = True
                Exit For
            End If
        Next pos
        If Not inserted Then sorted.Add item
    Next item
    
    Debug.Print "Total preguntas encontradas y ordenadas: " & sorted.Count
    Set ObtenerPreguntasPorPlantillaYSeccion = sorted
    Exit Function
    
ErrorHandler:
    Set ObtenerPreguntasPorPlantillaYSeccion = New Collection
    Call ErrorLogger2.Log("ChecklistRepository.ObtenerPreguntasPorPlantillaYSeccion", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerSecciones
' Propósito: Obtiene todas las secciones de tblSecciones.
' Retorna: Collection de arrays (ID, Nombre, TipoRespuesta).
' ----------------------------------------------------------------------
Public Function ObtenerSecciones() As Collection
    On Error GoTo ErrorHandler
    
    Dim wsChecklist As Worksheet
    Dim tblSecciones As ListObject
    Dim seccionRow As ListRow
    Dim resultado As New Collection
    
    Debug.Print "=== ObtenerSecciones ==="
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Debug.Print "Hoja: " & wsChecklist.Name
    Set tblSecciones = wsChecklist.ListObjects(Configuration2.TABLE_SECCIONES)
    Debug.Print "Tabla: " & tblSecciones.Name
    Debug.Print "Columnas disponibles: " & tblSecciones.ListColumns.Count
    
    ' Listar columnas
    Dim col As ListColumn
    For Each col In tblSecciones.ListColumns
        Debug.Print "  - Columna " & col.Index & ": " & col.Name
    Next col
    
    If tblSecciones.DataBodyRange Is Nothing Then
        Debug.Print "ADVERTENCIA: DataBodyRange vacío"
        Set ObtenerSecciones = resultado
        Exit Function
    End If
    
    Debug.Print "Filas en tabla: " & tblSecciones.ListRows.Count
    
    For Each seccionRow In tblSecciones.ListRows
        Dim sec(0 To 2) As Variant
        sec(0) = seccionRow.Range.Cells(1, tblSecciones.ListColumns("ID Seccion").Index).Value
        sec(1) = seccionRow.Range.Cells(1, tblSecciones.ListColumns("Nombre de sección").Index).Value
        sec(2) = seccionRow.Range.Cells(1, tblSecciones.ListColumns("Tipo de respuesta").Index).Value
        resultado.Add sec
        Debug.Print "  Sección: " & sec(0) & " - " & sec(1) & " (" & sec(2) & ")"
    Next seccionRow
    
    Debug.Print "Total secciones: " & resultado.Count
    Set ObtenerSecciones = resultado
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR en ObtenerSecciones: " & Err.Number & " - " & Err.Description
    Set ObtenerSecciones = New Collection
    Call ErrorLogger2.Log("ChecklistRepository.ObtenerSecciones", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerOpcionesRespuesta
' Propósito: Obtiene las opciones de respuesta para una sección y criticidad dadas.
' Retorna: Collection de arrays (ID_Opcion, Texto_opcion, Valor_puntaje).
' Parámetros:
'   idSeccion: ID de la sección para filtrar
'   idCriticidad: (Opcional) ID de criticidad para filtrar. Si es "", no filtra por criticidad
' ----------------------------------------------------------------------
Public Function ObtenerOpcionesRespuesta(ByVal idSeccion As String, Optional ByVal idCriticidad As String = "") As Collection
    On Error GoTo ErrorHandler
    
    Dim wsChecklist As Worksheet
    Dim tblOpciones As ListObject
    Dim opcionRow As ListRow
    Dim resultado As New Collection
    
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set tblOpciones = wsChecklist.ListObjects(Configuration2.TABLE_OPCIONES)
    
    If tblOpciones.DataBodyRange Is Nothing Then
        Set ObtenerOpcionesRespuesta = resultado
        Exit Function
    End If
    
    ' Validar que existen las columnas necesarias antes de usarlas
    Dim colIDSeccion As Long, colIDOpcion As Long, colOpcion As Long, colValorPuntaje As Long, colIDCriticidad As Long
    
    On Error Resume Next
    colIDSeccion = tblOpciones.ListColumns("ID Seccion").Index
    colIDOpcion = tblOpciones.ListColumns("ID Opcion").Index
    colOpcion = tblOpciones.ListColumns("Opción texto").Index
    colValorPuntaje = tblOpciones.ListColumns("Valor puntaje").Index
    colIDCriticidad = tblOpciones.ListColumns("ID Criticidad").Index
    
    If Err.Number <> 0 Then
        Debug.Print "ERROR ObtenerOpcionesRespuesta: Falta columna requerida"
        Err.Clear
        GoTo ErrorHandler
    End If
    On Error GoTo ErrorHandler
    
    For Each opcionRow In tblOpciones.ListRows
        Dim opSeccion As String
        Dim opCriticidad As String
        
        opSeccion = Trim(opcionRow.Range.Cells(1, colIDSeccion).Value)
        opCriticidad = Trim(opcionRow.Range.Cells(1, colIDCriticidad).Value)
        
        ' Filtrar por sección (obligatorio)
        If opSeccion <> Trim(idSeccion) Then GoTo SiguienteOpcion
        
        ' Filtrar por criticidad (si se especificó)
        If Len(idCriticidad) > 0 And opCriticidad <> Trim(idCriticidad) Then GoTo SiguienteOpcion
        
        ' Opción válida - agregar al resultado
        Dim op(0 To 2) As Variant
        op(0) = opcionRow.Range.Cells(1, colIDOpcion).Value
        op(1) = opcionRow.Range.Cells(1, colOpcion).Value
        op(2) = opcionRow.Range.Cells(1, colValorPuntaje).Value
        resultado.Add op
        
SiguienteOpcion:
    Next opcionRow
    Set ObtenerOpcionesRespuesta = resultado
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR en ObtenerOpcionesRespuesta: " & Err.Number & " - " & Err.Description
    Set ObtenerOpcionesRespuesta = New Collection
    Call ErrorLogger2.Log("ChecklistRepository.ObtenerOpcionesRespuesta", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerAreasPorPlanta
' Propósito: Obtiene las áreas únicas de tblEquipos filtradas por planta.
' Retorna: Collection de strings (nombres de áreas únicas).
' ----------------------------------------------------------------------
Public Function ObtenerAreasPorPlanta(ByVal planta As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim tblEquipos As ListObject
    Dim equipoRow As ListRow
    Dim resultado As New Collection
    Dim areasVistas As Object
    
    Set areasVistas = CreateObject("Scripting.Dictionary")
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tblEquipos = wsConfig.ListObjects(Configuration2.TABLE_EQUIPOS)
    
    If tblEquipos.DataBodyRange Is Nothing Then
        Set ObtenerAreasPorPlanta = resultado
        Exit Function
    End If
    
    For Each equipoRow In tblEquipos.ListRows
        Dim eqPlanta As String
        Dim eqArea As String
        
        eqPlanta = Trim(equipoRow.Range.Cells(1, tblEquipos.ListColumns("Planta").Index).Value)
        eqArea = Trim(equipoRow.Range.Cells(1, tblEquipos.ListColumns("Área").Index).Value)
        
        If eqPlanta = Trim(planta) And eqArea <> "" Then
            If Not areasVistas.Exists(eqArea) Then
                areasVistas.Add eqArea, True
                resultado.Add eqArea
            End If
        End If
    Next equipoRow
    
    Set ObtenerAreasPorPlanta = resultado
    Exit Function
    
ErrorHandler:
    Set ObtenerAreasPorPlanta = New Collection
    Call ErrorLogger2.Log("ChecklistRepository.ObtenerAreasPorPlanta", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerEquiposPorPlantaYArea
' Propósito: Obtiene los equipos/salas filtrados por planta y área.
' Retorna: Collection de strings (nombres de equipos).
' ----------------------------------------------------------------------
Public Function ObtenerEquiposPorPlantaYArea(ByVal planta As String, ByVal area As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim tblEquipos As ListObject
    Dim equipoRow As ListRow
    Dim resultado As New Collection
    
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tblEquipos = wsConfig.ListObjects(Configuration2.TABLE_EQUIPOS)
    
    If tblEquipos.DataBodyRange Is Nothing Then
        Set ObtenerEquiposPorPlantaYArea = resultado
        Exit Function
    End If
    
    For Each equipoRow In tblEquipos.ListRows
        Dim eqPlanta As String
        Dim eqArea As String
        Dim eqEquipo As String
        
        eqPlanta = Trim(equipoRow.Range.Cells(1, tblEquipos.ListColumns("Planta").Index).Value)
        eqArea = Trim(equipoRow.Range.Cells(1, tblEquipos.ListColumns("Área").Index).Value)
        eqEquipo = Trim(equipoRow.Range.Cells(1, tblEquipos.ListColumns("Equipo").Index).Value)
        
        If eqPlanta = Trim(planta) And eqArea = Trim(area) And eqEquipo <> "" Then
            resultado.Add eqEquipo
        End If
    Next equipoRow
    
    Set ObtenerEquiposPorPlantaYArea = resultado
    Exit Function
    
ErrorHandler:
    Set ObtenerEquiposPorPlantaYArea = New Collection
    Call ErrorLogger2.Log("ChecklistRepository.ObtenerEquiposPorPlantaYArea", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerPersonalPorPuestoYPlanta
' Propósito: Obtiene las iniciales de personal activo que tiene un puesto
'            específico activo y pertenece a una planta dada.
' Parámetros:
'   nombreColumnaPuesto: Nombre exacto de la columna de puesto en tblPersonal
'                        (ej: "Ayudante 2", "Operador")
'   planta: Nombre de la planta para filtrar
' Retorna: Collection de strings (iniciales).
' ----------------------------------------------------------------------
Public Function ObtenerPersonalPorPuestoYPlanta(ByVal nombreColumnaPuesto As String, ByVal planta As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim wsPersonal As Worksheet
    Dim tblPersonal As ListObject
    Dim personalRow As ListRow
    Dim resultado As New Collection
    
    Set wsPersonal = ThisWorkbook.Sheets(Configuration2.SHEET_PERSONAL)
    Set tblPersonal = wsPersonal.ListObjects(Configuration2.TABLE_PERSONAL)
    
    If tblPersonal.DataBodyRange Is Nothing Then
        Set ObtenerPersonalPorPuestoYPlanta = resultado
        Exit Function
    End If
    
    ' Verificar que la columna de puesto existe
    Dim colIndex As Long
    On Error Resume Next
    colIndex = tblPersonal.ListColumns(nombreColumnaPuesto).Index
    On Error GoTo ErrorHandler
    If colIndex = 0 Then
        Set ObtenerPersonalPorPuestoYPlanta = resultado
        Exit Function
    End If
    
    For Each personalRow In tblPersonal.ListRows
        Dim pPlanta As String
        Dim pActivo As String
        Dim pPuesto As String
        Dim pIniciales As String
        
        pPlanta = Trim(personalRow.Range.Cells(1, tblPersonal.ListColumns("Planta").Index).Value)
        pActivo = Trim(personalRow.Range.Cells(1, tblPersonal.ListColumns("Activo").Index).Value)
        pPuesto = Trim(personalRow.Range.Cells(1, colIndex).Value)
        
        If pPlanta = Trim(planta) And UCase(pActivo) = "SI" And UCase(pPuesto) = "SI" Then
            pIniciales = Trim(personalRow.Range.Cells(1, tblPersonal.ListColumns("Iniciales").Index).Value)
            If pIniciales <> "" Then resultado.Add pIniciales
        End If
    Next personalRow
    
    Set ObtenerPersonalPorPuestoYPlanta = resultado
    Exit Function
    
ErrorHandler:
    Set ObtenerPersonalPorPuestoYPlanta = New Collection
    Call ErrorLogger2.Log("ChecklistRepository.ObtenerPersonalPorPuestoYPlanta", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerEvaluadores
' Propósito: Obtiene las iniciales de los evaluadores autorizados
'            desde tblAseguramientoCalidad (columna 3).
' Retorna: Collection de strings (iniciales de evaluadores).
' ----------------------------------------------------------------------
Public Function ObtenerEvaluadores() As Collection
    On Error GoTo ErrorHandler
    
    Dim wsAseg As Worksheet
    Dim tblAseg As ListObject
    Dim asegRow As ListRow
    Dim resultado As New Collection
    
    Set wsAseg = ThisWorkbook.Sheets(Configuration2.SHEET_ASEGURAMIENTO)
    Set tblAseg = wsAseg.ListObjects(Configuration2.TABLE_ASEGURAMIENTO)
    
    If tblAseg.DataBodyRange Is Nothing Then
        Set ObtenerEvaluadores = resultado
        Exit Function
    End If
    
    For Each asegRow In tblAseg.ListRows
        Dim iniciales As String
        ' Columna 3 contiene las iniciales según el usuario
        iniciales = Trim(asegRow.Range.Cells(1, 3).Value)
        If iniciales <> "" Then resultado.Add iniciales
    Next asegRow
    
    Set ObtenerEvaluadores = resultado
    Exit Function
    
ErrorHandler:
    Set ObtenerEvaluadores = New Collection
    Call ErrorLogger2.Log("ChecklistRepository.ObtenerEvaluadores", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerMapeoPuestosIniciales
' Propósito: Obtiene el mapeo Puesto → Sigla desde tblPuestosIniciales.
'            Ejemplo: "Ayudante 2" → "AY2"
' Retorna: Dictionary con Key=NombrePuesto, Value=Sigla
' ----------------------------------------------------------------------
Public Function ObtenerMapeoPuestosIniciales() As Object
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim tblPuestos As ListObject
    Dim puestoRow As ListRow
    Dim resultado As Object
    
    Set resultado = CreateObject("Scripting.Dictionary")
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tblPuestos = wsConfig.ListObjects(Configuration2.TABLE_PUESTOS_INICIALES)
    
    If tblPuestos.DataBodyRange Is Nothing Then
        Set ObtenerMapeoPuestosIniciales = resultado
        Exit Function
    End If
    
    For Each puestoRow In tblPuestos.ListRows
        Dim nombrePuesto As String
        Dim sigla As String
        
        ' Leer dinámicamente desde nombres de columnas
        On Error Resume Next
        nombrePuesto = Trim(puestoRow.Range.Cells(1, tblPuestos.ListColumns("Puesto").Index).Value)
        sigla = Trim(puestoRow.Range.Cells(1, tblPuestos.ListColumns("Sigla").Index).Value)
        On Error GoTo 0
        
        If nombrePuesto <> "" And sigla <> "" Then
            If Not resultado.Exists(nombrePuesto) Then
                resultado.Add nombrePuesto, sigla
            End If
        End If
    Next puestoRow
    
    Set ObtenerMapeoPuestosIniciales = resultado
    Exit Function
    
ErrorHandler:
    Set ObtenerMapeoPuestosIniciales = CreateObject("Scripting.Dictionary")
    Call ErrorLogger2.Log("ChecklistRepository.ObtenerMapeoPuestosIniciales", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerPlantaPersonal
' Propósito: Obtiene la planta de una persona por sus iniciales.
' Retorna: String con nombre de planta o "" si no se encuentra.
' ----------------------------------------------------------------------
Public Function ObtenerPlantaPersonal(ByVal iniciales As String) As String
    On Error GoTo ErrorHandler
    
    Dim wsPersonal As Worksheet
    Dim tblPersonal As ListObject
    Dim personalRow As ListRow
    
    Set wsPersonal = ThisWorkbook.Sheets(Configuration2.SHEET_PERSONAL)
    Set tblPersonal = wsPersonal.ListObjects(Configuration2.TABLE_PERSONAL)
    
    If tblPersonal.DataBodyRange Is Nothing Then
        ObtenerPlantaPersonal = ""
        Exit Function
    End If
    
    For Each personalRow In tblPersonal.ListRows
        Dim pIniciales As String
        pIniciales = Trim(personalRow.Range.Cells(1, tblPersonal.ListColumns("Iniciales").Index).Value)
        
        If pIniciales = Trim(iniciales) Then
            ObtenerPlantaPersonal = Trim(personalRow.Range.Cells(1, tblPersonal.ListColumns("Planta").Index).Value)
            Exit Function
        End If
    Next personalRow
    
    ObtenerPlantaPersonal = ""
    Exit Function
    
ErrorHandler:
    ObtenerPlantaPersonal = ""
    Call ErrorLogger2.Log("ChecklistRepository.ObtenerPlantaPersonal", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerIDOpcionPorTexto
' Propósito: Busca el ID de una opción de respuesta por su texto y sección.
' Retorna: Array(ID_Opcion, Valor_puntaje) o Empty si no se encuentra.
' ----------------------------------------------------------------------
Public Function ObtenerIDOpcionPorTexto(ByVal idSeccion As String, ByVal textoOpcion As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim wsChecklist As Worksheet
    Dim tblOpciones As ListObject
    Dim opcionRow As ListRow
    
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set tblOpciones = wsChecklist.ListObjects(Configuration2.TABLE_OPCIONES)
    
    If tblOpciones.DataBodyRange Is Nothing Then
        ObtenerIDOpcionPorTexto = Empty
        Exit Function
    End If
    
    For Each opcionRow In tblOpciones.ListRows
        Dim opSeccion As String
        Dim opTexto As String
        
        opSeccion = Trim(opcionRow.Range.Cells(1, tblOpciones.ListColumns("ID Seccion").Index).Value)
        opTexto = Trim(opcionRow.Range.Cells(1, tblOpciones.ListColumns("Opción texto").Index).Value)
        
        If opSeccion = Trim(idSeccion) And opTexto = Trim(textoOpcion) Then
            Dim res(0 To 1) As Variant
            res(0) = opcionRow.Range.Cells(1, tblOpciones.ListColumns("ID Opcion").Index).Value
            res(1) = opcionRow.Range.Cells(1, tblOpciones.ListColumns("Valor puntaje").Index).Value
            ObtenerIDOpcionPorTexto = res
            Exit Function
        End If
    Next opcionRow
    
    ObtenerIDOpcionPorTexto = Empty
    Exit Function
    
ErrorHandler:
    ObtenerIDOpcionPorTexto = Empty
    Call ErrorLogger2.Log("ChecklistRepository.ObtenerIDOpcionPorTexto", Err.Description, Err.Number)
End Function
