' ======================================================================
' Módulo: TableValidator
' Descripción: Valida la integridad referencial entre las 5 tablas maestras.
'              Se ejecuta ANTES de guardar (preventivo) y bajo demanda
'              para detectar datos huérfanos o inconsistentes.
'
' Tablas validadas:
'   Nivel 1: tblCriticidad, tblSecciones, tblPlantillas (sin dependencias)
'   Nivel 2: tblOpcionesDeRespuesta (depende de Nivel 1)
'   Nivel 3: tblPreguntas (depende de Nivel 1 y 2)
'
' Dependencias: TableManager.bas, Configuration2.bas
' ======================================================================
Option Explicit

' ======================================================================
' SECCIÓN 1: VALIDACIÓN PRE-GUARDADO (usada por frmGestorTablas)
' ======================================================================

' ----------------------------------------------------------------------
' ValidarDatosParaGuardar
' Propósito: Valida los datos ANTES de guardarlos en la tabla.
'            Verifica campos obligatorios, formatos y referencias.
' Parámetros:
'   nombreLogico - tabla destino ("CRITICIDAD", "SECCIONES", etc.)
'   datos - Dictionary con los datos del formulario
'   modoOperacion - "NUEVO" o "EDITAR"
' Retorna: String vacío si todo OK, o mensaje de error detallado.
' ----------------------------------------------------------------------
Public Function ValidarDatosParaGuardar(ByVal nombreLogico As String, _
                                         ByVal datos As Object, _
                                         ByVal modoOperacion As String) As String
    Dim errores As String
    errores = ""
    
    Select Case UCase(nombreLogico)
        Case "CRITICIDAD"
            errores = ValidarCriticidad(datos, modoOperacion)
        Case "SECCIONES"
            errores = ValidarSeccion(datos, modoOperacion)
        Case "PLANTILLAS"
            errores = ValidarPlantilla(datos, modoOperacion)
        Case "OPCIONES"
            errores = ValidarOpcionRespuesta(datos, modoOperacion)
        Case "PREGUNTAS"
            errores = ValidarPregunta(datos, modoOperacion)
    End Select
    
    ValidarDatosParaGuardar = errores
End Function

' --- Validación individual por tabla ---

Private Function ValidarCriticidad(ByVal datos As Object, ByVal modo As String) As String
    Dim e As String: e = ""
    
    ' Campo obligatorio: Nombre
    If Not datos.Exists("Nombre") Or Trim(CStr(datos("Nombre"))) = "" Then
        e = e & "- El nombre de criticidad es obligatorio." & vbCrLf
    End If
    
    ' Campo obligatorio: Valor (debe ser numérico)
    If datos.Exists("Valor") Then
        If Trim(CStr(datos("Valor"))) <> "" And Not IsNumeric(datos("Valor")) Then
            e = e & "- El valor debe ser numérico." & vbCrLf
        End If
    End If
    
    ' Unicidad: verificar que no exista otro con el mismo nombre
    If modo = "NUEVO" And datos.Exists("Nombre") Then
        If ExisteValorEnColumna("CRITICIDAD", "Nombre de criticidad", CStr(datos("Nombre"))) Then
            e = e & "- Ya existe una criticidad con ese nombre." & vbCrLf
        End If
    End If
    
    ValidarCriticidad = e
End Function

Private Function ValidarSeccion(ByVal datos As Object, ByVal modo As String) As String
    Dim e As String: e = ""
    
    If Not datos.Exists("Nombre") Or Trim(CStr(datos("Nombre"))) = "" Then
        e = e & "- El nombre de sección es obligatorio." & vbCrLf
    End If
    
    ' Tipo de respuesta debe ser válido
    If datos.Exists("TipoRespuesta") Then
        Dim tipo As String
        tipo = Trim(CStr(datos("TipoRespuesta")))
        If tipo <> "" And tipo <> "Selección" And tipo <> "Puntaje" Then
            e = e & "- Tipo de respuesta debe ser 'Selección' o 'Puntaje'." & vbCrLf
        End If
    End If
    
    If modo = "NUEVO" And datos.Exists("Nombre") Then
        If ExisteValorEnColumna("SECCIONES", "Nombre de sección", CStr(datos("Nombre"))) Then
            e = e & "- Ya existe una sección con ese nombre." & vbCrLf
        End If
    End If
    
    ValidarSeccion = e
End Function

Private Function ValidarPlantilla(ByVal datos As Object, ByVal modo As String) As String
    Dim e As String: e = ""
    
    If Not datos.Exists("Nombre") Or Trim(CStr(datos("Nombre"))) = "" Then
        e = e & "- El nombre de plantilla es obligatorio." & vbCrLf
    End If
    
    ' Frecuencia debe ser numérica positiva
    If datos.Exists("Frecuencia") Then
        Dim freq As String
        freq = Trim(CStr(datos("Frecuencia")))
        If freq <> "" Then
            If Not IsNumeric(freq) Then
                e = e & "- La frecuencia debe ser un número." & vbCrLf
            ElseIf CDbl(freq) <= 0 Then
                e = e & "- La frecuencia debe ser mayor a 0." & vbCrLf
            End If
        End If
    End If
    
    If modo = "NUEVO" And datos.Exists("Nombre") Then
        If ExisteValorEnColumna("PLANTILLAS", "Nombre de plantilla", CStr(datos("Nombre"))) Then
            e = e & "- Ya existe una plantilla con ese nombre." & vbCrLf
        End If
    End If
    
    ValidarPlantilla = e
End Function

Private Function ValidarOpcionRespuesta(ByVal datos As Object, ByVal modo As String) As String
    Dim e As String: e = ""
    
    ' Referencia obligatoria: Sección debe existir
    If Not datos.Exists("IDSeccion") Or Trim(CStr(datos("IDSeccion"))) = "" Then
        e = e & "- Debe seleccionar una Sección." & vbCrLf
    Else
        If Not ExisteIDEnTabla("SECCIONES", CStr(datos("IDSeccion"))) Then
            e = e & "- La Sección seleccionada no existe en tblSecciones." & vbCrLf
        End If
    End If
    
    ' Referencia opcional: Criticidad (si se proporciona, debe existir)
    If datos.Exists("IDCriticidad") Then
        If Trim(CStr(datos("IDCriticidad"))) <> "" Then
            If Not ExisteIDEnTabla("CRITICIDAD", CStr(datos("IDCriticidad"))) Then
                e = e & "- La Criticidad seleccionada no existe en tblCriticidad." & vbCrLf
            End If
        End If
    End If
    
    ' Texto de opción obligatorio
    If Not datos.Exists("TextoOpcion") Or Trim(CStr(datos("TextoOpcion"))) = "" Then
        e = e & "- El texto de opción es obligatorio." & vbCrLf
    End If
    
    ValidarOpcionRespuesta = e
End Function

Private Function ValidarPregunta(ByVal datos As Object, ByVal modo As String) As String
    Dim e As String: e = ""
    
    ' Referencia obligatoria: Plantilla
    If Not datos.Exists("IDPlantilla") Or Trim(CStr(datos("IDPlantilla"))) = "" Then
        e = e & "- Debe seleccionar una Plantilla." & vbCrLf
    Else
        If Not ExisteIDEnTabla("PLANTILLAS", CStr(datos("IDPlantilla"))) Then
            e = e & "- La Plantilla seleccionada no existe en tblPlantillas." & vbCrLf
        End If
    End If
    
    ' Referencia obligatoria: Sección
    If Not datos.Exists("IDSeccion") Or Trim(CStr(datos("IDSeccion"))) = "" Then
        e = e & "- Debe seleccionar una Sección." & vbCrLf
    Else
        If Not ExisteIDEnTabla("SECCIONES", CStr(datos("IDSeccion"))) Then
            e = e & "- La Sección seleccionada no existe en tblSecciones." & vbCrLf
        End If
    End If
    
    ' Referencia obligatoria: Criticidad
    If Not datos.Exists("IDCriticidad") Or Trim(CStr(datos("IDCriticidad"))) = "" Then
        e = e & "- Debe seleccionar una Criticidad." & vbCrLf
    Else
        If Not ExisteIDEnTabla("CRITICIDAD", CStr(datos("IDCriticidad"))) Then
            e = e & "- La Criticidad seleccionada no existe en tblCriticidad." & vbCrLf
        End If
    End If
    
    ' Texto de pregunta obligatorio
    If Not datos.Exists("TextoPregunta") Or Trim(CStr(datos("TextoPregunta"))) = "" Then
        e = e & "- El texto de la pregunta es obligatorio." & vbCrLf
    End If
    
    ' Orden debe ser numérico
    If datos.Exists("Orden") Then
        Dim ord As String
        ord = Trim(CStr(datos("Orden")))
        If ord <> "" And Not IsNumeric(ord) Then
            e = e & "- El orden debe ser un número." & vbCrLf
        End If
    End If
    
    ' Validar que existan opciones de respuesta para la combinación sección+criticidad
    If datos.Exists("IDSeccion") And datos.Exists("IDCriticidad") Then
        If Trim(CStr(datos("IDSeccion"))) <> "" And Trim(CStr(datos("IDCriticidad"))) <> "" Then
            If Not ExistenOpcionesParaCombinacion(CStr(datos("IDSeccion")), CStr(datos("IDCriticidad"))) Then
                e = e & "- ADVERTENCIA: No existen opciones de respuesta para esta " & _
                         "combinación Sección+Criticidad. La pregunta no podrá " & _
                         "calificarse hasta que se definan opciones en Nivel 2." & vbCrLf
            End If
        End If
    End If
    
    ValidarPregunta = e
End Function

' ======================================================================
' SECCIÓN 2: VERIFICACIÓN DE DEPENDENCIAS (usada antes de eliminar)
' ======================================================================

' ----------------------------------------------------------------------
' VerificarDependencias
' Propósito: Verifica si un registro es referenciado por otras tablas.
'            Se ejecuta ANTES de permitir eliminación.
' Parámetros:
'   nombreLogico - tabla del registro a eliminar
'   idRegistro - ID del registro
' Retorna: String vacío si no tiene dependencias, o detalle de dependencias.
' ----------------------------------------------------------------------
Public Function VerificarDependencias(ByVal nombreLogico As String, _
                                       ByVal idRegistro As String) As String
    Dim deps As String: deps = ""
    
    Select Case UCase(nombreLogico)
        Case "CRITICIDAD"
            ' Criticidad es referenciada por: tblOpcionesDeRespuesta y tblPreguntas
            Dim nOpcCrit As Long
            nOpcCrit = ContarReferenciasEnTabla("OPCIONES", "ID Criticidad", idRegistro)
            If nOpcCrit > 0 Then
                deps = deps & "- tblOpcionesDeRespuesta: " & nOpcCrit & " opciones usan esta criticidad." & vbCrLf
            End If
            
            Dim nPregCrit As Long
            nPregCrit = ContarReferenciasEnTabla("PREGUNTAS", "ID Criticidad", idRegistro)
            If nPregCrit > 0 Then
                deps = deps & "- tblPreguntas: " & nPregCrit & " preguntas usan esta criticidad." & vbCrLf
            End If
            
        Case "SECCIONES"
            ' Secciones es referenciada por: tblOpcionesDeRespuesta y tblPreguntas
            Dim nOpcSec As Long
            nOpcSec = ContarReferenciasEnTabla("OPCIONES", "ID Seccion", idRegistro)
            If nOpcSec > 0 Then
                deps = deps & "- tblOpcionesDeRespuesta: " & nOpcSec & " opciones usan esta sección." & vbCrLf
            End If
            
            Dim nPregSec As Long
            nPregSec = ContarReferenciasEnTabla("PREGUNTAS", "ID Sección", idRegistro)
            If nPregSec > 0 Then
                deps = deps & "- tblPreguntas: " & nPregSec & " preguntas usan esta sección." & vbCrLf
            End If
            
        Case "PLANTILLAS"
            ' Plantillas es referenciada por: tblPreguntas
            Dim nPregPlt As Long
            nPregPlt = ContarReferenciasEnTabla("PREGUNTAS", "ID Plantilla", idRegistro)
            If nPregPlt > 0 Then
                deps = deps & "- tblPreguntas: " & nPregPlt & " preguntas pertenecen a esta plantilla." & vbCrLf
            End If
            
        Case "OPCIONES"
            ' Opciones no es referenciada directamente por otras tablas del MVP
            ' (Las respuestas de inspecciones futuras la referenciarán)
            deps = ""
            
        Case "PREGUNTAS"
            ' Preguntas puede ser referenciada por tblRespuestas (futuro)
            ' Por ahora, siempre se hace soft-delete
            deps = "- Regla arquitectónica: Las preguntas no se eliminan físicamente." & vbCrLf & _
                   "  Se marcan como Activo='No' para preservar auditoría histórica."
    End Select
    
    VerificarDependencias = deps
End Function

' ======================================================================
' SECCIÓN 3: VALIDACIÓN COMPLETA DEL SISTEMA
' ======================================================================

' ----------------------------------------------------------------------
' ValidarIntegridad
' Propósito: Ejecuta validación completa de todas las tablas.
'            Detecta registros huérfanos, referencias rotas y
'            datos inconsistentes.
' Retorna: String con reporte de errores, o vacío si todo OK.
' Uso: Botón "Validar Todo" en frmGestorTablas.
' ----------------------------------------------------------------------
Public Function ValidarIntegridad() As String
    Dim reporte As String: reporte = ""
    Dim erroresNivel2 As String
    Dim erroresNivel3 As String
    
    ' --- Validar Nivel 1: Tablas básicas (existencia y estructura) ---
    reporte = reporte & ValidarEstructuraTabla("CRITICIDAD", Array("ID Criticidad", "Nombre de criticidad", "Valor"))
    reporte = reporte & ValidarEstructuraTabla("SECCIONES", Array("ID Seccion", "Nombre de sección", "Tipo de respuesta"))
    reporte = reporte & ValidarEstructuraTabla("PLANTILLAS", Array("ID Plantilla", "Nombre de plantilla"))
    
    ' --- Validar Nivel 2: tblOpcionesDeRespuesta ---
    erroresNivel2 = ValidarReferenciasOpcionesRespuesta()
    If erroresNivel2 <> "" Then
        reporte = reporte & "=== tblOpcionesDeRespuesta ===" & vbCrLf & erroresNivel2
    End If
    
    ' --- Validar Nivel 3: tblPreguntas ---
    erroresNivel3 = ValidarReferenciasPreguntas()
    If erroresNivel3 <> "" Then
        reporte = reporte & "=== tblPreguntas ===" & vbCrLf & erroresNivel3
    End If
    
    ' --- Validar IDs duplicados en todas las tablas ---
    Dim erroresDups As String
    erroresDups = ValidarIDsDuplicados()
    If erroresDups <> "" Then
        reporte = reporte & "=== IDs Duplicados ===" & vbCrLf & erroresDups
    End If
    
    ValidarIntegridad = reporte
End Function

' --- Validar estructura de una tabla ---
Private Function ValidarEstructuraTabla(ByVal nombreLogico As String, _
                                         ByVal columnasRequeridas As Variant) As String
    Dim tbl As ListObject
    Set tbl = TableManager.ObtenerListObject(nombreLogico)
    
    Dim e As String: e = ""
    
    If tbl Is Nothing Then
        e = "ERROR CRÍTICO: La tabla " & nombreLogico & " no existe." & vbCrLf
        ValidarEstructuraTabla = e
        Exit Function
    End If
    
    Dim i As Long
    For i = LBound(columnasRequeridas) To UBound(columnasRequeridas)
        If TableManager.ObtenerIndiceColumna(tbl, CStr(columnasRequeridas(i))) = 0 Then
            e = e & "- " & nombreLogico & ": Falta la columna '" & columnasRequeridas(i) & "'." & vbCrLf
        End If
    Next i
    
    ValidarEstructuraTabla = e
End Function

' --- Validar referencias en tblOpcionesDeRespuesta ---
Private Function ValidarReferenciasOpcionesRespuesta() As String
    Dim tbl As ListObject
    Set tbl = TableManager.ObtenerListObject("OPCIONES")
    
    Dim e As String: e = ""
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
        ValidarReferenciasOpcionesRespuesta = ""
        Exit Function
    End If
    
    Dim colSeccion As Long
    colSeccion = TableManager.ObtenerIndiceColumna(tbl, "ID Seccion")
    
    Dim colCriticidad As Long
    colCriticidad = TableManager.ObtenerIndiceColumna(tbl, "ID Criticidad")
    
    Dim fila As Long
    For fila = 1 To tbl.ListRows.Count
        ' Verificar que la Sección referenciada existe
        If colSeccion > 0 Then
            Dim idSec As String
            idSec = CStr(tbl.ListRows(fila).Range.Cells(1, colSeccion).Value)
            If idSec <> "" And Not ExisteIDEnTabla("SECCIONES", idSec) Then
                e = e & "- Fila " & fila & ": ID Sección '" & idSec & "' no existe en tblSecciones." & vbCrLf
            End If
        End If
        
        ' Verificar que la Criticidad referenciada existe
        If colCriticidad > 0 Then
            Dim idCrit As String
            idCrit = CStr(tbl.ListRows(fila).Range.Cells(1, colCriticidad).Value)
            If idCrit <> "" And Not ExisteIDEnTabla("CRITICIDAD", idCrit) Then
                e = e & "- Fila " & fila & ": ID Criticidad '" & idCrit & "' no existe en tblCriticidad." & vbCrLf
            End If
        End If
    Next fila
    
    ValidarReferenciasOpcionesRespuesta = e
End Function

' --- Validar referencias en tblPreguntas ---
Private Function ValidarReferenciasPreguntas() As String
    Dim tbl As ListObject
    Set tbl = TableManager.ObtenerListObject("PREGUNTAS")
    
    Dim e As String: e = ""
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
        ValidarReferenciasPreguntas = ""
        Exit Function
    End If
    
    Dim colPlantilla As Long: colPlantilla = TableManager.ObtenerIndiceColumna(tbl, "ID Plantilla")
    Dim colSeccion As Long: colSeccion = TableManager.ObtenerIndiceColumna(tbl, "ID Sección")
    Dim colCriticidad As Long: colCriticidad = TableManager.ObtenerIndiceColumna(tbl, "ID Criticidad")
    
    Dim fila As Long
    For fila = 1 To tbl.ListRows.Count
        ' Verificar Plantilla
        If colPlantilla > 0 Then
            Dim idPlt As String
            idPlt = CStr(tbl.ListRows(fila).Range.Cells(1, colPlantilla).Value)
            If idPlt <> "" And Not ExisteIDEnTabla("PLANTILLAS", idPlt) Then
                e = e & "- Fila " & fila & ": ID Plantilla '" & idPlt & "' no existe en tblPlantillas." & vbCrLf
            End If
        End If
        
        ' Verificar Sección
        If colSeccion > 0 Then
            Dim idSec2 As String
            idSec2 = CStr(tbl.ListRows(fila).Range.Cells(1, colSeccion).Value)
            If idSec2 <> "" And Not ExisteIDEnTabla("SECCIONES", idSec2) Then
                e = e & "- Fila " & fila & ": ID Sección '" & idSec2 & "' no existe en tblSecciones." & vbCrLf
            End If
        End If
        
        ' Verificar Criticidad
        If colCriticidad > 0 Then
            Dim idCrit2 As String
            idCrit2 = CStr(tbl.ListRows(fila).Range.Cells(1, colCriticidad).Value)
            If idCrit2 <> "" And Not ExisteIDEnTabla("CRITICIDAD", idCrit2) Then
                e = e & "- Fila " & fila & ": ID Criticidad '" & idCrit2 & "' no existe en tblCriticidad." & vbCrLf
            End If
        End If
    Next fila
    
    ValidarReferenciasPreguntas = e
End Function

' --- Validar IDs duplicados ---
Private Function ValidarIDsDuplicados() As String
    Dim e As String: e = ""
    
    e = e & VerificarDuplicadosEnTabla("CRITICIDAD", "ID Criticidad")
    e = e & VerificarDuplicadosEnTabla("SECCIONES", "ID Seccion")
    e = e & VerificarDuplicadosEnTabla("PLANTILLAS", "ID Plantilla")
    e = e & VerificarDuplicadosEnTabla("OPCIONES", "ID Opcion")
    e = e & VerificarDuplicadosEnTabla("PREGUNTAS", "ID Pregunta")
    
    ValidarIDsDuplicados = e
End Function

Private Function VerificarDuplicadosEnTabla(ByVal nombreLogico As String, _
                                             ByVal nombreColumnaID As String) As String
    Dim tbl As ListObject
    Set tbl = TableManager.ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Or tbl.ListRows.Count <= 1 Then
        VerificarDuplicadosEnTabla = ""
        Exit Function
    End If
    
    Dim colID As Long
    colID = TableManager.ObtenerIndiceColumna(tbl, nombreColumnaID)
    If colID = 0 Then
        VerificarDuplicadosEnTabla = ""
        Exit Function
    End If
    
    Dim e As String: e = ""
    Dim ids As Object
    Set ids = CreateObject("Scripting.Dictionary")
    
    Dim fila As Long
    For fila = 1 To tbl.ListRows.Count
        Dim idVal As String
        idVal = CStr(tbl.ListRows(fila).Range.Cells(1, colID).Value)
        If idVal <> "" Then
            If ids.Exists(idVal) Then
                e = e & "- " & nombreLogico & ": ID duplicado '" & idVal & "' en filas " & ids(idVal) & " y " & fila & "." & vbCrLf
            Else
                ids.Add idVal, fila
            End If
        End If
    Next fila
    
    VerificarDuplicadosEnTabla = e
End Function

' ======================================================================
' SECCIÓN 4: FUNCIONES AUXILIARES DE BÚSQUEDA
' ======================================================================

' --- Verifica si un ID existe en la primera columna de una tabla ---
Private Function ExisteIDEnTabla(ByVal nombreLogico As String, _
                                  ByVal idBuscado As String) As Boolean
    ExisteIDEnTabla = (TableManager.BuscarFilaPorID(nombreLogico, idBuscado) > 0)
End Function

' --- Verifica si un valor existe en una columna específica ---
Private Function ExisteValorEnColumna(ByVal nombreLogico As String, _
                                       ByVal nombreColumna As String, _
                                       ByVal valorBuscado As String) As Boolean
    Dim tbl As ListObject
    Set tbl = TableManager.ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
        ExisteValorEnColumna = False
        Exit Function
    End If
    
    Dim colIdx As Long
    colIdx = TableManager.ObtenerIndiceColumna(tbl, nombreColumna)
    If colIdx = 0 Then
        ExisteValorEnColumna = False
        Exit Function
    End If
    
    Dim fila As Long
    For fila = 1 To tbl.ListRows.Count
        If LCase(Trim(CStr(tbl.ListRows(fila).Range.Cells(1, colIdx).Value))) = LCase(Trim(valorBuscado)) Then
            ExisteValorEnColumna = True
            Exit Function
        End If
    Next fila
    
    ExisteValorEnColumna = False
End Function

' --- Cuenta cuántas veces aparece un ID en una columna de otra tabla ---
Private Function ContarReferenciasEnTabla(ByVal nombreLogico As String, _
                                           ByVal nombreColumna As String, _
                                           ByVal idBuscado As String) As Long
    Dim tbl As ListObject
    Set tbl = TableManager.ObtenerListObject(nombreLogico)
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
        ContarReferenciasEnTabla = 0
        Exit Function
    End If
    
    Dim colIdx As Long
    colIdx = TableManager.ObtenerIndiceColumna(tbl, nombreColumna)
    If colIdx = 0 Then
        ContarReferenciasEnTabla = 0
        Exit Function
    End If
    
    Dim contador As Long: contador = 0
    Dim fila As Long
    For fila = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(fila).Range.Cells(1, colIdx).Value) = idBuscado Then
            contador = contador + 1
        End If
    Next fila
    
    ContarReferenciasEnTabla = contador
End Function

' --- Verifica si existen opciones de respuesta para una combinación sección+criticidad ---
Private Function ExistenOpcionesParaCombinacion(ByVal idSeccion As String, _
                                                  ByVal idCriticidad As String) As Boolean
    Dim tbl As ListObject
    Set tbl = TableManager.ObtenerListObject("OPCIONES")
    
    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
        ExistenOpcionesParaCombinacion = False
        Exit Function
    End If
    
    Dim colSec As Long: colSec = TableManager.ObtenerIndiceColumna(tbl, "ID Seccion")
    Dim colCrit As Long: colCrit = TableManager.ObtenerIndiceColumna(tbl, "ID Criticidad")
    
    ' Si la tabla no tiene columna de criticidad, solo validar sección
    If colCrit = 0 Then
        Dim fila2 As Long
        For fila2 = 1 To tbl.ListRows.Count
            If CStr(tbl.ListRows(fila2).Range.Cells(1, colSec).Value) = idSeccion Then
                ExistenOpcionesParaCombinacion = True
                Exit Function
            End If
        Next fila2
        ExistenOpcionesParaCombinacion = False
        Exit Function
    End If
    
    Dim fila As Long
    For fila = 1 To tbl.ListRows.Count
        If CStr(tbl.ListRows(fila).Range.Cells(1, colSec).Value) = idSeccion Then
            If CStr(tbl.ListRows(fila).Range.Cells(1, colCrit).Value) = idCriticidad Then
                ExistenOpcionesParaCombinacion = True
                Exit Function
            End If
        End If
    Next fila
    
    ExistenOpcionesParaCombinacion = False
End Function
