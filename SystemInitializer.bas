' ----------------------------------------------------------------------
' Módulo: SystemInitializer
' Propósito: Inicialización inteligente del sistema de inspecciones.
'            Detecta automáticamente si es la primera vez que se abre el libro
'            y configura todos los componentes necesarios para que funcione.
' Fecha creación: 14/04/2026
' Dependencias: Configuration2, InspectionScheduler, CronogramaResumen
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Subrutina: InicializarSistemaCompleto
' Propósito: Punto de entrada principal para inicialización automática.
'            Llamado desde ThisWorkbook.Workbook_Open.
' Lógica:
'   1. Valida datos maestros (personal, plantillas, equipos, evaluadores)
'   2. Configura filtro de planta en Menú Principal
'   3. Inicializa tblCronogramaInspecciones si está vacía
'   4. Actualiza tblResumenCronograma
'   5. Muestra mensaje de bienvenida si es primera vez
' ----------------------------------------------------------------------
Public Sub InicializarSistemaCompleto()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== INICIO: InicializarSistemaCompleto ==="
    
    ' --- PASO 1: Validar datos maestros ---
    Dim validacion As Object
    Set validacion = ValidarDatosMaestros()
    
    If Not validacion("Valido") Then
        MsgBox validacion("Mensaje"), vbExclamation, "Configuración incompleta"
        Debug.Print "ADVERTENCIA: " & validacion("Mensaje")
        Exit Sub
    End If
    
    Debug.Print "  ✓ Datos maestros válidos"
    
    ' --- PASO 2: Configurar filtro de planta ---
    Call ConfigurarFiltroPlantaPorDefecto(validacion("PlantaPorDefecto"))
    Debug.Print "  ✓ Filtro de planta configurado: " & validacion("PlantaPorDefecto")
    
    ' --- PASO 3: Inicializar cronograma si está vacío ---
    Dim cronogramaInicializado As Boolean
    cronogramaInicializado = InicializarCronogramaSiEstaVacio()
    
    ' --- PASO 4: Actualizar resumen cronograma ---
    Call CronogramaResumen.RefrescarResumenCronograma
    Debug.Print "  ✓ Resumen cronograma actualizado"
    
    ' --- PASO 5: Mensaje de bienvenida (solo si se inicializó por primera vez) ---
    If cronogramaInicializado Then
        MsgBox "¡Bienvenido al Sistema de Inspecciones!" & vbCrLf & vbCrLf & _
               "El sistema ha sido inicializado correctamente:" & vbCrLf & _
               "• Personal: " & validacion("CantPersonal") & " persona(s)" & vbCrLf & _
               "• Plantillas: " & validacion("CantPlantillas") & " plantilla(s)" & vbCrLf & _
               "• Cronogramas: " & validacion("CantCronogramas") & " registro(s)" & vbCrLf & vbCrLf & _
               "Haz doble clic en cualquier fila de la tabla para iniciar una inspección.", _
               vbInformation, "Sistema Inicializado"
    End If
    
    Debug.Print "=== FIN: InicializarSistemaCompleto ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en InicializarSistemaCompleto: " & Err.Number & " - " & Err.Description
    MsgBox "Error al inicializar el sistema: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Función: ValidarDatosMaestros
' Propósito: Verifica que existan datos mínimos necesarios para operar.
' Retorna: Dictionary con claves:
'   - Valido: Boolean (True si cumple requisitos mínimos)
'   - Mensaje: String (mensaje descriptivo)
'   - CantPersonal: Long
'   - CantPlantillas: Long
'   - CantEquipos: Long
'   - CantEvaluadores: Long
'   - PlantaPorDefecto: String
' ----------------------------------------------------------------------
Private Function ValidarDatosMaestros() As Object
    On Error GoTo ErrorHandler
    
    Dim resultado As Object
    Set resultado = CreateObject("Scripting.Dictionary")
    
    ' Valores por defecto
    resultado("Valido") = False
    resultado("Mensaje") = ""
    resultado("CantPersonal") = 0
    resultado("CantPlantillas") = 0
    resultado("CantEquipos") = 0
    resultado("CantEvaluadores") = 0
    resultado("PlantaPorDefecto") = "Todas"
    resultado("CantCronogramas") = 0
    
    ' --- Validar tblPersonal ---
    Dim tblPersonal As ListObject
    Set tblPersonal = ThisWorkbook.Sheets(Configuration2.SHEET_PERSONAL).ListObjects(Configuration2.TABLE_PERSONAL)
    
    If tblPersonal.DataBodyRange Is Nothing Then
        resultado("Mensaje") = "No hay personal registrado en la tabla tblPersonal." & vbCrLf & _
                               "Agrega al menos una persona antes de iniciar el sistema."
        Set ValidarDatosMaestros = resultado
        Exit Function
    End If
    
    resultado("CantPersonal") = tblPersonal.ListRows.Count
    
    ' Obtener primera planta disponible
    Dim primeraFila As ListRow
    Set primeraFila = tblPersonal.ListRows(1)
    resultado("PlantaPorDefecto") = Trim(primeraFila.Range.Cells(1, tblPersonal.ListColumns("Planta").Index).Value)
    
    ' --- Validar tblPlantillas ---
    Dim tblPlantillas As ListObject
    Set tblPlantillas = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST).ListObjects(Configuration2.TABLE_PLANTILLAS)
    
    If tblPlantillas.DataBodyRange Is Nothing Then
        resultado("Mensaje") = "No hay plantillas de inspección configuradas en tblPlantillas." & vbCrLf & _
                               "Crea al menos una plantilla antes de iniciar el sistema."
        Set ValidarDatosMaestros = resultado
        Exit Function
    End If
    
    resultado("CantPlantillas") = tblPlantillas.ListRows.Count
    
    ' --- Validar tblEquipos ---
    Dim tblEquipos As ListObject
    Set tblEquipos = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION).ListObjects(Configuration2.TABLE_EQUIPOS)
    
    If tblEquipos.DataBodyRange Is Nothing Then
        resultado("Mensaje") = "No hay equipos/áreas configurados en tblEquipos." & vbCrLf & _
                               "Agrega al menos un equipo/área antes de iniciar el sistema."
        Set ValidarDatosMaestros = resultado
        Exit Function
    End If
    
    resultado("CantEquipos") = tblEquipos.ListRows.Count
    
    ' --- Validar tblAseguramientoCalidad (evaluadores) ---
    Dim tblAseg As ListObject
    Set tblAseg = ThisWorkbook.Sheets(Configuration2.SHEET_ASEGURAMIENTO).ListObjects(Configuration2.TABLE_ASEGURAMIENTO)
    
    If tblAseg.DataBodyRange Is Nothing Then
        resultado("Mensaje") = "No hay evaluadores configurados en tblAseguramientoCalidad." & vbCrLf & _
                               "Agrega al menos un evaluador antes de iniciar el sistema."
        Set ValidarDatosMaestros = resultado
        Exit Function
    End If
    
    resultado("CantEvaluadores") = tblAseg.ListRows.Count
    
    ' --- Todo válido ---
    resultado("Valido") = True
    resultado("Mensaje") = "Sistema listo para operar"
    
    Set ValidarDatosMaestros = resultado
    Exit Function
    
ErrorHandler:
    resultado("Valido") = False
    resultado("Mensaje") = "Error al validar datos maestros: " & Err.Description
    Set ValidarDatosMaestros = resultado
End Function

'' ----------------------------------------------------------------------
' Subrutina: ConfigurarFiltroPlantaPorDefecto
' Propósito: Configura lista de validación de plantas en J15 del Menú Principal.
'            Lee plantas desde tblPlanta, agrega "Todas" al inicio,
'            crea validación de datos y asigna "Todas" por defecto.
' Parámetros:
'   plantaPorDefecto: No usado (se mantiene por compatibilidad)
' ----------------------------------------------------------------------
Private Sub ConfigurarFiltroPlantaPorDefecto(ByVal plantaPorDefecto As String)
    On Error GoTo ErrorHandler
    
    Dim wsMenu As Worksheet
    Dim wsConfig As Worksheet
    Dim tblPlanta As ListObject
    Dim celdaFiltro As Range
    Dim plantaRow As ListRow
    Dim listaValidacion As String
    Dim plantaNombre As String
    
    Set wsMenu = ThisWorkbook.Sheets(Configuration2.MAIN_MENU_SHEET)
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tblPlanta = wsConfig.ListObjects(Configuration2.TABLE_PLANTA)
    Set celdaFiltro = wsMenu.Range(Configuration2.RESUMEN_FILTRO_PLANTA_CELDA)
    
    ' --- Construir lista de validación: "Todas" + plantas desde tblPlanta ---
    listaValidacion = "Todas"
    
    If Not tblPlanta.DataBodyRange Is Nothing Then
        For Each plantaRow In tblPlanta.ListRows
            plantaNombre = Trim(plantaRow.Range.Cells(1, tblPlanta.ListColumns("Planta").Index).Value)
            If plantaNombre <> "" Then
                listaValidacion = listaValidacion & "," & plantaNombre
            End If
        Next plantaRow
    End If
    
    ' --- Configurar validación de datos en J15 ---
    Call SheetProtector2.UnprotectSheet(wsMenu, Configuration2.APP_PASSWORD)
    
    With celdaFiltro.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listaValidacion
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' --- Asignar "Todas" por defecto si está vacía ---
    If Trim(celdaFiltro.Value) = "" Then
        celdaFiltro.Value = "Todas"
        Debug.Print "  → Filtro planta asignado: Todas (por defecto)"
    Else
        Debug.Print "  → Filtro planta ya configurado: " & celdaFiltro.Value
    End If
    
    Call SheetProtector2.ProtectSheet(wsMenu, Configuration2.APP_PASSWORD)
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en ConfigurarFiltroPlantaPorDefecto: " & Err.Description
    Call SheetProtector2.ProtectSheet(wsMenu, Configuration2.APP_PASSWORD)
End Sub

'' ----------------------------------------------------------------------
' Función: InicializarCronogramaSiEstaVacio
' Propósito: Verifica si tblCronogramaInspecciones está vacía.
'            Si lo está, llama a InspectionScheduler.InicializarCronograma.
' Retorna: True si se inicializó, False si ya tenía datos.
' ----------------------------------------------------------------------
Private Function InicializarCronogramaSiEstaVacio() As Boolean
    On Error GoTo ErrorHandler
    
    Dim tblCronograma As ListObject
    Set tblCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA).ListObjects(Configuration2.TABLE_CRONOGRAMA)
    
    ' Si la tabla está vacía, inicializar
    If tblCronograma.DataBodyRange Is Nothing Then
        Debug.Print "  → Cronograma vacío. Inicializando..."
        
        ' Llamar al inicializador (versión silenciosa)
        Call InicializarCronogramaSilencioso(tblCronograma)
        
        InicializarCronogramaSiEstaVacio = True
        Debug.Print "  ✓ Cronograma inicializado con " & tblCronograma.ListRows.Count & " registros"
    Else
        InicializarCronogramaSiEstaVacio = False
        Debug.Print "  → Cronograma ya tiene " & tblCronograma.ListRows.Count & " registros"
    End If
    
    Exit Function
    
ErrorHandler:
    InicializarCronogramaSiEstaVacio = False
    Debug.Print "ERROR en InicializarCronogramaSiEstaVacio: " & Err.Description
End Function

'' ----------------------------------------------------------------------
' Subrutina: InicializarCronogramaSilencioso
' Propósito: Versión simplificada de InspectionScheduler.InicializarCronograma
'            sin mensajes al usuario (para ejecución automática).
' ----------------------------------------------------------------------
Private Sub InicializarCronogramaSilencioso(ByRef tblCronograma As ListObject)
    On Error GoTo ErrorHandler
    
    Dim tblPersonal As ListObject
    Dim tblPlantillas As ListObject
    Dim personaRow As ListRow
    Dim plantillaRow As ListRow
    Dim nuevoRow As ListRow
    Dim puestosArray As Variant
    
    Set tblPersonal = ThisWorkbook.Sheets(Configuration2.SHEET_PERSONAL).ListObjects(Configuration2.TABLE_PERSONAL)
    Set tblPlantillas = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST).ListObjects(Configuration2.TABLE_PLANTILLAS)
    
    puestosArray = Configuration2.GetPuestosColumns()
    
    ' Iterar sobre cada persona
    If Not tblPersonal.DataBodyRange Is Nothing Then
        For Each personaRow In tblPersonal.ListRows
            Dim iniciales As String
            Dim planta As String
            Dim activo As String
            
            iniciales = personaRow.Range.Cells(1, tblPersonal.ListColumns("Iniciales").Index).Value
            planta = personaRow.Range.Cells(1, tblPersonal.ListColumns("Planta").Index).Value
            activo = personaRow.Range.Cells(1, tblPersonal.ListColumns("Activo").Index).Value
            
            ' Para cada puesto
            Dim puestoColumna As Variant
            For Each puestoColumna In puestosArray
                Dim puestoNombre As String
                Dim puestoValor As String
                
                puestoNombre = CStr(puestoColumna)
                
                On Error Resume Next
                puestoValor = personaRow.Range.Cells(1, tblPersonal.ListColumns(puestoNombre).Index).Value
                On Error GoTo ErrorHandler
                
                ' Si el puesto está activo
                If UCase(Trim(puestoValor)) = "SI" Then
                    ' Buscar plantillas para este puesto
                    If Not tblPlantillas.DataBodyRange Is Nothing Then
                        For Each plantillaRow In tblPlantillas.ListRows
                            Dim plantillaPuesto As String
                            plantillaPuesto = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("Puesto").Index).Value
                            
                            If Trim(plantillaPuesto) = Trim(puestoNombre) Then
                                ' Crear registro en cronograma
                                Set nuevoRow = tblCronograma.ListRows.Add
                                
                                With nuevoRow.Range
                                    .Cells(1, tblCronograma.ListColumns("ID Cronograma").Index).Value = GenerarID()
                                    .Cells(1, tblCronograma.ListColumns("Iniciales personal").Index).Value = iniciales
                                    .Cells(1, tblCronograma.ListColumns("Puesto").Index).Value = puestoNombre
                                    .Cells(1, tblCronograma.ListColumns("ID Plantilla").Index).Value = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("ID Plantilla").Index).Value
                                    .Cells(1, tblCronograma.ListColumns("Nombre plantilla").Index).Value = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("Nombre de plantilla").Index).Value
                                    .Cells(1, tblCronograma.ListColumns("Frecuencia meses").Index).Value = plantillaRow.Range.Cells(1, tblPlantillas.ListColumns("Frecuencia meses").Index).Value
                                    .Cells(1, tblCronograma.ListColumns("Estado cronograma").Index).Value = Configuration2.ESTADO_NUNCA_INSPECCIONADO
                                    .Cells(1, tblCronograma.ListColumns("Planta personal").Index).Value = planta
                                End With
                            End If
                        Next plantillaRow
                    End If
                End If
            Next puestoColumna
        Next personaRow
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en InicializarCronogramaSilencioso: " & Err.Description
End Sub

'' ----------------------------------------------------------------------
' Función: GenerarID
' Propósito: Genera un ID único en formato XXXXXXXX-XXXXXXXX-XXXXXXXXXX
' ----------------------------------------------------------------------
Private Function GenerarID() As String
    Randomize
    GenerarID = GenerarSegmento(8) & "-" & GenerarSegmento(8) & "-" & GenerarSegmento(10)
End Function

Private Function GenerarSegmento(ByVal longitud As Long) As String
    Const CHARS As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    Dim i As Long
    Dim resultado As String
    
    resultado = ""
    For i = 1 To longitud
        Dim pos As Long
        pos = Int((Len(CHARS) * Rnd) + 1)
        resultado = resultado & Mid(CHARS, pos, 1)
    Next i
    
    GenerarSegmento = resultado
End Function
