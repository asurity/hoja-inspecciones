' ======================================================================
' UserForm: frmSelectorInspeccion
' Descripción: Formulario de selección para iniciar una inspección.
'              Permite elegir Puesto → Personal → Plantilla.
'              Al confirmar, abre frmChecklistVirtual con los datos.
' Fecha creación: 14/04/2026
' Última actualización: 14/04/2026 - Integración con frmChecklistVirtual
' Dependencias: CronogramaButtons, Configuration2, frmChecklistVirtual
'
' CONTROLES REQUERIDOS (crear manualmente en diseñador VBA):
'   - lblPlanta (Label): "Planta:"
'   - cboPlanta (ComboBox): Lista de plantas disponibles
'   - lblPuesto (Label): "Seleccione Puesto:"
'   - cboPuesto (ComboBox): Lista de puestos disponibles
'   - lblPersonal (Label): "Seleccione Personal:"
'   - cboPersonal (ComboBox): Lista de personal por puesto y planta
'   - lblPlantillas (Label): "Seleccione Plantilla:"
'   - cboPlantillas (ComboBox): Lista de plantillas disponibles
'   - btnAceptar (CommandButton): "Iniciar Inspección"
'   - btnCancelar (CommandButton): "Cancelar"
'
' LAYOUT SUGERIDO:
'   Top=35: lblPlanta
'   Top=55: cboPlanta (Width=380)
'   Top=90: lblPuesto
'   Top=110: cboPuesto (Width=380)
'   Top=145: lblPersonal
'   Top=165: cboPersonal (Width=380)
'   Top=200: lblPlantillas
'   Top=220: cboPlantillas (Width=380)
'   Top=265: btnAceptar, btnCancelar
' ======================================================================
Option Explicit

' --- Estado interno ---
Private mPlantaSeleccionada As String
Private mPuestoSeleccionado As String
Private mPersonalSeleccionado As String
Private mIDPlantilla As String
Private mNombrePlantilla As String
Private mAreaSeleccionada As String
Private mPlanta As String
Private mCancelado As Boolean

' --- Referencia global al formulario de checklist (para mantenerlo vivo) ---
Private oChecklistFormInstance As frmChecklistVirtual

' ======================================================================
' PROPIEDADES PÚBLICAS
' ======================================================================
Public Property Get Cancelado() As Boolean
    Cancelado = mCancelado
End Property

Public Property Get PuestoSeleccionado() As String
    PuestoSeleccionado = mPuestoSeleccionado
End Property

Public Property Get PersonalSeleccionado() As String
    PersonalSeleccionado = mPersonalSeleccionado
End Property

Public Property Get IDPlantilla() As String
    IDPlantilla = mIDPlantilla
End Property

Public Property Get Planta() As String
    Planta = mPlanta
End Property

' ======================================================================
' CONFIGURACIÓN DE DISEÑO (Posicionamiento y colores)
' ======================================================================

'' ----------------------------------------------------------------------
' Subrutina: ConfigurarFormulario
' Propósito: Configura las propiedades generales del formulario.
' ----------------------------------------------------------------------
Private Sub ConfigurarFormulario()
    With Me
        .Caption = "Nueva Inspección"
        .Width = 420
        .Height = 340  ' Aumentado para incluir campo Planta
        .BackColor = vbWhite  ' Blanco
        .StartUpPosition = 1  ' CenterOwner
    End With
End Sub

'' ----------------------------------------------------------------------
' Subrutina: ConfigurarTitulo
' Propósito: Configura el label de título del formulario.
' ----------------------------------------------------------------------
Private Sub ConfigurarTitulo()
    ' Crear label de título si no existe (se asume que existe en el diseñador)
    ' Si no existe, se puede crear dinámicamente
    
    On Error Resume Next
    Dim lblTitulo As MSForms.Label
    Set lblTitulo = Me.Controls("lblTitulo")
    
    If lblTitulo Is Nothing Then
        ' Si no existe, crear uno dinámico
        Set lblTitulo = Me.Controls.Add("Forms.Label.1", "lblTitulo")
    End If
    On Error GoTo 0
    
    With lblTitulo
        .Left = 12
        .Top = 8
        .Width = 396
        .Height = 22
        .Caption = "SELECCIONAR INSPECCIÓN"
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = RGB(114, 78, 39)  ' Marrón oscuro
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

'' ----------------------------------------------------------------------
' Subrutina: ConfigurarCampoPlanta
' Propósito: Configura label y combobox de Planta.
' ----------------------------------------------------------------------
Private Sub ConfigurarCampoPlanta()
    With Me.lblPlanta
        .Left = 20
        .Top = 35
        .Width = 100
        .Height = 18
        .Caption = "Planta:"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(60, 60, 60)
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    
    With Me.cboPlanta
        .Left = 20
        .Top = 55
        .Width = 380
        .Height = 22
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Style = fmStyleDropDownList
        .BackColor = vbWhite
        .ForeColor = RGB(0, 0, 0)
    End With
End Sub

'' ----------------------------------------------------------------------
' Subrutina: ConfigurarCampoPuesto
' Propósito: Configura label y combobox de Puesto.
' ----------------------------------------------------------------------
Private Sub ConfigurarCampoPuesto()
    With Me.lblPuesto
        .Left = 20
        .Top = 90
        .Width = 100
        .Height = 18
        .Caption = "Puesto:"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(60, 60, 60)
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    
    With Me.cboPuesto
        .Left = 20
        .Top = 110
        .Width = 380
        .Height = 22
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Style = fmStyleDropDownList
        .BackColor = vbWhite
        .ForeColor = RGB(0, 0, 0)
    End With
End Sub

'' ----------------------------------------------------------------------
' Subrutina: ConfigurarCampoPersonal
' Propósito: Configura label y combobox de Personal.
' ----------------------------------------------------------------------
Private Sub ConfigurarCampoPersonal()
    With Me.lblPersonal
        .Left = 20
        .Top = 145
        .Width = 100
        .Height = 18
        .Caption = "Personal:"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(60, 60, 60)
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    
    With Me.cboPersonal
        .Left = 20
        .Top = 165
        .Width = 380
        .Height = 22
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Style = fmStyleDropDownList
        .BackColor = vbWhite  ' Blanco
        .ForeColor = RGB(0, 0, 0)
    End With
End Sub

'' ----------------------------------------------------------------------
' Subrutina: ConfigurarCampoPlantillas
' Propósito: Configura label y combobox para seleccionar Plantilla.
' ----------------------------------------------------------------------
Private Sub ConfigurarCampoPlantillas()
    With Me.lblPlantillas
        .Left = 20
        .Top = 200
        .Width = 150
        .Height = 18
        .Caption = "Seleccione Plantilla:"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(60, 60, 60)
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    
    With Me.cboPlantillas
        .Left = 20
        .Top = 220
        .Width = 380
        .Height = 22
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Style = fmStyleDropDownList
        .BackColor = vbWhite  ' Blanco
        .ForeColor = RGB(0, 0, 0)
    End With
End Sub

'' ----------------------------------------------------------------------
' Subrutina: ConfigurarBotones
' Propósito: Configura los botones Aceptar y Cancelar.
' ----------------------------------------------------------------------
Private Sub ConfigurarBotones()
    ' Botón Aceptar
    With Me.btnAceptar
        .Left = 150
        .Top = 265
        .Width = 110
        .Height = 32
        .Caption = "Iniciar Inspección"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = RGB(192, 220, 192)  ' Verde claro
        .ForeColor = RGB(0, 0, 0)
    End With
    
    ' Botón Cancelar
    With Me.btnCancelar
        .Left = 270
        .Top = 265
        .Width = 110
        .Height = 32
        .Caption = "Cancelar"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackColor = RGB(192, 192, 192)  ' Gris
        .ForeColor = RGB(0, 0, 0)
    End With
End Sub

' ======================================================================
' EVENTOS
' ======================================================================

'' ----------------------------------------------------------------------
' Evento: UserForm_Initialize
' Propósito: Inicializa el formulario al abrirse.
'            Configura TODOS los controles (posición, tamaño, colores)
'            y carga la lista de plantas en cboPlanta.
' ----------------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    ' Inicializar estado
    mCancelado = True
    mPlantaSeleccionada = ""
    mPuestoSeleccionado = ""
    mPersonalSeleccionado = ""
    mIDPlantilla = ""
    mNombrePlantilla = ""
    mAreaSeleccionada = ""
    mPlanta = ""
    
    ' --- CONFIGURAR DISEÑO DEL FORMULARIO ---
    Call ConfigurarFormulario
    Call ConfigurarTitulo
    Call ConfigurarCampoPlanta
    Call ConfigurarCampoPuesto
    Call ConfigurarCampoPersonal
    Call ConfigurarCampoPlantillas
    Call ConfigurarBotones
    
    ' Cargar lista de plantas
    Call CargarPlantas
    
    ' Deshabilitar controles hasta que se seleccione algo
    cboPuesto.Enabled = False
    cboPersonal.Enabled = False
    cboPlantillas.Enabled = False
    btnAceptar.Enabled = False
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al inicializar formulario: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Evento: cboPlanta_Change
' Propósito: Se dispara cuando el usuario selecciona una planta.
'            Carga la lista de puestos disponibles para esa planta.
' ----------------------------------------------------------------------
Private Sub cboPlanta_Change()
    On Error GoTo ErrorHandler
    
    If cboPlanta.ListIndex = -1 Then Exit Sub
    
    mPlantaSeleccionada = cboPlanta.Value
    mPlanta = mPlantaSeleccionada
    
    ' Limpiar selecciones anteriores
    cboPuesto.Clear
    cboPersonal.Clear
    cboPlantillas.Clear
    mPuestoSeleccionado = ""
    mPersonalSeleccionado = ""
    mIDPlantilla = ""
    mNombrePlantilla = ""
    mAreaSeleccionada = ""
    
    ' Cargar puestos disponibles en esta planta
    Call CargarPuestosPorPlanta(mPlantaSeleccionada)
    
    ' Habilitar ComboBox de puesto
    cboPuesto.Enabled = True
    cboPuesto.BackColor = vbWhite
    cboPersonal.Enabled = False
    cboPersonal.BackColor = vbWhite
    cboPlantillas.Enabled = False
    cboPlantillas.BackColor = vbWhite
    btnAceptar.Enabled = False
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar puestos: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Evento: cboPuesto_Change
' Propósito: Se dispara cuando el usuario selecciona un puesto.
'            Carga la lista de personal disponible para ese puesto y planta.
' ----------------------------------------------------------------------
Private Sub cboPuesto_Change()
    On Error GoTo ErrorHandler
    
    If cboPuesto.ListIndex = -1 Then Exit Sub
    
    mPuestoSeleccionado = cboPuesto.Value
    Debug.Print "=== cboPuesto_Change ==="
    Debug.Print "Puesto seleccionado: [" & mPuestoSeleccionado & "]"
    
    ' Limpiar selecciones anteriores
    cboPersonal.Clear
    cboPlantillas.Clear
    mPersonalSeleccionado = ""
    mIDPlantilla = ""
    mNombrePlantilla = ""
    
    ' Cargar personal para este puesto y planta
    Call CargarPersonalPorPuestoYPlanta(mPuestoSeleccionado, mPlantaSeleccionada)
    
    ' Habilitar ComboBox de personal (cambiar color a blanco)
    cboPersonal.Enabled = True
    cboPersonal.BackColor = vbWhite
    cboPlantillas.Enabled = False
    cboPlantillas.BackColor = vbWhite
    btnAceptar.Enabled = False
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar personal: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Evento: cboPersonal_Change
' Propósito: Se dispara cuando el usuario selecciona un personal.
'            Carga plantillas disponibles para el puesto seleccionado.
' ----------------------------------------------------------------------
Private Sub cboPersonal_Change()
    On Error GoTo ErrorHandler
    
    If cboPersonal.ListIndex = -1 Then Exit Sub
    
    mPersonalSeleccionado = cboPersonal.Value
    
    Debug.Print "=== cboPersonal_Change ==="
    Debug.Print "Personal seleccionado: [" & mPersonalSeleccionado & "]"
    Debug.Print "Cargando plantillas para puesto: [" & mPuestoSeleccionado & "]"
    
    ' La planta ya está seleccionada en mPlanta desde cboPlanta_Change
    
    ' Cargar plantillas disponibles para este puesto
    Call CargarPlantillasDisponibles(mPuestoSeleccionado)
    
    ' Habilitar combo de plantillas
    cboPlantillas.Enabled = True
    cboPlantillas.BackColor = vbWhite
    mIDPlantilla = ""
    mNombrePlantilla = ""
    btnAceptar.Enabled = False
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar plantillas: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Evento: cboPlantillas_Change
' Propósito: Se dispara cuando el usuario selecciona una plantilla.
'            Actualiza el ID, nombre y área de la plantilla. Habilita Aceptar.
' Nota: El combo contiene objetos Dictionary con ID, Nombre y Área
' ----------------------------------------------------------------------
Private Sub cboPlantillas_Change()
    On Error GoTo ErrorHandler
    
    If cboPlantillas.ListIndex = -1 Then Exit Sub
    
    ' El valor es solo el nombre de la plantilla
    ' El ID y Área se obtienen del tag guardado en formato "ID|AREA"
    Dim tagInfo As String
    tagInfo = cboPlantillas.List(cboPlantillas.ListIndex, 1)  ' Columna oculta con ID|AREA
    
    Dim partes() As String
    partes = Split(tagInfo, "|")
    
    mIDPlantilla = partes(0)
    mAreaSeleccionada = partes(1)
    mNombrePlantilla = cboPlantillas.Value
    
    ' Habilitar botón Aceptar si tenemos planta
    If Len(mPlanta) > 0 Then
        btnAceptar.Enabled = True
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al seleccionar plantilla: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Evento: btnAceptar_Click
' Propósito: Confirma la selección y abre el ChecklistVirtual
' ----------------------------------------------------------------------
Private Sub btnAceptar_Click()
    On Error GoTo ErrorHandler
    
    ' Validar que todos los campos estén completos
    If Len(mPlantaSeleccionada) = 0 Then
        MsgBox "Por favor seleccione una Planta.", vbExclamation, "Validación"
        Exit Sub
    End If
    
    If Len(mPuestoSeleccionado) = 0 Then
        MsgBox "Por favor seleccione un Puesto.", vbExclamation, "Validación"
        Exit Sub
    End If
    
    If Len(mPersonalSeleccionado) = 0 Then
        MsgBox "Por favor seleccione Personal.", vbExclamation, "Validación"
        Exit Sub
    End If
    
    If Len(mIDPlantilla) = 0 Then
        MsgBox "Por favor seleccione una Plantilla.", vbExclamation, "Validación"
        Exit Sub
    End If
    
    If Len(mPlanta) = 0 Then
        MsgBox "No se pudo obtener la Planta del Personal.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Marcar como NO cancelado
    mCancelado = False
    
    ' Ocultar ANTES de abrir el siguiente formulario
    Me.Hide
    
    ' Abrir checklist virtual directamente
    ' idCronograma = "" porque es una inspección ad-hoc (no viene de tabla cronograma)
    Call ChecklistOrchestrator.AbrirChecklistVirtual(mPersonalSeleccionado, mIDPlantilla, mPuestoSeleccionado, "", mAreaSeleccionada)
    
    ' Cerrar este formulario después de que ChecklistVirtual se cierre
    Unload Me
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al abrir Checklist Virtual: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Evento: btnCancelar_Click
' Propósito: Cancela la operación y cierra el formulario.
' ----------------------------------------------------------------------
Private Sub btnCancelar_Click()
    mCancelado = True
    Me.Hide
End Sub

' ======================================================================
' MÉTODOS PRIVADOS (CARGA DE DATOS)
' ======================================================================

'' ----------------------------------------------------------------------
' Subrutina: CargarPlantas
' Propósito: Llena cboPlanta con la lista de plantas desde tblPlanta.
' ----------------------------------------------------------------------
Private Sub CargarPlantas()
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim tblPlanta As ListObject
    Dim plantaRow As ListRow
    
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tblPlanta = wsConfig.ListObjects(Configuration2.TABLE_PLANTA)
    
    cboPlanta.Clear
    
    If Not tblPlanta.DataBodyRange Is Nothing Then
        For Each plantaRow In tblPlanta.ListRows
            Dim nombrePlanta As String
            ' Columna 1 = Nombre de la planta
            nombrePlanta = Trim(plantaRow.Range.Cells(1, 1).Value)
            
            If nombrePlanta <> "" Then
                cboPlanta.AddItem nombrePlanta
            End If
        Next plantaRow
    End If
    
    If cboPlanta.ListCount = 0 Then
        MsgBox "No hay plantas configuradas en tblPlanta.", vbExclamation, "Sin plantas"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar plantas: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: CargarPuestosPorPlanta
' Propósito: Llena cboPuesto con los puestos que tienen personal activo
'            en la planta seleccionada.
' Parámetros:
'   planta: Nombre de la planta seleccionada
' ----------------------------------------------------------------------
Private Sub CargarPuestosPorPlanta(ByVal planta As String)
    On Error GoTo ErrorHandler
    
    Dim wsPersonal As Worksheet
    Dim tblPersonal As ListObject
    Dim personaRow As ListRow
    Dim puestosDict As Object
    Dim activo As String
    Dim plantaPersonal As String
    Dim puesto As String
    Dim puestoKey As Variant
    
    Set wsPersonal = ThisWorkbook.Sheets(Configuration2.SHEET_PERSONAL)
    Set tblPersonal = wsPersonal.ListObjects(Configuration2.TABLE_PERSONAL)
    Set puestosDict = CreateObject("Scripting.Dictionary")
    
    cboPuesto.Clear
    
    ' Recopilar puestos únicos del personal activo en esta planta
    If Not tblPersonal.DataBodyRange Is Nothing Then
        For Each personaRow In tblPersonal.ListRows
            ' Verificar si está activo
            On Error Resume Next
            activo = personaRow.Range.Cells(1, tblPersonal.ListColumns("Activo").Index).Value
            plantaPersonal = personaRow.Range.Cells(1, tblPersonal.ListColumns("Planta").Index).Value
            On Error GoTo ErrorHandler
            
            ' Solo personal activo de esta planta
            If UCase(Trim(activo)) = "SI" And Trim(plantaPersonal) = planta Then
                ' Buscar qué puesto(s) tiene activos (columnas con "SI")
                ' Las columnas de puestos son: Químico, Operador, Ayudante 2, etc.
                Dim col As ListColumn
                For Each col In tblPersonal.ListColumns
                    If col.Name <> "Iniciales" And col.Name <> "Nombre" And _
                       col.Name <> "Planta" And col.Name <> "Activo" And _
                       col.Name <> "ID Plantilla" And col.Name <> "Frecuencia (meses)" Then
                        
                        On Error Resume Next
                        puesto = personaRow.Range.Cells(1, col.Index).Value
                        On Error GoTo ErrorHandler
                        
                        If UCase(Trim(puesto)) = "SI" Then
                            ' Agregar este puesto al diccionario (clave = nombre puesto)
                            If Not puestosDict.Exists(col.Name) Then
                                puestosDict.Add col.Name, True
                            End If
                        End If
                    End If
                Next col
            End If
        Next personaRow
    End If
    
    ' Llenar combo con puestos encontrados
    For Each puestoKey In puestosDict.Keys
        cboPuesto.AddItem CStr(puestoKey)
    Next puestoKey
    
    If cboPuesto.ListCount = 0 Then
        MsgBox "No hay personal activo en la planta '" & planta & "'.", vbInformation, "Sin personal"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar puestos: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: CargarPersonalPorPuestoYPlanta
' Propósito: Llena cboPersonal con el personal activo para el puesto
'            y la planta seleccionados.
' Parámetros:
'   puesto: Nombre del puesto ("Químico", "Operador", etc.)
'   planta: Nombre de la planta
' ----------------------------------------------------------------------
Private Sub CargarPersonalPorPuestoYPlanta(ByVal puesto As String, ByVal planta As String)
    On Error GoTo ErrorHandler
    
    Dim wsPersonal As Worksheet
    Dim tblPersonal As ListObject
    
    Set wsPersonal = ThisWorkbook.Sheets(Configuration2.SHEET_PERSONAL)
    Set tblPersonal = wsPersonal.ListObjects(Configuration2.TABLE_PERSONAL)
    
    cboPersonal.Clear
    
    If Not tblPersonal.DataBodyRange Is Nothing Then
        Dim personaRow As ListRow
        Dim iniciales As String
        Dim activo As String
        Dim plantaPersonal As String
        Dim puestoValor As String
        
        For Each personaRow In tblPersonal.ListRows
            ' Verificar si está activo y en la planta correcta
            On Error Resume Next
            activo = personaRow.Range.Cells(1, tblPersonal.ListColumns("Activo").Index).Value
            plantaPersonal = personaRow.Range.Cells(1, tblPersonal.ListColumns("Planta").Index).Value
            On Error GoTo ErrorHandler
            
            If UCase(Trim(activo)) <> "SI" Then GoTo NextPersona
            If Trim(plantaPersonal) <> planta Then GoTo NextPersona
            
            ' Verificar si tiene el puesto
            On Error Resume Next
            puestoValor = personaRow.Range.Cells(1, tblPersonal.ListColumns(puesto).Index).Value
            On Error GoTo ErrorHandler
            
            If UCase(Trim(puestoValor)) = "SI" Then
                iniciales = Trim(personaRow.Range.Cells(1, tblPersonal.ListColumns("Iniciales").Index).Value)
                cboPersonal.AddItem iniciales
            End If
            
NextPersona:
        Next personaRow
    End If
    
    If cboPersonal.ListCount = 0 Then
        MsgBox "No hay personal activo para el puesto '" & puesto & "' en la planta '" & planta & "'.", _
               vbInformation, "Sin personal"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar personal: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: CargarPlantillasDisponibles
' Propósito: Carga en cboPlantillas todas las plantillas para el puesto.
' Parámetros:
'   puesto: Nombre del puesto
' Formato: Muestra solo el nombre de la plantilla (sin ID)
'          Guarda ID y Área en columna oculta para recuperar al seleccionar
' ----------------------------------------------------------------------
Private Sub CargarPlantillasDisponibles(ByVal puesto As String)
    On Error GoTo ErrorHandler
    
    Dim wsChecklist As Worksheet
    Dim tblPlantillas As ListObject
    
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set tblPlantillas = wsChecklist.ListObjects(Configuration2.TABLE_PLANTILLAS)
    
    Debug.Print "=== CargarPlantillasDisponibles ==="
    Debug.Print "Puesto buscado: [" & puesto & "]"
    
    cboPlantillas.Clear
    cboPlantillas.ColumnCount = 2  ' Columna 0: Nombre (visible), Columna 1: ID|AREA (oculta)
    cboPlantillas.ColumnWidths = "380;0"  ' Segunda columna oculta
    
    If Not tblPlantillas.DataBodyRange Is Nothing Then
        Dim plantillaRow As ListRow
        Dim puestoPLT As String
        Dim nombrePlantilla As String
        Dim idPlantilla As String
        Dim areaPlantilla As String
        
        ' Colecciones para almacenar plantillas exactas y flexibles
        Dim plantillasExactas As New Collection
        Dim plantillasFlexibles As New Collection
        
        For Each plantillaRow In tblPlantillas.ListRows
            ' Columna 4 = Puesto (según Configuration2.bas)
            puestoPLT = Trim(plantillaRow.Range.Cells(1, 4).Value)
            
            ' Columna 1 = ID Plantilla
            idPlantilla = Trim(plantillaRow.Range.Cells(1, 1).Value)
            ' Columna 2 = Nombre plantilla
            nombrePlantilla = Trim(plantillaRow.Range.Cells(1, 2).Value)
            ' Columna 3 = Área
            areaPlantilla = Trim(plantillaRow.Range.Cells(1, 3).Value)
            
            If nombrePlantilla <> "" And idPlantilla <> "" Then
                If puestoPLT = puesto Then
                    ' MATCH EXACTO
                    Dim arrExacta(0 To 2) As String
                    arrExacta(0) = nombrePlantilla
                    arrExacta(1) = idPlantilla
                    arrExacta(2) = areaPlantilla
                    plantillasExactas.Add arrExacta
                    Debug.Print "  ✓ Match EXACTO: [" & puestoPLT & "] = [" & puesto & "] → " & nombrePlantilla
                    
                ElseIf (Len(puestoPLT) > 0 And InStr(1, puesto, puestoPLT, vbTextCompare) = 1) Or _
                       (Len(puesto) > 0 And InStr(1, puestoPLT, puesto, vbTextCompare) = 1) Then
                    ' MATCH FLEXIBLE (solo si el puesto base coincide)
                    Dim arrFlexible(0 To 2) As String
                    arrFlexible(0) = nombrePlantilla
                    arrFlexible(1) = idPlantilla
                    arrFlexible(2) = areaPlantilla
                    plantillasFlexibles.Add arrFlexible
                    Debug.Print "  ~ Match flexible: [" & puestoPLT & "] ≈ [" & puesto & "] → " & nombrePlantilla
                End If
            End If
        Next plantillaRow
        
        ' PRIORIDAD: Mostrar EXACTAS si existen, sino FLEXIBLES
        Dim plantillasAMostrar As Collection
        If plantillasExactas.Count > 0 Then
            Set plantillasAMostrar = plantillasExactas
            Debug.Print ">>> Mostrando " & plantillasExactas.Count & " plantilla(s) con match EXACTO"
        Else
            Set plantillasAMostrar = plantillasFlexibles
            Debug.Print ">>> Mostrando " & plantillasFlexibles.Count & " plantilla(s) con match FLEXIBLE (no hay exactas)"
        End If
        
        ' Agregar al combo
        Dim plantilla As Variant
        For Each plantilla In plantillasAMostrar
            cboPlantillas.AddItem plantilla(0)  ' Nombre
            cboPlantillas.List(cboPlantillas.ListCount - 1, 1) = plantilla(1) & "|" & plantilla(2)  ' ID|AREA
        Next plantilla
    End If
    
    Debug.Print "Total plantillas en combo: " & cboPlantillas.ListCount
    
    If cboPlantillas.ListCount = 0 Then
        MsgBox "No hay plantillas disponibles para el puesto '" & puesto & "'.", _
               vbExclamation, "Sin plantillas"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar plantillas: " & Err.Description, vbCritical, "Error"
End Sub
