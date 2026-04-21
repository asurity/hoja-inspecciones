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
'   - lblPuesto (Label): "Seleccione Puesto:"
'   - cboPuesto (ComboBox): Lista de puestos disponibles
'   - lblPersonal (Label): "Seleccione Personal:"
'   - cboPersonal (ComboBox): Lista de personal por puesto
'   - lblPlantilla (Label): "Plantilla asignada:"
'   - txtPlantilla (TextBox, Locked): Nombre de plantilla (solo lectura)
'   - btnAceptar (CommandButton): "Iniciar Inspección"
'   - btnCancelar (CommandButton): "Cancelar"
'
' LAYOUT SUGERIDO:
'   Top=10: lblPuesto
'   Top=30: cboPuesto (Width=250)
'   Top=70: lblPersonal
'   Top=90: cboPersonal (Width=250)
'   Top=130: lblPlantilla
'   Top=150: txtPlantilla (Width=350, Locked=True)
'   Top=200: btnAceptar, btnCancelar
' ======================================================================
Option Explicit

' --- Estado interno ---
Private mPuestoSeleccionado As String
Private mPersonalSeleccionado As String
Private mIDPlantilla As String
Private mNombrePlantilla As String
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
        .Height = 280
        .BackColor = RGB(250, 248, 245)  ' Beige claro
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
' Subrutina: ConfigurarCampoPuesto
' Propósito: Configura label y combobox de Puesto.
' ----------------------------------------------------------------------
Private Sub ConfigurarCampoPuesto()
    With Me.lblPuesto
        .Left = 20
        .Top = 35
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
' Subrutina: ConfigurarCampoPersonal
' Propósito: Configura label y combobox de Personal.
' ----------------------------------------------------------------------
Private Sub ConfigurarCampoPersonal()
    With Me.lblPersonal
        .Left = 20
        .Top = 90
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
        .Top = 110
        .Width = 380
        .Height = 22
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Style = fmStyleDropDownList
        .BackColor = RGB(240, 240, 240)  ' Gris claro (deshabilitado inicialmente)
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
        .Top = 145
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
        .Top = 165
        .Width = 380
        .Height = 22
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Style = fmStyleDropDownList
        .BackColor = RGB(240, 240, 240)  ' Gris claro (deshabilitado inicialmente)
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
        .Top = 200
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
        .Top = 200
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
'            y carga la lista de puestos en cboPuesto.
' ----------------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    ' Inicializar estado
    mCancelado = True
    mPuestoSeleccionado = ""
    mPersonalSeleccionado = ""
    mIDPlantilla = ""
    mNombrePlantilla = ""
    mPlanta = ""
    
    ' --- CONFIGURAR DISEÑO DEL FORMULARIO ---
    Call ConfigurarFormulario
    Call ConfigurarTitulo
    Call ConfigurarCampoPuesto
    Call ConfigurarCampoPersonal
    Call ConfigurarCampoPlantillas
    Call ConfigurarBotones
    
    ' Cargar lista de puestos
    Call CargarPuestos
    
    ' Deshabilitar controles hasta que se seleccione algo
    cboPersonal.Enabled = False
    cboPlantillas.Enabled = False
    btnAceptar.Enabled = False
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al inicializar formulario: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Evento: cboPuesto_Change
' Propósito: Se dispara cuando el usuario selecciona un puesto.
'            Carga la lista de personal disponible para ese puesto.
' ----------------------------------------------------------------------
Private Sub cboPuesto_Change()
    On Error GoTo ErrorHandler
    
    If cboPuesto.ListIndex = -1 Then Exit Sub
    
    mPuestoSeleccionado = cboPuesto.Value
    
    ' Limpiar selecciones anteriores
    cboPersonal.Clear
    cboPlantillas.Clear
    mPersonalSeleccionado = ""
    mIDPlantilla = ""
    mNombrePlantilla = ""
    
    ' Cargar personal para este puesto
    Call CargarPersonalPorPuesto(mPuestoSeleccionado)
    
    ' Habilitar ComboBox de personal (cambiar color a blanco)
    cboPersonal.Enabled = True
    cboPersonal.BackColor = vbWhite
    cboPlantillas.Enabled = False
    cboPlantillas.BackColor = RGB(240, 240, 240)
    btnAceptar.Enabled = False
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar personal: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Evento: cboPersonal_Change
' Propósito: Se dispara cuando el usuario selecciona un personal.
'            Carga la planta del personal y plantillas disponibles.
' ----------------------------------------------------------------------
Private Sub cboPersonal_Change()
    On Error GoTo ErrorHandler
    
    If cboPersonal.ListIndex = -1 Then Exit Sub
    
    mPersonalSeleccionado = cboPersonal.Value
    
    ' Obtener planta del personal
    Call ObtenerPlantaDelPersonal(mPersonalSeleccionado)
    
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
'            Actualiza el nombre de la plantilla y habilita Aceptar.
' ----------------------------------------------------------------------
Private Sub cboPlantillas_Change()
    On Error GoTo ErrorHandler
    
    If cboPlantillas.ListIndex = -1 Then Exit Sub
    
    ' El combo tiene formato: "NombrePlantilla (ID)"
    ' Extraer ID de los últimos caracteres (UUID)
    Dim plantillaInfo As String: plantillaInfo = cboPlantillas.Value
    Dim posApertura As Long: posApertura = InStrRev(plantillaInfo, "(")
    
    ' Extraer solo el contenido entre paréntesis sin los paréntesis
    mIDPlantilla = Trim(Mid(plantillaInfo, posApertura + 1, Len(plantillaInfo) - posApertura - 1))
    mNombrePlantilla = Trim(Left(plantillaInfo, posApertura - 1))
    
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
    Call ChecklistOrchestrator.AbrirChecklistVirtual(mPersonalSeleccionado, mIDPlantilla, mPuestoSeleccionado, "")
    
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
' Subrutina: CargarPuestos
' Propósito: Llena cboPuesto con la lista de puestos desde tblPuesto.
'            Carga solo los puestos únicos disponibles en el sistema.
' ----------------------------------------------------------------------
Private Sub CargarPuestos()
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim tblPuestos As ListObject
    Dim puestoRow As ListRow
    
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tblPuestos = wsConfig.ListObjects(Configuration2.TABLE_PUESTO)
    
    cboPuesto.Clear
    
    If Not tblPuestos.DataBodyRange Is Nothing Then
        For Each puestoRow In tblPuestos.ListRows
            Dim nombrePuesto As String
            ' Columna 2 = Puesto (columna 1 es Etapa, no se usa)
            nombrePuesto = Trim(puestoRow.Range.Cells(1, 2).Value)
            
            If nombrePuesto <> "" Then
                cboPuesto.AddItem nombrePuesto
            End If
        Next puestoRow
    End If
    
    If cboPuesto.ListCount = 0 Then
        MsgBox "No hay puestos configurados en tblPuesto.", vbExclamation, "Sin puestos"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar puestos: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: CargarPersonalPorPuesto
' Propósito: Llena cboPersonal con el personal activo para el puesto.
' Parámetros:
'   puesto: Nombre del puesto ("Químico", "Operador", etc.)
' ----------------------------------------------------------------------
Private Sub CargarPersonalPorPuesto(ByVal puesto As String)
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
        Dim puestoValor As String
        
        For Each personaRow In tblPersonal.ListRows
            ' Verificar si está activo
            On Error Resume Next
            activo = personaRow.Range.Cells(1, tblPersonal.ListColumns("Activo").Index).Value
            On Error GoTo ErrorHandler
            
            If UCase(Trim(activo)) <> "SI" Then GoTo NextPersona
            
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
        MsgBox "No hay personal activo para el puesto '" & puesto & "'.", vbInformation, "Sin personal"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar personal: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: ObtenerPlantaDelPersonal
' Propósito: Obtiene la planta del personal seleccionado.
' Parámetros:
'   iniciales: Iniciales del personal seleccionado
' ----------------------------------------------------------------------
Private Sub ObtenerPlantaDelPersonal(ByVal iniciales As String)
    On Error GoTo ErrorHandler
    
    Dim wsPersonal As Worksheet
    Dim tblPersonal As ListObject
    
    Set wsPersonal = ThisWorkbook.Sheets(Configuration2.SHEET_PERSONAL)
    Set tblPersonal = wsPersonal.ListObjects(Configuration2.TABLE_PERSONAL)
    
    mPlanta = ""
    
    If Not tblPersonal.DataBodyRange Is Nothing Then
        Dim personaRow As ListRow
        For Each personaRow In tblPersonal.ListRows
            Dim personInit As String
            personInit = Trim(personaRow.Range.Cells(1, tblPersonal.ListColumns("Iniciales").Index).Value)
            If personInit = iniciales Then
                mPlanta = Trim(personaRow.Range.Cells(1, tblPersonal.ListColumns("Planta").Index).Value)
                Exit For
            End If
        Next personaRow
    End If
    
    Exit Sub
ErrorHandler:
    mPlanta = ""
End Sub

'' ----------------------------------------------------------------------
' Subrutina: CargarPlantillasDisponibles
' Propósito: Carga en cboPlantillas todas las plantillas para el puesto.
' Parámetros:
'   puesto: Nombre del puesto
' Formato: Cada item es "Nombre Plantilla (ID Plantilla)"
' ----------------------------------------------------------------------
Private Sub CargarPlantillasDisponibles(ByVal puesto As String)
    On Error GoTo ErrorHandler
    
    Dim wsChecklist As Worksheet
    Dim tblPlantillas As ListObject
    
    Set wsChecklist = ThisWorkbook.Sheets(Configuration2.SHEET_CHECKLIST)
    Set tblPlantillas = wsChecklist.ListObjects(Configuration2.TABLE_PLANTILLAS)
    
    cboPlantillas.Clear
    
    If Not tblPlantillas.DataBodyRange Is Nothing Then
        Dim plantillaRow As ListRow
        Dim puestoPLT As String
        Dim nombrePlantilla As String
        Dim idPlantilla As String
        
        For Each plantillaRow In tblPlantillas.ListRows
            ' Columna 4 = Puesto (según Configuration2.bas)
            puestoPLT = Trim(plantillaRow.Range.Cells(1, 4).Value)
            
            If puestoPLT = puesto Then
                ' Columna 1 = ID Plantilla
                idPlantilla = Trim(plantillaRow.Range.Cells(1, 1).Value)
                ' Columna 2 = Nombre plantilla
                nombrePlantilla = Trim(plantillaRow.Range.Cells(1, 2).Value)
                
                If nombrePlantilla <> "" And idPlantilla <> "" Then
                    cboPlantillas.AddItem nombrePlantilla & " (" & idPlantilla & ")"
                End If
            End If
        Next plantillaRow
    End If
    
    If cboPlantillas.ListCount = 0 Then
        MsgBox "No hay plantillas disponibles para el puesto '" & puesto & "'.", _
               vbExclamation, "Sin plantillas"
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar plantillas: " & Err.Description, vbCritical, "Error"
End Sub
