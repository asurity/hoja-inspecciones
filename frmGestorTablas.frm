' ======================================================================
' UserForm: frmGestorTablas
' Descripción: Formulario único CRUD para gestionar las 5 tablas maestras
'              del sistema de inspecciones.
' Tablas: tblCriticidad, tblSecciones, tblPlantillas,
'         tblOpcionesDeRespuesta, tblPreguntas
' Dependencias: TableManager.bas, TableValidator.bas, Configuration2.bas,
'               AuditLogger2.bas
' Controles requeridos en el diseñador VBA:
'   - lblTitulo (Label), lblTabla (Label), cmbTabla (ComboBox)
'   - lstDatos (ListBox)
'   - fraEdicion (Frame)
'   - lblCampo1..lblCampo8 (Labels dentro de fraEdicion)
'   - txtCampo1..txtCampo3, txtCampo7, txtCampo8 (TextBox dentro de fraEdicion)
'   - cmbCampo4, cmbCampo5, cmbCampo6 (ComboBox dentro de fraEdicion)
'   - chkActivo (CheckBox dentro de fraEdicion)
'   - btnNuevo, btnGuardar, btnEliminar, btnValidar, btnCerrar (CommandButton)
'   - lblEstado (Label)
' ======================================================================
Option Explicit

' --- Constantes de diseño (posiciones y tamaños en puntos) ---
Private Const FORM_WIDTH As Single = 650
Private Const FORM_HEIGHT As Single = 520
Private Const MARGIN_LEFT As Single = 12
Private Const MARGIN_RIGHT As Single = 12
Private Const CONTENT_WIDTH As Single = 596  ' FORM_WIDTH - MARGIN_LEFT - MARGIN_RIGHT

' --- Constantes de colores ---
Private Const COLOR_FONDO As Long = &HFFFFFF        ' Fondo formulario (blanco)
Private Const COLOR_TITULO As Long = &H724E27        ' Título (marrón oscuro)
Private Const COLOR_FRAME As Long = &HFFFFFF         ' Fondo Frame (blanco)
Private Const COLOR_BOTON_NUEVO As Long = &HC0DCC0    ' Verde claro
Private Const COLOR_BOTON_GUARDAR As Long = &HFFFFC0  ' Amarillo claro
Private Const COLOR_BOTON_ELIMINAR As Long = &HC0C0FF ' Rojo claro
Private Const COLOR_BOTON_VALIDAR As Long = &HFFFFC0  ' Amarillo
Private Const COLOR_BOTON_CERRAR As Long = &HC0C0C0   ' Gris
Private Const COLOR_ESTADO_OK As Long = &H8000&       ' Verde (mensajes ok)
Private Const COLOR_ESTADO_ERROR As Long = &HFF&      ' Rojo (mensajes error)
Private Const COLOR_READONLY As Long = &HFFFFFF       ' Blanco (campos no editables)

' --- Estado interno ---
Private mModoEdicion As String   ' "NUEVO" o "EDITAR"
Private mTablaActual As String   ' Nombre lógico de la tabla activa
Private mFilaSeleccionada As Long ' Fila seleccionada en lstDatos (-1 = ninguna)

' ======================================================================
' INICIALIZACIÓN: Posiciona y configura TODOS los controles
' ======================================================================
Private Sub UserForm_Initialize()
    ' --- Configurar formulario ---
    Me.Caption = "Gestor de Tablas Maestras"
    Me.Width = FORM_WIDTH
    Me.Height = FORM_HEIGHT
    Me.BackColor = COLOR_FONDO
    
    mModoEdicion = ""
    mTablaActual = ""
    mFilaSeleccionada = -1
    
    ' --- Posicionar controles ---
    Call ConfigurarTitulo
    Call ConfigurarSelectorTabla
    Call ConfigurarListaDatos
    Call ConfigurarFrameEdicion
    Call ConfigurarCamposEdicion
    Call ConfigurarBotones
    Call ConfigurarBarraEstado
    
    ' --- Estado inicial: sin tabla seleccionada ---
    Call HabilitarBotones(False, False, False)
    Call MostrarEstado("Seleccione una tabla para comenzar.", False)
End Sub

' ======================================================================
' SECCIÓN 1: CONFIGURACIÓN DE CONTROLES (Posicionamiento físico)
' ======================================================================

Private Sub ConfigurarTitulo()
    With Me.lblTitulo
        .Left = MARGIN_LEFT
        .Top = 6
        .Width = CONTENT_WIDTH
        .Height = 20
        .Caption = "GESTOR DE TABLAS MAESTRAS"
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = COLOR_TITULO
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

Private Sub ConfigurarSelectorTabla()
    With Me.lblTabla
        .Left = MARGIN_LEFT
        .Top = 32
        .Width = 90
        .Height = 18
        .Caption = "Seleccionar tabla:"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    
    With Me.cmbTabla
        .Left = 108
        .Top = 30
        .Width = CONTENT_WIDTH - 96
        .Height = 22
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Style = fmStyleDropDownList
        .Clear
        .AddItem "1. Criticidades"
        .AddItem "2. Secciones"
        .AddItem "3. Plantillas"
        .AddItem "4. Opciones de Respuesta"
        .AddItem "5. Preguntas"
    End With
End Sub

Private Sub ConfigurarListaDatos()
    With Me.lstDatos
        .Left = MARGIN_LEFT
        .Top = 58
        .Width = CONTENT_WIDTH
        .Height = 140
        .Font.Name = "Consolas"
        .Font.Size = 9
        .ColumnCount = 1   ' Se ajustará dinámicamente
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStylePlain
        .BackColor = vbWhite
        .BorderStyle = fmBorderStyleSingle
    End With
End Sub

Private Sub ConfigurarFrameEdicion()
    With Me.fraEdicion
        .Left = MARGIN_LEFT
        .Top = 204
        .Width = CONTENT_WIDTH
        .Height = 246
        .Caption = " Datos del registro "
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .BackColor = COLOR_FRAME
    End With
End Sub

Private Sub ConfigurarCamposEdicion()
    Dim labelWidth As Single: labelWidth = 100
    Dim inputLeft As Single: inputLeft = 110
    Dim inputWidth As Single: inputWidth = 468
    Dim rowHeight As Single: rowHeight = 24
    Dim startY As Single: startY = 18
    Dim i As Long
    
    ' --- Labels (lblCampo1 a lblCampo8) ---
    Dim labels As Variant
    labels = Array(Me.lblCampo1, Me.lblCampo2, Me.lblCampo3, Me.lblCampo4, _
                   Me.lblCampo5, Me.lblCampo6, Me.lblCampo7, Me.lblCampo8)
    
    For i = 0 To 7
        With labels(i)
            .Left = 8
            .Top = startY + (i * rowHeight)
            .Width = labelWidth
            .Height = 18
            .Caption = "Campo " & (i + 1) & ":"
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .TextAlign = fmTextAlignRight
            .BackStyle = fmBackStyleTransparent
            .Visible = False
        End With
    Next i
    
    ' --- TextBoxes (txtCampo1, txtCampo2, txtCampo3, txtCampo7, txtCampo8) ---
    Dim textboxes As Variant
    textboxes = Array(Me.txtCampo1, Me.txtCampo2, Me.txtCampo3, Me.txtCampo7, Me.txtCampo8)
    Dim textRows As Variant
    textRows = Array(0, 1, 2, 6, 7)  ' Fila donde se posiciona cada textbox
    
    For i = 0 To 4
        With textboxes(i)
            .Left = inputLeft
            .Top = startY + (CLng(textRows(i)) * rowHeight)
            .Width = inputWidth
            .Height = 20
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Visible = False
        End With
    Next i
    
    ' txtCampo1 siempre es ID = readonly
    Me.txtCampo1.BackColor = COLOR_READONLY
    Me.txtCampo1.Locked = True
    
    ' --- ComboBoxes (cmbCampo4, cmbCampo5, cmbCampo6) ---
    Dim combos As Variant
    combos = Array(Me.cmbCampo4, Me.cmbCampo5, Me.cmbCampo6)
    Dim comboRows As Variant
    comboRows = Array(3, 4, 5)  ' Filas 4, 5, 6
    
    For i = 0 To 2
        With combos(i)
            .Left = inputLeft
            .Top = startY + (CLng(comboRows(i)) * rowHeight)
            .Width = inputWidth
            .Height = 20
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Style = fmStyleDropDownList
            .Visible = False
        End With
    Next i
    
    ' --- CheckBox Activo ---
    With Me.chkActivo
        .Left = inputLeft
        .Top = startY + (8 * rowHeight)
        .Width = 120
        .Height = 18
        .Caption = "Activo"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Value = True
        .Visible = False
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

Private Sub ConfigurarBotones()
    Dim btnTop As Single: btnTop = 456
    Dim btnWidth As Single: btnWidth = 90
    Dim btnHeight As Single: btnHeight = 28
    Dim btnSpacing As Single: btnSpacing = 8
    
    ' --- Grupo CRUD (izquierda) ---
    With Me.btnNuevo
        .Left = MARGIN_LEFT
        .Top = btnTop
        .Width = btnWidth
        .Height = btnHeight
        .Caption = "Nuevo"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COLOR_BOTON_NUEVO
    End With
    
    With Me.btnGuardar
        .Left = MARGIN_LEFT + btnWidth + btnSpacing
        .Top = btnTop
        .Width = btnWidth
        .Height = btnHeight
        .Caption = "Guardar"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COLOR_BOTON_GUARDAR
    End With
    
    With Me.btnEliminar
        .Left = MARGIN_LEFT + (btnWidth + btnSpacing) * 2
        .Top = btnTop
        .Width = btnWidth
        .Height = btnHeight
        .Caption = "Eliminar"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COLOR_BOTON_ELIMINAR
    End With
    
    ' --- Grupo Sistema (derecha) ---
    With Me.btnValidar
        .Left = FORM_WIDTH - MARGIN_RIGHT - (btnWidth * 2) - btnSpacing - 8
        .Top = btnTop
        .Width = btnWidth
        .Height = btnHeight
        .Caption = "Validar Todo"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .BackColor = COLOR_BOTON_VALIDAR
    End With
    
    With Me.btnCerrar
        .Left = FORM_WIDTH - MARGIN_RIGHT - btnWidth - 8
        .Top = btnTop
        .Width = btnWidth
        .Height = btnHeight
        .Caption = "Cerrar"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackColor = COLOR_BOTON_CERRAR
    End With
End Sub

Private Sub ConfigurarBarraEstado()
    With Me.lblEstado
        .Left = 0
        .Top = FORM_HEIGHT - 32
        .Width = FORM_WIDTH
        .Height = 20
        .Caption = " Listo"
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .ForeColor = COLOR_ESTADO_OK
        .BackColor = &HFFFFFF
        .BackStyle = fmBackStyleOpaque
        .TextAlign = fmTextAlignLeft
    End With
End Sub

' ======================================================================
' SECCIÓN 2: EVENTOS DE CONTROLES
' ======================================================================

' --- Cambio de tabla seleccionada ---
Private Sub cmbTabla_Change()
    If Me.cmbTabla.ListIndex < 0 Then Exit Sub
    
    Select Case Me.cmbTabla.ListIndex
        Case 0: mTablaActual = "CRITICIDAD"
        Case 1: mTablaActual = "SECCIONES"
        Case 2: mTablaActual = "PLANTILLAS"
        Case 3: mTablaActual = "OPCIONES"
        Case 4: mTablaActual = "PREGUNTAS"
    End Select
    
    Call ConfigurarCamposParaTabla(mTablaActual)
    Call CargarDatosEnLista
    Call LimpiarCamposEdicion
    Call HabilitarBotones(True, False, False)
    Call MostrarEstado("Tabla cargada: " & mTablaActual, False)
End Sub

' --- Click en fila del listbox ---
Private Sub lstDatos_Click()
    If Me.lstDatos.ListIndex < 0 Then Exit Sub
    
    mFilaSeleccionada = Me.lstDatos.ListIndex
    Call CargarFilaEnCampos(mFilaSeleccionada)
    mModoEdicion = "EDITAR"
    Call HabilitarBotones(True, True, True)
    Call MostrarEstado("Registro seleccionado. Puede editar o eliminar.", False)
End Sub

' --- Botón NUEVO ---
Private Sub btnNuevo_Click()
    If mTablaActual = "" Then
        Call MostrarEstado("ERROR: Seleccione una tabla primero.", True)
        Exit Sub
    End If
    
    mModoEdicion = "NUEVO"
    mFilaSeleccionada = -1
    Call LimpiarCamposEdicion
    
    ' Generar nuevo ID automáticamente
    Me.txtCampo1.Value = TableManager.GenerarNuevoID(mTablaActual)
    
    ' Habilitar guardar, deshabilitar eliminar
    Call HabilitarBotones(True, True, False)
    
    ' Focus al primer campo editable
    If Me.txtCampo2.Visible Then Me.txtCampo2.SetFocus
    
    Call MostrarEstado("Modo NUEVO: Complete los campos y presione Guardar.", False)
End Sub

' --- Botón GUARDAR ---
Private Sub btnGuardar_Click()
    If mTablaActual = "" Or mModoEdicion = "" Then
        Call MostrarEstado("ERROR: No hay operación en curso.", True)
        Exit Sub
    End If
    
    ' Paso 1: Recopilar datos del formulario
    Dim datos As Object
    Set datos = RecopilarDatosDelFormulario()
    If datos Is Nothing Then Exit Sub  ' Error ya mostrado
    
    ' Paso 2: Validar integridad referencial
    Dim erroresValidacion As String
    erroresValidacion = TableValidator.ValidarDatosParaGuardar(mTablaActual, datos, mModoEdicion)
    
    If erroresValidacion <> "" Then
        Call MostrarEstado("ERROR: " & erroresValidacion, True)
        MsgBox "No se puede guardar:" & vbCrLf & vbCrLf & erroresValidacion, _
               vbExclamation, "Validación fallida"
        Exit Sub
    End If
    
    ' Paso 3: Guardar en tabla
    Dim resultado As Boolean
    
    If mModoEdicion = "NUEVO" Then
        resultado = TableManager.InsertarFila(mTablaActual, datos)
    Else
        resultado = TableManager.ActualizarFila(mTablaActual, mFilaSeleccionada, datos)
    End If
    
    If resultado Then
        ' Paso 4: Registrar en audit trail
        Call TableManager.RegistrarCambioAudit(mTablaActual, mModoEdicion, datos)
        
        ' Paso 5: Recargar lista
        Call CargarDatosEnLista
        Call LimpiarCamposEdicion
        mModoEdicion = ""
        Call HabilitarBotones(True, False, False)
        Call MostrarEstado("Registro guardado exitosamente.", False)
    Else
        Call MostrarEstado("ERROR: No se pudo guardar el registro.", True)
    End If
End Sub

' --- Botón ELIMINAR ---
Private Sub btnEliminar_Click()
    If mTablaActual = "" Or mFilaSeleccionada < 0 Then
        Call MostrarEstado("ERROR: Seleccione un registro primero.", True)
        Exit Sub
    End If
    
    ' Obtener el ID del registro seleccionado
    Dim idRegistro As String
    idRegistro = Me.txtCampo1.Value
    
    ' Verificar dependencias
    Dim dependencias As String
    dependencias = TableValidator.VerificarDependencias(mTablaActual, idRegistro)
    
    If dependencias <> "" Then
        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox("Este registro tiene dependencias:" & vbCrLf & vbCrLf & _
                          dependencias & vbCrLf & vbCrLf & _
                          "Se marcará como INACTIVO (soft-delete)." & vbCrLf & _
                          "¿Desea continuar?", _
                          vbQuestion + vbYesNo, "Dependencias encontradas")
        
        If respuesta = vbNo Then Exit Sub
        
        ' Soft-delete: marcar como inactivo
        Call TableManager.MarcarInactivo(mTablaActual, mFilaSeleccionada)
        Call MostrarEstado("Registro marcado como INACTIVO.", False)
    Else
        ' Sin dependencias: confirmar eliminación física
        Dim confirmar As VbMsgBoxResult
        confirmar = MsgBox("¿Está seguro de eliminar este registro?" & vbCrLf & _
                          "Esta acción no se puede deshacer.", _
                          vbQuestion + vbYesNo, "Confirmar eliminación")
        
        If confirmar = vbNo Then Exit Sub
        
        Call TableManager.EliminarFila(mTablaActual, mFilaSeleccionada)
        Call MostrarEstado("Registro eliminado.", False)
    End If
    
    ' Registrar en audit trail
    Dim datosAudit As Object
    Set datosAudit = CreateObject("Scripting.Dictionary")
    datosAudit.Add "ID", idRegistro
    datosAudit.Add "Accion", IIf(dependencias <> "", "SOFT-DELETE", "DELETE")
    Call TableManager.RegistrarCambioAudit(mTablaActual, "ELIMINAR", datosAudit)
    
    Call CargarDatosEnLista
    Call LimpiarCamposEdicion
    mFilaSeleccionada = -1
    mModoEdicion = ""
    Call HabilitarBotones(True, False, False)
End Sub

' --- Botón VALIDAR TODO ---
Private Sub btnValidar_Click()
    Call MostrarEstado("Ejecutando validación completa...", False)
    
    Dim reporte As String
    reporte = TableValidator.ValidarIntegridad()
    
    If reporte = "" Then
        MsgBox "Todas las tablas están íntegras. No se encontraron errores.", _
               vbInformation, "Validación completa"
        Call MostrarEstado("Validación completa: Sin errores.", False)
    Else
        MsgBox "Se encontraron problemas de integridad:" & vbCrLf & vbCrLf & reporte, _
               vbExclamation, "Errores de integridad"
        Call MostrarEstado("Validación completa: Errores encontrados.", True)
    End If
End Sub

' --- Botón CERRAR ---
Private Sub btnCerrar_Click()
    If mModoEdicion <> "" Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("Tiene cambios sin guardar. ¿Desea cerrar de todas formas?", _
                      vbQuestion + vbYesNo, "Cambios pendientes")
        If resp = vbNo Then Exit Sub
    End If
    
    Unload Me
End Sub

' ======================================================================
' SECCIÓN 3: CONFIGURACIÓN DINÁMICA DE CAMPOS POR TABLA
' ======================================================================

Private Sub ConfigurarCamposParaTabla(ByVal tabla As String)
    ' Primero ocultar todos los campos
    Call OcultarTodosCampos
    
    Select Case tabla
        Case "CRITICIDAD"
            Call MostrarCampo(1, "ID Criticidad:", "TXT", True)
            Call MostrarCampo(2, "Nombre:", "TXT", False)
            Call MostrarCampo(3, "Valor:", "TXT", False)
            Me.chkActivo.Visible = False
            
            Me.lstDatos.ColumnCount = 3
            Me.lstDatos.ColumnWidths = "0;280;80"
            
        Case "SECCIONES"
            Call MostrarCampo(1, "ID Sección:", "TXT", True)
            Call MostrarCampo(2, "Nombre sección:", "TXT", False)
            Call MostrarCampo(3, "Tipo respuesta:", "TXT", False)
            Me.chkActivo.Visible = False
            
            Me.lstDatos.ColumnCount = 3
            Me.lstDatos.ColumnWidths = "0;300;100"
            
        Case "PLANTILLAS"
            Call MostrarCampo(1, "ID Plantilla:", "TXT", True)
            Call MostrarCampo(2, "Nombre plantilla:", "TXT", False)
            Call MostrarCampo(3, "Etapa:", "TXT", False)
            Call MostrarCampo(7, "Puesto:", "TXT", False)
            Call MostrarCampo(8, "Frecuencia meses:", "TXT", False)
            Me.chkActivo.Visible = False
            
            Me.lstDatos.ColumnCount = 5
            Me.lstDatos.ColumnWidths = "0;180;90;130;60"
            
        Case "OPCIONES"
            Call MostrarCampo(1, "ID Opción:", "TXT", True)
            Call MostrarCampo(4, "Sección:", "CMB", False)
            Call MostrarCampo(5, "Criticidad:", "CMB", False)
            Call MostrarCampo(2, "Texto opción:", "TXT", False)
            Call MostrarCampo(3, "Valor puntaje:", "TXT", False)
            Me.chkActivo.Visible = False
            
            Call CargarLookup(Me.cmbCampo4, "SECCIONES")
            Call CargarLookup(Me.cmbCampo5, "CRITICIDAD")
            
            Me.lstDatos.ColumnCount = 5
            Me.lstDatos.ColumnWidths = "0;120;100;120;60"
            
        Case "PREGUNTAS"
            Call MostrarCampo(1, "ID Pregunta:", "TXT", True)
            Call MostrarCampo(6, "Plantilla:", "CMB", False)
            Call MostrarCampo(4, "Sección:", "CMB", False)
            Call MostrarCampo(5, "Criticidad:", "CMB", False)
            Call MostrarCampo(2, "Texto pregunta:", "TXT", False)
            Call MostrarCampo(7, "Orden:", "TXT", False)
            Call MostrarCampo(8, "Observaciones:", "TXT", False)
            Me.chkActivo.Visible = True
            Me.chkActivo.Value = True
            
            Call CargarLookup(Me.cmbCampo4, "SECCIONES")
            Call CargarLookup(Me.cmbCampo5, "CRITICIDAD")
            Call CargarLookup(Me.cmbCampo6, "PLANTILLAS")
            
            Me.lstDatos.ColumnCount = 10
            Me.lstDatos.ColumnWidths = "0;120;100;100;220;80;50;50;150;100"
    End Select
End Sub

' ======================================================================
' SECCIÓN 4: FUNCIONES AUXILIARES DE UI
' ======================================================================

Private Sub MostrarCampo(ByVal numCampo As Long, ByVal etiqueta As String, _
                          ByVal tipoCampo As String, ByVal soloLectura As Boolean)
    ' Hacer visible label + control correspondiente
    Select Case numCampo
        Case 1
            Me.lblCampo1.Visible = True
            Me.lblCampo1.Caption = etiqueta
            Me.txtCampo1.Visible = True
            Me.txtCampo1.Locked = soloLectura
            If soloLectura Then Me.txtCampo1.BackColor = COLOR_READONLY Else Me.txtCampo1.BackColor = vbWhite
        Case 2
            Me.lblCampo2.Visible = True
            Me.lblCampo2.Caption = etiqueta
            Me.txtCampo2.Visible = True
            Me.txtCampo2.Locked = soloLectura
        Case 3
            Me.lblCampo3.Visible = True
            Me.lblCampo3.Caption = etiqueta
            Me.txtCampo3.Visible = True
            Me.txtCampo3.Locked = soloLectura
        Case 4
            Me.lblCampo4.Visible = True
            Me.lblCampo4.Caption = etiqueta
            Me.cmbCampo4.Visible = True
            Me.cmbCampo4.Enabled = Not soloLectura
        Case 5
            Me.lblCampo5.Visible = True
            Me.lblCampo5.Caption = etiqueta
            Me.cmbCampo5.Visible = True
            Me.cmbCampo5.Enabled = Not soloLectura
        Case 6
            Me.lblCampo6.Visible = True
            Me.lblCampo6.Caption = etiqueta
            Me.cmbCampo6.Visible = True
            Me.cmbCampo6.Enabled = Not soloLectura
        Case 7
            Me.lblCampo7.Visible = True
            Me.lblCampo7.Caption = etiqueta
            Me.txtCampo7.Visible = True
            Me.txtCampo7.Locked = soloLectura
        Case 8
            Me.lblCampo8.Visible = True
            Me.lblCampo8.Caption = etiqueta
            Me.txtCampo8.Visible = True
            Me.txtCampo8.Locked = soloLectura
    End Select
End Sub

Private Sub OcultarTodosCampos()
    Dim i As Long
    Dim ctrl As Control
    
    For Each ctrl In Me.fraEdicion.Controls
        ctrl.Visible = False
    Next ctrl
End Sub

Private Sub LimpiarCamposEdicion()
    Me.txtCampo1.Value = ""
    Me.txtCampo2.Value = ""
    Me.txtCampo3.Value = ""
    Me.cmbCampo4.ListIndex = -1
    Me.cmbCampo5.ListIndex = -1
    Me.cmbCampo6.ListIndex = -1
    Me.txtCampo7.Value = ""
    Me.txtCampo8.Value = ""
    Me.chkActivo.Value = True
End Sub

Private Sub HabilitarBotones(ByVal nuevo As Boolean, ByVal guardar As Boolean, _
                              ByVal eliminar As Boolean)
    Me.btnNuevo.Enabled = nuevo
    Me.btnGuardar.Enabled = guardar
    Me.btnEliminar.Enabled = eliminar
End Sub

Private Sub MostrarEstado(ByVal mensaje As String, ByVal esError As Boolean)
    Me.lblEstado.Caption = " " & mensaje
    If esError Then
        Me.lblEstado.ForeColor = COLOR_ESTADO_ERROR
    Else
        Me.lblEstado.ForeColor = COLOR_ESTADO_OK
    End If
End Sub

' ======================================================================
' SECCIÓN 5: CARGA DE DATOS (delega a TableManager)
' ======================================================================

Private Sub CargarDatosEnLista()
    ' Limpia el ListBox y carga todas las filas de la tabla activa
    ' Reemplaza IDs con nombres para mejor legibilidad
    Me.lstDatos.Clear
    
    If mTablaActual = "" Then Exit Sub
    
    Dim datos As Variant
    datos = TableManager.ObtenerDatosTabla(mTablaActual)
    
    If IsEmpty(datos) Then
        Call MostrarEstado("La tabla está vacía.", False)
        Exit Sub
    End If
    
    Dim fila As Long
    Dim col As Long
    Dim numCols As Long
    numCols = Me.lstDatos.ColumnCount
    
    For fila = LBound(datos, 1) To UBound(datos, 1)
        ' Para PREGUNTAS: transformar ID Plantilla a Nombre Plantilla antes de AddItem
        Dim primerValor As String
        If mTablaActual = "PREGUNTAS" Then
            primerValor = TableManager.ObtenerNombrePorID("PLANTILLAS", CStr(datos(fila, 1)))
        Else
            primerValor = CStr(datos(fila, 1))
        End If
        
        Me.lstDatos.AddItem primerValor   ' Primera columna
        
        ' Procesar cada columna y hacer lookups si es necesario
        For col = 2 To numCols
            If col <= UBound(datos, 2) Then
                Dim valorMostrar As String
                valorMostrar = CStr(datos(fila, col))
                
                ' Reemplazar IDs con nombres para OPCIONES y PREGUNTAS
                If mTablaActual = "OPCIONES" Then
                    If col = 2 Then  ' ID Seccion
                        valorMostrar = TableManager.ObtenerNombrePorID("SECCIONES", valorMostrar)
                    ElseIf col = 3 Then  ' ID Criticidad
                        valorMostrar = TableManager.ObtenerNombrePorID("CRITICIDAD", valorMostrar)
                    End If
                ElseIf mTablaActual = "PREGUNTAS" Then
                    If col = 3 Then  ' ID Seccion → Nombre Sección (col 3 de datos)
                        valorMostrar = TableManager.ObtenerNombrePorID("SECCIONES", valorMostrar)
                    ElseIf col = 6 Then  ' ID Criticidad → Nombre Criticidad (col 6 de datos)
                        valorMostrar = TableManager.ObtenerNombrePorID("CRITICIDAD", valorMostrar)
                    End If
                End If
                
                Me.lstDatos.List(Me.lstDatos.ListCount - 1, col - 1) = valorMostrar
            End If
        Next col
    Next fila
    
    Call MostrarEstado(mTablaActual & ": " & (UBound(datos, 1) - LBound(datos, 1) + 1) & " registros.", False)
End Sub

Private Function ObtenerHeadersParaTabla(ByVal tabla As String) As String
    ' Esta función se mantiene para referencia futura pero no se usa actualmente
    ObtenerHeadersParaTabla = ""
End Function

Private Sub CargarFilaEnCampos(ByVal indiceFila As Long)
    ' Delega a TableManager para obtener la fila y carga en campos
    Dim fila As Variant
    fila = TableManager.ObtenerFilaPorIndice(mTablaActual, indiceFila)
    
    If IsEmpty(fila) Then Exit Sub
    
    Select Case mTablaActual
        Case "CRITICIDAD"
            Me.txtCampo1.Value = fila(1)   ' ID
            Me.txtCampo2.Value = fila(2)   ' Nombre
            Me.txtCampo3.Value = fila(3)   ' Valor
            
        Case "SECCIONES"
            Me.txtCampo1.Value = fila(1)   ' ID
            Me.txtCampo2.Value = fila(2)   ' Nombre
            Me.txtCampo3.Value = fila(3)   ' Tipo respuesta
            
        Case "PLANTILLAS"
            Me.txtCampo1.Value = fila(1)   ' ID
            Me.txtCampo2.Value = fila(2)   ' Nombre
            Me.txtCampo3.Value = fila(3)   ' Etapa
            Me.txtCampo7.Value = fila(4)   ' Puesto
            Me.txtCampo8.Value = fila(5)   ' Frecuencia
            
        Case "OPCIONES"
            Me.txtCampo1.Value = fila(1)   ' ID
            Call SeleccionarLookup(Me.cmbCampo4, CStr(fila(2)))  ' ID Seccion
            Call SeleccionarLookup(Me.cmbCampo5, CStr(fila(3)))  ' ID Criticidad
            Me.txtCampo2.Value = fila(4)   ' Texto opción
            Me.txtCampo3.Value = fila(5)   ' Valor puntaje
            
        Case "PREGUNTAS"
            Me.txtCampo1.Value = fila(2)     ' ID Pregunta
            Call SeleccionarLookup(Me.cmbCampo6, CStr(fila(1)))  ' ID Plantilla
            Call SeleccionarLookup(Me.cmbCampo4, CStr(fila(3)))  ' ID Sección
            Call SeleccionarLookup(Me.cmbCampo5, CStr(fila(6)))  ' ID Criticidad
            Me.txtCampo2.Value = fila(5)     ' Texto pregunta
            Me.txtCampo7.Value = fila(7)     ' Orden
            Me.txtCampo8.Value = fila(9)     ' Observaciones
            Me.chkActivo.Value = (CStr(fila(8)) = "Sí" Or CStr(fila(8)) = "Si" Or fila(8) = True)
    End Select
End Sub

Private Sub CargarLookup(ByRef cmb As MSForms.ComboBox, ByVal tablaOrigen As String)
    ' Carga un ComboBox con IDs y nombres de una tabla maestra
    cmb.Clear
    
    Dim items As Variant
    items = TableManager.ObtenerItemsLookup(tablaOrigen)
    
    If IsEmpty(items) Then Exit Sub
    
    Dim i As Long
    cmb.ColumnCount = 2
    cmb.ColumnWidths = "0;300"  ' Oculta ID, muestra nombre. ID se lee con .Column(0)
    cmb.BoundColumn = 1
    
    For i = LBound(items, 1) To UBound(items, 1)
        cmb.AddItem items(i, 1)                      ' ID (columna 0 = oculta)
        cmb.List(cmb.ListCount - 1, 1) = items(i, 2)  ' Nombre visible
    Next i
End Sub

Private Sub SeleccionarLookup(ByRef cmb As MSForms.ComboBox, ByVal idBuscado As String)
    ' Selecciona un item en el ComboBox por su ID
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If CStr(cmb.List(i, 0)) = idBuscado Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next i
    cmb.ListIndex = -1  ' No encontrado
End Sub

' ======================================================================
' SECCIÓN 6: RECOPILACIÓN DE DATOS DEL FORMULARIO
' ======================================================================

Private Function RecopilarDatosDelFormulario() As Object
    ' Retorna un Dictionary con los datos ingresados, o Nothing si hay campos vacíos obligatorios
    Dim datos As Object
    Set datos = CreateObject("Scripting.Dictionary")
    
    On Error GoTo ErrorHandler
    
    Select Case mTablaActual
        Case "CRITICIDAD"
            If Trim(Me.txtCampo2.Value) = "" Then
                Call MostrarEstado("ERROR: El nombre es obligatorio.", True)
                Set RecopilarDatosDelFormulario = Nothing
                Exit Function
            End If
            datos.Add "ID", Me.txtCampo1.Value
            datos.Add "Nombre", Trim(Me.txtCampo2.Value)
            datos.Add "Valor", Me.txtCampo3.Value
            
        Case "SECCIONES"
            If Trim(Me.txtCampo2.Value) = "" Then
                Call MostrarEstado("ERROR: El nombre de sección es obligatorio.", True)
                Set RecopilarDatosDelFormulario = Nothing
                Exit Function
            End If
            datos.Add "ID", Me.txtCampo1.Value
            datos.Add "Nombre", Trim(Me.txtCampo2.Value)
            datos.Add "TipoRespuesta", Trim(Me.txtCampo3.Value)
            
        Case "PLANTILLAS"
            If Trim(Me.txtCampo2.Value) = "" Then
                Call MostrarEstado("ERROR: El nombre de plantilla es obligatorio.", True)
                Set RecopilarDatosDelFormulario = Nothing
                Exit Function
            End If
            datos.Add "ID", Me.txtCampo1.Value
            datos.Add "Nombre", Trim(Me.txtCampo2.Value)
            datos.Add "Etapa", Trim(Me.txtCampo3.Value)
            datos.Add "Puesto", Trim(Me.txtCampo7.Value)
            datos.Add "Frecuencia", Trim(Me.txtCampo8.Value)
            
        Case "OPCIONES"
            If Me.cmbCampo4.ListIndex < 0 Then
                Call MostrarEstado("ERROR: Debe seleccionar una Sección.", True)
                Set RecopilarDatosDelFormulario = Nothing
                Exit Function
            End If
            datos.Add "ID", Me.txtCampo1.Value
            datos.Add "IDSeccion", Me.cmbCampo4.List(Me.cmbCampo4.ListIndex, 0)
            datos.Add "IDCriticidad", IIf(Me.cmbCampo5.ListIndex >= 0, _
                                     Me.cmbCampo5.List(Me.cmbCampo5.ListIndex, 0), "")
            datos.Add "TextoOpcion", Trim(Me.txtCampo2.Value)
            datos.Add "ValorPuntaje", Trim(Me.txtCampo3.Value)
            
        Case "PREGUNTAS"
            If Me.cmbCampo6.ListIndex < 0 Then
                Call MostrarEstado("ERROR: Debe seleccionar una Plantilla.", True)
                Set RecopilarDatosDelFormulario = Nothing
                Exit Function
            End If
            If Me.cmbCampo4.ListIndex < 0 Then
                Call MostrarEstado("ERROR: Debe seleccionar una Sección.", True)
                Set RecopilarDatosDelFormulario = Nothing
                Exit Function
            End If
            If Me.cmbCampo5.ListIndex < 0 Then
                Call MostrarEstado("ERROR: Debe seleccionar una Criticidad.", True)
                Set RecopilarDatosDelFormulario = Nothing
                Exit Function
            End If
            datos.Add "IDPlantilla", Me.cmbCampo6.List(Me.cmbCampo6.ListIndex, 0)
            datos.Add "IDPregunta", Me.txtCampo1.Value
            datos.Add "IDSeccion", Me.cmbCampo4.List(Me.cmbCampo4.ListIndex, 0)
            datos.Add "IDCriticidad", Me.cmbCampo5.List(Me.cmbCampo5.ListIndex, 0)
            datos.Add "TextoPregunta", Trim(Me.txtCampo2.Value)
            datos.Add "Orden", Trim(Me.txtCampo7.Value)
            datos.Add "Activo", IIf(Me.chkActivo.Value, "Sí", "No")
            datos.Add "Observaciones", Trim(Me.txtCampo8.Value)
    End Select
    
    Set RecopilarDatosDelFormulario = datos
    Exit Function
    
ErrorHandler:
    Call MostrarEstado("ERROR: " & Err.Description, True)
    Set RecopilarDatosDelFormulario = Nothing
End Function
