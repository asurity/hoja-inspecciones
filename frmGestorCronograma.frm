' ======================================================================
' UserForm: frmGestorCronograma
' Descripción: Formulario modal para gestionar el cronograma de inspecciones.
'              Permite filtrar, visualizar, pausar y reactivar inspecciones
'              en tblCronogramaInspecciones desde el Menú principal.
' Fecha creación: 28/04/2026
' Dependencias: CronogramaGestorService, CronogramaResumen, Configuration2
'
' CONTROLES REQUERIDOS (crear en diseñador VBA antes de importar):
'   - lblTitulo       (Label)     : título del formulario
'   - lblFiltroPlanta (Label)     : etiqueta "Planta:"
'   - cboFiltroPlanta (ComboBox)  : filtro por planta
'   - lblFiltroEstado (Label)     : etiqueta "Estado:"
'   - cboFiltroEstado (ComboBox)  : filtro por estado (Todas/Activas/Pausadas)
'   - lblColHeaders   (Label)     : cabecera visual de columnas del ListBox
'   - lstCronograma   (ListBox)   : lista principal (8 columnas, multiselección)
'   - lblActivas      (Label)     : etiqueta "Activas:"
'   - txtTotalActivas (TextBox)   : cantidad de inspecciones activas (ReadOnly)
'   - lblPausadas     (Label)     : etiqueta "Pausadas:"
'   - txtTotalPausadas(TextBox)   : cantidad de inspecciones pausadas (ReadOnly)
'   - lblSeleccionadas(Label)     : etiqueta "Seleccionadas:"
'   - txtSeleccionadas(TextBox)   : cantidad de filas seleccionadas (ReadOnly)
'   - btnPausar       (CommandButton): "Pausar Seleccionadas"
'   - btnReactivar    (CommandButton): "Reactivar Seleccionadas"
'   - btnRefrescar    (CommandButton): "Refrescar"
'   - btnCerrar       (CommandButton): "Cerrar"
'
' LAYOUT SUGERIDO (900x600):
'   lblTitulo         : Left=12, Top=6,   Width=864, Height=24
'   lblFiltroPlanta   : Left=12, Top=38,  Width=50,  Height=18
'   cboFiltroPlanta   : Left=66, Top=36,  Width=200, Height=22
'   lblFiltroEstado   : Left=278,Top=38,  Width=50,  Height=18
'   cboFiltroEstado   : Left=332,Top=36,  Width=160, Height=22
'   lblColHeaders     : Left=12, Top=64,  Width=864, Height=16
'   lstCronograma     : Left=12, Top=82,  Width=864, Height=360
'   lblActivas        : Left=12, Top=454, Width=55,  Height=18
'   txtTotalActivas   : Left=70, Top=452, Width=50,  Height=20
'   lblPausadas       : Left=130,Top=454, Width=60,  Height=18
'   txtTotalPausadas  : Left=194,Top=452, Width=50,  Height=20
'   lblSeleccionadas  : Left=254,Top=454, Width=80,  Height=18
'   txtSeleccionadas  : Left=338,Top=452, Width=50,  Height=20
'   btnPausar         : Left=530,Top=450, Width=140, Height=26
'   btnReactivar      : Left=680,Top=450, Width=140, Height=26
'   btnRefrescar      : Left=730,Top=484, Width=80,  Height=26
'   btnCerrar         : Left=820,Top=484, Width=80,  Height=26
' ======================================================================
Option Explicit

' --- Constantes de layout ---
Private Const FORM_WIDTH  As Single = 900
Private Const FORM_HEIGHT As Single = 560
Private Const MARGIN      As Single = 12
Private Const CONTENT_W   As Single = 876  ' FORM_WIDTH - 2*MARGIN

' --- Constantes de color ---
Private Const COLOR_FONDO          As Long = &HFFFFFF
Private Const COLOR_TITULO         As Long = &H724E27
Private Const COLOR_BOTON_PAUSAR   As Long = &HC0C0FF   ' Azul claro (acción de pausar)
Private Const COLOR_BOTON_REACTIVAR As Long = &HC0DCC0  ' Verde claro (acción de activar)
Private Const COLOR_BOTON_NEUTRO   As Long = &HC0C0C0   ' Gris
Private Const COLOR_READONLY       As Long = &HF0F0F0   ' Fondo lectura

' --- Opciones del filtro de estado ---
Private Const FILTRO_TODAS   As String = "Todas"
Private Const FILTRO_ACTIVAS As String = "Activas"
Private Const FILTRO_PAUSADAS As String = "Pausadas"

' --- Columnas del ListBox lstCronograma ---
' Col 0 (índice VBA): IDCronograma  — oculta (ancho 0)
' Col 1: Iniciales
' Col 2: Nombre Plantilla
' Col 3: Planta
' Col 4: Puesto
' Col 5: Frecuencia (meses)
' Col 6: Estado Cronograma
' Col 7: Activo en Cronograma
Private Const LST_COL_ID           As Integer = 0  ' Oculta
Private Const LST_COL_INICIALES    As Integer = 1
Private Const LST_COL_PLANTILLA    As Integer = 2
Private Const LST_COL_PLANTA       As Integer = 3
Private Const LST_COL_PUESTO       As Integer = 4
Private Const LST_COL_FRECUENCIA   As Integer = 5
Private Const LST_COL_ESTADO       As Integer = 6
Private Const LST_COL_ACTIVO       As Integer = 7

' --- Cache de datos cargados (para filtrado local sin re-consultar el servicio) ---
Private mDatosCompletos As Collection

' ======================================================================
' INICIALIZACIÓN
' ======================================================================

Private Sub UserForm_Initialize()
    Me.Caption = "Gestor de Cronograma de Inspecciones"
    Me.Width = FORM_WIDTH
    Me.Height = FORM_HEIGHT
    Me.BackColor = COLOR_FONDO

    Set mDatosCompletos = Nothing

    Call ConfigurarTitulo
    Call ConfigurarFiltros
    Call ConfigurarListBox
    Call ConfigurarEstadisticas
    Call ConfigurarBotones
    Call CargarDatos
End Sub

' ======================================================================
' SECCIÓN 1: CONFIGURACIÓN DE CONTROLES
' ======================================================================

Private Sub ConfigurarTitulo()
    With Me.lblTitulo
        .Left = MARGIN
        .Top = 6
        .Width = CONTENT_W
        .Height = 24
        .Caption = "GESTOR DE CRONOGRAMA DE INSPECCIONES"
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = COLOR_TITULO
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

Private Sub ConfigurarFiltros()
    With Me.lblFiltroPlanta
        .Left = MARGIN
        .Top = 38
        .Width = 50
        .Height = 18
        .Caption = "Planta:"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackStyle = fmBackStyleTransparent
    End With

    With Me.cboFiltroPlanta
        .Left = 66
        .Top = 36
        .Width = 200
        .Height = 22
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Style = fmStyleDropDownList
        .Clear
        .AddItem "Todas"
    End With

    With Me.lblFiltroEstado
        .Left = 278
        .Top = 38
        .Width = 50
        .Height = 18
        .Caption = "Estado:"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackStyle = fmBackStyleTransparent
    End With

    With Me.cboFiltroEstado
        .Left = 332
        .Top = 36
        .Width = 160
        .Height = 22
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Style = fmStyleDropDownList
        .Clear
        .AddItem FILTRO_TODAS
        .AddItem FILTRO_ACTIVAS
        .AddItem FILTRO_PAUSADAS
        .ListIndex = 0
    End With
End Sub

Private Sub ConfigurarListBox()
    ' --- Cabecera visual de columnas ---
    With Me.lblColHeaders
        .Left = MARGIN
        .Top = 64
        .Width = CONTENT_W
        .Height = 16
        .Caption = "  Iniciales  |  Nombre Plantilla                |  Planta        |  Puesto                   |  Frec.  |  Estado              |  Activo"
        .Font.Name = "Consolas"
        .Font.Size = 8
        .Font.Bold = True
        .BackColor = &HDCDCDC
        .BackStyle = fmBackStyleOpaque
    End With

    With Me.lstCronograma
        .Left = MARGIN
        .Top = 82
        .Width = CONTENT_W
        .Height = 360
        .Font.Name = "Consolas"
        .Font.Size = 9
        .BackColor = vbWhite
        .BorderStyle = fmBorderStyleSingle
        ' 8 columnas: col 0 = IDCronograma oculta, cols 1-7 visibles
        .ColumnCount = 8
        .ColumnWidths = "0;55;195;100;140;50;130;80"
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption
    End With
End Sub

Private Sub ConfigurarEstadisticas()
    With Me.lblActivas
        .Left = MARGIN
        .Top = 452
        .Width = 55
        .Height = 18
        .Caption = "Activas:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.txtTotalActivas
        .Left = 70
        .Top = 450
        .Width = 50
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = True
        .BackColor = COLOR_READONLY
        .TextAlign = fmTextAlignCenter
    End With

    With Me.lblPausadas
        .Left = 130
        .Top = 452
        .Width = 60
        .Height = 18
        .Caption = "Pausadas:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.txtTotalPausadas
        .Left = 194
        .Top = 450
        .Width = 50
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = True
        .BackColor = COLOR_READONLY
        .TextAlign = fmTextAlignCenter
    End With

    With Me.lblSeleccionadas
        .Left = 254
        .Top = 452
        .Width = 80
        .Height = 18
        .Caption = "Seleccionadas:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.txtSeleccionadas
        .Left = 338
        .Top = 450
        .Width = 50
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = True
        .BackColor = COLOR_READONLY
        .TextAlign = fmTextAlignCenter
    End With
End Sub

Private Sub ConfigurarBotones()
    With Me.btnPausar
        .Left = 510
        .Top = 448
        .Width = 160
        .Height = 26
        .Caption = "Pausar Seleccionadas"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .BackColor = COLOR_BOTON_PAUSAR
        .Enabled = False
    End With

    With Me.btnReactivar
        .Left = 680
        .Top = 448
        .Width = 170
        .Height = 26
        .Caption = "Reactivar Seleccionadas"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .BackColor = COLOR_BOTON_REACTIVAR
        .Enabled = False
    End With

    With Me.btnRefrescar
        .Left = 730
        .Top = 484
        .Width = 80
        .Height = 26
        .Caption = "Refrescar"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .BackColor = COLOR_BOTON_NEUTRO
    End With

    With Me.btnCerrar
        .Left = 820
        .Top = 484
        .Width = 68
        .Height = 26
        .Caption = "Cerrar"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .BackColor = COLOR_BOTON_NEUTRO
    End With
End Sub

' ======================================================================
' SECCIÓN 2: CARGA Y FILTRADO DE DATOS
' ======================================================================

'' ----------------------------------------------------------------------
' Subrutina: CargarDatos
' Propósito: Consulta CronogramaGestorService.ObtenerDatosParaGestor(),
'            carga plantas en el combo (primera vez), aplica filtros
'            y actualiza lstCronograma y estadísticas.
'            Separation of Concerns: UI solo renderiza, Service provee datos.
' ----------------------------------------------------------------------
Private Sub CargarDatos()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    ' Obtener todos los datos del servicio (cache local para filtrado)
    Set mDatosCompletos = CronogramaGestorService.ObtenerDatosParaGestor()

    ' Cargar plantas únicas en combo (solo en primera carga)
    If cboFiltroPlanta.ListCount <= 1 Then
        Call CargarComboPlanta
    End If

    ' Aplicar filtros y renderizar lista
    Call AplicarFiltrosYRenderizar

    ' Actualizar estadísticas totales (sin filtro de estado)
    Me.txtTotalActivas.Value = CronogramaGestorService.ContarInspeccionesPorEstado("Si")
    Me.txtTotalPausadas.Value = CronogramaGestorService.ContarInspeccionesPorEstado("No")
    Me.txtSeleccionadas.Value = "0"

    ' Deshabilitar botones de acción (ninguna fila seleccionada aún)
    Me.btnPausar.Enabled = False
    Me.btnReactivar.Enabled = False

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Call ErrorLogger2.Log("frmGestorCronograma.CargarDatos", Err.Description, Err.Number)
    MsgBox "Error al cargar datos: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: CargarComboPlanta
' Propósito: Llena cboFiltroPlanta con "Todas" + plantas únicas del servicio.
'            Se ejecuta una sola vez (primera carga) para no perder selección.
' ----------------------------------------------------------------------
Private Sub CargarComboPlanta()
    On Error GoTo ErrorHandler

    Dim plantaActual As String
    plantaActual = cboFiltroPlanta.Value
    If plantaActual = "" Then plantaActual = "Todas"

    cboFiltroPlanta.Clear
    cboFiltroPlanta.AddItem "Todas"

    Dim plantas As Collection
    Set plantas = CronogramaGestorService.ObtenerPlantasUnicas()

    Dim p As Variant
    For Each p In plantas
        cboFiltroPlanta.AddItem CStr(p)
    Next p

    ' Restaurar selección previa si sigue existiendo
    Dim i As Long
    For i = 0 To cboFiltroPlanta.ListCount - 1
        If cboFiltroPlanta.List(i) = plantaActual Then
            cboFiltroPlanta.ListIndex = i
            Exit For
        End If
    Next i
    If cboFiltroPlanta.ListIndex = -1 Then cboFiltroPlanta.ListIndex = 0

    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("frmGestorCronograma.CargarComboPlanta", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Subrutina: AplicarFiltrosYRenderizar
' Propósito: Filtra mDatosCompletos según combos activos y actualiza lstCronograma.
'            Principio DRY: único punto de renderizado del ListBox.
' ----------------------------------------------------------------------
Private Sub AplicarFiltrosYRenderizar()
    On Error GoTo ErrorHandler

    If mDatosCompletos Is Nothing Then Exit Sub

    Dim filtroPlanta As String
    Dim filtroEstado As String
    filtroPlanta = Trim(CStr(cboFiltroPlanta.Value))
    filtroEstado = Trim(CStr(cboFiltroEstado.Value))
    If filtroPlanta = "" Then filtroPlanta = "Todas"
    If filtroEstado = "" Then filtroEstado = FILTRO_TODAS

    Me.lstCronograma.Clear

    Dim item As Variant
    For Each item In mDatosCompletos
        ' Filtro planta
        If filtroPlanta <> "Todas" Then
            If CStr(item("Planta")) <> filtroPlanta Then GoTo SiguienteItem
        End If

        ' Filtro estado
        Dim activoVal As String
        activoVal = CStr(item("Activo"))
        If filtroEstado = FILTRO_ACTIVAS Then
            If UCase(activoVal) <> "SI" Then GoTo SiguienteItem
        ElseIf filtroEstado = FILTRO_PAUSADAS Then
            If UCase(activoVal) <> "NO" Then GoTo SiguienteItem
        End If

        ' Agregar fila al ListBox
        Me.lstCronograma.AddItem CStr(item("IDCronograma"))  ' Col 0 (oculta)
        Dim lastIdx As Long
        lastIdx = Me.lstCronograma.ListCount - 1
        Me.lstCronograma.List(lastIdx, LST_COL_INICIALES)  = CStr(item("Iniciales"))
        Me.lstCronograma.List(lastIdx, LST_COL_PLANTILLA)  = CStr(item("NombrePlantilla"))
        Me.lstCronograma.List(lastIdx, LST_COL_PLANTA)     = CStr(item("Planta"))
        Me.lstCronograma.List(lastIdx, LST_COL_PUESTO)     = CStr(item("Puesto"))
        Me.lstCronograma.List(lastIdx, LST_COL_FRECUENCIA) = CStr(item("Frecuencia"))
        Me.lstCronograma.List(lastIdx, LST_COL_ESTADO)     = CStr(item("EstadoCronograma"))
        Me.lstCronograma.List(lastIdx, LST_COL_ACTIVO)     = activoVal

SiguienteItem:
    Next item

    Me.txtSeleccionadas.Value = "0"
    Me.btnPausar.Enabled = False
    Me.btnReactivar.Enabled = False
    Exit Sub

ErrorHandler:
    Call ErrorLogger2.Log("frmGestorCronograma.AplicarFiltrosYRenderizar", Err.Description, Err.Number)
End Sub

' ======================================================================
' SECCIÓN 3: EVENTOS DE CONTROLES
' ======================================================================

Private Sub cboFiltroPlanta_Change()
    AplicarFiltrosYRenderizar
End Sub

Private Sub cboFiltroEstado_Change()
    AplicarFiltrosYRenderizar
End Sub

Private Sub lstCronograma_Click()
    Dim seleccionadas As Long
    seleccionadas = ContarSeleccionadas()
    Me.txtSeleccionadas.Value = seleccionadas
    Me.btnPausar.Enabled = (seleccionadas > 0)
    Me.btnReactivar.Enabled = (seleccionadas > 0)
End Sub

' ======================================================================
' SECCIÓN 4: ACCIONES DE BOTONES
' ======================================================================

Private Sub btnPausar_Click()
    On Error GoTo ErrorHandler

    Dim ids() As String
    ids = ObtenerIDsSeleccionados()

    If UBound(ids) < LBound(ids) Then
        MsgBox "Seleccione al menos una inspeccion.", vbInformation, "Sin seleccion"
        Exit Sub
    End If

    Dim cant As Long
    cant = UBound(ids) - LBound(ids) + 1

    If MsgBox("¿Pausar " & cant & " inspeccion(es) seleccionada(s)?" & vbCrLf & vbCrLf & _
              "Las inspecciones pausadas NO apareceran en el resumen del menu principal.", _
              vbQuestion + vbYesNo, "Confirmar pausa") = vbNo Then
        Exit Sub
    End If

    Dim actualizados As Long
    actualizados = CronogramaGestorService.PausarInspecciones(ids)

    MsgBox actualizados & " inspeccion(es) pausada(s) correctamente.", vbInformation, "Pausar"

    Call CargarDatos
    Exit Sub

ErrorHandler:
    Call ErrorLogger2.Log("frmGestorCronograma.btnPausar_Click", Err.Description, Err.Number)
    MsgBox "Error al pausar: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnReactivar_Click()
    On Error GoTo ErrorHandler

    Dim ids() As String
    ids = ObtenerIDsSeleccionados()

    If UBound(ids) < LBound(ids) Then
        MsgBox "Seleccione al menos una inspeccion.", vbInformation, "Sin seleccion"
        Exit Sub
    End If

    Dim cant As Long
    cant = UBound(ids) - LBound(ids) + 1

    If MsgBox("¿Reactivar " & cant & " inspeccion(es) seleccionada(s)?" & vbCrLf & vbCrLf & _
              "Las inspecciones reactivadas volvean a aparecer en el resumen del menu principal.", _
              vbQuestion + vbYesNo, "Confirmar reactivacion") = vbNo Then
        Exit Sub
    End If

    Dim actualizados As Long
    actualizados = CronogramaGestorService.ReactivarInspecciones(ids)

    MsgBox actualizados & " inspeccion(es) reactivada(s) correctamente.", vbInformation, "Reactivar"

    Call CargarDatos
    Exit Sub

ErrorHandler:
    Call ErrorLogger2.Log("frmGestorCronograma.btnReactivar_Click", Err.Description, Err.Number)
    MsgBox "Error al reactivar: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnRefrescar_Click()
    ' Fuerza recarga completa desde el servicio (invalida cache)
    Set mDatosCompletos = Nothing
    Call CargarDatos
End Sub

Private Sub btnCerrar_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Liberar cache de datos al cerrar
    Set mDatosCompletos = Nothing
End Sub

' ======================================================================
' SECCIÓN 5: HELPERS PRIVADOS
' ======================================================================

'' ----------------------------------------------------------------------
' Función: ContarSeleccionadas
' Propósito: Cuenta cuántas filas del ListBox están seleccionadas.
' Retorna: Long
' ----------------------------------------------------------------------
Private Function ContarSeleccionadas() As Long
    Dim total As Long
    total = 0
    Dim i As Long
    For i = 0 To Me.lstCronograma.ListCount - 1
        If Me.lstCronograma.Selected(i) Then total = total + 1
    Next i
    ContarSeleccionadas = total
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerIDsSeleccionados
' Propósito: Extrae los IDCronograma de las filas seleccionadas en lstCronograma.
'            El ID está en la columna 0 (oculta) del ListBox.
' Retorna: Array de Strings. Array vacío si no hay selección.
' ----------------------------------------------------------------------
Private Function ObtenerIDsSeleccionados() As String()
    Dim ids() As String
    Dim count As Long
    count = ContarSeleccionadas()

    If count = 0 Then
        ObtenerIDsSeleccionados = ids
        Exit Function
    End If

    ReDim ids(0 To count - 1)
    Dim idx As Long
    idx = 0
    Dim i As Long
    For i = 0 To Me.lstCronograma.ListCount - 1
        If Me.lstCronograma.Selected(i) Then
            ids(idx) = Me.lstCronograma.List(i, LST_COL_ID)
            idx = idx + 1
        End If
    Next i

    ObtenerIDsSeleccionados = ids
End Function
