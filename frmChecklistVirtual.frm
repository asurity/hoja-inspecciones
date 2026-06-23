' ======================================================================
' UserForm: frmChecklistVirtual
' Descripción: Formulario de Checklist Virtual para inspecciones.
'              Contiene cabecera con datos del evaluado, MultiPage con
'              2 pestañas de preguntas (Auditoría de procesos + Técnica
'              aséptica), observación general y botones Guardar/Cancelar.
' Fecha creación: 14/04/2026
' Dependencias: ChecklistRepository, ChecklistOrchestrator, Configuration2,
'               ErrorLogger2
'
' LAYOUT:
'   ┌─────────────────────┬──────────────────────────────┐
'   │ fraCabecera (35%)   │  mpPreguntas (65%)           │
'   │ ┌─────────────────┐ │  2 pestañas con              │
'   │ │ 14 campos       │ │  preguntas dinámicas         │
'   │ │ verticales      │ │                              │
'   │ ├─────────────────┤ │                              │
'   │ │ Inspección      │ │                              │
'   │ │ Recurrente      │ │                              │
'   │ │ (con scroll)    │ │                              │
'   │ └─────────────────┘ │                              │
'   ├─────────────────────┴──────────────────────────────┤
'   │ Obs. Generales              [Guardar] [Cancelar]   │
'   └────────────────────────────────────────────────────┘
'
' CONTROLES REQUERIDOS EN EL DISEÑADOR VBA:
'   Formulario: frmChecklistVirtual (Adaptativo a pantalla completa, 95%x85%, StartUpPosition=CenterScreen)
'
'   Frame "fraCabecera" (columna izquierda, 35% del ancho, mín 280pt, diseño vertical con scroll):
'     18 campos verticales (actualizado 23/04/2026):
'       - lblEvaluado, txtEvaluado (TextBox, Locked)
'       - lblPuesto, txtPuesto (TextBox, Locked)
'       - lblPlanta, txtPlanta (TextBox, Locked)
'       - lblArea, cboArea (ComboBox)
'       - lblLineaAuditada, cboLineaAuditada (ComboBox)
'       - lblFecha, txtFecha (TextBox)
'       - lblFechaAuditada, txtFechaAuditada (TextBox)
'       - lblHoraInicio, txtHoraInicio (TextBox)
'       - lblHoraTermino, txtHoraTermino (TextBox)
'       - lblEvaluador, cboEvaluador (ComboBox)
'       - lblAY1, cboAY1 (ComboBox)
'       - lblAY2, cboAY2 (ComboBox)
'       - lblOP, cboOP (ComboBox)
'       - lblLugar, cboLugar (ComboBox)
'       - lblCalificacionVestuario, cboCalificacionVestuario (ComboBox Si/No) [NUEVO 23/04/2026]
'       - lblFechaVencVestuario, txtFechaVencVestuario (TextBox con validación fecha) [NUEVO 23/04/2026]
'       - lblCalificacionOperador, cboCalificacionOperador (ComboBox Si/No) [NUEVO 23/04/2026]
'       - lblFechaVencOperador, txtFechaVencOperador (TextBox con validación fecha) [NUEVO 23/04/2026]
'     Inspecciones recurrentes (DEBAJO de los 18 campos):
'       - fraRecurrentInspection (Frame contenedor)
'         - btnBuscarHistorico (CommandButton)
'         - lblInfoHistorico (Label)
'         - lblNumeroInspeccion, txtNumeroInspeccion (Label, TextBox)
'         - lblRPNAnterior, txtRPNAnteriorAuto, txtRPNAnteriorManual (Label, 2 TextBox)
'         - lblModoRPN (Label)
'         - lblPorcRecuperacion, txtPorcRecuperacion (Label, TextBox) [NUEVO 23/04/2026]
'         - lblPorcOOL, txtPorcOOL (Label, TextBox) [NUEVO 23/04/2026]
'
'   MultiPage "mpPreguntas" (columna derecha, 65% del ancho, 2 páginas):
'     Page 0: "Auditoría de procesos" → fraPreguntas0 (Frame, ScrollBars=Vertical)
'     Page 1: "Técnica aséptica"      → fraPreguntas1 (Frame, ScrollBars=Vertical)
'
'   Sección inferior (ancho completo):
'     - lblObsGeneral (Label)
'     - txtObsGeneral (TextBox, MultiLine, ScrollBars=Vertical)
'     - btnGuardar (CommandButton)
'     - btnCancelar (CommandButton)
' ======================================================================
Option Explicit

' --- Win32 API para métricas de pantalla ---
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
#Else
    'Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    'Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    'Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    'Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

' --- Constantes de diseño ---
Private Const MARGIN As Single = 12
Private Const ROW_HEIGHT As Single = 22
Private Const LABEL_WIDTH As Single = 100
Private Const LEFT_COL_MIN As Single = 280  ' Mínimo para columna izquierda con scroll
Private Const COL_GAP As Single = 8
Private Const OBS_LABEL_H As Single = 18
Private Const OBS_TEXT_H As Single = 50
Private Const BTN_AREA_H As Single = 38

' --- Constantes para preguntas dinámicas ---
Private Const PREG_LABEL_HEIGHT As Single = 30
Private Const PREG_COMBO_HEIGHT As Single = 20
Private Const PREG_OBS_HEIGHT As Single = 20
Private Const PREG_BLOCK_HEIGHT As Single = 80  ' total por pregunta
Private Const PREG_MARGIN As Single = 6

' --- Colores (paleta consistente con frmGestorTablas) ---
Private Const COLOR_FONDO As Long = &HFFFFFF           ' Blanco (fondo formulario)
Private Const COLOR_TITULO As Long = &H724E27           ' Marrón oscuro (títulos)
Private Const COLOR_FRAME As Long = &HFFFFFF            ' Blanco (fondo frames/cabecera)
Private Const COLOR_FRAME_PREGUNTAS As Long = &HFFFFFF   ' Blanco (fondo preguntas)
Private Const COLOR_LABEL As Long = &H5B3A1A            ' Marrón medio (labels de campo)
Private Const COLOR_READONLY As Long = &HFFFFFF         ' Blanco (no editable)
Private Const COLOR_GUARDAR As Long = &HC0DCC0           ' Verde claro
Private Const COLOR_CANCELAR As Long = &HC0C0C0          ' Gris
Private Const COLOR_ESTADO_OK As Long = &H8000&          ' Verde (mensajes ok)
Private Const COLOR_ESTADO_ERROR As Long = &HFF&         ' Rojo (mensajes error)

' --- Variables de layout dinámico (calculadas desde resolución de pantalla) ---
Private FORM_WIDTH As Single
Private FORM_HEIGHT As Single
Private CONTENT_WIDTH As Single
Private mLeftColWidth As Single
Private mRightColLeft As Single
Private mRightColWidth As Single
Private mContentHeight As Single
Private mObsTop As Single
Private mBtnTop As Single

' --- Estado interno ---
Private mEvaluado As String
Private mPuesto As String
Private mIDPlantilla As String
Private mIDCronograma As String
Private mPlanta As String
Private mAreaPendiente As String  ' Área asignada antes de cargar combos
Private mFrecuenciaMeses As Long
Private mIDSeccionTA As String
Private mIDSeccionProcesos As String

' Colección de secciones: cada item es un array(ID, Nombre, TipoRespuesta)
Private mSecciones As Collection

' Diccionario de respuestas: Key=IDPregunta, Value=Dictionary("IDOpcion","ValorNumerico","Observacion","IDSeccion")
Private mRespuestas As Object

' Mapeo de controles dinámicos para acceder después
Private mPreguntaIDs As Collection      ' Collection de ID_Pregunta por sección
Private mPreguntaSecciones As Object    ' Dictionary: IDPregunta → IDSeccion

' --- Estado inspecciones recurrentes (FASE 2 - 21/04/2026) ---
Private mEsInspeccionRecurrente As Boolean
Private mNumeroInspeccion As Long
Private mRPNAnteriorManual As Double
Private mRPNAnteriorAuto As Double
Private mIDInspeccionAnterior As String
Private mModoRPN As String  ' "AUTO" o "MANUAL"
Private mModoRecurrenteManual As Boolean  ' True si el usuario activó modo recurrente sin historial (carga manual)

' --- Factores adicionales (FASE 6 - 23/04/2026) ---
Private mPorcRecuperacion As Double
Private mPorcOOL As Double
Private mRequiereFactoresAdicionales As Boolean  ' Depende del tipo de checklist

' --- Calificaciones y vencimientos (FASE 7 - 23/04/2026) ---
Private mCalificacionVestuario As String
Private mFechaVencVestuario As String
Private mCalificacionOperador As String
Private mFechaVencOperador As String

' ======================================================================
' PROPIEDADES PÚBLICAS (Let/Get)
' ======================================================================

Public Property Let Evaluado(ByVal v As String)
    mEvaluado = v
End Property
Public Property Get Evaluado() As String
    Evaluado = mEvaluado
End Property

Public Property Let Puesto(ByVal v As String)
    mPuesto = v
End Property
Public Property Get Puesto() As String
    Puesto = mPuesto
End Property

Public Property Let IDPlantilla(ByVal v As String)
    mIDPlantilla = v
End Property
Public Property Get IDPlantilla() As String
    IDPlantilla = mIDPlantilla
End Property

Public Property Let IDCronograma(ByVal v As String)
    mIDCronograma = v
End Property
Public Property Get IDCronograma() As String
    IDCronograma = mIDCronograma
End Property

Public Property Let Planta(ByVal v As String)
    mPlanta = v
End Property
Public Property Get Planta() As String
    Planta = mPlanta
End Property

' ACTUALIZADO 23/06/2026: Usar .Text en vez de .Value para garantizar String
' (evita conversión implícita Date→String con locale del sistema)
Public Property Get FechaInspeccion() As String
    FechaInspeccion = Trim(txtFecha.Text)
End Property

Public Property Get FechaAuditada() As String
    FechaAuditada = Trim(txtFechaAuditada.Text)
End Property

Public Property Get HoraInicio() As String
    HoraInicio = Trim(txtHoraInicio.Value)
End Property

Public Property Get HoraTermino() As String
    HoraTermino = Trim(txtHoraTermino.Value)
End Property

Public Property Let Area(ByVal v As String)
    ' Guardar el valor del área para procesarlo después de cargar combos
    mAreaPendiente = v
End Property

Public Property Get Area() As String
    Area = Trim(cboArea.Value)
End Property

Public Property Get LineaAuditada() As String
    LineaAuditada = Trim(cboLineaAuditada.Value)
End Property

Public Property Get Evaluador() As String
    Evaluador = Trim(cboEvaluador.Value)
End Property

Public Property Get AY1() As String
    AY1 = Trim(cboAY1.Value)
End Property

Public Property Get AY2() As String
    AY2 = Trim(cboAY2.Value)
End Property

Public Property Get OP() As String
    OP = Trim(cboOP.Value)
End Property

Public Property Get LugarAuditoria() As String
    LugarAuditoria = Trim(cboLugar.Value)
End Property

Public Property Get ObservacionGeneral() As String
    ObservacionGeneral = Trim(txtObsGeneral.Value)
End Property

Public Property Get CalificacionVestuario() As String
    ' Si el control no está visible, el campo no aplica para este puesto
    If Not cboCalificacionVestuario.Visible Then
        CalificacionVestuario = "-"
    Else
        CalificacionVestuario = Trim(cboCalificacionVestuario.Value)
    End If
End Property

' ACTUALIZADO 23/06/2026: Usar .Text en vez de .Value para garantizar String
Public Property Get FechaVencVestuario() As String
    ' Si el control no está visible, el campo no aplica para este puesto
    If Not txtFechaVencVestuario.Visible Then
        FechaVencVestuario = "-"
    Else
        FechaVencVestuario = Trim(txtFechaVencVestuario.Text)
    End If
End Property

Public Property Get CalificacionOperador() As String
    ' Si el control no está visible, el campo no aplica para este puesto
    If Not cboCalificacionOperador.Visible Then
        CalificacionOperador = "-"
    Else
        CalificacionOperador = Trim(cboCalificacionOperador.Value)
    End If
End Property

Public Property Get FechaVencOperador() As String
    ' Si el control no está visible, el campo no aplica para este puesto
    If Not txtFechaVencOperador.Visible Then
        FechaVencOperador = "-"
    Else
        FechaVencOperador = Trim(txtFechaVencOperador.Text)
    End If
End Property

Public Property Get FrecuenciaMeses() As Long
    FrecuenciaMeses = mFrecuenciaMeses
End Property

Public Property Get IDSeccionTA() As String
    IDSeccionTA = mIDSeccionTA
End Property

Public Property Get IDSeccionProcesos() As String
    IDSeccionProcesos = mIDSeccionProcesos
End Property

' --- Propiedades inspecciones recurrentes (FASE 2 - 21/04/2026) ---
Public Property Get EsInspeccionRecurrente() As Boolean
    EsInspeccionRecurrente = mEsInspeccionRecurrente
End Property

Public Property Get NumeroInspeccion() As Long
    NumeroInspeccion = mNumeroInspeccion
End Property

Public Property Get RPNAnteriorManual() As Double
    RPNAnteriorManual = mRPNAnteriorManual
End Property

Public Property Get RPNAnteriorAuto() As Double
    RPNAnteriorAuto = mRPNAnteriorAuto
End Property

Public Property Get IDInspeccionAnterior() As String
    IDInspeccionAnterior = mIDInspeccionAnterior
End Property

Public Property Get ModoRPN() As String
    ModoRPN = mModoRPN
End Property

' --- Factores adicionales (FASE 6 - 23/04/2026) ---
Public Property Get PorcRecuperacion() As Double
    PorcRecuperacion = mPorcRecuperacion
End Property

Public Property Get PorcOOL() As Double
    PorcOOL = mPorcOOL
End Property

Public Property Get RequiereFactoresAdicionales() As Boolean
    RequiereFactoresAdicionales = mRequiereFactoresAdicionales
End Property

' ======================================================================
' INICIALIZACIÓN
' ======================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    ' Inicializar colecciones
    Set mRespuestas = CreateObject("Scripting.Dictionary")
    Set mPreguntaIDs = New Collection
    Set mPreguntaSecciones = CreateObject("Scripting.Dictionary")
    Set mSecciones = New Collection
    
    ' Inicializar estado inspecciones recurrentes (FASE 2)
    mEsInspeccionRecurrente = False
    mNumeroInspeccion = 1
    mRPNAnteriorManual = 0
    mRPNAnteriorAuto = 0
    mIDInspeccionAnterior = ""
    mModoRPN = "NINGUNO"
    mModoRecurrenteManual = False
    mAreaPendiente = ""
    
    ' Inicializar factores adicionales (FASE 6 - 23/04/2026)
    mPorcRecuperacion = 0
    mPorcOOL = 0
    mRequiereFactoresAdicionales = False
    
    ' --- Configurar todos los controles ---
    Call ConfigurarFormulario
    Call ConfigurarCabecera
    Call ConfigurarMultiPage
    Call ConfigurarObservacionGeneral
    Call ConfigurarBotones
    
    ' Debug.Print "UserForm_Initialize completado OK"
    
    Exit Sub
    
ErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    ' Debug.Print "ERROR UserForm_Initialize: [" & errNum & "] " & errDesc
    Call ErrorLogger2.Log("frmChecklistVirtual.UserForm_Initialize", errDesc, errNum)
End Sub

' ======================================================================
' SECCIÓN: CONFIGURACIÓN DE CONTROLES (Posicionamiento y estilos)
' ======================================================================

Private Sub ConfigurarFormulario()
    ' --- Obtener métricas de pantalla para dimensionar el formulario ---
    Dim scrW_px As Long, scrH_px As Long
    Dim scrW_pt As Single, scrH_pt As Single
    
    scrW_px = GetSystemMetrics(SM_CXSCREEN)
    scrH_px = GetSystemMetrics(SM_CYSCREEN)
    
    If scrW_px > 0 And scrH_px > 0 Then
        Dim dpiX As Long, dpiY As Long
        #If VBA7 Then
            Dim hDC As LongPtr
        #Else
            Dim hDC As Long
        #End If
        hDC = GetDC(0)
        dpiX = GetDeviceCaps(hDC, LOGPIXELSX)
        dpiY = GetDeviceCaps(hDC, LOGPIXELSY)
        ReleaseDC 0, hDC
        If dpiX = 0 Then dpiX = 96
        If dpiY = 0 Then dpiY = 96
        
        ' Convertir píxeles a puntos (1 punto = 1/72 pulgada)
        scrW_pt = CSng(scrW_px) * 72# / CSng(dpiX)
        scrH_pt = CSng(scrH_px) * 72# / CSng(dpiY)
        
        ' Aprovechar más espacio de pantalla (95% ancho, 85% alto)
        FORM_WIDTH = scrW_pt * 0.95
        FORM_HEIGHT = scrH_pt * 0.85
    Else
        ' Fallback si la API no responde
        FORM_WIDTH = 900
        FORM_HEIGHT = 700
        scrW_pt = 0
        scrH_pt = 0
    End If
    
    ' Límites razonables ampliados
    If FORM_WIDTH < 900 Then FORM_WIDTH = 900
    If FORM_WIDTH > 1400 Then FORM_WIDTH = 1400
    If FORM_HEIGHT < 650 Then FORM_HEIGHT = 650
    If FORM_HEIGHT > 900 Then FORM_HEIGHT = 900
    
    ' Calcular ancho total de contenido
    CONTENT_WIDTH = FORM_WIDTH - (MARGIN * 2) - 4
    
    ' Calcular anchos de columnas (cabecera 35%, preguntas 65% para mejor distribución)
    mLeftColWidth = CONTENT_WIDTH * 0.35
    If mLeftColWidth < 280 Then mLeftColWidth = 280
    If mLeftColWidth > 380 Then mLeftColWidth = 380
    mRightColLeft = MARGIN + mLeftColWidth + COL_GAP
    mRightColWidth = CONTENT_WIDTH - mLeftColWidth - COL_GAP
    
    ' Calcular posiciones verticales dinámicas
    ' InsideHeight ≈ Height - barra de título (~26pt) - bordes (~6pt)
    Dim availH As Single
    availH = FORM_HEIGHT - 46
    mBtnTop = availH - BTN_AREA_H
    mObsTop = mBtnTop - OBS_TEXT_H - OBS_LABEL_H - 4
    mContentHeight = mObsTop - MARGIN - 4
    If mContentHeight < 280 Then mContentHeight = 280
    
    ' --- Aplicar propiedades del formulario ---
    Me.Caption = "Checklist Virtual — Inspección"
    Me.Width = FORM_WIDTH
    Me.Height = FORM_HEIGHT
    Me.BackColor = COLOR_FONDO
    
    ' Centrar en pantalla
    If scrW_pt > 0 Then
        Me.StartUpPosition = 0  ' Manual
        Me.Left = (scrW_pt - FORM_WIDTH) / 2
        Me.Top = (scrH_pt - FORM_HEIGHT) / 2
    Else
        Me.StartUpPosition = 1  ' CenterOwner como fallback
    End If
End Sub

Private Sub ConfigurarCabecera()
    ' --- Frame principal de cabecera (columna izquierda, 1 columna vertical con scroll) ---
    With Me.fraCabecera
        .Left = MARGIN
        .Top = MARGIN
        .Width = mLeftColWidth
        .Height = mContentHeight
        .Caption = " Datos de la inspección "
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_TITULO
        .BackColor = COLOR_FRAME
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = &HD0C8C0
        .SpecialEffect = fmSpecialEffectFlat
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = 840  ' Altura total: 18 campos (396pt) + frame recurrente (240pt) + márgenes (actualizad o 23/04/2026)
        .KeepScrollBarsVisible = fmScrollBarsVertical
    End With
    
    ' --- Layout interno: 1 columna vertical ---
    Dim lblLeft As Single: lblLeft = 4
    Dim ctrlLeft As Single: ctrlLeft = 100
    Dim ctrlW As Single: ctrlW = mLeftColWidth - ctrlLeft - 20  ' Espacio para scrollbar
    Dim rowTop As Single: rowTop = 20
    Dim rowIdx As Long: rowIdx = 0
    
    If ctrlW < 120 Then ctrlW = 120
    
    ' FILA 1: Evaluado
    With Me.lblEvaluado
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Evaluado:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.txtEvaluado
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = True
        .BackColor = COLOR_READONLY
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 2: Puesto
    With Me.lblPuesto
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Puesto:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.txtPuesto
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = True
        .BackColor = COLOR_READONLY
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 3: Planta
    With Me.lblPlanta
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Planta:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.txtPlanta
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = True
        .BackColor = COLOR_READONLY
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 4: Área
    With Me.lblArea
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Área:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.cboArea
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Style = fmStyleDropDownList
        .TabIndex = 0
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 5: Línea/Equipo
    With Me.lblLineaAuditada
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Línea/Equipo:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.cboLineaAuditada
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .Style = fmStyleDropDownList
        .TabIndex = 1
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 6: Fecha
    With Me.lblFecha
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Fecha:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.txtFecha
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .TabIndex = 2
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 7: Fecha Auditada
    With Me.lblFechaAuditada
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Fecha Auditada:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.txtFechaAuditada
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .TabIndex = 3
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 8: Hora inicio
    With Me.lblHoraInicio
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Hora inicio:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.txtHoraInicio
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .TabIndex = 4
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 9: Hora término
    With Me.lblHoraTermino
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Hora término:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.txtHoraTermino
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .TabIndex = 5
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 10: Evaluador
    With Me.lblEvaluador
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Evaluador:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.cboEvaluador
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Style = fmStyleDropDownList
        .TabIndex = 6
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 11: Ayudante 1
    With Me.lblAY1
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Ayudante 1:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.cboAY1
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Style = fmStyleDropDownList
        .TabIndex = 7
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 12: Ayudante 2
    With Me.lblAY2
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Ayudante 2:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.cboAY2
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Style = fmStyleDropDownList
        .TabIndex = 8
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 13: Operador
    With Me.lblOP
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Operador:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.cboOP
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Style = fmStyleDropDownList
        .TabIndex = 9
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 14: Lugar auditoría
    With Me.lblLugar
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Lugar:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.cboLugar
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Style = fmStyleDropDownList
        .TabIndex = 10
    End With
    rowIdx = rowIdx + 1
    
    ' ═══════════════════════════════════════════════════════════════════
    ' NUEVOS CAMPOS: CALIFICACIONES Y VENCIMIENTOS (FASE 7 - 23/04/2026)
    ' ═══════════════════════════════════════════════════════════════════
    
    ' FILA 15: Calificación de Vestuario
    With Me.lblCalificacionVestuario
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Calif. Vestuario:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.cboCalificacionVestuario
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Style = fmStyleDropDownList
        .TabIndex = 11
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 16: Fecha Vencimiento Vestuario
    With Me.lblFechaVencVestuario
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Venc. Vestuario:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.txtFechaVencVestuario
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .TabIndex = 12
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 17: Calificación de Operador
    With Me.lblCalificacionOperador
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Calif. Operador:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.cboCalificacionOperador
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Style = fmStyleDropDownList
        .TabIndex = 13
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 18: Fecha Vencimiento Operador
    With Me.lblFechaVencOperador
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = 95
        .Height = 18
        .Caption = "Venc. Operador:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    With Me.txtFechaVencOperador
        .Left = ctrlLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = ctrlW
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .TabIndex = 14
    End With
    rowIdx = rowIdx + 1
    
    ' ═══════════════════════════════════════════════════════════════════
    ' SECCIÓN: INSPECCIONES RECURRENTES (FASE 2 - 21/04/2026)
    ' BOTÓN TOGGLE FUERA DEL FRAME - SIEMPRE VISIBLE
    ' ═══════════════════════════════════════════════════════════════════
    
    ' Botón toggle para modo recurrente (FUERA del frame, debajo de los nuevos campos)
    With Me.btnToggleRecurrente
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = mLeftColWidth - 12
        .Height = 24
        .Caption = "Activar Modo Recurrente"
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .Font.Bold = False
        .TabIndex = 15
        .TabStop = True
        .BackColor = &HFFFFFF
        .Enabled = True
        .Visible = True
        .ZOrder 0
    End With
    rowIdx = rowIdx + 1
    
    ' Calcular posición vertical para el frame (debajo del botón con MÁS separación)
    Dim recTop As Single
    recTop = rowTop + (rowIdx * ROW_HEIGHT) + 30
    
    ' Frame contenedor (ancho completo de la columna)
    With Me.fraRecurrentInspection
        .Left = lblLeft
        .Top = recTop
        .Width = mLeftColWidth - 12
        .Height = 240
        .Caption = " Inspección Recurrente "
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .ForeColor = &H5B3A1A  ' Marrón medio
        .BackColor = &HFFFFFF  ' Blanco
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = &HD0C8C0
        .SpecialEffect = fmSpecialEffectFlat
        .ZOrder 1  ' CRÍTICO: Frame detrás del checkbox
    End With
    
    ' Ancho interno para controles del frame
    Dim frameInnerW As Single
    frameInnerW = mLeftColWidth - 28
    
    ' Botón búsqueda histórico
    With Me.btnBuscarHistorico
        .Left = 8
        .Top = 42
        .Width = frameInnerW
        .Height = 28
        .Caption = "Actualizar historial"
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .TabIndex = 12
        .BackColor = &HFFFFFF
    End With
    
    ' Label info histórico
    With Me.lblInfoHistorico
        .Left = 8
        .Top = 74
        .Width = frameInnerW
        .Height = 32
        .Caption = "(Info de inspecciones previas aparecerá aquí)"
        .Font.Name = "Segoe UI"
        .Font.Size = 7
        .ForeColor = &H808080  ' Gris
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
    End With
    
    ' Label número inspección
    With Me.lblNumeroInspeccion
        .Left = 8
        .Top = 110
        .Width = 100
        .Height = 16
        .Caption = "Inspección N°:"
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignLeft
        .Visible = False
    End With
    
    ' TextBox número inspección
    With Me.txtNumeroInspeccion
        .Left = 112
        .Top = 110
        .Width = frameInnerW - 104
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = True
        .BackColor = COLOR_READONLY
        .TabStop = False
        .Visible = False
    End With
    
    ' Label RPN Anterior
    With Me.lblRPNAnterior
        .Left = 8
        .Top = 136
        .Width = 100
        .Height = 16
        .Caption = "%TA anterior:"
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignLeft
        .Visible = False
    End With
    
    ' TextBox RPN Anterior Automático
    With Me.txtRPNAnteriorAuto
        .Left = 112
        .Top = 136
        .Width = frameInnerW - 104
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = True
        .BackColor = &HFFFFFF  ' Blanco
        .TabStop = False
        .Visible = False
    End With
    
    ' TextBox RPN Anterior Manual
    With Me.txtRPNAnteriorManual
        .Left = 112
        .Top = 136
        .Width = frameInnerW - 104
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = False
        .BackColor = &HFFFFFF  ' Blanco
        .Visible = False
    End With
    
    ' Label estado modo
    With Me.lblModoRPN
        .Left = 8
        .Top = 162
        .Width = frameInnerW
        .Height = 36
        .Caption = "[Modo RPN: no determinado]"
        .Font.Name = "Segoe UI"
        .Font.Size = 7
        .ForeColor = &H808080  ' Gris
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
        .Visible = False
    End With
    
    ' ═══════════════════════════════════════════════════════════════════
    ' FACTORES ADICIONALES (FASE 6 - 23/04/2026)
    ' Solo visible en inspecciones recurrentes de ciertos tipos de checklist
    ' ═══════════════════════════════════════════════════════════════════
    
    ' Label % Recuperación
    With Me.lblPorcRecuperacion
        .Left = 8
        .Top = 204
        .Width = 100
        .Height = 16
        .Caption = "% Recuperación:"
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignLeft
        .Visible = False
    End With
    
    ' TextBox % Recuperación
    With Me.txtPorcRecuperacion
        .Left = 112
        .Top = 204
        .Width = frameInnerW - 104
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = False
        .BackColor = &HFFFFFF  ' Blanco
        .Visible = False
    End With
    
    ' Label % OOL
    With Me.lblPorcOOL
        .Left = 8
        .Top = 230
        .Width = 100
        .Height = 16
        .Caption = "% OOL:"
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignLeft
        .Visible = False
    End With
    
    ' TextBox % OOL
    With Me.txtPorcOOL
        .Left = 112
        .Top = 230
        .Width = frameInnerW - 104
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = False
        .BackColor = &HFFFFFF  ' Blanco
        .Visible = False
    End With
    
    ' Ajustar altura del frame para incluir factores adicionales
    Me.fraRecurrentInspection.Height = 260
    
    ' Ocultar frame de inspección recurrente al inicio
    ' (Se mostrará automáticamente si hay historial al activar el formulario)
    fraRecurrentInspection.Visible = False
End Sub

Private Sub ConfigurarMultiPage()
    ' --- MultiPage para preguntas (columna derecha, tamaño dinámico) ---
    ' Los frames internos se redimensionan en UserForm_Activate
    ' donde el form ya está renderizado y las métricas son confiables.
    With Me.mpPreguntas
        .Left = mRightColLeft
        .Top = MARGIN
        .Width = mRightColWidth
        .Height = mContentHeight
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .Style = fmTabStyleTabs
        .MultiRow = False
        .TabIndex = 11
    End With
End Sub

Private Sub ConfigurarObservacionGeneral()
    ' --- Label para observación general ---
    With Me.lblObsGeneral
        .Left = MARGIN
        .Top = mObsTop
        .Width = 200
        .Height = OBS_LABEL_H
        .Caption = "Observaciones generales:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = COLOR_TITULO
        .BackStyle = fmBackStyleTransparent
    End With
    
    ' --- TextBox multilínea (debajo del label) ---
    With Me.txtObsGeneral
        .Left = MARGIN
        .Top = mObsTop + OBS_LABEL_H
        .Width = CONTENT_WIDTH
        .Height = OBS_TEXT_H
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .EnterKeyBehavior = True
        .WordWrap = True
        .TabIndex = 12
    End With
End Sub

Private Sub ConfigurarBotones()
    Dim btnWidth As Single: btnWidth = 110
    Dim btnHeight As Single: btnHeight = 34
    Dim btnSpacing As Single: btnSpacing = 12
    
    ' --- Botón Guardar (izquierda) ---
    With Me.btnGuardar
        .Left = CONTENT_WIDTH - (btnWidth * 2) - btnSpacing
        .Top = mBtnTop
        .Width = btnWidth
        .Height = btnHeight
        .Caption = "Guardar"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = COLOR_GUARDAR
        .TabIndex = 13
    End With
    
    ' --- Botón Cancelar (derecha) ---
    With Me.btnCancelar
        .Left = CONTENT_WIDTH - btnWidth
        .Top = mBtnTop
        .Width = btnWidth
        .Height = btnHeight
        .Caption = "Cancelar"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackColor = COLOR_CANCELAR
        .TabIndex = 14
    End With
End Sub

'' ----------------------------------------------------------------------
' Subrutina: RedimensionarFramesPreguntas
' Propósito: Ajusta los frames fraPreguntas0/1 al tamaño real del MultiPage.
'            Debe llamarse en Activate (NO en Initialize) porque los
'            tamaños reales de las páginas solo están disponibles después
'            de que el formulario se ha renderizado.
' ----------------------------------------------------------------------
Private Sub RedimensionarFramesPreguntas()
    On Error Resume Next
    
    Dim i As Long
    Dim fraName As String
    Dim fra As MSForms.Frame
    Dim pgW As Single, pgH As Single
    
    For i = 0 To mpPreguntas.Pages.Count - 1
        ' Activar la página para asegurar métricas correctas
        mpPreguntas.Value = i
        DoEvents
        
        fraName = "fraPreguntas" & i
        Set fra = Nothing
        Set fra = mpPreguntas.Pages(i).Controls(fraName)
        
        If Not fra Is Nothing Then
            pgW = mpPreguntas.Pages(i).InsideWidth
            pgH = mpPreguntas.Pages(i).InsideHeight
            
            ' Fallback si InsideWidth/Height no son confiables
            If pgW < 100 Then pgW = mRightColWidth - 12
            If pgH < 50 Then pgH = mContentHeight - 30
            
            fra.Left = 0
            fra.Top = 0
            fra.Width = pgW
            fra.Height = pgH
            fra.ScrollBars = fmScrollBarsVertical
            fra.BackColor = COLOR_FRAME
            fra.BorderStyle = fmBorderStyleNone
            fra.SpecialEffect = fmSpecialEffectFlat
            fra.Caption = ""
            
            ' Debug.Print "  Frame " & fraName & " redimensionado: " & pgW & " x " & pgH
        Else
            ' Debug.Print "  Frame " & fraName & " NO encontrado"
        End If
    Next i
    
    ' Volver a primera página
    mpPreguntas.Value = 0
    DoEvents
    
    On Error GoTo 0
End Sub

'' ----------------------------------------------------------------------
' Subrutina: UserForm_Activate
' Propósito: Se ejecuta al mostrar el form. Carga datos después de que
'            las propiedades han sido asignadas por el orquestador.
' ----------------------------------------------------------------------
Private Sub UserForm_Activate()
    On Error GoTo ErrorHandler
    
    ' Validar que tenemos datos mínimos
    If Len(mEvaluado) = 0 Or Len(mPuesto) = 0 Or Len(mIDPlantilla) = 0 Then
        MsgBox "ERROR: Datos incompletos." & vbCrLf & _
               "Evaluado: [" & mEvaluado & "]" & vbCrLf & _
               "Puesto: [" & mPuesto & "]" & vbCrLf & _
               "Plantilla: [" & mIDPlantilla & "]" & vbCrLf & vbCrLf & _
               "No se pueden cargar los controles dinámicos.", _
               vbCritical, "Datos inválidos"
        Exit Sub
    End If
    
    ' Rellenar campos de solo lectura
    On Error Resume Next
    txtEvaluado.Value = mEvaluado
    txtEvaluado.Locked = True
    txtEvaluado.BackColor = &H8000000F  ' Color gris (bloqueado)
    
    txtPuesto.Value = mPuesto
    txtPuesto.Locked = True
    txtPuesto.BackColor = &H8000000F
    
    txtPlanta.Value = mPlanta
    txtPlanta.Locked = True
    txtPlanta.BackColor = &H8000000F
    
    txtFecha.Value = Format(Date, "dd/mm/yyyy")
    txtFechaAuditada.Value = Format(Date, "dd/mm/yyyy")
    
    ' Inicializar horas con la hora actual
    txtHoraInicio.Value = Format(Now, "HH:mm")
    txtHoraTermino.Value = Format(Now, "HH:mm")
    On Error GoTo ErrorHandler
    
    ' Debug.Print "Campos básicos cargados OK"
    
    ' --- Redimensionar frames de preguntas (ahora que el form está renderizado) ---
    ' Debug.Print "Redimensionando frames de preguntas..."
    Call RedimensionarFramesPreguntas
    
    ' --- Cargar combos con manejo seguro ---
    ' Debug.Print "Intentando cargar combos..."
    
    On Error Resume Next
    Call CargarComboAreas
    ' Debug.Print "  Areas: OK"
    Call ConfigurarCampoLineaPorPuesto  ' Configurar línea/equipo según puesto
    ' Debug.Print "  Campo Línea configurado según puesto: OK"
    Call AplicarFiltroArea  ' Aplicar filtrado/bloqueo según planta y área
    ' Debug.Print "  Filtro de área aplicado: OK"
    Call CargarCombosPersonal
    ' Debug.Print "  Personal: OK"
    Call CargarComboEvaluadores
    ' Debug.Print "  Evaluadores: OK"
    Call CargarComboLugar
    ' Debug.Print "  Lugar: OK"
    Call CargarCombosCalificacion  ' NUEVO: Cargar combos de calificación (FASE 7 - 23/04/2026)
    ' Debug.Print "  Calificaciones: OK"
    Call ConfigurarVisibilidadCalificaciones  ' FASE 7: Configurar visibilidad según puesto
    ' Debug.Print "  Visibilidad calificaciones configurada: OK"
    Call CargarFrecuenciaPlantilla
    ' Debug.Print "  Frecuencia: OK"
    On Error GoTo ErrorHandler
    
    ' Debug.Print "Combos cargados OK"
    
    ' --- Cargar secciones y preguntas (lo más crítico) ---
    ' Debug.Print "Intentando cargar secciones y preguntas..."
    
    Call CargarSecciones
    ' Debug.Print "Secciones cargadas: " & mSecciones.Count
    
    If mSecciones.Count > 0 Then
        Call CargarPreguntasDinamicas
        ' Debug.Print "CargarPreguntasDinamicas completado"
    Else
        ' Debug.Print "ERROR: No hay secciones. Mostrando form sin preguntas."
        MsgBox "No se encontraron secciones configuradas para la plantilla '" & mIDPlantilla & "'." & vbCrLf & vbCrLf & _
               "El formulario se cerrará. Verifique la configuración de secciones en tblSecciones.", _
               vbCritical, "Error de inicialización"
        Unload Me
        Exit Sub
    End If
    
    ' VERIFICACIÓN CRÍTICA: asegurar que se cargaron preguntas
    If mPreguntaSecciones.Count = 0 Then
        MsgBox "ERROR CRÍTICO: No se pudieron cargar las preguntas de la plantilla '" & mIDPlantilla & "'." & vbCrLf & vbCrLf & _
               "Verifique que existan preguntas en tblPreguntas para las secciones configuradas." & vbCrLf & vbCrLf & _
               "El formulario se cerrará.", vbCritical, "Error de inicialización"
        Unload Me
        Exit Sub
    End If
    
    ' --- Búsqueda automática de historial (silenciosa) ---
    Call BuscarHistorialSilencioso
    
    Exit Sub
    
ErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    ' Debug.Print "ERROR UserForm_Activate: [" & errNum & "] " & errDesc
    
    On Error Resume Next
    Call ErrorLogger2.Log("frmChecklistVirtual.UserForm_Activate", errDesc, errNum)
    On Error GoTo 0
    
    MsgBox "Error al cargar el formulario:" & vbCrLf & _
           "[" & errNum & "] " & errDesc & vbCrLf & vbCrLf & _
           "El formulario se abrirá con funcionalidad limitada." & vbCrLf & _
           "Por favor, revise los datos y la configuración.", _
           vbExclamation, "Error de inicialización"
End Sub

' ======================================================================
' MÉTODOS PÚBLICOS
' ======================================================================

'' ----------------------------------------------------------------------
' Función: ObtenerRespuestas
' Propósito: Retorna la colección de respuestas para persistencia en
'            tblRespuestas (sin IDSeccion, para InspectionRepository).
' Retorna: Collection de Dictionary con claves:
'   "IDPregunta", "IDOpcion", "ValorNumerico", "Observacion"
' ----------------------------------------------------------------------
Public Function ObtenerRespuestas() As Collection
    On Error GoTo ErrorHandler
    
    Dim resultado As New Collection
    Dim key As Variant
    
    For Each key In mRespuestas.Keys
        Dim dictOrig As Object
        Set dictOrig = mRespuestas(key)
        
        Dim dictResp As Object
        Set dictResp = CreateObject("Scripting.Dictionary")
        dictResp("IDPregunta") = CStr(key)
        dictResp("IDOpcion") = dictOrig("IDOpcion")
        dictResp("ValorNumerico") = dictOrig("ValorNumerico")
        dictResp("Observacion") = dictOrig("Observacion")
        
        resultado.Add dictResp
    Next key
    
    Set ObtenerRespuestas = resultado
    Exit Function
    
ErrorHandler:
    Set ObtenerRespuestas = New Collection
    Call ErrorLogger2.Log("frmChecklistVirtual.ObtenerRespuestas", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerRespuestasConSeccion
' Propósito: Retorna la colección de respuestas incluyendo IDSeccion e IDCriticidad
'            para cálculos de scoring (InspectionCalculator).
' Retorna: Collection de Dictionary con claves:
'   "IDPregunta", "IDOpcion", "ValorNumerico", "Observacion", "IDSeccion", "IDCriticidad"
' ----------------------------------------------------------------------
Public Function ObtenerRespuestasConSeccion() As Collection
    On Error GoTo ErrorHandler
    
    Dim resultado As New Collection
    Dim key As Variant
    Dim contador As Long
    contador = 0
    
    For Each key In mRespuestas.Keys
        contador = contador + 1
        Dim dictOrig As Object
        Set dictOrig = mRespuestas(key)
        
        Dim dictResp As Object
        Set dictResp = CreateObject("Scripting.Dictionary")
        dictResp("IDPregunta") = CStr(key)
        dictResp("IDOpcion") = dictOrig("IDOpcion")
        dictResp("ValorNumerico") = dictOrig("ValorNumerico")
        dictResp("Observacion") = dictOrig("Observacion")
        dictResp("IDSeccion") = dictOrig("IDSeccion")
        dictResp("IDCriticidad") = dictOrig("IDCriticidad")
        
        resultado.Add dictResp
    Next key
    
    Set ObtenerRespuestasConSeccion = resultado
    Exit Function
    
ErrorHandler:
    Set ObtenerRespuestasConSeccion = New Collection
    Call ErrorLogger2.Log("frmChecklistVirtual.ObtenerRespuestasConSeccion", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerCantidadPreguntas
' Propósito: Retorna la cantidad total de preguntas cargadas.
' ----------------------------------------------------------------------
Public Function ObtenerCantidadPreguntas() As Long
    ' Contar el total de preguntas cargadas en el formulario
    ' contando los ComboBox de respuesta (cboR_) en todas las páginas
    On Error GoTo ErrorHandler
    
    Dim totalPreguntas As Long
    totalPreguntas = 0
    
    Dim secIdx As Long
    For secIdx = 0 To mSecciones.Count - 1
        Dim fraName As String
        fraName = "fraPreguntas" & secIdx
        
        Dim fraContainer As MSForms.Frame
        On Error Resume Next
        Set fraContainer = mpPreguntas.Pages(secIdx).Controls(fraName)
        On Error GoTo ErrorHandler
        
        If Not fraContainer Is Nothing Then
            Dim preguntasSeccion As Long
            preguntasSeccion = 0
            
            Dim ctrl As MSForms.Control
            For Each ctrl In fraContainer.Controls
                ' Contar ComboBoxes de respuesta (cboR_)
                If Left(ctrl.Name, 5) = "cboR_" Then
                    totalPreguntas = totalPreguntas + 1
                    preguntasSeccion = preguntasSeccion + 1
                End If
            Next ctrl
        End If
    Next secIdx
    
    ObtenerCantidadPreguntas = totalPreguntas
    Exit Function
    
ErrorHandler:
    ' En caso de error, devolver 0 para que la validación falle
    ObtenerCantidadPreguntas = 0
End Function

'' ----------------------------------------------------------------------
' Subrutina: RecopilarObservacionesPublic
' Propósito: Punto de entrada público para recopilar todas las respuestas
'            y observaciones del formulario. Llamado por el orquestador
'            antes de validar o guardar.
' ----------------------------------------------------------------------
Public Sub RecopilarObservacionesPublic()
    Call RecopilarObservaciones
End Sub

' ======================================================================
' EVENTOS DE CONTROLES
' ======================================================================

Private Sub cboArea_Change()
    On Error GoTo ErrorHandler
    
    ' Si el campo está bloqueado (puestos sin línea), no hacer nada
    If Not cboLineaAuditada.Enabled Then Exit Sub
    
    ' Cascada: al cambiar área, recargar equipos
    cboLineaAuditada.Clear
    
    If Len(Trim(cboArea.Value)) = 0 Then Exit Sub
    
    Dim equipos As Collection
    Set equipos = ChecklistRepository.ObtenerEquiposPorPlantaYArea(mPlanta, Trim(cboArea.Value))
    
    Dim eq As Variant
    For Each eq In equipos
        cboLineaAuditada.AddItem CStr(eq)
    Next eq
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.cboArea_Change", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: txtHoraTermino_Exit
' Propósito: Valida que la hora de término sea mayor o igual a la hora de inicio
' ----------------------------------------------------------------------
Private Sub txtHoraTermino_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    
    Dim horaInicio As String
    Dim horaTermino As String
    Dim dtInicio As Date
    Dim dtTermino As Date
    
    horaInicio = Trim(txtHoraInicio.Value)
    horaTermino = Trim(txtHoraTermino.Value)
    
    ' Validar solo si ambos campos tienen valor
    If Len(horaInicio) = 0 Or Len(horaTermino) = 0 Then Exit Sub
    
    ' Validar formato HH:mm
    If Not IsDate(horaInicio) Or Not IsDate(horaTermino) Then
        MsgBox "El formato de hora debe ser HH:mm (por ejemplo: 14:30).", vbExclamation, "Formato inválido"
        Cancel = True
        Exit Sub
    End If
    
    ' Convertir a Date para comparar
    dtInicio = CDate(horaInicio)
    dtTermino = CDate(horaTermino)
    
    ' Validar que término >= inicio
    If dtTermino < dtInicio Then
        MsgBox "La hora de término debe ser mayor o igual a la hora de inicio." & vbCrLf & vbCrLf & _
               "Hora inicio: " & horaInicio & vbCrLf & _
               "Hora término: " & horaTermino, _
               vbExclamation, "Validación de horarios"
        Cancel = True
        Exit Sub
    End If
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.txtHoraTermino_Exit", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: txtHoraInicio_Exit
' Propósito: Valida formato de hora y coherencia con hora de término
' ----------------------------------------------------------------------
Private Sub txtHoraInicio_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    
    Dim horaInicio As String
    Dim horaTermino As String
    Dim dtInicio As Date
    Dim dtTermino As Date
    
    horaInicio = Trim(txtHoraInicio.Value)
    horaTermino = Trim(txtHoraTermino.Value)
    
    ' Validar solo si ambos campos tienen valor
    If Len(horaInicio) = 0 Or Len(horaTermino) = 0 Then Exit Sub
    
    ' Validar formato HH:mm
    If Not IsDate(horaInicio) Or Not IsDate(horaTermino) Then
        MsgBox "El formato de hora debe ser HH:mm (por ejemplo: 14:30).", vbExclamation, "Formato inválido"
        Cancel = True
        Exit Sub
    End If
    
    ' Convertir a Date para comparar
    dtInicio = CDate(horaInicio)
    dtTermino = CDate(horaTermino)
    
    ' Validar que término >= inicio
    If dtTermino < dtInicio Then
        MsgBox "La hora de inicio debe ser menor o igual a la hora de término." & vbCrLf & vbCrLf & _
               "Hora inicio: " & horaInicio & vbCrLf & _
               "Hora término: " & horaTermino, _
               vbExclamation, "Validación de horarios"
        Cancel = True
        Exit Sub
    End If
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.txtHoraInicio_Exit", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: txtFechaVencVestuario_Exit
' Propósito: Valida que la fecha de vencimiento de vestuario tenga formato válido (dd/mm/yyyy)
'            y que NO esté vencida (debe ser mayor a la fecha actual)
' Fecha: 23/04/2026 - FASE 7
' Actualizado: 24/04/2026 - Validación estricta de fecha futura
' ACTUALIZADO: 23/06/2026 — Usa ParseFechaDMY (independiente de locale) en vez de IsDate/CDate
' ----------------------------------------------------------------------
Private Sub txtFechaVencVestuario_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    
    ' ACTUALIZADO 23/06/2026: Usar .Text para obtener siempre String
    Dim fechaTexto As String
    fechaTexto = Trim(txtFechaVencVestuario.Text)
    
    ' Si está vacío, permitir (campo opcional)
    If Len(fechaTexto) = 0 Then Exit Sub
    
    ' Validar formato de fecha usando ParseFechaDMY (independiente de locale)
    Dim fechaParsed As Variant
    fechaParsed = ChecklistValidator.ParseFechaDMY(fechaTexto)
    
    If IsEmpty(fechaParsed) Then
        MsgBox "La fecha de vencimiento de vestuario debe tener un formato válido (dd/mm/yyyy)." & vbCrLf & _
               "Ejemplo: 31/12/2026", vbExclamation, "Formato de fecha inválido"
        Cancel = True
        Exit Sub
    End If
    
    ' VALIDACIÓN CRÍTICA: La fecha de vencimiento DEBE ser mayor a la fecha actual
    ' Si está vencida, la calificación no es válida y NO se puede continuar
    If CDate(fechaParsed) <= Date Then
        MsgBox "La fecha de vencimiento de vestuario ya pasó (" & fechaTexto & ")." & vbCrLf & vbCrLf & _
               "La calificación de vestuario está VENCIDA." & vbCrLf & _
               "El personal debe renovar su calificación antes de realizar la inspección." & vbCrLf & vbCrLf & _
               "No se puede continuar con una calificación vencida.", _
               vbCritical, "Calificación vencida"
        Cancel = True
        Exit Sub
    End If
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.txtFechaVencVestuario_Exit", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: txtFechaVencOperador_Exit
' Propósito: Valida que la fecha de vencimiento de operador tenga formato válido (dd/mm/yyyy)
'            y que NO esté vencida (debe ser mayor a la fecha actual)
' Fecha: 23/04/2026 - FASE 7
' Actualizado: 24/04/2026 - Validación estricta de fecha futura
' ACTUALIZADO: 23/06/2026 — Usa ParseFechaDMY (independiente de locale) en vez de IsDate/CDate
' ----------------------------------------------------------------------
Private Sub txtFechaVencOperador_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    
    ' ACTUALIZADO 23/06/2026: Usar .Text para obtener siempre String
    Dim fechaTexto As String
    fechaTexto = Trim(txtFechaVencOperador.Text)
    
    ' Si está vacío, permitir (campo opcional)
    If Len(fechaTexto) = 0 Then Exit Sub
    
    ' Validar formato de fecha usando ParseFechaDMY (independiente de locale)
    Dim fechaParsed As Variant
    fechaParsed = ChecklistValidator.ParseFechaDMY(fechaTexto)
    
    If IsEmpty(fechaParsed) Then
        MsgBox "La fecha de vencimiento de operador debe tener un formato válido (dd/mm/yyyy)." & vbCrLf & _
               "Ejemplo: 31/12/2026", vbExclamation, "Formato de fecha inválido"
        Cancel = True
        Exit Sub
    End If
    
    ' VALIDACIÓN CRÍTICA: La fecha de vencimiento DEBE ser mayor a la fecha actual
    ' Si está vencida, la calificación no es válida y NO se puede continuar
    If CDate(fechaParsed) <= Date Then
        MsgBox "La fecha de vencimiento de operador ya pasó (" & fechaTexto & ")." & vbCrLf & vbCrLf & _
               "La calificación de operador está VENCIDA." & vbCrLf & _
               "El personal debe renovar su calificación antes de realizar la inspección." & vbCrLf & vbCrLf & _
               "No se puede continuar con una calificación vencida.", _
               vbCritical, "Calificación vencida"
        Cancel = True
        Exit Sub
    End If
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.txtFechaVencOperador_Exit", Err.Description, Err.Number)
End Sub

Private Sub btnGuardar_Click()
    On Error GoTo ErrorHandler
    
    ' PASO 1: Validar cabecera CON AUTO-CORRECCIÓN en primer lugar
    Dim resultValidacion As Object
    Set resultValidacion = ChecklistValidator.ValidarCabeceraConAutoCorrecion(Me)
    
    ' Si hay errores → Mostrar y detener
    If Not resultValidacion("valido") Then
        Dim mensajeError As String
        Dim i As Long
        mensajeError = "ERRORES EN LOS DATOS:" & vbCrLf & vbCrLf
        Dim errores As Variant
        errores = resultValidacion("errores")
        For i = LBound(errores) To UBound(errores)
            mensajeError = mensajeError & "- " & errores(i) & vbCrLf
        Next i
        MsgBox mensajeError, vbExclamation, "Validacion fallida"
        Exit Sub
    End If
    
    ' Si hay correcciones → Mostrar información
    Dim mensajeInfo As String
    Dim correcciones As Variant
    correcciones = resultValidacion("correcciones")
    If UBound(correcciones) >= 0 Then
        mensajeInfo = "AJUSTES REALIZADOS:" & vbCrLf & vbCrLf
        For i = LBound(correcciones) To UBound(correcciones)
            mensajeInfo = mensajeInfo & "- " & correcciones(i) & vbCrLf
        Next i
        MsgBox mensajeInfo, vbInformation, "Datos corregidos"
    End If
    
    ' PASO 2: Validar datos de inspección recurrente si aplica
    If Not ValidarDatosRecurrentes() Then
        Exit Sub
    End If
    
    ' PASO 2.5: Confirmar factores adicionales (% OOL y % Recuperación) si aplica
    If mEsInspeccionRecurrente And mRequiereFactoresAdicionales Then
        Dim mensajeFactores As String
        mensajeFactores = "CONFIRMAR FACTORES ADICIONALES:" & vbCrLf & vbCrLf & _
                          "Los siguientes valores se guardarán con la inspección:" & vbCrLf & vbCrLf & _
                          "  • % Recuperación: " & mPorcRecuperacion & " %" & vbCrLf & _
                          "  • % OOL: " & mPorcOOL & " %" & vbCrLf & vbCrLf & _
                          "¿Está seguro que estos valores son correctos?" & vbCrLf & vbCrLf & _
                          "Seleccione 'No' para regresar y corregirlos."
        
        If MsgBox(mensajeFactores, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar factores adicionales") <> vbYes Then
            Exit Sub
        End If
    End If
    
    ' PASO 3: Delegar al orquestador (él se encarga de recopilar, validar respuestas y guardar)
    Call ChecklistOrchestrator.GuardarInspeccionCompleta(Me)
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.btnGuardar_Click", Err.Description, Err.Number)
    MsgBox "Error al guardar: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Función: RequiereCalificaciones
' Propósito: Determina si el puesto actual requiere calificaciones de
'            vestuario y operador.
' Retorna: True si el puesto contiene: Operador, Ayudante, Sanitizador
' FASE 7 - 23/04/2026
' CORREGIDO 27/04/2026: Usar InStr para detectar puestos con variantes (ej: "Operador Electrolitos")
' ----------------------------------------------------------------------
Private Function RequiereCalificaciones() As Boolean
    Dim puestoUpper As String
    puestoUpper = UCase(Trim(mPuesto))
    
    ' Lista de puestos que requieren calificaciones (usando InStr para detectar variantes)
    RequiereCalificaciones = (InStr(1, puestoUpper, "OPERADOR") > 0 Or _
                             InStr(1, puestoUpper, "AYUDANTE") > 0 Or _
                             InStr(1, puestoUpper, "SANITIZADOR") > 0)
End Function

'' ----------------------------------------------------------------------
' Subrutina: ConfigurarVisibilidadCalificaciones
' Propósito: Muestra u oculta los controles de calificaciones según el puesto.
'            - Operadores: Vestuario + Operador (ambos)
'            - Ayudantes: Solo Vestuario
'            - Sanitizador: Solo Vestuario
' FASE 7 - 23/04/2026
' CORREGIDO 27/04/2026: Diferenciar entre campos de vestuario y operador según tipo de puesto
' ----------------------------------------------------------------------
Private Sub ConfigurarVisibilidadCalificaciones()
    On Error Resume Next
    
    Dim puestoUpper As String
    puestoUpper = UCase(Trim(mPuesto))
    
    Dim esOperador As Boolean
    Dim esAyudante As Boolean
    Dim esSanitizador As Boolean
    
    ' Detectar tipo de puesto (usando InStr para incluir variantes como "Operador Electrolitos")
    esOperador = (InStr(1, puestoUpper, "OPERADOR") > 0)
    esAyudante = (InStr(1, puestoUpper, "AYUDANTE") > 0)
    esSanitizador = (InStr(1, puestoUpper, "SANITIZADOR") > 0)
    
    ' Campos de VESTUARIO: Mostrar para Operadores, Ayudantes y Sanitizadores
    Dim mostrarVestuario As Boolean
    mostrarVestuario = (esOperador Or esAyudante Or esSanitizador)
    
    lblCalificacionVestuario.Visible = mostrarVestuario
    cboCalificacionVestuario.Visible = mostrarVestuario
    lblFechaVencVestuario.Visible = mostrarVestuario
    txtFechaVencVestuario.Visible = mostrarVestuario
    
    ' Campos de OPERADOR: Mostrar SOLO para Operadores (NO para Ayudantes ni Sanitizadores)
    Dim mostrarOperador As Boolean
    mostrarOperador = esOperador
    
    lblCalificacionOperador.Visible = mostrarOperador
    cboCalificacionOperador.Visible = mostrarOperador
    lblFechaVencOperador.Visible = mostrarOperador
    txtFechaVencOperador.Visible = mostrarOperador
    
    ' Limpiar valores si no se muestran (para evitar datos residuales)
    If Not mostrarVestuario Then
        cboCalificacionVestuario.Value = ""
        txtFechaVencVestuario.Value = ""
    End If
    
    If Not mostrarOperador Then
        cboCalificacionOperador.Value = ""
        txtFechaVencOperador.Value = ""
    End If
    
    On Error GoTo 0
End Sub

Private Sub btnCancelar_Click()
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Cancelar la inspección? Los datos no se guardarán.", _
                       vbQuestion + vbYesNo, "Confirmar cancelación")
    
    If respuesta = vbYes Then Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Si el usuario cierra con la X, tratar como cancelar
    If CloseMode = vbFormControlMenu Then
        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox("¿Cancelar la inspección? Los datos no se guardarán.", _
                           vbQuestion + vbYesNo, "Confirmar cierre")
        If respuesta <> vbYes Then Cancel = 1
    End If
End Sub

' ======================================================================
' CARGA DE COMBOS
' ======================================================================

Private Sub CargarComboAreas()
    On Error GoTo ErrorHandler
    
    cboArea.Clear
    
    Dim areas As Collection
    Set areas = ChecklistRepository.ObtenerAreasPorPlanta(mPlanta)
    
    Dim a As Variant
    For Each a In areas
        cboArea.AddItem CStr(a)
    Next a
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.CargarComboAreas", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Subrutina: AplicarFiltroArea
' Propósito: Aplica la lógica de filtrado/bloqueo del combo de Área
'            según la planta y el área de la plantilla.
' Lógica:
'   - Therapia iv Santiago + NPT → Filtrar solo opciones NPT (dejar elegir)
'   - Todos los demás casos → Bloquear con valor específico
' ----------------------------------------------------------------------
Private Sub AplicarFiltroArea()
    On Error GoTo ErrorHandler
    
    Dim esPlantaSantiago As Boolean
    Dim areaEsNPT As Boolean
    Dim areaEsGenerica As Boolean
    Dim areaTrimmed As String
    
    areaTrimmed = Trim(mAreaPendiente)
    esPlantaSantiago = (Trim(mPlanta) = "Therapia iv Santiago")
    areaEsNPT = (UCase(areaTrimmed) = "NPT")
    
    ' Detectar si es plantilla genérica (sin área específica o "TODAS")
    areaEsGenerica = (Len(areaTrimmed) = 0 Or UCase(areaTrimmed) = "TODAS" Or UCase(areaTrimmed) = "GENERAL")
    
    ' Debug.Print "=== AplicarFiltroArea ==="
    ' Debug.Print "Planta: [" & mPlanta & "]"
    ' Debug.Print "Área pendiente: [" & mAreaPendiente & "]"
    ' Debug.Print "Es Santiago: " & esPlantaSantiago
    ' Debug.Print "Es NPT: " & areaEsNPT
    ' Debug.Print "Es Genérica: " & areaEsGenerica
    ' Debug.Print "Items en combo ANTES: " & cboArea.ListCount
    
    If areaEsGenerica Then
        ' PLANTILLA GENÉRICA: Mostrar todas las opciones disponibles
        ' Debug.Print "PLANTILLA GENÉRICA: Mostrando todas las áreas..."
        cboArea.Enabled = True
        cboArea.Locked = False
        ' El usuario deberá elegir entre todas las opciones cargadas
        
    ElseIf esPlantaSantiago And areaEsNPT Then
        ' CASO ESPECIAL: Santiago + NPT → Filtrar solo opciones NPT
        ' Guardar todas las opciones actuales
        Dim opcionesOriginales As New Collection
        Dim i As Long
        For i = 0 To cboArea.ListCount - 1
            opcionesOriginales.Add cboArea.List(i)
            ' Debug.Print "  Opción original " & i & ": [" & cboArea.List(i) & "]"
        Next i
        
        ' Debug.Print "Filtrando opciones con 'NPT'..."
        
        ' Limpiar y recargar solo las que contienen "NPT"
        cboArea.Clear
        Dim opcion As Variant
        Dim contador As Long
        contador = 0
        For Each opcion In opcionesOriginales
            If InStr(1, CStr(opcion), "NPT", vbTextCompare) > 0 Then
                cboArea.AddItem CStr(opcion)
                contador = contador + 1
                ' Debug.Print "  ✓ Agregada: [" & CStr(opcion) & "]"
            Else
                ' Debug.Print "  ✗ Omitida: [" & CStr(opcion) & "]"
            End If
        Next opcion
        
        ' Debug.Print "Items agregados: " & contador
        ' Debug.Print "Items en combo DESPUÉS: " & cboArea.ListCount
        
        cboArea.Enabled = True
        cboArea.Locked = False
        ' No asignamos valor, el usuario deberá elegir
    Else
        ' CASO NORMAL: Asignar valor y bloquear
        ' Debug.Print "CASO NORMAL: Bloqueando..."
        
        Dim valorCombo As String
        If UCase(areaTrimmed) = "ONCO" Then
            valorCombo = "Oncología"
        ElseIf UCase(areaTrimmed) = "NPT" Then
            valorCombo = "NPT"
        Else
            valorCombo = areaTrimmed
        End If
        
        cboArea.Value = valorCombo
        cboArea.Enabled = False
        cboArea.Locked = True
        
        ' Debug.Print "Valor bloqueado: [" & valorCombo & "]"
    End If
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.AplicarFiltroArea", Err.Description, Err.Number)
End Sub

Private Sub CargarCombosPersonal()
    On Error GoTo ErrorHandler
    
    ' Limpiar combos
    cboAY1.Clear
    cboAY2.Clear
    cboOP.Clear
    
    ' Cargar AY1 (Ayudante 1)
    Dim personalAY1 As Collection
    Set personalAY1 = ChecklistRepository.ObtenerPersonalPorPuestoYPlanta("Ayudante 1", mPlanta)
    Dim p As Variant
    For Each p In personalAY1
        cboAY1.AddItem CStr(p)
    Next p
    
    ' Cargar AY2 (Ayudante 2)
    Dim personalAY2 As Collection
    Set personalAY2 = ChecklistRepository.ObtenerPersonalPorPuestoYPlanta("Ayudante 2", mPlanta)
    For Each p In personalAY2
        cboAY2.AddItem CStr(p)
    Next p
    
    ' Cargar OP (Operador)
    Dim personalOP As Collection
    Set personalOP = ChecklistRepository.ObtenerPersonalPorPuestoYPlanta("Operador", mPlanta)
    For Each p In personalOP
        cboOP.AddItem CStr(p)
    Next p
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.CargarCombosPersonal", Err.Description, Err.Number)
End Sub

Private Sub CargarComboEvaluadores()
    On Error GoTo ErrorHandler
    
    ' Debug.Print "=== CargarComboEvaluadores ==="
    
    cboEvaluador.Clear
    
    Dim evaluadores As Collection
    Set evaluadores = ChecklistRepository.ObtenerEvaluadores()
    
    Dim e As Variant
    For Each e In evaluadores
        cboEvaluador.AddItem CStr(e)
    Next e
    
    ' Debug.Print "Total evaluadores cargados en combo: " & cboEvaluador.ListCount
    
    ' Pre-seleccionar evaluador basado en usuario de Windows
    Dim nombreUsuarioWindows As String
    Dim inicialesUsuario As String
    
    ' IMPORTANTE: Usar Application.UserName (nombre completo) NO Environ("USERNAME") (cuenta Windows)
    nombreUsuarioWindows = Application.UserName  ' Ejemplo: "NIEVES CARRERO"
    ' Debug.Print "Usuario Windows (Application.UserName): [" & nombreUsuarioWindows & "]"
    
    If Len(nombreUsuarioWindows) > 0 Then
        ' Buscar iniciales del evaluador por nombre
        inicialesUsuario = ChecklistRepository.ObtenerInicialesEvaluadorPorNombre(nombreUsuarioWindows)
        
        ' Debug.Print "Iniciales obtenidas de búsqueda: [" & inicialesUsuario & "]"
        
        If Len(inicialesUsuario) > 0 Then
            ' Intentar seleccionar en el combo
            Dim i As Long
            Dim encontrado As Boolean
            encontrado = False
            
            ' Debug.Print "Buscando en combo las iniciales: [" & inicialesUsuario & "]"
            For i = 0 To cboEvaluador.ListCount - 1
                ' Debug.Print "  Opción " & i & ": [" & cboEvaluador.List(i) & "] | Match: " & (Trim(cboEvaluador.List(i)) = inicialesUsuario)
                If Trim(cboEvaluador.List(i)) = inicialesUsuario Then
                    cboEvaluador.ListIndex = i
                    encontrado = True
                    ' Debug.Print ">>> Evaluador pre-seleccionado: [" & inicialesUsuario & "] en índice " & i
                    Exit For
                End If
            Next i
            
            If Not encontrado Then
                ' Debug.Print ">>> Las iniciales [" & inicialesUsuario & "] NO se encontraron en el combo"
            End If
        Else
            ' Debug.Print ">>> No se encontró evaluador para usuario Windows: [" & nombreUsuarioWindows & "]"
        End If
    Else
        ' Debug.Print ">>> Environ(USERNAME) está vacío"
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Debug.Print "ERROR en CargarComboEvaluadores: " & Err.Description
    Call ErrorLogger2.Log("frmChecklistVirtual.CargarComboEvaluadores", Err.Description, Err.Number)
End Sub

Private Sub CargarComboLugar()
    On Error GoTo ErrorHandler
    
    cboLugar.Clear
    
    On Error Resume Next
    cboLugar.AddItem Configuration2.LUGAR_DENTRO_AREA
    cboLugar.AddItem Configuration2.LUGAR_FUERA_AREA
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    ' Debug.Print "  >> ERROR CargarComboLugar: " & Err.Description
    Call ErrorLogger2.Log("frmChecklistVirtual.CargarComboLugar", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Subrutina: CargarCombosCalificacion
' Propósito: Carga los combos de calificación de vestuario y operador
'            con valores "Si" y "No", estableciendo "Si" como predeterminado.
' Fecha: 23/04/2026 - FASE 7
' ----------------------------------------------------------------------
Private Sub CargarCombosCalificacion()
    On Error GoTo ErrorHandler
    
    ' Combo Calificación de Vestuario
    cboCalificacionVestuario.Clear
    cboCalificacionVestuario.AddItem "Si"
    cboCalificacionVestuario.AddItem "No"
    cboCalificacionVestuario.Value = "Si"  ' Valor predeterminado
    
    ' Combo Calificación de Operador
    cboCalificacionOperador.Clear
    cboCalificacionOperador.AddItem "Si"
    cboCalificacionOperador.AddItem "No"
    cboCalificacionOperador.Value = "Si"  ' Valor predeterminado
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR CargarCombosCalificacion: " & Err.Description
    Call ErrorLogger2.Log("frmChecklistVirtual.CargarCombosCalificacion", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Subrutina: ConfigurarCampoLineaPorPuesto
' Propósito: Configura el campo de Línea/Equipo según el puesto evaluado.
'            Para ciertos puestos técnicos que no tienen línea/equipo asignado,
'            bloquea el campo y establece "N/A".
' ----------------------------------------------------------------------
Private Sub ConfigurarCampoLineaPorPuesto()
    On Error GoTo ErrorHandler
    
    ' Puestos que NO tienen línea/equipo asignada
    Dim puestosSinLinea As Variant
    puestosSinLinea = Array( _
        "Técnico de producción - grado C", _
        "Técnico de producción - grado D" _
    )
    
    Dim puesto As String
    puesto = Trim(mPuesto)
    
    Dim esPuestoSinLinea As Boolean
    esPuestoSinLinea = False
    
    ' Verificar si el puesto actual está en la lista
    Dim i As Long
    For i = LBound(puestosSinLinea) To UBound(puestosSinLinea)
        If StrComp(puesto, puestosSinLinea(i), vbTextCompare) = 0 Then
            esPuestoSinLinea = True
            Exit For
        End If
    Next i
    
    ' Configurar el campo según el resultado
    If esPuestoSinLinea Then
        ' Bloquear y establecer Configuration2.VALOR_NO_APLICA
        cboLineaAuditada.Clear
        cboLineaAuditada.AddItem Configuration2.VALOR_NO_APLICA
        cboLineaAuditada.Value = Configuration2.VALOR_NO_APLICA
        cboLineaAuditada.Enabled = False
        cboLineaAuditada.BackColor = &H8000000F  ' Gris (bloqueado)
    Else
        ' Habilitar campo normal
        cboLineaAuditada.Enabled = True
        cboLineaAuditada.BackColor = &H80000005  ' Blanco (habilitado)
    End If
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.ConfigurarCampoLineaPorPuesto", Err.Description, Err.Number)
End Sub

Private Sub CargarFrecuenciaPlantilla()
    On Error GoTo ErrorHandler
    
    ' Debug.Print "  >> CargarFrecuenciaPlantilla iniciado"
    
    Dim plantillaData As Variant
    On Error Resume Next
    plantillaData = ChecklistRepository.ObtenerPlantillaPorPuesto(mPuesto)
    On Error GoTo ErrorHandler
    
    If Not IsEmpty(plantillaData) Then
        On Error Resume Next
        mFrecuenciaMeses = CLng(plantillaData(2))
        On Error GoTo ErrorHandler
        ' Debug.Print "    - Frecuencia cargada: " & mFrecuenciaMeses & " meses"
    Else
        ' Debug.Print "    - ADVERTENCIA: Plantilla no encontrada, usando default (3 meses)"
        mFrecuenciaMeses = 3
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Debug.Print "  >> ERROR CargarFrecuenciaPlantilla: " & Err.Description
    mFrecuenciaMeses = 3
    Call ErrorLogger2.Log("frmChecklistVirtual.CargarFrecuenciaPlantilla", Err.Description, Err.Number)
End Sub

' ======================================================================
' CARGA DINÁMICA DE PREGUNTAS
' ======================================================================

Private Sub CargarSecciones()
    On Error GoTo ErrorHandler
    
    ' Debug.Print "  >> CargarSecciones iniciado"
    
    Set mSecciones = ChecklistRepository.ObtenerSecciones()
    ' Debug.Print "  >> Secciones obtenidas: " & mSecciones.Count
    
    If mSecciones.Count = 0 Then
        ' Debug.Print "  >> ADVERTENCIA: No hay secciones configuradas"
        Exit Sub
    End If
    
    ' Identificar ID de sección TA
    Dim sec As Variant
    For Each sec In mSecciones
        Dim arrSec() As Variant
        arrSec = sec
        
        Dim nombreSeccion As String
        nombreSeccion = CStr(arrSec(1))
        ' Debug.Print "  >> Sección: " & arrSec(0) & " - " & nombreSeccion
        
        If InStr(1, LCase(nombreSeccion), "aséptica") > 0 Or _
           InStr(1, LCase(nombreSeccion), "aseptica") > 0 Or _
           InStr(1, LCase(nombreSeccion), "técnica") > 0 Then
            mIDSeccionTA = CStr(arrSec(0))
            ' Debug.Print "  >> Identificada sección TA: " & mIDSeccionTA
        ElseIf InStr(1, LCase(nombreSeccion), "procesos") > 0 Or _
               InStr(1, LCase(nombreSeccion), "auditoría de procesos") > 0 Or _
               InStr(1, LCase(nombreSeccion), "auditoria de procesos") > 0 Then
            mIDSeccionProcesos = CStr(arrSec(0))
            ' Debug.Print "  >> Identificada sección Auditoría de Procesos: " & mIDSeccionProcesos
        End If
    Next sec
    
    Exit Sub
    
ErrorHandler:
    ' Debug.Print "  >> ERROR CargarSecciones: " & Err.Description
    Set mSecciones = New Collection
    Call ErrorLogger2.Log("frmChecklistVirtual.CargarSecciones", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Subrutina: CargarPreguntasDinamicas
' Propósito: Crea controles dinámicos (Label + ComboBox + TextBox) dentro
'            de cada página del MultiPage, una página por sección.
' ----------------------------------------------------------------------
Private Sub CargarPreguntasDinamicas()
    On Error GoTo ErrorHandler
    
    ' Validar precondiciones
    If mSecciones Is Nothing Then Exit Sub
    If mSecciones.Count = 0 Then Exit Sub
    
    Dim pageIndex As Long
    Dim arrSec() As Variant
    Dim idSeccion As String
    Dim nombreSeccion As String
    Dim preguntas As Collection
    Dim opciones As Collection
    Dim fraContainer As MSForms.Frame
    Dim fraName As String
    
    pageIndex = 0
    
    ' Máximo de páginas disponibles en el diseñador (fraPreguntas0, fraPreguntas1)
    Dim maxPages As Long
    maxPages = mpPreguntas.Pages.Count  ' Debería ser 2
    
    Dim sec As Variant
    For Each sec In mSecciones
        ' --- GUARDIA: No procesar más secciones que páginas disponibles ---
        If pageIndex >= maxPages Then GoTo SiguienteSeccion
        
        ' --- Extraer datos de la sección ---
        arrSec = sec
        idSeccion = CStr(arrSec(0))
        nombreSeccion = CStr(arrSec(1))
        
        ' --- Verificar que la página existe en el MultiPage ---
        If pageIndex >= mpPreguntas.Pages.Count Then GoTo SiguienteSeccion
        
        ' Configurar caption de la página
        mpPreguntas.Pages(pageIndex).Caption = nombreSeccion
        
        ' --- Obtener preguntas ---
        Set preguntas = Nothing
        
        On Error Resume Next
        Set preguntas = ChecklistRepository.ObtenerPreguntasPorPlantillaYSeccion(mIDPlantilla, idSeccion)
        On Error GoTo ErrorHandler
        
        If preguntas Is Nothing Then GoTo SiguienteSeccion
        
        ' --- Acceder al frame existente en el diseñador ---
        Set fraContainer = Nothing
        fraName = "fraPreguntas" & pageIndex
        
        On Error Resume Next
        Set fraContainer = mpPreguntas.Pages(pageIndex).Controls(fraName)
        On Error GoTo ErrorHandler
        
        If fraContainer Is Nothing Then GoTo SiguienteSeccion
        
        ' --- Activar la página antes de crear controles dinámicos ---
        mpPreguntas.Value = pageIndex
        DoEvents
        
        ' --- Crear controles dinámicos ---
        Call CrearControlesPreguntas(fraContainer, preguntas, idSeccion)
        
SiguienteSeccion:
        pageIndex = pageIndex + 1
    Next sec
    
    ' Volver a primera página
    On Error Resume Next
    DoEvents
    mpPreguntas.Value = 0
    DoEvents
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.CargarPreguntasDinamicas", "pageIndex=" & pageIndex & " | " & Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Función auxiliar: SanitizarNombreControl
' Propósito: Convierte un string en un nombre válido para controles VBA
'            (solo A-Z, 0-9, _) y añade prefijo si comienza con número
' Parámetro: nombre original (puede tener guiones, puntos, etc)
' Retorna: nombre válido para control VBA
' ----------------------------------------------------------------------
Private Function SanitizarNombreControl(ByVal original As String) As String
    Dim sanitized As String
    Dim i As Long
    Dim c As String
    
    sanitized = ""
    For i = 1 To Len(original)
        c = Mid(original, i, 1)
        ' Solo mantener A-Z, a-z, 0-9, guion bajo
        If ((c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or _
            (c >= "0" And c <= "9") Or c = "_") Then
            sanitized = sanitized & c
        Else
            sanitized = sanitized & "_"  ' Reemplazar caracteres inválidos con
        End If
    Next i
    
    ' Si empieza con número, agregar prefijo
    If Len(sanitized) > 0 Then
        c = Left(sanitized, 1)
        If c >= "0" And c <= "9" Then
            sanitized = "P_" & sanitized
        End If
    End If
    
    SanitizarNombreControl = sanitized
End Function

'' ----------------------------------------------------------------------
' Subrutina: CrearControlesPreguntas
' Propósito: Crea Label + ComboBox + TextBox por cada pregunta dentro
'            de un Frame contenedor.
' Parámetros:
'   fra: Frame contenedor
'   preguntas: Collection de arrays (ID, Numero, Texto, IDSeccion, IDCriticidad, Orden)
'   idSeccion: ID de la sección para vincular respuestas
' CAMBIOS: Ahora obtiene opciones por cada pregunta según su ID Criticidad
' ----------------------------------------------------------------------
Private Sub CrearControlesPreguntas(ByRef fra As MSForms.Frame, _
                                    ByVal preguntas As Collection, _
                                    ByVal idSeccion As String)
    On Error GoTo ErrorHandler
    
    ' Validar parámetros
    If fra Is Nothing Then Exit Sub
    If preguntas.Count = 0 Then Exit Sub
    
    Dim topPos As Single
    topPos = PREG_MARGIN
    
    Dim pregIndex As Long
    Dim preg As Variant
    Dim lblPreg As MSForms.Label
    Dim cboResp As MSForms.ComboBox
    Dim lblObs As MSForms.Label
    Dim txtObs As MSForms.TextBox
    Dim dictResp As Object
    
    For Each preg In preguntas
        Dim arrPreg() As Variant
        arrPreg = preg
        
        Dim idPregunta As String
        Dim numPregunta As String
        Dim textoPregunta As String
        Dim idCriticidad As String
        Dim nomSanitizado As String
        
        idPregunta = CStr(arrPreg(0))
        numPregunta = CStr(arrPreg(1))
        textoPregunta = CStr(arrPreg(2))
        ' arrPreg(3) es idSeccion (ya lo tenemos como parámetro)
        idCriticidad = CStr(arrPreg(4))  ' ID Criticidad de esta pregunta
        nomSanitizado = SanitizarNombreControl(idPregunta)
        
        pregIndex = pregIndex + 1
        
        ' --- OBTENER OPCIONES PARA ESTA PREGUNTA (filtradas por sección Y criticidad) ---
        Dim opciones As Collection
        On Error Resume Next
        Set opciones = ChecklistRepository.ObtenerOpcionesRespuesta(idSeccion, idCriticidad)
        On Error GoTo ErrorHandler
        
        If opciones Is Nothing Then GoTo SiguientePregunta
        
        ' --- Label: número + texto de la pregunta ---
        On Error Resume Next
        Set lblPreg = Nothing
        Set lblPreg = fra.Controls.Add("Forms.Label.1", "lblP_" & nomSanitizado)
        
        If Not lblPreg Is Nothing Then
            On Error GoTo ErrorHandler
            With lblPreg
                .Left = PREG_MARGIN
                .Top = topPos
                .Width = fra.InsideWidth - (PREG_MARGIN * 2) - 16
                .Height = PREG_LABEL_HEIGHT
                .Caption = numPregunta & ". " & textoPregunta
                .WordWrap = True
                .Font.Size = 7
                .Tag = idPregunta  ' Guardar ID original en Tag
            End With
            topPos = topPos + PREG_LABEL_HEIGHT + 2
        Else
            On Error GoTo ErrorHandler
            GoTo SiguientePregunta
        End If
        
        ' --- ComboBox: opciones de respuesta ---
        On Error Resume Next
        Set cboResp = Nothing
        Set cboResp = fra.Controls.Add("Forms.ComboBox.1", "cboR_" & nomSanitizado)
        
        If Not cboResp Is Nothing Then
            On Error GoTo ErrorHandler
            With cboResp
                .Left = PREG_MARGIN
                .Top = topPos
                .Width = 160
                .Height = PREG_COMBO_HEIGHT
                .Style = fmStyleDropDownList
                .Tag = idPregunta & "|" & idSeccion & "|" & idCriticidad  ' Guardar IDs en Tag
            End With
            
            ' Cargar opciones (ya filtradas por sección y criticidad)
            Dim op As Variant
            For Each op In opciones
                Dim arrOp() As Variant
                arrOp = op
                cboResp.AddItem CStr(arrOp(1))
            Next op
        Else
            On Error GoTo ErrorHandler
            GoTo SiguientePregunta
        End If
        
        ' --- Label: "Obs:" ---
        On Error Resume Next
        Set lblObs = Nothing
        Set lblObs = fra.Controls.Add("Forms.Label.1", "lblO_" & nomSanitizado)
        
        If Not lblObs Is Nothing Then
            On Error GoTo ErrorHandler
            With lblObs
                .Left = 180
                .Top = topPos + 2
                .Width = 30
                .Height = PREG_OBS_HEIGHT
                .Caption = "Obs:"
                .Font.Size = 8
                .Tag = idPregunta
            End With
        End If
        
        On Error GoTo ErrorHandler
        
        ' --- TextBox: observación individual ---
        On Error Resume Next
        Set txtObs = Nothing
        Set txtObs = fra.Controls.Add("Forms.TextBox.1", "txtO_" & nomSanitizado)
        
        If Not txtObs Is Nothing Then
            On Error GoTo ErrorHandler
            With txtObs
                .Left = 212
                .Top = topPos
                .Width = fra.InsideWidth - 212 - PREG_MARGIN - 16
                .Height = PREG_OBS_HEIGHT
                .Font.Size = 8
                .Tag = idPregunta
            End With
        End If
        
        On Error GoTo ErrorHandler
        
        topPos = topPos + PREG_COMBO_HEIGHT + PREG_MARGIN + 4
        
        ' Registrar pregunta en el diccionario
        On Error Resume Next
        mPreguntaSecciones(idPregunta) = idSeccion
        
        ' Inicializar entrada en diccionario de respuestas
        Set dictResp = CreateObject("Scripting.Dictionary")
        dictResp("IDOpcion") = ""
        dictResp("ValorNumerico") = 0
        dictResp("Observacion") = ""
        dictResp("IDSeccion") = idSeccion
        dictResp("IDCriticidad") = idCriticidad
        mRespuestas(idPregunta) = dictResp
        
        ' Debug.Print "[frmChecklistVirtual.CrearControlesPreguntas] Pregunta " & idPregunta & " inicializada con IDCriticidad: " & idCriticidad
        
        On Error GoTo ErrorHandler

SiguientePregunta:
    Next preg
    
    ' Ajustar ScrollHeight del frame (con límite de seguridad)
    On Error Resume Next
    Dim newHeight As Single
    newHeight = topPos + PREG_MARGIN
    
    ' Límite de seguridad: no exceder 5000
    If newHeight > 5000 Then newHeight = 5000
    
    fra.ScrollHeight = newHeight
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.CrearControlesPreguntas", Err.Description, Err.Number)
End Sub

' ======================================================================
' CAPTURA DE RESPUESTAS
' ======================================================================

'' ----------------------------------------------------------------------
' Subrutina: RecopilarRespuestas
' Propósito: Recorre todos los ComboBox dinámicos de respuestas y
'            actualiza el diccionario mRespuestas con las selecciones.
'            Llamado antes de guardar.
' CAMBIO: Ahora usa ID Criticidad del Tag para obtener opciones correctas
' ----------------------------------------------------------------------
Private Sub RecopilarRespuestas()
    On Error GoTo ErrorHandler
    
    Dim pageIdx As Long
    Dim secIdx As Long
    secIdx = 0
    
    Dim sec As Variant
    For Each sec In mSecciones
        ' Recorrer controles del frame
        Dim fraName As String
        fraName = "fraPreguntas" & secIdx
        
        Dim fraContainer As MSForms.Frame
        On Error Resume Next
        Set fraContainer = mpPreguntas.Pages(secIdx).Controls(fraName)
        On Error GoTo ErrorHandler
        
        If Not fraContainer Is Nothing Then
            Dim respuestasSeccion As Long
            respuestasSeccion = 0
            Dim cbosEncontrados As Long
            cbosEncontrados = 0
            
            Dim ctrl As MSForms.Control
            For Each ctrl In fraContainer.Controls
                ' Buscar ComboBoxes de respuesta (prefijo "cboR_")
                If Left(ctrl.Name, 5) = "cboR_" Then
                    cbosEncontrados = cbosEncontrados + 1
                    
                    Dim cboResp As MSForms.ComboBox
                    Set cboResp = ctrl
                    
                    ' Extraer IDPregunta, IDSeccion e IDCriticidad del Tag
                    Dim tagParts() As String
                    tagParts = Split(cboResp.Tag, "|")
                    Dim idPregunta As String
                    Dim idSeccion As String
                    Dim idCriticidad As String
                    
                    idPregunta = tagParts(0)
                    If UBound(tagParts) >= 1 Then idSeccion = tagParts(1)
                    If UBound(tagParts) >= 2 Then idCriticidad = tagParts(2)
                    
                    If cboResp.ListIndex >= 0 Then
                        respuestasSeccion = respuestasSeccion + 1
                        
                        ' Obtener opciones filtradas por sección Y criticidad
                        Dim opciones As Collection
                        On Error Resume Next
                        Set opciones = ChecklistRepository.ObtenerOpcionesRespuesta(idSeccion, idCriticidad)
                        On Error GoTo ErrorHandler
                        
                        If Not opciones Is Nothing Then
                            ' Encontrar la opción seleccionada
                            Dim textoSeleccionado As String
                            textoSeleccionado = cboResp.Value
                            
                            ' Buscar ID y valor en las opciones
                            Dim op As Variant
                            For Each op In opciones
                                Dim arrOp() As Variant
                                arrOp = op
                                If CStr(arrOp(1)) = textoSeleccionado Then
                                    Dim dictResp As Object
                                    Set dictResp = CreateObject("Scripting.Dictionary")
                                    dictResp("IDOpcion") = CStr(arrOp(0))
                                    dictResp("ValorNumerico") = CDbl(arrOp(2))
                                    dictResp("Observacion") = ""
                                    dictResp("IDSeccion") = idSeccion
                                    dictResp("IDCriticidad") = idCriticidad
                                    Set mRespuestas(idPregunta) = dictResp
                                    
                                    Exit For
                                End If
                            Next op
                        End If
                    End If
                End If
            Next ctrl
        End If  ' Cierre de If Not fraContainer Is Nothing
        
        secIdx = secIdx + 1
    Next sec
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.RecopilarRespuestas", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Subrutina: RecopilarObservaciones
' Propósito: Recorre todos los TextBox de observación y los añade al
'            diccionario de respuestas. Combina con RecopilarRespuestas.
' ----------------------------------------------------------------------
Private Sub RecopilarObservaciones()
    On Error GoTo ErrorHandler
    
    ' Primero recopilar las respuestas de los ComboBox
    Call RecopilarRespuestas
    
    ' Luego las observaciones de los TextBox "txtO_"
    Dim secIdx As Long
    secIdx = 0
    
    Dim sec As Variant
    For Each sec In mSecciones
        Dim fraName As String
        fraName = "fraPreguntas" & secIdx
        
        Dim fraContainer As MSForms.Frame
        On Error Resume Next
        Set fraContainer = mpPreguntas.Pages(secIdx).Controls(fraName)
        On Error GoTo ErrorHandler
        
        If Not fraContainer Is Nothing Then
            Dim ctrl As MSForms.Control
            For Each ctrl In fraContainer.Controls
                If Left(ctrl.Name, 5) = "txtO_" Then
                    Dim txtObs As MSForms.TextBox
                    Set txtObs = ctrl
                    
                    Dim idPregunta As String
                    idPregunta = txtObs.Tag
                    
                    If mRespuestas.Exists(idPregunta) Then
                        Dim dictResp As Object
                        Set dictResp = mRespuestas(idPregunta)
                        dictResp("Observacion") = Trim(txtObs.Value)
                        Set mRespuestas(idPregunta) = dictResp
                    End If
                End If
            Next ctrl
        End If
        
        secIdx = secIdx + 1
    Next sec
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.RecopilarObservaciones", Err.Description, Err.Number)
End Sub

' ======================================================================
' EVENTOS: INSPECCIONES RECURRENTES (FASE 2 - 21/04/2026)
' ======================================================================

' ----------------------------------------------------------------------
' btnToggleRecurrente_Click
' Propósito: Evento disparado cuando el USUARIO hace clic en el botón toggle
'            Alterna entre mostrar/ocultar el frame de inspección recurrente
'            Actualiza el caption del botón según el estado
'            IMPORTANTE: Actualiza mEsInspeccionRecurrente para mantener
'            consistencia con el resto del sistema
' ----------------------------------------------------------------------
Private Sub btnToggleRecurrente_Click()
    On Error GoTo ErrorHandler
    
    If mEsInspeccionRecurrente Then
        ' ===== DESACTIVAR MODO RECURRENTE =====
        mEsInspeccionRecurrente = False
        mModoRecurrenteManual = False
        fraRecurrentInspection.Visible = False
        btnToggleRecurrente.Caption = "Activar Modo Recurrente"
        
        ' Limpiar campos
        txtNumeroInspeccion.Value = ""
        txtRPNAnteriorManual.Value = ""
        txtRPNAnteriorAuto.Value = ""
        txtPorcRecuperacion.Value = ""
        txtPorcOOL.Value = ""
        lblInfoHistorico.Caption = "(Info de inspecciones previas aparecerá aquí)"
        lblInfoHistorico.ForeColor = &H808080  ' Gris
    Else
        ' ===== ACTIVAR MODO RECURRENTE =====
        mEsInspeccionRecurrente = True
        mModoRecurrenteManual = True  ' Activación manual
        
        ' Configurar factores adicionales según tipo de checklist
        Call ConfigurarFactoresAdicionales(mPuesto)
        
        fraRecurrentInspection.Visible = True
        fraRecurrentInspection.ZOrder 1  ' Frame al fondo
        btnToggleRecurrente.ZOrder 0      ' Botón al frente
        btnToggleRecurrente.Caption = "Desactivar Modo Recurrente"
        
        ' Mostrar controles básicos
        lblNumeroInspeccion.Visible = True
        txtNumeroInspeccion.Visible = True
        lblRPNAnterior.Visible = True
        txtRPNAnteriorManual.Visible = True
        lblModoRPN.Visible = True
        
        ' MODO MANUAL: Habilitar edición de número de inspección
        txtNumeroInspeccion.Locked = False
        txtNumeroInspeccion.BackColor = &HFFFFFF  ' Blanco
        txtNumeroInspeccion.Value = ""  ' Vacío para que el usuario ingrese
        
        ' Mostrar textbox manual para RPN
        txtRPNAnteriorManual.Visible = True
        txtRPNAnteriorManual.Locked = False
        txtRPNAnteriorManual.BackColor = &HFFFFFF
        txtRPNAnteriorManual.Value = ""
        txtRPNAnteriorAuto.Visible = False
        
        ' Configurar modo RPN
        mModoRPN = "MANUAL"
        lblModoRPN.Caption = "[Modo MANUAL - Carga de datos históricos]" & vbCrLf & _
                            "Ingrese los valores de la inspección anterior realizada"
        lblModoRPN.ForeColor = &HFF8000  ' Naranja (indica que es manual)
        
        ' Mostrar mensaje informativo
        lblInfoHistorico.Caption = "MODO CARGA MANUAL: Ingrese los datos de la inspección anterior" & vbCrLf & _
                                  "(inspección realizada pero no registrada en el sistema)"
        lblInfoHistorico.ForeColor = &HFF8000  ' Naranja
        
        ' Mostrar factores adicionales si aplica
        If mRequiereFactoresAdicionales Then
            lblPorcRecuperacion.Visible = True
            txtPorcRecuperacion.Visible = True
            txtPorcRecuperacion.Locked = False
            txtPorcRecuperacion.BackColor = &HFFFFFF
            txtPorcRecuperacion.Value = ""
            
            lblPorcOOL.Visible = True
            txtPorcOOL.Visible = True
            txtPorcOOL.Locked = False
            txtPorcOOL.BackColor = &HFFFFFF
            txtPorcOOL.Value = ""
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.btnToggleRecurrente_Click", Err.Description, Err.Number)
End Sub

' ----------------------------------------------------------------------
' BuscarHistorialSilencioso
' Propósito: Búsqueda automática de inspecciones previas al inicializar
'            el formulario. Sin mensajes al usuario, solo logs.
'            Si encuentra historial: activa modo recurrente automáticamente
'            Si NO encuentra: oculta el frame de inspección recurrente
' LLAMADO POR: UserForm_Activate
' ----------------------------------------------------------------------
Private Sub BuscarHistorialSilencioso()
    On Error GoTo ErrorHandler
    
    ' Validar datos mínimos (silencioso, sin mensajes)
    If Len(Trim(mEvaluado)) = 0 Or Len(Trim(Me.txtPuesto.Value)) = 0 Then
        fraRecurrentInspection.Visible = False
        Exit Sub
    End If
    
    ' Buscar inspecciones previas
    ' CRITERIO RECURRENCIA: Solo iniciales + puesto (sin filtrar por área/planta)
    Dim inspecciones As Object
    Dim iniciales As String
    Dim puestoEval As String
    
    iniciales = Trim(mEvaluado)
    puestoEval = Trim(Me.txtPuesto.Value)
    
    Debug.Print "[AUTO-BUSCAR] Criterio: Iniciales='" & iniciales & "' + Puesto='" & puestoEval & "' (sin filtro de plantilla)"
    
    On Error Resume Next
    Set inspecciones = InspectionHistoryService.BuscarInspeccionesPrevias( _
        iniciales, True, puestoEval, "")
    
    ' Verificar si hubo error
    If Err.Number <> 0 Then
        Debug.Print "[AUTO-BUSCAR] ERROR en BuscarInspeccionesPrevias: " & Err.Number & " - " & Err.Description
        Err.Clear
        fraRecurrentInspection.Visible = False
        Exit Sub
    End If
    Err.Clear  ' Limpiar cualquier error residual
    On Error GoTo ErrorHandler
    
    ' Verificar resultado
    If inspecciones Is Nothing Then
        fraRecurrentInspection.Visible = False
        Exit Sub
    End If
    
    ' ===== SI NO HAY HISTORIAL =====
    If inspecciones.Count = 0 Then
        Debug.Print "[AUTO-BUSCAR] Sin historial previo - No se encontraron inspecciones anteriores"
        Debug.Print "[AUTO-BUSCAR] Frame oculto - El usuario podrá activar modo recurrente manualmente si lo desea"
        ' Ocultar frame de inspección recurrente
        fraRecurrentInspection.Visible = False
        btnToggleRecurrente.Caption = "Activar Modo Recurrente"
        Exit Sub
    End If
    
    ' ===== SI HAY HISTORIAL: ACTIVAR MODO RECURRENTE AUTOMÁTICAMENTE =====
    Debug.Print "[AUTO-BUSCAR] Historial encontrado - Activando modo recurrente automáticamente"
    
    ' Obtener última inspección (sin filtro de plantilla)
    Dim ultInsp As Object
    
    On Error Resume Next
    Set ultInsp = InspectionHistoryService.ObtenerUltimaInspeccion( _
        iniciales, True, puestoEval, "")
    
    If Err.Number <> 0 Then
        Debug.Print "[AUTO-BUSCAR] ERROR en ObtenerUltimaInspeccion: " & Err.Number & " - " & Err.Description
        Err.Clear
        fraRecurrentInspection.Visible = False
        Exit Sub
    End If
    Err.Clear  ' Limpiar cualquier error residual
    On Error GoTo ErrorHandler
    
    If ultInsp Is Nothing Then
        fraRecurrentInspection.Visible = False
        Exit Sub
    End If
    
    ' Validar que el campo RPN existe
    On Error Resume Next
    Dim tempRPN As Variant
    tempRPN = ultInsp("RPN")
    If Err.Number <> 0 Then
        Debug.Print "[AUTO-BUSCAR] ERROR: Campo 'RPN' no existe en ultInsp. Error: " & Err.Description
        Debug.Print "[AUTO-BUSCAR] Campos disponibles: " & Join(ultInsp.Keys, ", ")
        Err.Clear
        fraRecurrentInspection.Visible = False
        Exit Sub
    End If
    Err.Clear  ' Limpiar cualquier error residual
    On Error GoTo ErrorHandler
    
    ' Actualizar variable interna
    mEsInspeccionRecurrente = True
    
    ' Configurar factores adicionales según tipo de checklist
    Call ConfigurarFactoresAdicionales(mPuesto)
    
    ' Mostrar frame y controles básicos
    fraRecurrentInspection.Visible = True
    
    ' CRÍTICO: Traer botón al frente para que no sea cubierto por el frame
    btnToggleRecurrente.ZOrder 0
    
    lblNumeroInspeccion.Visible = True
    txtNumeroInspeccion.Visible = True
    lblRPNAnterior.Visible = True
    txtRPNAnteriorAuto.Visible = True
    lblModoRPN.Visible = True
    
    ' Mostrar factores adicionales SOLO si aplica
    If mRequiereFactoresAdicionales Then
        lblPorcRecuperacion.Visible = True
        txtPorcRecuperacion.Visible = True
        lblPorcOOL.Visible = True
        txtPorcOOL.Visible = True
    End If
    
    ' Calcular número de inspección
    ' CORREGIDO (15/06/2026): Usar CalcularNumeroInspeccionSiguiente() que
    ' usa MAX(NumeroInspeccion) en vez de depender de ObtenerUltimaInspeccion.
    ' Esto es más robusto ante desordenamiento por fechas iguales.
    Dim numInspeccion As Long
    numInspeccion = InspectionHistoryService.CalcularNumeroInspeccionSiguiente( _
        iniciales, True, puestoEval, "")
    mNumeroInspeccion = numInspeccion
    txtNumeroInspeccion.Value = CStr(numInspeccion)
    
    ' Cargar RPN anterior (automático)
    Dim rpnAnterior As Double
    rpnAnterior = ultInsp("RPN")
    mRPNAnteriorAuto = rpnAnterior
    mModoRPN = "AUTOMATICO"
    txtRPNAnteriorAuto.Value = Format(rpnAnterior, "0.00")
    lblModoRPN.Caption = "[Modo RPN: AUTOMÁTICO]"
    
    ' Guardar ID de inspección anterior
    mIDInspeccionAnterior = ultInsp("IDInspeccion")
    
    ' Actualizar etiqueta informativa
    Dim fechaInsp As String
    If ultInsp.Exists("FechaInspeccion") Then
        On Error Resume Next
        fechaInsp = Format(CDate(ultInsp("FechaInspeccion")), "dd/mm/yyyy")
        If Err.Number <> 0 Then
            fechaInsp = CStr(ultInsp("FechaInspeccion"))  ' Si falla, mostrar como texto
            Err.Clear
        End If
        Err.Clear  ' Limpiar cualquier error residual
        On Error GoTo ErrorHandler
    Else
        fechaInsp = "(sin fecha)"
    End If
    
    lblInfoHistorico.Caption = "Última: " & fechaInsp & " | RPN: " & Format(rpnAnterior, "0.00")
    lblInfoHistorico.ForeColor = &H8000&  ' Verde
    
    ' Actualizar caption del botón
    btnToggleRecurrente.Caption = "Desactivar Modo Recurrente"
    
    Exit Sub
    
ErrorHandler:
    ' En caso de error, ocultar frame para no confundir al usuario
    On Error Resume Next
    fraRecurrentInspection.Visible = False
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------
' btnBuscarHistorico_Click
' Propósito: Actualiza manualmente el historial de inspecciones
'            (ahora usado como "refrescar" ya que la búsqueda inicial es automática)
' ACTUALIZADO: Fase 3 - Usa InspectionHistoryService
' ----------------------------------------------------------------------
Private Sub btnBuscarHistorico_Click()
    On Error GoTo ErrorHandler
    
    ' Validar que hay personal seleccionado
    If Len(Trim(mEvaluado)) = 0 Then
        MsgBox "Error: No se ha definido el personal evaluado (Iniciales)." & vbCrLf & vbCrLf & _
               "Por favor, regrese al selector y elija el personal a inspeccionar.", _
               vbExclamation, "Datos incompletos"
        Exit Sub
    End If
    
    If Len(Trim(Me.txtPuesto.Value)) = 0 Then
        MsgBox "Error: No se ha definido el puesto del personal." & vbCrLf & vbCrLf & _
               "Por favor, regrese al selector y elija el puesto a inspeccionar.", _
               vbExclamation, "Datos incompletos"
        Exit Sub
    End If
    
    ' Buscar inspecciones previas usando InspectionHistoryService
    ' CRITERIO RECURRENCIA: Solo iniciales + puesto (sin filtrar por área/planta)
    Dim inspecciones As Object
    Dim iniciales As String
    Dim puestoEval As String
    
    iniciales = Trim(mEvaluado)
    puestoEval = Trim(Me.txtPuesto.Value)
    
    Debug.Print "[BUSCAR] Llamando a InspectionHistoryService.BuscarInspeccionesPrevias..."
    Debug.Print "[BUSCAR]   - iniciales: [" & iniciales & "]"
    Debug.Print "[BUSCAR]   - filtroPorPuesto: True"
    Debug.Print "[BUSCAR]   - puesto: [" & puestoEval & "]"
    Debug.Print "[BUSCAR]   - plantillaID: [''] (SIN FILTRO - Recurrencia por persona+puesto solamente)"
    
    On Error Resume Next
    Set inspecciones = InspectionHistoryService.BuscarInspeccionesPrevias( _
        iniciales, True, puestoEval, "")
    
    ' Verificar si hubo error en la llamada
    If Err.Number <> 0 Then
        Dim errNum As Long: errNum = Err.Number
        Dim errDesc As String: errDesc = Err.Description
        Err.Clear  ' Limpiar error antes de continuar
        On Error GoTo 0  ' Desactivar manejo temporal
        
        MsgBox "Error al buscar inspecciones previas:" & vbCrLf & vbCrLf & _
               "Número: " & errNum & vbCrLf & _
               "Descripción: " & errDesc & vbCrLf & vbCrLf & _
               "Verifique que los datos del personal sean correctos.", _
               vbCritical, "Error de búsqueda"
        Exit Sub
    End If
    Err.Clear  ' Limpiar cualquier error residual
    On Error GoTo ErrorHandler
    
    ' Verificar que el resultado no sea Nothing
    If inspecciones Is Nothing Then
        MsgBox "Error: No se pudo realizar la búsqueda de inspecciones previas." & vbCrLf & vbCrLf & _
               "El servicio de historial no respondió correctamente." & vbCrLf & vbCrLf & _
               "Verifique que la tabla tblInspecciones existe y tiene datos.", _
               vbCritical, "Error de sistema"
        Exit Sub
    End If
    
    If inspecciones.Count = 0 Then
        ' No hay inspecciones anteriores
        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox("No se encontraron inspecciones anteriores para:" & vbCrLf & _
               "Personal: " & mEvaluado & vbCrLf & _
               "Puesto: " & Me.txtPuesto.Value & vbCrLf & vbCrLf & _
               "Esta es la PRIMERA inspección según el sistema." & vbCrLf & vbCrLf & _
               "¿Desea activar el modo recurrente para cargar datos de" & vbCrLf & _
               "una inspección anterior realizada pero no registrada?" & vbCrLf & vbCrLf & _
               "(Podrá ingresar manualmente: N° de inspección, RPN anterior, % Recuperación y % OOL)", _
               vbQuestion + vbYesNo, "Cargar datos históricos")
        
        If respuesta = vbYes Then
            ' ===== ACTIVAR MODO RECURRENTE MANUAL =====
            mEsInspeccionRecurrente = True
            mModoRecurrenteManual = True
            mModoRPN = "MANUAL"
            
            ' Configurar factores adicionales según tipo de checklist
            Call ConfigurarFactoresAdicionales(mPuesto)
            
            ' Mostrar frame y controles
            fraRecurrentInspection.Visible = True
            btnToggleRecurrente.ZOrder 0
            
            lblNumeroInspeccion.Visible = True
            txtNumeroInspeccion.Visible = True
            lblRPNAnterior.Visible = True
            lblModoRPN.Visible = True
            
            ' MODO MANUAL: Permitir editar número de inspección
            txtNumeroInspeccion.Locked = False
            txtNumeroInspeccion.BackColor = &HFFFFFF
            txtNumeroInspeccion.Value = "2"  ' Sugerir 2 como valor inicial (esta sería la segunda)
            
            ' Mostrar textbox manual para RPN
            txtRPNAnteriorManual.Visible = True
            txtRPNAnteriorManual.Locked = False
            txtRPNAnteriorManual.BackColor = &HFFFFFF
            txtRPNAnteriorManual.Value = ""
            txtRPNAnteriorAuto.Visible = False
            
            ' Configurar labels informativos
            lblModoRPN.Caption = "[Modo MANUAL - Carga de datos históricos]" & vbCrLf & _
                                "Ingrese los valores de la inspección anterior realizada"
            lblModoRPN.ForeColor = &HFF8000  ' Naranja
            
            lblInfoHistorico.Caption = "MODO CARGA MANUAL: Ingrese los datos de la inspección anterior" & vbCrLf & _
                                      "(inspección realizada pero no registrada en el sistema)"
            lblInfoHistorico.ForeColor = &HFF8000  ' Naranja
            
            ' Mostrar factores adicionales si aplica
            If mRequiereFactoresAdicionales Then
                lblPorcRecuperacion.Visible = True
                txtPorcRecuperacion.Visible = True
                txtPorcRecuperacion.Locked = False
                txtPorcRecuperacion.BackColor = &HFFFFFF
                txtPorcRecuperacion.Value = ""
                
                lblPorcOOL.Visible = True
                txtPorcOOL.Visible = True
                txtPorcOOL.Locked = False
                txtPorcOOL.BackColor = &HFFFFFF
                txtPorcOOL.Value = ""
            End If
            
            ' Actualizar caption del botón
            btnToggleRecurrente.Caption = "Desactivar Modo Recurrente"
            
            ' Mensaje de confirmación
            MsgBox "Modo recurrente activado para carga manual." & vbCrLf & vbCrLf & _
                   "Por favor, complete los siguientes campos:" & vbCrLf & _
                   "• Número de inspección (ej: 2 si esta es la segunda)" & vbCrLf & _
                   "• RPN anterior (0-100)" & vbCrLf & _
                   IIf(mRequiereFactoresAdicionales, "• % Recuperación" & vbCrLf & "• % OOL" & vbCrLf, "") & vbCrLf & _
                   "Estos datos corresponden a la inspección anterior realizada.", _
                   vbInformation, "Modo manual activado"
        Else
            ' El usuario NO quiere cargar datos manuales
            mEsInspeccionRecurrente = False
            mModoRecurrenteManual = False
            mNumeroInspeccion = 1
            mRPNAnteriorManual = 0
            mRPNAnteriorAuto = 0
            mIDInspeccionAnterior = ""
            mModoRPN = "NINGUNO"
            
            ' Ocultar frame completo
            fraRecurrentInspection.Visible = False
            
            ' Actualizar caption del botón
            btnToggleRecurrente.Caption = "Activar Modo Recurrente"
        End If
        
        Exit Sub
    End If
    
    ' Obtener la última inspección (sin filtro de plantilla - criterio: persona+puesto)
    
    Dim ultInsp As Object
    
    On Error Resume Next
    Set ultInsp = InspectionHistoryService.ObtenerUltimaInspeccion( _
        iniciales, True, puestoEval, "")
    
    ' Verificar si hubo error en la llamada
    If Err.Number <> 0 Then
        Dim errNum2 As Long: errNum2 = Err.Number
        Dim errDesc2 As String: errDesc2 = Err.Description
        Err.Clear  ' Limpiar error antes de continuar
        On Error GoTo 0  ' Desactivar manejo temporal
        
        MsgBox "Error al obtener la última inspección:" & vbCrLf & vbCrLf & _
               "Número: " & errNum2 & vbCrLf & _
               "Descripción: " & errDesc2, _
               vbCritical, "Error de búsqueda"
        Exit Sub
    End If
    Err.Clear  ' Limpiar cualquier error residual
    On Error GoTo ErrorHandler
    
    If ultInsp Is Nothing Then
        MsgBox "Error interno: Se encontraron inspecciones pero no se pudo obtener la última." & vbCrLf & vbCrLf & _
               "Por favor, contacte al administrador del sistema.", _
               vbCritical, "Error de sistema"
        Exit Sub
    End If
    
    ' Actualizar variable interna
    mEsInspeccionRecurrente = True
    
    ' Mostrar frame y controles
    fraRecurrentInspection.Visible = True
    
    ' CRÍTICO: Traer botón al frente para que no sea cubierto por el frame
    btnToggleRecurrente.ZOrder 0
    
    lblNumeroInspeccion.Visible = True
    txtNumeroInspeccion.Visible = True
    lblRPNAnterior.Visible = True
    lblModoRPN.Visible = True
    
    ' Determinar número de inspección
    ' CORREGIDO (15/06/2026): Usar CalcularNumeroInspeccionSiguiente() que
    ' usa MAX(NumeroInspeccion) en vez de depender de ObtenerUltimaInspeccion.
    ' Esto es más robusto ante desordenamiento por fechas iguales.
    Dim numInspeccion As Long
    numInspeccion = InspectionHistoryService.CalcularNumeroInspeccionSiguiente( _
        iniciales, True, puestoEval, "")
    mNumeroInspeccion = numInspeccion
    txtNumeroInspeccion.Value = numInspeccion
    
    ' Guardar ID de inspección anterior
    mIDInspeccionAnterior = CStr(ultInsp("IDInspeccion"))
    
    ' Extraer RPN de la inspección anterior
    Dim rpnAnterior As Double
    If ultInsp.Exists("RPN") Then
        rpnAnterior = CDbl(ultInsp("RPN"))
    Else
        rpnAnterior = 0
    End If
    
    ' Cambiar a modo AUTOMÁTICO
    mModoRPN = "AUTO"
    txtRPNAnteriorAuto.Value = Format(rpnAnterior, "0.00")
    txtRPNAnteriorAuto.Visible = True
    txtRPNAnteriorManual.Visible = False
    txtRPNAnteriorManual.Value = ""
    
    lblModoRPN.Caption = "[Modo AUTO - Detectado]"
    lblModoRPN.ForeColor = &H8000&  ' Verde
    
    ' Mostrar información
    Dim fechaInsp As String
    If ultInsp.Exists("FechaInspeccion") Then
        On Error Resume Next
        fechaInsp = Format(CDate(ultInsp("FechaInspeccion")), "dd/mm/yyyy")
        If Err.Number <> 0 Then
            fechaInsp = CStr(ultInsp("FechaInspeccion"))  ' Si falla, mostrar como texto
            Err.Clear
        End If
        Err.Clear  ' Limpiar cualquier error residual
        On Error GoTo ErrorHandler
    Else
        fechaInsp = "(sin fecha)"
    End If
    
    lblInfoHistorico.Caption = "Última: " & ultInsp("IDInspeccion") & " (" & fechaInsp & ") - RPN: " & Format(rpnAnterior, "0.00")
    lblInfoHistorico.ForeColor = &H8000&  ' Verde
    
    ' Actualizar caption del botón
    btnToggleRecurrente.Caption = "Desactivar Modo Recurrente"
    
    MsgBox "Historial actualizado correctamente:" & vbCrLf & vbCrLf & _
           "Inspecciones previas: " & inspecciones.Count & vbCrLf & _
           "Última inspección: " & ultInsp("IDInspeccion") & vbCrLf & _
           "Fecha: " & fechaInsp & vbCrLf & _
           "RPN anterior: " & Format(rpnAnterior, "0.00") & vbCrLf & vbCrLf & _
           "Esta será la inspección #" & numInspeccion & " del puesto.", _
           vbInformation, "Historial actualizado"
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.btnBuscarHistorico_Click", Err.Description, Err.Number)
    
    ' Mensaje de error mejorado
    Dim mensajeError As String
    mensajeError = "Error al buscar historial de inspecciones:" & vbCrLf & vbCrLf
    
    If Err.Number = 0 Then
        mensajeError = mensajeError & "Error inesperado sin código de error." & vbCrLf & vbCrLf & _
                      "Esto puede indicar un problema en el servicio de historial." & vbCrLf & _
                      "Revise la ventana Inmediato (Ctrl+G) para ver los logs detallados."
    Else
        mensajeError = mensajeError & "Número de error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & vbCrLf & _
                      "Datos buscados:" & vbCrLf & _
                      "  - Personal: " & mEvaluado & vbCrLf & _
                      "  - Puesto: " & Me.txtPuesto.Value & vbCrLf & _
                      "  - Plantilla: " & mIDPlantilla
    End If
    
    MsgBox mensajeError, vbCritical, "Error de búsqueda de historial"
End Sub

' ----------------------------------------------------------------------
' ValidarDatosRecurrentes
' Propósito: Valida que los campos de inspección recurrente estén completos
'            antes de guardar la inspección.
' Retorna: True si los datos son válidos, False en caso contrario
' ----------------------------------------------------------------------
Public Function ValidarDatosRecurrentes() As Boolean
    ValidarDatosRecurrentes = True  ' Asumimos válido por defecto
    
    On Error GoTo ErrorHandler
    
    ' Si NO es inspección recurrente, no hay nada que validar
    If Not mEsInspeccionRecurrente Then
        Exit Function
    End If
    
    ' Validar número de inspección
    ' En modo manual, el usuario debe ingresar el número
    If mModoRecurrenteManual Then
        If Len(Trim(txtNumeroInspeccion.Value)) = 0 Then
            MsgBox "Debe ingresar el número de inspección." & vbCrLf & vbCrLf & _
                   "Ejemplo: Si esta es la segunda inspección, ingrese 2.", _
                   vbExclamation, "Validación"
            ValidarDatosRecurrentes = False
            Exit Function
        End If
        
        If Not IsNumeric(txtNumeroInspeccion.Value) Then
            MsgBox "El número de inspección debe ser un valor numérico.", _
                   vbExclamation, "Validación"
            ValidarDatosRecurrentes = False
            Exit Function
        End If
        
        mNumeroInspeccion = CLng(txtNumeroInspeccion.Value)
    End If
    
    If mNumeroInspeccion < 2 Then
        MsgBox "El número de inspección debe ser >= 2 para inspecciones recurrentes." & vbCrLf & vbCrLf & _
               "Si esta es la primera inspección, desactive el modo recurrente.", _
               vbExclamation, "Validación"
        ValidarDatosRecurrentes = False
        Exit Function
    End If
    
    ' Validar que hay RPN anterior (manual o automático)
    Dim tieneRPNManual As Boolean
    Dim tieneRPNAuto As Boolean
    
    tieneRPNManual = (Len(Trim(txtRPNAnteriorManual.Value)) > 0)
    tieneRPNAuto = (Len(Trim(txtRPNAnteriorAuto.Value)) > 0)
    
    If Not tieneRPNManual And Not tieneRPNAuto Then
        MsgBox "Debe proporcionar el RPN anterior (automático o manual) para inspecciones recurrentes.", _
               vbExclamation, "Validación"
        ValidarDatosRecurrentes = False
        Exit Function
    End If
    
    ' Validar formato numérico del RPN manual si se proporcionó
    If tieneRPNManual Then
        If Not IsNumeric(txtRPNAnteriorManual.Value) Then
            MsgBox "El RPN anterior manual debe ser un valor numérico.", _
                   vbExclamation, "Validación"
            ValidarDatosRecurrentes = False
            Exit Function
        End If
        
        Dim rpnVal As Double
        rpnVal = CDbl(txtRPNAnteriorManual.Value)
        
        If rpnVal < 0 Or rpnVal > 100 Then
            MsgBox "El RPN anterior debe estar entre 0 y 100." & vbCrLf & _
                   "(0 = desempeño perfecto, 100 = máximo riesgo)", _
                   vbExclamation, "Validación"
            ValidarDatosRecurrentes = False
            Exit Function
        End If
        
        ' Guardar valor validado
        mRPNAnteriorManual = rpnVal
        mModoRPN = "MANUAL"
    End If
    
    ' ═══════════════════════════════════════════════════════════════════
    ' VALIDAR FACTORES ADICIONALES (FASE 6 - 23/04/2026)
    ' Solo para checklists específicos: Operador, Muestreador, Ayudante 1, Ayudante 2
    ' ═══════════════════════════════════════════════════════════════════
    If mRequiereFactoresAdicionales Then
        ' Validar % Recuperación
        If Len(Trim(txtPorcRecuperacion.Value)) = 0 Then
            MsgBox "Debe proporcionar el % Recuperación para inspecciones recurrentes de este tipo de checklist.", _
                   vbExclamation, "Validación"
            ValidarDatosRecurrentes = False
            Exit Function
        End If
        
        If Not IsNumeric(txtPorcRecuperacion.Value) Then
            MsgBox "El % Recuperación debe ser un valor numérico.", _
                   vbExclamation, "Validación"
            ValidarDatosRecurrentes = False
            Exit Function
        End If
        
        mPorcRecuperacion = CDbl(txtPorcRecuperacion.Value)
        If mPorcRecuperacion < 0 Then
            MsgBox "El % Recuperación no puede ser negativo.", _
                   vbExclamation, "Validación"
            ValidarDatosRecurrentes = False
            Exit Function
        End If
        
        ' Validar % OOL
        If Len(Trim(txtPorcOOL.Value)) = 0 Then
            MsgBox "Debe proporcionar el % OOL para inspecciones recurrentes de este tipo de checklist.", _
                   vbExclamation, "Validación"
            ValidarDatosRecurrentes = False
            Exit Function
        End If
        
        If Not IsNumeric(txtPorcOOL.Value) Then
            MsgBox "El % OOL debe ser un valor numérico.", _
                   vbExclamation, "Validación"
            ValidarDatosRecurrentes = False
            Exit Function
        End If
        
        mPorcOOL = CDbl(txtPorcOOL.Value)
        If mPorcOOL < 0 Then
            MsgBox "El % OOL no puede ser negativo.", _
                   vbExclamation, "Validación"
            ValidarDatosRecurrentes = False
            Exit Function
        End If
    Else
        ' No requiere factores, resetear a 0
        ' Se guardarán como 0 en la BD para evitar columnas vacías
        mPorcRecuperacion = 0
        mPorcOOL = 0
    End If
    
    ' Si llegamos aquí, todo es válido
    ' Debug.Print "Validación inspección recurrente OK:"
    ' Debug.Print "  Número inspección: " & mNumeroInspeccion
    ' Debug.Print "  Modo RPN: " & mModoRPN
    If mModoRPN = "MANUAL" Then
        ' Debug.Print "  RPN anterior (manual): " & mRPNAnteriorManual
    Else
        ' Debug.Print "  RPN anterior (auto): " & mRPNAnteriorAuto
        ' Debug.Print "  ID inspección anterior: " & mIDInspeccionAnterior
    End If
    
    Exit Function
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.ValidarDatosRecurrentes", Err.Description, Err.Number)
    ValidarDatosRecurrentes = False
End Function

' ----------------------------------------------------------------------
' ConfigurarFactoresAdicionales
' Propósito: Determina si requiere factores adicionales según el tipo de
'            checklist y configura las etiquetas con el grado apropiado
' Entrada:  puesto (String) - Puesto del evaluado
' Salida:   Actualiza mRequiereFactoresAdicionales y labels
' ----------------------------------------------------------------------
Private Sub ConfigurarFactoresAdicionales(ByVal puesto As String)
    On Error GoTo ErrorHandler
    
    Dim puestoUpper As String
    puestoUpper = Trim(UCase(puesto))
    
    Dim grado As String
    Dim requiere As Boolean
    
    ' Determinar si requiere factores y qué grado
    If InStr(puestoUpper, "OPERADOR") > 0 Then
        ' Operador → Grado A
        requiere = True
        grado = "Grado A"
        
    ElseIf InStr(puestoUpper, "MUESTREADOR") > 0 Or _
           InStr(puestoUpper, "AYUDANTE 1") > 0 Or _
           InStr(puestoUpper, "AYUDANTE 2") > 0 Then
        ' Muestreador, Ayudante 1, Ayudante 2 → Grado B
        requiere = True
        grado = "Grado B"
        
    ElseIf InStr(puestoUpper, "TÉCNICO") > 0 Or _
           InStr(puestoUpper, "TECNICO") > 0 Or _
           InStr(puestoUpper, "SANITIZADOR") > 0 Then
        ' Técnico Grado C, Técnico Grado D, Sanitizador → NO requiere
        requiere = False
        grado = ""
        
    Else
        ' Tipo desconocido → NO requiere por seguridad
        requiere = False
        grado = ""
        Debug.Print "[FACTORES] Tipo de checklist desconocido: " & puesto & " → NO requiere factores"
    End If
    
    ' Guardar estado
    mRequiereFactoresAdicionales = requiere
    
    ' Configurar labels
    If requiere Then
        lblPorcRecuperacion.Caption = "% Recuperación " & grado & ":"
        lblPorcOOL.Caption = "% OOL " & grado & ":"
        Debug.Print "[FACTORES] Configurado para " & puesto & " → Requiere factores (" & grado & ")"
    Else
        Debug.Print "[FACTORES] Configurado para " & puesto & " → NO requiere factores"
    End If
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.ConfigurarFactoresAdicionales", Err.Description, Err.Number)
    ' En caso de error, asumir que NO requiere factores
    mRequiereFactoresAdicionales = False
End Sub
