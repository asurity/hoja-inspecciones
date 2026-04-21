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
'   ┌──────────────────┬───────────────────────────────────┐
'   │ fraCabecera      │                                   │
'   │ (1/3, min 260pt) │     mpPreguntas (2/3)             │
'   │ 13 campos en     │     2 pestañas con preguntas      │
'   │ 1 columna        │     dinámicas                     │
'   ├──────────────────┴───────────────────────────────────┤
'   │ Obs. Generales              [Guardar] [Cancelar]     │
'   └─────────────────────────────────────────────────────-┘
'
' CONTROLES REQUERIDOS EN EL DISEÑADOR VBA:
'   Formulario: frmChecklistVirtual (Adaptativo a pantalla, Min=720x650, StartUpPosition=CenterScreen)
'
'   Frame "fraCabecera" (columna izquierda, 1/3 del ancho, min 260pt):
'     - lblEvaluado, txtEvaluado (TextBox, Locked)
'     - lblPuesto, txtPuesto (TextBox, Locked)
'     - lblPlanta, txtPlanta (TextBox, Locked)
'     - lblArea, cboArea (ComboBox)
'     - lblLineaAuditada, cboLineaAuditada (ComboBox)
'     - lblFecha, txtFecha (TextBox)
'     - lblFechaAuditada, txtFechaAuditada (TextBox)
'     - lblHoraInicio, txtHoraInicio (TextBox)
'     - lblHoraTermino, txtHoraTermino (TextBox)
'     - lblEvaluador, cboEvaluador (ComboBox)
'     - lblAY1, cboAY1 (ComboBox)
'     - lblAY2, cboAY2 (ComboBox)
'     - lblOP, cboOP (ComboBox)
'     - lblLugar, cboLugar (ComboBox)
'
'   MultiPage "mpPreguntas" (columna derecha, 2/3 del ancho, 2 páginas):
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
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

' --- Constantes de diseño ---
Private Const MARGIN As Single = 12
Private Const ROW_HEIGHT As Single = 22
Private Const LABEL_WIDTH As Single = 100
Private Const LEFT_COL_MIN As Single = 260
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
Private Const COLOR_FONDO As Long = &HFAF8F5           ' Beige claro (fondo formulario)
Private Const COLOR_TITULO As Long = &H724E27           ' Marrón oscuro (títulos)
Private Const COLOR_FRAME As Long = &HFFFFFF            ' Blanco (fondo frames/cabecera)
Private Const COLOR_FRAME_PREGUNTAS As Long = &HFAF8F5   ' Beige claro (fondo preguntas)
Private Const COLOR_LABEL As Long = &H5B3A1A            ' Marrón medio (labels de campo)
Private Const COLOR_READONLY As Long = &HF0F0F0         ' Gris claro (no editable)
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

Public Property Get FechaInspeccion() As String
    FechaInspeccion = Trim(txtFecha.Value)
End Property

Public Property Get FechaAuditada() As String
    FechaAuditada = Trim(txtFechaAuditada.Value)
End Property

Public Property Get HoraInicio() As String
    HoraInicio = Trim(txtHoraInicio.Value)
End Property

Public Property Get HoraTermino() As String
    HoraTermino = Trim(txtHoraTermino.Value)
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
    
    ' --- Configurar todos los controles ---
    Call ConfigurarFormulario
    Call ConfigurarCabecera
    Call ConfigurarMultiPage
    Call ConfigurarObservacionGeneral
    Call ConfigurarBotones
    
    Debug.Print "UserForm_Initialize completado OK"
    
    Exit Sub
    
ErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    Debug.Print "ERROR UserForm_Initialize: [" & errNum & "] " & errDesc
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
        
        FORM_WIDTH = scrW_pt * 0.92
        FORM_HEIGHT = scrH_pt * 0.78
    Else
        ' Fallback si la API no responde
        FORM_WIDTH = 720
        FORM_HEIGHT = 650
        scrW_pt = 0
        scrH_pt = 0
    End If
    
    ' Límites razonables
    If FORM_WIDTH < 720 Then FORM_WIDTH = 720
    If FORM_WIDTH > 1200 Then FORM_WIDTH = 1200
    If FORM_HEIGHT < 550 Then FORM_HEIGHT = 550
    If FORM_HEIGHT > 820 Then FORM_HEIGHT = 820
    
    ' Calcular ancho total de contenido
    CONTENT_WIDTH = FORM_WIDTH - (MARGIN * 2) - 4
    
    ' Calcular anchos de columnas (cabecera 1/3, preguntas 2/3)
    mLeftColWidth = CONTENT_WIDTH * 0.33
    If mLeftColWidth < LEFT_COL_MIN Then mLeftColWidth = LEFT_COL_MIN
    If mLeftColWidth > 350 Then mLeftColWidth = 350
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
    ' --- Frame principal de cabecera (columna izquierda, 1 columna vertical) ---
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
        .ScrollBars = fmScrollBarsVertical  ' FASE 2: Scroll para contenido extenso
        .ScrollHeight = 580  ' Altura total del contenido interno (ajustado para 14 filas + frame recurrente)
        .KeepScrollBarsVisible = fmScrollBarsVertical
    End With
    
    ' --- Layout interno: 1 columna ---
    Dim lblLeft As Single: lblLeft = 4
    Dim ctrlLeft As Single: ctrlLeft = 100
    Dim ctrlW As Single: ctrlW = mLeftColWidth - ctrlLeft - 14
    Dim rowTop As Single: rowTop = 20
    Dim rowIdx As Long: rowIdx = 0
    
    If ctrlW < 120 Then ctrlW = 120
    
    ' FILA 1: Evaluado
    With Me.lblEvaluado
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = LABEL_WIDTH
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
        .Width = LABEL_WIDTH
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
        .Width = LABEL_WIDTH
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
        .Width = LABEL_WIDTH
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
        .Width = LABEL_WIDTH
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
        .Font.Size = 9
        .Style = fmStyleDropDownList
        .TabIndex = 1
    End With
    rowIdx = rowIdx + 1
    
    ' FILA 6: Fecha
    With Me.lblFecha
        .Left = lblLeft
        .Top = rowTop + (rowIdx * ROW_HEIGHT)
        .Width = LABEL_WIDTH
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
        .Width = LABEL_WIDTH
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
        .Width = LABEL_WIDTH
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
        .Width = LABEL_WIDTH
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
        .Width = LABEL_WIDTH
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
        .Width = LABEL_WIDTH
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
        .Width = LABEL_WIDTH
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
        .Width = LABEL_WIDTH
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
        .Width = LABEL_WIDTH
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
    ' SECCIÓN: INSPECCIONES RECURRENTES (FASE 2 - 21/04/2026)
    ' ═══════════════════════════════════════════════════════════════════
    
    ' Calcular posición del frame contenedor
    Dim recTop As Single
    recTop = rowTop + (rowIdx * ROW_HEIGHT) + 12  ' Gap de 12pt (aumentado)
    
    ' Frame contenedor (altura aumentada para scroll visible)
    With Me.fraRecurrentInspection
        .Left = lblLeft
        .Top = recTop
        .Width = mLeftColWidth - 8
        .Height = 180  ' Altura optimizada
        .Caption = " Inspección Recurrente "
        .Font.Name = "Segoe UI"
        .Font.Size = 8  ' Fuente más pequeña para ahorrar espacio
        .Font.Bold = True
        .ForeColor = &H5B3A1A  ' Marrón medio
        .BackColor = &HF5F5F5  ' Gris muy claro
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = &HD0C8C0
        .SpecialEffect = fmSpecialEffectFlat
    End With
    
    ' Checkbox principal (mejor posicionado)
    With Me.chkEsRecurrente
        .Left = 8
        .Top = 18
        .Width = 220
        .Height = 18
        .Caption = "Esta NO es la primera inspección"
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = False
        .Value = False
        .TabIndex = 11
    End With
    
    ' Botón búsqueda histórico (mejor posicionado)
    With Me.btnBuscarHistorico
        .Left = 8
        .Top = 40
        .Width = 140
        .Height = 24
        .Caption = "🔍 Buscar historial"
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .TabIndex = 12
        .BackColor = &HE0E0E0
    End With
    
    ' Label info histórico (mejor posicionado con WordWrap)
    With Me.lblInfoHistorico
        .Left = 8
        .Top = 68
        .Width = 220
        .Height = 20
        .Caption = "(Info de inspecciones previas aparecerá aquí)"
        .Font.Name = "Segoe UI"
        .Font.Size = 7
        .ForeColor = &H808080  ' Gris
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
    End With
    
    ' Label número inspección (mejor alineado)
    With Me.lblNumeroInspeccion
        .Left = 8
        .Top = 92
        .Width = 100
        .Height = 18
        .Caption = "Inspección N°:"
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignLeft
        .Visible = False
    End With
    
    ' TextBox número inspección (mejor posicionado)
    With Me.txtNumeroInspeccion
        .Left = 112
        .Top = 92
        .Width = 50
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = True
        .BackColor = COLOR_READONLY
        .TabStop = False
        .Visible = False
    End With
    
    ' Label RPN Anterior (mejor alineado)
    With Me.lblRPNAnterior
        .Left = 8
        .Top = 118
        .Width = 100
        .Height = 18
        .Caption = "RPN anterior:"
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .ForeColor = COLOR_LABEL
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignLeft
        .Visible = False
    End With
    
    ' TextBox RPN Anterior Automático (mejor posicionado)
    With Me.txtRPNAnteriorAuto
        .Left = 112
        .Top = 118
        .Width = 70
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = True
        .BackColor = &HE0FFE0  ' Verde claro
        .TabStop = False
        .Visible = False
    End With
    
    ' TextBox RPN Anterior Manual (mejor posicionado)
    With Me.txtRPNAnteriorManual
        .Left = 112
        .Top = 118
        .Width = 70
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Locked = False
        .BackColor = &HFFFFE0  ' Amarillo claro
        .Visible = False
    End With
    
    ' Label estado modo (mejor posicionado con WordWrap)
    With Me.lblModoRPN
        .Left = 8
        .Top = 144
        .Width = 220
        .Height = 26
        .Caption = "[Modo RPN: no determinado]"
        .Font.Name = "Segoe UI"
        .Font.Size = 7
        .ForeColor = &H808080  ' Gris
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
        .Visible = False
    End With
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
            
            Debug.Print "  Frame " & fraName & " redimensionado: " & pgW & " x " & pgH
        Else
            Debug.Print "  Frame " & fraName & " NO encontrado"
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
    
    Debug.Print "=== UserForm_Activate INICIADO ==="
    Debug.Print "Valores recibidos en UserForm_Activate:"
    Debug.Print "  mEvaluado: [" & mEvaluado & "]"
    Debug.Print "  mPuesto: [" & mPuesto & "]"
    Debug.Print "  mIDPlantilla: [" & mIDPlantilla & "]"
    Debug.Print "  mPlanta: [" & mPlanta & "]"
    Debug.Print "  mIDCronograma: [" & mIDCronograma & "]"
    
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
    txtPuesto.Value = mPuesto
    txtPlanta.Value = mPlanta
    txtFecha.Value = Format(Date, "dd/mm/yyyy")
    txtFechaAuditada.Value = Format(Date, "dd/mm/yyyy")
    On Error GoTo ErrorHandler
    
    Debug.Print "Campos básicos cargados OK"
    
    ' --- Redimensionar frames de preguntas (ahora que el form está renderizado) ---
    Debug.Print "Redimensionando frames de preguntas..."
    Call RedimensionarFramesPreguntas
    
    ' --- Cargar combos con manejo seguro ---
    Debug.Print "Intentando cargar combos..."
    
    On Error Resume Next
    Call CargarComboAreas
    Debug.Print "  Areas: OK"
    Call CargarCombosPersonal
    Debug.Print "  Personal: OK"
    Call CargarComboEvaluadores
    Debug.Print "  Evaluadores: OK"
    Call CargarComboLugar
    Debug.Print "  Lugar: OK"
    Call CargarFrecuenciaPlantilla
    Debug.Print "  Frecuencia: OK"
    On Error GoTo ErrorHandler
    
    Debug.Print "Combos cargados OK"
    
    ' --- Cargar secciones y preguntas (lo más crítico) ---
    Debug.Print "Intentando cargar secciones y preguntas..."
    
    Call CargarSecciones
    Debug.Print "Secciones cargadas: " & mSecciones.Count
    
    If mSecciones.Count > 0 Then
        Call CargarPreguntasDinamicas
        Debug.Print "CargarPreguntasDinamicas completado"
    Else
        Debug.Print "ERROR: No hay secciones. Mostrando form sin preguntas."
        MsgBox "No se encontraron secciones configuradas. El formulario mostrará solo la cabecera.", _
               vbExclamation, "Advertencia"
    End If
    
    Exit Sub
    
ErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    Debug.Print "ERROR UserForm_Activate: [" & errNum & "] " & errDesc
    
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
    
    Debug.Print "[frmChecklistVirtual.ObtenerRespuestasConSeccion] Iniciando - Total respuestas en mRespuestas: " & mRespuestas.Count
    
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
        
        Debug.Print "  [ObtenerRespuestasConSeccion] Respuesta #" & contador & " - IDPregunta: " & key & ", IDSeccion: " & dictOrig("IDSeccion") & ", IDCriticidad: " & dictOrig("IDCriticidad") & ", IDOpcion: " & dictOrig("IDOpcion")
        
        resultado.Add dictResp
    Next key
    
    Debug.Print "[frmChecklistVirtual.ObtenerRespuestasConSeccion] Completado - Total items en Collection: " & resultado.Count
    
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
    ObtenerCantidadPreguntas = mRespuestas.Count
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

Private Sub btnGuardar_Click()
    On Error GoTo ErrorHandler
    
    ' FASE 2: Validar datos de inspección recurrente si aplica
    If Not ValidarDatosRecurrentes() Then
        Exit Sub
    End If
    
    ' Delegar al orquestador (él se encarga de recopilar y validar)
    Call ChecklistOrchestrator.GuardarInspeccionCompleta(Me)
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.btnGuardar_Click", Err.Description, Err.Number)
    MsgBox "Error al guardar: " & Err.Description, vbCritical, "Error"
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
    
    cboEvaluador.Clear
    
    Dim evaluadores As Collection
    Set evaluadores = ChecklistRepository.ObtenerEvaluadores()
    
    Dim e As Variant
    For Each e In evaluadores
        cboEvaluador.AddItem CStr(e)
    Next e
    
    Exit Sub
    
ErrorHandler:
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
    Debug.Print "  >> ERROR CargarComboLugar: " & Err.Description
    Call ErrorLogger2.Log("frmChecklistVirtual.CargarComboLugar", Err.Description, Err.Number)
End Sub

Private Sub CargarFrecuenciaPlantilla()
    On Error GoTo ErrorHandler
    
    Debug.Print "  >> CargarFrecuenciaPlantilla iniciado"
    
    Dim plantillaData As Variant
    On Error Resume Next
    plantillaData = ChecklistRepository.ObtenerPlantillaPorPuesto(mPuesto)
    On Error GoTo ErrorHandler
    
    If Not IsEmpty(plantillaData) Then
        On Error Resume Next
        mFrecuenciaMeses = CLng(plantillaData(2))
        On Error GoTo ErrorHandler
        Debug.Print "    - Frecuencia cargada: " & mFrecuenciaMeses & " meses"
    Else
        Debug.Print "    - ADVERTENCIA: Plantilla no encontrada, usando default (3 meses)"
        mFrecuenciaMeses = 3
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "  >> ERROR CargarFrecuenciaPlantilla: " & Err.Description
    mFrecuenciaMeses = 3
    Call ErrorLogger2.Log("frmChecklistVirtual.CargarFrecuenciaPlantilla", Err.Description, Err.Number)
End Sub

' ======================================================================
' CARGA DINÁMICA DE PREGUNTAS
' ======================================================================

Private Sub CargarSecciones()
    On Error GoTo ErrorHandler
    
    Debug.Print "  >> CargarSecciones iniciado"
    
    Set mSecciones = ChecklistRepository.ObtenerSecciones()
    Debug.Print "  >> Secciones obtenidas: " & mSecciones.Count
    
    If mSecciones.Count = 0 Then
        Debug.Print "  >> ADVERTENCIA: No hay secciones configuradas"
        Exit Sub
    End If
    
    ' Identificar ID de sección TA
    Dim sec As Variant
    For Each sec In mSecciones
        Dim arrSec() As Variant
        arrSec = sec
        
        Dim nombreSeccion As String
        nombreSeccion = CStr(arrSec(1))
        Debug.Print "  >> Sección: " & arrSec(0) & " - " & nombreSeccion
        
        If InStr(1, LCase(nombreSeccion), "aséptica") > 0 Or _
           InStr(1, LCase(nombreSeccion), "aseptica") > 0 Or _
           InStr(1, LCase(nombreSeccion), "técnica") > 0 Then
            mIDSeccionTA = CStr(arrSec(0))
            Debug.Print "  >> Identificada sección TA: " & mIDSeccionTA
        ElseIf InStr(1, LCase(nombreSeccion), "procesos") > 0 Or _
               InStr(1, LCase(nombreSeccion), "auditoría de procesos") > 0 Or _
               InStr(1, LCase(nombreSeccion), "auditoria de procesos") > 0 Then
            mIDSeccionProcesos = CStr(arrSec(0))
            Debug.Print "  >> Identificada sección Auditoría de Procesos: " & mIDSeccionProcesos
        End If
    Next sec
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "  >> ERROR CargarSecciones: " & Err.Description
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
                .Font.Size = 9
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
        
        Debug.Print "[frmChecklistVirtual.CrearControlesPreguntas] Pregunta " & idPregunta & " inicializada con IDCriticidad: " & idCriticidad
        
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
            Dim ctrl As MSForms.Control
            For Each ctrl In fraContainer.Controls
                ' Buscar ComboBoxes de respuesta (prefijo "cboR_")
                If Left(ctrl.Name, 5) = "cboR_" Then
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
                                    
                                    Debug.Print "[frmChecklistVirtual.RecopilarRespuestas] Pregunta " & idPregunta & ": IDOpcion=" & arrOp(0) & ", ValorNum=" & arrOp(2) & ", IDCriticidad=" & idCriticidad
                                    Exit For
                                End If
                            Next op
                        End If
                    End If
                End If
            Next ctrl
        End If
        
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
' chkEsRecurrente_Click
' Propósito: Muestra/oculta controles de inspección recurrente al marcar el checkbox
' ----------------------------------------------------------------------
Private Sub chkEsRecurrente_Click()
    On Error GoTo ErrorHandler
    
    If chkEsRecurrente.Value = True Then
        ' Mostrar controles de inspección recurrente
        lblNumeroInspeccion.Visible = True
        txtNumeroInspeccion.Visible = True
        lblRPNAnterior.Visible = True
        lblModoRPN.Visible = True
        
        mEsInspeccionRecurrente = True
        
        ' El usuario debe hacer clic en "Buscar historial" o ingresar datos manualmente
        If mNumeroInspeccion <= 1 Then
            txtNumeroInspeccion.Value = "2"  ' Default: segunda inspección
            mNumeroInspeccion = 2
        Else
            txtNumeroInspeccion.Value = CStr(mNumeroInspeccion)
        End If
    Else
        ' Ocultar y resetear controles
        lblNumeroInspeccion.Visible = False
        txtNumeroInspeccion.Visible = False
        lblRPNAnterior.Visible = False
        txtRPNAnteriorAuto.Visible = False
        txtRPNAnteriorManual.Visible = False
        lblModoRPN.Visible = False
        
        mEsInspeccionRecurrente = False
        mNumeroInspeccion = 1
        mRPNAnteriorManual = 0
        mRPNAnteriorAuto = 0
        mIDInspeccionAnterior = ""
        mModoRPN = "NINGUNO"
        
        txtNumeroInspeccion.Value = ""
        txtRPNAnteriorAuto.Value = ""
        txtRPNAnteriorManual.Value = ""
        lblModoRPN.Caption = "[Modo RPN: no determinado]"
    End If
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.chkEsRecurrente_Click", Err.Description, Err.Number)
End Sub

' ----------------------------------------------------------------------
' btnBuscarHistorico_Click
' Propósito: Busca inspecciones previas del personal evaluado
'            y autocompleta los campos de inspección recurrente
' ----------------------------------------------------------------------
Private Sub btnBuscarHistorico_Click()
    On Error GoTo ErrorHandler
    
    ' Validar que hay personal seleccionado
    If Len(mEvaluado) = 0 Or Len(mIDPlantilla) = 0 Then
        MsgBox "Debe seleccionar un personal y plantilla antes de buscar historial.", _
               vbExclamation, "Datos incompletos"
        Exit Sub
    End If
    
    ' TODO FASE 3: Llamar a InspectionHistoryService.BuscarInspeccionesPrevias()
    ' Por ahora, simulación básica
    
    MsgBox "NOTA TEMPORAL (FASE 2):" & vbCrLf & vbCrLf & _
           "La búsqueda de historial se implementará en FASE 3." & vbCrLf & _
           "Por ahora, puede marcar el checkbox e ingresar los datos manualmente:" & vbCrLf & vbCrLf & _
           "1. Marque 'Esta NO es la primera inspección'" & vbCrLf & _
           "2. Ingrese el RPN anterior en el campo manual", _
           vbInformation, "Función en desarrollo - FASE 3"
    
    ' Activar checkbox automáticamente para facilitar el flujo
    chkEsRecurrente.Value = True
    
    ' Habilitar modo manual (ya que no hay búsqueda automática aún)
    txtRPNAnteriorManual.Visible = True
    txtRPNAnteriorAuto.Visible = False
    mModoRPN = "MANUAL"
    lblModoRPN.Caption = "[Modo MANUAL - Ingrese RPN]"
    lblModoRPN.ForeColor = &H0080FF  ' Naranja
    
    ' Simular datos temporales para prueba (esto se reemplazará en FASE 3)
    lblInfoHistorico.Caption = "Búsqueda histórica pendiente (FASE 3)"
    lblInfoHistorico.ForeColor = &H808080
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.btnBuscarHistorico_Click", Err.Description, Err.Number)
    MsgBox "Error al buscar historial: " & Err.Description, vbCritical, "Error"
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
    If mNumeroInspeccion < 2 Then
        MsgBox "El número de inspección debe ser >= 2 para inspecciones recurrentes.", _
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
        
        If rpnVal <= 0 Or rpnVal > 100 Then
            MsgBox "El RPN anterior debe estar entre 0 y 100.", _
                   vbExclamation, "Validación"
            ValidarDatosRecurrentes = False
            Exit Function
        End If
        
        ' Guardar valor validado
        mRPNAnteriorManual = rpnVal
        mModoRPN = "MANUAL"
    End If
    
    ' Si llegamos aquí, todo es válido
    Debug.Print "Validación inspección recurrente OK:"
    Debug.Print "  Número inspección: " & mNumeroInspeccion
    Debug.Print "  Modo RPN: " & mModoRPN
    If mModoRPN = "MANUAL" Then
        Debug.Print "  RPN anterior (manual): " & mRPNAnteriorManual
    Else
        Debug.Print "  RPN anterior (auto): " & mRPNAnteriorAuto
        Debug.Print "  ID inspección anterior: " & mIDInspeccionAnterior
    End If
    
    Exit Function
    
ErrorHandler:
    Call ErrorLogger2.Log("frmChecklistVirtual.ValidarDatosRecurrentes", Err.Description, Err.Number)
    ValidarDatosRecurrentes = False
End Function
