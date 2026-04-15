
' ----------------------------------------------------------------------
' Módulo: ThisWorkbook
' Descripción: Controla los eventos globales del libro, la gestión de roles de usuario,
'              la visibilidad y protección de hojas, y la inicialización de la sesión.
'              Centraliza la lógica de seguridad y navegación principal.
' Última actualización: 12/03/2026 - Refactorizado para usar constantes centralizadas
'                       en Configuration2 (mejora de portabilidad entre proyectos).
' ----------------------------------------------------------------------
Option Explicit

Private m_oldValues As Variant ' Almacena valores previos para auditoría (si aplica)

'' ----------------------------------------------------------------------
' Evento: Workbook_Open
' Propósito: Inicializa el libro al abrirlo, estableciendo el rol de usuario,
'            mostrando el nombre en el menú, activando la hoja principal y
'            ocultando/protegiendo el resto de las hojas.
' Lógica:
'   1. Establece el rol por defecto (definido en Configuration2.INITIAL_USER_ROLE).
'   2. Muestra el nombre del usuario en la hoja de menú.
'   3. Activa la hoja principal (definida en Configuration2.MAIN_MENU_SHEET).
'   4. Oculta y protege todas las hojas excepto el menú principal.
'   5. Inicializa rastreo de navegación y audita apertura del libro.
'   6. Maneja errores y registra en el log si ocurre alguno.
'' ----------------------------------------------------------------------
Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    
    ' ========== ACTIVACIÓN DE PROTECCIONES DE SEGURIDAD (URS-22, URS-20) ==========
    ' • URS-22: Proteger estructura del libro (evita eliminar/mover hojas)
    ' • URS-20: Proteger hojas individuales (evita editar celdas específicas)
    ' Controladas por booleanos en Configuration2.bas para fácil debugging
    ' IMPORTANTE: Los booleanos controlan tanto protección como desprotección
    ' ===========================================================================
    
    ' Control de protección de estructura (URS-22)
    If Configuration2.ENABLE_WORKBOOK_PROTECTION Then
        Call WorkbookProtector2.ProtectWorkbook()
        Debug.Print "[INIT] Protección de estructura (URS-22): ACTIVADA"
    Else
        ' Si está en False, desproteger (permite cambios durante desarrollo)
        Call WorkbookProtector2.UnprotectWorkbook()
        Debug.Print "[INIT] Protección de estructura (URS-22): DESACTIVADA (MODO DESARROLLO)"
    End If
    
    ' Control de protección de hojas individuales (URS-20/21)
    ' Si el booleano está en False, desproteger TODAS las hojas
    If Not Configuration2.ENABLE_SHEET_PROTECTION Then
        Call DesprotegerTodasLasHojas()
        Debug.Print "[INIT] Protección de hojas (URS-20/21): DESACTIVADA (MODO DESARROLLO)"
    Else
        Debug.Print "[INIT] Protección de hojas (URS-20/21): ACTIVADA (según rol)"
    End If
    
    ' Inicializar rol de usuario
    m_userRole = Configuration2.INITIAL_USER_ROLE
    
    ' ========== INICIALIZACIÓN AUTOMÁTICA DEL SISTEMA ==========
    ' Detecta automáticamente si es la primera vez y configura:
    ' - Filtro de planta en Menú Principal
    ' - Cronograma de inspecciones (si está vacío)
    ' - Resumen de cronograma
    ' ===========================================================
    Call SystemInitializer.InicializarSistemaCompleto
    
    ' ========== SISTEMA DE NAVEGACIÓN (desactivado hasta completar diseño UI) ==========
    ' Descomentar cuando NavigationService2, SheetService2 y UserManager2 estén listos
    ' ==================================================================================
    
    ' Call SheetService2.HideAndProtectAllSheetsExcept(Configuration2.MAIN_MENU_SHEET)
    ' ThisWorkbook.Sheets(Configuration2.MAIN_MENU_SHEET).Activate
    ' Call UserManager2.DisplayUserName
    ' g_PreviousSheetName = Configuration2.MAIN_MENU_SHEET
    
    ' Dim userName As String
    ' userName = Environ("USERNAME")
    ' Call AuditLogger2.LogAction( _
    '     action:="Apertura del libro", _
    '     sheetName:="Sistema", _
    '     dataModified:="Sesión iniciada", _
    '     beforeChange:="N/A", _
    '     afterChange:="Usuario: " & userName & " | Rol inicial: " & m_userRole, _
    '     moduleAndSubroutine:="ThisWorkbook.Workbook_Open" _
    ' )
    
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("ThisWorkbook.Workbook_Open", VBA.Err.Description, VBA.Err.Number)
    MsgBox "Error al abrir: " & Err.Description, vbCritical
End Sub

'' ----------------------------------------------------------------------
' Subrutina: DesprotegerTodasLasHojas
' Propósito: Desprotege TODAS las hojas del libro cuando ENABLE_SHEET_PROTECTION = False.
'            Esta función se ejecuta en Workbook_Open() para facilitar desarrollo/debugging.
'            Permite pasar rápidamente de PRODUCCIÓN (protegido) a DESARROLLO (libre).
' Lógica:
'   1. Itera sobre todas las hojas del libro
'   2. Para cada hoja, intenta desprotegerla con APP_PASSWORD
'   3. Registra errores pero continúa (fail-safe)
'' ----------------------------------------------------------------------
Private Sub DesprotegerTodasLasHojas()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim hojasSinProteger As Long
    
    hojasSinProteger = 0
    
    For Each ws In ThisWorkbook.Sheets
        ' Intenta desproteger
        On Error Resume Next
        Call SheetProtector2.UnprotectSheet(ws, Configuration2.APP_PASSWORD)
        On Error GoTo ErrorHandler
        
        hojasSinProteger = hojasSinProteger + 1
    Next ws
    
    Debug.Print "[INIT] Desprotegidas " & hojasSinProteger & " hojas para modo desarrollo"
    Exit Sub
    
ErrorHandler:
    ' Si hay error, registra pero continúa
    Debug.Print "[INIT] AVISO: Error desprotegiendo hojas: " & Err.Description
End Sub

'' ----------------------------------------------------------------------
' Función: GetUserRole
' Propósito: Devuelve el rol actual del usuario para controlar permisos en el libro.
'' ----------------------------------------------------------------------
Public Function GetUserRole() As String
    GetUserRole = m_userRole
End Function

'' ----------------------------------------------------------------------
' Evento: Workbook_SheetDeactivate
' Propósito: Evento deshabilitado intencionalmente. La visibilidad de hojas
'            es gestionada exclusivamente por NavigationService2 y eventos
'            de activación para evitar conflictos con la protección de
'            estructura del libro (Error 1004 en Sh.Visible = xlSheetHidden).
' Nota: No implementar cambios de visibilidad aquí.
'' ----------------------------------------------------------------------
Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    ' La visibilidad de las hojas es gestionada exclusivamente por NavigationService2.
    ' No se realizan cambios de visibilidad aquí para evitar conflictos con la
    ' protección de estructura del libro (Error 1004 en Sh.Visible = xlSheetHidden).
End Sub

'' ----------------------------------------------------------------------
' Evento: Workbook_SheetActivate
' Propósito: Asegura que hojas normales activadas estén visibles (excepto Audit Trail).
'            Audita la navegación entre hojas para trazabilidad.
' Lógica:
'   1. Hace visible la hoja SOLO si no es Audit Trail ni la hoja principal.
'   2. Registra el cambio de navegación en el Audit Trail (sin duplicados).
' Nota (10/03/2026): Las hojas Audit Trail NUNCA se hacen visibles automáticamente.
'                     Solo se muestran cuando el usuario navega explícitamente mediante
'                     ShowAuditTrailGroup().
' ESTADO: TEMPORALMENTE DESACTIVADO PARA DESARROLLO
'' ----------------------------------------------------------------------
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' ========== SISTEMA DE NAVEGACIÓN TEMPORALMENTE DESACTIVADO ==========
    ' Para reactivar, descomentar el bloque de código a continuación
    ' =====================================================================
    
    ' If Sh.Name <> Configuration2.MAIN_MENU_SHEET And Not IsAuditSheet(Sh.Name) Then
    '     Sh.Visible = xlSheetVisible
    ' End If
    
    ' If Sh.Name <> g_PreviousSheetName Then
    '     On Error Resume Next
    '     Call AuditLogger2.LogAction( _
    '         action:="Navegación entre hojas", _
    '         sheetName:=Sh.Name, _
    '         dataModified:="Cambio de vista", _
    '         beforeChange:="Hoja anterior: " & g_PreviousSheetName, _
    '         afterChange:="Hoja actual: " & Sh.Name, _
    '         moduleAndSubroutine:="ThisWorkbook.Workbook_SheetActivate" _
    '     )
    '     On Error GoTo 0
    '     g_PreviousSheetName = Sh.Name
    ' End If
    
    ' Evento desactivado - navegación libre sin restricciones
End Sub

'' ----------------------------------------------------------------------
' Función: IsAuditSheet
' Propósito: Determina si una hoja pertenece al grupo Audit Trail.
'            Usado para evitar hacer visibles las hojas Audit Trail automáticamente.
'' ----------------------------------------------------------------------
Private Function IsAuditSheet(ByVal sheetName As String) As Boolean
    Dim i As Long
    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        If sheetName = AuditRotation2.ObtenerNombreHoja(i) Then
            IsAuditSheet = True
            Exit Function
        End If
    Next i
    IsAuditSheet = False
End Function

'' ----------------------------------------------------------------------
' Evento: Workbook_BeforeSave
' Propósito: Crea una copia de seguridad automática del libro antes de cada
'            guardado voluntario del usuario. El backup se genera en la misma
'            carpeta con el nombre "Copia de Seguridad [NombreOriginal].xlsm".
'            Si el archivo aún no tiene ruta (libro nuevo sin guardar),
'            mod_BackupManager omite el backup sin generar error.
' ESTADO: TEMPORALMENTE DESACTIVADO PARA DESARROLLO
'' ----------------------------------------------------------------------
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    ' ========== BACKUP AUTOMÁTICO TEMPORALMENTE DESACTIVADO ==========
    ' Para reactivar, descomentar la siguiente línea
    ' =================================================================
    ' Call mod_BackupManager.CrearBackupAutomatico
    
    Exit Sub
ErrorHandler:
    ' Call ErrorLogger2.Log("ThisWorkbook.Workbook_BeforeSave", VBA.Err.Description, VBA.Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: Workbook_BeforeClose
' Propósito: Ejecuta EjecutarAnalisis si hubo guardados pendientes desde la
'            última ejecución del análisis. Evita que el análisis se ejecute
'            N veces durante una sesión de carga masiva (FASE 4, 22/02/2026).
' Variante manual: el usuario también puede forzar el recálculo desde el
'            botón "Actualizar Análisis" en la hoja de menú (tarea 4.2.4).
' ESTADO: TEMPORALMENTE DESACTIVADO PARA DESARROLLO
'' ----------------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    ' ========== ANÁLISIS AUTOMÁTICO TEMPORALMENTE DESACTIVADO ==========
    ' Para reactivar, descomentar el bloque de código a continuación
    ' ===================================================================
    ' If g_AnalisisPendiente Then
    '     Call dataProcessAnalysis.EjecutarAnalisis
    '     g_AnalisisPendiente = False
    ' End If
    
    Exit Sub
ErrorHandler:
    ' Call ErrorLogger2.Log("ThisWorkbook.Workbook_BeforeClose", VBA.Err.Description, VBA.Err.Number)
End Sub