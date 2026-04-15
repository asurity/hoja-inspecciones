' ==============================================================================
' Módulo: AuditRotation2
' Descripción: Sistema de rotación automática de hojas Audit Trail para el
'              proyecto TH-HC-002. Cuando una hoja alcanza AUDIT_MAX_ROWS
'              filas de datos, la siguiente hoja (previamente oculta) se activa
'              y recibe los registros siguientes. Soporta hasta AUDIT_MAX_SHEETS
'              hojas (5 por defecto = hasta 5 000 000 registros en producción).
'
' Convención de nomenclatura (debe coincidir exactamente con el libro):
'   Hoja 1 : "Audit trail 1"  tabla "tblAudit1"
'   Hoja 2 : "Audit trail 2"  tabla "tblAudit2"
'   Hoja 3 : "Audit trail 3"  tabla "tblAudit3"
'   Hoja 4 : "Audit trail 4"  tabla "tblAudit4"
'   Hoja 5 : "Audit trail 5"  tabla "tblAudit5"
'
' Visibilidad inicial en el libro:
'   "Audit trail 1" ? xlSheetVisible       (ya existe y es visible)
'   "Audit trail 2" ? xlSheetVeryHidden    (crear y ocultar)
'   "Audit trail 3" ? xlSheetVeryHidden    (crear y ocultar)
'   "Audit trail 4" ? xlSheetVeryHidden    (crear y ocultar)
'   "Audit trail 5" ? xlSheetVeryHidden    (crear y ocultar)
'
' Dependencias: Configuration2, SheetProtector2, WorkbookProtector2, ErrorLogger2
' Autor: TH-HC-002
' Última modificación: 22/02/2026
' ==============================================================================
Option Explicit

' ------------------------------------------------------------------------------
' Función: ObtenerNombreHoja
' Retorna el nombre de la hoja Audit Trail para el número de secuencia dado.
'   1 ? "Audit trail 1"
'   2 ? "Audit trail 2"
'   5 ? "Audit trail 5"
' ------------------------------------------------------------------------------
Public Function ObtenerNombreHoja(ByVal numeroHoja As Long) As String
    ObtenerNombreHoja = Configuration2.AUDIT_BASE_NAME & " " & numeroHoja
End Function

' ------------------------------------------------------------------------------
' Función: ObtenerNombreTabla
' Retorna el nombre de la tabla ListObject para el número de secuencia dado.
'   1 ? "tblAudit1"
'   2 ? "tblAudit2"
'   5 ? "tblAudit5"
' ------------------------------------------------------------------------------
Public Function ObtenerNombreTabla(ByVal numeroHoja As Long) As String
    ObtenerNombreTabla = Configuration2.AUDIT_TABLE_PREFIX & numeroHoja
End Function

' ------------------------------------------------------------------------------
' Función: ContarFilasTabla
' Cuenta filas de datos en un ListObject. Retorna:
'   >= 0  filas reales
'   -1    tabla no existe o error
' ------------------------------------------------------------------------------
Public Function ContarFilasTabla(ByVal ws As Worksheet, ByVal nombreTabla As String) As Long
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = ws.ListObjects(nombreTabla)
    If tbl Is Nothing Then
        ContarFilasTabla = -1
        Exit Function
    End If
    If tbl.DataBodyRange Is Nothing Then
        ContarFilasTabla = 0
    Else
        ContarFilasTabla = tbl.ListRows.Count
    End If
    On Error GoTo 0
End Function

' ------------------------------------------------------------------------------
' Función: DetectarSiHojaLlena
' Retorna True si la tabla ha alcanzado o superado AUDIT_MAX_ROWS.
' Retorna False si tiene espacio o si no se puede leer (fail-safe).
' ------------------------------------------------------------------------------
Public Function DetectarSiHojaLlena(ByVal ws As Worksheet, ByVal nombreTabla As String) As Boolean
    Dim filas As Long
    filas = ContarFilasTabla(ws, nombreTabla)
    If filas < 0 Then
        ' No se pudo leer la tabla — tratar como NO llena para no bloquear la escritura.
        Debug.Print "[AuditRotation2] AVISO: No se pudo contar filas en '" & ws.Name & "' / '" & nombreTabla & "'. Se asume no llena."
        DetectarSiHojaLlena = False
        Exit Function
    End If
    If filas >= Configuration2.AUDIT_MAX_ROWS Then
        Debug.Print "[AuditRotation2] Hoja LLENA: '" & ws.Name & "' (" & filas & "/" & Configuration2.AUDIT_MAX_ROWS & " filas)"
        DetectarSiHojaLlena = True
    Else
        DetectarSiHojaLlena = False
    End If
End Function

' ------------------------------------------------------------------------------
' Función: ObtenerHojaAuditActiva
' Función principal del módulo. Recorre las hojas 1?AUDIT_MAX_SHEETS y retorna
' la primera que NO está llena. Si una hoja adicional estaba oculta, la hace
' visible en ese momento (rotación automática).
' Si todas están llenas, retorna la última hoja (no se pierde ningún registro).
' Retorna Nothing solo ante un error inesperado.
' ------------------------------------------------------------------------------
Public Function ObtenerHojaAuditActiva() As Worksheet
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim nombreHoja As String
    Dim nombreTabla As String
    Dim ws As Worksheet
    Dim wsLast As Worksheet

    Debug.Print "[AuditRotation2] Buscando hoja Audit Trail activa (máx " & Configuration2.AUDIT_MAX_SHEETS & " hojas, máx " & Configuration2.AUDIT_MAX_ROWS & " filas/hoja)..."

    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        nombreHoja = ObtenerNombreHoja(i)
        nombreTabla = ObtenerNombreTabla(i)

        ' Verificar si la hoja existe en el libro.
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(nombreHoja)
        On Error GoTo ErrorHandler

        If ws Is Nothing Then
            Debug.Print "[AuditRotation2]   Hoja " & i & " ('" & nombreHoja & "'): NO EXISTE en el libro — se omite."
        Else
            Set wsLast = ws   ' Guardar referencia a la última hoja encontrada.

            If DetectarSiHojaLlena(ws, nombreTabla) Then
                Debug.Print "[AuditRotation2]   Hoja " & i & " ('" & nombreHoja & "'): llena — continuando búsqueda."
            Else
                ' Hoja encontrada con espacio disponible.
                ' FIX (10/03/2026): NO hacer visible la hoja automáticamente.
                ' Las hojas Audit Trail solo deben ser visibles cuando el usuario
                ' navega explícitamente mediante ShowAuditTrailGroup().
                ' Excel puede escribir en hojas ocultas sin problemas.
                Debug.Print "[AuditRotation2]   Hoja " & i & " ('" & nombreHoja & "'): seleccionada (permanece oculta hasta navegación explícita)."
                Set ObtenerHojaAuditActiva = ws
                Exit Function
            End If
        End If
    Next i

    ' Todas las hojas encontradas están llenas.
    If Not wsLast Is Nothing Then
        Debug.Print "[AuditRotation2] *** CRÍTICO: todas las hojas Audit Trail están llenas. Se usará '" & wsLast.Name & "' (excederá el límite)."
        Set ObtenerHojaAuditActiva = wsLast
    Else
        Debug.Print "[AuditRotation2] *** ERROR CRÍTICO: no se encontró ninguna hoja Audit Trail en el libro."
        Set ObtenerHojaAuditActiva = Nothing
    End If
    Exit Function

ErrorHandler:
    Debug.Print "[AuditRotation2] ERROR en ObtenerHojaAuditActiva: Nº" & Err.Number & " — " & Err.Description
    Call ErrorLogger2.Log("AuditRotation2.ObtenerHojaAuditActiva", VBA.Err.Description, VBA.Err.Number)
    Set ObtenerHojaAuditActiva = Nothing
End Function

' ==============================================================================
' SECCIÓN DE TESTING
' Las funciones TEST_* son solo para entornos de desarrollo/validación.
' Para activarlas, cambiar Configuration2.AUDIT_MAX_ROWS a 100, ejecutar las
' pruebas y luego revertir AUDIT_MAX_ROWS a 1000000 antes de desplegar.
' ==============================================================================

' ------------------------------------------------------------------------------
' Sub: TEST_GenerarRegistrosAudit
' Genera N registros ficticios para probar la rotación automática.
' Uso típico: TEST_GenerarRegistrosAudit(250) con AUDIT_MAX_ROWS = 100
'             ? distribuye: 100 + 100 + 50 registros en hojas 1, 2 y 3.
' ------------------------------------------------------------------------------
Public Sub TEST_GenerarRegistrosAudit(Optional ByVal cantidadRegistros As Long = 250)
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim t0 As Double
    t0 = Timer

    Application.ScreenUpdating = False

    Debug.Print "[TEST] ========================================================"
    Debug.Print "[TEST] Generando " & cantidadRegistros & " registros de prueba..."
    Debug.Print "[TEST] AUDIT_MAX_ROWS = " & Configuration2.AUDIT_MAX_ROWS & "  |  AUDIT_MAX_SHEETS = " & Configuration2.AUDIT_MAX_SHEETS
    Debug.Print "[TEST] ========================================================"

    For i = 1 To cantidadRegistros
        Call AuditLogger2.LogAction( _
            action:="TEST - Acción " & i, _
            sheetName:="HojaPrueba", _
            dataModified:="Campo" & (i Mod 10), _
            beforeChange:="Antes_" & i, _
            afterChange:="Después_" & i, _
            moduleAndSubroutine:="AuditRotation2.TEST_GenerarRegistrosAudit" _
        )
        If i Mod 50 = 0 Then
            Debug.Print "[TEST]   Generados " & i & " / " & cantidadRegistros & " registros..."
        End If
    Next i

    Application.ScreenUpdating = True
    Debug.Print "[TEST] ? Generación completada en " & Format(Timer - t0, "0.00") & " segundos."
    Debug.Print "[TEST] Ejecuta TEST_VerificarDistribucion para ver el resumen."
    Debug.Print "[TEST] ========================================================"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print "[TEST] ERROR: Nº" & Err.Number & " — " & Err.Description
End Sub

' ------------------------------------------------------------------------------
' Sub: TEST_VerificarDistribucion
' Muestra en la Ventana Inmediato cuántas filas tiene cada hoja Audit Trail.
' ------------------------------------------------------------------------------
Public Sub TEST_VerificarDistribucion()
    On Error Resume Next

    Dim i As Long
    Dim ws As Worksheet
    Dim filas As Long
    Dim totalFilas As Long
    Dim estadoVis As String

    Debug.Print "[TEST] ========================================================"
    Debug.Print "[TEST] DISTRIBUCIÓN DE REGISTROS AUDIT TRAIL"
    Debug.Print "[TEST] ========================================================"
    totalFilas = 0

    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        Dim nombreHoja As String
        Dim nombreTabla As String
        nombreHoja = ObtenerNombreHoja(i)
        nombreTabla = ObtenerNombreTabla(i)

        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(nombreHoja)

        If ws Is Nothing Then
            Debug.Print "[TEST]   Hoja " & i & " ('" & nombreHoja & "'): NO EXISTE en el libro"
        Else
            filas = ContarFilasTabla(ws, nombreTabla)
            If ws.Visible = xlSheetVisible Then
                estadoVis = "Visible"
            Else
                estadoVis = "Oculta"
            End If
            Dim label As String
            label = nombreHoja & String(20 - Len(nombreHoja), " ")
            Debug.Print "[TEST]   " & label & ": " & IIf(filas < 0, "tabla no encontrada", filas & " filas") & "  (" & estadoVis & ")"
            If filas > 0 Then totalFilas = totalFilas + filas
        End If
    Next i

    Debug.Print "[TEST] --------------------------------------------------------"
    Debug.Print "[TEST]   TOTAL                   : " & totalFilas & " registros"
    Debug.Print "[TEST] ========================================================"
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' Sub: TEST_LimpiarRegistrosPrueba
' *** ADVERTENCIA: elimina TODOS los datos de todas las tablas Audit Trail ***
' Usar SOLO en entorno de testing. Revierte el libro a estado inicial:
' hojas 2-5 ocultas y tabla de hoja 1 vacía.
' ------------------------------------------------------------------------------
Public Sub TEST_LimpiarRegistrosPrueba()
    On Error GoTo ErrorHandler

    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox( _
        "??  ADVERTENCIA" & vbCrLf & vbCrLf & _
        "Esto eliminará TODOS los registros de todas las tablas Audit Trail " & _
        "y ocultará las hojas 2 a 5." & vbCrLf & vbCrLf & _
        "¿Confirmar? (Solo usar en pruebas)", _
        vbYesNo + vbCritical, "Limpiar Audit Trail — Solo Testing")

    If respuesta <> vbYes Then Exit Sub

    Dim estadoEventos As Boolean
    estadoEventos = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Call WorkbookProtector2.UnprotectWorkbook

    Dim i As Long
    Dim ws As Worksheet
    Dim tbl As ListObject

    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        Dim nombreHoja As String
        Dim nombreTabla As String
        nombreHoja = ObtenerNombreHoja(i)
        nombreTabla = ObtenerNombreTabla(i)

        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(nombreHoja)
        On Error GoTo ErrorHandler

        If Not ws Is Nothing Then
            Call SheetProtector2.UnprotectSheet(ws, Configuration2.AUDIT_PASSWORD)

            ' Limpiar filas de la tabla dejando encabezados.
            Set tbl = Nothing
            On Error Resume Next
            Set tbl = ws.ListObjects(nombreTabla)
            On Error GoTo ErrorHandler

            If Not tbl Is Nothing Then
                If Not tbl.DataBodyRange Is Nothing Then
                    tbl.DataBodyRange.Delete
                    Debug.Print "[TEST] Tabla '" & nombreTabla & "' en '" & nombreHoja & "' vaciada."
                End If
            End If

            Call SheetProtector2.ProtectSheet(ws, Configuration2.AUDIT_PASSWORD)

            ' Ocultar hojas 2..5; dejar visible solo la 1.
            If i > 1 Then
                ws.Visible = xlSheetVeryHidden
                Debug.Print "[TEST] Hoja '" & nombreHoja & "' ocultada (xlSheetVeryHidden)."
            End If
        End If
    Next i

    Call WorkbookProtector2.ProtectWorkbook
    Application.EnableEvents = estadoEventos
    Application.ScreenUpdating = True

    Debug.Print "[TEST] ? Limpieza completada. Estado restaurado: solo 'Audit trail' visible y vacía."
    MsgBox "? Limpieza completada. Estado restaurado.", vbInformation, "Testing"
    Exit Sub

ErrorHandler:
    Call WorkbookProtector2.ProtectWorkbook
    Application.EnableEvents = estadoEventos
    Application.ScreenUpdating = True
    Debug.Print "[TEST] ERROR en limpieza: Nº" & Err.Number & " — " & Err.Description
    MsgBox "Error durante la limpieza: " & Err.Description, vbCritical, "Error"
End Sub

' ==============================================================================
' Sub: InicializarHojasAuditTrail
' Descripción: Crea automáticamente las 5 hojas "Audit trail" necesarias y sus
'              tablas ListObject correspondientes. Las hojas 2-5 se crean ocultas
'              (xlSheetVeryHidden) y se activan según sea necesario mediante
'              AuditRotation2.ObtenerHojaAuditActiva().
' Uso: Ejecutar UNA SOLA VEZ al inicializar el libro (o desde ThisWorkbook.Workbook_Open)
' ==============================================================================
Public Sub InicializarHojasAuditTrail()
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim nombreHoja As String
    Dim nombreTabla As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rango As Range

    Application.ScreenUpdating = False
    Call WorkbookProtector2.UnprotectWorkbook

    Debug.Print "[INIT] Inicializando hojas Audit Trail..."
    Debug.Print "[INIT] Cantidad de hojas a crear: " & Configuration2.AUDIT_MAX_SHEETS

    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        nombreHoja = ObtenerNombreHoja(i)
        nombreTabla = ObtenerNombreTabla(i)

        ' Verificar si la hoja ya existe
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(nombreHoja)
        On Error GoTo ErrorHandler

        If ws Is Nothing Then
            ' Crear la hoja nueva
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            ws.Name = nombreHoja
            Debug.Print "[INIT]   Hoja " & i & " ('" & nombreHoja & "'): CREADA"

            ' Crear encabezados estándar en fila 8
            With ws.Range("B8:K8")
                .Value = Array("Fecha", "Hora", "Usuario", "Hoja", "Acción", "Campo modificado", "Valor anterior", "Valor después", "Módulo/Subrutina")
            End With

            ' Crear tabla ListObject
            Set rango = ws.Range("B8:K8")
            Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=rango, _
                                         XlListObjectHasHeaders:=xlYes, TableStyleName:="TableStyleMedium2")
            tbl.Name = nombreTabla
            Debug.Print "[INIT]     Tabla '" & nombreTabla & "': CREADA"

            ' Proteger la hoja
            Call SheetProtector2.ProtectSheet(ws, Configuration2.AUDIT_PASSWORD)

            ' Ocultar las hojas 2-5 (la hoja 1 permanece visible)
            If i > 1 Then
                ws.Visible = xlSheetVeryHidden
                Debug.Print "[INIT]     Hoja " & i & " ('" & nombreHoja & "'): OCULTA (xlSheetVeryHidden)"
            End If
        Else
            Debug.Print "[INIT]   Hoja " & i & " ('" & nombreHoja & "'): YA EXISTE — se omite."
        End If
    Next i

    Call WorkbookProtector2.ProtectWorkbook
    Application.ScreenUpdating = True

    Debug.Print "[INIT] ? Inicialización completada. Todas las hojas Audit Trail listas."
    MsgBox "✓ Hojas Audit Trail inicializadas correctamente.", vbInformation, "Sistema de Auditoría"
    Exit Sub

ErrorHandler:
    Call WorkbookProtector2.ProtectWorkbook
    Application.ScreenUpdating = True
    Debug.Print "[INIT] ERROR: Nº" & Err.Number & " — " & Err.Description
    MsgBox "Error durante la inicialización: " & Err.Description, vbCritical, "Error"
End Sub