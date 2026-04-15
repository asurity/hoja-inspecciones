'******************************************************************************
' Módulo: mod_AuditRotation
' Proyecto: CA-HC-004 PROCESO DE VALIDACION
' Descripción: Sistema de rotación automática de hojas Audit Trail cuando
'              se alcanza el límite de filas configurado. Gestiona múltiples
'              hojas de auditoría para evitar saturar una sola tabla.
'
' Responsabilidades:
'   - Detectar cuándo una hoja Audit Trail está llena
'   - Obtener la hoja Audit Trail activa con espacio disponible
'   - Gestionar visibilidad de hojas (ocultar/mostrar según necesidad)
'   - Construir nombres de hojas y tablas dinámicamente
'   - Proporcionar funciones de testing y diagnóstico
'
' Características clave:
'   - Reutilizable: Copiar y pegar en cualquier proyecto VBA
'   - Desacoplado: Sin dependencias circulares
'   - Configurable: Usa constantes de Configuration.bas
'   - Testing: Funciones para generar datos de prueba y verificar distribución
'   - Logging: Debug.Print para trazabilidad completa
'
' Dependencias:
'   - Configuration.bas: Constantes AUDIT_*
'   - WorkbookProtector.bas: Protección/desprotección de estructura
'   - SheetProtector.bas: Protección de hojas individuales
'   - AuditLogger.bas: Escritura de registros (llamante)
'
' Integración:
'   - AuditLogger.LogAction() llama a ObtenerHojaAuditActiva()
'   - Reemplaza referencia fija a "Audit Trail" por detección dinámica
'
' Para reutilizar en otros proyectos:
'   1. Copiar este módulo completo
'   2. Ajustar constantes en Configuration.bas (AUDIT_*)
'   3. Crear hojas Audit Trail siguiendo nomenclatura estándar
'   4. Modificar AuditLogger para usar ObtenerHojaAuditActiva()
'
' Autor: Sistema CA-HC-004
' Fecha creación: 19/02/2026
' Última modificación: 19/02/2026 - Creación inicial
'******************************************************************************
Option Explicit

'******************************************************************************
' Función: ObtenerNombreTabla
' Descripción: Construye el nombre de la tabla ListObject correspondiente a
'              una hoja Audit Trail específica según su número de secuencia.
'
' Parámetros:
'   - numeroHoja (Long): Número de secuencia de la hoja (1 = primera hoja)
'
' Retorno:
'   - (String): Nombre de la tabla ListObject
'               - Hoja 1: "tblAudit" (sin sufijo, compatibilidad hoja original)
'               - Hojas 2+: "tblAudit2", "tblAudit3"... (sin guion bajo)
'
' Ejemplos:
'   ObtenerNombreTabla(1) ? "tblAudit"
'   ObtenerNombreTabla(2) ? "tblAudit2"
'   ObtenerNombreTabla(5) ? "tblAudit5"
'
' Usa configuración:
'   - Configuration.AUDIT_TABLE_PREFIX: Prefijo de las tablas ("tblAudit")
'
' Autor: Sistema CA-HC-004
' Fecha: 19/02/2026
'******************************************************************************
Public Function ObtenerNombreTabla(ByVal numeroHoja As Long) As String
    If numeroHoja = 1 Then
        ' Primera hoja: nombre sin sufijo para compatibilidad con tabla original
        ObtenerNombreTabla = Configuration2.AUDIT_TABLE_PREFIX
    Else
        ' Hojas adicionales: agregar número sin guion bajo
        ObtenerNombreTabla = Configuration2.AUDIT_TABLE_PREFIX & numeroHoja
    End If
End Function

'******************************************************************************
' Función: ObtenerNombreHoja
' Descripción: Construye el nombre de la hoja Audit Trail correspondiente a
'              un número de secuencia específico.
'
' Parámetros:
'   - numeroHoja (Long): Número de secuencia de la hoja (1 = primera hoja)
'
' Retorno:
'   - (String): Nombre de la hoja
'               - Hoja 1: "Audit Trail" (sin número)
'               - Hojas 2+: "Audit Trail 2", "Audit Trail 3"... (con "Trail" mayúscula)
'
' Ejemplos:
'   ObtenerNombreHoja(1) ? "Audit Trail"
'   ObtenerNombreHoja(2) ? "Audit Trail 2"
'   ObtenerNombreHoja(5) ? "Audit Trail 5"
'
' Usa configuración:
'   - Configuration.AUDIT_BASE_NAME: Nombre base de las hojas ("Audit Trail")
'
' Autor: Sistema CA-HC-004
' Fecha: 19/02/2026
'******************************************************************************
Public Function ObtenerNombreHoja(ByVal numeroHoja As Long) As String
    If numeroHoja = 1 Then
        ' Primera hoja: nombre sin número
        ObtenerNombreHoja = Configuration2.AUDIT_BASE_NAME
    Else
        ' Hojas adicionales: agregar número con espacio
        ObtenerNombreHoja = Configuration2.AUDIT_BASE_NAME & " " & numeroHoja
    End If
End Function

'******************************************************************************
' Función: ContarFilasTabla
' Descripción: Cuenta el número de filas con datos en una tabla ListObject.
'              Maneja correctamente tablas vacías y errores de tabla no encontrada.
'
' Parámetros:
'   - ws (Worksheet): Hoja que contiene la tabla
'   - nombreTabla (String): Nombre del ListObject a contar
'
' Retorno:
'   - (Long): Cantidad de filas con datos
'             - 0 si tabla está vacía (DataBodyRange Is Nothing)
'             - -1 si tabla no existe o error al acceder
'
' Ejemplo de uso:
'   Dim filas As Long
'   filas = ContarFilasTabla(Sheets("Audit Trail"), "tblAudit")
'   If filas >= 0 Then Debug.Print "Filas: " & filas
'
' Manejo de errores:
'   - On Error Resume Next: Si tabla no existe, retorna -1
'   - Verifica DataBodyRange Is Nothing para tablas vacías
'
' Autor: Sistema CA-HC-004
' Fecha: 19/02/2026
'******************************************************************************
Public Function ContarFilasTabla(ByVal ws As Worksheet, ByVal nombreTabla As String) As Long
    On Error Resume Next
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects(nombreTabla)
    
    ' Si la tabla no existe o hubo error al obtenerla
    If tbl Is Nothing Then
        ContarFilasTabla = -1
        Exit Function
    End If
    
    ' Si la tabla está vacía (solo tiene encabezados)
    If tbl.DataBodyRange Is Nothing Then
        ContarFilasTabla = 0
    Else
        ' Tabla tiene datos: contar filas
        ContarFilasTabla = tbl.ListRows.Count
    End If
    
    On Error GoTo 0
End Function

'******************************************************************************
' Función: DetectarSiHojaLlena
' Descripción: Verifica si una hoja Audit Trail ha alcanzado el límite de
'              filas configurado y necesita rotar a la siguiente hoja.
'
' Parámetros:
'   - ws (Worksheet): Hoja a verificar
'   - nombreTabla (String): Nombre de la tabla ListObject en esa hoja
'
' Retorno:
'   - (Boolean): True si hoja está llena (>= AUDIT_MAX_ROWS)
'                False si tiene espacio disponible o error
'
' Flujo:
'   1. Llama a ContarFilasTabla() para obtener cantidad actual
'   2. Compara con Configuration.AUDIT_MAX_ROWS
'   3. Debug.Print con información de estado
'
' Ejemplo de uso:
'   If DetectarSiHojaLlena(ws, "tblAudit") Then
'       Debug.Print "Hoja llena, necesita rotar"
'   End If
'
' Usa configuración:
'   - Configuration.AUDIT_MAX_ROWS: Límite de filas por hoja
'
' Autor: Sistema CA-HC-004
' Fecha: 19/02/2026
'******************************************************************************
Public Function DetectarSiHojaLlena(ByVal ws As Worksheet, ByVal nombreTabla As String) As Boolean
    Dim filasActuales As Long
    
    filasActuales = ContarFilasTabla(ws, nombreTabla)
    
    ' Si hubo error al contar (-1), considerar que NO está llena
    If filasActuales < 0 Then
        'Debug.Print "[AuditRotation] ERROR: No se pudo contar filas en " & ws.Name & " - " & nombreTabla
        DetectarSiHojaLlena = False
        Exit Function
    End If
    
    ' Comparar con el límite configurado
    If filasActuales >= Configuration2.AUDIT_MAX_ROWS Then
        'Debug.Print "[AuditRotation] Hoja LLENA: " & ws.Name & " (" & filasActuales & "/" & Configuration.AUDIT_MAX_ROWS & " filas)"
        DetectarSiHojaLlena = True
    Else
        'Debug.Print "[AuditRotation] Hoja con espacio: " & ws.Name & " (" & filasActuales & "/" & Configuration.AUDIT_MAX_ROWS & " filas)"
        DetectarSiHojaLlena = False
    End If
End Function

'******************************************************************************
' Función: ObtenerHojaAuditActiva
' Descripción: Obtiene la hoja Audit Trail activa con espacio disponible para
'              nuevos registros. Si la hoja actual está llena, busca la siguiente
'              y la hace visible automáticamente. Función central del sistema de rotación.
'
' Parámetros: Ninguno
'
' Retorno:
'   - (Worksheet): Hoja Audit Trail con espacio disponible
'                  - Si todas llenas: retorna última hoja (continúa escribiendo)
'                  - Si error: retorna Nothing
'
' Flujo de ejecución:
'   1. Recorre hojas 1 a AUDIT_MAX_SHEETS
'   2. Para cada hoja:
'      a. Construye nombre con ObtenerNombreHoja()
'      b. Verifica si existe en el libro
'      c. Obtiene nombre de tabla con ObtenerNombreTabla()
'      d. Verifica si está llena con DetectarSiHojaLlena()
'      e. Si NO está llena:
'         - Si está oculta (xlSheetVeryHidden): hace visible
'         - Retorna esa hoja
'   3. Si todas están llenas:
'      - Debug.Print CRÍTICO
'      - Retorna última hoja (no pierde datos)
'
' Visibilidad de hojas:
'   - Hoja 1: Siempre visible
'   - Hojas 2+: Ocultas (xlSheetVeryHidden) hasta que se necesiten
'   - Al necesitarse: Se hacen visibles automáticamente
'
' Manejo de estructura protegida:
'   - Desprotege estructura del libro para cambiar visibilidad
'   - Vuelve a proteger después del cambio
'
' Ejemplo de uso:
'   Dim wsAudit As Worksheet
'   Set wsAudit = ObtenerHojaAuditActiva()
'   If Not wsAudit Is Nothing Then
'       ' Escribir registro en wsAudit
'   End If
'
' Usa configuración:
'   - Configuration.AUDIT_MAX_SHEETS: Cantidad máxima de hojas
'   - Configuration.AUDIT_BASE_NAME: Nombre base de hojas
'
' Autor: Sistema CA-HC-004
' Fecha: 19/02/2026
'******************************************************************************
Public Function ObtenerHojaAuditActiva() As Worksheet
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim nombreHoja As String
    Dim nombreTabla As String
    Dim ws As Worksheet
    Dim estadoEventos As Boolean
    
    'Debug.Print "[AuditRotation] Buscando hoja Audit Trail activa..."
    
    ' Recorrer todas las hojas configuradas
    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        nombreHoja = ObtenerNombreHoja(i)
        nombreTabla = ObtenerNombreTabla(i)
        
        ' Verificar si la hoja existe
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(nombreHoja)
        On Error GoTo ErrorHandler
        
        If ws Is Nothing Then
            'Debug.Print "[AuditRotation] Hoja no encontrada: " & nombreHoja
            ' Continuar con la siguiente
        Else
            ' Verificar si la hoja está llena
            If Not DetectarSiHojaLlena(ws, nombreTabla) Then
                ' Esta hoja tiene espacio disponible
                ' NO cambiar visibilidad aquí. Se puede escribir en hojas ocultas.
                ' La visibilidad se controla exclusivamente desde NavigationService.
                
                ' Retornar esta hoja
                Set ObtenerHojaAuditActiva = ws
                'Debug.Print "[AuditRotation] ? Hoja activa seleccionada: " & nombreHoja
                Exit Function
            End If
        End If
    Next i
    
    ' Si llegamos aquí, TODAS las hojas están llenas
    Debug.Print "[AuditRotation] ?? CRÍTICO: TODAS las hojas Audit Trail están llenas!"
    Debug.Print "[AuditRotation] Se continuará escribiendo en la última hoja (excederá límite)"
    
    ' Retornar la última hoja para no perder datos
    nombreHoja = ObtenerNombreHoja(Configuration2.AUDIT_MAX_SHEETS)
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo ErrorHandler
    
    Set ObtenerHojaAuditActiva = ws
    Exit Function
    
ErrorHandler:
    'Debug.Print "[AuditRotation] ERROR: " & Err.Number & " - " & Err.Description
    Set ObtenerHojaAuditActiva = Nothing
End Function

'******************************************************************************
' Función: ObtenerTablaAuditActiva
' Descripción: Obtiene la tabla ListObject de la hoja Audit Trail activa.
'              Función de conveniencia que combina ObtenerHojaAuditActiva()
'              con la obtención del ListObject correspondiente.
'
' Parámetros: Ninguno
'
' Retorno:
'   - (ListObject): Tabla Audit Trail con espacio disponible
'                   - Nothing si error o no se encuentra
'
' Flujo:
'   1. Llama a ObtenerHojaAuditActiva() para obtener hoja correcta
'   2. Determina número de hoja (1, 2, 3...) según nombre
'   3. Construye nombre de tabla con ObtenerNombreTabla()
'   4. Retorna el ListObject de esa hoja
'
' Ejemplo de uso:
'   Dim tblAudit As ListObject
'   Set tblAudit = ObtenerTablaAuditActiva()
'   If Not tblAudit Is Nothing Then
'       tblAudit.ListRows.Add
'   End If
'
' Nota:
'   - Esta función podría usarse directamente en AuditLogger
'   - Simplifica el código al encapsular la lógica completa
'
' Autor: Sistema CA-HC-004
' Fecha: 19/02/2026
'******************************************************************************
Public Function ObtenerTablaAuditActiva() As ListObject
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim nombreTabla As String
    Dim numeroHoja As Long
    Dim i As Long
    
    ' Obtener la hoja activa
    Set ws = ObtenerHojaAuditActiva()
    
    If ws Is Nothing Then
        Set ObtenerTablaAuditActiva = Nothing
        Exit Function
    End If
    
    ' Determinar el número de hoja comparando nombres
    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        If ws.Name = ObtenerNombreHoja(i) Then
            numeroHoja = i
            Exit For
        End If
    Next i
    
    ' Obtener el nombre de la tabla correspondiente
    nombreTabla = ObtenerNombreTabla(numeroHoja)
    
    ' Retornar la tabla
    On Error Resume Next
    Set ObtenerTablaAuditActiva = ws.ListObjects(nombreTabla)
    On Error GoTo ErrorHandler
    
    If ObtenerTablaAuditActiva Is Nothing Then
        'Debug.Print "[AuditRotation] ERROR: Tabla no encontrada: " & nombreTabla & " en " & ws.Name
    End If
    
    Exit Function
    
ErrorHandler:
    'Debug.Print "[AuditRotation] ERROR en ObtenerTablaAuditActiva: " & Err.Number & " - " & Err.Description
    Set ObtenerTablaAuditActiva = Nothing
End Function

'******************************************************************************
' Sub: TEST_GenerarRegistrosAudit
' Descripción: Función de testing que genera N registros de auditoría ficticios
'              para probar el sistema de rotación automática.
'
' Parámetros:
'   - cantidadRegistros (Long): Cantidad de registros a generar (default: 250)
'
' Uso:
'   - Testing con límite de 100 filas: TEST_GenerarRegistrosAudit(250)
'   - Genera 250 registros distribuidos en: 100 + 100 + 50 entre 3 hojas
'
' Flujo:
'   1. Desactiva ScreenUpdating para velocidad
'   2. Loop de 1 a cantidadRegistros
'   3. Llama a AuditLogger.LogAction con datos ficticios
'   4. Debug.Print resumen cada 50 registros
'   5. Al finalizar: muestra distribución total
'
' ADVERTENCIA:
'   - Solo usar en ambiente de testing
'   - Configurar AUDIT_MAX_ROWS = 100 antes de ejecutar
'   - Limpiar con TEST_LimpiarRegistrosPrueba() después
'
' Autor: Sistema CA-HC-004
' Fecha: 19/02/2026
'   NOTA: Esta subrutina fue reemplazada por la versión en AuditRotation2.bas
'         Eliminada el 14/04/2026 para resolver ambigüedad de nombres.
'******************************************************************************

'******************************************************************************
' Sub: TEST_VerificarDistribucion
' Descripción: Función de testing que muestra la distribución de registros
'              entre todas las hojas Audit Trail.
'
' Parámetros: Ninguno
'
' Uso:
'   - Ejecutar después de TEST_GenerarRegistrosAudit()
'   - Muestra tabla resumen con conteo por hoja
'
' Output (Debug.Print):
'   ================================================
'   DISTRIBUCIÓN DE REGISTROS AUDIT TRAIL
'   ================================================
'   Audit Trail      : 100 filas (Visible)
'   Audit Trail 2    : 100 filas (Visible)
'   Audit Trail 3    : 50 filas (Visible)
'   Audit Trail 4    : 0 filas (Oculta)
'   Audit Trail 5    : 0 filas (Oculta)
'   ------------------------------------------------
'   TOTAL            : 250 registros
'   ================================================
'
' Autor: Sistema CA-HC-004
' Fecha: 19/02/2026
'******************************************************************************
Public Sub TEST_VerificarDistribucion()
    On Error Resume Next
    
    Dim i As Long
    Dim nombreHoja As String
    Dim nombreTabla As String
    Dim ws As Worksheet
    Dim filas As Long
    Dim totalFilas As Long
    Dim estadoVisibilidad As String
    
    Debug.Print "[TEST] ================================================"
    Debug.Print "[TEST] DISTRIBUCIÓN DE REGISTROS AUDIT TRAIL"
    Debug.Print "[TEST] ================================================"
    
    totalFilas = 0
    
    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        nombreHoja = ObtenerNombreHoja(i)
        nombreTabla = ObtenerNombreTabla(i)
        
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(nombreHoja)
        
        If ws Is Nothing Then
            Debug.Print "[TEST] " & nombreHoja & String(15 - Len(nombreHoja), " ") & ": NO EXISTE"
        Else
            filas = ContarFilasTabla(ws, nombreTabla)
            
            ' Determinar estado de visibilidad
            If ws.Visible = xlSheetVisible Then
                estadoVisibilidad = "(Visible)"
            ElseIf ws.Visible = xlSheetHidden Then
                estadoVisibilidad = "(Oculta)"
            Else
                estadoVisibilidad = "(Muy oculta)"
            End If
            
            If filas >= 0 Then
                Debug.Print "[TEST] " & nombreHoja & String(15 - Len(nombreHoja), " ") & ": " & _
                           Format(filas, "0") & " filas " & estadoVisibilidad
                totalFilas = totalFilas + filas
            Else
                Debug.Print "[TEST] " & nombreHoja & String(15 - Len(nombreHoja), " ") & ": ERROR al contar"
            End If
        End If
    Next i
    
    Debug.Print "[TEST] ------------------------------------------------"
    Debug.Print "[TEST] TOTAL" & String(15 - 5, " ") & ": " & totalFilas & " registros"
    Debug.Print "[TEST] ================================================"
    
    On Error GoTo 0
End Sub

'******************************************************************************
' Sub: TEST_LimpiarRegistrosPrueba
' Descripción: Función de testing que limpia todos los registros de prueba
'              de las tablas Audit Trail y oculta las hojas secundarias.
'
' Parámetros: Ninguno
'
' Uso:
'   - Ejecutar después de completar pruebas
'   - Restaura el estado inicial del sistema
'
' ADVERTENCIA:
'   - ?? ELIMINA TODOS LOS DATOS de las tablas Audit Trail
'   - Solo usar en ambiente de testing
'   - NO ejecutar en producción con datos reales
'
' Flujo:
'   1. Confirma con usuario (MsgBox)
'   2. Limpia datos de todas las tablas (deja encabezados)
'   3. Oculta hojas 2-5 (xlSheetVeryHidden)
'   4. Deja visible solo la primera hoja
'
' Autor: Sistema CA-HC-004
' Fecha: 19/02/2026
'******************************************************************************
Public Sub TEST_LimpiarRegistrosPrueba()
    On Error GoTo ErrorHandler
    
    Dim respuesta As VbMsgBoxResult
    Dim i As Long
    Dim nombreHoja As String
    Dim nombreTabla As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim estadoEventos As Boolean
    
    ' Confirmación de seguridad
    respuesta = MsgBox("?? ADVERTENCIA: Esto eliminará TODOS los registros de las tablas Audit Trail." & vbCrLf & vbCrLf & _
                       "¿Estás seguro de continuar?", vbYesNo + vbExclamation, "Confirmar Limpieza")
    
    If respuesta <> vbYes Then
        Debug.Print "[TEST] Limpieza cancelada por el usuario"
        Exit Sub
    End If
    
    Debug.Print "[TEST] ================================================"
    Debug.Print "[TEST] Limpiando registros de prueba..."
    Debug.Print "[TEST] ================================================"
    
    ' Deshabilitar eventos para evitar bucles al cambiar visibilidad
    estadoEventos = Application.EnableEvents
    Application.EnableEvents = False
    
    Application.ScreenUpdating = False
    
    ' Desproteger estructura del libro
    Call WorkbookProtector2.UnprotectWorkbook
    
    For i = 1 To Configuration2.AUDIT_MAX_SHEETS
        nombreHoja = ObtenerNombreHoja(i)
        nombreTabla = ObtenerNombreTabla(i)
        
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(nombreHoja)
        On Error GoTo ErrorHandler
        
        If Not ws Is Nothing Then
            ' Desproteger hoja
            Call SheetProtector2.UnprotectSheet(ws, Configuration2.AUDIT_PASSWORD)
            
            ' Obtener tabla
            On Error Resume Next
            Set tbl = Nothing
            Set tbl = ws.ListObjects(nombreTabla)
            On Error GoTo ErrorHandler
            
            If Not tbl Is Nothing Then
                ' Limpiar datos (dejar solo encabezados)
                On Error Resume Next
                If Not tbl.DataBodyRange Is Nothing Then
                    tbl.DataBodyRange.Delete
                    Debug.Print "[TEST] ? Limpiada: " & nombreHoja & " (" & nombreTabla & ")"
                Else
                    Debug.Print "[TEST] ? Ya vacía: " & nombreHoja
                End If
                On Error GoTo ErrorHandler
            End If
            
            ' Ocultar hojas secundarias (2+)
            If i > 1 Then
                ws.Visible = xlSheetVeryHidden
                Debug.Print "[TEST] ? Ocultada: " & nombreHoja
            End If
            
            ' Proteger hoja nuevamente
            Call SheetProtector2.ProtectSheet(ws, Configuration2.AUDIT_PASSWORD)
        End If
    Next i
    
    ' Proteger estructura del libro
    Call WorkbookProtector2.ProtectWorkbook
    
    ' Restaurar eventos
    Application.EnableEvents = estadoEventos
    Application.ScreenUpdating = True
    
    Debug.Print "[TEST] ================================================"
    Debug.Print "[TEST] ? Limpieza completada"
    Debug.Print "[TEST] Estado restaurado: Solo 'Audit Trail' visible y vacía"
    Debug.Print "[TEST] ================================================"
    
    MsgBox "? Limpieza completada exitosamente", vbInformation, "Testing"
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = estadoEventos
    Application.ScreenUpdating = True
    Call WorkbookProtector2.ProtectWorkbook
    Debug.Print "[TEST] ERROR: " & Err.Number & " - " & Err.Description
    MsgBox "Error durante limpieza: " & Err.Description, vbCritical, "Error"
End Sub