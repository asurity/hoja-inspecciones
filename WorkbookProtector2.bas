
' ----------------------------------------------------------------------
' Módulo: WorkbookProtector
' Descripción: Gestiona la protección y desprotección de la estructura del libro.
'              Permite separar la lógica de protección del libro de la de las hojas individuales,
'              facilitando el mantenimiento y la seguridad global del archivo.
' ----------------------------------------------------------------------
Option Explicit

' ## NAVEGACIÓN ## Caché de estado de protección (FASE 5, 08/06/2026)
' Evita desproteger si ya está desprotegido, reduciendo operaciones redundantes.
Private m_IsProtected As Boolean

' ----------------------------------------------------------------------
' Subrutina: ToggleProtection
' Propósito: Activa o desactiva la protección de estructura según el parámetro.
'            Simplifica las llamadas desde el motor de navegación.
' Creado: 08/06/2026 - FASE 5 Refactorización Navegación
' ----------------------------------------------------------------------
Public Sub ToggleProtection(ByVal enable As Boolean)
    If enable Then
        Call ProtectWorkbook
    Else
        Call UnprotectWorkbook
    End If
End Sub

' ----------------------------------------------------------------------
' Subrutina: ProtectWorkbook
' Propósito: Protege la estructura del libro (no las hojas individuales) usando la contraseña definida.
' Lógica:
'   1. Si ya está protegido (según caché), sale sin hacer nada.
'   2. Protege la estructura del libro para evitar que se agreguen, eliminen o muevan hojas.
'   3. Maneja errores y registra en el log si ocurre alguno.
' Refactorizado: 08/06/2026 - FASE 5 (caché de estado)
' ----------------------------------------------------------------------
Public Sub ProtectWorkbook()
    If m_IsProtected Then Exit Sub  ' ## NAVEGACIÓN ## caché (FASE 5)
    
    Dim pwd As String
    pwd = Configuration2.APP_PASSWORD
    
    On Error Resume Next
    ThisWorkbook.Protect Password:=pwd, Structure:=True, Windows:=False
    If VBA.Err.Number <> 0 Then
        ' NOTA (17/06/2026): Si el error es porque el workbook YA está protegido
        ' (ej. archivo .xlsm guardado con protección), igual marcamos el caché.
        ' Así UnprotectWorkbook sabrá que debe desproteger antes de ShowOnly.
        m_IsProtected = True
        Call ErrorLogger2.Log("WorkbookProtector2.ProtectWorkbook", VBA.Err.Description, VBA.Err.Number)
    Else
        m_IsProtected = True
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------
' Subrutina: UnprotectWorkbook
' Propósito: Desprotege la estructura del libro. Primero intenta con APP_PASSWORD,
'            y si falla, prueba con TODAS las contraseñas históricas conocidas
'            del proyecto para detectar cuál es la correcta.
' Lógica:
'   1. Intenta con APP_PASSWORD de Configuration2 primero.
'   2. Si falla, prueba todas las contraseñas históricas conocidas.
'   3. Reporta cuál funcionó (si alguna).
' ----------------------------------------------------------------------
Public Sub UnprotectWorkbook()
    ' NOTA (17/06/2026): Ya NO se sale temprano cuando m_IsProtected=False.
    ' Si el archivo .xlsm se guardó con protección de estructura activa,
    ' ProtectWorkbook() falla silenciosamente (Err=1004, workbook ya protegido)
    ' y m_IsProtected se queda en False. Esto causaba que ShowOnly() tronara
    ' con error 1004 al intentar cambiar visibilidad de hojas.
    ' Ahora SIEMPRE intenta desproteger, actualizando el caché si funciona.
    
    Dim pwd As String
    pwd = Configuration2.APP_PASSWORD
    
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=pwd
    If VBA.Err.Number <> 0 Then
        ' Fallback: intentar con contraseñas históricas conocidas
        Dim historicas As Variant
        historicas = Array( _
            "Inspecciones2026", "Validacion003", "1234", "5678", _
            "supervisor002.", "Validacion2025", "Inspecciones2025", _
            "Aseguramiento2026", "CronoAdmin2026*", "" _
        )
        Dim i As Long
        For i = LBound(historicas) To UBound(historicas)
            Err.Clear
            ThisWorkbook.Unprotect Password:=historicas(i)
            If Err.Number = 0 Then
                m_IsProtected = False
                Exit Sub
            End If
        Next i
        Call ErrorLogger2.Log("WorkbookProtector2.UnprotectWorkbook", "Ninguna contraseña conocida funcionó", 1004)
    Else
        m_IsProtected = False
    End If
    On Error GoTo 0
End Sub


' ----------------------------------------------------------------------
' Subrutina: DiagnosticarContrasenasHojas
' Propósito: Itera TODAS las hojas del libro, detecta si están protegidas,
'            prueba desproteger cada una con todas las contraseñas conocidas
'            del sistema, y muestra un MsgBox con el resultado.
' Uso: Llamar manualmente desde la ventana Immediate (Ctrl+G):
'         Call WorkbookProtector2.DiagnosticarContrasenasHojas
' Creado: 17/06/2026 - Diagnóstico de conflictos de contraseñas
' ----------------------------------------------------------------------
Public Sub DiagnosticarContrasenasHojas()
    Dim ws As Worksheet
    Dim passwords As Variant
    Dim resultado As String
    Dim i As Long
    Dim foundPassword As String
    Dim isProtected As Boolean
    Dim intentos As Long
    
    ' Todas las contraseñas conocidas del sistema + históricas
    passwords = Array( _
        Configuration2.APP_PASSWORD, _
        Configuration2.AUDIT_PASSWORD, _
        Configuration2.ADMIN_PASSWORD, _
        Configuration2.CRONOGRAMA_ADMIN_PASSWORD, _
        "Inspecciones2026", "Validacion003", "1234", "5678", _
        "Validacion2025", "Inspecciones2025", _
        "Aseguramiento2026", "CronoAdmin2026*", "" _
    )
    
    resultado = "DIAGNÓSTICO DE CONTRASEÑAS POR HOJA" & vbCrLf & _
                String(45, "=") & vbCrLf & vbCrLf
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        isProtected = ws.ProtectContents
        On Error GoTo 0
        
        If Not isProtected Then
            resultado = resultado & ws.Name & "  →  SIN PROTECCIÓN" & vbCrLf
        Else
            foundPassword = ""
            For i = LBound(passwords) To UBound(passwords)
                intentos = intentos + 1
                On Error Resume Next
                ws.Unprotect Password:=CStr(passwords(i))
                If Err.Number = 0 Then
                    foundPassword = CStr(passwords(i))
                    ' Re-proteger inmediatamente con la misma contraseña
                    ws.Protect Password:=foundPassword
                    On Error GoTo 0
                    Exit For
                End If
                On Error GoTo 0
            Next i
            
            If foundPassword = "" Then
                resultado = resultado & ws.Name & "  →  ⚠ DESCONOCIDA (ninguna funcionó)" & vbCrLf
            ElseIf Len(foundPassword) = 0 Then
                resultado = resultado & ws.Name & "  →  [vacía / sin contraseña]" & vbCrLf
            Else
                resultado = resultado & ws.Name & "  →  """ & foundPassword & """" & vbCrLf
            End If
        End If
    Next ws
    
    resultado = resultado & vbCrLf & _
                "--- CONTRASEÑAS DEL SISTEMA ---" & vbCrLf & _
                "APP_PASSWORD               = ""supervisor002.""" & vbCrLf & _
                "AUDIT_PASSWORD             = ""validacion002.""" & vbCrLf & _
                "ADMIN_PASSWORD             = ""validacion002.""" & vbCrLf & _
                "CRONOGRAMA_ADMIN_PASSWORD  = ""validacion002.""" & vbCrLf & _
                vbCrLf & _
                "Total hojas analizadas: " & ThisWorkbook.Worksheets.Count & vbCrLf & _
                "Total intentos: " & intentos
    
    MsgBox resultado, vbInformation, "Diagnóstico de Contraseñas"
End Sub


' ----------------------------------------------------------------------
' Subrutina: ProtegerHojasCliente
' Propósito: Protege todas las hojas del sistema que están sin protección
'            con APP_PASSWORD ("supervisor002."), usando SheetProtector2.ProtectSheet.
'            Es una operación única solicitada por el cliente.
' Hojas a proteger: Graficos, Formulario de inspeccion, Cronograma, Checklist,
'                   Historico, Configuración, Aseguramiento de calidad,
'                   Personal, Registro de errores, Plantilla Certificado
' Uso: Llamar desde Immediate (Ctrl+G):
'         Call WorkbookProtector2.ProtegerHojasCliente
' Creado: 17/06/2026 - Solicitud del cliente
' ----------------------------------------------------------------------
Public Sub ProtegerHojasCliente()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim exitos As Long, fallos As Long
    Dim reporte As String
    
    sheetNames = Array( _
        "Graficos", _
        "Formulario de inspeccion", _
        "Cronograma", _
        "Checklist", _
        "Historico", _
        "Configuración", _
        "Aseguramiento de calidad", _
        "Personal", _
        "Registro de errores", _
        "Plantilla Certificado" _
    )
    
    reporte = "PROTECCIÓN DE HOJAS (APP_PASSWORD)" & vbCrLf & _
              String(45, "=") & vbCrLf & vbCrLf
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sheetNames(i)))
        On Error GoTo 0
        
        If ws Is Nothing Then
            reporte = reporte & "✗ " & sheetNames(i) & "  →  NO ENCONTRADA" & vbCrLf
            fallos = fallos + 1
        Else
            ' Desproteger workbook si es necesario para poder proteger hojas
            Call UnprotectWorkbook
            
            ' Desproteger hoja primero (por si tiene protección residual)
            On Error Resume Next
            ws.Unprotect Password:=Configuration2.APP_PASSWORD
            On Error GoTo 0
            
            ' Proteger con el método estándar del sistema
            On Error Resume Next
            Call SheetProtector2.ProtectSheet(ws, Configuration2.APP_PASSWORD)
            If Err.Number <> 0 Then
                reporte = reporte & "✗ " & ws.Name & "  →  ERROR: " & Err.Description & vbCrLf
                fallos = fallos + 1
            Else
                reporte = reporte & "✓ " & ws.Name & "  →  PROTEGIDA" & vbCrLf
                exitos = exitos + 1
            End If
            On Error GoTo 0
            
            Set ws = Nothing
        End If
    Next i
    
    ' Re-proteger workbook
    Call ProtectWorkbook
    
    reporte = reporte & vbCrLf & _
              "Éxitos: " & exitos & " | Fallos: " & fallos
    
    MsgBox reporte, vbInformation, "Protección de Hojas - Cliente"
End Sub