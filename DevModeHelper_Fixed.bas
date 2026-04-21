'═══════════════════════════════════════════════════════════════════════════════
' MÓDULO: DevModeHelper
' PROPÓSITO: Facilitar el desarrollo desprotegiendo y mostrando todas las hojas
' VERSIÓN: 2.0 (con ErrorHandler mejorado)
'═══════════════════════════════════════════════════════════════════════════════
' USO:
' - Call DevModeHelper.ActivarModoDesarrollo()
'     → Desprotege todo, muestra todas las hojas
' - Call DevModeHelper.DesactivarModoDesarrollo()
'     → Oculta hojas sensibles, protege todo
' - Call DevModeHelper.DesprotegerHojasPrincipales()
'     → Desprotege solo 5 hojas principales
'═══════════════════════════════════════════════════════════════════════════════
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' FUNCIÓN: ActivarModoDesarrollo
' PROPÓSITO: Unprotect all, show all, para desarrollo
'═══════════════════════════════════════════════════════════════════════════════
Public Sub ActivarModoDesarrollo()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim contador As Long
    contador = 0
    
    Application.ScreenUpdating = False
    
    ' 1. Desproteger estructura del libro
    Debug.Print "════════════════════════════════════════════════════════"
    Debug.Print "ACTIVANDO MODO DESARROLLO"
    Debug.Print "════════════════════════════════════════════════════════"
    
    If ThisWorkbook.ProtectStructure Then
        On Error Resume Next
        ThisWorkbook.Unprotect Configuration2.APP_PASSWORD
        If Err.Number <> 0 Then
            Debug.Print "⚠ No se pudo desproteger libro: " & Err.Description
            Err.Clear
        Else
            Debug.Print "✓ Libro desprotegido"
        End If
        On Error GoTo ErrorHandler
    Else
        Debug.Print "  Libro ya estaba desprotegido"
    End If
    
    ' 2. Hacer visibles TODAS las hojas (incluyendo muy ocultas)
    contador = 0
    Debug.Print ""
    Debug.Print "Haciendo todas las hojas visibles..."
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            Debug.Print "  → " & ws.Name & " (ahora visible)"
            contador = contador + 1
        End If
    Next ws
    
    If contador = 0 Then
        Debug.Print "  (Todas las hojas ya eran visibles)"
    Else
        Debug.Print "✓ " & contador & " hojas hechas visibles"
    End If
    
    ' 3. Desproteger TODAS las hojas
    contador = 0
    Debug.Print ""
    Debug.Print "Desprotegiendo todas las hojas..."
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Then
            On Error Resume Next
            ' Intentar primero con APP_PASSWORD
            ws.Unprotect Configuration2.APP_PASSWORD
            If Err.Number <> 0 Then
                ' Si falla, intentar con AUDIT_PASSWORD
                Err.Clear
                ws.Unprotect Configuration2.AUDIT_PASSWORD
            End If
            
            If Err.Number <> 0 Then
                Debug.Print "  ⚠ No se pudo desproteger: " & ws.Name & " (" & Err.Description & ")"
                Err.Clear
            Else
                Debug.Print "  → " & ws.Name
                contador = contador + 1
            End If
            On Error GoTo ErrorHandler
        End If
    Next ws
    
    If contador = 0 Then
        Debug.Print "  (Todas las hojas ya estaban desprotegidas)"
    Else
        Debug.Print "✓ " & contador & " hojas desprotegidas"
    End If
    
    Application.ScreenUpdating = True
    
    Debug.Print ""
    Debug.Print "════════════════════════════════════════════════════════"
    Debug.Print "MODO DESARROLLO ACTIVADO"
    Debug.Print "════════════════════════════════════════════════════════"
    
    MsgBox "Modo Desarrollo Activado" & vbCrLf & vbCrLf & _
           "• Libro desprotegido" & vbCrLf & _
           "• Todas las hojas visibles" & vbCrLf & _
           "• Todas las hojas desprotegidas" & vbCrLf & vbCrLf & _
           "Cuando termines, ejecuta:" & vbCrLf & _
           "   Call DevModeHelper.DesactivarModoDesarrollo()", _
           vbInformation, "Desarrollo"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print ""
    Debug.Print "✗ ERROR: " & Err.Number & " - " & Err.Description
    MsgBox "Error al activar modo desarrollo:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical
End Sub

'═══════════════════════════════════════════════════════════════════════════════
' FUNCIÓN: DesactivarModoDesarrollo
' PROPÓSITO: Volver a ocultar y proteger para producción
'═══════════════════════════════════════════════════════════════════════════════
Public Sub DesactivarModoDesarrollo()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim contador As Long
    
    Application.ScreenUpdating = False
    
    Debug.Print "════════════════════════════════════════════════════════"
    Debug.Print "DESACTIVANDO MODO DESARROLLO"
    Debug.Print "════════════════════════════════════════════════════════"
    
    ' 1. Ocultar hojas que deberían estar ocultas
    contador = 0
    Debug.Print "Ocultando hojas sensibles..."
    
    ' Plantilla Certificado → Muy oculta
    On Error Resume Next
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets(Configuration2.SHEET_PLANTILLA_CERTIFICADO)
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVeryHidden
        Debug.Print "  → " & ws.Name & " (muy oculta)"
        contador = contador + 1
    Else
        Debug.Print "  (Plantilla Certificado no encontrada)"
    End If
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' Audit trails → Muy ocultos (excepto el primero que puede estar visible)
    Dim i As Long
    For i = 2 To Configuration2.AUDIT_MAX_SHEETS
        On Error Resume Next
        Set ws = Nothing
        If i = 1 Then
            Set ws = ThisWorkbook.Sheets(Configuration2.AUDIT_BASE_NAME)
        Else
            Set ws = ThisWorkbook.Sheets(Configuration2.AUDIT_BASE_NAME & " " & i)
        End If
        
        If Not ws Is Nothing And i > 1 Then
            ws.Visible = xlSheetVeryHidden
            Debug.Print "  → " & ws.Name & " (muy oculta)"
            contador = contador + 1
        End If
        Err.Clear
        On Error GoTo ErrorHandler
    Next i
    
    If contador = 0 Then
        Debug.Print "  (No hay hojas para ocultar)"
    Else
        Debug.Print "✓ " & contador & " hojas ocultadas"
    End If
    
    ' 2. Proteger hojas importantes
    contador = 0
    Debug.Print ""
    Debug.Print "Protegiendo hojas..."
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        
        ' Proteger Audit trails con contraseña especial
        If InStr(ws.Name, Configuration2.AUDIT_BASE_NAME) > 0 Then
            ws.Protect Password:=Configuration2.AUDIT_PASSWORD, _
                       DrawingObjects:=True, _
                       Contents:=True, _
                       Scenarios:=True, _
                       AllowFiltering:=True
            If Err.Number = 0 Then
                Debug.Print "  → " & ws.Name & " (audit)"
                contador = contador + 1
            End If
        ' Proteger hojas normales (excepto menú principal si existe)
        ElseIf ws.Name <> Configuration2.MAIN_MENU_SHEET Then
            ws.Protect Password:=Configuration2.APP_PASSWORD, _
                       DrawingObjects:=True, _
                       Contents:=True, _
                       Scenarios:=True, _
                       AllowFiltering:=True
            If Err.Number = 0 Then
                Debug.Print "  → " & ws.Name
                contador = contador + 1
            End If
        End If
        Err.Clear
        On Error GoTo ErrorHandler
    Next ws
    
    If contador = 0 Then
        Debug.Print "  (No se protegieron hojas)"
    Else
        Debug.Print "✓ " & contador & " hojas protegidas"
    End If
    
    ' 3. Proteger estructura del libro
    Debug.Print ""
    
    On Error Resume Next
    Dim enableProtection As Boolean
    enableProtection = Configuration2.ENABLE_WORKBOOK_PROTECTION
    If Err.Number <> 0 Then
        enableProtection = True ' Default si la constante no existe
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    If enableProtection Then
        ThisWorkbook.Protect Password:=Configuration2.APP_PASSWORD, Structure:=True
        Debug.Print "✓ Libro protegido (estructura)"
    Else
        Debug.Print "  Libro sin protección (ENABLE_WORKBOOK_PROTECTION = False)"
    End If
    
    Application.ScreenUpdating = True
    
    Debug.Print ""
    Debug.Print "════════════════════════════════════════════════════════"
    Debug.Print "MODO DESARROLLO DESACTIVADO"
    Debug.Print "════════════════════════════════════════════════════════"
    
    MsgBox "Modo Desarrollo Desactivado" & vbCrLf & vbCrLf & _
           "• Hojas sensibles ocultas" & vbCrLf & _
           "• Hojas protegidas" & vbCrLf & _
           "• Libro protegido" & vbCrLf & vbCrLf & _
           "Sistema listo para producción.", _
           vbInformation, "Producción"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print ""
    Debug.Print "✗ ERROR: " & Err.Number & " - " & Err.Description
    MsgBox "Error al desactivar modo desarrollo:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical
End Sub

'═══════════════════════════════════════════════════════════════════════════════
' FUNCIÓN: DesprotegerHojasPrincipales
' PROPÓSITO: Desproteger solo las 5 hojas principales para desarrollo rápido
'═══════════════════════════════════════════════════════════════════════════════
Public Sub DesprotegerHojasPrincipales()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Debug.Print "════════════════════════════════════════════════════════"
    Debug.Print "DESPROTEGIENDO HOJAS PRINCIPALES"
    Debug.Print "════════════════════════════════════════════════════════"
    
    Dim hojasPrincipales As Variant
    hojasPrincipales = Array( _
        Configuration2.SHEET_HISTORICO, _
        Configuration2.SHEET_CHECKLIST, _
        Configuration2.SHEET_PERSONAL, _
        Configuration2.SHEET_CRONOGRAMA, _
        Configuration2.SHEET_CONFIGURACION _
    )
    
    Dim hoja As Variant
    Dim ws As Worksheet
    Dim contador As Long
    contador = 0
    
    For Each hoja In hojasPrincipales
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(CStr(hoja))
        
        If Not ws Is Nothing Then
            If ws.ProtectContents Then
                ws.Unprotect Configuration2.APP_PASSWORD
                If Err.Number = 0 Then
                    Debug.Print "  → " & ws.Name
                    contador = contador + 1
                Else
                    Debug.Print "  ⚠ No se pudo desproteger: " & ws.Name
                End If
            Else
                Debug.Print "  • " & ws.Name & " (ya desprotegida)"
            End If
        Else
            Debug.Print "  ⚠ No encontrada: " & hoja
        End If
        Err.Clear
        On Error GoTo ErrorHandler
    Next hoja
    
    Application.ScreenUpdating = True
    
    Debug.Print ""
    Debug.Print "════════════════════════════════════════════════════════"
    Debug.Print "HOJAS PRINCIPALES DESPROTEGIDAS"
    Debug.Print "════════════════════════════════════════════════════════"
    
    MsgBox "Hojas principales desprotegidas:" & vbCrLf & vbCrLf & _
           "• Historico" & vbCrLf & _
           "• Checklist" & vbCrLf & _
           "• Personal" & vbCrLf & _
           "• Cronograma" & vbCrLf & _
           "• Configuración" & vbCrLf & vbCrLf & _
           "(Estructura del libro sigue protegida)", _
           vbInformation, "Desarrollo"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print ""
    Debug.Print "✗ ERROR: " & Err.Number & " - " & Err.Description
    MsgBox "Error al desproteger hojas principales:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical
End Sub

'═══════════════════════════════════════════════════════════════════════════════
' FIN DEL MÓDULO
'═══════════════════════════════════════════════════════════════════════════════
