' ══════════════════════════════════════════════════════════════
' Módulo: DevModeHelper
' Propósito: Facilitar desarrollo desprotegiendo y mostrando todo
' Fecha: 17/04/2026
' ══════════════════════════════════════════════════════════════
Option Explicit

' ══════════════════════════════════════════════════════════════
' MODO DESARROLLO: Desproteger y mostrar todo
' ══════════════════════════════════════════════════════════════
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
    
    On Error Resume Next
    If ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Unprotect Configuration2.APP_PASSWORD
        If Err.Number = 0 Then
            Debug.Print "✓ Libro desprotegido"
        Else
            Debug.Print "⚠ No se pudo desproteger libro: " & Err.Description
        End If
    Else
        Debug.Print "  Libro ya estaba desprotegido"
    End If
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' 2. Hacer visibles todas las hojas (incluso las muy ocultas)
    Debug.Print ""
    Debug.Print "Haciendo visibles todas las hojas..."
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then
            Debug.Print "  → " & ws.Name & " (era: " & GetVisibilityName(ws.Visible) & ")"
            ws.Visible = xlSheetVisible
            contador = contador + 1
        End If
    Next ws
    Debug.Print "✓ " & contador & " hojas ahora visibles"
    
    ' 3. Desproteger todas las hojas
    contador = 0
    Debug.Print ""
    Debug.Print "Desprotegiendo todas las hojas..."
    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
            On Error Resume Next
            ws.Unprotect Configuration2.APP_PASSWORD
            If Err.Number = 0 Then
                Debug.Print "  → " & ws.Name
                contador = contador + 1
            Else
                ' Intentar con contraseña de audit trail
                ws.Unprotect Configuration2.AUDIT_PASSWORD
                If Err.Number = 0 Then
                    Debug.Print "  → " & ws.Name & " (audit trail)"
                    contador = contador + 1
                Else
                    Debug.Print "  ✗ " & ws.Name & " (error: " & Err.Description & ")"
                End If
            End If
            On Error GoTo 0
        End If
    Next ws
    Debug.Print "✓ " & contador & " hojas desprotegidas"
    
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
    MsgBox "EGoTo ErrorHandler
    
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
        Debug.Print "  (Plantilla Certificado no existe)"
    End If
    On Error GoTo ErrorHandler
    ' Plantilla Certificado → Muy oculta
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Configuration2.SHEET_PLANTILLA_CERTIFICADO)
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVeryHidden
        Debug.Print "  → " & ws.Name & " (muy oculta)"
        contador = contador + 1
    End If
    On Error GoTo 0
    
    ' Audit trails → Muy ocultos (excepto el primero que puede estar visible)
    Dim i As Long
    For i = 2 To Configuration2.AUDIT_MAX_SHEETS
        On Error Resume Next
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
        On Error GoTo 0
    Next i
    
    Debug.Print "✓ " & contador & " hojas ocultadas"
    
    ' 2. Proteger hojas importantes
    contador = 0
    Debug.Print ""
    Debug.Print "Protegiendo hojas..."
    
    For Each ws In ThisWorkbook.Worksheets
        ' Proteger Audit trails con contraseña especial
        If InStr(ws.Name, Configuration2.AUDIT_BASE_NAME) > 0 Then
            ws.Protect Password:=Configuration2.AUDIT_PASSWORD, _
                       DrawingObjects:=True, _
                       Contents:=True, _
                       Scenarios:=True, _
                       AllowFiltering:=True
            Debug.Print "  → " & ws.Name & " (audit)"
            contador = contador + 1
        ' Proteger hojas normales
        ElseIf ws.Name <> Configuration2.MAIN_MENU_SHEET Then
            ws.Protect Password:=Configuration2.APP_PASSWORD, _
                       DrawingObjects:=True, _
                       Contents:=True, _
                       Scenarios:=True, _
                       AllowFiltering:=True
            Debug.Print "  → " & ws.Name
            contador = contador + 1
        End If
    Next ws
    Debug.Print "✓ " & contador & " hojas protegidas"
    On Error Resume Next
    Dim enableProtection As Boolean
    enableProtection = Configuration2.ENABLE_WORKBOOK_PROTECTION
    If Err.Number <> 0 Then
        enableProtection = True ' Default si la constante no existe
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    If enableProtection
    ' 3. Proteger estructura del libro
    Debug.Print ""
    If Configuration2.ENABLE_WORKBOOK_PROTECTION Then
        ThisWorkbook.Protect Password:=Configuration2.APP_PASSWORD, Structure:=True
        Debug.Print "✓ Libro protegido (estructura)"
    Else
        Debug.Print "  Libro sin protección (ENABLE_WORKBOOK_PROTECTION = False)"
    End If
    
    Application.ScreenUpdating = True
    
    Debug.Print ""
    Debug.Print "════════════════════════════════════════════════════════"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print ""
    Debug.Print "✗ ERROR: " & Err.Number & " - " & Err.Description
    MsgBox "Error al desactivar modo desarrollo:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical
    Debug.Print "MODO DESARROLLO DESACTIVADO"
    Debug.Print "════════════════════════════════════════════════════════"
    
    MsgBox "Modo Desarrollo Desactivado" & vbCrLf & vbCrLf & _
           "• Hojas sensibles ocultas" & vbCrLf & _
           "• Hojas protegidas" & vbCrLf & _
           "• Libro protegido" & vbCrLf & vbCrLf & _
           "Sistema listo para producción.", _
           vbInformation, "Producción"
End Sub

' ══════════════════════════════════════════════════════════════
' Helper: Obtener nombre legible de visibilidad
' ══════════════════════════════════════════════════════════════
Private Function GetVisibilityName(vis As XlSheetVisibility) As String
    Select Case vis
        Case xlSheetVisible
            GetVisibilityName = "visible"
        Case xlSheetHidden
            GetVisibilityName = "oculta"
        Case xlSheetVeryHidden
            GetVisibilityName = "muy oculta"
        Case Else
            GetVisibilityName = "desconocido"
    End Select
End Function

' ══════════════════════════════════════════════════════════════
' ATAJO RÁPIDO: Desproteger solo las hojas principales
' (sin afectar visibilidad ni estructura del libro)
' ══════════════════════════════════════════════════════════════
Public Sub DesprotegerHojasPrincipales()
    On Error Resume Next
    
    Dim hojas As Variant
    hojas = Array("Historico", "Checklist", "Personal", "Cronograma", "Configuración")
    
    Dim nombreHoja As Variant
    Dim ws As Worksheet
    
    Debug.Print "Desprotegiendo hojas principales..."
    
    For Each nombreHoja In hojas
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(CStr(nombreHoja))
        
        If Not ws Is Nothing Then
            ws.Unprotect Configuration2.APP_PASSWORD
            Debug.Print "  ✓ " & nombreHoja
        End If
    Next nombreHoja
    
    Debug.Print "Listo."
    MsgBox "Hojas principales desprotegidas.", vbInformation
End Sub
