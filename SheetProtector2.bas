
' ----------------------------------------------------------------------
' Módulo: SheetProtector
' Descripción: Proporciona utilidades centralizadas para proteger y desproteger hojas de Excel
'              de manera segura y controlada, aplicando restricciones específicas según las
'              necesidades de la aplicación.
'              Facilita la administración de permisos y la gestión de errores en la protección.
' ----------------------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------------
' Subrutina: UnprotectSheet
' Propósito: Desprotege una hoja de Excel usando la contraseña proporcionada.
' Argumentos:
'   - ws: Referencia a la hoja a desproteger.
'   - sheetPassword: Contraseña para desproteger la hoja.
' Lógica:
'   1. Intenta desproteger la hoja con la contraseña.
'   2. Si ocurre un error, puede registrarse en el log (comentado por defecto).
' ----------------------------------------------------------------------
Public Sub UnprotectSheet(ByRef ws As Worksheet, ByVal sheetPassword As String)
    ' Si la protección de hojas está deshabilitada (modo desarrollo),
    ' esta función no hace nada (la hoja ya está desprotegida de todos modos)
    If Not Configuration2.ENABLE_SHEET_PROTECTION Then
        Debug.Print "[SheetProtector.UnprotectSheet] ■ SKIP: ENABLE_SHEET_PROTECTION = False (modo desarrollo)"
        Exit Sub
    End If
    
    Debug.Print "[SheetProtector.UnprotectSheet] INICIO | Hoja: '" & ws.Name & "' | Contraseña: '" & sheetPassword & "' (" & Len(sheetPassword) & " caracteres)"
    
    On Error GoTo ErrorHandler
    ws.Unprotect Password:=sheetPassword
    Debug.Print "[SheetProtector.UnprotectSheet] OK - Hoja '" & ws.Name & "' desprotegida exitosamente"
    Exit Sub
    
ErrorHandler:
    Debug.Print "[SheetProtector.UnprotectSheet] ⚠ ERROR desprotegiendo hoja '" & ws.Name & "': N°" & Err.Number & " - " & Err.Description
    ' Para producción, descomentar la siguiente línea para registrar el error:
    ' Call ErrorLogger.Log("SheetProtector.UnprotectSheet", VBA.Err.Description, VBA.Err.Number)
End Sub

' ----------------------------------------------------------------------
' Subrutina: ProtectSheetForReading
' Propósito: Protege una hoja de Excel permitiendo al usuario SOLO seleccionar
'            y copiar datos (lectura), sin poder editar, insertar ni eliminar.
'            Pensado para el rol "Usuario" en hojas de consulta como tblDetecciones.
' Argumentos:
'   - ws: Referencia a la hoja a proteger.
'   - sheetPassword: Contraseña para proteger la hoja.
' Escalabilidad: esta función puede ser invocada por cualquier hoja que requiera
'   acceso de solo lectura + copiado, sin duplicar lógica de protección (DRY).
' FASE 2 (22/02/2026): Creado para satisfacer el permiso de copiado del rol "Usuario" en Hoja5.
' ----------------------------------------------------------------------
Public Sub ProtectSheetForReading(ByRef ws As Worksheet, ByVal sheetPassword As String)
    ' Si la protección de hojas está deshabilitada (modo desarrollo),
    ' no hace nada para permitir edición libre durante testing
    If Not Configuration2.ENABLE_SHEET_PROTECTION Then
        Debug.Print "[SheetProtector.ProtectSheetForReading] ■ SKIP: ENABLE_SHEET_PROTECTION = False (modo desarrollo) - Hoja: '" & ws.Name & "'"
        Exit Sub
    End If
    
    Debug.Print "[SheetProtector.ProtectSheetForReading] INICIO | Hoja: '" & ws.Name & "' | Password: '" & sheetPassword & "'"
    
    On Error GoTo ErrorHandler
    ' Protege con las mismas restricciones que ProtectSheet, EXCEPTO:
    ' - ws.EnableSelection NO se establece a xlNoSelection, lo que permite
    '   seleccionar celdas y usar Ctrl+C para copiar sin poder modificarlas.
    ws.Protect Password:=sheetPassword, _
        UserInterfaceOnly:=False, _
        AllowFormattingCells:=False, _
        AllowFormattingColumns:=False, _
        AllowFormattingRows:=False, _
        AllowInsertingRows:=False, _
        AllowInsertingColumns:=False, _
        AllowInsertingHyperlinks:=False, _
        AllowDeletingRows:=False, _
        AllowDeletingColumns:=False, _
        AllowSorting:=False, _
        AllowFiltering:=True, _
        AllowUsingPivotTables:=False
    ' xlNoRestrictions permite seleccionar cualquier celda (bloqueada o no),
    ' habilitando el copiado sin otorgar permisos de escritura.
    ws.EnableSelection = xlNoRestrictions
    Debug.Print "[SheetProtector.ProtectSheetForReading] OK - Hoja '" & ws.Name & "' protegida (solo lectura + copiado)"
    Exit Sub
ErrorHandler:
    Debug.Print "[SheetProtector.ProtectSheetForReading] ⚠ ERROR: N°" & Err.Number & " - " & Err.Description & " | Hoja: '" & ws.Name & "'"
    Call ErrorLogger2.Log("SheetProtector2.ProtectSheetForReading", VBA.Err.Description, VBA.Err.Number)
End Sub

' ----------------------------------------------------------------------
Public Sub ProtectSheet(ByRef ws As Worksheet, ByVal sheetPassword As String)
    ' Si la protección de hojas está deshabilitada (modo desarrollo),
    ' no hace nada para permitir edición libre durante testing
    If Not Configuration2.ENABLE_SHEET_PROTECTION Then
        Debug.Print "[SheetProtector.ProtectSheet] ■ SKIP: ENABLE_SHEET_PROTECTION = False (modo desarrollo) - Hoja: '" & ws.Name & "'"
        Exit Sub
    End If
    
    Debug.Print "[SheetProtector.ProtectSheet] INICIO | Hoja: '" & ws.Name & "' | Password: '" & sheetPassword & "'"
    
    On Error GoTo ErrorHandler
    ws.Protect Password:=sheetPassword, _
        UserInterfaceOnly:=False, _
        AllowFormattingCells:=False, _
        AllowFormattingColumns:=False, _
        AllowFormattingRows:=False, _
        AllowInsertingRows:=False, _
        AllowInsertingColumns:=False, _
        AllowInsertingHyperlinks:=False, _
        AllowDeletingRows:=False, _
        AllowDeletingColumns:=False, _
        AllowSorting:=False, _
        AllowFiltering:=True, _
        AllowUsingPivotTables:=False
    ws.EnableSelection = xlNoSelection
    Debug.Print "[SheetProtector.ProtectSheet] OK - Hoja '" & ws.Name & "' protegida (sin selección)"
    Exit Sub
ErrorHandler:
    Debug.Print "[SheetProtector.ProtectSheet] ⚠ ERROR: N°" & Err.Number & " - " & Err.Description & " | Hoja: '" & ws.Name & "'"
    Call ErrorLogger2.Log("SheetProtector2.ProtectSheet", VBA.Err.Description, VBA.Err.Number)
End Sub

' ----------------------------------------------------------------------
' Subrutina: ApplyRoleBasedProtection
' Propósito: Centraliza la decisión de qué nivel de protección aplicar a una hoja
'            basándose en el rol activo del usuario (m_userRole).
'            Cualquier punto del sistema que necesite "proteger la hoja al terminar"
'            debe llamar esta función en lugar de llamar directamente a ProtectSheet,
'            garantizando que los permisos de cada rol se respeten siempre (DRY).
' Reglas:
'   - Admin   → UnprotectSheet (puede editar libremente).
'   - Usuario → ProtectSheetForReading (puede seleccionar y copiar, no editar).
'   - Otros   → ProtectSheet (sin acceso ni selección).
' FASE 1 (22/02/2026): Creada para corregir la brecha donde SaveDataToTable
'   llamaba a ProtectSheet directamente (sin rol), pisando el permiso de copiado
'   del rol "Usuario" establecido en Hoja5.Worksheet_Activate.
' ----------------------------------------------------------------------
Public Sub ApplyRoleBasedProtection(ByRef ws As Worksheet, ByVal sheetPassword As String)
    Debug.Print "[SheetProtector.ApplyRoleBasedProtection] INICIO | Hoja: '" & ws.Name & "' | Rol: '" & m_userRole & "' | Password: '" & sheetPassword & "'"
    
    On Error GoTo ErrorHandler
    ' FASE 1 (22/02/2026): Se desprotege siempre ANTES de re-proteger.
    ' Llamar a ws.Protect sobre una hoja ya protegida puede lanzar error 1004 en algunas
    ' versiones de Excel. El On Error de ProtectSheetForReading capturaba ese error en
    ' silencio, impidiendo que ws.EnableSelection = xlNoRestrictions se ejecutara, y dejando
    ' la hoja con xlNoSelection (sin copiado) para el rol "Usuario".
    ' Desproteger primero garantiza que ws.Protect siempre opera sobre una hoja sin protección.
    Call UnprotectSheet(ws, sheetPassword)
    
    If m_userRole = "Admin" Then
        ' Admin: queda desprotegida (ya se llamó UnprotectSheet arriba).
        Debug.Print "[SheetProtector.ApplyRoleBasedProtection] → Admin: hoja '" & ws.Name & "' queda DESPROTEGIDA"
    ElseIf m_userRole = "Usuario" Then
        Debug.Print "[SheetProtector.ApplyRoleBasedProtection] → Usuario: aplicando ProtectSheetForReading a '" & ws.Name & "'"
        Call ProtectSheetForReading(ws, sheetPassword)
    Else
        Debug.Print "[SheetProtector.ApplyRoleBasedProtection] → Otro ('" & m_userRole & "'): aplicando ProtectSheet a '" & ws.Name & "'"
        Call ProtectSheet(ws, sheetPassword)
    End If
    
    Debug.Print "[SheetProtector.ApplyRoleBasedProtection] FIN - Hoja: '" & ws.Name & "'"
    Exit Sub
ErrorHandler:
    Debug.Print "[SheetProtector.ApplyRoleBasedProtection] ⚠ ERROR: N°" & Err.Number & " - " & Err.Description & " | Hoja: '" & ws.Name & "' | Rol: '" & m_userRole & "'"
    Call ErrorLogger2.Log("SheetProtector2.ApplyRoleBasedProtection", VBA.Err.Description, VBA.Err.Number)
End Sub
