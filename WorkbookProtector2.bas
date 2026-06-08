
' ----------------------------------------------------------------------
' Módulo: WorkbookProtector
' Descripción: Gestiona la protección y desprotección de la estructura del libro.
'              Permite separar la lógica de protección del libro de la de las hojas individuales,
'              facilitando el mantenimiento y la seguridad global del archivo.
' ----------------------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------------
' Subrutina: ProtectWorkbook
' Propósito: Protege la estructura del libro (no las hojas individuales) usando la contraseña definida.
' Lógica:
'   1. Protege la estructura del libro para evitar que se agreguen, eliminen o muevan hojas.
'   2. Maneja errores y registra en el log si ocurre alguno.
' ----------------------------------------------------------------------
Public Sub ProtectWorkbook()
    Dim pwd As String
    pwd = Configuration2.APP_PASSWORD
    Debug.Print "[WorkbookProtector.ProtectWorkbook] INICIO | Contraseña a usar: '" & pwd & "' (" & Len(pwd) & " caracteres)"
    
    ' On Error Resume Next: si el libro ya está protegido, Excel lanza Error 1004.
    ' Ignoramos ese caso en silencio para evitar interrumpir el flujo del llamador.
    On Error Resume Next
    ThisWorkbook.Protect Password:=pwd, Structure:=True, Windows:=False
    If VBA.Err.Number <> 0 Then
        Debug.Print "[WorkbookProtector.ProtectWorkbook] ERROR de protección: N°" & VBA.Err.Number & " - " & VBA.Err.Description
        Call ErrorLogger2.Log("WorkbookProtector2.ProtectWorkbook", VBA.Err.Description, VBA.Err.Number)
    Else
        Debug.Print "[WorkbookProtector.ProtectWorkbook] OK - Estructura del libro protegida exitosamente"
    End If
    On Error GoTo 0
    Debug.Print "[WorkbookProtector.ProtectWorkbook] FIN"
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
    Dim pwd As String
    pwd = Configuration2.APP_PASSWORD
    Debug.Print "[WorkbookProtector.UnprotectWorkbook] INICIO | Contraseña a usar: '" & pwd & "' (" & Len(pwd) & " caracteres)"
    
    ' On Error Resume Next: si el libro ya está desprotegido o la contraseña no coincide,
    ' Excel lanza error. Lo capturamos aquí para que el fallo no sea silencioso pero
    ' tampoco interrumpa el flujo del llamador con un error no controlado.
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=pwd
    If VBA.Err.Number <> 0 Then
        Debug.Print "[WorkbookProtector.UnprotectWorkbook] ERROR con APP_PASSWORD: N°" & VBA.Err.Number & " - " & VBA.Err.Description
        Debug.Print "[WorkbookProtector.UnprotectWorkbook] ? LA CONTRASEÑA NO COINCIDE con el candado físico del libro"
        Debug.Print "[WorkbookProtector.UnprotectWorkbook] Intentando con contraseñas históricas conocidas..."
        
        ' Lista de todas las contraseñas históricas conocidas en el proyecto:
        ' Inspecciones2026 (OLD_APP), Validacion003 (OLD_AUDIT), 1234/5678 (Fábrica),
        ' supervisor002. (Nueva APP), 2025s (Anteriores), Aseguramiento/Crono (Antiguas ADMIN)
        Dim historicas As Variant
        historicas = Array( _
            "Inspecciones2026", _
            "Validacion003", _
            "1234", _
            "5678", _
            "supervisor002.", _
            "Validacion2025", _
            "Inspecciones2025", _
            "Aseguramiento2026", _
            "CronoAdmin2026*", _
            "" _
        )
       
        Dim i As Long
        Dim encontrada As Boolean
        encontrada = False
        
        For i = LBound(historicas) To UBound(historicas)
            Err.Clear
            ThisWorkbook.Unprotect Password:=historicas(i)
            If Err.Number = 0 Then
                Debug.Print "[WorkbookProtector.UnprotectWorkbook] ? CONTRASEÑA ENCONTRADA: '" & historicas(i) & "' (" & Len(historicas(i)) & " caracteres)"
                Debug.Print "[WorkbookProtector.UnprotectWorkbook]   Esta es la contraseña que tiene físicamente el libro ahora."
                Debug.Print "[WorkbookProtector.UnprotectWorkbook]   Debes ejecutar MigrarContrasenasAlNuevo para cambiarla a: '" & Configuration2.APP_PASSWORD & "'"
                encontrada = True
                Exit For
            End If
        Next i
        
        If Not encontrada Then
            Debug.Print "[WorkbookProtector.UnprotectWorkbook] ? NINGUNA contraseña conocida funcionó."
            Debug.Print "[WorkbookProtector.UnprotectWorkbook]   El libro tiene una contraseña desconocida en su estructura."
            Call ErrorLogger2.Log("WorkbookProtector2.UnprotectWorkbook", "Ninguna contraseña conocida funcionó para desproteger la estructura del libro", 1004)
        End If
        
        On Error GoTo 0
    Else
        Debug.Print "[WorkbookProtector.UnprotectWorkbook] OK - Estructura del libro desprotegida exitosamente con APP_PASSWORD"
    End If
    On Error GoTo 0
    Debug.Print "[WorkbookProtector.UnprotectWorkbook] FIN"
End Sub


