
'' ----------------------------------------------------------------------
' Módulo: WorkbookProtector
' Descripción: Gestiona la protección y desprotección de la estructura del libro.
'              Permite separar la lógica de protección del libro de la de las hojas individuales,
'              facilitando el mantenimiento y la seguridad global del archivo.
'' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Subrutina: ProtectWorkbook
' Propósito: Protege la estructura del libro (no las hojas individuales) usando la contraseña definida.
' Lógica:
'   1. Protege la estructura del libro para evitar que se agreguen, eliminen o muevan hojas.
'   2. Maneja errores y registra en el log si ocurre alguno.
'' ----------------------------------------------------------------------
Public Sub ProtectWorkbook()
    ' On Error Resume Next: si el libro ya está protegido, Excel lanza Error 1004.
    ' Ignoramos ese caso en silencio para evitar interrumpir el flujo del llamador.
    On Error Resume Next
    ThisWorkbook.Protect Password:=Configuration2.APP_PASSWORD, Structure:=True, Windows:=False
    If VBA.Err.Number <> 0 Then
        Call ErrorLogger2.Log("WorkbookProtector2.ProtectWorkbook", VBA.Err.Description, VBA.Err.Number)
    End If
    On Error GoTo 0
End Sub

'' ----------------------------------------------------------------------
' Subrutina: UnprotectWorkbook
' Propósito: Desprotege la estructura del libro usando la contraseña definida.
' Lógica:
'   1. Quita la protección de la estructura del libro para permitir modificaciones globales.
'   2. Maneja errores y registra en el log si ocurre alguno.
'' ----------------------------------------------------------------------
Public Sub UnprotectWorkbook()
    ' On Error Resume Next: si el libro ya está desprotegido o la contraseña no coincide,
    ' Excel lanza error. Lo capturamos aquí para que el fallo no sea silencioso pero
    ' tampoco interrumpa el flujo del llamador con un error no controlado.
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=Configuration2.APP_PASSWORD
    If VBA.Err.Number <> 0 Then
        Call ErrorLogger2.Log("WorkbookProtector2.UnprotectWorkbook", VBA.Err.Description, VBA.Err.Number)
    End If
    On Error GoTo 0
End Sub