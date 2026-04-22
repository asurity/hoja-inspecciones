
' ----------------------------------------------------------------------
' Módulo: UserManager
' Descripción: Servicio para gestionar la información y visualización del usuario activo.
'              Permite mostrar el nombre del usuario en la hoja principal del libro.
' Última actualización: 12/03/2026 - Refactorizado para usar constantes centralizadas.
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Subrutina: DisplayUserName
' Propósito: Muestra el nombre del usuario activo en la hoja principal del menú.
' Lógica:
'   1. Obtiene el nombre del usuario desde la aplicación.
'   2. Desprotege la hoja principal (Configuration2.MAIN_MENU_SHEET).
'   3. Escribe el nombre del usuario en la celda configurada (Configuration2.USER_DISPLAY_CELL).
'   4. Protege nuevamente la hoja para mantener la seguridad.
'   5. Maneja errores mostrando un mensaje y (opcionalmente) registrando en el log.
' Actualización: 12/03/2026 - Refactorizado para usar constantes centralizadas
'' ----------------------------------------------------------------------
Public Sub DisplayUserName()
    Dim ws As Worksheet
    Dim userName As String
    On Error GoTo ErrorHandler
    userName = Application.userName
    Set ws = ThisWorkbook.Sheets(Configuration2.MAIN_MENU_SHEET)
    Call SheetProtector2.UnprotectSheet(ws, Configuration2.APP_PASSWORD)
    ws.Range(Configuration2.USER_DISPLAY_CELL).Value = "Usuario: " & userName
    Call SheetProtector2.ProtectSheet(ws, Configuration2.APP_PASSWORD)
    Exit Sub
ErrorHandler:
    MsgBox "Ocurrió un error al mostrar el nombre de usuario.", vbExclamation, "Error del Sistema"
    ' Para registrar el error en el log, descomentar la siguiente línea:
    ' Call ErrorLogger.Log("UserManager.DisplayUserName", VBA.Err.Description, VBA.Err.Number)
End Sub