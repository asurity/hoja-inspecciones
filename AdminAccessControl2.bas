
' ----------------------------------------------------------------------
' Módulo: AdminAccessControl
' Descripción: Proporciona funciones para controlar el acceso de administrador en la aplicación.
'              Permite solicitar y validar la contraseña de administrador, y asigna el rol adecuado
'              al usuario según la autenticación. Incluye manejo de errores y registro de intentos fallidos.
' Dependencias:
'   - frmInput: Formulario de entrada de contraseña.
'   - Configuration: Contiene la contraseña de administrador (APP_PASSWORD).
'   - ErrorLogger: Para registrar errores en el sistema.
'   - Variable de módulo m_userRole: Debe estar declarada a nivel de módulo para almacenar el rol actual.
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' Función: CheckAdminAccess
' Propósito: Solicita la contraseña de administrador y valida el acceso.
'            Si la contraseña es correcta, asigna el rol "Admin" al usuario.
'            Si ya es Admin, no solicita la contraseña nuevamente.
' Retorno: Boolean - True si el acceso es concedido, False si es denegado o hay error.
' ----------------------------------------------------------------------
Public Function CheckAdminAccess() As Boolean
    On Error GoTo ErrorHandler

    Dim frm As New frmInput ' Formulario para ingresar la contraseña
    Dim enteredPassword As String ' Contraseña ingresada por el usuario

    ' Si el usuario ya es Admin, no es necesario volver a pedir la contraseña.
    If m_userRole = "Admin" Then
        CheckAdminAccess = True
        Exit Function
    End If

    ' Mostrar el formulario de inicio de sesión de forma modal.
    frm.Show vbModal

    ' Obtener la contraseña ingresada en el formulario.
    enteredPassword = frm.txtContrasena.Value

    ' Limpiar la instancia del formulario para liberar memoria.
    Unload frm
    Set frm = Nothing

    ' Validar la contraseña ingresada.
    If enteredPassword = Configuration2.ADMIN_PASSWORD Then
        m_userRole = "Admin" ' Asignar el rol de Admin.
        
        ' FASE 2.2 (10/03/2026): Auditar autenticación exitosa de admin
        Call AuditLogger2.LogAction( _
            action:="Cambio de rol a Admin", _
            sheetName:="Sistema", _
            dataModified:="Usuario autenticado como administrador", _
            beforeChange:="Rol: Usuario", _
            afterChange:="Rol: Admin", _
            moduleAndSubroutine:="AdminAccessControl2.CheckAdminAccess" _
        )
        
        ' (Opcional) Desproteger el libro para permitir cambios estructurales.
        ThisWorkbook.Unprotect Password:=Configuration2.APP_PASSWORD
        CheckAdminAccess = True
        MsgBox "Acceso de administrador concedido. Ahora puede acceder a las funcionalidades restringidas.", vbInformation, "Acceso Concedido"
    Else
        m_userRole = "Usuario" ' Mantener el rol de Usuario si la contraseña es incorrecta.
        
        ' FASE 2.3 (10/03/2026): CRÍTICO PARA SEGURIDAD - Auditar intento fallido
        Call AuditLogger2.LogAction( _
            action:="Intento fallido de autenticación Admin", _
            sheetName:="Sistema", _
            dataModified:="Contraseña incorrecta ingresada", _
            beforeChange:="Rol: Usuario", _
            afterChange:="Acceso denegado (rol mantenido: Usuario)", _
            moduleAndSubroutine:="AdminAccessControl2.CheckAdminAccess" _
        )
        
        CheckAdminAccess = False
        MsgBox "Contraseña incorrecta. Se mantendrá el rol de Usuario Estándar.", vbCritical, "Acceso Denegado"
    End If

    Exit Function

ErrorHandler:
    ' Registrar el error y notificar al usuario.
    Call ErrorLogger.Log("ThisWorkbook.CheckAdminAccess", VBA.Err.Description, VBA.Err.Number)
    MsgBox "Ocurrió un error inesperado al procesar el acceso. Consulte el log de errores.", vbCritical
    CheckAdminAccess = False
End Function

' ----------------------------------------------------------------------
' Subrutina: AdminAccess
' Propósito: Solicita la contraseña de administrador y asigna el rol si es correcta.
'            Si ya es Admin, informa al usuario y no solicita la contraseña.
'            No retorna valor, solo muestra mensajes informativos.
' ----------------------------------------------------------------------
Public Sub AdminAccess()
    On Error GoTo ErrorHandler

    Dim frm As New frmInput ' Formulario para ingresar la contraseña
    Dim enteredPassword As String ' Contraseña ingresada por el usuario

    ' Si el usuario ya es Admin, no es necesario pedir la contraseña de nuevo.
    If m_userRole = "Admin" Then
        MsgBox "Ya tiene acceso de administrador. No es necesario volver a iniciar sesión.", vbInformation, "Acceso Existente"
        Exit Sub
    End If

    ' Mostrar el formulario de inicio de sesión de forma modal.
    frm.Show vbModal

    ' Obtener la contraseña ingresada en el formulario.
    enteredPassword = frm.txtContrasena.Value

    ' Limpiar la instancia del formulario para liberar memoria.
    Unload frm
    Set frm = Nothing

    ' Validar la contraseña ingresada.
    If enteredPassword = Configuration2.ADMIN_PASSWORD Then
        m_userRole = "Admin" ' Asignar el rol de Admin.
        
        ' FASE 2.2 (10/03/2026): Auditar autenticación exitosa de admin
        Call AuditLogger2.LogAction( _
            action:="Cambio de rol a Admin", _
            sheetName:="Sistema", _
            dataModified:="Usuario autenticado como administrador", _
            beforeChange:="Rol: Usuario", _
            afterChange:="Rol: Admin", _
            moduleAndSubroutine:="AdminAccessControl2.AdminAccess" _
        )
        
        ' (Opcional) Desproteger el libro para permitir cambios estructurales.
        'ThisWorkbook.Unprotect PASSWORD:=Configuration2.APP_PASSWORD
        MsgBox "Acceso de administrador concedido. Ahora puede acceder a las funcionalidades restringidas.", vbInformation, "Acceso Concedido"
    Else
        m_userRole = "Usuario" ' Mantener el rol de Usuario si la contraseña es incorrecta.
        
        ' FASE 2.3 (10/03/2026): CRÍTICO PARA SEGURIDAD - Auditar intento fallido
        Call AuditLogger2.LogAction( _
            action:="Intento fallido de autenticación Admin", _
            sheetName:="Sistema", _
            dataModified:="Contraseña incorrecta ingresada", _
            beforeChange:="Rol: Usuario", _
            afterChange:="Acceso denegado (rol mantenido: Usuario)", _
            moduleAndSubroutine:="AdminAccessControl2.AdminAccess" _
        )
        MsgBox "Contraseña incorrecta. Se mantendrá el rol de Usuario Estándar.", vbCritical, "Acceso Denegado"
    End If

    Exit Sub

ErrorHandler:
    ' Registrar el error y notificar al usuario.
    Call ErrorLogger.Log("ThisWorkbook.CheckAdminAccess", VBA.Err.Description, VBA.Err.Number)
    MsgBox "Ocurrió un error inesperado al procesar el acceso. Consulte el log de errores.", vbCritical, "Error"
End Sub