
' ----------------------------------------------------------------------
' Módulo: AdminAccessControl
' Descripción: Proporciona funciones para controlar el acceso de administrador en la aplicación.
'              Permite solicitar y validar la contraseña de administrador, y asigna el rol adecuado
'              al usuario según la autenticación. Incluye manejo de errores, registro de intentos
'              fallidos y bloqueo por fuerza bruta (M05).
' Dependencias:
'   - frmInput: Formulario de entrada de contraseña.
'   - Configuration: Contiene la contraseña de administrador (ADMIN_PASSWORD).
'   - ErrorLogger2: Para registrar errores en el sistema.
'   - AuditLogger2: Para registrar intentos de autenticación.
'   - Variable de módulo m_userRole: Debe estar declarada a nivel de módulo para almacenar el rol actual.
' ----------------------------------------------------------------------
Option Explicit

' ============================================================================
' CONSTANTES DE SEGURIDAD - M05: Bloqueo por intentos fallidos
' ============================================================================
Private Const MAX_INTENTOS_FALLIDOS As Long = 3          ' Máximo de intentos antes del bloqueo
Private Const MINUTOS_BLOQUEO       As Long = 5          ' Duración del bloqueo en minutos
Private ms_ultimoBloqueoHasta        As Date              ' Timestamp hasta cuando dura el bloqueo
Private ml_intentosFallidos          As Long              ' Contador de intentos fallidos en sesión actual

' ============================================================================
' Función privada: EstaBloqueado
' Propósito: Verifica si la autenticación está temporalmente bloqueada por
'            exceso de intentos fallidos.
' Retorno: True si el sistema está bloqueado, False si se permite autenticar.
' ============================================================================
Private Function EstaBloqueado() As Boolean
    If ml_intentosFallidos >= MAX_INTENTOS_FALLIDOS Then
        If Now < ms_ultimoBloqueoHasta Then
            EstaBloqueado = True
        Else
            ' El bloqueo expiró, reiniciar contador
            ml_intentosFallidos = 0
            ms_ultimoBloqueoHasta = 0
            EstaBloqueado = False
        End If
    Else
        EstaBloqueado = False
    End If
End Function

' ============================================================================
' Función privada: AplicarBloqueo
' Propósito: Activa el bloqueo temporal y registra el evento en auditoría.
' ============================================================================
Private Sub AplicarBloqueo()
    ms_ultimoBloqueoHasta = DateAdd("n", MINUTOS_BLOQUEO, Now)
    
    Call AuditLogger2.LogAction( _
        action:="Bloqueo por intentos fallidos", _
        sheetName:="Sistema", _
        dataModified:="Cuenta bloqueada por " & MINUTOS_BLOQUEO & " minutos tras " & _
                       MAX_INTENTOS_FALLIDOS & " intentos fallidos", _
        beforeChange:="Intentos: " & ml_intentosFallidos, _
        afterChange:="Bloqueado hasta: " & Format(ms_ultimoBloqueoHasta, "dd/mm/yyyy hh:nn"), _
        moduleAndSubroutine:="AdminAccessControl2.AplicarBloqueo" _
    )
End Sub

' ============================================================================
' Función pública: ReiniciarContadorIntentos
' Propósito: Permite reiniciar el contador de intentos fallidos (por ejemplo,
'            después de una autenticación exitosa desde otro punto).
' ============================================================================
Public Sub ReiniciarContadorIntentos()
    ml_intentosFallidos = 0
    ms_ultimoBloqueoHasta = 0
End Sub

' ============================================================================
' Función pública: ObtenerTiempoRestanteBloqueo
' Propósito: Devuelve los minutos restantes de bloqueo, para mostrar al usuario.
' Retorno: Long con minutos restantes (0 si no hay bloqueo activo).
' ============================================================================
Public Function ObtenerTiempoRestanteBloqueo() As Long
    If EstaBloqueado Then
        ObtenerTiempoRestanteBloqueo = DateDiff("n", Now, ms_ultimoBloqueoHasta)
    Else
        ObtenerTiempoRestanteBloqueo = 0
    End If
End Function

' ----------------------------------------------------------------------
' Función: CheckAdminAccess
' Propósito: Solicita la contraseña de administrador y valida el acceso.
'            Si la contraseña es correcta, asigna el rol "Admin" al usuario.
'            Si ya es Admin, no solicita la contraseña nuevamente.
'            Incluye bloqueo por intentos fallidos (M05).
' Retorno: Boolean - True si el acceso es concedido, False si es denegado o hay error.
' ----------------------------------------------------------------------
Public Function CheckAdminAccess() As Boolean
    On Error GoTo ErrorHandler

    ' Si el usuario ya es Admin, no es necesario volver a pedir la contraseña.
    If m_userRole = "Admin" Then
        CheckAdminAccess = True
        Exit Function
    End If

    ' M05: Verificar bloqueo por intentos fallidos
    If EstaBloqueado Then
        Dim minsRestantes As Long
        minsRestantes = ObtenerTiempoRestanteBloqueo()
        MsgBox "Acceso temporalmente bloqueado por múltiples intentos fallidos." & vbCrLf & vbCrLf & _
               "Por favor, intente nuevamente en " & minsRestantes & " minuto(s).", _
               vbExclamation, "Acceso Bloqueado"
        CheckAdminAccess = False
        Exit Function
    End If

    ' M05: Informar al usuario cuántos intentos le quedan
    If ml_intentosFallidos > 0 Then
        Dim intentosRestantes As Long
        intentosRestantes = MAX_INTENTOS_FALLIDOS - ml_intentosFallidos
        If intentosRestantes = 1 Then
            MsgBox "ADVERTENCIA: Le queda 1 intento antes de que el acceso se bloquee por " & _
                   MINUTOS_BLOQUEO & " minutos.", vbExclamation, "Último Intento"
        End If
    End If

    Dim frm As New frmInput ' Formulario para ingresar la contraseña
    Dim enteredPassword As String ' Contraseña ingresada por el usuario

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
        
        ' M05: Reiniciar contador de intentos fallidos tras autenticación exitosa
        ml_intentosFallidos = 0
        ms_ultimoBloqueoHasta = 0
        
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
        
        ' ========== FASE 5 (08/06/2026): Desproteger TODAS las hojas de módulo ==========
        ' Al autenticarse como Admin, se desprotegen proactivamente todas las hojas
        ' de módulo (excepto Audit Trail) para que el administrador pueda editar
        ' libremente sin tener que navegar hoja por hoja.
        Call UnprotectAllModuleSheetsForAdmin
        
        CheckAdminAccess = True
        MsgBox "Acceso de administrador concedido. Ahora puede acceder a las funcionalidades restringidas.", vbInformation, "Acceso Concedido"
    Else
        m_userRole = "Usuario" ' Mantener el rol de Usuario si la contraseña es incorrecta.
        
        ' M05: Incrementar contador de intentos fallidos
        ml_intentosFallidos = ml_intentosFallidos + 1
        
        ' FASE 2.3 (10/03/2026): CRÍTICO PARA SEGURIDAD - Auditar intento fallido
        Call AuditLogger2.LogAction( _
            action:="Intento fallido de autenticación Admin", _
            sheetName:="Sistema", _
            dataModified:="Contraseña incorrecta ingresada (intento " & ml_intentosFallidos & _
                          " de " & MAX_INTENTOS_FALLIDOS & ")", _
            beforeChange:="Rol: Usuario", _
            afterChange:="Acceso denegado (rol mantenido: Usuario)", _
            moduleAndSubroutine:="AdminAccessControl2.CheckAdminAccess" _
        )
        
        ' M05: Aplicar bloqueo si se excedió el límite
        If ml_intentosFallidos >= MAX_INTENTOS_FALLIDOS Then
            Call AplicarBloqueo
            MsgBox "Ha excedido el número máximo de intentos permitidos (" & MAX_INTENTOS_FALLIDOS & ")." & vbCrLf & _
                   "El acceso ha sido bloqueado por " & MINUTOS_BLOQUEO & " minutos.", _
                   vbCritical, "Acceso Bloqueado"
        Else
            Dim restantes As Long
            restantes = MAX_INTENTOS_FALLIDOS - ml_intentosFallidos
            MsgBox "Contraseña incorrecta. Se mantendrá el rol de Usuario Estándar." & vbCrLf & _
                   "Intentos restantes: " & restantes, vbCritical, "Acceso Denegado"
        End If
        
        CheckAdminAccess = False
    End If

    Exit Function

ErrorHandler:
    ' Registrar el error y notificar al usuario.
    Call ErrorLogger2.Log("AdminAccessControl2.CheckAdminAccess", VBA.Err.Description, VBA.Err.Number)
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

    ' Si el usuario ya es Admin, no es necesario pedir la contraseña de nuevo.
    If m_userRole = "Admin" Then
        MsgBox "Ya tiene acceso de administrador. No es necesario volver a iniciar sesión.", vbInformation, "Acceso Existente"
        Exit Sub
    End If

    ' M05: Verificar bloqueo por intentos fallidos
    If EstaBloqueado Then
        Dim minsRestantes As Long
        minsRestantes = ObtenerTiempoRestanteBloqueo()
        MsgBox "Acceso temporalmente bloqueado por múltiples intentos fallidos." & vbCrLf & vbCrLf & _
               "Por favor, intente nuevamente en " & minsRestantes & " minuto(s).", _
               vbExclamation, "Acceso Bloqueado"
        Exit Sub
    End If

    Dim frm As New frmInput ' Formulario para ingresar la contraseña
    Dim enteredPassword As String ' Contraseña ingresada por el usuario

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
        
        ' M05: Reiniciar contador de intentos fallidos tras autenticación exitosa
        ml_intentosFallidos = 0
        ms_ultimoBloqueoHasta = 0
        
        ' FASE 2.2 (10/03/2026): Auditar autenticación exitosa de admin
        Call AuditLogger2.LogAction( _
            action:="Cambio de rol a Admin", _
            sheetName:="Sistema", _
            dataModified:="Usuario autenticado como administrador", _
            beforeChange:="Rol: Usuario", _
            afterChange:="Rol: Admin", _
            moduleAndSubroutine:="AdminAccessControl2.AdminAccess" _
        )
        
        ' ========== FASE 5 (08/06/2026): Desproteger TODAS las hojas de módulo ==========
        ' Al autenticarse como Admin, se desprotegen proactivamente todas las hojas
        ' de módulo (excepto Audit Trail) para que el administrador pueda editar
        ' libremente sin tener que navegar hoja por hoja.
        ThisWorkbook.Unprotect Password:=Configuration2.APP_PASSWORD
        Call UnprotectAllModuleSheetsForAdmin
        
        MsgBox "Acceso de administrador concedido. Ahora puede acceder a las funcionalidades restringidas.", vbInformation, "Acceso Concedido"
    Else
        m_userRole = "Usuario" ' Mantener el rol de Usuario si la contraseña es incorrecta.
        
        ' M05: Incrementar contador de intentos fallidos
        ml_intentosFallidos = ml_intentosFallidos + 1
        
        ' FASE 2.3 (10/03/2026): CRÍTICO PARA SEGURIDAD - Auditar intento fallido
        Call AuditLogger2.LogAction( _
            action:="Intento fallido de autenticación Admin", _
            sheetName:="Sistema", _
            dataModified:="Contraseña incorrecta ingresada (intento " & ml_intentosFallidos & _
                          " de " & MAX_INTENTOS_FALLIDOS & ")", _
            beforeChange:="Rol: Usuario", _
            afterChange:="Acceso denegado (rol mantenido: Usuario)", _
            moduleAndSubroutine:="AdminAccessControl2.AdminAccess" _
        )
        
        ' M05: Aplicar bloqueo si se excedió el límite
        If ml_intentosFallidos >= MAX_INTENTOS_FALLIDOS Then
            Call AplicarBloqueo
            MsgBox "Ha excedido el número máximo de intentos permitidos (" & MAX_INTENTOS_FALLIDOS & ")." & vbCrLf & _
                   "El acceso ha sido bloqueado por " & MINUTOS_BLOQUEO & " minutos.", _
                   vbCritical, "Acceso Bloqueado"
        Else
            MsgBox "Contraseña incorrecta. Se mantendrá el rol de Usuario Estándar.", vbCritical, "Acceso Denegado"
        End If
    End If

    Exit Sub

ErrorHandler:
    ' Registrar el error y notificar al usuario.
    Call ErrorLogger2.Log("AdminAccessControl2.AdminAccess", VBA.Err.Description, VBA.Err.Number)
    MsgBox "Ocurrió un error inesperado al procesar el acceso. Consulte el log de errores.", vbCritical, "Error"
End Sub


' ============================================================================
' Subrutina: UnprotectAllModuleSheetsForAdmin
' Propósito: Desprotege proactivamente TODAS las hojas de módulo del sistema
'            (excepto las de Audit Trail) cuando un administrador se autentica.
'            Evita que el admin tenga que navegar hoja por hoja para poder editarlas.
'
'            Para las hojas Audit Trail, aplica protección EXPLÍCITA con
'            AUDIT_PASSWORD para garantizar su inmutabilidad incluso para admin.
'
' Fecha: 08/06/2026 - FASE 5 Activación frmInput Admin
' Lógica:
'   1. Itera sobre GetAllModuleSheetNames() (lista centralizada en Configuration2).
'   2. Para hojas NO Audit Trail: llama ApplyRoleBasedProtection (como Admin → desprotege).
'   3. Para hojas Audit Trail: aplica ProtectSheet con AUDIT_PASSWORD explícitamente.
'   4. Hace Visible la hoja "Menú principal" al final (fallback).
' Dependencias:
'   - Configuration2: GetAllModuleSheetNames, GetAuditTrailSheetNames, APP_PASSWORD, AUDIT_PASSWORD
'   - SheetProtector2: ApplyRoleBasedProtection, ProtectSheet
'   - VariablesGlobales2: m_userRole (debe ser "Admin" al llamar esta sub)
' ============================================================================
Private Sub UnprotectAllModuleSheetsForAdmin()
    On Error GoTo ErrorHandler
    
    Dim allModules As Variant
    allModules = Configuration2.GetAllModuleSheetNames()
    
    Dim auditNames As Variant
    auditNames = Configuration2.GetAuditTrailSheetNames()
    
    Dim moduleName As String
    Dim esAudit As Boolean
    Dim i As Long, j As Long
    Dim ws As Worksheet
    Dim hojasProcesadas As Long
    Dim hojasFallidas As Long
    
    hojasProcesadas = 0
    hojasFallidas = 0
    
    Debug.Print "[AdminAccessControl.UnprotectAllModuleSheetsForAdmin] INICIO — Rol actual: " & m_userRole
    
    For i = LBound(allModules) To UBound(allModules)
        moduleName = CStr(allModules(i))
        
        ' Determinar si esta hoja es Audit Trail
        esAudit = False
        For j = LBound(auditNames) To UBound(auditNames)
            If StrComp(moduleName, CStr(auditNames(j)), vbTextCompare) = 0 Then
                esAudit = True
                Exit For
            End If
        Next j
        
        ' Intentar obtener referencia a la hoja
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(moduleName)
        On Error GoTo ErrorHandler
        
        If ws Is Nothing Then
            Debug.Print "[AdminAccessControl.UnprotectAllModuleSheetsForAdmin] ⚠ Hoja no encontrada: '" & moduleName & "'"
            Set ws = Nothing
            GoTo ContinueLoop
        End If
        
        If esAudit Then
            ' Audit Trail: Proteger EXPLÍCITAMENTE (incluso para admin)
            ' ApplyRoleBasedProtection con Admin dejaría la hoja desprotegida;
            ' usamos ProtectSheet directamente para garantizar inmutabilidad.
            Call SheetProtector2.UnprotectSheet(ws, Configuration2.AUDIT_PASSWORD)
            Call SheetProtector2.ProtectSheet(ws, Configuration2.AUDIT_PASSWORD)
            Debug.Print "[AdminAccessControl.UnprotectAllModuleSheetsForAdmin] 🔒 Audit Trail PROTEGIDA: '" & moduleName & "'"
        Else
            ' Hoja de módulo normal: ApplyRoleBasedProtection con Admin → desprotege
            Call SheetProtector2.ApplyRoleBasedProtection(ws, Configuration2.APP_PASSWORD)
            Debug.Print "[AdminAccessControl.UnprotectAllModuleSheetsForAdmin] 🔓 Desprotegida: '" & moduleName & "'"
        End If
        
        hojasProcesadas = hojasProcesadas + 1
        
ContinueLoop:
        Set ws = Nothing
    Next i
    
    ' Asegurar que "Menú principal" esté visible y desprotegida
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Menú principal")
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVisible
        Call SheetProtector2.ApplyRoleBasedProtection(ws, Configuration2.APP_PASSWORD)
    End If
    On Error GoTo ErrorHandler
    
    Debug.Print "[AdminAccessControl.UnprotectAllModuleSheetsForAdmin] FIN — Procesadas: " & hojasProcesadas & " | Fallidas: " & hojasFallidas
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "[AdminAccessControl.UnprotectAllModuleSheetsForAdmin] ⚠ ERROR: N°" & Err.Number & " - " & Err.Description
    Call ErrorLogger2.Log("AdminAccessControl2.UnprotectAllModuleSheetsForAdmin", VBA.Err.Description, VBA.Err.Number)
    Resume ContinueLoop
End Sub

