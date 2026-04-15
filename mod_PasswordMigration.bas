' ============================================================================
' Módulo: mod_PasswordMigration
' Descripción: Macro de migración ONE-TIME para cambiar contraseñas físicas del libro.
'              Desbloquea todo con las contraseñas ANTERIORES y vuelve a bloquear
'              con las nuevas contraseñas definidas en Configuration2.bas.
'
' INSTRUCCIONES DE USO:
'   1. Importar este módulo al proyecto VBA del archivo TH-HC-002 V00.xlsm
'   2. Importar el módulo Configuration2.bas ACTUALIZADO (con las nuevas contraseñas)
'   3. Ejecutar la macro "MigrarContrasenasAlNuevo" UNA SOLA VEZ
'   4. Verificar que el libro funcione correctamente
'   5. Opcional: Eliminar este módulo del proyecto VBA (ya no es necesario)
'
' FECHA DE MIGRACIÓN: 10-03-2026 (FASE 1 Mejoras Finales)
' ============================================================================

Option Explicit

' ============================================================================
' CONTRASEÑAS ANTERIORES (hardcodeadas aquí para poder desbloquear)
' No modificar estos valores — reflejan el estado ACTUAL del sistema
' Actualizadas: 10/03/2026 - Segunda migración (desde contraseñas FASE 1)
' ============================================================================
Private Const OLD_APP_PASSWORD   As String = "AppTH-HC-002_Sec2026!"
Private Const OLD_AUDIT_PASSWORD As String = "AuditTH-HC-002_Immut!"

' ============================================================================
' Subrutina principal de migración
' ============================================================================
Public Sub MigrarContrasenasAlNuevo()
    On Error GoTo ErrorHandler

    Dim ws          As Worksheet
    Dim countOK     As Long
    Dim countFail   As Long
    Dim failList    As String

    countOK = 0
    countFail = 0
    failList = ""

    ' --- Confirmación antes de ejecutar ---
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("Este proceso cambiará las contraseñas de protección de todas las hojas " & _
                       "y la estructura del libro TH-HC-002." & vbCrLf & vbCrLf & _
                       "SEGUNDA MIGRACIÓN (FASE 1B):" & vbCrLf & _
                       "  • APP_PASSWORD:   AppTH-HC-002_Sec2026!  →  Detecciones004." & vbCrLf & _
                       "  • ADMIN_PASSWORD: AdminTH-HC-002_Auth!   →  AdministradorDetecciones" & vbCrLf & _
                       "  • AUDIT_PASSWORD: AuditTH-HC-002_Immut!  →  Validaciones003" & vbCrLf & vbCrLf & _
                       "⚠️ IMPORTANTE: Asegúrate de tener un backup antes de continuar." & vbCrLf & vbCrLf & _
                       "¿Deseas continuar?", _
                       vbQuestion + vbYesNo, "Segunda Migración de Contraseñas TH-HC-002")

    If respuesta = vbNo Then
        MsgBox "Migración cancelada. No se realizaron cambios.", vbInformation, "Cancelado"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' -----------------------------------------------------------------------
    ' PASO 1: Desproteger estructura del libro con contraseña ANTERIOR
    ' -----------------------------------------------------------------------
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=OLD_APP_PASSWORD
    If Err.Number <> 0 Then
        MsgBox "No se pudo desproteger la estructura del libro con la contraseña anterior." & vbCrLf & _
               "Verifica que la contraseña anterior sea '1234'." & vbCrLf & vbCrLf & _
               "Error: " & Err.Description, vbCritical, "Error en Migración"
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler

    ' -----------------------------------------------------------------------
    ' PASO 2: Migrar cada hoja — desbloquear con contraseña vieja,
    '         re-bloquear con la nueva (lee Configuration.bas actualizado)
    ' -----------------------------------------------------------------------
    For Each ws In ThisWorkbook.Worksheets

        On Error Resume Next

        ' Determinar si es hoja Audit Trail
        Dim esAudit As Boolean
        esAudit = (ws.Name Like "Audit Trail*")

        ' Desproteger con contraseña ANTERIOR correspondiente.
        ' NOTA: Debido al bug original en SheetService (coincidencia exacta en lugar de Like),
        '       solo "Audit Trail" (sin número) estaba bloqueada con OLD_AUDIT_PASSWORD.
        '       "Audit Trail 2", etc. estaban bloqueadas con OLD_APP_PASSWORD.
        '       Para manejar cualquier estado mixto, se intentan ambas contraseñas.
        If esAudit Then
            ' Intentar primero con OLD_AUDIT_PASSWORD
            ws.Unprotect Password:=OLD_AUDIT_PASSWORD
            If Err.Number <> 0 Then
                Err.Clear
                ' Si falla, intentar con OLD_APP_PASSWORD (bug del SheetService anterior)
                ws.Unprotect Password:=OLD_APP_PASSWORD
            End If
        Else
            ws.Unprotect Password:=OLD_APP_PASSWORD
        End If

        If Err.Number <> 0 Then
            countFail = countFail + 1
            failList = failList & "  - " & ws.Name & " (" & Err.Description & ")" & vbCrLf
            Err.Clear
        Else
            ' Re-proteger con contraseña NUEVA desde Configuration2
            If esAudit Then
                Call SheetProtector2.ProtectSheet(ws, Configuration2.AUDIT_PASSWORD)
            Else
                Call SheetProtector2.ProtectSheet(ws, Configuration2.APP_PASSWORD)
            End If
            countOK = countOK + 1
        End If

        On Error GoTo ErrorHandler

    Next ws

    ' -----------------------------------------------------------------------
    ' PASO 3: Re-proteger estructura del libro con contraseña NUEVA
    ' -----------------------------------------------------------------------
    Call WorkbookProtector2.ProtectWorkbook

    ' -----------------------------------------------------------------------
    ' PASO 4: Guardar el libro para persistir los nuevos locks
    ' -----------------------------------------------------------------------
    ThisWorkbook.Save

    ' -----------------------------------------------------------------------
    ' Reporte final
    ' -----------------------------------------------------------------------
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim mensaje As String
    mensaje = "Migración completada." & vbCrLf & vbCrLf & _
              "  ✔ Hojas migradas correctamente: " & countOK & vbCrLf

    If countFail > 0 Then
        mensaje = mensaje & "  ✘ Hojas con error: " & countFail & vbCrLf & vbCrLf & _
                  "Detalle de errores:" & vbCrLf & failList & vbCrLf & _
                  "Revisa esas hojas manualmente."
        MsgBox mensaje, vbExclamation, "Migración con advertencias"
    Else
        mensaje = mensaje & vbCrLf & _
                  "El libro fue guardado con las nuevas contraseñas." & vbCrLf & vbCrLf & _
                  "Puedes eliminar el módulo 'mod_PasswordMigration' del proyecto VBA " & _
                  "ya que no se necesitará más."
        MsgBox mensaje, vbInformation, "Migración Exitosa"
    End If

    Exit Sub

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error inesperado durante la migración." & vbCrLf & _
           "Número: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, vbCritical, "Error en Migración"
End Sub

' ============================================================================
' Subrutina: VerificarContrasenasActuales
' Descripción: Prueba de solo-lectura que verifica si las contraseñas en
'              Configuration.bas coinciden con los locks físicos actuales del libro.
'              NO modifica nada — desprotege temporalmente y vuelve a proteger.
'
' USO: Alt+F8 → VerificarContrasenasActuales → Ejecutar
' ============================================================================
Public Sub VerificarContrasenasActuales()

    Dim reporte     As String
    Dim okApp       As Boolean
    Dim okWb        As Boolean
    Dim okAudit     As Boolean
    Dim wsApp       As Worksheet
    Dim ws          As Worksheet

    reporte = "══════════════════════════════════════" & vbCrLf & _
              "  VERIFICACIÓN DE CONTRASEÑAS" & vbCrLf & _
              "  TH-HC-002 — " & Format(Now, "dd/mm/yyyy hh:mm") & vbCrLf & _
              "══════════════════════════════════════" & vbCrLf & vbCrLf

    Application.ScreenUpdating = False

    ' ------------------------------------------------------------------
    ' 1. APP_PASSWORD — estructura del libro (WorkbookProtector)
    ' ------------------------------------------------------------------
    On Error Resume Next
    Err.Clear
    ThisWorkbook.Unprotect Password:=Configuration2.APP_PASSWORD
    okWb = (Err.Number = 0)
    Err.Clear
    ' Re-proteger inmediatamente sin importar el resultado
    ThisWorkbook.Protect Password:=Configuration2.APP_PASSWORD, Structure:=True, Windows:=False
    On Error GoTo 0

    reporte = reporte & "1. APP_PASSWORD (estructura del libro)" & vbCrLf
    If okWb Then
        reporte = reporte & "   [OK] La contraseña coincide con el lock del libro." & vbCrLf
    Else
        reporte = reporte & "   [FALLO] La contraseña NO coincide. " & _
                            "El libro sigue bloqueado con una clave distinta." & vbCrLf
    End If
    reporte = reporte & vbCrLf

    ' ------------------------------------------------------------------
    ' 2. APP_PASSWORD — hojas normales (primera hoja no-Audit que aparezca)
    ' ------------------------------------------------------------------
    Set wsApp = Nothing
    For Each ws In ThisWorkbook.Worksheets
        If Not (ws.Name Like "Audit Trail*") Then
            Set wsApp = ws
            Exit For
        End If
    Next ws

    reporte = reporte & "2. APP_PASSWORD (hojas normales)" & vbCrLf
    If wsApp Is Nothing Then
        reporte = reporte & "   [AVISO] No se encontró ninguna hoja normal para probar." & vbCrLf
    Else
        On Error Resume Next
        Err.Clear
        wsApp.Unprotect Password:=Configuration2.APP_PASSWORD
        okApp = (Err.Number = 0)
        Err.Clear
        Call SheetProtector2.ProtectSheet(wsApp, Configuration2.APP_PASSWORD)
        On Error GoTo 0

        reporte = reporte & "   Hoja probada: """ & wsApp.Name & """" & vbCrLf
        If okApp Then
            reporte = reporte & "   [OK] La contraseña coincide con el lock de la hoja." & vbCrLf
        Else
            reporte = reporte & "   [FALLO] La contraseña NO coincide. " & _
                                "La hoja sigue bloqueada con una clave distinta." & vbCrLf
        End If
    End If
    reporte = reporte & vbCrLf

    ' ------------------------------------------------------------------
    ' 3. AUDIT_PASSWORD — hojas Audit Trail (se prueban TODAS)
    ' ------------------------------------------------------------------
    Dim countAuditOK   As Long
    Dim countAuditFail As Long
    Dim auditFailList  As String
    countAuditOK = 0
    countAuditFail = 0
    auditFailList = ""

    reporte = reporte & "3. AUDIT_PASSWORD (hojas Audit Trail — se prueban todas)" & vbCrLf

    Dim foundAudit As Boolean
    foundAudit = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Audit Trail*" Then
            foundAudit = True
            On Error Resume Next
            Err.Clear
            ws.Unprotect Password:=Configuration2.AUDIT_PASSWORD
            Dim auditOK As Boolean
            auditOK = (Err.Number = 0)
            Err.Clear
            Call SheetProtector2.ProtectSheet(ws, Configuration2.AUDIT_PASSWORD)
            On Error GoTo 0

            If auditOK Then
                countAuditOK = countAuditOK + 1
                reporte = reporte & "   [OK]    " & ws.Name & vbCrLf
            Else
                countAuditFail = countAuditFail + 1
                auditFailList = auditFailList & ws.Name & ", "
                reporte = reporte & "   [FALLO] " & ws.Name & " — bloqueada con clave distinta" & vbCrLf
            End If
        End If
    Next ws

    If Not foundAudit Then
        reporte = reporte & "   [AVISO] No se encontró ninguna hoja Audit Trail." & vbCrLf
        okAudit = True ' No hay hojas que verificar, no es error
    Else
        okAudit = (countAuditFail = 0)
        reporte = reporte & "   Total: " & (countAuditOK + countAuditFail) & " hojas | " & _
                            countAuditOK & " OK | " & countAuditFail & " con fallo" & vbCrLf
    End If
    reporte = reporte & vbCrLf

    ' ------------------------------------------------------------------
    ' 4. ADMIN_PASSWORD — solo comparación de texto (sin lock físico)
    ' ------------------------------------------------------------------
    reporte = reporte & "4. ADMIN_PASSWORD (autenticación de usuario)" & vbCrLf
    If Len(Configuration2.ADMIN_PASSWORD) > 0 Then
        reporte = reporte & "   [OK] Definida en Configuration2 — " & _
                            Len(Configuration2.ADMIN_PASSWORD) & " caracteres." & vbCrLf & _
                            "   (No tiene lock físico: se valida por comparación de texto.)" & vbCrLf
    Else
        reporte = reporte & "   [FALLO] La constante ADMIN_PASSWORD está vacía." & vbCrLf
    End If
    reporte = reporte & vbCrLf

    ' ------------------------------------------------------------------
    ' Resumen final
    ' ------------------------------------------------------------------
    Dim todoBien As Boolean
    todoBien = okWb And okApp And okAudit And Len(Configuration2.ADMIN_PASSWORD) > 0

    reporte = reporte & "══════════════════════════════════════" & vbCrLf
    If todoBien Then
        reporte = reporte & "  RESULTADO: Todo correcto." & vbCrLf & _
                            "  El sistema está listo para usar."
        Application.ScreenUpdating = True
        MsgBox reporte, vbInformation, "Verificación de Contraseñas"
    Else
        reporte = reporte & "  RESULTADO: Se detectaron problemas." & vbCrLf & _
                            "  Ejecuta 'MigrarContrasenasAlNuevo' si aún no lo hiciste."
        Application.ScreenUpdating = True
        MsgBox reporte, vbExclamation, "Verificación de Contraseñas"
    End If

End Sub

' ============================================================================
' Subrutina: CorregirHojasAuditSecundarias
' Descripción: Corrección manual para hojas Audit Trail 2-5 que fallaron
'              en la migración inicial. Intenta múltiples métodos para
'              desproteger y re-proteger con la nueva contraseña.
'
' USO: Alt+F8 → CorregirHojasAuditSecundarias → Ejecutar
' ============================================================================
Public Sub CorregirHojasAuditSecundarias()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim hojasAudit As Variant
    Dim i As Long
    Dim countOK As Long
    Dim countFail As Long
    Dim reporte As String
    Dim intentos() As String
    
    ' Lista de contraseñas a intentar (de más probable a menos)
    ReDim intentos(0 To 5)
    intentos(0) = OLD_APP_PASSWORD      ' Contraseña ACTUAL de APP (más probable para hojas secundarias)
    intentos(1) = OLD_AUDIT_PASSWORD    ' Contraseña ACTUAL de AUDIT
    intentos(2) = ""                    ' Sin contraseña
    intentos(3) = "1234"                ' Contraseña original de APP
    intentos(4) = "5678"                ' Contraseña original de AUDIT
    intentos(5) = "AppTH-HC-002_Sec2026!" ' Por si quedó con esta explícitamente
    
    hojasAudit = Array("Audit trail 2", "Audit trail 3", "Audit trail 4", "Audit trail 5")
    
    countOK = 0
    countFail = 0
    reporte = "══════════════════════════════════════" & vbCrLf & _
              "  CORRECCIÓN HOJAS AUDIT SECUNDARIAS" & vbCrLf & _
              "  TH-HC-002 — " & Format(Now, "dd/mm/yyyy hh:mm") & vbCrLf & _
              "══════════════════════════════════════" & vbCrLf & vbCrLf
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Desproteger estructura del libro temporalmente
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=Configuration2.APP_PASSWORD
    On Error GoTo ErrorHandler
    
    For i = LBound(hojasAudit) To UBound(hojasAudit)
        Dim nombreHoja As String
        nombreHoja = hojasAudit(i)
        
        ' Buscar la hoja
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(nombreHoja)
        On Error GoTo ErrorHandler
        
        If ws Is Nothing Then
            reporte = reporte & "[SKIP] " & nombreHoja & " — No existe en el libro" & vbCrLf
        Else
            ' Hacer visible temporalmente para poder trabajar con ella
            Dim visibilidadOriginal As XlSheetVisibility
            visibilidadOriginal = ws.Visible
            ws.Visible = xlSheetVisible
            
            ' Intentar desproteger con cada contraseña
            Dim desprotegida As Boolean
            Dim j As Long
            desprotegida = False
            
            For j = LBound(intentos) To UBound(intentos)
                On Error Resume Next
                Err.Clear
                
                If intentos(j) = "" Then
                    ' Intentar sin contraseña
                    ws.Unprotect
                Else
                    ws.Unprotect Password:=intentos(j)
                End If
                
                If Err.Number = 0 Then
                    desprotegida = True
                    On Error GoTo ErrorHandler
                    Exit For
                End If
            Next j
            
            On Error GoTo ErrorHandler
            
            If desprotegida Then
                ' Re-proteger con la nueva contraseña
                Call SheetProtector2.ProtectSheet(ws, Configuration2.AUDIT_PASSWORD)
                
                ' Restaurar visibilidad
                ws.Visible = visibilidadOriginal
                
                countOK = countOK + 1
                reporte = reporte & "[OK] " & nombreHoja & " — Migrada correctamente" & vbCrLf
            Else
                ' No se pudo desproteger
                ws.Visible = visibilidadOriginal
                countFail = countFail + 1
                reporte = reporte & "[FAIL] " & nombreHoja & " — No se pudo desproteger con ninguna contraseña conocida" & vbCrLf
            End If
        End If
    Next i
    
    ' Re-proteger estructura del libro
    Call WorkbookProtector2.ProtectWorkbook
    
    ' Guardar cambios
    ThisWorkbook.Save
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Mostrar reporte
    reporte = reporte & vbCrLf & "══════════════════════════════════════" & vbCrLf
    reporte = reporte & "Resultado: " & countOK & " corregidas | " & countFail & " con error" & vbCrLf
    
    If countFail = 0 Then
        reporte = reporte & vbCrLf & "✓ Todas las hojas Audit secundarias fueron corregidas." & vbCrLf & _
                           "  Ejecuta 'VerificarContrasenasActuales' para confirmar."
        MsgBox reporte, vbInformation, "Corrección Exitosa"
    Else
        reporte = reporte & vbCrLf & "⚠ Algunas hojas no pudieron corregirse." & vbCrLf & _
                           "  Contacta al desarrollador para investigar."
        MsgBox reporte, vbExclamation, "Corrección con advertencias"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error inesperado durante la corrección." & vbCrLf & _
           "Número: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, vbCritical, "Error en Corrección"
End Sub
