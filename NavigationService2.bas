
' ----------------------------------------------------------------------
' Módulo: NavigationService
' Descripción: Proporciona subrutinas para navegar de manera segura entre las diferentes
'              hojas de cálculo de la aplicación. Centraliza la lógica de selección de hojas
'              y la integra con el sistema de protección para un flujo de trabajo controlado.
' Dependencias:
'   - SheetProtector2: Para aplicar protección según el rol del usuario.
'   - Configuration2: Para las contraseñas de cada hoja.
'   - ErrorLogger2: Para registrar errores de navegación.
' ----------------------------------------------------------------------
Option Explicit


' ----------------------------------------------------------------------
' Subrutina: NavigateToSheet
' Propósito: Navega a una hoja específica y aplica la protección correcta según el rol:
'            - Admin   → sin protección (puede editar todo).
'            - Usuario → protección de solo lectura (puede ver y copiar, no editar).
' Parámetros:
'   - targetSheetName: Nombre de la hoja de destino.
' ----------------------------------------------------------------------
Public Sub NavigateToSheet(ByVal targetSheetName As String)
    Debug.Print "═══════════════════════════════════════════════════════════"
    Debug.Print "[NavigationService.NavigateToSheet] INICIO"
    Debug.Print "[NavigationService] Hoja destino: '" & targetSheetName & "'"
    Debug.Print "[NavigationService] Rol actual: '" & m_userRole & "'"
    Debug.Print "───────────────────────────────────────────────────────────"
    
    On Error GoTo ErrorHandler
    
    ' ========== BLOQUE DE SUPRESIÓN DE PARPADEO ==========
    ' Se manejan aquí para que HideAndProtectAllSheetsExcept no los restaure
    ' hasta después del .Select, evitando parpadeo visual al cambiar de hoja.
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ' Verificar que la hoja destino existe
    Debug.Print "[NavigationService] Verificando existencia de hoja '" & targetSheetName & "'..."
    Dim wsExist As Worksheet
    On Error Resume Next
    Set wsExist = ThisWorkbook.Sheets(targetSheetName)
    On Error GoTo ErrorHandler
    If wsExist Is Nothing Then
        Debug.Print "[NavigationService] ⚠ ERROR: La hoja '" & targetSheetName & "' NO EXISTE en el libro"
        Debug.Print "[NavigationService]   Hojas disponibles en el libro:"
        Dim tmpWs As Worksheet
        For Each tmpWs In ThisWorkbook.Sheets
            Debug.Print "[NavigationService]   - '" & tmpWs.Name & "' (Visible: " & tmpWs.Visible & ")"
        Next tmpWs
        MsgBox "Error: La hoja '" & targetSheetName & "' no existe.", vbCritical, "Error de Navegación"
        Exit Sub
    End If
    Debug.Print "[NavigationService]   ✔ Hoja '" & targetSheetName & "' existe en el libro"
    
    ' Llamar al servicio de ocultación/protección
    Debug.Print "[NavigationService] Llamando a SheetService2.HideAndProtectAllSheetsExcept('" & targetSheetName & "')..."
    Call SheetService2.HideAndProtectAllSheetsExcept(targetSheetName)
    Debug.Print "[NavigationService]   ✔ SheetService completado exitosamente"
    
    ' Seleccionar la hoja destino (ScreenUpdating/EnableEvents ya están deshabilitados arriba)
    Debug.Print "[NavigationService] Seleccionando hoja '" & targetSheetName & "'..."
    On Error Resume Next
    ThisWorkbook.Sheets(targetSheetName).Select
    If Err.Number <> 0 Then
        Debug.Print "[NavigationService] ⚠ ERROR al seleccionar hoja: N°" & Err.Number & " - " & Err.Description
    Else
        Debug.Print "[NavigationService]   ✔ Hoja '" & targetSheetName & "' seleccionada"
    End If
    On Error GoTo ErrorHandler
    
    ' ========== RESTAURAR ESTADOS DE APLICACIÓN ==========
    ' Se restauran aquí, DESPUÉS del .Select, para evitar parpadeo visual.
    ' HideAndProtectAllSheetsExcept ya no los restaura (lo hace el llamador).
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    ' Registrar navegación en Audit Trail
    Debug.Print "[NavigationService] Registrando navegación en Audit Trail..."
    On Error Resume Next
    Call AuditLogger2.LogAction( _
        action:="Navegación", _
        sheetName:=targetSheetName, _
        dataModified:="Acceso a módulo", _
        beforeChange:="N/A", _
        afterChange:="Usuario accedió a: " & targetSheetName, _
        moduleAndSubroutine:="NavigationService2.NavigateToSheet" _
    )
    If Err.Number <> 0 Then
        Debug.Print "[NavigationService] ⚠ AVISO: No se pudo registrar en Audit Trail: N°" & Err.Number & " - " & Err.Description
    End If
    On Error GoTo ErrorHandler
    
    Debug.Print "[NavigationService.NavigateToSheet] FIN EXITOSO"
    Debug.Print "═══════════════════════════════════════════════════════════"
    
    Exit Sub
ErrorHandler:
    ' Restaurar estados de aplicación por si ocurrió un error antes de la restauración normal
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Debug.Print "[NavigationService.NavigateToSheet] *** ERROR CAPTURADO ***"
    Debug.Print "[NavigationService]   Hoja destino: '" & targetSheetName & "'"
    Debug.Print "[NavigationService]   N° de error: " & VBA.Err.Number
    Debug.Print "[NavigationService]   Descripción: " & VBA.Err.Description
    Debug.Print "[NavigationService]   Fuente: " & VBA.Err.Source
    Debug.Print "[NavigationService] *******************************"
    Call ErrorLogger2.Log("NavigationService.NavigateToSheet", "No se pudo navegar a la hoja: " & targetSheetName, VBA.Err.Number)
    MsgBox "Error: No se pudo navegar a la hoja '" & targetSheetName & "'.", vbCritical, "Error de Navegación"
End Sub


' ----------------------------------------------------------------------
' Subrutinas de Navegación Específicas
' Propósito: Proporcionan puntos de entrada públicos y amigables para la navegación,
'            permitiendo ir a hojas específicas mediante nombres claros.
'            Se pueden agregar validaciones de permisos si es necesario.
' ----------------------------------------------------------------------

Public Sub NavigateToDetecciones()
    Debug.Print "[NavigationService.NavigateToDetecciones] Solicitando navegación a 'Detecciones'"
    Call NavigateToSheet("Detecciones")
End Sub

Public Sub NavigateToDashboard()
    Debug.Print "[NavigationService.NavigateToDashboard] Solicitando navegación a 'Dashboard'"
    Call NavigateToSheet("Dashboard")
End Sub

Public Sub NavigateToObservaciones()
    Debug.Print "[NavigationService.NavigateToObservaciones] Solicitando navegación a 'Observaciones'"
    Call NavigateToSheet("Observaciones")
End Sub

Public Sub NavigateToRechazo()
    Debug.Print "[NavigationService.NavigateToRechazo] Solicitando navegación a 'Análisis Rechazo'"
    Call NavigateToSheet("Análisis Rechazo")
End Sub

Public Sub NavigateToDesvio()
    Debug.Print "[NavigationService.NavigateToDesvio] Solicitando navegación a 'Análisis Desvío'"
    Call NavigateToSheet("Análisis Desvío")
End Sub

Public Sub NavigateToConfiguracion()
    Debug.Print "[NavigationService.NavigateToConfiguracion] Solicitando navegación a 'Configuración'"
    Call NavigateToSheet("Configuración")
End Sub

' ----------------------------------------------------------------------
' Subrutina: NavigateToPersonalProduccion
' Propósito: Navega a la hoja "Personal" que contiene la tabla tblPersonal
'            con todo el personal de producción del sistema.
' Características:
'   - Tabla tblPersonal con iniciales, nombre, puesto, planta
'   - Personal activo/inactivo
'   - Asignación de plantillas de inspección
' ----------------------------------------------------------------------
Public Sub NavigateToPersonalProduccion()
    Debug.Print "[NavigationService.NavigateToPersonalProduccion] Solicitando navegación a '" & Configuration2.SHEET_PERSONAL & "'"
    Call NavigateToSheet(Configuration2.SHEET_PERSONAL)
End Sub

' ----------------------------------------------------------------------
' Subrutina: NavigateToAseguramientoCalidad
' Propósito: Navega a la hoja "Aseguramiento de calidad" que contiene
'            la tabla tblAseguramientoCalidad con el personal de QA.
' Características:
'   - Tabla tblAseguramientoCalidad
'   - Personal de aseguramiento de calidad
'   - Evaluadores autorizados para inspecciones
' ----------------------------------------------------------------------
Public Sub NavigateToAseguramientoCalidad()
    Debug.Print "[NavigationService.NavigateToAseguramientoCalidad] Solicitando navegación a '" & Configuration2.SHEET_ASEGURAMIENTO & "'"
    Call NavigateToSheet(Configuration2.SHEET_ASEGURAMIENTO)
End Sub

Public Sub NavigateToTecControlProceso()
    Debug.Print "[NavigationService.NavigateToTecControlProceso] Solicitando navegación a 'TecControlProceso'"
    Call NavigateToSheet("TecControlProceso")
End Sub

Public Sub NavigateToAuditTrail()
    Debug.Print "═══════════════════════════════════════════════════════════"
    Debug.Print "[NavigationService.NavigateToAuditTrail] INICIO"
    Debug.Print "[NavigationService] Rol actual: '" & m_userRole & "'"
    Debug.Print "───────────────────────────────────────────────────────────"
    
    On Error GoTo ErrorHandler
    
    ' Registrar navegación en Audit Trail (antes de mostrar las hojas para evitar recursión)
    Debug.Print "[NavigationService] Registrando navegación al grupo Audit Trail..."
    On Error Resume Next
    Call AuditLogger2.LogAction( _
        action:="Navegación", _
        sheetName:="Audit Trail (Grupo)", _
        dataModified:="Acceso a módulo de auditoría", _
        beforeChange:="N/A", _
        afterChange:="Usuario accedió al grupo de hojas Audit Trail", _
        moduleAndSubroutine:="NavigationService2.NavigateToAuditTrail" _
    )
    If Err.Number <> 0 Then
        Debug.Print "[NavigationService] ⚠ AVISO: No se pudo registrar en Audit Trail: N°" & Err.Number & " - " & Err.Description
    End If
    On Error GoTo ErrorHandler
    
    ' Muestra las hojas de Audit Trail simultáneamente y oculta el resto.
    Debug.Print "[NavigationService] Llamando a SheetService2.ShowAuditTrailGroup..."
    Call SheetService2.ShowAuditTrailGroup
    Debug.Print "[NavigationService]   ✔ ShowAuditTrailGroup completado exitosamente"
    
    Debug.Print "[NavigationService.NavigateToAuditTrail] FIN EXITOSO"
    Debug.Print "═══════════════════════════════════════════════════════════"
    Exit Sub
ErrorHandler:
    Debug.Print "[NavigationService.NavigateToAuditTrail] *** ERROR: " & Err.Number & " - " & Err.Description
    Call ErrorLogger2.Log("NavigationService.NavigateToAuditTrail", "No se pudo navegar al grupo Audit trail", VBA.Err.Number)
    MsgBox "Error: No se pudo navegar a Audit trail.", vbCritical, "Error de Navegación"
End Sub

Public Sub NavigateToControlDeCambios()
    Debug.Print "[NavigationService.NavigateToControlDeCambios] Solicitando navegación a 'Control de cambios'"
    Call NavigateToSheet("Control de cambios")
End Sub

Public Sub NavigateToMenu()
    ' Esta navegación está permitida para todos los usuarios, sin necesidad de validación.
    Debug.Print "[NavigationService.NavigateToMenu] Solicitando navegación a 'Menú principal'"
    Call NavigateToSheet("Menú principal")
End Sub

' ----------------------------------------------------------------------
' Subrutina: NavigateToChecklistVirtual
' Propósito: Abre el formulario selector para iniciar una inspección
'            con el checklist virtual. El selector permite elegir
'            Puesto → Personal → Plantilla y luego abre frmChecklistVirtual.
' Flujo:
'   1. Abre frmSelectorInspeccion (modal)
'   2. Usuario selecciona Personal y Plantilla
'   3. Al aceptar, el formulario llama a ChecklistOrchestrator.AbrirChecklistVirtual
' ----------------------------------------------------------------------
Public Sub NavigateToChecklistVirtual()
    On Error GoTo ErrorHandler
    
    Debug.Print "[NavigationService.NavigateToChecklistVirtual] INICIO - Abriendo frmSelectorInspeccion"
    
    ' Registrar apertura del selector en Audit Trail
    Call AuditLogger2.LogAction( _
        action:="Apertura Selector Inspección", _
        sheetName:="Formulario", _
        dataModified:="frmSelectorInspeccion", _
        beforeChange:="N/A", _
        afterChange:="Usuario solicitó nueva inspección", _
        moduleAndSubroutine:="NavigationService2.NavigateToChecklistVirtual" _
    )
    
    ' Abrir formulario selector modal
    Dim frmSelector As frmSelectorInspeccion
    Set frmSelector = New frmSelectorInspeccion
    frmSelector.Show vbModal
    
    ' Limpiar referencia
    Set frmSelector = Nothing
    
    Debug.Print "[NavigationService.NavigateToChecklistVirtual] FIN - Selector cerrado"
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("NavigationService.NavigateToChecklistVirtual", _
                         "Error al abrir selector de inspección: " & Err.Description, Err.Number)
    MsgBox "Error: No se pudo abrir el selector de inspección.", vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------
' Subrutina: NavigateToResultados
' Propósito: Navega a la hoja "Historico" que contiene los resultados
'            de inspecciones completadas (tabla tblInspecciones).
'            Permite consultar historial y generar certificados PDF
'            haciendo doble clic en una inspección.
' Características:
'   - Tabla con todas las inspecciones guardadas
'   - Doble clic en fila → Genera certificado PDF
'   - Permite filtros y búsquedas de inspecciones pasadas
' ----------------------------------------------------------------------
Public Sub NavigateToResultados()
    Debug.Print "[NavigationService.NavigateToResultados] Solicitando navegación a '" & Configuration2.SHEET_HISTORICO & "'"
    Call NavigateToSheet(Configuration2.SHEET_HISTORICO)
End Sub

' ----------------------------------------------------------------------
' Subrutina: NavigateToConfiguracionChecklist
' Propósito: Navega a la hoja "Checklist" que contiene la configuración
'            de plantillas de inspección y sus ítems.
' Características:
'   - Configuración de plantillas
'   - Definición de ítems de inspección
'   - Administración de criterios de evaluación
' ----------------------------------------------------------------------
Public Sub NavigateToConfiguracionChecklist()
    Debug.Print "[NavigationService.NavigateToConfiguracionChecklist] Solicitando navegación a 'Checklist'"
    Call NavigateToSheet("Checklist")
End Sub

' ----------------------------------------------------------------------
' Subrutina: NavigateToCronograma
' Propósito: Navega a la hoja "Cronograma" que contiene el cronograma
'            de inspecciones programadas por persona y puesto.
' Características:
'   - Tabla tblCronogramaInspecciones con todas las inspecciones programadas
'   - Información de frecuencias de inspección por puesto
'   - Estado de cumplimiento del cronograma
'   - Próximas inspecciones pendientes
'   - Inspecciones vencidas o atrasadas
' ----------------------------------------------------------------------
Public Sub NavigateToCronograma()
    Debug.Print "[NavigationService.NavigateToCronograma] Solicitando navegación a '" & Configuration2.SHEET_CRONOGRAMA & "'"
    Call NavigateToSheet(Configuration2.SHEET_CRONOGRAMA)
End Sub

' ----------------------------------------------------------------------
' Subrutina: NavigateToPlantillaCertificado
' Propósito: Navega a la hoja "Plantilla Certificado" que contiene
'            la plantilla base para la generación de certificados PDF.
' Características:
'   - Plantilla de diseño del certificado
'   - Estructura de 4 columnas (A-D)
'   - Secciones: Categoría, Datos, AP, Resultados, Preguntas, Feedback, Observaciones
'   - Hoja normalmente oculta (xlSheetVeryHidden)
'   - Uso exclusivo para generación automática de PDFs
' Advertencia: Esta hoja se puebla dinámicamente durante la generación de PDFs.
'              No se recomienda modificar manualmente sin consultar la documentación.
' ----------------------------------------------------------------------
Public Sub NavigateToPlantillaCertificado()
    Debug.Print "[NavigationService.NavigateToPlantillaCertificado] Solicitando navegación a '" & Configuration2.SHEET_PLANTILLA_CERTIFICADO & "'"
    Call NavigateToSheet(Configuration2.SHEET_PLANTILLA_CERTIFICADO)
End Sub
