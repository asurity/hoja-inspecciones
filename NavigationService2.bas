
'' ----------------------------------------------------------------------
' Módulo: NavigationService
' Descripción: Proporciona subrutinas para navegar de manera segura entre las diferentes
'              hojas de cálculo de la aplicación. Centraliza la lógica de selección de hojas
'              y la integra con el sistema de protección para un flujo de trabajo controlado.
' Dependencias:
'   - SheetProtector2: Para aplicar protección según el rol del usuario.
'   - Configuration2: Para las contraseñas de cada hoja.
'   - ErrorLogger2: Para registrar errores de navegación.
'' ----------------------------------------------------------------------
Option Explicit


'' ----------------------------------------------------------------------
' Subrutina: NavigateToSheet
' Propósito: Navega a una hoja específica y aplica la protección correcta según el rol:
'            - Admin   ? sin protección (puede editar todo).
'            - Usuario ? protección de solo lectura (puede ver y copiar, no editar).
' Parámetros:
'   - targetSheetName: Nombre de la hoja de destino.
'' ----------------------------------------------------------------------
Public Sub NavigateToSheet(ByVal targetSheetName As String)
    On Error GoTo ErrorHandler
    ' HideAndProtectAllSheetsExcept oculta todo lo demás, hace visible la hoja
    ' destino y aplica la protección correcta según el rol activo. También
    ' envuelve el cambio de visibilidad con Unprotect/ProtectWorkbook para
    ' evitar Error 1004 cuando la estructura del libro está protegida.
    Call SheetService2.HideAndProtectAllSheetsExcept(targetSheetName)
    ThisWorkbook.Sheets(targetSheetName).Select
    
    ' Registrar navegación en Audit Trail
    Call AuditLogger2.LogAction( _
        action:="Navegación", _
        sheetName:=targetSheetName, _
        dataModified:="Acceso a módulo", _
        beforeChange:="N/A", _
        afterChange:="Usuario accedió a: " & targetSheetName, _
        moduleAndSubroutine:="NavigationService2.NavigateToSheet" _
    )
    
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("NavigationService.NavigateToSheet", "No se pudo navegar a la hoja: " & targetSheetName, VBA.Err.Number)
    MsgBox "Error: No se pudo navegar a la hoja '" & targetSheetName & "'.", vbCritical, "Error de Navegación"
End Sub


'' ----------------------------------------------------------------------
' Subrutinas de Navegación Específicas
' Propósito: Proporcionan puntos de entrada públicos y amigables para la navegación,
'            permitiendo ir a hojas específicas mediante nombres claros.
'            Se pueden agregar validaciones de permisos si es necesario.
'' ----------------------------------------------------------------------

Public Sub NavigateToDetecciones()
    Call NavigateToSheet("Detecciones")
End Sub

Public Sub NavigateToDashboard()
    Call NavigateToSheet("Dashboard")
End Sub

Public Sub NavigateToObservaciones()
    Call NavigateToSheet("Observaciones")
End Sub

Public Sub NavigateToRechazo()
    Call NavigateToSheet("Análisis Rechazo")
End Sub

Public Sub NavigateToDesvio()
    Call NavigateToSheet("Análisis Desvío")
End Sub

Public Sub NavigateToConfiguracion()
    Call NavigateToSheet("Configuración")
End Sub

'' ----------------------------------------------------------------------
' Subrutina: NavigateToPersonalProduccion
' Propósito: Navega a la hoja "Personal" que contiene la tabla tblPersonal
'            con todo el personal de producción del sistema.
' Características:
'   - Tabla tblPersonal con iniciales, nombre, puesto, planta
'   - Personal activo/inactivo
'   - Asignación de plantillas de inspección
' ----------------------------------------------------------------------
Public Sub NavigateToPersonalProduccion()
    Call NavigateToSheet(Configuration2.SHEET_PERSONAL)
End Sub

'' ----------------------------------------------------------------------
' Subrutina: NavigateToAseguramientoCalidad
' Propósito: Navega a la hoja "Aseguramiento de calidad" que contiene
'            la tabla tblAseguramientoCalidad con el personal de QA.
' Características:
'   - Tabla tblAseguramientoCalidad
'   - Personal de aseguramiento de calidad
'   - Evaluadores autorizados para inspecciones
' ----------------------------------------------------------------------
Public Sub NavigateToAseguramientoCalidad()
    Call NavigateToSheet(Configuration2.SHEET_ASEGURAMIENTO)
End Sub

Public Sub NavigateToTecControlProceso()
    Call NavigateToSheet("TecControlProceso")
End Sub

Public Sub NavigateToAuditTrail()
    On Error GoTo ErrorHandler
    
    ' Registrar navegación en Audit Trail (antes de mostrar las hojas para evitar recursión)
    Call AuditLogger2.LogAction( _
        action:="Navegación", _
        sheetName:="Audit Trail (Grupo)", _
        dataModified:="Acceso a módulo de auditoría", _
        beforeChange:="N/A", _
        afterChange:="Usuario accedió al grupo de hojas Audit Trail", _
        moduleAndSubroutine:="NavigationService2.NavigateToAuditTrail" _
    )
    
    ' Muestra las 5 hojas de Audit Trail simultáneamente y oculta el resto.
    Call SheetService2.ShowAuditTrailGroup
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("NavigationService.NavigateToAuditTrail", "No se pudo navegar al grupo Audit trail", VBA.Err.Number)
    MsgBox "Error: No se pudo navegar a Audit trail.", vbCritical, "Error de Navegación"
End Sub

Public Sub NavigateToControlDeCambios()
    Call NavigateToSheet("Control de cambios")
End Sub

Public Sub NavigateToMenu()
    ' Esta navegación está permitida para todos los usuarios, sin necesidad de validación.
    Call NavigateToSheet("Menú principal")
End Sub

'' ----------------------------------------------------------------------
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
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("NavigationService.NavigateToChecklistVirtual", _
                         "Error al abrir selector de inspección: " & Err.Description, Err.Number)
    MsgBox "Error: No se pudo abrir el selector de inspección.", vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
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
    Call NavigateToSheet(Configuration2.SHEET_HISTORICO)
End Sub

'' ----------------------------------------------------------------------
' Subrutina: NavigateToConfiguracionChecklist
' Propósito: Navega a la hoja "Checklist" que contiene la configuración
'            de plantillas de inspección y sus ítems.
' Características:
'   - Configuración de plantillas
'   - Definición de ítems de inspección
'   - Administración de criterios de evaluación
' ----------------------------------------------------------------------
Public Sub NavigateToConfiguracionChecklist()
    Call NavigateToSheet("Checklist")
End Sub

'' ----------------------------------------------------------------------
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
    Call NavigateToSheet(Configuration2.SHEET_CRONOGRAMA)
End Sub

'' ----------------------------------------------------------------------
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
    Call NavigateToSheet(Configuration2.SHEET_PLANTILLA_CERTIFICADO)
End Sub