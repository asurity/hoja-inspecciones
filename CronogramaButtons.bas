Attribute VB_Name = "CronogramaButtons"
' ----------------------------------------------------------------------
' Módulo: CronogramaButtons
' Descripción: Vincula los botones de la hoja Cronograma con las funciones
'              del sistema de inspecciones. Proporciona interfaz simple
'              para usuario final.
' Fecha creación: 12/03/2026
' Dependencias: InspectionScheduler
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Subrutina: btnInicializarCronograma_Click
' Propósito: Manejador del botón "INICIALIZAR CRONOGRAMA".
'            Solicita confirmación antes de ejecutar InicializarCronograma.
' ADVERTENCIA: Esta operación limpia tblCronogramaInspecciones y reconstruye
'              desde cero. Usar solo en configuración inicial o reset completo.
' ----------------------------------------------------------------------
Public Sub btnInicializarCronograma_Click()
    On Error GoTo ErrorHandler
    
    Dim respuesta As VbMsgBoxResult
    
    ' Solicitar confirmación
    respuesta = MsgBox( _
        "¿Está seguro de inicializar el cronograma?" & vbCrLf & vbCrLf & _
        "Esta operación:" & vbCrLf & _
        "• Limpiará la tabla de cronograma actual" & vbCrLf & _
        "• Creará registros nuevos para cada Persona × Plantilla válida" & vbCrLf & _
        "• Puede tardar varios segundos" & vbCrLf & vbCrLf & _
        "NOTA: Solo los registros faltantes serán agregados si la tabla ya tiene datos.", _
        vbQuestion + vbYesNo, "Confirmar Inicialización")
    
    If respuesta = vbYes Then
        ' Ejecutar inicialización
        Call InspectionScheduler.InicializarCronograma
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al ejecutar inicialización: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: btnRecalcularCronograma_Click
' Propósito: Manejador del botón "RECALCULAR CRONOGRAMA".
'            Actualiza todos los registros del cronograma con información
'            actual de inspecciones completadas.
' USO NORMAL: Este botón se usa rutinariamente en operación diaria.
' ----------------------------------------------------------------------
Public Sub btnRecalcularCronograma_Click()
    On Error GoTo ErrorHandler
    
    ' Ejecutar recálculo (sin confirmación, es operación segura)
    Call InspectionScheduler.RecalcularCronograma
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al recalcular cronograma: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: btnNuevaInspeccion_Click
' Propósito: Manejador del botón "NUEVA INSPECCIÓN".
'            Abre el formulario de checklist virtual para la persona
'            seleccionada en tblCronogramaInspecciones.
' ----------------------------------------------------------------------
Public Sub btnNuevaInspeccion_Click()
    On Error GoTo ErrorHandler
    
    ' Obtener fila seleccionada en tblCronogramaInspecciones
    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    
    If tblCronograma.DataBodyRange Is Nothing Then
        MsgBox "No hay registros en el cronograma.", vbInformation, "Sin datos"
        Exit Sub
    End If
    
    ' Verificar que la celda activa está dentro de la tabla
    If Intersect(ActiveCell, tblCronograma.DataBodyRange) Is Nothing Then
        MsgBox "Seleccione una fila en la tabla de cronograma antes de crear una inspección.", _
               vbInformation, "Seleccione un registro"
        Exit Sub
    End If
    
    ' Obtener datos de la fila seleccionada
    Dim selectedRow As Long
    selectedRow = ActiveCell.Row - tblCronograma.DataBodyRange.Row + 1
    
    Dim iniciales As String
    Dim idPlantilla As String
    Dim puesto As String
    Dim idCronograma As String
    
    iniciales = Trim(tblCronograma.DataBodyRange.Cells(selectedRow, tblCronograma.ListColumns("Iniciales personal").Index).Value)
    idPlantilla = Trim(tblCronograma.DataBodyRange.Cells(selectedRow, tblCronograma.ListColumns("ID Plantilla").Index).Value)
    puesto = Trim(tblCronograma.DataBodyRange.Cells(selectedRow, tblCronograma.ListColumns("Puesto").Index).Value)
    idCronograma = Trim(tblCronograma.DataBodyRange.Cells(selectedRow, tblCronograma.ListColumns("ID Cronograma").Index).Value)
    
    If Len(iniciales) = 0 Or Len(idPlantilla) = 0 Then
        MsgBox "La fila seleccionada no tiene datos válidos.", vbExclamation, "Datos incompletos"
        Exit Sub
    End If
    
    ' Abrir checklist virtual
    Call ChecklistOrchestrator.AbrirChecklistVirtual(iniciales, idPlantilla, puesto, idCronograma)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al crear inspección: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: btnNuevaInspeccionDirecta_Click
' Propósito: Abre un formulario de selección para crear una inspección SIN
'            necesidad de que exista cronograma previo.
'            Permite elegir: Puesto → Personal → se carga Plantilla automáticamente
' Flujo:
'   1. Abre frmSelectorInspeccion modal
'   2. El usuario selecciona un Puesto (ComboBox)
'   3. Se carga Personal disponible en cboPersonal automáticamente
'   4. El usuario selecciona Personal → se carga Plantilla automáticamente
'   5. Al aceptar, abre frmChecklistVirtual con los datos seleccionados
' Uso: Botón en Menú Principal para crear inspecciones sin cronograma
' ----------------------------------------------------------------------
Public Sub btnNuevaInspeccionDirecta_Click()
    On Error GoTo ErrorHandler
    
    ' Abre el formulario de selección
    Call AbrirSelectorInspeccion
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al crear inspección: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: AbrirSelectorInspeccion
' Propósito: Abre el formulario visual de selección (frmSelectorInspeccion)
'            y luego inicia el checklist virtual si el usuario confirma.
' ----------------------------------------------------------------------
Public Sub AbrirSelectorInspeccion()
    On Error GoTo ErrorHandler
    
    ' Crear instancia del formulario selector
    Dim frmSelector As New frmSelectorInspeccion
    
    ' Mostrar modal (espera hasta que el usuario acepte o cancele)
    frmSelector.Show vbModal
    
    ' Verificar si el usuario canceló
    If frmSelector.Cancelado Then
        Unload frmSelector
        Exit Sub
    End If
    
    ' Obtener datos seleccionados
    Dim puesto As String
    Dim iniciales As String
    Dim idPlantilla As String
    Dim planta As String
    
    puesto = frmSelector.PuestoSeleccionado
    iniciales = frmSelector.PersonalSeleccionado
    idPlantilla = frmSelector.IDPlantilla
    planta = frmSelector.Planta
    
    ' Liberar formulario
    Unload frmSelector
    
    ' Validar datos
    If Len(idPlantilla) = 0 Or Len(iniciales) = 0 Then
        MsgBox "Datos incompletos. No se puede iniciar la inspección.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Abrir checklist virtual
    ' idCronograma = "" porque es una inspección ad-hoc (no viene de tabla cronograma)
    Call ChecklistOrchestrator.AbrirChecklistVirtual(iniciales, idPlantilla, puesto, "")
    
    Exit Sub
ErrorHandler:
    MsgBox "Error en selector: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: btnFiltrarVencidos_Click
' Propósito: Filtra la tabla de cronograma para mostrar solo registros vencidos
'            o próximos a vencer.
' ----------------------------------------------------------------------
Public Sub btnFiltrarVencidos_Click()
    On Error GoTo ErrorHandler
    
    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    
    ' Limpiar filtros existentes
    If tblCronograma.AutoFilter.FilterMode Then
        tblCronograma.AutoFilter.ShowAllData
    End If
    
    ' Aplicar filtro a columna "Estado cronograma"
    Dim colEstado As Long
    colEstado = tblCronograma.ListColumns("Estado cronograma").Index
    
    ' Filtrar por: Vencido o Por vencer
    tblCronograma.Range.AutoFilter Field:=colEstado, _
        Criteria1:=Configuration2.ESTADO_VENCIDO, _
        Operator:=xlOr, _
        Criteria2:=Configuration2.ESTADO_POR_VENCER
    
    MsgBox "Filtro aplicado: mostrando solo inspecciones vencidas o por vencer.", vbInformation, "Filtro Aplicado"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al filtrar: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: btnMostrarTodos_Click
' Propósito: Quita todos los filtros de la tabla de cronograma.
' ----------------------------------------------------------------------
Public Sub btnMostrarTodos_Click()
    On Error GoTo ErrorHandler
    
    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    
    ' Limpiar todos los filtros
    If tblCronograma.AutoFilter.FilterMode Then
        tblCronograma.AutoFilter.ShowAllData
    End If
    
    MsgBox "Filtros removidos: mostrando todos los registros.", vbInformation, "Filtros Removidos"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al quitar filtros: " & Err.Description, vbCritical, "Error"
End Sub

'' ----------------------------------------------------------------------
' Subrutina: btnFiltrarPersona_Click
' Propósito: Muestra un InputBox para filtrar por iniciales de persona.
' ----------------------------------------------------------------------
Public Sub btnFiltrarPersona_Click()
    On Error GoTo ErrorHandler
    
    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    Dim iniciales As String
    
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    
    ' Solicitar iniciales
    iniciales = InputBox("Ingrese las iniciales del personal a filtrar:", "Filtrar por Persona", "")
    
    If Len(Trim(iniciales)) = 0 Then
        MsgBox "No se ingresaron iniciales. Operación cancelada.", vbExclamation, "Cancelado"
        Exit Sub
    End If
    
    ' Limpiar filtros existentes
    If tblCronograma.AutoFilter.FilterMode Then
        tblCronograma.AutoFilter.ShowAllData
    End If
    
    ' Aplicar filtro a columna "Iniciales personal"
    Dim colIniciales As Long
    colIniciales = tblCronograma.ListColumns("Iniciales personal").Index
    
    tblCronograma.Range.AutoFilter Field:=colIniciales, Criteria1:="=" & UCase(Trim(iniciales))
    
    ' Contar resultados visibles
    Dim registrosVisibles As Long
    If Not tblCronograma.DataBodyRange Is Nothing Then
        On Error Resume Next
        registrosVisibles = tblCronograma.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count
        On Error GoTo ErrorHandler
    End If
    
    MsgBox "Filtro aplicado para: " & UCase(Trim(iniciales)) & vbCrLf & _
           "Registros encontrados: " & registrosVisibles, vbInformation, "Filtro Aplicado"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al filtrar persona: " & Err.Description, vbCritical, "Error"
End Sub
