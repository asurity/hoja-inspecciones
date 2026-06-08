
'' ----------------------------------------------------------------------
' Módulo: Hoja1 ("Menú principal")
' Descripción: Controla el comportamiento de la hoja principal del menú.
'              Incluye eventos para la activación y desactivación de la hoja,
'              y la gestión de la protección de la misma.
' Procedimientos públicos:
'   - AbrirGestorCronograma: Punto de entrada para gestión de cronograma.
'                            Requiere contraseña CRONOGRAMA_ADMIN_PASSWORD.
'                            Asignado al botón "Gestor de Cronograma" en la hoja.
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Evento: Worksheet_Activate
' Propósito: Se ejecuta cuando el usuario selecciona la hoja "Menú principal"
'            (ya sea haciendo clic en la pestaña o por código).
'            Oculta todas las hojas visibles excepto esta, y refresca
'            el cronograma resumen.
' Lógica:
'   1. Desprotege la estructura del libro.
'   2. Oculta todas las hojas excepto "Menú principal".
'   3. Re-protege la estructura del libro.
'   4. Restringe la selección de celdas.
'   5. Refresca cronograma resumen de inspecciones.
' ----------------------------------------------------------------------
Private Sub Worksheet_Activate()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    ' Aplicar protección centralizada según el rol del usuario
    ' Admin → desprotegido (puede editar filtros, configuraciones)
    ' Usuario → solo lectura con copiado
    ' Otros → sin selección
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.APP_PASSWORD)
    
    ' Refrescar cronograma resumen
    Call CronogramaResumen.RefrescarResumenCronograma
    
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Call ErrorLogger2.Log("Menú principal.Worksheet_Activate", VBA.Err.Description, VBA.Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_BeforeDoubleClick
' Propósito: Detecta doble clic en una fila de tblResumenCronograma y abre
'            frmSelectorInspeccion con los campos prellenados (Planta,
'            Puesto, Personal e ID Plantilla).
' Parámetros:
'   Target: Celda donde se hizo doble clic
'   Cancel: Cancela el comportamiento por defecto
' Fecha implementación: 23/04/2026
' ----------------------------------------------------------------------
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    Dim tblResumen As ListObject
    Dim filaTabla As ListRow
    Dim iniciales As String
    Dim puesto As String
    Dim idPlantilla As String
    Dim idCronograma As String
    Dim planta As String
    
    ' Obtener referencia a la tabla tblResumenCronograma
    On Error Resume Next
    Set tblResumen = Me.ListObjects(Configuration2.TABLE_RESUMEN_CRONOGRAMA)
    On Error GoTo ErrorHandler
    
    ' Verificar que la tabla existe
    If tblResumen Is Nothing Then Exit Sub
    
    ' Verificar que hay datos en la tabla
    If tblResumen.DataBodyRange Is Nothing Then Exit Sub
    
    ' Verificar que el doble clic fue dentro del DataBodyRange
    If Not Intersect(Target, tblResumen.DataBodyRange) Is Nothing Then
        
        ' Obtener la fila de la tabla donde se hizo doble clic
        Set filaTabla = tblResumen.ListRows(Target.row - tblResumen.Range.row)
        
        ' Leer los valores de la fila
        iniciales = Trim(filaTabla.Range.Cells(1, tblResumen.ListColumns("Iniciales").Index).Value)
        puesto = Trim(filaTabla.Range.Cells(1, tblResumen.ListColumns("Puesto").Index).Value)
        idPlantilla = Trim(filaTabla.Range.Cells(1, tblResumen.ListColumns("ID Plantilla").Index).Value)
        idCronograma = Trim(filaTabla.Range.Cells(1, tblResumen.ListColumns("ID Cronograma").Index).Value)
        
        ' Validar que tenemos los datos mínimos
        If Len(iniciales) = 0 Or Len(puesto) = 0 Or Len(idPlantilla) = 0 Then
            MsgBox "Datos incompletos en la fila seleccionada.", vbExclamation, "Error"
            Cancel = True
            Exit Sub
        End If
        
        ' Obtener la planta del personal usando ChecklistRepository
        planta = ChecklistRepository.ObtenerPlantaPersonal(iniciales)
        
        If Len(planta) = 0 Then
            MsgBox "No se pudo obtener la planta del personal '" & iniciales & "'." & vbCrLf & _
                   "Verifique que el personal esté registrado en tblPersonal.", _
                   vbExclamation, "Error"
            Cancel = True
            Exit Sub
        End If
        
        ' Cancelar el comportamiento por defecto del doble clic
        Cancel = True
        
        ' Crear y configurar el formulario de selección con valores prellenados
        Dim frmSelector As frmSelectorInspeccion
        Set frmSelector = New frmSelectorInspeccion
        
        ' Prellenar los valores iniciales
        frmSelector.PlantaInicial = planta
        frmSelector.PuestoInicial = puesto
        frmSelector.PersonalInicial = iniciales
        frmSelector.IDPlantillaInicial = idPlantilla
        
        ' IMPORTANTE: Aplicar prellenado DESPUÉS de asignar las propiedades
        frmSelector.AplicarPrellenado
        
        ' Mostrar el formulario de forma modal
        frmSelector.Show vbModal
        
        ' Liberar referencia
        Set frmSelector = Nothing
        
        ' Refrescar el cronograma por si hubo cambios
        Call CronogramaResumen.RefrescarResumenCronograma
    End If
    
    Exit Sub
    
ErrorHandler:
    Cancel = True
    Debug.Print "[Menú principal.BeforeDoubleClick] ERROR: " & Err.Description
    Call ErrorLogger2.Log("Menú principal.Worksheet_BeforeDoubleClick", Err.Description, Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_Deactivate
' Propósito: Se ejecuta cuando el usuario sale de la hoja "Menú principal".
'            Restaura la protección de la hoja para evitar modificaciones accidentales.
' Lógica:
'   1. Llama a SheetProtector para proteger la hoja con la contraseña de la aplicación.
' ----------------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.APP_PASSWORD)
End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_Change
' Propósito: Detecta cambios en la celda del filtro de planta (I16) y
'            actualiza automáticamente el cronograma resumen.
' Lógica:
'   1. Verifica si el cambio fue en la celda I16 (filtro de planta)
'   2. Si es así, refresca el cronograma resumen automáticamente
'   3. Esto permite cambio dinámico entre plantas sin necesidad de botón
' ----------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    ' Verificar si el cambio fue en la celda del filtro de planta (I16)
    If Not Intersect(Target, Me.Range(Configuration2.RESUMEN_FILTRO_PLANTA_CELDA)) Is Nothing Then
        ' Refrescar cronograma resumen automáticamente
        Application.EnableEvents = False ' Evitar recursión
        Call CronogramaResumen.RefrescarResumenCronograma
        Application.EnableEvents = True
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Call ErrorLogger2.Log("Menú principal.Worksheet_Change", VBA.Err.Description, VBA.Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Subrutina pública: AbrirGestorCronograma
' Propósito: Punto de entrada para el Gestor de Cronograma.
'            Solicita contraseña, valida con CronogramaGestorService y,
'            si es correcta, abre frmGestorCronograma en modo modal.
'            Al cerrar el form, refresca tblResumenCronograma automáticamente.
' Seguridad: Validación de contraseña antes de mostrar UI (seguridad en capas).
' Asignación: Botón "Gestor de Cronograma" en hoja "Menú principal".
' Fecha creación: 28/04/2026
' ----------------------------------------------------------------------
Public Sub AbrirGestorCronograma()
    On Error GoTo ErrorHandler

    ' ========== M04: Estandarizado con frmInput (ocultación de contraseña) ==========
    ' Ahora usa el mismo formulario modal que AdminAccessControl2,
    ' con la contraseña oculta (caracteres ***) en lugar del InputBox estándar.
    Dim frm As New frmInput
    frm.Show vbModal
    
    Dim contrasena As String
    contrasena = frm.txtContrasena.Value
    
    ' Limpiar la instancia del formulario
    Unload frm
    Set frm = Nothing

    ' Si el usuario no ingresó contraseña, cancelar
    If contrasena = "" Then Exit Sub

    ' Validar contraseña contra CRONOGRAMA_ADMIN_PASSWORD (incluye auditoría interna)
    If Not CronogramaGestorService.ValidarContrasenaGestor(contrasena) Then
        MsgBox "Contraseña incorrecta. Acceso denegado.", vbCritical, "Acceso denegado"
        Exit Sub
    End If

    ' Abrir gestor en modo modal
    Dim frmGestor As frmGestorCronograma
    Set frmGestor = New frmGestorCronograma
    frmGestor.Show vbModal

    Set frmGestor = Nothing

    ' Refrescar resumen al cerrar el gestor (puede haber pausas/reactivaciones)
    Call CronogramaResumen.RefrescarResumenCronograma

    Exit Sub

ErrorHandler:
    Call ErrorLogger2.Log("Menú principal.AbrirGestorCronograma", VBA.Err.Description, VBA.Err.Number)
    MsgBox "Error al abrir el Gestor de Cronograma: " & VBA.Err.Description, vbCritical, "Error"
End Sub