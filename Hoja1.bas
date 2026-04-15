
'' ----------------------------------------------------------------------
' Módulo: Hoja1 ("Menú principal")
' Descripción: Controla el comportamiento de la hoja principal del menú.
'              Incluye eventos para la activación y desactivación de la hoja,
'              y la gestión de la protección de la misma.
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
    Call WorkbookProtector2.UnprotectWorkbook
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> Me.Name Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    Application.DisplayAlerts = True
    Call WorkbookProtector2.ProtectWorkbook
    Me.EnableSelection = xlNoSelection
    
    ' Refrescar cronograma resumen
    Call CronogramaResumen.RefrescarResumenCronograma
    
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Call WorkbookProtector2.ProtectWorkbook
    Call ErrorLogger2.Log("Menú principal.Worksheet_Activate", VBA.Err.Description, VBA.Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_BeforeDoubleClick
' Propósito: DESHABILITADO - El flujo de inspecciones ahora usa botón con
'            selector de Puesto → Personal → Plantilla.
' ----------------------------------------------------------------------
' Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'     ' FLUJO DESCARTADO: Ver CronogramaButtons.btnNuevaInspeccionDirecta_Click
' End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_Deactivate
' Propósito: Se ejecuta cuando el usuario sale de la hoja "Menú principal".
'            Restaura la protección de la hoja para evitar modificaciones accidentales.
' Lógica:
'   1. Llama a SheetProtector para proteger la hoja con la contraseña de la aplicación.
' ----------------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    Call SheetProtector2.ProtectSheet(Me, Configuration2.APP_PASSWORD)
End Sub