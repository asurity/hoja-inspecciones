' ----------------------------------------------------------------------
' Módulo: TestDataGenerator
' Propósito: Genera datos de prueba para testing del sistema de inspecciones
' Fecha: 14/04/2026
' NOTA: Este módulo es TEMPORAL - eliminar antes de producción
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Subrutina: GenerarDatosPrueba
' Propósito: Crea un registro completo de prueba en tblCronogramaInspecciones
'            para que aparezca en tblResumenCronograma y se pueda hacer doble clic.
' ----------------------------------------------------------------------
Public Sub GenerarDatosPrueba()
    On Error GoTo ErrorHandler
    
    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    Dim nuevaFila As ListRow
    
    ' Desproteger si es necesario
    Application.ScreenUpdating = False
    
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    
    ' Verificar si ya existe un registro de prueba para ACF/Operador
    Dim filaExistente As ListRow
    Dim yaExiste As Boolean: yaExiste = False
    
    If Not tblCronograma.DataBodyRange Is Nothing Then
        For Each filaExistente In tblCronograma.ListRows
            Dim inics As String
            Dim psto As String
            inics = Trim(filaExistente.Range.Cells(1, tblCronograma.ListColumns("Iniciales").Index).Value)
            psto = Trim(filaExistente.Range.Cells(1, tblCronograma.ListColumns("Puesto").Index).Value)
            
            If inics = "ACF" And psto = "Operador" Then
                yaExiste = True
                Debug.Print "Ya existe registro de prueba para ACF/Operador"
                Exit For
            End If
        Next filaExistente
    End If
    
    ' Si no existe, crear nuevo registro
    If Not yaExiste Then
        Set nuevaFila = tblCronograma.ListRows.Add
        
        With nuevaFila.Range
            ' ID Cronograma - generado automáticamente
            .Cells(1, tblCronograma.ListColumns("ID Cronograma").Index).Value = GenerarID()
            
            ' Iniciales personal
            .Cells(1, tblCronograma.ListColumns("Iniciales personal").Index).Value = "ACF"
            
            ' Puesto
            .Cells(1, tblCronograma.ListColumns("Puesto").Index).Value = "Operador"
            
            ' ID Plantilla (la que ya existe)
            .Cells(1, tblCronograma.ListColumns("ID Plantilla").Index).Value = "fxEJV01C-xC6PKG6C-pVOj2dMa"
            
            ' Frecuencia meses
            .Cells(1, tblCronograma.ListColumns("Frecuencia meses").Index).Value = 1
            
            ' Fecha última inspección - dejar en blanco para "Nunca inspeccionado"
            .Cells(1, tblCronograma.ListColumns("Fecha ultima inspeccion").Index).Value = ""
            
            ' Fecha próxima inspección - dejar en blanco
            .Cells(1, tblCronograma.ListColumns("Fecha proxima inspeccion").Index).Value = ""
            
            ' Estado cronograma - se calculará automáticamente
            .Cells(1, tblCronograma.ListColumns("Estado cronograma").Index).Value = Configuration2.ESTADO_NUNCA_INSPECCIONADO
            
            ' Planta personal
            .Cells(1, tblCronograma.ListColumns("Planta personal").Index).Value = "Therapia iv Santiago"
            
            ' Activo
            .Cells(1, tblCronograma.ListColumns("Activo").Index).Value = "Si"
            
            ' Fecha de creación
            .Cells(1, tblCronograma.ListColumns("Fecha de creacion").Index).Value = Now
        End With
        
        Debug.Print "Registro de prueba creado en tblCronogramaInspecciones"
        Debug.Print "  Iniciales: ACF"
        Debug.Print "  Puesto: Operador"
        Debug.Print "  Plantilla: fxEJV01C-xC6PKG6C-pVOj2dMa"
    End If
    
    ' Actualizar tabla resumen
    Debug.Print "Actualizando tblResumenCronograma..."
    Call CronogramaResumen.RefrescarResumenCronograma
    Debug.Print "tblResumenCronograma actualizada OK"
    
    Application.ScreenUpdating = True
    
    MsgBox "Datos de prueba generados correctamente." & vbCrLf & vbCrLf & _
           "Ahora puedes hacer doble clic en la fila de ACF/Operador en el Menú principal.", _
           vbInformation, "Datos de prueba"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error al generar datos de prueba: " & Err.Description, vbCritical, "Error"
    Debug.Print "ERROR en GenerarDatosPrueba: " & Err.Number & " - " & Err.Description
End Sub

'' ----------------------------------------------------------------------
' Función: GenerarID
' Propósito: Genera un ID único en formato XXXXXXXX-XXXXXXXX-XXXXXXXXXX
' ----------------------------------------------------------------------
Private Function GenerarID() As String
    Dim parte1 As String
    Dim parte2 As String
    Dim parte3 As String
    
    Randomize
    
    parte1 = GenerarSegmento(8)
    parte2 = GenerarSegmento(8)
    parte3 = GenerarSegmento(10)
    
    GenerarID = parte1 & "-" & parte2 & "-" & parte3
End Function

Private Function GenerarSegmento(ByVal longitud As Long) As String
    Const CHARS As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    Dim i As Long
    Dim resultado As String
    
    resultado = ""
    For i = 1 To longitud
        Dim pos As Long
        pos = Int((Len(CHARS) * Rnd) + 1)
        resultado = resultado & Mid(CHARS, pos, 1)
    Next i
    
    GenerarSegmento = resultado
End Function
