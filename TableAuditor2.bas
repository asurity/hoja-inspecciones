'' ----------------------------------------------------------------------
' Módulo: TableAuditor
' Descripción: Servicio para auditar y registrar cambios en tablas de Excel.
'              Agrupa los cambios de celdas múltiples en una sola entrada de auditoría,
'              marca celdas vacías explícitamente y maneja el límite de caracteres de Excel.
'' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Subrutina: AuditTableChanges
' Propósito: Audita los cambios realizados en una o varias celdas de una tabla específica,
'            agrupando los cambios y registrando los valores antes y después de la modificación.
' Argumentos:
'   - changedSheet: Hoja donde ocurrió el cambio.
'   - changedRange: Rango de celdas modificadas.
'   - tablesToAudit: Array con los nombres de las tablas a auditar.
'   - beforeChange: Valores previos al cambio (puede ser array o valor simple).
' Lógica:
'   1. Identifica la tabla afectada dentro de las tablas a auditar.
'   2. Recorre cada celda modificada y compara el valor anterior y el actual.
'   3. Si hay cambios, construye cadenas de texto agrupando los detalles.
'   4. Si alguna cadena excede el límite de Excel, la trunca y lo indica.
'   5. Registra la acción en el log de auditoría solo si hubo cambios.
'   6. Maneja errores y restablece el estado de la aplicación si ocurre una excepción.
'' ----------------------------------------------------------------------
Public Sub AuditTableChanges(ByVal changedSheet As Worksheet, ByVal changedRange As Range, ByVal tablesToAudit As Variant, ByVal beforeChange As Variant)
    On Error GoTo ErrorHandler
    Dim changedTable As ListObject
    Dim tableName As Variant
    Dim cell As Range
    Dim auditIndex As Long
    Dim cellBeforeValue As Variant
    Dim cellAfterValue As Variant
    Dim changedCellsString As String ' Celdas modificadas (agrupadas)
    Dim beforeChangesString As String ' Valores antes del cambio
    Dim afterChangesString As String  ' Valores después del cambio
    Dim auditorModule As String ' Identificador del módulo que realiza la acción
    auditorModule = "Acción realizada por el usuario"
    Const EXCEL_CELL_LIMIT As Long = 32767
    ' 1. Identificar la tabla que se ha modificado.
    For Each tableName In tablesToAudit
        On Error Resume Next
        Set changedTable = changedSheet.ListObjects(tableName)
        On Error GoTo 0
        If Not changedTable Is Nothing Then
            If Not Intersect(changedRange, changedTable.DataBodyRange) Is Nothing Then
                Exit For
            Else
                Set changedTable = Nothing
            End If
        End If
    Next tableName
    If changedTable Is Nothing Then Exit Sub
    ' 2. Recorrer cada celda en el rango modificado para construir las cadenas de texto.
    auditIndex = 0
    For Each cell In Intersect(changedRange, changedTable.DataBodyRange).Cells
        ' 3. Obtener el valor original (antes del cambio).
        If IsArray(beforeChange) Then
            If auditIndex < UBound(beforeChange, 1) * UBound(beforeChange, 2) Then
                cellBeforeValue = beforeChange(cell.Row - changedRange.Row + 1, cell.Column - changedRange.Column + 1)
            Else
                cellBeforeValue = "Error: Valor no encontrado"
            End If
        Else
            cellBeforeValue = beforeChange
        End If
        ' 4. Obtener el valor actual y normalizarlo.
        cellAfterValue = cell.Value
        ' 5. Comparar los valores y construir las cadenas si hay un cambio.
        If Not CStr(cellBeforeValue) = CStr(cellAfterValue) Then
            changedCellsString = changedCellsString & changedTable.Name & " - Celda: " & cell.Address(False, False) & vbCrLf
            beforeChangesString = beforeChangesString & cell.Address(False, False) & ": " & IIf(CStr(cellBeforeValue) = "", "vacío", CStr(cellBeforeValue)) & vbCrLf
            afterChangesString = afterChangesString & cell.Address(False, False) & ": " & IIf(CStr(cellAfterValue) = "", "vacío", CStr(cellAfterValue)) & vbCrLf
        End If
        auditIndex = auditIndex + 1
    Next cell
    ' 6. Realizar una única llamada al AuditLogger si se detectaron cambios.
    If Len(changedCellsString) > 0 Then
        Dim maxLength As Long
        Dim truncationMessage As String
        maxLength = Len(changedCellsString)
        If Len(beforeChangesString) > maxLength Then maxLength = Len(beforeChangesString)
        If Len(afterChangesString) > maxLength Then maxLength = Len(afterChangesString)
        ' Si alguna cadena excede el límite de Excel, truncalas de manera consistente.
        If maxLength > EXCEL_CELL_LIMIT Then
            truncationMessage = " (Truncado)"
            changedCellsString = Left(changedCellsString, EXCEL_CELL_LIMIT - Len(truncationMessage)) & Trim(truncationMessage)
            beforeChangesString = Left(beforeChangesString, EXCEL_CELL_LIMIT - Len(truncationMessage)) & Trim(truncationMessage)
            afterChangesString = Left(afterChangesString, EXCEL_CELL_LIMIT - Len(truncationMessage)) & Trim(truncationMessage)
        End If
        Call AuditLogger2.LogAction( _
            action:="Modificación en tabla", _
            sheetName:=changedSheet.Name, _
            dataModified:=Trim(changedCellsString), _
            beforeChange:=Trim(beforeChangesString), _
            afterChange:=Trim(afterChangesString), _
            moduleAndSubroutine:=auditorModule _
        )
    End If
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("TableAuditor2.AuditTableChanges", VBA.Err.Description, VBA.Err.Number)
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub