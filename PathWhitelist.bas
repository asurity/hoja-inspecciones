' ----------------------------------------------------------------------
' Módulo: PathWhitelist
' Propósito: Valida que el archivo se esté ejecutando desde una ruta
'            autorizada según la tabla tblWhiteList en la hoja Configuración.
'            Si la ruta actual no está en la whitelist, el libro se cierra.
' Fecha creación: 17/06/2026
' Dependencias: Configuration2 (TABLE_WHITELIST, SHEET_CONFIGURACION)
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Función: IsPathInWhitelist
' Propósito: Verifica si la ruta actual del archivo coincide con alguna
'            de las rutas permitidas en tblWhiteList.
' Retorna: Boolean
'   - True  = La ruta está autorizada (o la tabla no existe aún)
'   - False = La ruta NO está autorizada
' Lógica:
'   1. Si la tabla tblWhiteList no existe en Configuración, permite abrir
'      (para que el usuario pueda configurarla por primera vez).
'   2. Si la tabla existe pero está vacía, bloquea el acceso.
'   3. Compara la ruta actual (ThisWorkbook.Path) contra cada entrada
'      usando comparación case-insensitive y soportando rutas parciales
'      (ej: "C:\Proyectos" cubre "C:\Proyectos\SubCarpeta\").
' ----------------------------------------------------------------------
Public Function IsPathInWhitelist() As Boolean
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim tblWhitelist As ListObject
    Dim rutaActual As String
    Dim rutaPermitida As String
    Dim fila As ListRow
    
    ' Obtener ruta actual del archivo (sin nombre de archivo)
    rutaActual = LCase$(Trim$(ThisWorkbook.Path))
    
    ' Si el archivo aún no se ha guardado (Path vacío), permitir abrir
    If Len(rutaActual) = 0 Then
        IsPathInWhitelist = True
        Exit Function
    End If
    
    ' Intentar obtener la tabla tblWhiteList de la hoja Configuración
    Set wsConfig = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    
    ' Verificar si la tabla existe
    Dim tblExists As Boolean
    tblExists = False
    On Error Resume Next
    Set tblWhitelist = wsConfig.ListObjects(Configuration2.TABLE_WHITELIST)
    If Err.Number = 0 Then tblExists = True
    On Error GoTo ErrorHandler
    
    ' Si la tabla no existe aún, permitir abrir (primera configuración)
    If Not tblExists Then
        Debug.Print "[PathWhitelist] tblWhiteList no existe en Configuración. Acceso permitido (modo configuración inicial)."
        IsPathInWhitelist = True
        Exit Function
    End If
    
    ' Si la tabla no tiene datos, bloquear
    If tblWhitelist.DataBodyRange Is Nothing Then
        Debug.Print "[PathWhitelist] tblWhiteList existe pero está vacía. Acceso DENEGADO."
        IsPathInWhitelist = False
        Exit Function
    End If
    
    ' Recorrer cada fila de la whitelist y comparar rutas
    Dim columnaRuta As Long
    columnaRuta = tblWhitelist.ListColumns("Ruta").Index
    
    For Each fila In tblWhitelist.ListRows
        rutaPermitida = LCase$(Trim$(CStr(fila.Range.Cells(1, columnaRuta).Value)))
        
        ' Saltar filas vacías
        If Len(rutaPermitida) = 0 Then GoTo ContinueLoop
        
        ' Normalizar: asegurar que termine con backslash para comparación de prefijos
        If Right$(rutaPermitida, 1) <> "\" Then
            rutaPermitida = rutaPermitida & "\"
        End If
        If Right$(rutaActual, 1) <> "\" Then
            rutaActual = rutaActual & "\"
        End If
        
        ' Comparar: la ruta actual debe comenzar con la ruta permitida
        ' Esto soporta tanto rutas exactas como subcarpetas
        If Left$(rutaActual, Len(rutaPermitida)) = rutaPermitida Then
            Debug.Print "[PathWhitelist] Ruta autorizada: '" & rutaActual & "' coincide con '" & rutaPermitida & "'"
            IsPathInWhitelist = True
            Exit Function
        End If
        
ContinueLoop:
    Next fila
    
    ' Ninguna ruta coincide
    Debug.Print "[PathWhitelist] Ruta NO autorizada: '" & rutaActual & "'"
    IsPathInWhitelist = False
    Exit Function
    
ErrorHandler:
    ' Si hay error al leer la tabla (ej: columna "Ruta" no existe), permitir acceso
    Debug.Print "[PathWhitelist] Error al verificar whitelist: " & Err.Description & ". Acceso permitido por seguridad."
    IsPathInWhitelist = True
End Function

'' ----------------------------------------------------------------------
' Subrutina: EnforceWhitelist
' Propósito: Punto de entrada llamado desde Workbook_Open.
'            Si la ruta no está en la whitelist, muestra un mensaje
'            y cierra el libro inmediatamente.
' ----------------------------------------------------------------------
Public Sub EnforceWhitelist()
    If Not IsPathInWhitelist() Then
        MsgBox "ACCESO DENEGADO" & vbCrLf & vbCrLf & _
               "Este archivo no puede ejecutarse desde esta ubicación:" & vbCrLf & _
               ThisWorkbook.Path & vbCrLf & vbCrLf & _
               "Contacte al administrador del sistema si requiere acceso.", _
               vbCritical, "Seguridad - Ruta no autorizada"
        
        ' Cerrar sin guardar cambios
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub
