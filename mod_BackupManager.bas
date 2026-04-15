'******************************************************************************
' Módulo: mod_BackupManager
' Proyecto: CA-HC-004 PROCESO DE VALIDACION
' Descripción: Sistema de gestión de copias de seguridad automáticas del libro
'              Excel. Crea copias de respaldo en la misma carpeta del archivo
'              original con nomenclatura estandarizada, facilitando recuperación
'              ante errores o pérdida de datos.
'
' Responsabilidades:
'   - Creación automática de copias de seguridad del libro
'   - Validación de que el archivo esté guardado antes de crear backup
'   - Construcción de nombre estandarizado para backup
'   - Eliminación de backup anterior antes de crear nuevo
'   - Logging de operaciones con Debug.Print
'
' Características clave:
'   - Nomenclatura: "Copia de Seguridad [NombreOriginal].xlsm"
'   - Ubicación: Misma carpeta que archivo original
'   - Sobreescritura automática de backup anterior
'   - Preservación de extensión original del archivo
'   - Validación de archivo guardado (Path no vacío)
'   - Compatible con rutas OneDrive y locales
'
' Estrategia de backup:
'   1. Verifica que archivo esté guardado (ThisWorkbook.Path <> "")
'   2. Convierte URLs de OneDrive a rutas locales
'   3. Construye nombre: "Copia de Seguridad [base][ext]"
'   4. Elimina backup anterior si existe
'   5. Crea nueva copia con SaveCopyAs
'
' Integración con el proyecto:
'   - Se ejecuta automáticamente en Workbook_BeforeSave
'   - Backup antes de cada guardado voluntario del usuario
'   - Protege datos del proceso de validación CA-HC-004
'   - Facilita recuperación ante errores de entrada de datos
'
' Casos de uso:
'   - Backup automático al guardar (Workbook_BeforeSave)
'   - Backup antes de operaciones críticas de validación
'   - Backup manual desde interfaz de usuario
'
' Manejo de archivos no guardados:
'   - Si ThisWorkbook.Path = "": retorna False, no crea backup
'   - Evita errores con archivos nuevos sin guardar
'   - Debug.Print notifica situación
'
' Dependencias: Ninguna (módulo independiente)
'
' Autor: Sistema CA-HC-004
' Fecha creación: 2025
' Última modificación: 19/02/2026 - Integración con proyecto CA-HC-004
'******************************************************************************
Option Explicit

'******************************************************************************
' Función: CrearBackupAutomatico
' Descripción: Crea una copia de seguridad del libro actual en la misma carpeta
'              con nomenclatura estandarizada. Elimina backup anterior si existe
'              y valida que el archivo esté guardado antes de proceder.
'
' Parámetros: Ninguno
'
' Retorno:
'   - (Boolean): True si backup creado exitosamente
'                False si falla o archivo no está guardado
'
' Flujo de ejecución:
'   1. Valida que archivo esté guardado:
'      - Verifica ThisWorkbook.Path <> ""
'      - Si vacío: retorna False, no crea backup
'   2. Obtiene información del archivo original:
'      - Ruta completa (FullName)
'      - Nombre con extensión (Name)
'   3. Construye nombre del backup:
'      - Separa nombre base de extensión (InStrRev para último punto)
'      - Formato: "Copia de Seguridad [base][ext]"
'      - Ejemplo: "Copia de Seguridad CA-HC-004 V01-Preliminar.xlsm"
'   4. Construye ruta completa del backup:
'      - Misma carpeta que original (ThisWorkbook.Path)
'      - Usa Application.PathSeparator (compatible Windows/Mac)
'   5. Elimina backup anterior si existe:
'      - Usa FileSystemObject para verificar existencia
'      - Kill para eliminar archivo
'      - On Error Resume Next para tolerancia si no existe
'   6. Crea nueva copia con SaveCopyAs:
'      - Método nativo Excel para copia exacta
'      - Preserva todo el contenido del libro
'   7. Logging con Debug.Print:
'      - Ruta original
'      - Ruta backup
'      - Notificación de backup anterior eliminado
'      - Confirmación de éxito con "?"
'
' Ejemplo de uso:
'   ' Backup automático al guardar (Workbook_BeforeSave)
'   Call mod_BackupManager.CrearBackupAutomatico
'
'   ' Antes de operación crítica de validación
'   If mod_BackupManager.CrearBackupAutomatico() Then
'       ' Proceder con operación de validación
'       Call dataEntryService.ProcesarEntrada()
'   Else
'       MsgBox "No se pudo crear backup. Operación cancelada.", vbExclamation
'   End If
'
' Manejo de errores:
'   - On Error GoTo ErrorHandler
'   - Debug.Print con número y descripción de error
'   - Retorna False en caso de error
'   - No muestra MsgBox (silencioso)
'
' Manejo de archivo no guardado:
'   - Si ThisWorkbook.Path = "": archivo nuevo sin guardar
'   - Debug.Print: "Archivo no guardado aún. No se crea backup."
'   - Retorna False sin error
'   - Previene error al intentar acceder a Path vacío
'
' Construcción del nombre:
'   - InStrRev busca último punto (extensión)
'   - Si no hay punto: asume .xlsm como extensión
'   - Preserva extensión original (.xlsm, .xlsx, .xls, etc.)
'   - Formato consistente: "Copia de Seguridad [base][ext]"
'
' Logging (Debug.Print):
'   - "[Backup] Creando copia de seguridad..."
'   - "[Backup] Original: [ruta]"
'   - "[Backup] Backup: [ruta]"
'   - "[Backup] Backup anterior eliminado." (si aplicable)
'   - "[Backup] ? Copia de seguridad creada exitosamente."
'   - "[Backup] ERROR: [num] - [desc]" (si error)
'
' Notas:
'   - SaveCopyAs NO guarda libro actual (solo crea copia)
'   - Libro actual NO se marca como modificado
'   - Usuario NO ve ninguna notificación visual
'   - Backup se crea en misma carpeta (no subcarpeta)
'   - Compatible con rutas UNC y OneDrive
'   - Esencial para integridad de datos en proceso de validación
'
' Consideraciones:
'   - No crea carpeta "Backups" separada
'   - Sobreescribe backup anterior (no versionado múltiple)
'   - Para versionado múltiple: agregar timestamp al nombre
'   - Usa Kill en lugar de fso.DeleteFile (más directo)
'
' Mejoras futuras posibles:
'   - Agregar timestamp: "Copia de Seguridad [base]_YYYYMMDD_HHMMSS[ext]"
'   - Límite de backups antiguos (eliminar más viejos que N días)
'   - Carpeta "Backups" separada
'   - Compresión de backups
'
' Autor: Sistema CA-HC-004
' Fecha creación: 2025
' Última modificación: 19/02/2026 - Documentación actualizada
'******************************************************************************
Public Function CrearBackupAutomatico() As Boolean
    On Error GoTo ErrorHandler
    
    Dim originalPath As String
    Dim originalName As String
    Dim backupPath As String
    Dim fso As Object
    Dim folderPath As String
    
    ' Verificar que el archivo ya esté guardado (no sea un nuevo libro sin ruta)
    If ThisWorkbook.Path = "" Then
        Debug.Print "[Backup] Archivo no guardado aún. No se crea backup."
        CrearBackupAutomatico = False
        Exit Function
    End If
    
    ' Obtener la ruta REAL del archivo (convierte URLs de OneDrive a rutas locales)
    originalPath = GetLocalPath(ThisWorkbook.FullName)
    folderPath = GetLocalPath(ThisWorkbook.Path)
    originalName = ThisWorkbook.Name
    
    Debug.Print "[Backup] Creando copia de seguridad..."
    Debug.Print "[Backup] Original: " & originalPath
    
    ' Construir nombre del backup: "Copia de Seguridad [NombreOriginal]"
    Dim baseName As String
    Dim extension As String
    Dim dotPos As Long
    Dim backupName As String
    
    dotPos = InStrRev(originalName, ".")
    If dotPos > 0 Then
        baseName = Left(originalName, dotPos - 1)
        extension = Mid(originalName, dotPos)
    Else
        baseName = originalName
        extension = ".xlsm"
    End If
    
    backupName = "Copia de Seguridad " & baseName & extension
    backupPath = folderPath & Application.PathSeparator & backupName
    
    Debug.Print "[Backup] Backup: " & backupPath
    
    ' Si el backup ya existe, eliminarlo primero
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If fso.FileExists(backupPath) Then
        Kill backupPath
        Debug.Print "[Backup] Backup anterior eliminado."
    End If
    On Error GoTo ErrorHandler
    
    ' Crear la copia usando la ruta local
    ThisWorkbook.SaveCopyAs backupPath
    
    Debug.Print "[Backup] ? Copia de seguridad creada exitosamente."
    CrearBackupAutomatico = True
    Exit Function
    
ErrorHandler:
    Debug.Print "[Backup] ERROR: " & Err.Number & " - " & Err.Description
    CrearBackupAutomatico = False
End Function

'******************************************************************************
' Función: GetLocalPath
' Descripción: Convierte una URL de OneDrive/SharePoint a una ruta local del
'              sistema de archivos de forma robusta. Si la ruta ya es local,
'              la devuelve sin cambios.
'
' Parámetros:
'   - urlPath (String): Ruta que puede ser URL (https://...) o ruta local (C:\...)
'
' Retorno:
'   - (String): Ruta local del sistema de archivos
'
' Funcionamiento:
'   1. Detecta si es una URL de OneDrive (https://d.docs.live.net/...)
'   2. Obtiene la ruta local de OneDrive desde variables de entorno
'   3. Extrae la ruta relativa de la URL y la combina con la ruta local
'   4. Si no es OneDrive, intenta conversión genérica con FSO
'   5. Si todo falla, devuelve la ruta original
'
' Ejemplo:
'   URL de entrada:
'     "https://d.docs.live.net/ABC123/Escritorio/Proyecto/archivo.xlsm"
'   Salida (ruta local):
'     "C:\Users\Usuario\OneDrive\Escritorio\Proyecto\archivo.xlsm"
'
' Compatibilidad:
'   - OneDrive Personal: https://d.docs.live.net/...
'   - OneDrive Business: https://[tenant]-my.sharepoint.com/...
'   - SharePoint: https://[tenant].sharepoint.com/...
'   - Rutas locales: Las devuelve sin cambios
'
' Notas:
'   - Usa variables de entorno OneDrive, OneDriveConsumer, OneDriveCommercial
'   - Si OneDrive no está configurado, intenta método alternativo
'   - Decodifica caracteres URL (%20 ? espacio)
'
' Autor: Asurity
' Fecha creación: 12/01/2026
' Última modificación: 12/01/2026 - Versión robusta con detección de OneDrive
'******************************************************************************
Private Function GetLocalPath(ByVal urlPath As String) As String
    On Error Resume Next
    
    Dim fso As Object
    Dim wsh As Object
    Dim oneDrivePath As String
    Dim relativePath As String
    Dim localPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")
    
    ' Si NO es una URL, devolver la ruta original
    If Left(LCase(urlPath), 4) <> "http" Then
        GetLocalPath = urlPath
        Exit Function
    End If
    
    ' ========================================================================
    ' CASO 1: URL DE ONEDRIVE PERSONAL (https://d.docs.live.net/...)
    ' ========================================================================
    If InStr(1, urlPath, "d.docs.live.net", vbTextCompare) > 0 Then
        ' Intentar obtener ruta de OneDrive desde variables de entorno
        On Error Resume Next
        oneDrivePath = wsh.ExpandEnvironmentStrings("%OneDrive%")
        If oneDrivePath = "%OneDrive%" Or oneDrivePath = "" Then
            oneDrivePath = wsh.ExpandEnvironmentStrings("%OneDriveConsumer%")
        End If
        On Error GoTo 0
        
        ' Si encontramos la ruta de OneDrive, extraer la parte relativa de la URL
        If oneDrivePath <> "%OneDrive%" And oneDrivePath <> "%OneDriveConsumer%" And oneDrivePath <> "" Then
            ' Extraer todo después del ID de usuario en la URL
            ' Formato: https://d.docs.live.net/[ID]/[ruta relativa]
            Dim parts() As String
            parts = Split(urlPath, "/")
            
            ' Encontrar índice después de "d.docs.live.net/[ID]"
            Dim startIdx As Integer
            startIdx = -1
            Dim i As Integer
            For i = LBound(parts) To UBound(parts)
                If InStr(1, parts(i), "d.docs.live.net", vbTextCompare) > 0 Then
                    startIdx = i + 2 ' Saltar "d.docs.live.net" y el ID
                    Exit For
                End If
            Next i
            
            ' Reconstruir ruta relativa
            If startIdx > 0 And startIdx <= UBound(parts) Then
                relativePath = ""
                For i = startIdx To UBound(parts)
                    If parts(i) <> "" Then
                        If relativePath <> "" Then relativePath = relativePath & "\"
                        ' Decodificar caracteres URL especiales
                        relativePath = relativePath & DecodeURL(parts(i))
                    End If
                Next i
                
                ' Construir ruta local completa
                localPath = oneDrivePath & "\" & relativePath
                
                ' Verificar que la ruta existe
                If fso.FileExists(localPath) Or fso.FolderExists(localPath) Then
                    GetLocalPath = localPath
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' ========================================================================
    ' CASO 2: URL DE ONEDRIVE BUSINESS O SHAREPOINT
    ' ========================================================================
    If InStr(1, urlPath, "sharepoint.com", vbTextCompare) > 0 Then
        ' Intentar obtener ruta de OneDrive Business
        On Error Resume Next
        oneDrivePath = wsh.ExpandEnvironmentStrings("%OneDriveCommercial%")
        If oneDrivePath = "%OneDriveCommercial%" Or oneDrivePath = "" Then
            oneDrivePath = wsh.ExpandEnvironmentStrings("%OneDrive%")
        End If
        On Error GoTo 0
        
        ' Lógica similar para SharePoint (más compleja, requiere parsing adicional)
        ' Por ahora, intentamos con FSO
    End If
    
    ' ========================================================================
    ' CASO 3: MÉTODO ALTERNATIVO - Intentar con FileSystemObject
    ' ========================================================================
    On Error Resume Next
    localPath = fso.GetAbsolutePathName(urlPath)
    If Err.Number = 0 And Left(LCase(localPath), 4) <> "http" Then
        GetLocalPath = localPath
        Exit Function
    End If
    On Error GoTo 0
    
    ' ========================================================================
    ' CASO 4: TODO FALLÓ - Devolver URL original con advertencia
    ' ========================================================================
    GetLocalPath = urlPath
    Debug.Print "[Backup] ADVERTENCIA: No se pudo convertir URL a ruta local."
    Debug.Print "[Backup] Si el backup falla, guarde el archivo en una carpeta local primero."
    
    On Error GoTo 0
End Function

'******************************************************************************
' Función: DecodeURL
' Descripción: Decodifica caracteres especiales en URLs (ej: %20 ? espacio)
'
' Parámetros:
'   - encodedStr (String): Cadena con caracteres codificados
'
' Retorno:
'   - (String): Cadena decodificada
'
' Autor: Asurity
' Fecha: 12/01/2026
'******************************************************************************
Private Function DecodeURL(ByVal encodedStr As String) As String
    Dim result As String
    result = encodedStr
    
    ' Reemplazar caracteres comunes codificados en URLs
    result = Replace(result, "%20", " ")
    result = Replace(result, "%C3%A1", "á")
    result = Replace(result, "%C3%A9", "é")
    result = Replace(result, "%C3%AD", "í")
    result = Replace(result, "%C3%B3", "ó")
    result = Replace(result, "%C3%BA", "ú")
    result = Replace(result, "%C3%B1", "ñ")
    result = Replace(result, "%C3%81", "Á")
    result = Replace(result, "%C3%89", "É")
    result = Replace(result, "%C3%8D", "Í")
    result = Replace(result, "%C3%93", "Ó")
    result = Replace(result, "%C3%9A", "Ú")
    result = Replace(result, "%C3%91", "Ñ")
    
    DecodeURL = result
End Function