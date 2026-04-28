Attribute VB_Name = "CronogramaGestorService"
' ----------------------------------------------------------------------
' Módulo: CronogramaGestorService
' Descripción: Servicio de negocio para la gestión de cronogramas.
'              Permite pausar y reactivar inspecciones en tblCronogramaInspecciones
'              controlando su visibilidad en el resumen del menú principal.
' Fecha creación: 28/04/2026
' Dependencias: Configuration2, AuditLogger2, ErrorLogger2
' Expone:
'   - ValidarContrasenaGestor   : autenticación para acceso al gestor
'   - ObtenerDatosParaGestor    : todos los registros del cronograma para la UI
'   - ObtenerPlantasUnicas      : valores únicos de planta para filtro de UI
'   - PausarInspecciones        : cambia estado a "No" + auditoría
'   - ReactivarInspecciones     : cambia estado a "Si" + auditoría
'   - ContarInspeccionesPorEstado: estadísticas para la UI
' NOTAS DE ARQUITECTURA:
'   - Las inspecciones con Activo="No" NO aparecen en tblResumenCronograma.
'   - InspectionHistoryService NO filtra por esta columna (comportamiento intencional).
'   - Acceso directo a tblCronogramaInspecciones (patrón del proyecto, igual que
'     CronogramaResumen e InspectionScheduler).
'   - Backward compatible: filas sin valor en col. 20 son tratadas como "Si".
' ----------------------------------------------------------------------
Option Explicit

' Nombre de la columna de control en tblCronogramaInspecciones
Private Const COL_ACTIVO As String = "Activo en cronograma"

' ============================================================================
' AUTENTICACIÓN
' ============================================================================

'' ----------------------------------------------------------------------
' Función: ValidarContrasenaGestor
' Propósito: Verifica la contraseña para acceder al gestor de cronograma.
'            Registra el intento en AuditLogger2 (exitoso o fallido).
' Parámetros:
'   contrasena: Contraseña ingresada por el usuario
' Retorna: True si coincide con CRONOGRAMA_ADMIN_PASSWORD, False si no.
' ----------------------------------------------------------------------
Public Function ValidarContrasenaGestor(ByVal contrasena As String) As Boolean
    On Error GoTo ErrorHandler

    Dim esValida As Boolean
    esValida = (contrasena = Configuration2.CRONOGRAMA_ADMIN_PASSWORD)

    If esValida Then
        Call AuditLogger2.LogAction( _
            "CRONOGRAMA_ACCESO", _
            Configuration2.SHEET_CRONOGRAMA, _
            "Acceso al Gestor de Cronograma", _
            "", "Acceso exitoso", _
            "CronogramaGestorService.ValidarContrasenaGestor")
    Else
        Call AuditLogger2.LogAction( _
            "CRONOGRAMA_ACCESO_FALLIDO", _
            Configuration2.SHEET_CRONOGRAMA, _
            "Intento fallido de acceso al Gestor de Cronograma", _
            "", "Contrasena incorrecta", _
            "CronogramaGestorService.ValidarContrasenaGestor")
    End If

    ValidarContrasenaGestor = esValida
    Exit Function

ErrorHandler:
    Call ErrorLogger2.Log("CronogramaGestorService.ValidarContrasenaGestor", Err.Description, Err.Number)
    ValidarContrasenaGestor = False
End Function

' ============================================================================
' CONSULTAS
' ============================================================================

'' ----------------------------------------------------------------------
' Función: ObtenerDatosParaGestor
' Propósito: Retorna todos los registros de tblCronogramaInspecciones
'            con las columnas necesarias para frmGestorCronograma.
'            NO filtra por estado — retorna pausadas y activas.
' Retorna: Collection de Dictionary con claves:
'   "IDCronograma", "Iniciales", "NombrePlantilla", "Planta",
'   "Puesto", "Frecuencia", "EstadoCronograma",
'   "FechaUltimaInspeccion", "Activo"
' Nota: El filtrado por planta/estado es responsabilidad de la UI.
' ----------------------------------------------------------------------
Public Function ObtenerDatosParaGestor() As Collection
    On Error GoTo ErrorHandler

    Dim resultado As New Collection
    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    Dim cronRow As ListRow

    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)

    If tblCronograma.DataBodyRange Is Nothing Then
        Set ObtenerDatosParaGestor = resultado
        Exit Function
    End If

    For Each cronRow In tblCronograma.ListRows
        Dim datos As Object
        Set datos = CreateObject("Scripting.Dictionary")

        datos("IDCronograma") = Trim(CStr(cronRow.Range.Cells(1, tblCronograma.ListColumns("ID Cronograma").Index).Value))
        datos("Iniciales") = Trim(CStr(cronRow.Range.Cells(1, tblCronograma.ListColumns("Iniciales personal").Index).Value))
        datos("NombrePlantilla") = Trim(CStr(cronRow.Range.Cells(1, tblCronograma.ListColumns("Nombre plantilla").Index).Value))
        datos("Planta") = Trim(CStr(cronRow.Range.Cells(1, tblCronograma.ListColumns("Planta personal").Index).Value))
        datos("Puesto") = Trim(CStr(cronRow.Range.Cells(1, tblCronograma.ListColumns("Puesto").Index).Value))
        datos("Frecuencia") = cronRow.Range.Cells(1, tblCronograma.ListColumns("Frecuencia meses").Index).Value
        datos("EstadoCronograma") = Trim(CStr(cronRow.Range.Cells(1, tblCronograma.ListColumns("Estado cronograma").Index).Value))
        datos("FechaUltimaInspeccion") = cronRow.Range.Cells(1, tblCronograma.ListColumns("Fecha ultima inspeccion").Index).Value
        datos("Activo") = LeerActivoConFallback(cronRow, tblCronograma)

        resultado.Add datos
    Next cronRow

    Set ObtenerDatosParaGestor = resultado
    Exit Function

ErrorHandler:
    Call ErrorLogger2.Log("CronogramaGestorService.ObtenerDatosParaGestor", Err.Description, Err.Number)
    Set ObtenerDatosParaGestor = New Collection
End Function

'' ----------------------------------------------------------------------
' Función: ObtenerPlantasUnicas
' Propósito: Retorna los valores únicos de "Planta personal" en
'            tblCronogramaInspecciones, ordenados alfabéticamente.
'            Usado por frmGestorCronograma para poblar el combo de planta.
' Retorna: Collection de Strings
' ----------------------------------------------------------------------
Public Function ObtenerPlantasUnicas() As Collection
    On Error GoTo ErrorHandler

    Dim resultado As New Collection
    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    Dim cronRow As ListRow
    Dim dictPlantas As Object
    Dim planta As String
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim temp As String

    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)

    If tblCronograma.DataBodyRange Is Nothing Then
        Set ObtenerPlantasUnicas = resultado
        Exit Function
    End If

    ' Usar Dictionary para deduplicar (CompareMode=vbTextCompare = case-insensitive)
    Set dictPlantas = CreateObject("Scripting.Dictionary")
    dictPlantas.CompareMode = vbTextCompare

    For Each cronRow In tblCronograma.ListRows
        planta = Trim(CStr(cronRow.Range.Cells(1, tblCronograma.ListColumns("Planta personal").Index).Value))
        If Len(planta) > 0 And Not dictPlantas.Exists(planta) Then
            dictPlantas.Add planta, planta
        End If
    Next cronRow

    ' Ordenar alfabéticamente (bubble sort — cantidad de plantas es pequeña)
    arr = dictPlantas.Keys
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i

    For i = LBound(arr) To UBound(arr)
        resultado.Add arr(i)
    Next i

    Set ObtenerPlantasUnicas = resultado
    Exit Function

ErrorHandler:
    Call ErrorLogger2.Log("CronogramaGestorService.ObtenerPlantasUnicas", Err.Description, Err.Number)
    Set ObtenerPlantasUnicas = New Collection
End Function

'' ----------------------------------------------------------------------
' Función: ContarInspeccionesPorEstado
' Propósito: Cuenta cuántas filas de tblCronogramaInspecciones tienen
'            un valor específico en "Activo en cronograma".
' Parámetros:
'   estado: "Si" para activas, "No" para pausadas, "" para todas
' Retorna: Cantidad de registros
' ----------------------------------------------------------------------
Public Function ContarInspeccionesPorEstado(ByVal estado As String) As Long
    On Error GoTo ErrorHandler

    Dim total As Long
    total = 0

    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    Dim cronRow As ListRow

    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)

    If tblCronograma.DataBodyRange Is Nothing Then
        ContarInspeccionesPorEstado = 0
        Exit Function
    End If

    ' Sin filtro de estado: retornar total de filas
    If estado = "" Then
        ContarInspeccionesPorEstado = tblCronograma.ListRows.Count
        Exit Function
    End If

    For Each cronRow In tblCronograma.ListRows
        Dim activoVal As String
        activoVal = LeerActivoConFallback(cronRow, tblCronograma)
        If UCase(Trim(activoVal)) = UCase(Trim(estado)) Then
            total = total + 1
        End If
    Next cronRow

    ContarInspeccionesPorEstado = total
    Exit Function

ErrorHandler:
    Call ErrorLogger2.Log("CronogramaGestorService.ContarInspeccionesPorEstado", Err.Description, Err.Number)
    ContarInspeccionesPorEstado = 0
End Function

' ============================================================================
' OPERACIONES DE ESCRITURA
' ============================================================================

'' ----------------------------------------------------------------------
' Función: PausarInspecciones
' Propósito: Actualiza "Activo en cronograma" a "No" para cada ID recibido.
'            Si un registro ya está en "No", registra advertencia y continúa.
'            Si un ID no existe en la tabla, lo ignora silenciosamente.
' Parámetros:
'   idsInspecciones: Array de strings con IDs de tblCronogramaInspecciones
' Retorna: Número de registros modificados (excluye los ya pausados y no encontrados)
' Pipeline: UI → PausarInspecciones → LeerActivoConFallback → AuditLogger2
' ----------------------------------------------------------------------
Public Function PausarInspecciones(ByVal idsInspecciones() As String) As Long
    On Error GoTo ErrorHandler

    Dim actualizados As Long
    actualizados = 0

    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)

    Dim i As Long
    Dim idBuscado As String
    Dim cronRow As ListRow
    Dim idFila As String
    Dim estadoAnterior As String
    Dim accion As String

    For i = LBound(idsInspecciones) To UBound(idsInspecciones)
        idBuscado = Trim(idsInspecciones(i))
        If Len(idBuscado) = 0 Then GoTo SiguientePausar

        For Each cronRow In tblCronograma.ListRows
            idFila = Trim(CStr(cronRow.Range.Cells(1, tblCronograma.ListColumns("ID Cronograma").Index).Value))

            If idFila = idBuscado Then
                estadoAnterior = LeerActivoConFallback(cronRow, tblCronograma)

                ' Actualizar columna independientemente del estado previo
                cronRow.Range.Cells(1, tblCronograma.ListColumns(COL_ACTIVO).Index).Value = "No"

                If estadoAnterior = "No" Then
                    accion = "CRONOGRAMA_PAUSAR_ADVERTENCIA"  ' Ya estaba pausada
                Else
                    accion = "CRONOGRAMA_PAUSAR"
                    actualizados = actualizados + 1
                End If

                Call AuditLogger2.LogAction( _
                    accion, _
                    Configuration2.SHEET_CRONOGRAMA, _
                    COL_ACTIVO & " | ID: " & idBuscado, _
                    estadoAnterior, "No", _
                    "CronogramaGestorService.PausarInspecciones")

                Exit For  ' ID encontrado, no seguir buscando en filas
            End If
        Next cronRow

SiguientePausar:
    Next i

    PausarInspecciones = actualizados
    Exit Function

ErrorHandler:
    Call ErrorLogger2.Log("CronogramaGestorService.PausarInspecciones", Err.Description, Err.Number)
    PausarInspecciones = actualizados
End Function

'' ----------------------------------------------------------------------
' Función: ReactivarInspecciones
' Propósito: Actualiza "Activo en cronograma" a "Si" para cada ID recibido.
'            Si un registro ya está en "Si", registra advertencia y continúa.
'            Si un ID no existe en la tabla, lo ignora silenciosamente.
' Parámetros:
'   idsInspecciones: Array de strings con IDs de tblCronogramaInspecciones
' Retorna: Número de registros modificados (excluye los ya activos y no encontrados)
' Pipeline: UI → ReactivarInspecciones → LeerActivoConFallback → AuditLogger2
' ----------------------------------------------------------------------
Public Function ReactivarInspecciones(ByVal idsInspecciones() As String) As Long
    On Error GoTo ErrorHandler

    Dim actualizados As Long
    actualizados = 0

    Dim wsCronograma As Worksheet
    Dim tblCronograma As ListObject
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)

    Dim i As Long
    Dim idBuscado As String
    Dim cronRow As ListRow
    Dim idFila As String
    Dim estadoAnterior As String
    Dim accion As String

    For i = LBound(idsInspecciones) To UBound(idsInspecciones)
        idBuscado = Trim(idsInspecciones(i))
        If Len(idBuscado) = 0 Then GoTo SiguienteReactivar

        For Each cronRow In tblCronograma.ListRows
            idFila = Trim(CStr(cronRow.Range.Cells(1, tblCronograma.ListColumns("ID Cronograma").Index).Value))

            If idFila = idBuscado Then
                estadoAnterior = LeerActivoConFallback(cronRow, tblCronograma)

                ' Actualizar columna independientemente del estado previo
                cronRow.Range.Cells(1, tblCronograma.ListColumns(COL_ACTIVO).Index).Value = "Si"

                If estadoAnterior = "Si" Then
                    accion = "CRONOGRAMA_REACTIVAR_ADVERTENCIA"  ' Ya estaba activa
                Else
                    accion = "CRONOGRAMA_REACTIVAR"
                    actualizados = actualizados + 1
                End If

                Call AuditLogger2.LogAction( _
                    accion, _
                    Configuration2.SHEET_CRONOGRAMA, _
                    COL_ACTIVO & " | ID: " & idBuscado, _
                    estadoAnterior, "Si", _
                    "CronogramaGestorService.ReactivarInspecciones")

                Exit For  ' ID encontrado, no seguir buscando en filas
            End If
        Next cronRow

SiguienteReactivar:
    Next i

    ReactivarInspecciones = actualizados
    Exit Function

ErrorHandler:
    Call ErrorLogger2.Log("CronogramaGestorService.ReactivarInspecciones", Err.Description, Err.Number)
    ReactivarInspecciones = actualizados
End Function

' ============================================================================
' HELPERS PRIVADOS
' ============================================================================

'' ----------------------------------------------------------------------
' Función privada: LeerActivoConFallback
' Propósito: Lee el valor de "Activo en cronograma" de una fila con
'            compatibilidad hacia atrás. Si la columna no existe o está
'            vacía, retorna "Si" (activo por defecto).
' Parámetros:
'   cronRow      : Fila de tblCronogramaInspecciones
'   tblCronograma: Tabla para acceder a columnas por nombre
' Retorna: "Si" o "No"
' ----------------------------------------------------------------------
Private Function LeerActivoConFallback(ByVal cronRow As ListRow, ByVal tblCronograma As ListObject) As String
    On Error Resume Next

    Dim colIdx As Long
    colIdx = tblCronograma.ListColumns(COL_ACTIVO).Index

    If Err.Number <> 0 Or colIdx = 0 Then
        ' Columna no existe en la tabla — asumir activo (backward compatible)
        Err.Clear
        LeerActivoConFallback = "Si"
        Exit Function
    End If

    On Error GoTo 0

    Dim val As String
    val = Trim(CStr(cronRow.Range.Cells(1, colIdx).Value))

    LeerActivoConFallback = IIf(Len(val) = 0, "Si", val)
End Function
