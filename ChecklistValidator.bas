' ----------------------------------------------------------------------
' Módulo: ChecklistValidator
' Descripción: Validación de campos del formulario de checklist virtual.
'              Valida cabecera (campos obligatorios, formatos) y respuestas
'              (completitud de preguntas respondidas).
' Fecha creación: 14/04/2026
' Dependencias: Configuration2, ErrorLogger2
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Función: ValidarCabecera
' Propósito: Valida todos los campos obligatorios de la cabecera del formulario.
' Parámetros:
'   datos: Dictionary con claves de campos del formulario
' Retorna: "" si válido, o mensaje de error específico si hay problema.
' ----------------------------------------------------------------------
Public Function ValidarCabecera(ByVal datos As Object) As String
    On Error GoTo ErrorHandler
    
    ' --- Campos auto-rellenados (siempre deben existir) ---
    If Trim(CStr(datos("Iniciales"))) = "" Then
        ValidarCabecera = "El campo 'Evaluado' está vacío. Datos del cronograma incorrectos."
        Exit Function
    End If
    
    If Trim(CStr(datos("Puesto"))) = "" Then
        ValidarCabecera = "El campo 'Puesto' está vacío. Datos del cronograma incorrectos."
        Exit Function
    End If
    
    If Trim(CStr(datos("Planta"))) = "" Then
        ValidarCabecera = "El campo 'Planta' está vacío. Datos del cronograma incorrectos."
        Exit Function
    End If
    
    ' --- Campos de selección obligatorios ---
    If Trim(CStr(datos("Area"))) = "" Then
        ValidarCabecera = "Debe seleccionar un Área."
        Exit Function
    End If
    
    If Trim(CStr(datos("LineaAuditada"))) = "" Then
        ValidarCabecera = "Debe seleccionar un Equipo / Línea auditada."
        Exit Function
    End If
    
    If Trim(CStr(datos("Evaluador"))) = "" Then
        ValidarCabecera = "Debe seleccionar un Evaluador."
        Exit Function
    End If
    
    ' Los campos opcionales AY1, AY2, OP pueden estar vacíos
    
    If Trim(CStr(datos("LugarAuditoria"))) = "" Then
        ValidarCabecera = "Debe seleccionar el Lugar de auditoría (Dentro/Fuera del área)."
        Exit Function
    End If
    
    ' --- Fecha ---
    Dim fechaStr As String
    fechaStr = Trim(CStr(datos("FechaInspeccion")))
    
    If fechaStr = "" Then
        ValidarCabecera = "Debe ingresar la Fecha evaluada."
        Exit Function
    End If
    
    ' Usar ParseFechaDMY (independiente de locale) en vez de IsDate/CDate
    Dim fechaParsed As Variant
    fechaParsed = ParseFechaDMY(fechaStr)
    
    If IsEmpty(fechaParsed) Then
        ValidarCabecera = "La Fecha evaluada no tiene un formato válido (DD/MM/AAAA)."
        Exit Function
    End If
    
    If CDate(fechaParsed) > Date Then
        ValidarCabecera = "La Fecha evaluada no puede ser posterior a hoy."
        Exit Function
    End If
    
    ' --- Fecha Auditada ---
    Dim fechaAuditadaStr As String
    fechaAuditadaStr = Trim(CStr(datos("FechaAuditada")))
    
    If fechaAuditadaStr = "" Then
        ValidarCabecera = "Debe ingresar la Fecha Auditada."
        Exit Function
    End If
    
    Dim fechaAudParsed As Variant
    fechaAudParsed = ParseFechaDMY(fechaAuditadaStr)
    
    If IsEmpty(fechaAudParsed) Then
        ValidarCabecera = "La Fecha Auditada no tiene un formato válido (DD/MM/AAAA)."
        Exit Function
    End If
    
    If CDate(fechaAudParsed) > Date Then
        ValidarCabecera = "La Fecha Auditada no puede ser posterior a hoy."
        Exit Function
    End If
    
    ' --- Horas ---
    Dim horaInicio As String
    Dim horaTermino As String
    horaInicio = Trim(CStr(datos("HoraInicio")))
    horaTermino = Trim(CStr(datos("HoraTermino")))
    
    If horaInicio = "" Then
        ValidarCabecera = "Debe ingresar la Hora de inicio."
        Exit Function
    End If
    
    If Not ValidarFormatoHora(horaInicio) Then
        ValidarCabecera = "La Hora de inicio no tiene formato válido (HH:MM)."
        Exit Function
    End If
    
    If horaTermino = "" Then
        ValidarCabecera = "Debe ingresar la Hora de término."
        Exit Function
    End If
    
    If Not ValidarFormatoHora(horaTermino) Then
        ValidarCabecera = "La Hora de término no tiene formato válido (HH:MM)."
        Exit Function
    End If
    
    ' Validar que hora término > hora inicio
    If horaTermino <= horaInicio Then
        ValidarCabecera = "La Hora de término debe ser posterior a la Hora de inicio."
        Exit Function
    End If
    
    ' --- Todo válido ---
    ValidarCabecera = ""
    Exit Function
    
ErrorHandler:
    ValidarCabecera = "Error de validación: " & Err.Description
    Call ErrorLogger2.Log("ChecklistValidator.ValidarCabecera", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ValidarRespuestasCompletas
' Propósito: Verifica que todas las preguntas hayan sido respondidas.
' Parámetros:
'   respuestas: Dictionary con Key=ID_Pregunta, Value=ID_Opcion
'   totalPreguntas: número total de preguntas esperadas
' Retorna: "" si válido, o mensaje con preguntas faltantes.
' ----------------------------------------------------------------------
Public Function ValidarRespuestasCompletas(ByVal respuestas As Object, ByVal totalPreguntas As Long) As String
    On Error GoTo ErrorHandler
    
    ' Verificar que se hayan respondido todas las preguntas
    Dim preguntasRespondidas As Long
    preguntasRespondidas = respuestas.Count
    
    If preguntasRespondidas < totalPreguntas Then
        Dim faltantes As Long
        faltantes = totalPreguntas - preguntasRespondidas
        ValidarRespuestasCompletas = "Faltan " & faltantes & " pregunta(s) por responder." & vbCrLf & _
                                     "Todas las preguntas deben tener una opción seleccionada."
        Exit Function
    End If
    
    ' Verificar que ninguna respuesta esté vacía (doble verificación de seguridad)
    Dim key As Variant
    Dim vacias As Long
    vacias = 0
    
    For Each key In respuestas.Keys
        If Trim(CStr(respuestas(key))) = "" Then
            vacias = vacias + 1
        End If
    Next key
    
    If vacias > 0 Then
        ValidarRespuestasCompletas = "Hay " & vacias & " pregunta(s) sin respuesta seleccionada." & vbCrLf & _
                                     "Todas las preguntas deben tener una opción seleccionada."
        Exit Function
    End If
    
    ValidarRespuestasCompletas = ""
    Exit Function
    
ErrorHandler:
    ValidarRespuestasCompletas = "Error al validar respuestas: " & Err.Description
    Call ErrorLogger2.Log("ChecklistValidator.ValidarRespuestasCompletas", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ValidarTodo
' Propósito: Ejecuta todas las validaciones (cabecera + respuestas).
'            Acumula TODOS los errores encontrados para mostrarlos juntos.
' Retorna: "" si todo válido, o lista de errores encontrados.
' ----------------------------------------------------------------------
Public Function ValidarTodo(ByVal datos As Object, ByVal respuestas As Object, ByVal totalPreguntas As Long) As String
    On Error GoTo ErrorHandler
    
    Dim errores As String
    errores = ""
    
    ' Validar cabecera
    Dim errorCabecera As String
    errorCabecera = ValidarCabecera(datos)
    If errorCabecera <> "" Then
        errores = "DATOS DE LA INSPECCIÓN:" & vbCrLf & _
                  "  • " & errorCabecera & vbCrLf & vbCrLf
    End If
    
    ' Validar respuestas
    Dim errorRespuestas As String
    errorRespuestas = ValidarRespuestasCompletas(respuestas, totalPreguntas)
    If errorRespuestas <> "" Then
        errores = errores & "RESPUESTAS DE PREGUNTAS:" & vbCrLf & _
                  "  • " & errorRespuestas & vbCrLf
    End If
    
    ValidarTodo = errores
    Exit Function
    
ErrorHandler:
    ValidarTodo = "Error de validación general: " & Err.Description
    Call ErrorLogger2.Log("ChecklistValidator.ValidarTodo", Err.Description, Err.Number)
End Function

'' ----------------------------------------------------------------------
' Función: ValidarFormatoHora
' Propósito: Verifica que una cadena tenga formato HH:MM válido.
' Retorna: True si formato válido, False si no.
' ----------------------------------------------------------------------
Private Function ValidarFormatoHora(ByVal hora As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificar longitud
    If Len(hora) <> 5 Then
        ValidarFormatoHora = False
        Exit Function
    End If
    
    ' Verificar separador
    If Mid(hora, 3, 1) <> ":" Then
        ValidarFormatoHora = False
        Exit Function
    End If
    
    ' Verificar partes numéricas
    Dim hh As String
    Dim mm As String
    hh = Left(hora, 2)
    mm = Right(hora, 2)
    
    If Not IsNumeric(hh) Or Not IsNumeric(mm) Then
        ValidarFormatoHora = False
        Exit Function
    End If
    
    ' Verificar rangos
    If CInt(hh) < 0 Or CInt(hh) > 23 Then
        ValidarFormatoHora = False
        Exit Function
    End If
    
    If CInt(mm) < 0 Or CInt(mm) > 59 Then
        ValidarFormatoHora = False
        Exit Function
    End If
    
    ValidarFormatoHora = True
    Exit Function
    
ErrorHandler:
    ValidarFormatoHora = False
End Function

'' ======================================================================
' Función: IsNumeric
' Propósito: Verifica si una cadena contiene solo números.
' Parámetros:
'   valor: Cadena a verificar
' Retorna: True si es numérico, False si no.
' ======================================================================
Private Function IsNumeric(ByVal valor As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim c As String
    
    If Len(valor) = 0 Then
        IsNumeric = False
        Exit Function
    End If
    
    For i = 1 To Len(valor)
        c = Mid(valor, i, 1)
        If Not (c >= "0" And c <= "9") Then
            IsNumeric = False
            Exit Function
        End If
    Next i
    
    IsNumeric = True
    Exit Function
    
ErrorHandler:
    IsNumeric = False
End Function

'' ======================================================================
' Función: ParseFechaDMY
' Propósito: Convierte una cadena de fecha en formato dd/mm/yyyy o dd-mm-yyyy
'            a un valor Date de forma INDEPENDIENTE de la configuración regional.
'            Resuelve el bug de CDate() que interpreta erróneamente las fechas
'            cuando el sistema usa separador decimal "." (formato mm/dd/yyyy).
' Parámetros:
'   fechaStr: Cadena con fecha en formato dd/mm/yyyy o dd-mm-yyyy
' Retorna: Variant con el Date si es válido, o Empty si no se pudo parsear.
' Fecha: 22/06/2026 — Fix de locale-dependence en CDate/IsDate
' ======================================================================
Public Function ParseFechaDMY(ByVal fechaStr As String) As Variant
    On Error GoTo ErrorHandler
    
    ParseFechaDMY = Empty
    
    ' Limpiar espacios
    fechaStr = Trim(fechaStr)
    If fechaStr = "" Then Exit Function
    
    ' Normalizar separadores: aceptar "/" o "-"
    Dim sep As String
    If InStr(fechaStr, "/") > 0 Then
        sep = "/"
    ElseIf InStr(fechaStr, "-") > 0 Then
        sep = "-"
    Else
        Exit Function  ' Sin separador reconocido
    End If
    
    Dim partes() As String
    partes = Split(fechaStr, sep)
    
    ' Debe tener exactamente 3 partes: día, mes, año
    If UBound(partes) <> 2 Then Exit Function
    
    Dim d As Long, m As Long, a As Long
    
    ' Convertir a números (independiente de locale)
    On Error Resume Next
    d = CLng(Trim(partes(0)))
    m = CLng(Trim(partes(1)))
    a = CLng(Trim(partes(2)))
    On Error GoTo ErrorHandler
    
    ' Si hubo error de conversión, salir
    If d = 0 Or m = 0 Or a = 0 Then Exit Function
    
    ' Validar rangos
    If d < 1 Or d > 31 Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If a < 1900 Or a > 2100 Then Exit Function
    
    ' Validar días según el mes (incluyendo años bisiestos)
    Dim diasPorMes As Variant
    diasPorMes = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    
    ' Ajustar febrero para años bisiestos
    If m = 2 Then
        If (a Mod 4 = 0 And a Mod 100 <> 0) Or (a Mod 400 = 0) Then
            If d > 29 Then Exit Function
        Else
            If d > 28 Then Exit Function
        End If
    Else
        If d > diasPorMes(m - 1) Then Exit Function
    End If
    
    ' Construir fecha con DateSerial (SIEMPRE interpreta año, mes, día en ese orden)
    Dim fechaResult As Date
    fechaResult = DateSerial(a, m, d)
    
    ' Verificación adicional: DateSerial ajusta fechas inválidas (ej: 31/Feb → 3/Mar)
    ' Si el mes resultante no coincide, la fecha era inválida
    If Month(fechaResult) <> m Or Day(fechaResult) <> d Or Year(fechaResult) <> a Then
        Exit Function
    End If
    
    ParseFechaDMY = fechaResult
    Exit Function
    
ErrorHandler:
    ParseFechaDMY = Empty
End Function

'' ======================================================================
' Función: CorregirYValidarFecha
' Propósito: Intenta convertir entrada de fecha a formato dd-mm-yyyy.
'            Acepta múltiples formatos: dd/mm/yyyy, dd-mm-yyyy, yyyy-mm-dd, etc.
' Parámetros:
'   fechaStr: Cadena con fecha a corregir
' Retorna: Dictionary con:
'   "valido": True/False
'   "valor": Fecha en formato dd-mm-yyyy (si válido) o vacío (si inválido)
'   "mensaje": Mensaje para el usuario
' ACTUALIZADO: 22/06/2026 — Usa ParseFechaDMY (independiente de locale) en vez de CDate
' ======================================================================
Public Function CorregirYValidarFecha(ByVal fechaStr As String) As Object
    On Error GoTo ErrorHandler
    
    Dim resultado As Object
    Set resultado = CreateObject("Scripting.Dictionary")
    resultado("valido") = False
    resultado("valor") = ""
    resultado("mensaje") = ""
    
    ' Limpiar espacios
    fechaStr = Trim(fechaStr)
    
    ' Validar que no esté vacío
    If fechaStr = "" Then
        resultado("mensaje") = "El campo de fecha está vacío."
        Set CorregirYValidarFecha = resultado
        Exit Function
    End If
    
    ' Intentar interpretar como fecha usando parseo independiente de locale
    Dim fechaParsed As Variant
    fechaParsed = ParseFechaDMY(fechaStr)
    
    If IsEmpty(fechaParsed) Then
        resultado("mensaje") = "Fecha no válida. Use formato dd-mm-yyyy o dd/mm/yyyy."
        Set CorregirYValidarFecha = resultado
        Exit Function
    End If
    
    Dim fechaDate As Date
    fechaDate = fechaParsed  ' ParseFechaDMY ya devuelve Date (no usar CDate que depende del locale)
    
    ' Verificar que no sea futura
    If fechaDate > Date Then
        resultado("mensaje") = "La fecha no puede ser posterior a hoy."
        Set CorregirYValidarFecha = resultado
        Exit Function
    End If
    
    ' Convertir al formato dd-mm-yyyy
    Dim fechaCorregida As String
    fechaCorregida = Format(fechaDate, "dd-mm-yyyy")
    
    ' Comparar con original
    Dim esCambio As Boolean
    esCambio = (fechaCorregida <> fechaStr)
    
    resultado("valido") = True
    resultado("valor") = fechaCorregida
    
    If esCambio Then
        resultado("mensaje") = "Fecha convertida a formato dd-mm-yyyy: " & fechaCorregida
    Else
        resultado("mensaje") = ""
    End If
    
    Set CorregirYValidarFecha = resultado
    Exit Function
    
ErrorHandler:
    resultado("mensaje") = "Error al validar fecha: " & Err.Description
    Call ErrorLogger2.Log("ChecklistValidator.CorregirYValidarFecha", Err.Description, Err.Number)
    Set CorregirYValidarFecha = resultado
End Function

'' ======================================================================
' Función: CorregirYValidarFechaVencimiento
' Propósito: Intenta convertir entrada de fecha de VENCIMIENTO a formato dd-mm-yyyy.
'            VALIDA que la fecha sea FUTURA (no vencida).
'            Acepta múltiples formatos: dd/mm/yyyy, dd-mm-yyyy, yyyy-mm-dd, etc.
' Parámetros:
'   fechaStr: Cadena con fecha a corregir
' Retorna: Dictionary con:
'   "valido": True/False
'   "valor": Fecha en formato dd-mm-yyyy (si válido) o vacío (si inválido)
'   "mensaje": Mensaje para el usuario
' FASE 7 - 23/04/2026
' ACTUALIZADO: 24/04/2026 - Validación de fecha futura obligatoria
' ACTUALIZADO: 22/06/2026 — Usa ParseFechaDMY (independiente de locale) en vez de CDate
' ======================================================================
Public Function CorregirYValidarFechaVencimiento(ByVal fechaStr As String) As Object
    On Error GoTo ErrorHandler
    
    Dim resultado As Object
    Set resultado = CreateObject("Scripting.Dictionary")
    resultado("valido") = False
    resultado("valor") = ""
    resultado("mensaje") = ""
    
    ' Limpiar espacios
    fechaStr = Trim(fechaStr)
    
    ' Validar que no esté vacío
    If fechaStr = "" Then
        resultado("mensaje") = "El campo de fecha está vacío."
        Set CorregirYValidarFechaVencimiento = resultado
        Exit Function
    End If
    
    ' Intentar interpretar como fecha usando parseo independiente de locale
    Dim fechaParsed As Variant
    fechaParsed = ParseFechaDMY(fechaStr)
    
    If IsEmpty(fechaParsed) Then
        resultado("mensaje") = "Fecha no válida. Use formato dd-mm-yyyy o dd/mm/yyyy."
        Set CorregirYValidarFechaVencimiento = resultado
        Exit Function
    End If
    
    Dim fechaDate As Date
    fechaDate = fechaParsed  ' ParseFechaDMY ya devuelve Date (no usar CDate que depende del locale)
    
    ' VALIDACIÓN CRÍTICA: La fecha de vencimiento DEBE ser mayor a la fecha actual
    ' Si está vencida, el personal NO puede realizar la inspección
    If fechaDate <= Date Then
        resultado("mensaje") = "La fecha de vencimiento ya pasó (" & Format(fechaDate, "dd/mm/yyyy") & "). " & _
                               "La calificación está vencida y debe renovarse antes de realizar la inspección."
        Set CorregirYValidarFechaVencimiento = resultado
        Exit Function
    End If
    
    ' Convertir al formato dd-mm-yyyy
    Dim fechaCorregida As String
    fechaCorregida = Format(fechaDate, "dd-mm-yyyy")
    
    ' Comparar con original
    Dim esCambio As Boolean
    esCambio = (Format(fechaDate, "dd-mm-yyyy") <> fechaStr)
    
    resultado("valido") = True
    resultado("valor") = fechaCorregida
    
    If esCambio Then
        resultado("mensaje") = "Fecha convertida a formato dd-mm-yyyy: " & fechaCorregida
    Else
        resultado("mensaje") = ""
    End If
    
    Set CorregirYValidarFechaVencimiento = resultado
    Exit Function
    
ErrorHandler:
    resultado("mensaje") = "Error al validar fecha: " & Err.Description
    Call ErrorLogger2.Log("ChecklistValidator.CorregirYValidarFechaVencimiento", Err.Description, Err.Number)
    Set CorregirYValidarFechaVencimiento = resultado
End Function

'' ======================================================================
' Función: CorregirYValidarHora
' Propósito: Intenta convertir entrada de hora a formato HH:MM.
'            Acepta: HH:MM, HHMM, H:MM, etc.
' Parámetros:
'   horaStr: Cadena con hora a corregir
' Retorna: Dictionary con:
'   "valido": True/False
'   "valor": Hora en formato HH:MM (si válido) o vacío (si inválido)
'   "mensaje": Mensaje para el usuario
' ======================================================================
Public Function CorregirYValidarHora(ByVal horaStr As String) As Object
    On Error GoTo ErrorHandler
    
    Dim resultado As Object
    Set resultado = CreateObject("Scripting.Dictionary")
    resultado("valido") = False
    resultado("valor") = ""
    resultado("mensaje") = ""
    
    ' Limpiar espacios
    horaStr = Trim(horaStr)
    
    ' Validar que no esté vacío
    If horaStr = "" Then
        resultado("mensaje") = "El campo de hora está vacío."
        Set CorregirYValidarHora = resultado
        Exit Function
    End If
    
    Dim hh As String
    Dim mm As String
    Dim horaCorregida As String
    Dim esCambio As Boolean
    esCambio = False
    
    ' Caso 1: Formato HH:MM (correcto)
    If Len(horaStr) = 5 And Mid(horaStr, 3, 1) = ":" Then
        hh = Left(horaStr, 2)
        mm = Right(horaStr, 2)
    ' Caso 2: Formato HHMM (sin separador)
    ElseIf Len(horaStr) = 4 And IsNumeric(horaStr) Then
        hh = Left(horaStr, 2)
        mm = Right(horaStr, 2)
        esCambio = True
    ' Caso 3: Formato H:MM o HH:M (sin ceros a la izquierda)
    ElseIf InStr(horaStr, ":") > 0 Then
        Dim partes() As String
        partes = Split(horaStr, ":")
        If UBound(partes) = 1 Then
            On Error Resume Next
            hh = CInt(partes(0))
            mm = CInt(partes(1))
            On Error GoTo ErrorHandler
            esCambio = True
        Else
            resultado("mensaje") = "Hora no válida. Use formato HH:MM (ej: 14:30)."
            Set CorregirYValidarHora = resultado
            Exit Function
        End If
    Else
        resultado("mensaje") = "Hora no válida. Use formato HH:MM (ej: 14:30)."
        Set CorregirYValidarHora = resultado
        Exit Function
    End If
    
    ' Validar que hh y mm sean numéricos y válidos
    If Not IsNumeric(hh) Or Not IsNumeric(mm) Then
        resultado("mensaje") = "Hora contiene caracteres inválidos. Use solo números."
        Set CorregirYValidarHora = resultado
        Exit Function
    End If
    
    Dim hhInt As Integer
    Dim mmInt As Integer
    
    On Error Resume Next
    hhInt = CInt(hh)
    mmInt = CInt(mm)
    On Error GoTo ErrorHandler
    
    ' Validar rangos
    If hhInt < 0 Or hhInt > 23 Then
        resultado("mensaje") = "Hora debe estar entre 00 y 23 (formato 24 horas)."
        Set CorregirYValidarHora = resultado
        Exit Function
    End If
    
    If mmInt < 0 Or mmInt > 59 Then
        resultado("mensaje") = "Minutos deben estar entre 00 y 59."
        Set CorregirYValidarHora = resultado
        Exit Function
    End If
    
    ' Construir hora con ceros a la izquierda
    horaCorregida = Format(hhInt, "00") & ":" & Format(mmInt, "00")
    
    resultado("valido") = True
    resultado("valor") = horaCorregida
    
    If esCambio Then
        resultado("mensaje") = "Hora convertida a formato HH:MM: " & horaCorregida
    Else
        resultado("mensaje") = ""
    End If
    
    Set CorregirYValidarHora = resultado
    Exit Function
    
ErrorHandler:
    resultado("mensaje") = "Error al validar hora: " & Err.Description
    Call ErrorLogger2.Log("ChecklistValidator.CorregirYValidarHora", Err.Description, Err.Number)
    Set CorregirYValidarHora = resultado
End Function

'' ======================================================================
' Función: ValidarCabeceraConAutoCorrecion
' Propósito: Valida cabecera con intento de auto-corrección de fechas y horas.
'            Si hay errores en formato, intenta corregir automáticamente.
' Parámetros:
'   frm: El formulario (frmChecklistVirtual) con los datos
' Retorna: Dictionary con:
'   "valido": True/False
'   "errores": Array de mensajes de error (si hay)
'   "correcciones": Array de mensajes de correcciones aplicadas
' ======================================================================
Public Function ValidarCabeceraConAutoCorrecion(ByRef frm As Object) As Object
    On Error GoTo ErrorHandler
    
    Dim resultado As Object
    Set resultado = CreateObject("Scripting.Dictionary")
    resultado("valido") = False
    resultado("errores") = Array()
    resultado("correcciones") = Array()
    
    Dim errores As Collection
    Dim correcciones As Collection
    Set errores = New Collection
    Set correcciones = New Collection
    
    ' --- Verificar campos obligatorios básicos ---
    If Trim(frm.txtEvaluado.Value) = "" Then
        errores.Add "El campo 'Evaluado' está vacío. Datos del cronograma incorrectos."
    End If
    
    If Trim(frm.txtPuesto.Value) = "" Then
        errores.Add "El campo 'Puesto' está vacío. Datos del cronograma incorrectos."
    End If
    
    If Trim(frm.txtPlanta.Value) = "" Then
        errores.Add "El campo 'Planta' está vacío. Datos del cronograma incorrectos."
    End If
    
    If Trim(frm.cboArea.Value) = "" Then
        errores.Add "Debe seleccionar un Área."
    End If
    
    If Trim(frm.cboLineaAuditada.Value) = "" Then
        errores.Add "Debe seleccionar un Equipo / Línea auditada."
    End If
    
    If Trim(frm.cboEvaluador.Value) = "" Then
        errores.Add "Debe seleccionar un Evaluador."
    End If
    
    If Trim(frm.cboLugar.Value) = "" Then
        errores.Add "Debe seleccionar el Lugar de auditoría (Dentro/Fuera del área)."
    End If
    
    ' --- Validar y corregir FECHA ---
    ' ACTUALIZADO 23/06/2026: NO escribir de vuelta al textbox (.Value)
    ' porque MSForms reinterpreta la fecha con el locale del sistema, corrompiendo el valor.
    ' El texto original del usuario se preserva y ParseFechaDMY lo convierte correctamente.
    Dim resutlFecha As Object
    Set resutlFecha = CorregirYValidarFecha(Trim(frm.txtFecha.Text))
    
    If resutlFecha("valido") Then
        If resutlFecha("mensaje") <> "" Then
            correcciones.Add resutlFecha("mensaje")
        End If
    Else
        errores.Add resutlFecha("mensaje")
    End If
    
    ' --- Validar y corregir FECHA AUDITADA ---
    Dim resultFechaAuditada As Object
    Set resultFechaAuditada = CorregirYValidarFecha(Trim(frm.txtFechaAuditada.Text))
    
    If resultFechaAuditada("valido") Then
        If resultFechaAuditada("mensaje") <> "" Then
            correcciones.Add resultFechaAuditada("mensaje")
        End If
    Else
        errores.Add resultFechaAuditada("mensaje")
    End If
    
    ' --- Validar y corregir HORA INICIO ---
    Dim resultHoraInicio As Object
    Set resultHoraInicio = CorregirYValidarHora(Trim(frm.txtHoraInicio.Value))
    
    If resultHoraInicio("valido") Then
        frm.txtHoraInicio.Value = resultHoraInicio("valor")
        If resultHoraInicio("mensaje") <> "" Then
            correcciones.Add resultHoraInicio("mensaje")
        End If
    Else
        errores.Add resultHoraInicio("mensaje")
    End If
    
    ' --- Validar y corregir HORA TERMINO ---
    Dim resultHoraTermino As Object
    Set resultHoraTermino = CorregirYValidarHora(Trim(frm.txtHoraTermino.Value))
    
    If resultHoraTermino("valido") Then
        frm.txtHoraTermino.Value = resultHoraTermino("valor")
        If resultHoraTermino("mensaje") <> "" Then
            correcciones.Add resultHoraTermino("mensaje")
        End If
    Else
        errores.Add resultHoraTermino("mensaje")
    End If
    
    ' --- Validar coherencia de horas (solo si ambas son válidas) ---
    If resultHoraInicio("valido") And resultHoraTermino("valido") Then
        If resultHoraTermino("valor") <= resultHoraInicio("valor") Then
            errores.Add "La Hora de término (" & resultHoraTermino("valor") & ") debe ser posterior a la Hora de inicio (" & resultHoraInicio("valor") & ")."
        End If
    End If
    
    ' ═══════════════════════════════════════════════════════════════════
    ' VALIDAR CALIFICACIONES (FASE 7 - 23/04/2026)
    ' - Operadores: Vestuario + Operador (ambos)
    ' - Ayudantes y Sanitizador: Solo Vestuario
    ' CORREGIDO 27/04/2026: Validación selectiva según tipo de puesto
    ' ═══════════════════════════════════════════════════════════════════
    If RequiereCalificaciones(frm) Then
        Dim puestoUpper As String
        puestoUpper = UCase(Trim(frm.txtPuesto.Value))
        
        Dim esOperador As Boolean
        esOperador = (InStr(1, puestoUpper, "OPERADOR") > 0)
        
        ' --- Calificación Vestuario (obligatorio para todos) ---
        If Trim(frm.cboCalificacionVestuario.Value) = "" Then
            errores.Add "Debe seleccionar la Calificación de Vestuario (Si/No)."
        End If
        
        ' --- Fecha Vencimiento Vestuario (OBLIGATORIO) ---
        ' ACTUALIZADO 23/06/2026: Usar .Text para evitar conversión Date→String con locale
        Dim fechaVencVestuario As String
        fechaVencVestuario = Trim(frm.txtFechaVencVestuario.Text)
        If fechaVencVestuario = "" Then
            errores.Add "Debe ingresar la Fecha de Vencimiento de Vestuario."
        Else
            Dim resultFechaVencVestuario As Object
            Set resultFechaVencVestuario = CorregirYValidarFechaVencimiento(fechaVencVestuario)
            
            If resultFechaVencVestuario("valido") Then
                ' ACTUALIZADO 23/06/2026: NO escribir de vuelta al textbox
                ' (evita corrupción por reinterpretación MSForms con locale)
                If resultFechaVencVestuario("mensaje") <> "" Then
                    correcciones.Add resultFechaVencVestuario("mensaje")
                End If
            Else
                errores.Add "Fecha Venc. Vestuario: " & resultFechaVencVestuario("mensaje")
            End If
        End If
        
        ' --- Calificación Operador (obligatorio SOLO para Operadores) ---
        If esOperador Then
            If Trim(frm.cboCalificacionOperador.Value) = "" Then
                errores.Add "Debe seleccionar la Calificación de Operador (Si/No)."
            End If
            
            ' --- Fecha Vencimiento Operador (OBLIGATORIO) ---
            ' ACTUALIZADO 23/06/2026: Usar .Text para evitar conversión Date→String con locale
            Dim fechaVencOperador As String
            fechaVencOperador = Trim(frm.txtFechaVencOperador.Text)
            If fechaVencOperador = "" Then
                errores.Add "Debe ingresar la Fecha de Vencimiento de Operador."
            Else
                Dim resultFechaVencOperador As Object
                Set resultFechaVencOperador = CorregirYValidarFechaVencimiento(fechaVencOperador)
                
                If resultFechaVencOperador("valido") Then
                    ' ACTUALIZADO 23/06/2026: NO escribir de vuelta al textbox
                    ' (evita corrupción por reinterpretación MSForms con locale)
                    If resultFechaVencOperador("mensaje") <> "" Then
                        correcciones.Add resultFechaVencOperador("mensaje")
                    End If
                Else
                    errores.Add "Fecha Venc. Operador: " & resultFechaVencOperador("mensaje")
                End If
            End If
        End If
    End If
    
    ' --- Construir resultado final ---
    Dim errorArray() As String
    Dim corrArray() As String
    Dim i As Long
    
    If errores.Count > 0 Then
        ReDim errorArray(errores.Count - 1)
        For i = 1 To errores.Count
            errorArray(i - 1) = errores(i)
        Next i
        resultado("errores") = errorArray
    End If
    
    If correcciones.Count > 0 Then
        ReDim corrArray(correcciones.Count - 1)
        For i = 1 To correcciones.Count
            corrArray(i - 1) = correcciones(i)
        Next i
        resultado("correcciones") = corrArray
    End If
    
    resultado("valido") = (UBound(resultado("errores")) = -1 Or UBound(resultado("errores")) < 0)
    
    Set ValidarCabeceraConAutoCorrecion = resultado
    Exit Function
    
ErrorHandler:
    errores.Add "Error en validación: " & Err.Description
    resultado("errores") = errorArray
    Call ErrorLogger2.Log("ChecklistValidator.ValidarCabeceraConAutoCorrecion", Err.Description, Err.Number)
    Set ValidarCabeceraConAutoCorrecion = resultado
End Function

'' ----------------------------------------------------------------------
' Función: RequiereCalificaciones
' Propósito: Determina si el puesto actual requiere validación de
'            calificaciones de vestuario y operador.
' Parámetros:
'   frm: Referencia al formulario frmChecklistVirtual
' Retorna: True si el puesto contiene: Operador, Ayudante, Sanitizador
' FASE 7 - 23/04/2026
' CORREGIDO 27/04/2026: Usar InStr para detectar puestos con variantes (ej: "Operador Electrolitos")
' ----------------------------------------------------------------------
Private Function RequiereCalificaciones(ByRef frm As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim puestoUpper As String
    puestoUpper = UCase(Trim(frm.txtPuesto.Value))
    
    ' Lista de puestos que requieren calificaciones (usando InStr para detectar variantes)
    RequiereCalificaciones = (InStr(1, puestoUpper, "OPERADOR") > 0 Or _
                             InStr(1, puestoUpper, "AYUDANTE") > 0 Or _
                             InStr(1, puestoUpper, "SANITIZADOR") > 0)
    Exit Function
    
ErrorHandler:
    RequiereCalificaciones = False
    Call ErrorLogger2.Log("ChecklistValidator.RequiereCalificaciones", Err.Description, Err.Number)
End Function
