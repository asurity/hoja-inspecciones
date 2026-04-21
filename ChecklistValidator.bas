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
    
    If Not IsDate(fechaStr) Then
        ValidarCabecera = "La Fecha evaluada no tiene un formato válido (DD/MM/AAAA)."
        Exit Function
    End If
    
    If CDate(fechaStr) > Date Then
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
    
    If Not IsDate(fechaAuditadaStr) Then
        ValidarCabecera = "La Fecha Auditada no tiene un formato válido (DD/MM/AAAA)."
        Exit Function
    End If
    
    If CDate(fechaAuditadaStr) > Date Then
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
