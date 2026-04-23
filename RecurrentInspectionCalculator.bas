' ══════════════════════════════════════════════════════════════════════
' Módulo: RecurrentInspectionCalculator
' Descripción: Cálculos de RPN para inspecciones recurrentes (2da inspección en adelante)
' Fecha creación: 21/04/2026
' Autor: Sistema TH-HC-001
' Versión: 1.0.0
' ══════════════════════════════════════════════════════════════════════
'
' Responsabilidad:
'   - Calcular RPN Promedio entre inspecciones
'   - Calcular RPN Total (preparado para futura integración de microbiología)
'   - Determinar categoría basada en RPN Total
'   - Validaciones de coherencia de datos
'
' Dependencias:
'   - ErrorLogger2: Logging de errores
'   - Configuration2: Acceso a tblCategoriasRPN
'   - VariablesGlobales2: Constantes del sistema
'
' Uso:
'   Dim rpnProm As Double
'   rpnProm = CalcularRPNPromedio(rpnAnterior:=75.5, rpnActual:=82.3)
'   
'   Dim rpnTotal As Double
'   rpnTotal = CalcularRPNTotal(rpnPromedio:=rpnProm)
'   
'   Dim categoria As Long
'   categoria = DeterminarCategoriaRPNTotal(rpnTotal:=rpnTotal)
'
' ══════════════════════════════════════════════════════════════════════
Option Explicit

' ══════════════════════════════════════════════════════════════════════
' CONSTANTES DE VALIDACIÓN
' ══════════════════════════════════════════════════════════════════════
Private Const RPN_MIN As Double = 0#      ' RPN mínimo válido (0 = desempeño perfecto)
Private Const RPN_MAX As Double = 100#    ' RPN máximo válido
Private Const MODULO_NOMBRE As String = "RecurrentInspectionCalculator"

' ══════════════════════════════════════════════════════════════════════
' FUNCIÓN: CalcularRPNPromedio
' ══════════════════════════════════════════════════════════════════════
' Descripción:
'   Calcula el RPN Promedio para inspecciones recurrentes.
'   Fórmula: (RPN Anterior + RPN Actual) / 2
'
' Parámetros:
'   rpnAnterior (Double) - RPN de la inspección anterior (puede ser manual o del sistema)
'   rpnActual (Double)   - RPN de la inspección actual (calculado desde TA porcentaje)
'
' Retorna:
'   (Double) - Promedio aritmético simple de ambos valores
'
' Validaciones:
'   - Ambos valores deben ser > 0
'   - Ambos valores dentro de rango [RPN_MIN, RPN_MAX]
'
' Excepciones:
'   - Error 1001: RPN anterior inválido
'   - Error 1002: RPN actual inválido
'
' Ejemplo:
'   rpnProm = CalcularRPNPromedio(75.5, 82.3) → 78.9
' ══════════════════════════════════════════════════════════════════════
Public Function CalcularRPNPromedio( _
    ByVal rpnAnterior As Double, _
    ByVal rpnActual As Double _
) As Double
    
    On Error GoTo ErrorHandler
    
    ' ─────────────────────────────────────────────────────────────────
    ' VALIDACIÓN #1: RPN Anterior
    ' ─────────────────────────────────────────────────────────────────
    If rpnAnterior < 0 Then
        Err.Raise vbObjectError + 1001, MODULO_NOMBRE & ".CalcularRPNPromedio", _
            "RPN Anterior no puede ser negativo. Valor recibido: " & rpnAnterior
    End If
    
    If rpnAnterior < RPN_MIN Or rpnAnterior > RPN_MAX Then
        Err.Raise vbObjectError + 1001, MODULO_NOMBRE & ".CalcularRPNPromedio", _
            "RPN Anterior fuera de rango válido [" & RPN_MIN & " - " & RPN_MAX & "]. " & _
            "Valor recibido: " & rpnAnterior
    End If
    
    ' ─────────────────────────────────────────────────────────────────
    ' VALIDACIÓN #2: RPN Actual
    ' ─────────────────────────────────────────────────────────────────
    If rpnActual < 0 Then
        Err.Raise vbObjectError + 1002, MODULO_NOMBRE & ".CalcularRPNPromedio", _
            "RPN Actual no puede ser negativo. Valor recibido: " & rpnActual
    End If
    
    If rpnActual < RPN_MIN Or rpnActual > RPN_MAX Then
        Err.Raise vbObjectError + 1002, MODULO_NOMBRE & ".CalcularRPNPromedio", _
            "RPN Actual fuera de rango válido [" & RPN_MIN & " - " & RPN_MAX & "]. " & _
            "Valor recibido: " & rpnActual
    End If
    
    ' ─────────────────────────────────────────────────────────────────
    ' CÁLCULO: Promedio aritmético simple
    ' ─────────────────────────────────────────────────────────────────
    CalcularRPNPromedio = (rpnAnterior + rpnActual) / 2
    
    ' ─────────────────────────────────────────────────────────────────
    ' LOGGING: Registro detallado para auditoría
    ' ─────────────────────────────────────────────────────────────────
    Debug.Print "[RPN RECURRENTE] CalcularRPNPromedio():"
    Debug.Print "  RPN Anterior: " & Format(rpnAnterior, "0.00")
    Debug.Print "  RPN Actual:   " & Format(rpnActual, "0.00")
    Debug.Print "  RPN Promedio: " & Format(CalcularRPNPromedio, "0.00")
    
    Exit Function
    
ErrorHandler:
    ' Registrar error y re-lanzar
    Call ErrorLogger2.Log(MODULO_NOMBRE & ".CalcularRPNPromedio", Err.Description, Err.Number)
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ══════════════════════════════════════════════════════════════════════
' FUNCIÓN: CalcularRPNTotal
' ══════════════════════════════════════════════════════════════════════
' Descripción:
'   Calcula el RPN Total para inspecciones recurrentes.
'   Fórmula: RPN Promedio + % Recuperación + % OOL
'
'   FASE ACTUAL: Solo usa RPN Promedio (% Recuperación y % OOL = 0)
'   FASE FUTURA: Integrará datos de microbiología cuando estén disponibles
'
' Parámetros:
'   rpnPromedio (Double)       - RPN Promedio calculado (requerido)
'   porcRecuperacion (Double)  - % Recuperación microbiología (opcional, default = 0)
'   porcOOL (Double)           - % Out of Limits microbiología (opcional, default = 0)
'
' Retorna:
'   (Double) - RPN Total resultante
'
' Validaciones:
'   - rpnPromedio debe ser > 0
'   - porcRecuperacion >= 0 (puede ser 0 si no aplica)
'   - porcOOL >= 0 (puede ser 0 si no aplica)
'
' Excepciones:
'   - Error 1003: RPN Promedio inválido
'   - Error 1004: Porcentaje Recuperación inválido
'   - Error 1005: Porcentaje OOL inválido
'
' Ejemplo:
'   ' Fase actual (sin microbiología)
'   rpnTotal = CalcularRPNTotal(rpnPromedio:=78.9) → 78.9
'   
'   ' Fase futura (con microbiología)
'   rpnTotal = CalcularRPNTotal(rpnPromedio:=78.9, porcRecuperacion:=5.2, porcOOL:=2.1) → 86.2
' ══════════════════════════════════════════════════════════════════════
Public Function CalcularRPNTotal( _
    ByVal rpnPromedio As Double, _
    Optional ByVal porcRecuperacion As Double = 0, _
    Optional ByVal porcOOL As Double = 0 _
) As Double
    
    On Error GoTo ErrorHandler
    
    ' ─────────────────────────────────────────────────────────────────
    ' VALIDACIÓN #1: RPN Promedio
    ' ─────────────────────────────────────────────────────────────────
    If rpnPromedio < 0 Then
        Err.Raise vbObjectError + 1003, MODULO_NOMBRE & ".CalcularRPNTotal", _
            "RPN Promedio no puede ser negativo. Valor recibido: " & rpnPromedio
    End If
    
    If rpnPromedio < RPN_MIN Or rpnPromedio > RPN_MAX Then
        Err.Raise vbObjectError + 1003, MODULO_NOMBRE & ".CalcularRPNTotal", _
            "RPN Promedio fuera de rango válido [" & RPN_MIN & " - " & RPN_MAX & "]. " & _
            "Valor recibido: " & rpnPromedio
    End If
    
    ' ─────────────────────────────────────────────────────────────────
    ' VALIDACIÓN #2: % Recuperación (debe ser >= 0)
    ' ─────────────────────────────────────────────────────────────────
    If porcRecuperacion < 0 Then
        Err.Raise vbObjectError + 1004, MODULO_NOMBRE & ".CalcularRPNTotal", _
            "Porcentaje Recuperación no puede ser negativo. Valor recibido: " & porcRecuperacion
    End If
    
    ' ─────────────────────────────────────────────────────────────────
    ' VALIDACIÓN #3: % OOL (debe ser >= 0)
    ' ─────────────────────────────────────────────────────────────────
    If porcOOL < 0 Then
        Err.Raise vbObjectError + 1005, MODULO_NOMBRE & ".CalcularRPNTotal", _
            "Porcentaje OOL no puede ser negativo. Valor recibido: " & porcOOL
    End If
    
    ' ─────────────────────────────────────────────────────────────────
    ' CÁLCULO: Suma de componentes
    ' ─────────────────────────────────────────────────────────────────
    CalcularRPNTotal = rpnPromedio + porcRecuperacion + porcOOL
    
    ' ─────────────────────────────────────────────────────────────────
    ' LOGGING: Registro detallado para auditoría
    ' ─────────────────────────────────────────────────────────────────
    Debug.Print "[RPN RECURRENTE] CalcularRPNTotal():"
    Debug.Print "  RPN Promedio:        " & Format(rpnPromedio, "0.00")
    Debug.Print "  % Recuperación:      " & Format(porcRecuperacion, "0.00")
    Debug.Print "  % OOL:               " & Format(porcOOL, "0.00")
    Debug.Print "  RPN Total:           " & Format(CalcularRPNTotal, "0.00")
    
    If porcRecuperacion = 0 And porcOOL = 0 Then
        Debug.Print "  [NOTA] Modo actual: Sin datos microbiología (RPN Total = RPN Promedio)"
    Else
        Debug.Print "  [NOTA] Modo futuro: Con datos microbiología integrados"
    End If
    
    Exit Function
    
ErrorHandler:
    ' Registrar error y re-lanzar
    Call ErrorLogger2.Log(MODULO_NOMBRE & ".CalcularRPNTotal", Err.Description, Err.Number)
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ══════════════════════════════════════════════════════════════════════
' FUNCIÓN: DeterminarCategoriaRPNTotal
' ══════════════════════════════════════════════════════════════════════
' Descripción:
'   Determina la categoría de resultado basada en el RPN Total.
'   Usa la misma tabla de categorización (tblCategoriasRPN) que
'   el sistema original, garantizando consistencia.
'
' Parámetros:
'   rpnTotal (Double) - RPN Total calculado
'
' Retorna:
'   (Long) - Número de categoría (1-5)
'
' Lógica:
'   1. Lee tblCategoriasRPN (columnas: RPN minimo, RPN maximo, Numero categoria)
'   2. Busca en qué rango cae el rpnTotal
'   3. Retorna el número de categoría correspondiente
'
' Categorías estándar (según sistema actual):
'   - Categoría 1: RPN 0    - 20   → Aprobado (Sin observaciones)
'   - Categoría 2: RPN 20.1 - 40   → Aprobado con observaciones menores
'   - Categoría 3: RPN 40.1 - 60   → Aprobado con observaciones mayores
'   - Categoría 4: RPN 60.1 - 80   → Condicional
'   - Categoría 5: RPN 80.1 - 100  → No Calificado
'
' Validaciones:
'   - rpnTotal >= 0
'   - tblCategoriasRPN existe y contiene datos
'
' Excepciones:
'   - Error 1006: RPN Total inválido
'   - Error 1007: Tabla de categorías no encontrada
'   - Error 1008: No se encontró categoría para el valor
'
' Ejemplo:
'   categoria = DeterminarCategoriaRPNTotal(78.9) → 4 (Condicional)
' ══════════════════════════════════════════════════════════════════════
Public Function DeterminarCategoriaRPNTotal( _
    ByVal rpnTotal As Double _
) As Long
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tblCategorias As ListObject
    Dim fila As ListRow
    Dim rangoInferior As Double
    Dim rangoSuperior As Double
    Dim categoriaEncontrada As Boolean
    
    ' ─────────────────────────────────────────────────────────────────
    ' VALIDACIÓN #1: RPN Total
    ' ─────────────────────────────────────────────────────────────────
    If rpnTotal < 0 Then
        Err.Raise vbObjectError + 1006, MODULO_NOMBRE & ".DeterminarCategoriaRPNTotal", _
            "RPN Total no puede ser negativo. Valor recibido: " & rpnTotal
    End If
    
    ' ─────────────────────────────────────────────────────────────────
    ' ACCESO A TABLA: tblCategoriasRPN
    ' ─────────────────────────────────────────────────────────────────
    Debug.Print "[DeterminarCategoriaRPNTotal] Accediendo a tblCategoriasRPN..."
    Debug.Print "  RPN Total a clasificar: " & Format(rpnTotal, "0.00")
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tblCategorias = ws.ListObjects(Configuration2.TABLE_CATEGORIAS_RPN)
    On Error GoTo ErrorHandler
    
    If tblCategorias Is Nothing Then
        Debug.Print "[ERROR] tblCategoriasRPN no encontrada"
        Err.Raise vbObjectError + 1007, MODULO_NOMBRE & ".DeterminarCategoriaRPNTotal", _
            "No se pudo acceder a la tabla tblCategoriasRPN en la hoja Configuracion"
    End If
    
    Debug.Print "  Tabla encontrada. Filas: " & tblCategorias.ListRows.Count
    
    ' Verificar que las columnas existen
    Dim colRPNMin As ListColumn
    Dim colRPNMax As ListColumn
    Dim colNumCat As ListColumn
    
    On Error Resume Next
    Set colRPNMin = tblCategorias.ListColumns("RPN minimo")
    Set colRPNMax = tblCategorias.ListColumns("RPN maximo")
    Set colNumCat = tblCategorias.ListColumns("Numero categoria")
    On Error GoTo ErrorHandler
    
    If colRPNMin Is Nothing Then
        Debug.Print "[ERROR] Columna 'RPN minimo' no encontrada"
        Err.Raise vbObjectError + 1009, MODULO_NOMBRE & ".DeterminarCategoriaRPNTotal", _
            "Columna 'RPN minimo' no existe en tblCategoriasRPN"
    End If
    
    If colRPNMax Is Nothing Then
        Debug.Print "[ERROR] Columna 'RPN maximo' no encontrada"
        Err.Raise vbObjectError + 1010, MODULO_NOMBRE & ".DeterminarCategoriaRPNTotal", _
            "Columna 'RPN maximo' no existe en tblCategoriasRPN"
    End If
    
    If colNumCat Is Nothing Then
        Debug.Print "[ERROR] Columna 'Numero categoria' no encontrada"
        Err.Raise vbObjectError + 1011, MODULO_NOMBRE & ".DeterminarCategoriaRPNTotal", _
            "Columna 'Numero categoria' no existe en tblCategoriasRPN"
    End If
    
    Debug.Print "  Columnas verificadas OK"
    
    ' ─────────────────────────────────────────────────────────────────
    ' BÚSQUEDA: Iterar por cada fila de tabla
    ' ─────────────────────────────────────────────────────────────────
    categoriaEncontrada = False
    
    For Each fila In tblCategorias.ListRows
        On Error Resume Next
        rangoInferior = CDbl(fila.Range.Cells(1, colRPNMin.Index).Value)
        rangoSuperior = CDbl(fila.Range.Cells(1, colRPNMax.Index).Value)
        On Error GoTo ErrorHandler
        
        Debug.Print "  Verificando rango [" & rangoInferior & " - " & rangoSuperior & "]"
        
        ' Verificar si rpnTotal está dentro del rango [inferior, superior]
        If rpnTotal >= rangoInferior And rpnTotal <= rangoSuperior Then
            DeterminarCategoriaRPNTotal = CLng(fila.Range.Cells(1, colNumCat.Index).Value)
            categoriaEncontrada = True
            
            ' Logging detallado
            Debug.Print "[RPN RECURRENTE] DeterminarCategoriaRPNTotal():"
            Debug.Print "  RPN Total:        " & Format(rpnTotal, "0.00")
            Debug.Print "  Rango encontrado: [" & rangoInferior & " - " & rangoSuperior & "]"
            Debug.Print "  Categoría:        " & DeterminarCategoriaRPNTotal
            
            Exit For
        End If
    Next fila
    
    ' ─────────────────────────────────────────────────────────────────
    ' VALIDACIÓN POST-BÚSQUEDA: Categoría encontrada
    ' ─────────────────────────────────────────────────────────────────
    If Not categoriaEncontrada Then
        Err.Raise vbObjectError + 1008, MODULO_NOMBRE & ".DeterminarCategoriaRPNTotal", _
            "No se encontró categoría para RPN Total = " & rpnTotal & ". " & _
            "Verificar configuración de tblCategoriasRPN."
    End If
    
    Exit Function
    
ErrorHandler:
    ' Registrar error y re-lanzar
    Call ErrorLogger2.Log(MODULO_NOMBRE & ".DeterminarCategoriaRPNTotal", Err.Description, Err.Number)
    
    ' Error de fallback: Si no se puede categorizar, asignar categoría más conservadora
    If Err.Number = vbObjectError + 1008 Then
        Debug.Print "[ERROR CRÍTICO] RPN Total sin categoría. Asignando Cat 5 por precaución"
        DeterminarCategoriaRPNTotal = 5 ' Categoría más restrictiva
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

' ══════════════════════════════════════════════════════════════════════
' FUNCIÓN: ValidarConsistenciaRPN
' ══════════════════════════════════════════════════════════════════════
' Descripción:
'   Función auxiliar para validar consistencia entre RPN Anterior y RPN Actual.
'   Advertencia si la diferencia es > 50% (puede indicar error de captura).
'
' Parámetros:
'   rpnAnterior (Double) - RPN de inspección anterior
'   rpnActual (Double)   - RPN de inspección actual
'
' Retorna:
'   (Boolean) - True si es consistente, False si hay advertencia
'
' Ejemplo:
'   ValidarConsistenciaRPN(75, 80) → True (diferencia razonable)
'   ValidarConsistenciaRPN(20, 85) → False (advertencia: cambio brusco)
' ══════════════════════════════════════════════════════════════════════
Public Function ValidarConsistenciaRPN( _
    ByVal rpnAnterior As Double, _
    ByVal rpnActual As Double _
) As Boolean
    
    Dim diferencia As Double
    Dim porcCambio As Double
    
    ' Caso especial: Si RPN Anterior = 0 (desempeño perfecto previo)
    ' No se puede calcular porcentaje de cambio
    If rpnAnterior = 0 Then
        If rpnActual > 0 Then
            Debug.Print "[INFO] RPN cambió de 0 (perfecto) a " & Format(rpnActual, "0.00")
            Debug.Print "  Esto es normal - indica deterioro desde desempeño perfecto"
        End If
        ValidarConsistenciaRPN = True  ' No es error, es caso válido
        Exit Function
    End If
    
    ' Calcular diferencia absoluta y porcentual
    diferencia = Abs(rpnActual - rpnAnterior)
    porcCambio = (diferencia / rpnAnterior) * 100
    
    ' Umbral de advertencia: 50%
    If porcCambio > 50 Then
        Debug.Print "[ADVERTENCIA] Cambio brusco en RPN detectado:"
        Debug.Print "  RPN Anterior: " & Format(rpnAnterior, "0.00")
        Debug.Print "  RPN Actual:   " & Format(rpnActual, "0.00")
        Debug.Print "  Diferencia:   " & Format(diferencia, "0.00") & " (" & Format(porcCambio, "0.0") & "%)"
        Debug.Print "  Recomendación: Validar que datos sean correctos"
        
        ValidarConsistenciaRPN = False
    Else
        ValidarConsistenciaRPN = True
    End If
End Function
