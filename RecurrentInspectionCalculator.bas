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
'   Fórmula: RPN Promedio + ValorRecuperación(%Rec) + ValorOOL(%OOL)
'
'   Los porcentajes de Recuperación y OOL NO se suman directamente.
'   En su lugar, se convierten a un valor numérico según tablas de
'   rangos definidas en Excel (tblRangosRecuperacion y tblRangosOOL).
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
'   - rpnPromedio debe estar en rango [RPN_MIN, RPN_MAX]
'   - porcRecuperacion >= 0 (puede ser 0 si no aplica)
'   - porcOOL >= 0 (puede ser 0 si no aplica)
'
' Excepciones:
'   - Error 1003: RPN Promedio inválido
'   - Error 1004: Porcentaje Recuperación inválido
'   - Error 1005: Porcentaje OOL inválido
'
' Ejemplo:
'   ' Sin microbiología (todo en 0)
'   rpnTotal = CalcularRPNTotal(rpnPromedio:=78.9) → 78.9
'
'   ' Con microbiología (se convierten por rango)
'   rpnTotal = CalcularRPNTotal(rpnPromedio:=78.9, porcRecuperacion:=0.35, porcOOL:=1.2)
'   → valorRec=6, valorOOL=12, RPN Total = 78.9 + 6 + 12 = 96.9
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
    ' CONVERSIÓN: Obtener valores numéricos desde tablas de rangos
    ' ─────────────────────────────────────────────────────────────────
    Dim valorRecuperacion As Double
    Dim valorOOL As Double
    
    valorRecuperacion = ObtenerValorRecuperacion(porcRecuperacion)
    valorOOL = ObtenerValorOOL(porcOOL)
    
    ' ─────────────────────────────────────────────────────────────────
    ' CÁLCULO: RPN Promedio + valores convertidos de rangos
    ' ─────────────────────────────────────────────────────────────────
    CalcularRPNTotal = rpnPromedio + valorRecuperacion + valorOOL
    
    ' ─────────────────────────────────────────────────────────────────
    ' LOGGING: Registro detallado para auditoría
    ' ─────────────────────────────────────────────────────────────────
    Debug.Print "[RPN RECURRENTE] CalcularRPNTotal():"
    Debug.Print "  RPN Promedio:        " & Format(rpnPromedio, "0.00")
    Debug.Print "  % Recuperación:      " & Format(porcRecuperacion, "0.00") & " → Valor: " & valorRecuperacion
    Debug.Print "  % OOL:               " & Format(porcOOL, "0.00") & " → Valor: " & valorOOL
    Debug.Print "  RPN Total:           " & Format(CalcularRPNTotal, "0.00")
    Debug.Print "  Fórmula: " & Format(rpnPromedio, "0.00") & " + " & valorRecuperacion & " + " & valorOOL & " = " & Format(CalcularRPNTotal, "0.00")
    
    If porcRecuperacion = 0 And porcOOL = 0 Then
        Debug.Print "  [NOTA] Sin datos microbiología (RPN Total = RPN Promedio)"
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
'   rpnTotal (Double) - RPN Total calculado (con factores)
'   iniciales (String) - Iniciales del personal evaluado
'   idPlantilla (String) - ID de la plantilla usada
'
' Retorna:
'   (Long) - Número de categoría (1-5)
'
' Lógica (ACTUALIZADA 16/06/2026):
'   1. Si RPN Total > 20, buscar últimas 2 inspecciones del mismo puesto
'   2. Si las 3 (actual + 2 anteriores) tienen RPN Total > 20 → Categoría 5
'   3. Si no → buscar en tblCategoriasRPN el rango que contiene el RPN Total
'   4. Si ningún rango coincide → Categoría 4 (más restrictiva sin ser Cat 5)
'
' Categorías estándar (según sistema actual):
'   - Categoría 1: RPN 0    - 14   → Aprobado (Sin observaciones)
'   - Categoría 2: RPN 14.1 - 19   → Aprobado con observaciones menores
'   - Categoría 3: RPN 19.1 - 40   → Aprobado con observaciones mayores
'   - Categoría 4: RPN 40.1 - 100   → Condicional
'   - Categoría 5: RPN 20 - 100  → No Calificado (tres veces consecutivas con este resultado de 20)
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
    ByVal rpnTotal As Double, _
    ByVal iniciales As String, _
    ByVal idPlantilla As String _
) As Long
    
    On Error GoTo ErrorHandler
    
    ' ═══════════════════════════════════════════════════════════════
    ' PASO 0: Validar Categoría 5 por historial (3 consecutivas > 20)
    ' ACTUALIZADO (16/06/2026): Evalúa RPN Total (con factores), no %TA puro
    ' ═══════════════════════════════════════════════════════════════
    If rpnTotal > 20 Then
        Dim historial As Collection
        Set historial = InspectionRepository.ObtenerUltimasNInspecciones(iniciales, idPlantilla, 2)
        
        If historial.Count >= 2 Then
            If CDbl(historial(1)) > 20 And CDbl(historial(2)) > 20 Then
                Debug.Print "[RPN RECURRENTE] CATEGORÍA 5: 3 inspecciones consecutivas con RPN Total > 20"
                Debug.Print "  Actual: " & Format(rpnTotal, "0.00") & _
                    " | Anterior 1: " & Format(CDbl(historial(1)), "0.00") & _
                    " | Anterior 2: " & Format(CDbl(historial(2)), "0.00")
                DeterminarCategoriaRPNTotal = 5
                Exit Function
            End If
        End If
    End If
    
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
    
    ' Error de fallback: Si no se puede categorizar, asignar Categoría 4
    ' (la más restrictiva entre las categorías por rango, sin ser Cat 5 que es por historial)
    ' ACTUALIZADO (16/06/2026): Antes asignaba Cat 5 como "precaución", pero Cat 5
    ' solo debe asignarse por la regla de 3 inspecciones consecutivas > 20.
    If Err.Number = vbObjectError + 1008 Then
        Debug.Print "[ERROR CRÍTICO] RPN Total sin categoría en tblCategoriasRPN. Asignando Cat 4 (máxima por rango)"
        DeterminarCategoriaRPNTotal = 4 ' Categoría más restrictiva por rango (no Cat 5)
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

' ══════════════════════════════════════════════════════════════════════
' FUNCIÓN: ObtenerValorRecuperacion
' ══════════════════════════════════════════════════════════════════════
' Descripción:
'   Convierte un % de Recuperación a su valor numérico según la tabla
'   de rangos tblRangosRecuperacion en la hoja Configuración.
'
' Rangos (definidos en Excel):
'   0%              → 0
'   0.01% - 0.25%   → 4
'   0.26% - 0.50%   → 6
'   0.51% - 0.75%   → 8
'   0.76% - 0.99%   → 10
'   1.0% o más      → 12
'
' Parámetros:
'   porcentaje (Double) - % de Recuperación a convertir
'
' Retorna:
'   (Double) - Valor numérico asignado según el rango
' ══════════════════════════════════════════════════════════════════════
Public Function ObtenerValorRecuperacion(ByVal porcentaje As Double) As Double
    
    On Error GoTo ErrorHandler
    
    ' Si es 0, retornar 0 sin consultar la tabla
    If porcentaje = 0 Then
        ObtenerValorRecuperacion = 0
        Exit Function
    End If
    
    Dim ws As Worksheet
    Dim tblRangos As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tblRangos = ws.ListObjects(Configuration2.TABLE_RANGOS_RECUPERACION)
    
    If tblRangos Is Nothing Then
        Debug.Print "[ObtenerValorRecuperacion] ERROR: tblRangosRecuperacion no encontrada"
        ObtenerValorRecuperacion = 0
        Exit Function
    End If
    
    If tblRangos.DataBodyRange Is Nothing Then
        Debug.Print "[ObtenerValorRecuperacion] ADVERTENCIA: tblRangosRecuperacion vacía"
        ObtenerValorRecuperacion = 0
        Exit Function
    End If
    
    ' Buscar índices de columnas (robusto: case-insensitive, accent-insensitive, trim)
    Dim colMin As Long, colMax As Long, colValor As Long
    colMin = 0: colMax = 0: colValor = 0
    
    Dim col As ListColumn
    Dim normalizedName As String
    For Each col In tblRangos.ListColumns
        ' Normalizar: trim + uppercase + reemplazar acentos para matching robusto
        normalizedName = UCase(Trim(col.Name))
        normalizedName = Replace(normalizedName, "Á", "A")
        normalizedName = Replace(normalizedName, "É", "E")
        normalizedName = Replace(normalizedName, "Í", "I")
        normalizedName = Replace(normalizedName, "Ó", "O")
        normalizedName = Replace(normalizedName, "Ú", "U")
        
        If normalizedName = "RANGO MINIMO" Then colMin = col.Index
        If normalizedName = "RANGO MAXIMO" Then colMax = col.Index
        If normalizedName = "VALOR" Then colValor = col.Index
    Next col
    
    If colMin = 0 Or colMax = 0 Or colValor = 0 Then
        ' Diagnóstico: imprimir nombres reales de columnas para facilitar depuración
        Debug.Print "[ObtenerValorRecuperacion] ERROR: Columnas no encontradas (Min=" & colMin & ", Max=" & colMax & ", Valor=" & colValor & ")"
        Debug.Print "[ObtenerValorRecuperacion] Columnas reales en la tabla:"
        For Each col In tblRangos.ListColumns
            Debug.Print "  Col " & col.Index & ": '" & col.Name & "' (normalizado: '" & UCase(Trim(col.Name)) & "')"
        Next col
        ObtenerValorRecuperacion = 0
        Exit Function
    End If
    
    ' Buscar rango que contiene el porcentaje
    Dim rangoMin As Double, rangoMax As Double
    For Each fila In tblRangos.ListRows
        rangoMin = CDbl(fila.Range.Cells(1, colMin).Value)
        rangoMax = CDbl(fila.Range.Cells(1, colMax).Value)
        
        If porcentaje >= rangoMin And porcentaje <= rangoMax Then
            ObtenerValorRecuperacion = CDbl(fila.Range.Cells(1, colValor).Value)
            
            Debug.Print "[ObtenerValorRecuperacion] % Recuperación: " & Format(porcentaje, "0.00") & _
                        " → Rango [" & rangoMin & " - " & rangoMax & "] → Valor: " & ObtenerValorRecuperacion
            Exit Function
        End If
    Next fila
    
    ' Fallback: si no encuentra rango, retornar 0
    Debug.Print "[ObtenerValorRecuperacion] ADVERTENCIA: % " & Format(porcentaje, "0.00") & " no coincide con ningún rango. Retornando 0."
    ObtenerValorRecuperacion = 0
    Exit Function
    
ErrorHandler:
    Call ErrorLogger2.Log(MODULO_NOMBRE & ".ObtenerValorRecuperacion", Err.Description, Err.Number)
    ObtenerValorRecuperacion = 0
End Function

' ══════════════════════════════════════════════════════════════════════
' FUNCIÓN: ObtenerValorOOL
' ══════════════════════════════════════════════════════════════════════
' Descripción:
'   Convierte un % de OOL a su valor numérico según la tabla
'   de rangos tblRangosOOL en la hoja Configuración.
'
' Rangos (definidos en Excel):
'   0%              → 0
'   0.01% - 0.50%   → 4
'   0.51% - 1.00%   → 8
'   1.01% - 2.00%   → 12
'   2.01% - 3.00%   → 16
'   3.01% o más     → 20
'
' Parámetros:
'   porcentaje (Double) - % de OOL a convertir
'
' Retorna:
'   (Double) - Valor numérico asignado según el rango
' ══════════════════════════════════════════════════════════════════════
Public Function ObtenerValorOOL(ByVal porcentaje As Double) As Double
    
    On Error GoTo ErrorHandler
    
    ' Si es 0, retornar 0 sin consultar la tabla
    If porcentaje = 0 Then
        ObtenerValorOOL = 0
        Exit Function
    End If
    
    Dim ws As Worksheet
    Dim tblRangos As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets(Configuration2.SHEET_CONFIGURACION)
    Set tblRangos = ws.ListObjects(Configuration2.TABLE_RANGOS_OOL)
    
    If tblRangos Is Nothing Then
        Debug.Print "[ObtenerValorOOL] ERROR: tblRangosOOL no encontrada"
        ObtenerValorOOL = 0
        Exit Function
    End If
    
    If tblRangos.DataBodyRange Is Nothing Then
        Debug.Print "[ObtenerValorOOL] ADVERTENCIA: tblRangosOOL vacía"
        ObtenerValorOOL = 0
        Exit Function
    End If
    
    ' Buscar índices de columnas (robusto: case-insensitive, accent-insensitive, trim)
    Dim colMin As Long, colMax As Long, colValor As Long
    colMin = 0: colMax = 0: colValor = 0
    
    Dim col As ListColumn
    Dim normalizedName As String
    For Each col In tblRangos.ListColumns
        ' Normalizar: trim + uppercase + reemplazar acentos para matching robusto
        normalizedName = UCase(Trim(col.Name))
        normalizedName = Replace(normalizedName, "Á", "A")
        normalizedName = Replace(normalizedName, "É", "E")
        normalizedName = Replace(normalizedName, "Í", "I")
        normalizedName = Replace(normalizedName, "Ó", "O")
        normalizedName = Replace(normalizedName, "Ú", "U")
        
        If normalizedName = "RANGO MINIMO" Then colMin = col.Index
        If normalizedName = "RANGO MAXIMO" Then colMax = col.Index
        If normalizedName = "VALOR" Then colValor = col.Index
    Next col
    
    If colMin = 0 Or colMax = 0 Or colValor = 0 Then
        ' Diagnóstico: imprimir nombres reales de columnas para facilitar depuración
        Debug.Print "[ObtenerValorOOL] ERROR: Columnas no encontradas (Min=" & colMin & ", Max=" & colMax & ", Valor=" & colValor & ")"
        Debug.Print "[ObtenerValorOOL] Columnas reales en la tabla:"
        For Each col In tblRangos.ListColumns
            Debug.Print "  Col " & col.Index & ": '" & col.Name & "' (normalizado: '" & UCase(Trim(col.Name)) & "')"
        Next col
        ObtenerValorOOL = 0
        Exit Function
    End If
    
    ' Buscar rango que contiene el porcentaje
    Dim rangoMin As Double, rangoMax As Double
    For Each fila In tblRangos.ListRows
        rangoMin = CDbl(fila.Range.Cells(1, colMin).Value)
        rangoMax = CDbl(fila.Range.Cells(1, colMax).Value)
        
        If porcentaje >= rangoMin And porcentaje <= rangoMax Then
            ObtenerValorOOL = CDbl(fila.Range.Cells(1, colValor).Value)
            
            Debug.Print "[ObtenerValorOOL] % OOL: " & Format(porcentaje, "0.00") & _
                        " → Rango [" & rangoMin & " - " & rangoMax & "] → Valor: " & ObtenerValorOOL
            Exit Function
        End If
    Next fila
    
    ' Fallback: si no encuentra rango, retornar 0
    Debug.Print "[ObtenerValorOOL] ADVERTENCIA: % " & Format(porcentaje, "0.00") & " no coincide con ningún rango. Retornando 0."
    ObtenerValorOOL = 0
    Exit Function
    
ErrorHandler:
    Call ErrorLogger2.Log(MODULO_NOMBRE & ".ObtenerValorOOL", Err.Description, Err.Number)
    ObtenerValorOOL = 0
End Function
