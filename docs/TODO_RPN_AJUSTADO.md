# 📐 PLAN DETALLADO: Sistema de Inspecciones Recurrentes con RPN Histórico

## 🎯 Objetivo
Implementar sistema de inspecciones recurrentes que permita:
- Identificar inspecciones 2da, 3ra, etc. del mismo personal
- Capturar RPN anterior (manual o automático desde histórico)
- Calcular RPN Promedio y preparar estructura para RPN Total
- Mantener retrocompatibilidad total con flujo actual (1ra inspección)

---

## 🏗️ Arquitectura de la Solución

### Principios aplicados:
- **Clean Architecture**: Separación de capas (UI → Business Logic → Data)
- **DRY**: Funciones reutilizables para cálculos de RPN
- **Open/Closed**: Extensión sin modificación de código existente
- **Single Responsibility**: Cada módulo con responsabilidad única
- **Desacoplamiento**: Servicios independientes comunicados por interfaces

### Estructura modular propuesta:
```
InspectionHistoryService.bas    → Nuevo: Consultas de histórico
RecurrentInspectionCalculator.bas → Nuevo: Lógica RPN recurrente
frmChecklistVirtual.frm         → Modificado: UI controles recurrencia
ChecklistOrchestrator.bas       → Modificado: Pipeline decisión RPN
InspectionRepository.bas        → Modificado: Queries búsqueda histórico
TableManager.bas                → Modificado: Agregar columnas nuevas
```

---

## 📊 Fase 0: Análisis y Preparación (CHECKPOINT #0)

### 0.1 Análisis de Configuration2.bas
**CRÍTICO:** Validar estructura actual de tablas

**Tareas:**
- [x] Leer Configuration2.bas líneas 260-600 (estructuras de tablas)
- [ ] Validar que documentación coincide con Excel real
- [ ] Identificar todas las referencias a columnas por índice vs por nombre
- [ ] Mapear dependencias de columnas críticas:
  - `tblInspecciones[10]` - Iniciales personal
  - `tblInspecciones[13]` - Fecha inspeccion
  - `tblInspecciones[20]` - RPN calculado
  - `tblInspecciones[21]` - Categoria resultado

**Output esperado:**
- Tabla comparativa: Configuration2 vs Excel real
- Lista de módulos que usan índices hardcodeados
- Identificación de puntos de ruptura potenciales

---

### 0.2 Análisis de impacto en código ✅ COMPLETADO

#### ✅ HALLAZGO #1: Columna Extra en tblInspecciones
**DESCUBIERTO:** tblInspecciones tiene **31 columnas**, no 30 como documenta Configuration2.bas

**Columna nueva encontrada:**
- **[20] Auditoria Procesos Resultado** (String) - Entre "TA porcentaje" y "RPN calculado"
- Mencionada en Q2/Q3: "Auditoría Procesos NO se toca, queda separada"
- Configuration2.bas está DESACTUALIZADO

**Estructura REAL de Excel:**
```
[01-19] Columnas originales (igual que Configuration2)
[20] Auditoria Procesos Resultado  ← NUEVA (no documentada)
[21] RPN calculado                 ← Era col 20 en Configuration2
[22] Categoria resultado           ← Era col 21 en Configuration2
[23-31] Resto de columnas originales
```

**IMPACTO EN EL PLAN:**
- ❌ Las 9 nuevas columnas NO serán 31-39
- ✅ Las 9 nuevas columnas serán **32-40**
- ✅ Actualizar Configuration2.bas para documentar col [20]

---

#### ⚠️ HALLAZGO #2: BUG CRÍTICO en CertificadoPDFGenerator.bas

**ARCHIVO:** `CertificadoPDFGenerator.bas` líneas 220-267  
**FUNCIÓN:** `ObtenerDatosInspeccion()`

**PROBLEMA:** Usa índices HARDCODEADOS que son INCORRECTOS ahora:
```vb
' Línea 240-241 - INCORRECTO ❌
datos("RPN") = fila.Cells(1, 20).Value          ' Lee Auditoria Procesos! ❌
datos("Categoria") = fila.Cells(1, 21).Value    ' Lee RPN calculado! ❌

' Línea 267 - INCORRECTO ❌
datos("ObservacionesGenerales") = fila.Cells(1, 26).Value  ' Lee col equivocada ❌
```

**DEBERÍA SER:**
```vb
' CORRECTO ✅
datos("RPN") = fila.Cells(1, 21).Value          ' RPN calculado
datos("Categoria") = fila.Cells(1, 22).Value    ' Categoria resultado
datos("ObservacionesGenerales") = fila.Cells(1, 27).Value  ' Observaciones generales
```

**NOTA POSITIVA:** El mismo código YA tiene lógica flexible para "Auditoria Procesos Resultado" usando `Application.Match()` (líneas 245-263). Solo faltan RPN/Categoria/Observaciones.

**ACCIÓN REQUERIDA:**
- ✅ **FIX INMEDIATO antes de continuar:** Cambiar líneas 240, 241, 267 por `.ListColumns().Index`
- ✅ Esto es BLOQUEANTE para el resto del plan

---

#### ✅ HALLAZGO #3: ObtenerUltimasNInspecciones() YA EXISTE

**ARCHIVO:** `InspectionRepository.bas` líneas 291-393  
**FUNCIÓN:** `ObtenerUltimasNInspecciones(iniciales, idPlantilla, n)`

**DESCUBRIMIENTO CLAVE:**
- ✅ La función historical query YA EXISTE y funciona
- ✅ Usada por `DeterminarCategoria()` para validar Cat 5 (3 RPN > 20 consecutivos)
- ✅ Ya ordena por fecha DESC y retorna Collection de RPNs
- ✅ Filtra por Iniciales + ID Plantilla

**LIMITACIÓN ACTUAL:**
- ❌ Solo filtra por Iniciales + ID Plantilla
- ❌ NO filtra por "Puesto" (porque no existe esa columna aún)
- ❌ Para Q1 (PUESTO + INICIALES), necesitamos agregar columna [33] "Puesto Evaluado"

**REUTILIZACIÓN:**
- ✅ Modificar esta función para agregar filtro por Puesto (columna 33)
- ✅ No necesitamos crear función nueva, solo extender existente

---

#### ✅ HALLAZGO #4: DeterminarCategoria() tiene lógica de recurrencia

**ARCHIVO:** `InspectionCalculator.bas` líneas 193-250  
**FUNCIÓN:** `DeterminarCategoria(rpn, iniciales, idPlantilla)`

**LÓGICA ACTUAL:**
```vb
' Si RPN > 20:
'   1. Buscar últimas 2 inspecciones (ObtenerUltimasNInspecciones)
'   2. Si las 3 (actual + 2 anteriores) tienen RPN > 20 → Cat 5
'   3. Si no → Categorizar por tabla tblCategoriasRPN
```

**IMPLICACIONES:**
- ✅ Ya existe un PRECEDENTE de sistema recurrente parcial
- ✅ Solo valida Cat 5, no afecta RPN como tal
- ⚠️ Para nuestro plan: Necesitamos extender esta lógica para:
  - Detectar si es 2da+ inspección
  - Llamar a RecurrentInspectionCalculator si aplica
  - Usar RPN Total en lugar de RPN para categorización

---

#### ✅ HALLAZGO #5: CalcularRPN() es simple y NO requiere cambios

**ARCHIVO:** `InspectionCalculator.bas` línea 179  
**FIRMA:** `Public Function CalcularRPN(ByVal taData As Object) As Double`

**LÓGICA:**
```vb
CalcularRPN = CDbl(taData("porcentaje"))  ' RPN = TA porcentaje
```

**CONCLUSIÓN:**
- ✅ NO necesita modificación
- ✅ La lógica de RPN Promedio será EXTERNA en `RecurrentInspectionCalculator.bas`
- ✅ CalcularRPN() sigue calculando RPN individual (Actual)
- ✅ ChecklistOrchestrator decidirá si usar RPN o RPN Total para categorización

---

#### ✅ HALLAZGO #6: InspectionRepository usa nombres de columna (BIEN)

**ARCHIVO:** `InspectionRepository.bas`  
**FUNCIONES:** `CrearInspeccion()`, `ActualizarCalculosInspeccion()`

**PATRÓN USADO:**
```vb
.Cells(1, tblInspecciones.ListColumns("Nombre columna").Index).Value = valor
```

**CONCLUSIÓN:**
- ✅ NO usa índices hardcodeados
- ✅ SEGURO ante cambios de estructura
- ✅ Solo necesita agregar las 9 nuevas columnas al Dictionary de datos

---

### 📋 RESUMEN DE ACCIONES FASE 0

**BLOQUEANTES (FIX INMEDIATO):**
1. ❌ Corregir CertificadoPDFGenerator.bas líneas 240, 241, 267 (índices hardcodeados)
2. ❌ Actualizar Configuration2.bas para documentar columna [20] "Auditoria Procesos Resultado"

**AJUSTES AL PLAN:**
3. ✅ Cambiar numeración: Nuevas columnas serán 32-40 (no 31-39)
4. ✅ Modificar ObtenerUltimasNInspecciones() para agregar filtro por Puesto
5. ✅ Extender DeterminarCategoria() para usar RPN Total en inspecciones recurrentes

**REUTILIZACIÓN:**
6. ✅ No crear ObtenerUltimasNInspecciones() nueva (ya existe)
7. ✅ No modificar CalcularRPN() (seguirá siendo simple)

---

### 0.3 Backup inicial
```bash
git checkout -b feature/recurrent-inspections
git add .
git commit -m "CHECKPOINT #0: Estado base antes de inspecciones recurrentes"
git push origin feature/recurrent-inspections
```

### 0.3 Validación de requisitos
**PREGUNTAS CRÍTICAS:**

1. **Sobre el puesto de trabajo:**
   - Si JDG tiene puesto "ACF" y "NPT", ¿el histórico debe filtrar por puesto específico?
   - ¿O todas las inspecciones de JDG sin importar puesto?
   - **Ejemplo**: JDG inspeccionado como ACF hace 3 meses, hoy se inspecciona como NPT → ¿Se usa el RPN de ACF o se considera "primera" en NPT?

2. **Sobre la búsqueda de RPN anterior:**
   - ¿Buscar última inspección por fecha de ejecución?
   - ¿O última inspección APROBADA solamente?
   - ¿Qué pasa si la última fue Cat 5 (No Calificado)?

3. **Sobre datos microbiología (futuro):**
   - ¿Los % Recuperación y % OOL serán campos calculados o ingreso manual?
   - ¿Van en tblInspecciones o nueva tabla tblDatosMicrobiologia?

---

## 📦 Fase 1: Extensión de Base de Datos (CHECKPOINT #1)

### 1.0 Análisis de Estructura Actual (OBLIGATORIO - Antes de cualquier cambio)

**Objetivo:** Mapear estructura real de tablas según Configuration2.bas

#### Tabla 1: tblInspecciones (Hoja "Historico")
**Columnas actuales: 31 (REAL - Configuration2.bas desactualizado con 30)**

| # | Nombre Columna | Tipo | Uso Actual | Módulos Dependientes |
|---|----------------|------|------------|---------------------|
| 01 | ID Inspeccion | String (UUID) | PK única | InspectionRepository, InspectionScheduler |
| 02 | Area | String | Área de inspección | ChecklistOrchestrator |
| 03 | Linea Auditada | String | Línea auditada | ChecklistOrchestrator |
| 04 | Hora inicio | String | Hora inicio | ChecklistOrchestrator |
| 05 | Hora termino | String | Hora término | ChecklistOrchestrator |
| 06 | Iniciales AY1 | String | Ayudante 1 | ChecklistOrchestrator |
| 07 | Iniciales AY2 | String | Ayudante 2 | ChecklistOrchestrator |
| 08 | Iniciales OP | String | Operador | ChecklistOrchestrator |
| 09 | Lugar Auditoria | String | Lugar auditoría | ChecklistOrchestrator |
| 10 | Iniciales personal | String | FK tblPersonal | **CRÍTICO: Filtro búsqueda** |
| 11 | ID Plantilla | String | FK tblPlantillas | InspectionScheduler, ChecklistOrchestrator |
| 12 | Planta | String | Planta ejecutora | ChecklistOrchestrator |
| 13 | Fecha inspeccion | Date | Fecha realización | **CRÍTICO: Ordenamiento** |
| 14 | Auditor | String | Auditor | ChecklistOrchestrator |
| 15 | Estado | String | Estado inspección | ChecklistOrchestrator |
| 16 | TA puntaje obtenido | Double | Puntaje TA | InspectionCalculator |
| 17 | TA puntos maximos | Double | Máximos TA | InspectionCalculator |
| 18 | TA puntos no aplica | Double | No aplica TA | InspectionCalculator |
| 19 | TA porcentaje | Double | % TA | InspectionCalculator |
| 20 | Auditoria Procesos Resultado | String | **NUEVA - no en Configuration2** | ChecklistOrchestrator, CertificadoPDFGenerator |
| 21 | RPN calculado | Double | **RPN actual** | **CRÍTICO: Cálculo categoría** |
| 22 | Categoria resultado | String | Categoría | **CRÍTICO: Certificado PDF** |
| 23 | Requiere accion | String | Acción requerida | ChecklistOrchestrator |
| 24 | Fecha proxima inspeccion | Date | Próxima inspección | InspectionScheduler |
| 25 | Dias para vencimiento | Long | Días vencimiento | InspectionScheduler |
| 26 | Estado programacion | String | Estado programación | InspectionScheduler |
| 27 | Observaciones generales | String | Observaciones | ChecklistOrchestrator |
| 28 | Fecha calculo | Date | Timestamp cálculos | InspectionRepository |
| 29 | Usuario calculo | String | Usuario cálculos | InspectionRepository |
| 30 | Fecha completado | Date | Timestamp completado | InspectionRepository |
| 31 | Usuario completado | String | Usuario completado | InspectionRepository |

**⚠️ NOTA CRÍTICA:** Columnas 10 (Iniciales personal), 13 (Fecha inspeccion), 21 (RPN calculado), 22 (Categoria resultado) son fundamentales para búsqueda histórica.

---

### 1.1 Diseño de nuevas columnas en tblInspecciones

**NUEVAS COLUMNAS A AGREGAR (9 columnas - Numeración ACTUALIZADA):**

| # | Nombre Columna | Tipo | Default | Null? | Descripción | Flujo |
|---|----------------|------|---------|-------|-------------|-------|
| **32** | **Numero Inspeccion** | Long | 1 | NO | Secuencia: 1, 2, 3, 4... | Ambos |
| **33** | **Es Inspeccion Recurrente** | String | "No" | NO | "Si"/"No" flag | Ambos |
| **34** | **Puesto Evaluado** | String | "" | NO | Puesto en esta inspección | **CRÍTICO: Filtro** |
| **35** | **RPN Anterior Manual** | Double | Null | SÍ | RPN ingresado manualmente | Recurrente |
| **36** | **ID Inspeccion Anterior** | String | Null | SÍ | UUID si vino del sistema | Recurrente |
| **37** | **RPN Promedio** | Double | Null | SÍ | (RPN Ant + RPN Act)/2 | Recurrente |
| **38** | **Porcentaje Recuperacion** | Double | 0 | SÍ | Futuro: microbiología | Futuro |
| **39** | **Porcentaje OOL** | Double | 0 | SÍ | Futuro: microbiología | Futuro |
| **40** | **RPN Total** | Double | Null | SÍ | RPN Prom + %Rec + %OOL | Recurrente |

**COLUMNA MODIFICADA:**

| # Actual | Nombre | Cambio Conceptual | Compatibilidad |
|----------|--------|-------------------|----------------|
| **22** | Categoria resultado | Ahora basada en RPN o RPN Total | COMPATIBLE (mismo algoritmo) |

**TOTAL COLUMNAS DESPUÉS DE MIGRACIÓN: 40 (31 actuales + 9 nuevas)**

---

### 1.2 Justificación de Columna 34: "Puesto Evaluado"

**PROBLEMA ACTUAL:**
- tblInspecciones NO tiene columna "Puesto" explícita
- Puesto se infiere de tblPersonal mediante iniciales
- Si persona tiene múltiples puestos → ambigüedad
- Búsqueda histórica requiere filtro por PUESTO + INICIALES (Q1)

**SOLUCIÓN:**
- Agregar columna "Puesto Evaluado" que registra el puesto ESPECÍFICO de esta inspección
- Se llena desde frmChecklistVirtual al seleccionar plantilla
- Permite búsqueda histórica precisa: `WHERE Iniciales='JDG' AND PuestoEvaluado='Operador'`

**EJEMPLO:**
```
Personal: JDG tiene puestos "Operador" + "Mezclador"
Inspección 1: JDG como Operador → Col[34]="Operador"
Inspección 2: JDG como Mezclador → Col[34]="Mezclador"
Búsqueda histórica Operador → Solo encuentra Inspección 1
```

---

### 1.3 Estrategia de Migración Detallada

**Script:** `TableManager.AgregarColumnasInspeccionesRecurrentes()`

```vb
Public Sub AgregarColumnasInspeccionesRecurrentes()
    ' FASE 1: Pre-validación
    ' - Verificar que tblInspecciones existe
    ' - Contar registros actuales
    ' - Hacer backup interno (copiar datos a array)
    
    ' FASE 2: Agregar columnas una por una
    ' Columna 32: Numero Inspeccion
    tbl.ListColumns.Add Position:=32
    tbl.ListColumns(32).Name = "Numero Inspeccion"
    tbl.ListColumns(32).DataBodyRange.Value = 1  ' Default: todas son "primera"
    
    ' Columna 33: Es Inspeccion Recurrente
    tbl.ListColumns.Add Position:=33
    tbl.ListColumns(33).Name = "Es Inspeccion Recurrente"
    tbl.ListColumns(33).DataBodyRange.Value = "No"
    
    ' Columna 34: Puesto Evaluado (CRÍTICA)
    tbl.ListColumns.Add Position:=34
    tbl.ListColumns(34).Name = "Puesto Evaluado"
    ' Intentar inferir desde tblPlantillas (ID Plantilla → Puesto)
    Call InferirPuestoHistorico(tbl)
    
    ' Columnas 35-40: Campos recurrentes (NULL por defecto)
    ' ... agregar una por una con DataBodyRange.Value = Empty
    
    ' FASE 3: Validación post-migración
    ' - Verificar count de registros igual
    ' - Verificar que columna 34 tiene valores válidos
    ' - Log detallado
    
    ' FASE 4: Actualizar Configuration2.bas (manual)
    ' - Documentar nuevas columnas
    ' - Actualizar comentarios de módulos dependientes
End Sub

Private Sub InferirPuestoHistorico(tbl As ListObject)
    ' Para cada registro histórico:
    ' 1. Leer ID Plantilla (col 11)
    ' 2. Buscar en tblPlantillas → obtener Puesto
    ' 3. Escribir en col 34
    ' 4. Si no encuentra → escribir "DESCONOCIDO" (requiere corrección manual)
End Sub
```

**⚠️ VALIDACIÓN CRÍTICA:**
```vb
' Después de migración, validar:
Dim filasSinPuesto As Long
filasSinPuesto = Application.WorksheetFunction.CountIf( _
    tbl.ListColumns("Puesto Evaluado").DataBodyRange, "DESCONOCIDO")

If filasSinPuesto > 0 Then
    MsgBox "ADVERTENCIA: " & filasSinPuesto & " registros históricos " & _
           "sin puesto asignado. Revisar manualmente."
End If
```

---

### 1.4 Tabla de Referencias Cruzadas por Módulo

**Módulos que necesitarán actualización:**

| Módulo | Lee tblInspecc | Escribe | Columnas Usadas | Cambios Requeridos |
|--------|----------------|---------|-----------------|-------------------|
| InspectionRepository.bas | ✅ | ✅ | Todas | ✅ Agregar campos recurrentes (32-40) |
| InspectionScheduler.bas | ✅ | ❌ | 10,13,21,22 | ✅ Leer col 32,34,40 |
| ChecklistOrchestrator.bas | ❌ | ✅ (vía Repo) | Via Repository | ✅ Pasar datos recurrentes |
| InspectionCalculator.bas | ❌ | ❌ | Recibe datos | ⚠️ Mod menor (modo cálculo) |
| CertificadoPDFGenerator.bas | ✅ | ❌ | 1-31 | ✅ Leer 32-40 para PDF recurrente |
| frmChecklistVirtual.frm | ❌ | ✅ (vía Orch) | Via Orchestrator | ✅ Capturar Puesto Evaluado |

---

### 1.5 Validación Post-Migración (Checklist)

- [ ] **Integridad de datos:**
  - [ ] Todas las inspecciones tienen `Numero Inspeccion = 1`
  - [ ] Todas las inspecciones tienen `Es Inspeccion Recurrente = "No"`
  - [ ] Ninguna inspección tiene valores NULL en columnas obligatorias
  - [ ] Columna "Puesto Evaluado" tiene <5% de "DESCONOCIDO"

- [ ] **Compatibilidad hacia atrás:**
  - [ ] InspectionCalculator.CalcularRPN() sigue funcionando sin cambios
  - [ ] Certificados PDF de inspecciones antiguas se generan correctamente
  - [ ] Cronograma no se rompe con nuevas columnas

- [ ] **Performance:**
  - [ ] Tiempo de carga de tblInspecciones < 2 segundos (con 1000+ registros)
  - [ ] Búsquedas históricas en InspectionHistoryService < 1 segundo

**CHECKPOINT #1 - Git Commit:**
```bash
git add TableManager.bas Configuration2.bas
git commit -m "CHECKPOINT #1: Migración BD - Columnas inspecciones recurrentes"
git push
```

---

### 1.6 Análisis de Impacto: tblCronogramaInspecciones

**Ubicación:** Hoja "Cronograma"  
**Columnas actuales:** 24 (según Configuration2.bas línea 270-308)

**PREGUNTA:** ¿Necesita modificaciones?

**ANÁLISIS:**

| Columna Relevante | Uso Actual | ¿Afectada? | Justificación |
|-------------------|------------|-----------|---------------|
| [10] RPN ultima inspeccion | Almacena último RPN | ✅ SÍ | Podría ser RPN o RPN Total |
| [11] Categoria ultima inspeccion | Categoría actual | ✅ SÍ | Podría basarse en RPN Total |
| [06] Total inspecciones | Contador | ⚠️ REVISAR | ¿Cuenta todas o solo primeras? |

**DECISIÓN RECOMENDADA: NO MODIFICAR POR AHORA**

**Razones:**
1. **RPN ultima inspeccion** ya almacena el "RPN final" de la última inspección
   - Si es 1ra inspección → almacena RPN (TA)
   - Si es 2da+ inspección → almacena RPN Total
   - La columna es genérica y sirve para ambos casos

2. **Categoria ultima inspeccion** ya almacena la categoría final
   - Basada en RPN o RPN Total según corresponda
   - No necesita distinguir el origen

3. **Total inspecciones** cuenta TODAS las inspecciones del personal+puesto
   - Útil para estadísticas
   - No requiere cambios

**⚠️ POSIBLE MEJORA FUTURA (Fase 2.0):**
Si se requiere análisis detallado, agregar columnas opcionales:
- `Numero Ultima Inspeccion` (Long) - Para saber si última fue 1ra, 2da, 3ra...
- `Tipo RPN Ultima` (String) - "TA" o "RPN_TOTAL" para reportes

**Por ahora:** Mantener tblCronogramaInspecciones SIN cambios

---

### 1.7 Análisis de Impacto: Otras Tablas

**Tablas que NO requieren modificaciones:**

| Tabla | Razón |
|-------|-------|
| tblRespuestas | Solo almacena respuestas individuales, independiente del número de inspección |
| tblPlantillas | Ya tiene "Frecuencia meses", suficiente para cálculo de validez |
| tblPersonal | No necesita saber cuántas inspecciones tiene cada persona (eso está en Histórico) |
| tblPreguntas | Definición de preguntas es independiente del número de inspección |
| tblOpcionesDeRespuesta | Opciones son fijas independiente del número de inspección |
| tblCategoriasRPN | Rangos de categorización sirven tanto para RPN como RPN Total |

---

## 🎨 Fase 2: Interfaz de Usuario - frmChecklistVirtual (CHECKPOINT #2)

### 2.1 Diseño de controles nuevos

**Ubicación:** Después de selector de personal, antes de iniciar checklist

**Controles a agregar:**
```vb
' Frame contenedor
Frame: frmRecurrentInspection
  Caption: "Inspección Recurrente"
  Visible: True
  
  ' CheckBox principal
  CheckBox: chkEsRecurrente
    Caption: "☐ Esta NO es la primera inspección de este personal"
    Default: False
    
  ' Botón de búsqueda histórico
  CommandButton: btnBuscarHistorico
    Caption: "🔍 Buscar inspecciones previas"
    Enabled: True (siempre visible)
    
  ' Frame de datos históricos (solo visible si chkEsRecurrente = True)
  Frame: frmDatosHistorico
    Visible: False (activado por chkEsRecurrente)
    
    ' Label info
    Label: lblInfoHistorico
      Caption: "[Se mostrará info de inspecciones previas aquí]"
      ForeColor: Blue
    
    ' Número de inspección
    Label: lblNumeroInspeccion
      Caption: "Esta es la inspección número:"
    TextBox: txtNumeroInspeccion
      Locked: True (calculado automáticamente)
      BackColor: LightGray
    
    ' RPN anterior - modo automático
    Label: lblRPNAnteriorAuto
      Caption: "RPN anterior (del sistema):"
      Visible: False
    TextBox: txtRPNAnteriorAuto
      Locked: True
      Visible: False
    
    ' RPN anterior - modo manual
    Label: lblRPNAnteriorManual
      Caption: "RPN anterior (ingreso manual):"
      Visible: False
    TextBox: txtRPNAnteriorManual
      Locked: False
      Visible: False
      
    ' Indicador de modo
    Label: lblModoRPN
      Caption: "[Modo: Automático/Manual]"
      ForeColor: Green/Orange
```

### 2.2 Eventos y lógica de UI

**Event: chkEsRecurrente_Click()**
```vb
If chkEsRecurrente.Value = True Then
    frmDatosHistorico.Visible = True
    ' Limpiar campos
    txtNumeroInspeccion = ""
    txtRPNAnteriorAuto = ""
    txtRPNAnteriorManual = ""
Else
    frmDatosHistorico.Visible = False
    ' Resetear valores
End If
```

**Event: btnBuscarHistorico_Click()**
```vb
' 1. Validar que haya personal seleccionado
' 2. Llamar a InspectionHistoryService.BuscarInspeccionesPrevias()
' 3. Mostrar resultado en lblInfoHistorico
' 4. Si encuentra inspecciones:
'    - Autocompletar txtNumeroInspeccion = última + 1
'    - Mostrar txtRPNAnteriorAuto con último RPN
'    - Activar chkEsRecurrente automáticamente
'    - lblModoRPN = "Modo: AUTOMÁTICO"
' 5. Si NO encuentra:
'    - Mostrar mensaje "No hay inspecciones previas en el sistema"
'    - Habilitar txtRPNAnteriorManual
'    - lblModoRPN = "Modo: MANUAL"
```

### 2.3 Validaciones

**Validar antes de iniciar checklist:**
```vb
If chkEsRecurrente = True Then
    If txtNumeroInspeccion = "" Or Val(txtNumeroInspeccion) < 2 Then
        MsgBox "Debe indicar número de inspección >= 2"
        Exit Sub
    End If
    
    If txtRPNAnteriorAuto = "" And txtRPNAnteriorManual = "" Then
        MsgBox "Debe tener RPN anterior (automático o manual)"
        Exit Sub
    End If
    
    If txtRPNAnteriorManual <> "" Then
        If Not IsNumeric(txtRPNAnteriorManual) Or Val(txtRPNAnteriorManual) <= 0 Then
            MsgBox "RPN anterior manual debe ser numérico > 0"
            Exit Sub
        End If
    End If
End If
```

**CHECKPOINT #2 - Git Commit:**
```bash
git add frmChecklistVirtual.frm frmChecklistVirtual.frx
git commit -m "CHECKPOINT #2: UI - Controles inspecciones recurrentes"
git push
```

---

## 🔧 Fase 3: Servicio de Histórico (CHECKPOINT #3)

### 3.1 Crear InspectionHistoryService.bas

**Responsabilidad:** Consultas de inspecciones previas

```vb
' ══════════════════════════════════════════════════════════════
' Módulo: InspectionHistoryService
' Descripción: Gestión de histórico de inspecciones del personal
' Fecha creación: 21/04/2026
' Autor: Sistema TH-HC-001
' ══════════════════════════════════════════════════════════════
Option Explicit

' ══════════════════════════════════════════════════════════════
' Buscar todas las inspecciones previas de un personal
' ══════════════════════════════════════════════════════════════
Public Function BuscarInspeccionesPrevias( _
    ByVal iniciales As String, _
    Optional ByVal filtroPorPuesto As Boolean = True, _
    Optional ByVal puesto As String = "" _
) As Collection
    ' Retorna Collection de Dictionary con:
    ' - ID Inspeccion
    ' - Fecha Inspeccion
    ' - Puesto
    ' - RPN
    ' - Categoria
    ' - Numero Inspeccion
    ' Ordenados por Fecha DESC (más reciente primero)
End Function

' ══════════════════════════════════════════════════════════════
' Obtener última inspección de un personal
' ══════════════════════════════════════════════════════════════
Public Function ObtenerUltimaInspeccion( _
    ByVal iniciales As String, _
    Optional ByVal filtroPorPuesto As Boolean = True, _
    Optional ByVal puesto As String = "" _
) As Object
    ' Retorna Dictionary con datos de última inspección
    ' Nothing si no hay inspecciones previas
End Function

' ══════════════════════════════════════════════════════════════
' Calcular número de inspección siguiente
' ══════════════════════════════════════════════════════════════
Public Function CalcularNumeroInspeccionSiguiente( _
    ByVal iniciales As String, _
    Optional ByVal filtroPorPuesto As Boolean = True, _
    Optional ByVal puesto As String = "" _
) As Long
    ' Si no hay inspecciones previas → 1
    ' Si hay inspecciones → MAX(Numero Inspeccion) + 1
End Function

' ══════════════════════════════════════════════════════════════
' Validar coherencia de RPN anterior
' ══════════════════════════════════════════════════════════════
Public Function ValidarRPNAnteriorManual( _
    ByVal rpnManual As Double, _
    ByVal iniciales As String _
) As Boolean
    ' Validaciones de negocio:
    ' 1. RPN > 0
    ' 2. RPN dentro de rangos lógicos (ej: entre 0.1 y 100)
    ' 3. Opcional: Advertir si difiere mucho del histórico automático
End Function
```

### 3.2 Implementación de queries

**Estrategia:**
- Usar InspectionRepository donde sea posible
- Queries ordenadas y optimizadas
- Manejo robusto de errores
- Logging detallado

**CHECKPOINT #3 - Git Commit:**
```bash
git add InspectionHistoryService.bas
git commit -m "CHECKPOINT #3: Servicio histórico de inspecciones"
git push
```

---

## 🧮 Fase 4: Calculadora de RPN Recurrente (CHECKPOINT #4)

### 4.1 Crear RecurrentInspectionCalculator.bas

**Responsabilidad:** Cálculos de RPN para inspecciones 2da+

```vb
' ══════════════════════════════════════════════════════════════
' Módulo: RecurrentInspectionCalculator
' Descripción: Cálculos de RPN para inspecciones recurrentes
' Fecha creación: 21/04/2026
' ══════════════════════════════════════════════════════════════
Option Explicit

' ══════════════════════════════════════════════════════════════
' Calcular RPN Promedio (primera fase cálculo recurrente)
' ══════════════════════════════════════════════════════════════
Public Function CalcularRPNPromedio( _
    ByVal rpnAnterior As Double, _
    ByVal rpnActual As Double _
) As Double
    ' Fórmula: (RPN Anterior + RPN Actual) / 2
    ' Validaciones:
    ' - Ambos > 0
    ' - Resultado lógico
    
    If rpnAnterior <= 0 Or rpnActual <= 0 Then
        Err.Raise vbObjectError + 1000, "RecurrentInspectionCalculator", _
            "RPN anterior y actual deben ser > 0"
    End If
    
    CalcularRPNPromedio = (rpnAnterior + rpnActual) / 2
    
    Debug.Print "[RPN RECURRENTE] RPN Promedio = (" & rpnAnterior & " + " & _
                rpnActual & ") / 2 = " & CalcularRPNPromedio
End Function

' ══════════════════════════════════════════════════════════════
' Calcular RPN Total (fase completa - preparado para futuro)
' ══════════════════════════════════════════════════════════════
Public Function CalcularRPNTotal( _
    ByVal rpnPromedio As Double, _
    Optional ByVal porcRecuperacion As Double = 0, _
    Optional ByVal porcOOL As Double = 0 _
) As Double
    ' Fórmula: RPN Promedio + % Recuperación + % OOL
    ' Actualmente solo usa RPN Promedio (porcRecup y OOL = 0)
    ' Preparado para fase futura de microbiología
    
    CalcularRPNTotal = rpnPromedio + porcRecuperacion + porcOOL
    
    Debug.Print "[RPN RECURRENTE] RPN Total = " & rpnPromedio & " + " & _
                porcRecuperacion & " + " & porcOOL & " = " & CalcularRPNTotal
End Function

' ══════════════════════════════════════════════════════════════
' Determinar categoría basada en RPN Total
' ══════════════════════════════════════════════════════════════
Public Function DeterminarCategoriaRPNTotal( _
    ByVal rpnTotal As Double _
) As Long
    ' Usa misma tabla tblCategoriasRPN pero con RPN Total
    ' Retorna número de categoría (1-5)
    ' Reutiliza lógica de InspectionCalculator.DeterminarCategoria()
End Function
```

### 4.2 Integración con InspectionCalculator existente

**Modificación mínima:**
```vb
' En InspectionCalculator.bas
' Agregar parámetro opcional para diferenciar modo

Public Function DeterminarCategoria( _
    ByVal rpnValue As Double, _
    Optional ByVal modoCalculo As String = "TA" _
) As Long
    ' modoCalculo = "TA" → lógica actual (default)
    ' modoCalculo = "RPN_TOTAL" → mismo algoritmo, nombre diferente
    
    ' El algoritmo de categorización es el MISMO
    ' Solo cambia el origen del valor RPN
End Function
```

**CHECKPOINT #4 - Git Commit:**
```bash
git add RecurrentInspectionCalculator.bas InspectionCalculator.bas
git commit -m "CHECKPOINT #4: Calculadora RPN recurrente + integración"
git push
```

---

## 🔄 Fase 5: Pipeline de Orquestación (CHECKPOINT #5)

### 5.1 Modificar ChecklistOrchestrator.bas

**Método principal a modificar:** `FinalizarInspeccion()`

**Nuevo flujo de decisión:**
```vb
' ══════════════════════════════════════════════════════════════
' Pipeline de cálculo RPN (con soporte recurrente)
' ══════════════════════════════════════════════════════════════
Private Function CalcularMetricasInspeccion( _
    datosInspeccion As Object, _
    esRecurrente As Boolean, _
    numeroInspeccion As Long, _
    rpnAnterior As Double _
) As Object
    
    Dim metricas As Object
    Set metricas = CreateObject("Scripting.Dictionary")
    
    ' PASO 1: Calcular TA (SIEMPRE, es la base)
    Dim rpnTA As Double
    rpnTA = InspectionCalculator.CalcularRPN(datosInspeccion)
    metricas("RPN_TA") = rpnTA
    
    ' PASO 2: Decidir flujo según tipo inspección
    If Not esRecurrente Or numeroInspeccion = 1 Then
        ' ═══ FLUJO ACTUAL (1ra inspección) ═══
        metricas("Numero Inspeccion") = 1
        metricas("Es Inspeccion Recurrente") = False
        metricas("RPN Final") = rpnTA
        metricas("Categoria Aplicada") = "TA"
        metricas("Categoria") = InspectionCalculator.DeterminarCategoria(rpnTA, "TA")
        
    Else
        ' ═══ FLUJO NUEVO (2da+ inspección) ═══
        metricas("Numero Inspeccion") = numeroInspeccion
        metricas("Es Inspeccion Recurrente") = True
        metricas("RPN Anterior") = rpnAnterior
        
        ' Calcular RPN Promedio
        Dim rpnPromedio As Double
        rpnPromedio = RecurrentInspectionCalculator.CalcularRPNPromedio(rpnAnterior, rpnTA)
        metricas("RPN Promedio") = rpnPromedio
        
        ' Calcular RPN Total (por ahora = RPN Promedio, futuro +micro)
        Dim rpnTotal As Double
        rpnTotal = RecurrentInspectionCalculator.CalcularRPNTotal(rpnPromedio, 0, 0)
        metricas("RPN Total") = rpnTotal
        metricas("RPN Final") = rpnTotal
        
        ' Categoría basada en RPN Total
        metricas("Categoria Aplicada") = "RPN_TOTAL"
        metricas("Categoria") = RecurrentInspectionCalculator.DeterminarCategoriaRPNTotal(rpnTotal)
    End If
    
    Set CalcularMetricasInspeccion = metricas
End Function
```

### 5.2 Modificar guardado en tblInspecciones

**InspectionRepository.GuardarInspeccion()** - Agregar campos nuevos:
```vb
' Campos para todas las inspecciones
nuevaFila.Cells(1, colNumeroInspeccion).Value = metricas("Numero Inspeccion")
nuevaFila.Cells(1, colEsRecurrente).Value = metricas("Es Inspeccion Recurrente")
nuevaFila.Cells(1, colCategoriaAplicada).Value = metricas("Categoria Aplicada")

' Campos solo para inspecciones recurrentes
If metricas("Es Inspeccion Recurrente") = True Then
    nuevaFila.Cells(1, colRPNAnterior).Value = metricas("RPN Anterior")
    nuevaFila.Cells(1, colRPNPromedio).Value = metricas("RPN Promedio")
    nuevaFila.Cells(1, colRPNTotal).Value = metricas("RPN Total")
    
    ' ID inspeccion anterior (si vino del sistema)
    If metricas.Exists("ID Inspeccion Anterior") Then
        nuevaFila.Cells(1, colIDInspeccionAnterior).Value = metricas("ID Inspeccion Anterior")
    End If
End If
```

**CHECKPOINT #5 - Git Commit:**
```bash
git add ChecklistOrchestrator.bas InspectionRepository.bas
git commit -m "CHECKPOINT #5: Pipeline orquestación inspecciones recurrentes"
git push
```

---

## ✅ Fase 6: Testing y Validación (CHECKPOINT #6)

### 6.1 Plan de pruebas

**Test Case 1: Primera inspección (flujo actual sin cambios)**
- [ ] Personal nuevo sin historial
- [ ] Checkbox recurrente DESMARCADO
- [ ] Resultado = categoría basada en TA
- [ ] Campos recurrentes = NULL en BD

**Test Case 2: Segunda inspección con histórico en sistema**
- [ ] Personal con 1 inspección previa
- [ ] Click "Buscar inspecciones previas" → encuentra 1
- [ ] Checkbox marcado automáticamente
- [ ] RPN anterior autocargado
- [ ] Número inspección = 2
- [ ] Resultado = categoría basada en RPN Total

**Test Case 3: Segunda inspección SIN histórico (ingreso manual)**
- [ ] Personal nuevo (datos históricos externos)
- [ ] Marcar checkbox manualmente
- [ ] Ingresar RPN anterior manual
- [ ] Número inspección = 2
- [ ] Validar cálculo correcto

**Test Case 4: Validaciones**
- [ ] No permitir inspección recurrente sin RPN anterior
- [ ] Validar RPN anterior numérico > 0
- [ ] Advertir si número inspección < 2 con checkbox marcado

### 6.2 Datos de prueba

Crear registros de prueba:
```vb
' Personal: TEST_JDG
' Inspección 1: Cat 2, RPN = 15.50
' Inspección 2: Cat 1, RPN = 12.00
' Expected RPN Promedio = (15.50 + 12.00)/2 = 13.75
' Expected Categoría = f(13.75) según tblCategoriasRPN
```

**CHECKPOINT #6 - Git Commit:**
```bash
git add TestDataGenerator.bas
git commit -m "CHECKPOINT #6: Tests validación inspecciones recurrentes"
git push
```

---

## 📚 Fase 7: Documentación (CHECKPOINT #7)

### 7.1 Crear documentación técnica

**Archivo:** `docs/SISTEMA_INSPECCIONES_RECURRENTES.md`

Contenido:
- Diagrama de flujo decisión RPN
- Diccionario de datos (nuevas columnas)
- Reglas de negocio
- Fórmulas matemáticas
- Casos de uso
- Troubleshooting

### 7.2 Actualizar README.md

Agregar sección "Inspecciones Recurrentes"

**CHECKPOINT #7 - Git Commit:**
```bash
git add docs/SISTEMA_INSPECCIONES_RECURRENTES.md README.md
git commit -m "CHECKPOINT #7: Documentación inspecciones recurrentes"
git push
```

---

## 🚀 Fase 8: Merge y Despliegue

### 8.1 Code Review
- [ ] Revisar cada archivo modificado
- [ ] Verificar que flujo actual (1ra inspección) sigue funcionando
- [ ] Validar que no hay código duplicado (DRY)
- [ ] Confirmar desacoplamiento de módulos

### 8.2 Merge a main
```bash
git checkout main
git merge feature/recurrent-inspections
git tag v2.0.0-recurrent-inspections
git push origin main --tags
```

---

## ❓ PREGUNTAS CRÍTICAS A RESOLVER ANTES DE EMPEZAR

**Por favor confirma o aclara:**

### Q1: Filtro por puesto
Cuando busco inspecciones previas de "JDG":
- **Opción A**: Solo inspecciones del MISMO puesto (JDG-ACF busca solo ACF)
- **Opción B**: Todas las inspecciones de JDG sin importar puesto
- **Opción C**: Preguntar al usuario en el formulario (checkbox "Filtrar por puesto")

**DECISIÓN:** ✅ **Opción A - Filtro por PUESTO + INICIALES**
- Búsqueda histórica debe filtrar por AMBOS criterios
- Ejemplo: JDG-Operador solo compara con otras inspecciones de JDG-Operador
- JDG-Mezclador se considera inspección independiente (primera)
- Lógica: `WHERE Iniciales = 'JDG' AND Puesto = 'Operador'`

### Q2: Categorización mixta
Si RPN Total cae en Cat 3, pero el usuario en inspección actual tuvo respuestas "No Cumple" (que lo llevarían a Cat 5 por TA puro):
- **¿Prevalece RPN Total siempre?**
- **¿O hay regla de "worst case" (tomar el peor)?**

**DECISIÓN:** ✅ **RPN Total prevalece para categorización**
- 1ra inspección → Categoría determinada por RPN (TA)
- 2da+ inspección → Categoría determinada por RPN Total
- **IMPORTANTE**: "Auditoría de Procesos" es un sistema SEPARADO
  - Arroja solo Cumple/No Cumple
  - NO se compara con inspecciones anteriores
  - NO afecta categorización de Técnica Aséptica
- No hay conflicto: cada sistema tiene su propósito independiente

### Q3: Microbiología futura
Los % Recuperación y % OOL:
- **¿Son valores 0-100?** (ej: 85% recuperación)
- **¿O son factores de penalización?** (ej: +5 puntos por OOL)
- **¿Necesitas que deje estructura de tabla tblDatosMicrobiologia lista?**

**DECISIÓN:** ⏸️ **PENDIENTE - Fase futura**
- Por ahora: Columnas en tblInspecciones con valor NULL/0
- Se definirá en fase posterior cuando se implemente microbiología
- Estructura preparada pero no implementada

### Q4: Certificado PDF
¿El certificado debe mostrar RPN Promedio y RPN Total cuando es inspección recurrente?
- **Opción A**: Mostrar solo resultado final (como ahora)
- **Opción B**: Mostrar desglose: "RPN TA: X | RPN Promedio: Y | RPN Total: Z"

**DECISIÓN:** ✅ **Opción B - Mostrar desglose diferenciado**

**1ra Inspección (flujo actual):**
```
RPN: 15.50
Categoría: 2 - ESTABLE
```

**2da+ Inspección (flujo nuevo):**
```
RPN Anterior: 15.50
RPN Actual: 12.00
RPN Total: 13.75
Categoría: 1 - ÓPTIMO (basada en RPN Total)
```

**Cambios requeridos:**
- Modificar PlantillaCertificadoSetup.bas: Agregar filas condicionales
- Modificar CertificadoPDFGenerator.bas: Lógica de población condicional
- Layout adaptativo según tipo de inspección

### Q5: Auditoría
¿Necesitas auditar cambios de RPN anterior manual?
- Log cuando usuario edita RPN anterior manualmente

**DECISIÓN:** ✅ **SÍ - Auditoría habilitada**
- Registrar en AuditLogger cuando se ingresa RPN manual
- Campos a auditar:
  - Usuario que ingresó RPN manual
  - Valor ingresado
  - Fecha/hora de ingreso
  - ID de inspección relacionada
- Formato: `[RPN MANUAL] Usuario: JDG | RPN ingresado: 15.50 | Inspección: uuid-123`

---

## ⏱️ Estimación de Tiempo por Fase

| Fase | Descripción | Tiempo estimado |
|------|-------------|-----------------|
| 0 | Análisis y preparación | 1-2 horas |
| 1 | Migración BD | 2-3 horas |
| 2 | UI frmChecklistVirtual | 3-4 horas |
| 3 | InspectionHistoryService | 2-3 horas |
| 4 | RecurrentInspectionCalculator | 2 horas |
| 5 | Pipeline orquestación | 3-4 horas |
| 6 | Testing y validación | 3-4 horas |
| 7 | Documentación | 2 horas |
| **TOTAL** | **18-25 horas** |

---

## 📝 Notas de Implementación

### Fórmulas Matemáticas

#### Primera Inspección (Flujo Actual)
```
RPN = f(TA Puntaje, TA Máximos, TA No Aplica)
Categoría = f(RPN) según tblCategoriasRPN
```

#### Inspecciones Recurrentes (Flujo Nuevo)
```
RPN_TA = f(TA Puntaje, TA Máximos, TA No Aplica)
RPN_Promedio = (RPN_Anterior + RPN_TA) / 2

// Fase actual (solo TA)
RPN_Total = RPN_Promedio + 0 + 0

// Fase futura (con microbiología)
RPN_Total = RPN_Promedio + %Recuperación + %OOL

Categoría = f(RPN_Total) según tblCategoriasRPN
```

### Compatibilidad Hacia Atrás

**Garantías:**
1. Todas las inspecciones existentes se marcan como "Primera inspección"
2. Flujo actual funciona SIN cambios si checkbox NO se marca
3. Categorización actual permanece idéntica para inspecciones tradicionales
4. Nuevas columnas permiten NULL para registros legacy

### Escalabilidad Futura

**Preparación para microbiología:**
- Columnas % Recuperación y % OOL ya en esquema
- Función CalcularRPNTotal() preparada con parámetros opcionales
- Fácil agregar tabla tblDatosMicrobiologia posteriormente
- Pipeline de cálculo modular para agregar nuevos factores

### Impacto en Certificado PDF (MVP recién completado)

**⚠️ IMPORTANTE:** El certificado MVP que acabamos de finalizar necesitará ajustes:

**Modificaciones necesarias en sección Resultados:**
```vb
' Actual (1ra inspección):
Fila 23: RPN: [XX.XX]
Fila 24: Categoría resultado: [N - Descripción]

' Nuevo (2da+ inspección):
Fila 23: RPN Anterior: [XX.XX]
Fila 24: RPN Actual: [XX.XX]
Fila 25: RPN Total: [XX.XX]
Fila 26: Categoría resultado: [N - Descripción]
```

**Estrategia:**
1. Layout dinámico: PlantillaCertificadoSetup detecta tipo inspección
2. Poblado condicional: CertificadoPDFGenerator.PoblarPlantillaCertificado()
3. Filas 23-26 adaptativas según `Es Inspeccion Recurrente`
4. Mantener compatibilidad con certificados de 1ra inspección

---

## 🎯 Estado del Plan

- [ ] **Fase 0**: ✅ LISTO PARA INICIAR - Todas las decisiones tomadas
- [x] **Preguntas Q1-Q5**: ✅ RESPONDIDAS
  - Q1: Filtrar por Puesto + Iniciales (ambos)
  - Q2: RPN Total prevalece en inspecciones recurrentes
  - Q3: Microbiología pendiente para fase futura
  - Q4: Certificado PDF con desglose completo
  - Q5: Auditoría habilitada para RPN manual
- [ ] **Revisión arquitectura**: Pendiente aprobación usuario
- [ ] **Inicio implementación**: Esperando autorización

---

## 🟢 SIGUIENTE PASO

**¿Procedo con la Fase 0?**
- Análisis de impacto en código actual
- Backup inicial en Git (branch feature/recurrent-inspections)
- Preparación de estructura de trabajo

**Responde con "ADELANTE" para comenzar la implementación.**

---

## 📝 APÉNDICE A: Actualización de Configuration2.bas Post-Migración

### Documentación a Actualizar en Configuration2.bas

Después de completar la migración (Fase 1), actualizar manualmente:

```vb
' ----------------------------------------------------------------------
' TABLA: tblInspecciones
' Ubicación: Hoja "Historico"
' Propósito: Registro de inspecciones completadas con resultados de scoring,
'            RPN y categorización. ACTUALIZADO para inspecciones recurrentes.
' Total columnas: 40 (31 originales + 9 nuevas)
' Última actualización: 21/04/2026 - Sistema de inspecciones recurrentes
' ----------------------------------------------------------------------
' COLUMNAS VERIFICADAS (21/04/2026):
'   [01-19] Columnas originales sin cambios
'   [20] Auditoria Procesos Resultado - String (NUEVA - agregada recientemente)
'   [21-31] Resto de columnas originales sin cambios
'   
'   NUEVAS COLUMNAS - Sistema Recurrente:
'   [32] Numero Inspeccion           - Long (1=primera, 2=segunda, 3+...)
'   [33] Es Inspeccion Recurrente    - String ("Si"/"No")
'   [34] Puesto Evaluado             - String (puesto ESPECÍFICO de esta inspección)
'   [35] RPN Anterior Manual         - Double NULL (RPN ingresado manualmente)
'   [36] ID Inspeccion Anterior      - String NULL (UUID si histórico automático)
'   [37] RPN Promedio                - Double NULL ((RPN Ant + RPN Act)/2)
'   [38] Porcentaje Recuperacion     - Double NULL (Futuro: microbiología)
'   [39] Porcentaje OOL              - Double NULL (Futuro: microbiología)
'   [40] RPN Total                   - Double NULL (RPN Prom + %Rec + %OOL)
'
' MÓDULOS QUE USAN ESTA TABLA (ACTUALIZADOS):
'   - InspectionRepository.bas (ESCRITURA - columnas 32-40)
'   - InspectionScheduler.bas (LECTURA - columnas 32,34,40)
'   - ChecklistOrchestrator.bas (vía Repository - columnas 32-40)
'   - CertificadoPDFGenerator.bas (LECTURA - columnas 32-40 para PDF)
'   - InspectionHistoryService.bas (NUEVO - LECTURA columnas 10,13,21,32,34,40)
'   - RecurrentInspectionCalculator.bas (NUEVO - usa datos de columnas 37,40)
' ----------------------------------------------------------------------
```

### Nuevas Constantes a Agregar (Opcional)

```vb
' ============================================================================
' CONFIGURACIÓN DE INSPECCIONES RECURRENTES
' Última actualización: 21/04/2026 - Sistema RPN ajustado
' ============================================================================

' ----------------------------------------------------------------------
' Constantes de tipo de cálculo RPN
' ----------------------------------------------------------------------
Public Const TIPO_CALCULO_TA As String = "TA"
Public Const TIPO_CALCULO_RPN_TOTAL As String = "RPN_TOTAL"

' ----------------------------------------------------------------------
' Constantes de modo de RPN anterior
' ----------------------------------------------------------------------
Public Const MODO_RPN_AUTOMATICO As String = "AUTOMATICO"
Public Const MODO_RPN_MANUAL As String = "MANUAL"

' ----------------------------------------------------------------------
' Constantes de validación de número de inspección
' ----------------------------------------------------------------------
Public Const NUMERO_INSPECCION_MINIMO As Long = 1
Public Const NUMERO_INSPECCION_RECURRENTE_MINIMO As Long = 2
```

---

## 📝 APÉNDICE B: Checklist de Validación de Columnas en Excel

**Antes de ejecutar migración, verificar manualmente en Excel:**

### tblInspecciones (Hoja "Historico")
- [x] La tabla tiene exactamente 31 columnas (verificado)
- [x] Columna 01 se llama "ID Inspeccion" (sin acento)
- [x] Columna 10 se llama "Iniciales personal"
- [x] Columna 11 se llama "ID Plantilla"
- [x] Columna 13 se llama "Fecha inspeccion" (sin acento)
- [x] Columna 20 se llama "Auditoria Procesos Resultado" (NUEVA)
- [x] Columna 21 se llama "RPN calculado"
- [x] Columna 22 se llama "Categoria resultado"
- [x] Columna 31 se llama "Usuario completado"
- [x] No hay columnas extra después de la 31

### tblPlantillas (Hoja "Checklist")
- [ ] Columna "Puesto" existe
- [ ] Columna "Frecuencia meses" existe
- [ ] Todos los registros tienen valor numérico en "Frecuencia meses"

### tblPersonal (Hoja "Personal")
- [ ] Columna 01 se llama "Iniciales"
- [ ] Columnas 03-13 son puestos (ver GetPuestosColumns)
- [ ] Columna 14 se llama "Activo"

**Si alguna validación falla:** DETENER y corregir antes de migración.

---

**Fecha creación:** 21/04/2026
**Última actualización:** 21/04/2026 - Análisis detallado de estructura BD
**Versión:** 1.2 - Plan con análisis completo de Configuration2.bas
