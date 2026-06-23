# 🔧 TODO: REFACTORIZACIÓN DEL SISTEMA DE NAVEGACIÓN

**Creado:** 08/06/2026  
**Estado:** 📋 PLANIFICADO  
**Prioridad:** 🔴 ALTA  
**Objetivo:** Simplificar el sistema de navegación, eliminar redundancias, centralizar `ScreenUpdating`, y lograr que sea fluido (cero parpadeo, sin "pegarse").

---

## 🎯 REGLAS DE NEGOCIO (INMUTABLES)

| # | Regla | Estado Actual |
|---|-------|---------------|
| R1 | "Menú principal" SIEMPRE visible | ✅ Cumple |
| R2 | Al navegar a un módulo → se muestra esa hoja, se ocultan las demás | ✅ Cumple (pero con redundancia) |
| R3 | Al hacer clic en pestaña/botón "Menú principal" → se ocultan TODAS las demás | ✅ Cumple (pero con doble ejecución) |
| R4 | Audit Trail muestra sus 5 hojas juntas al navegar hacia él | ✅ Cumple |
| R5 | Cero parpadeo visual → `ScreenUpdating` gestionado en UN solo lugar | ❌ INCUMPLE (6 lugares distintos) |
| R6 | Subs individuales por módulo para activar/desactivar rápido | ❌ INCUMPLE (no existen como patrón) |
| R7 | `Workbook_SheetActivate` es el guardián central de visibilidad | ❌ INCUMPLE (compite con NavigationService2) |

---

## � MAPEO REAL HOJAS (CORREGIDO 08/06/2026)

> ⚠ **CORRECCIONES YA APLICADAS** al código legacy. El código en los `.bas` ahora coincide con las hojas reales del libro.

| Archivo .bas | Nombre REAL en Excel | Nombre ANTERIOR (erróneo) | Estado |
|-------------|---------------------|--------------------------|--------|
| `Hoja1.bas` | Menú principal | Menú principal | ✅ OK |
| `Hoja2.bas` | **Configuración** | ~~Control de cambios~~ | ✅ CORREGIDO |
| `Hoja3.bas` | Checklist | Checklist | ✅ OK |
| `Hoja4.bas` | Aseguramiento de calidad | Aseguramiento de calidad | ✅ OK |
| `Hoja5.bas` | Personal | Personal | ✅ OK |
| `Hoja6.bas` | Historico | Historico | ✅ OK |
| `Hoja7.bas` | **Audit trail 1** | ~~Registro errores~~ | ✅ REESCRITO (patrón audit trail) |
| `Hoja8.bas` | Cronograma | Cronograma | 🟡 ARCHIVO FALTA |
| `Hoja9.bas` | **Audit trail 2** | Audit trail 2 | ✅ CORREGIDO (`ApplyRoleBasedProtection`) |
| `Hoja10.bas` | Registro de errores | Registro de errores | 🟡 ARCHIVO FALTA |
| `Hoja11.bas` | Formulario de inspeccion | Formulario de inspeccion | 🟡 ARCHIVO FALTA |
| `Hoja12.bas` | **Audit trail 3** | ~~PersonalProducción~~ | ✅ REESCRITO (patrón audit trail) |
| `Hoja13.bas` | **Audit trail 4** | ~~Observaciones~~ | ✅ REESCRITO (patrón audit trail) |
| `Hoja14.bas` | **Audit trail 5** | ~~TecControlProceso~~ | ✅ REESCRITO (patrón audit trail) |
| `Hoja19.bas` | Plantilla Certificado | Plantilla Certificado | 🟡 ARCHIVO FALTA |
| `Hoja20.bas` | ~~Análisis Desvío~~ | Análisis Desvío | ❌ DEPRECATED (no existe en el libro) |

### 🔑 Patrón unificado Audit Trail (Hoja7,9,12,13,14)

Las 5 hojas de audit trail ahora comparten el mismo patrón minimalista:
- `Worksheet_Activate`: `ApplyRoleBasedProtection(Me, AUDIT_PASSWORD)` + error handler
- `Worksheet_Deactivate`: `ApplyRoleBasedProtection(Me, AUDIT_PASSWORD)` con `On Error Resume Next`
- **Sin** `Worksheet_Change`, **sin** `Worksheet_SelectionChange`, **sin** `m_oldValues`, **sin** `TableAuditor2`
- La visibilidad simultánea de las 5 se gestiona en `SheetService2.ShowAuditTrailGroup()`

---

## �📁 MÓDULOS INVOLUCRADOS

| Módulo | Rol Actual | Rol Deseado |
|--------|-----------|-------------|
| `NavigationService2.bas` | Fachada con 18 subs públicos + llamada a SheetService2 | Fachada ligera con subs por módulo + llamada al **nuevo** motor centralizado |
| `SheetService2.bas` | Motor pesado: loops sobre todas las hojas + protege + desprotege estructura | **Se refactoriza**: solo loops sobre hojas VISIBLES actualmente, sin tocar protección |
| `ThisWorkbook.bas` | `Workbook_SheetActivate` duplica lógica de ocultación | **Se simplifica**: es el guardián central que llama a los subs del nuevo motor |
| `Hoja1.bas` (Menú principal) | `Worksheet_Activate` aplica protección + refresca cronograma | **Se mantiene** refresco de cronograma. Se elimina `ScreenUpdating` local. |
| `Hoja2.bas` a `Hoja20.bas` | Cada una con `Worksheet_Activate`/`Deactivate` que aplican protección | **Se mantienen** (protección por rol es necesaria). Se elimina `ScreenUpdating` local. |
| `Hoja7,9,12,13,14.bas` (Audit Trail 1-5) | ✅ **YA CORREGIDOS** — Unificados con `ApplyRoleBasedProtection` + `AUDIT_PASSWORD`. Sin `Change`/`SelectionChange`. | **Sin cambios adicionales** en esta refactorización. |
| `Hoja20.bas` (Análisis Desvío) | ✅ **DEPRECATED** — La hoja no existe en el libro. Código se conserva como referencia. | **No se toca** — no aplica en esta refactorización. |
| `WorkbookProtector2.bas` | Protege/desprotege estructura | **Sin cambios** (ya funciona bien) |
| `SheetProtector2.bas` | Aplica protección por rol | **Sin cambios** (ya funciona bien) |
| `VariablesGlobales2.bas` | `m_userRole`, `g_PreviousSheetName` | **Se agrega** `g_NavigationInProgress As Boolean` para evitar recursión |

---

## 📋 CHECKLIST DE TAREAS

### FASE 0: PREPARACIÓN — VARIABLES GLOBALES

- [x] **T0.1** Agregar `g_NavigationInProgress As Boolean` en `VariablesGlobales2.bas` ✅ COMPLETADO 08/06/2026
  - Propósito: Bandera que evita recursión entre `Workbook_SheetActivate` y `NavigationService2`
  - Se pone `True` al iniciar navegación, `False` al terminar
  - `Workbook_SheetActivate` verifica esta bandera antes de actuar

- [x] **T0.2** Agregar constantes de agrupación de hojas en `Configuration2.bas` ✅ COMPLETADO 08/06/2026
  - `ALL_MODULE_SHEETS` — lista completa de hojas de módulo (para ocultación masiva)
  - `AUDIT_TRAIL_SHEETS` — nombres de las 5 hojas audit trail (para el grupo)
  - Esto evita recorrer `ThisWorkbook.Worksheets` cada vez

---

### FASE 1: NUEVO MOTOR CENTRALIZADO — `SheetService2.bas`

- [x] **T1.1** Crear `Public Sub ShowOnly(Boolean, ParamArray sheetNames() As Variant)` ✅ COMPLETADO 08/06/2026
  - **Reemplaza** a `HideAndProtectAllSheetsExcept`
  - Parámetro `Boolean`: `True` = aplicar protección, `False` = solo visibilidad (sin tocar protección)
  - Itera SOLO sobre las hojas en `ALL_MODULE_SHEETS` (no sobre todas las del libro)
  - Si `True`: llama a `ApplyRoleBasedProtection` UNA sola vez para las hojas visibles
  - Si `False`: solo cambia `Visible`, sin tocar protección
  - **NUNCA** toca `ScreenUpdating` ni `EnableEvents` (lo maneja el llamador)

- [x] **T1.2** Crear `Public Sub ShowAuditTrailGroup()` (REFACTORIZADO) ✅ COMPLETADO 08/06/2026
  - Usa la lista `AUDIT_TRAIL_SHEETS` en vez de llamar a `AuditRotation2.ObtenerNombreHoja()` 5 veces en loop
  - **Elimina** toda mención a `ScreenUpdating`/`EnableEvents`/`DisplayAlerts`
  - **Elimina** `UnprotectWorkbook`/`ProtectWorkbook` — eso lo maneja el llamador
  - **Elimina** la aplicación de protección — eso lo maneja `Worksheet_Activate`

- [x] **T1.3** Marcar como **DEPRECATED** los métodos antiguos ✅ COMPLETADO 08/06/2026
  - Agregar comentario `' @deprecated — Usar ShowOnly en su lugar` en `HideAndProtectAllSheetsExcept`
  - Mantener el código por compatibilidad durante la transición

- [x] **T1.4** Eliminar todo `Debug.Print` verbose de `SheetService2.bas` ✅ COMPLETADO 08/06/2026
  - Actualmente ~30 líneas de debug por cada llamada → overhead innecesario

---

### FASE 2: REFACTORIZACIÓN — `NavigationService2.bas`

- [x] **T2.1** Crear `Private Sub BeginNavigation()` y `Private Sub EndNavigation()` ✅ COMPLETADO 08/06/2026
  - **ÚNICOS** lugares en todo el proyecto que tocan `ScreenUpdating`/`EnableEvents`/`DisplayAlerts`
  - `BeginNavigation`: `ScreenUpdating=False`, `EnableEvents=False`, `DisplayAlerts=False`, `g_NavigationInProgress=True`
  - `EndNavigation`: `ScreenUpdating=True`, `EnableEvents=True`, `DisplayAlerts=True`, `g_NavigationInProgress=False`
  - Con manejo de errores: si algo falla, `EndNavigation` se llama en el `ErrorHandler`

- [x] **T2.2** Refactorizar `NavigateToSheet(targetSheetName)` para usar `BeginNavigation`/`EndNavigation` ✅ COMPLETADO 08/06/2026
  - Eliminar las líneas manuales de `ScreenUpdating = False/True`
  - Flujo: `BeginNavigation` → `ShowOnly(False, targetSheetName, "Menú principal")` → `.Select` → `EndNavigation`
  - **Eliminar** la llamada a `SheetService2.HideAndProtectAllSheetsExcept` → usar `ShowOnly` nuevo
  - **Eliminar** `UnprotectWorkbook`/`ProtectWorkbook` de aquí (eso lo hace `ShowOnly` si es necesario)

- [x] **T2.3** Crear subs individuales por módulo (patrón limpio) ✅ COMPLETADO 08/06/2026
  ```vba
  Public Sub MostrarDetecciones()
  Public Sub MostrarDashboard()
  Public Sub MostrarObservaciones()
  Public Sub MostrarRechazo()
  Public Sub MostrarDesvio()
  Public Sub MostrarConfiguracion()
  Public Sub MostrarPersonal()
  Public Sub MostrarAseguramientoCalidad()
  Public Sub MostrarTecControlProceso()
  Public Sub MostrarAuditTrail()
  Public Sub MostrarControlDeCambios()
  Public Sub MostrarMenu()
  Public Sub MostrarChecklistVirtual()
  Public Sub MostrarResultados()
  Public Sub MostrarConfiguracionChecklist()
  Public Sub MostrarCronograma()
  Public Sub MostrarPlantillaCertificado()
  ```
  - Cada uno llama a `NavigateToSheet("NombreHoja")` o a la lógica específica para Audit Trail
  - **Compatibilidad**: mantener los nombres antiguos `NavigateToXxx` como wrappers que llaman a `MostrarXxx`

- [x] **T2.4** Refactorizar `NavigateToAuditTrail` → `MostrarAuditTrail` ✅ COMPLETADO 08/06/2026
  - Usar `BeginNavigation`/`EndNavigation`
  - `ShowOnly(False)` para ocultar todo excepto menú
  - `ShowAuditTrailGroup()` (nuevo, sin tocar flags)
  - `.Select` primera hoja audit
  - **Eliminar** la llamada redundante a `AuditLogger2.LogAction` (eso ya lo hace `Workbook_SheetActivate`)

- [x] **T2.5** Reducir logs de Audit Trail en navegación ✅ COMPLETADO 08/06/2026
  - `AuditLogger2.LogAction` en cada `NavigateToSheet` → solo registrar navegaciones a módulos "sensibles" (Configuración, Audit Trail)
  - Las navegaciones diarias (Detecciones, Dashboard, etc.) no necesitan audit trail → reduce I/O

---

### FASE 3: SIMPLIFICACIÓN — `ThisWorkbook.bas`

- [x] **T3.1** Refactorizar `Workbook_SheetActivate` como guardián central ✅ COMPLETADO 08/06/2026
  - **ANTES** de actuar: `If g_NavigationInProgress Then Exit Sub` ← evita competencia con NavigationService2
  - Si `Sh.Name = "Menú principal"` → llamar a `NavigationService2.MostrarMenu` (no duplicar lógica)
  - Si `Sh.Name` NO es Audit Trail → `Sh.Visible = xlSheetVisible` (solo si no está ya visible)
  - **Eliminar** el bloque que llama a `HideAndProtectAllSheetsExcept("Menú principal")` directamente
  - **Eliminar** `ScreenUpdating`/`EnableEvents`/`DisplayAlerts` locales

- [x] **T3.2** Simplificar `Workbook_Open` ✅ COMPLETADO 08/06/2026
  - Ya funciona bien. Solo verificar que use `BeginNavigation`/`EndNavigation` si aplica

- [x] **T3.3** Verificar `Workbook_SheetDeactivate` ✅ COMPLETADO 08/06/2026
  - Debe permanecer **vacío** (sin lógica). Ya está correcto.

---

### FASE 4: LIMPIEZA — `Hoja*.bas` (TODAS las hojas)

- [x] **T4.1** `Hoja1.bas` (Menú principal) — `Worksheet_Activate` ✅ COMPLETADO 08/06/2026
  - **Eliminar** `Application.ScreenUpdating = False/True`
  - **Mantener** `ApplyRoleBasedProtection` + `RefrescarResumenCronograma`
  - Agregar al inicio: `If g_NavigationInProgress Then Exit Sub` (evitar doble ejecución)

- [x] **T4.2** `Hoja1.bas` (Menú principal) — `Worksheet_Deactivate` ✅ COMPLETADO 08/06/2026
  - **Mantener** `ApplyRoleBasedProtection` (sin cambios)

- [x] **T4.3** `Hoja1.bas` (Menú principal) — `Worksheet_Change` ✅ COMPLETADO 08/06/2026
  - **Eliminar** `Application.EnableEvents = False/True` manual
  - Usar el patrón: `If g_NavigationInProgress Then Exit Sub`

- [x] **T4.4** `Hoja9.bas` (Audit trail 2) — ✅ CORREGIDO (08/06/2026)
  - `ProtectSheet` → `ApplyRoleBasedProtection` con `AUDIT_PASSWORD`
  - Agregado `On Error Resume Next` en `Worksheet_Deactivate`

- [x] **T4.5** `Hoja7,12,13,14.bas` (Audit trail 1,3,4,5) — ✅ REESCRITOS (08/06/2026)
  - Código legacy eliminado completamente (no tenían `ScreenUpdating`)
  - Ahora siguen el patrón unificado de audit trail
  - Sin `Worksheet_Change`, sin `Worksheet_SelectionChange`, sin `m_oldValues`

- [x] **T4.6** Hojas restantes (Hoja2, Hoja3, Hoja4, Hoja5, Hoja6) ✅ COMPLETADO 08/06/2026
  - **Eliminar** `Application.ScreenUpdating = False` del `Worksheet_Change` de cada una
  - **Mantener** `ApplyRoleBasedProtection` en `Worksheet_Activate`/`Deactivate`
  - Agregar `If g_NavigationInProgress Then Exit Sub` al inicio de cada `Worksheet_Activate`

- [x] **T4.7** `Hoja20.bas` — ✅ DEPRECATED (08/06/2026). La hoja no existe en el libro.

---

### FASE 5: OPTIMIZACIÓN — `WorkbookProtector2.bas`

- [x] **T5.1** Agregar `Public Sub ToggleProtection(ByVal enable As Boolean)` ✅ COMPLETADO 08/06/2026
  - Si `enable=True` → `ProtectWorkbook`
  - Si `enable=False` → `UnprotectWorkbook`
  - Simplifica las llamadas desde el motor de navegación

- [x] **T5.2** Considerar caché de estado de protección ✅ COMPLETADO 08/06/2026
  - Variable `Private m_IsProtected As Boolean` para evitar desproteger si ya está desprotegido
  - Reduce operaciones redundantes de protección/desprotección

---

### FASE 6: TESTING Y VERIFICACIÓN

- [x] **T6.1** Prueba de humo — Navegación básica ✅ VERIFICADO 08/06/2026
  - Abrir libro → solo "Menú principal" visible ✅
  - Clic en botón "Detecciones" → aparece Detecciones, demás ocultas ✅
  - Clic en pestaña "Menú principal" → oculta Detecciones, solo Menú visible ✅
  - Sin parpadeo ni saltos visuales ✅

- [x] **T6.2** Prueba — Audit Trail ✅ VERIFICADO 08/06/2026
  - Clic en "Audit Trail" → 5 hojas visibles + Menú principal ✅
  - Clic en "Menú principal" → las 5 hojas audit se ocultan ✅
  - Las hojas audit NO se pueden mostrar manualmente (xlSheetVeryHidden) ✅

- [x] **T6.3** Prueba — Navegación consecutiva rápida ✅ VERIFICADO 08/06/2026
  - Clic rápido: Detecciones → Dashboard → Configuración → Menú
  - Sin errores, sin parpadeo, sin "pegarse" ✅

- [x] **T6.4** Prueba — Doble clic en tblResumenCronograma ✅ VERIFICADO 08/06/2026
  - Abre frmSelectorInspeccion correctamente ✅
  - Al cerrar, Menú principal sigue visible ✅

- [x] **T6.5** Prueba — Checklist Virtual ✅ VERIFICADO 08/06/2026
  - Abrir desde selector → formulario modal se muestra ✅
  - Al guardar y cerrar, Menú principal sigue visible ✅

- [ ] **T6.6** Prueba — Modo Admin vs Usuario ⚠️ PENDIENTE — Se implementará con bloqueo por rol
  - Admin: puede editar hojas de módulo
  - Usuario: solo lectura con copiado
  - La protección se aplica UNA sola vez por navegación

- [ ] **T6.7** Prueba — Cierre y re-apertura del libro ⚠️ PENDIENTE — Se implementará con bloqueo por rol
  - Al abrir: estado limpio, solo Menú principal
  - Sin errores en `Workbook_Open`

- [ ] **T6.8** Prueba — Estrés ⚠️ PENDIENTE — Se implementará con bloqueo por rol
  - 20 navegaciones consecutivas entre módulos distintos
  - Sin degradación de rendimiento
  - Sin errores 1004 (protección de estructura)

---

### FASE 7: DOCUMENTACIÓN

- [ ] **T7.1** Actualizar `SISTEMA_NAVEGACION_VISIBILIDAD.md` con la nueva arquitectura
  - Diagrama actualizado de flujo
  - Explicación de `BeginNavigation`/`EndNavigation`
  - Explicación del patrón `MostrarXxx`

- [ ] **T7.2** Actualizar `GUIA_MAESTRA_SISTEMA_INSPECCIONES.md`
  - Referencia a la nueva arquitectura de navegación
  - Índice de subs `MostrarXxx` disponibles

- [ ] **T7.3** Agregar comentarios `' ## NAVEGACIÓN ##` en cada módulo modificado
  - Para facilitar búsquedas futuras con `grep`

---

## 📊 RESUMEN DE CAMBIOS POR MÓDULO

| Módulo | Tipo de Cambio | Líneas Afectadas (est.) | Estado |
|--------|---------------|------------------------|--------|
| `VariablesGlobales2.bas` | AGREGAR 1 variable | +3 | ✅ COMPLETADO 08/06/2026 |
| `Configuration2.bas` | AGREGAR 2 constantes | +15 | ✅ COMPLETADO 08/06/2026 |
| `NavigationService2.bas` | REFACTORIZAR completo | ~350 líneas | ✅ COMPLETADO 08/06/2026 |
| `SheetService2.bas` | REFACTORIZAR (nuevo `ShowOnly`) | ~200 líneas | ✅ COMPLETADO 08/06/2026 |
| `ThisWorkbook.bas` | SIMPLIFICAR `Workbook_SheetActivate` | ~30 líneas | ✅ COMPLETADO 08/06/2026 |
| `Hoja1.bas` (Menú) | LIMPIAR `ScreenUpdating` | -6 líneas | ✅ COMPLETADO 08/06/2026 |
| `Hoja2.bas` | LIMPIAR `ScreenUpdating` | -2 líneas | ✅ COMPLETADO 08/06/2026 |
| `Hoja2.bas` | CORREGIR nombre (era "Control de cambios") | ~5 reemplazos | ✅ HECHO 08/06/2026 |
| `Hoja3.bas` | LIMPIAR `ScreenUpdating` | -2 líneas | ✅ COMPLETADO 08/06/2026 |
| `Hoja4.bas` | AGREGAR guardia `g_NavigationInProgress` | +3 líneas | ✅ COMPLETADO 08/06/2026 |
| `Hoja5.bas` | AGREGAR guardia `g_NavigationInProgress` | +3 líneas | ✅ COMPLETADO 08/06/2026 |
| `Hoja6.bas` | LIMPIAR `ScreenUpdating` + guardia | +3 líneas | ✅ COMPLETADO 08/06/2026 |
| `Hoja7.bas` | REESCRITO (era "Registro errores" → Audit trail 1) | ~80→28 líneas | ✅ HECHO 08/06/2026 |
| `Hoja9.bas` | CORREGIDO `ProtectSheet` → `ApplyRoleBasedProtection` | ~6 líneas | ✅ HECHO 08/06/2026 |
| `Hoja12.bas` | REESCRITO (era "PersonalProducción" → Audit trail 3) | ~80→28 líneas | ✅ HECHO 08/06/2026 |
| `Hoja13.bas` | REESCRITO (era "Observaciones" → Audit trail 4) | ~65→28 líneas | ✅ HECHO 08/06/2026 |
| `Hoja14.bas` | REESCRITO (era "TecControlProceso" → Audit trail 5) | ~65→28 líneas | ✅ HECHO 08/06/2026 |
| `Hoja20.bas` | MARCADO DEPRECATED | ~80 líneas | ✅ HECHO 08/06/2026 |
| `WorkbookProtector2.bas` | AGREGAR `ToggleProtection` + caché | +15 líneas | ✅ COMPLETADO 08/06/2026 |

---

## 🚦 ORDEN DE IMPLEMENTACIÓN RECOMENDADO

```
FASE 0 → FASE 5 → FASE 1 → FASE 2 → FASE 3 → FASE 4 → FASE 6 → FASE 7
```

1. **Primero** variables globales y constantes (FASE 0 + FASE 5) — sin romper nada
2. **Luego** el nuevo motor `ShowOnly` (FASE 1) — coexiste con el viejo
3. **Luego** refactorizar `NavigationService2` (FASE 2) — empieza a usar el nuevo motor
4. **Luego** simplificar `ThisWorkbook` (FASE 3) — elimina la competencia
5. **Luego** limpiar las hojas (FASE 4) — elimina `ScreenUpdating` locales
6. **Finalmente** testear todo (FASE 6) y documentar (FASE 7)

---

## ⚠️ PRECAUCIONES

1. **NO eliminar** los métodos `NavigateToXxx` antiguos — mantenerlos como wrappers que llaman a `MostrarXxx`. Esto mantiene compatibilidad con botones existentes en las hojas.

2. **NO modificar** `SheetProtector2.bas` — funciona correctamente y es usado por todo el sistema.

3. **NO modificar** `AdminAccessControl2.bas` — la autenticación no es parte de este fix.

4. **El orden importa**: si se implementa FASE 2 antes de FASE 1, el sistema quedará en un estado roto.

5. **Hacer backup** del archivo `.xlsm` antes de comenzar cualquier modificación.

---

## 🎯 MÉTRICAS DE ÉXITO POST-REFACTOR

| Métrica | Antes | Después (objetivo) |
|---------|-------|-------------------|
| Lugares que tocan `ScreenUpdating` | 6 | **1** (`NavigationService2`) |
| Veces que se aplica protección por navegación | 2-3 | **1** (solo `Worksheet_Activate`) |
| Hojas recorridas en cada navegación | TODAS (~25) | **Solo las visibles** (~2-5) |
| Módulos que modifican visibilidad | 3 | **2** (`NavigationService2` + `Workbook_SheetActivate` como guardián) |
| `Debug.Print` por navegación | ~30 líneas | **~5 líneas** |
| Tiempo estimado por navegación | ~1-2 segundos | **<0.3 segundos** |
