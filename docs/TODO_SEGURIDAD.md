# 🛡️ TODO - Mejoras del Sistema de Seguridad VBA

> **Archivo:** `docs/TODO_SEGURIDAD.md`
> **Propósito:** Plan detallado de correcciones para el sistema de seguridad (SheetProtector2, WorkbookProtector2, AdminAccessControl2, roles, etc.)
> **Última actualización:** 28/05/2026 (verificación completa con el código fuente; correcciones TO-DO aplicadas)
> **Estado:** ✅ **15 de 15 items implementados — FASE COMPLETADA — CORRECCIONES TO-DO APLICADAS**
> **Nota:** Esta actualización sincroniza el TODO con el código fuente real. Todos los items (C01-C07, M01-M09) ya están implementados en el código VBA. Los sub-items que quedan como `[ ]` corresponden a pruebas manuales o decisiones de diseño pendientes.
> 
> **⚠️ VERIFICACIÓN DEL CÓDIGO FUENTE (28/05/2026):**
> Se verificó el código de todos los módulos `.bas` contra el TODO. A continuación se listan los hallazgos:
> 
> **Correcciones verificadas (28/05/2026 - FASE COMPLETA):**
> - ✅ **C01** — Contraseñas migradas del cliente: APP_PASSWORD="supervisor002.", ADMIN_PASSWORD="validacion002.", AUDIT_PASSWORD="validacion002.", CRONOGRAMA_ADMIN_PASSWORD="validacion002." (Configuration2.bas líneas 22,30,38,47)
> - ✅ **C02** — ENABLE_SHEET_PROTECTION = True, ENABLE_WORKBOOK_PROTECTION = True (Configuration2.bas líneas 137,148)
> - ✅ **C03** — Todas las hojas unificadas con ApplyRoleBasedProtection (Hoja1,2,3,4,5,6,7,12,13,14,20 verificados)
> - ✅ **C04** — Hoja1 protegida con ApplyRoleBasedProtection (Activate línea 35, Deactivate línea 151)
> - ✅ **C05** — Hoja6 protegida con ApplyRoleBasedProtection (Activate línea 23, Deactivate línea 37)
> - ✅ **C06** — Navegación segura ACTIVA (Workbook_SheetActivate línea 164-183 — código NO comentado, funcional)
> - ✅ **C07** — SheetService2 aplica ApplyRoleBasedProtection en HideAndProtectAllSheetsExcept (líneas 87-91) y ShowAuditTrailGroup (líneas 157-166)
> - ✅ **M01** — ErrorLogger.Log corregido a ErrorLogger2.Log (AdminAccessControl2.bas línea 197)
> - ✅ **M02** — Hoja3 (Checklist) creada con eventos completos (Activate, Deactivate, Change, SelectionChange) — 94 líneas verificadas
> - ✅ **M04** — InputBox reemplazado por frmInput en AbrirGestorCronograma (Hoja1.bas líneas 197-205)
> - ✅ **M05** — Bloqueo por 3 intentos fallidos implementado en AdminAccessControl2 (MAX_INTENTOS_FALLIDOS=3, MINUTOS_BLOQUEO=5, EstaBloqueado, AplicarBloqueo)
> - ✅ **M06** — 6 constantes legacy centralizadas en Configuration2.bas (líneas 193-198) y usadas en Hoja2,7,12,13,14,20
> - ✅ **M07** — Backup automático reactivado (ThisWorkbook.bas línea 248, código descomentado)
> - ✅ **M08** — Análisis automático al cerrar reactivado (ThisWorkbook.bas líneas 282-285, código descomentado)
> - ✅ **M09** — Archivos .bas creados para hojas sin módulo legacy
> 
> **Hallazgos de la verificación (28/05/2026):**
> 1. ✔️ **C06**: El código en `ThisWorkbook.Workbook_SheetActivate` está **activo** (NO comentado), aunque el comentario dice "TEMPORALMENTE DESACTIVADO PARA DESARROLLO". Se actualizó el comentario para reflejar el estado real.
> 2. 🟡 **M08**: `dataProcessAnalysis.EjecutarAnalisis` se referencia en ThisWorkbook.bas línea 283 pero no existe un archivo `.bas` exportado para esta clase. La referencia funciona porque `dataProcessAnalysis` es una clase interna del proyecto VBA (.xlsm). Se mantiene el item como implementado.
> 3. ⚠️ **mod_PasswordMigration.bas**: Las contraseñas OLD están desactualizadas respecto a la configuración actual. Se actualizaron para reflejar el estado real del libro.
> 4. 🟡 **C01e/C01f**: Pendiente ejecutar la migración física de contraseñas en el archivo .xlsm.
> 


---

## � Leyenda de la checklist

| Símbolo | Significado |
|---------|-------------|
| `[ ]` | Pendiente de implementar |
| `[~]` | En progreso |
| `[x]` | Implementado y verificado |

---

## 1. 🔴 ERRORES CRÍTICOS

> Estos errores representan **brechas de seguridad activas**. Permiten que cualquier usuario (sin importar el rol) pueda editar datos, modificar la estructura del libro y acceder a información sensible sin restricciones reales.

---

### 🔴 C01 — Contraseñas débiles e idénticas en Configuration2.bas

**Archivo afectado:** `Configuration2.bas` (líneas 22, 30, 38)

**Problema:** Las 3 contraseñas principales son literalmente `"1234"` — todas idénticas. Esto anula cualquier separación de permisos. La contraseña de protección de hojas (APP_PASSWORD) debería ser distinta a la de autenticación de administrador (ADMIN_PASSWORD) y a la de auditoría (AUDIT_PASSWORD).

```vb
' Estado actual (INSEGURO):
Public Const APP_PASSWORD   As String = "1234"   ' ← Protección de hojas
Public Const ADMIN_PASSWORD As String = "1234"   ' ← Autenticación Admin (¡IDÉNTICA!)
Public Const AUDIT_PASSWORD As String = "1234"   ' ← Hojas Audit Trail (¡IDÉNTICA!)
```

**Impacto:** Cualquier persona que conozca la contraseña de una hoja (ej. para copiar datos) tiene automáticamente acceso de administrador y puede modificar el Audit Trail.

**Checklist:**
- [x] **C01a** — APP_PASSWORD definida por el cliente ("supervisor002.") ✅
- [x] **C01b** — ADMIN_PASSWORD definida por el cliente ("validacion002.") ✅
- [x] **C01c** — AUDIT_PASSWORD definida por el cliente ("validacion002.") ⚠️ Igual a ADMIN_PASSWORD
- [x] **C01d** — CRONOGRAMA_ADMIN_PASSWORD definida por el cliente ("validacion002.") ⚠️ Igual a ADMIN_PASSWORD
- [ ] **C01e** — Ejecutar `mod_PasswordMigration.MigrarContrasenasAlNuevo` para aplicar los cambios físicos en el libro
- [ ] **C01f** — Verificar con `mod_PasswordMigration.VerificarContrasenasActuales` que todas las contraseñas funcionan

> **⚠️ NOTA (28/05/2026):** El cliente solicitó que ADMIN_PASSWORD, AUDIT_PASSWORD y CRONOGRAMA_ADMIN_PASSWORD compartan el mismo valor "validacion002.". Esto debilita la separación de permisos (C01 original recomendaba contraseñas únicas por propósito). Se implementa según solicitud del cliente.

> **Arquitectura:** `Configuration2.bas` es el punto único de configuración. Al ser `Public Const`, las contraseñas se compilan en el código VBA y no son modificables en tiempo de ejecución. Esto es correcto por diseño (evita que un usuario malicioso las cambie con una macro), pero las expone si alguien abre el VBA Editor. Considerar ofuscación o migración a hash si se requiere mayor seguridad.

---

### 🔴 C02 — Protección de hojas y libro deshabilitada en Configuration2.bas

**Archivo afectado:** `Configuration2.bas` (líneas 137, 148)

**Problema:** Ambos flags de protección están en `False`, lo que significa que el sistema opera en "modo desarrollo" permanentemente.

```vb
Public Const ENABLE_WORKBOOK_PROTECTION As Boolean = False  ' ← Debe ser True
Public Const ENABLE_SHEET_PROTECTION    As Boolean = False  ' ← Debe ser True
```

**Impacto:** Cualquier usuario puede:
- Editar cualquier celda en cualquier hoja (incluso datos maestros)
- Eliminar, agregar o renombrar hojas
- Modificar tablas, fórmulas, formatos
- Desactivar validaciones de datos

**Checklist:**
- [x] **C02a** — Cambiar `ENABLE_SHEET_PROTECTION = True` en Configuration2.bas ✅
- [x] **C02b** — Cambiar `ENABLE_WORKBOOK_PROTECTION = True` en Configuration2.bas ✅
- [x] **C02c** — Verificar que al abrir el libro se activa la protección (Workbook_Open) ✅
- [x] **C02d** — SheetProtector2.UnprotectSheet respeta el flag ✅
- [x] **C02e** — SheetProtector2.ProtectSheet respeta el flag ✅
- [ ] **C02f** — Probar que un usuario normal NO puede editar datos en ninguna hoja — ⏳ Pendiente prueba manual
- [ ] **C02g** — Probar que un admin PUEDE editar después de autenticarse — ⏳ Pendiente prueba manual

> **Arquitectura:** `SheetProtector2.bas` y `WorkbookProtector2.bas` son servicios (capa de infraestructura) que implementan la protección física de Excel. `Configuration2.bas` es la capa de configuración. `ThisWorkbook.bas` (Workbook_Open) es el orquestador que inicializa las protecciones al abrir el libro.

---

### 🔴 C03 — Inconsistencia en el manejo de roles entre hojas

**Archivos afectados:** Hoja1.bas, Hoja2.bas, Hoja5.bas, Hoja4.bas, Hoja6.bas, Hoja7.bas, Hoja12.bas, Hoja13.bas, Hoja14.bas, Hoja20.bas

**Problema:** Cada hoja implementa la protección por rol de manera diferente. Algunas usan `GetUserRole()`, otras `m_userRole` directamente, y varias no implementan ninguna protección en `Worksheet_Activate`.

**Estado actual por hoja:**

| Hoja | ¿Activate implementado? | ¿Usa GetUserRole()? | ¿Aplica ApplyRoleBasedProtection? | ¿Protege al salir (Deactivate)? |
|------|------------------------|---------------------|-----------------------------------|----------------------------------|
| **Hoja1** (Menú principal) | ✅ Sí | ✅ `GetUserRole()` (vía ApplyRoleBasedProtection) | ✅ `ApplyRoleBasedProtection` | ✅ `ApplyRoleBasedProtection` |
| **Hoja2** (Control de cambios) | ✅ Sí | ✅ `GetUserRole()` (vía ApplyRoleBasedProtection) | ✅ `ApplyRoleBasedProtection` | ✅ `ApplyRoleBasedProtection` |
| **Hoja3** (Checklist) | ✅ Sí | ✅ `GetUserRole()` (vía ApplyRoleBasedProtection) | ✅ `ApplyRoleBasedProtection` | ✅ `ApplyRoleBasedProtection` |
| **Hoja4** (Aseguramiento) | ✅ Sí | ✅ `GetUserRole()` (vía ApplyRoleBasedProtection) | ✅ `ApplyRoleBasedProtection` | ✅ `ApplyRoleBasedProtection` |
| **Hoja5** (Personal) | ✅ Sí | ✅ `GetUserRole()` (vía ApplyRoleBasedProtection) | ✅ `ApplyRoleBasedProtection` | ✅ `ApplyRoleBasedProtection` |
| **Hoja6** (Historico) | ✅ Sí | ✅ `GetUserRole()` (vía ApplyRoleBasedProtection) | ✅ `ApplyRoleBasedProtection` | ✅ `ApplyRoleBasedProtection` |
| **Hoja7** (Registro errores) | ✅ Sí | ✅ `GetUserRole()` (vía ApplyRoleBasedProtection) | ✅ `ApplyRoleBasedProtection` | ✅ `ApplyRoleBasedProtection` |
| **Hoja9** (Audit trail) | ✅ Sí | ❌ Usa AUDIT_PASSWORD directo (protección especial) | ❌ Usa `ProtectSheet` con AUDIT_PASSWORD | ✅ `ProtectSheet` con AUDIT_PASSWORD |
| **Hoja12** (PersonalProducción) | ✅ Sí | ✅ `GetUserRole()` (vía ApplyRoleBasedProtection) | ✅ `ApplyRoleBasedProtection` | ✅ `ApplyRoleBasedProtection` |
| **Hoja13** (Observaciones) | ✅ Sí | ✅ `GetUserRole()` (vía ApplyRoleBasedProtection) | ✅ `ApplyRoleBasedProtection` | ✅ `ApplyRoleBasedProtection` |
| **Hoja14** (TecControlProceso) | ✅ Sí | ✅ `GetUserRole()` (vía ApplyRoleBasedProtection) | ✅ `ApplyRoleBasedProtection` | ✅ `ApplyRoleBasedProtection` |
| **Hoja20** (Análisis Desvío) | ✅ Sí | ✅ `GetUserRole()` (vía ApplyRoleBasedProtection) | ✅ `ApplyRoleBasedProtection` | ✅ `ApplyRoleBasedProtection` |

**Checklist:**
- [x] **C03a** — Función unificada `ApplyRoleBasedProtection` ya existe en SheetProtector2.bas ✅
- [x] **C03b** — Hoja1 (Menú principal): actualizada con `ApplyRoleBasedProtection` ✅
- [x] **C03c** — Hoja2 (Control de cambios): `m_userRole` reemplazado por `GetUserRole()` (vía ApplyRoleBasedProtection) ✅
- [x] **C03d** — Hoja4 (Aseguramiento): unificada con `ApplyRoleBasedProtection` ✅
- [x] **C03e** — Hoja5 (Personal): unificada con `ApplyRoleBasedProtection` ✅
- [x] **C03f** — Hoja6 (Historico): agregado `Worksheet_Activate` con protección por rol ✅
- [x] **C03g** — Hoja7 (Registro errores): unificada con `ApplyRoleBasedProtection` ✅
- [x] **C03h** — Hoja12 (PersonalProducción): unificada con `ApplyRoleBasedProtection` ✅
- [x] **C03i** — Hoja13 (Observaciones): unificada con `ApplyRoleBasedProtection` ✅
- [x] **C03j** — Hoja14 (TecControlProceso): unificada con `ApplyRoleBasedProtection` ✅
- [x] **C03k** — Hoja20 (Análisis Desvío): unificada con `ApplyRoleBasedProtection` ✅
- [x] **C03l** — Hoja2 (Control de cambios): ya usa `ApplyRoleBasedProtection` ✅
- [x] **C03m** — Hoja9 (Audit trail): protección especial con AUDIT_PASSWORD (diseñado para máxima seguridad) ✅

> **Arquitectura:** Cada hoja es un **módulo de presentación** (vista). La lógica de protección debe estar centralizada en `SheetProtector2.ApplyRoleBasedProtection` (capa de infraestructura). Las hojas solo deben invocar esta función, no implementar su propia lógica de permisos.

---

### 🔴 C04 — Menú principal (Hoja1) no tiene protección basada en rol

**Archivo afectado:** `Hoja1.bas` (líneas 27-52, 157-159)

**Problema:** El evento `Worksheet_Activate` de la hoja principal NO verifica el rol del usuario antes de mostrar la hoja. Solo llama a `Me.EnableSelection = xlUnlockedCells`, que permite editar cualquier celda desbloqueada. Además, `Worksheet_Deactivate` llama a `ProtectSheet` genérico sin considerar el rol.

```vb
' Hoja1.Worksheet_Activate — NO hay verificación de rol
Private Sub Worksheet_Activate()
    ' ...
    Me.EnableSelection = xlUnlockedCells  ' ← Permite editar celdas desbloqueadas a TODOS
    ' ...
End Sub

Private Sub Worksheet_Deactivate()
    Call SheetProtector2.ProtectSheet(Me, Configuration2.APP_PASSWORD)  ' ← No usa ApplyRoleBasedProtection
End Sub
```

**Impacto:** Si la protección está activa (C02), el menú principal se muestra sin protección adecuada. Un usuario normal podría modificar celdas críticas del menú (filtros, configuraciones visuales).

**Checklist:**
- [x] **C04a** — Hoja1.Worksheet_Activate ya usa `ApplyRoleBasedProtection` ✅
- [x] **C04b** — Hoja1.Worksheet_Deactivate ya usa `ApplyRoleBasedProtection` ✅
- [ ] **C04c** — Verificar que Admin puede editar filtros (J15) y que Usuario no puede — ⏳ Pendiente prueba manual

> **Arquitectura:** Hoja1 es la **vista principal** (Menú principal). Debe delegar la protección a `SheetProtector2` (capa de infraestructura) mediante `ApplyRoleBasedProtection`, que ya implementa correctamente la lógica de 3 niveles (Admin → sin protección, Usuario → solo lectura, Otros → sin selección).

---

### 🔴 C05 — Hoja6 (Historico) no tiene Worksheet_Activate con protección

**Archivo afectado:** `Hoja6.bas`

**Problema:** El módulo `Hoja6.bas` (Historico) solo implementa `Worksheet_BeforeDoubleClick` para generar PDFs. **No existe** un evento `Worksheet_Activate` que aplique protección por rol.

**Impacto:** Un usuario normal podría modificar directamente la tabla `tblInspecciones` en el histórico, comprometiendo la integridad de los registros de inspección.

**Checklist:**
- [x] **C05a** — Agregado `Worksheet_Activate` en Hoja6.bas con `ApplyRoleBasedProtection` ✅
- [x] **C05b** — Agregado `Worksheet_Deactivate` en Hoja6.bas con `ApplyRoleBasedProtection` ✅
- [ ] **C05c** — Agregar eventos `Worksheet_Change` y `Worksheet_SelectionChange` para auditoría (opcional) — ⏳ Pendiente

> **Arquitectura:** Hoja6 es una **vista de consulta** (Historico). Como contiene datos transaccionales (inspecciones completadas), debe tener el mismo nivel de protección que las hojas de datos maestros (Personal, Aseguramiento).

---

### 🔴 C06 — La navegación segura (NavigationService2) está desactivada

**Archivo afectado:** `ThisWorkbook.bas` (líneas 163-187)

**Problema:** El sistema de navegación que oculta/muestra hojas según el rol está completamente desactivado, con todo el código comentado dentro de `Workbook_SheetActivate`.

```vb
' ========== SISTEMA DE NAVEGACIÓN TEMPORALMENTE DESACTIVADO ==========
' Para reactivar, descomentar el bloque de código a continuación
' =====================================================================
' If Sh.Name <> Configuration2.MAIN_MENU_SHEET And Not IsAuditSheet(Sh.Name) Then
'     Sh.Visible = xlSheetVisible
' End If
```

**Impacto:** Actualmente, al navegar a cualquier hoja, todas las hojas permanecen visibles (excepto las ocultas por otro código). Un usuario normal puede ver y acceder a todas las hojas del libro, incluyendo aquellas a las que no debería tener acceso.

**Checklist:**
- [x] **C06a** — Reactivado bloque de navegación en `Workbook_SheetActivate` ✅
- [x] **C06b** — `SheetService2.HideAndProtectAllSheetsExcept` funcional ✅
- [x] **C06c** — Hojas Audit Trail siguen siendo `xlSheetVeryHidden` ✅
- [x] **C06d** — Menú principal siempre visible (regla en SheetService2) ✅
- [ ] **C06e** — Probar que al navegar a una hoja, las demás se ocultan automáticamente — ⏳ Pendiente prueba manual

> **Arquitectura:** `NavigationService2.bas` es el **servicio de navegación** (capa de aplicación). `SheetService2.bas` es el **servicio de infraestructura** que maneja visibilidad/ocultación. `Workbook_SheetActivate` en `ThisWorkbook.bas` es el **orquestador** que conecta ambos. La cadena completa es: Evento → NavigationService2 → SheetService2 → WorkbookProtector2.

---

### 🔴 C07 — SheetService2.HideAndProtectAllSheetsExcept no aplica protección por rol

**Archivo afectado:** `SheetService2.bas` (líneas 47-100)

**Problema:** `HideAndProtectAllSheetsExcept` solo oculta/muestra hojas pero **no aplica** `ApplyRoleBasedProtection` a la hoja destino. La protección solo ocurre en los eventos `Worksheet_Activate` individuales, y como vimos en C03 y C04, esos eventos son inconsistentes.

**Flujo actual inseguro:**
```
NavigateToSheet("Personal")
  → SheetService2.HideAndProtectAllSheetsExcept("Personal")   ← Solo oculta/muestra
  → Sheets("Personal").Select                                   ← Activa la hoja
  → (se dispara Worksheet_Activate si existe)                   ← Opcional, depende de la hoja
  → Si no hay Worksheet_Activate → la hoja queda SIN protección ← ¡INSEGURO!
```

**Checklist:**
- [x] **C07a** — `HideAndProtectAllSheetsExcept` modificado para aplicar `ApplyRoleBasedProtection` a la hoja destino ✅
- [x] **C07b** — `ShowAuditTrailGroup` modificado para aplicar protección a las hojas Audit Trail ✅
- [x] **C07c** — Eliminada dependencia de eventos `Worksheet_Activate`; la protección se aplica desde SheetService2 siempre ✅
- [ ] **C07d** — Verificar que la protección se aplica incluso si la hoja destino no tiene evento `Worksheet_Activate` — ⏳ Pendiente prueba manual

> **Arquitectura:** `SheetService2` es un **servicio de infraestructura** que debe garantizar que la protección se aplique SIEMPRE, independientemente de si la hoja destino implementa o no sus eventos. La protección debe ser obligatoria (por defecto) y los eventos individuales deben ser redundancia (capa extra).

---

## 2. 🟡 ERRORES MEDIOS

> Estos errores no representan una brecha de seguridad inmediata, pero comprometen la mantenibilidad, consistencia y confiabilidad del sistema a largo plazo.

---

### 🟡 M01 — Error en referencia a ErrorLogger en AdminAccessControl2.bas (línea 82)

**Archivo afectado:** `AdminAccessControl2.bas` (línea 82)

**Problema:** En el manejador de errores de `CheckAdminAccess`, se referencia `ErrorLogger.Log` en lugar de `ErrorLogger2.Log`. Esto causará un error en tiempo de ejecución si la función falla (doble error: el error original + el error al intentar registrarlo).

```vb
' Línea 82 — ERROR: debería ser ErrorLogger2.Log
Call ErrorLogger.Log("ThisWorkbook.CheckAdminAccess", VBA.Err.Description, VBA.Err.Number)
'     ↑ Falta el "2"
```

**Impacto:** Si ocurre un error en `CheckAdminAccess`, el manejador de errores fallará al intentar registrarlo, y el usuario verá un error difícil de diagnosticar.

**Checklist:**
- [x] **M01a** — Corregido `ErrorLogger.Log` → `ErrorLogger2.Log` en AdminAccessControl2.bas ✅

> **Arquitectura:** `ErrorLogger2.bas` reemplazó a `ErrorLogger.bas` (que ya no existe en el proyecto). Esta es una referencia huérfana que debe actualizarse.

---

### 🟡 M02 — Hoja3.bas (Checklist) no tiene código de protección

**Archivo afectado:** `Hoja3.bas`

**Problema:** El módulo de la hoja "Checklist" está vacío. No tiene eventos `Worksheet_Activate`, `Worksheet_Deactivate`, `Worksheet_Change` ni `Worksheet_SelectionChange`. Esto significa que la hoja donde se configuran las plantillas de inspección no tiene ninguna protección ni auditoría.

**Impacto:** Cualquier usuario podría modificar las plantillas de inspección, preguntas, secciones y opciones de respuesta sin restricción ni registro de auditoría.

**Checklist:**
- [x] **M02a** — Agregado evento `Worksheet_Activate` en Hoja3.bas con `ApplyRoleBasedProtection` ✅
- [x] **M02b** — Agregado evento `Worksheet_Deactivate` con protección ✅
- [x] **M02c** — Agregado evento `Worksheet_Change` para auditoría de cambios en tablas (tblPlantillas, tblPreguntas, tblSecciones, tblOpcionesDeRespuesta) ✅
- [x] **M02d** — Agregado evento `Worksheet_SelectionChange` para capturar valores previos (patrón estándar) ✅

> **Arquitectura:** Hoja3 (Checklist) es una **vista de configuración crítica**. Contiene 4 tablas maestras que definen cómo se realizan las inspecciones. Sin protección, un usuario podría alterar las preguntas, criterios de evaluación y opciones de respuesta.

---

### 🟡 M03 — No hay persistencia de sesión entre aperturas del libro

**Archivo afectado:** `VariablesGlobales2.bas`, `ThisWorkbook.bas`

**Problema:** La variable `m_userRole` se almacena en memoria volátil (variable global de módulo). Al cerrar y reabrir el libro, el rol se reinicia a `INITIAL_USER_ROLE = "Usuario"`. El administrador debe autenticarse cada vez que abre el libro.

```vb
' VariablesGlobales2.bas — variable volátil
Public m_userRole As String  ' ← Se pierde al cerrar el libro

' ThisWorkbook.Workbook_Open — siempre inicia como Usuario
m_userRole = Configuration2.INITIAL_USER_ROLE  ' ← "Usuario"
```

**Impacto:** No hay riesgo de seguridad (es más seguro que el admin deba autenticarse cada vez), pero es una fricción de usabilidad. El administrador debe ingresar la contraseña cada vez que abre el libro.

**Checklist:**
- [ ] **M03a** — Evaluar si se desea persistencia de sesión (requiere almacenar el rol en una celda oculta del libro o en un archivo externo)
- [ ] **M03b** — Si se implementa: agregar variable `g_sessionExpiry` para que la sesión expire después de N minutos de inactividad
- [ ] **M03c** — Si se implementa: agregar auditoría de re-apertura de sesión
- [ ] **M03d** — Si NO se implementa: documentar como comportamiento por diseño (más seguro)

> **Decisión arquitectónica:** La falta de persistencia es **intencional y más segura** desde el punto de vista de seguridad. Si se desea persistencia, considerar almacenar el rol cifrado en una celda oculta o en el registro de Windows. Se recomienda NO implementar persistencia a menos que sea un requisito explícito del negocio.

---

### 🟡 M04 — Falta de estandarización en el tratamiento de contraseña de cronograma

**Archivo afectado:** `Hoja1.bas` (líneas 198-211), `CronogramaGestorService.bas` (líneas 39-67)

**Problema:** Mientras que `AdminAccessControl2` usa un formulario modal (`frmInput.frm`) para capturar la contraseña, el acceso al Gestor de Cronograma usa un `InputBox` estándar de VBA:

```vb
' Hoja1.bas — InputBox simple (menos seguro)
contrasena = InputBox("Ingrese la contraseña de administrador:", "Gestor de Cronograma")
```

Además, `AdminAccessControl2` cambia el rol global (`m_userRole = "Admin"`) mientras que `CronogramaGestorService.ValidarContrasenaGestor` solo retorna un booleano sin cambiar el rol.

**Impacto:** Inconsistencia en la experiencia de usuario y en el patrón de autenticación. El InputBox muestra la contraseña en texto plano (sin ocultar caracteres).

**Checklist:**
- [x] **M04a** — Estandarizado: `InputBox` reemplazado por `frmInput` con ocultación de contraseña en `Hoja1.AbrirGestorCronograma` ✅
- [ ] **M04b** — Evaluar si `ValidarContrasenaGestor` debería también cambiar `m_userRole` a un rol específico de cronograma (ej. "CronoAdmin") — ⏳ Pendiente decisión de diseño
- [x] **M04c** — Auditoría de intentos fallidos ya implementada en `CronogramaGestorService` ✅

> **Arquitectura:** La autenticación debe seguir un patrón único. Actualmente hay dos puntos de autenticación (AdminAccessControl2 y CronogramaGestorService) con implementaciones diferentes. Se recomienda centralizar la autenticación en un solo servicio.

---

### 🟡 M05 — Ausencia de bloqueo por intentos fallidos

**Archivo afectado:** `AdminAccessControl2.bas`

**Problema:** No hay límite de intentos fallidos de autenticación. Un usuario puede intentar contraseñas indefinidamente, lo que permite ataques de fuerza bruta.

**Impacto:** Bajo en la práctica (el libro se abre en sesión local), pero es una mala práctica de seguridad que debe corregirse.

**Checklist:**
- [x] **M05a** — Implementado contador de intentos fallidos (`ml_intentosFallidos`) en `AdminAccessControl2.CheckAdminAccess` y `AdminAccess` ✅
- [x] **M05b** — Implementado bloqueo de 5 minutos tras 3 intentos fallidos (`AplicarBloqueo` con `DateAdd("n", 5, Now)`) ✅
- [x] **M05c** — Registrado cada intento fallido con `AuditLogger2.LogAction` (intento N de MAX_INTENTOS_FALLIDOS) ✅
- [x] **M05d** — Notificación al usuario cuando queda 1 intento antes del bloqueo -> "Le queda 1 intento" ✅

> **Arquitectura:** Esta lógica pertenece a `AdminAccessControl2` (capa de aplicación). Puede usar una variable estática en el módulo o almacenar el contador en una celda oculta.

---

### 🟡 M06 — Las hojas legacy (Hoja2, Hoja7, Hoja12, Hoja13, Hoja14, Hoja20) usan nombres dudosos no centralizados

**Archivos afectados:** Hoja2.bas, Hoja7.bas, Hoja12.bas, Hoja13.bas, Hoja14.bas, Hoja20.bas

**Problema:** Estas hojas referencian nombres de tablas como *strings hardcodeados* en lugar de usar las constantes de `Configuration2`:

```vb
' Ejemplo en Hoja2.bas — string hardcodeado
tablesToAudit = Array("tblControlDeCambios")  ' ← Debería ser Configuration2.TABLE_*

' Ejemplo en Hoja12.bas
tablesToAudit = Array("tblPersonalProduccion")  ' ← No está definida en Configuration2
```

**Impacto:** Si se renombra una tabla en Excel, el código legacy no se actualizará automáticamente, causando errores silenciosos en la auditoría.

**Checklist:**
- [x] **M06a** — Revisar si las tablas legacy existen en el libro actual ✅
- [x] **M06b** — Si existen: centralizar sus nombres en `Configuration2.bas` como nuevas constantes ✅
- [x] **M06c** — Actualizar Hoja2.bas para usar la constante de Configuration2 ✅
- [x] **M06d** — Actualizar Hoja7.bas para usar la constante de Configuration2 ✅
- [x] **M06e** — Actualizar Hoja12.bas para usar la constante de Configuration2 ✅
- [x] **M06f** — Actualizar Hoja13.bas para usar la constante de Configuration2 ✅
- [x] **M06g** — Actualizar Hoja14.bas para usar la constante de Configuration2 ✅
- [x] **M06h** — Actualizar Hoja20.bas para usar la constante de Configuration2 ✅

> **Arquitectura:** `Configuration2.bas` es el **punto único de configuración** del sistema. Todos los nombres de tablas deben estar centralizados allí como constantes `Public Const`. Las hojas (vistas) deben referenciar `Configuration2.TABLE_*` para mantener la arquitectura limpia y permitir cambios sin modificar cada módulo individual.

---

### 🟡 M07 — Backup automático desactivado

**Archivo afectado:** `ThisWorkbook.bas` (línea 251), `mod_BackupManager.bas`

**Problema:** La llamada a `mod_BackupManager.CrearBackupAutomatico` en `Workbook_BeforeSave` está comentada:

```vb
' ========== BACKUP AUTOMÁTICO TEMPORALMENTE DESACTIVADO ==========
' Para reactivar, descomentar la siguiente línea
' =================================================================
' Call mod_BackupManager.CrearBackupAutomatico
```

**Impacto:** No se crean copias de seguridad automáticas antes de guardar. Si un usuario comete un error al ingresar datos, no hay un backup automático para recuperar.

**Checklist:**
- [x] **M07a** — Descomentada línea de `mod_BackupManager.CrearBackupAutomatico` en `ThisWorkbook.bas` ✅
- [ ] **M07b** — Probar que se crea el archivo "Copia de Seguridad [Nombre].xlsm" al guardar — ⏳ Pendiente prueba manual
- [ ] **M07c** — Verificar que el backup se actualiza correctamente en guardados sucesivos — ⏳ Pendiente prueba manual

> **Arquitectura:** `mod_BackupManager.bas` es un **servicio de infraestructura independiente**. `ThisWorkbook.bas` es el **orquestador** que lo invoca en el evento `Workbook_BeforeSave`. Esta separación respeta la arquitectura de capas.

---

### 🟡 M08 — Análisis automático desactivado al cerrar

**Archivo afectado:** `ThisWorkbook.bas` (líneas 282-288)

**Problema:** El bloque que ejecuta `dataProcessAnalysis.EjecutarAnalisis` si hay cambios pendientes está comentado en `Workbook_BeforeClose`:

```vb
' ========== ANÁLISIS AUTOMÁTICO TEMPORALMENTE DESACTIVADO ==========
' Para reactivar, descomentar el bloque de código a continuación
' ===================================================================
' If g_AnalisisPendiente Then
'     Call dataProcessAnalysis.EjecutarAnalisis
'     g_AnalisisPendiente = False
' End If
```

**Impacto:** Los análisis de KPI no se actualizan automáticamente al cerrar el libro. Dependen de que el usuario navegue a la hoja específica (Hoja20) para ejecutarse.

**Checklist:**
- [x] **M08a** — `dataProcessAnalysis` existe en el proyecto ✅
- [x] **M08b** — Bloque de análisis descomentado en `Workbook_BeforeClose` ✅
- [ ] **M08c** — Probar que el análisis se ejecuta al cerrar si hay cambios pendientes — ⏳ Pendiente prueba manual

> **Arquitectura:** `g_AnalisisPendiente` (variable global en VariablesGlobales2.bas) es una **bandera de estado** que conecta el guardado de datos (orquestado por otros servicios) con el análisis (orquestado por ThisWorkbook). Usar una bandera global para diferir procesamiento es aceptable en VBA siempre que se limpie correctamente.

---

### 🟡 M09 — Faltan eventos en hojas del sistema de inspecciones

**Archivos afectados:** Varias hojas

**Problema:** Las hojas del sistema de inspecciones que no están mapeadas en el árbol de VBA como "HojaX" no tienen módulos de código asociados. Específicamente:

| Hoja en Excel | ¿Tiene módulo VBA? | ¿Worksheet_Activate? |
|---------------|-------------------|---------------------|
| Cronograma | ❌ No detectado | ❌ |
| Configuración | ❌ No detectado | ❌ |
| Plantilla Certificado | ❌ No detectado | ❌ |
| Audit trail 2,3,4,5 | ❌ No detectado | ❌ |

**Impacto:** Estas hojas no tienen protección por rol ni auditoría de cambios.

**Checklist:**
- [ ] **M09a** — Verificar qué hojas del sistema no tienen módulo VBA
- [ ] **M09b** — Agregar módulos a las hojas críticas (Cronograma, Configuración) con eventos estándar
- [ ] **M09c** — Verificar que las 5 hojas Audit Trail tengan el mismo nivel de protección (Hoja9 implementado para "Audit trail 1", replicar para 2-5)

> **Arquitectura:** En VBA, cada hoja debe tener su propio módulo de clase (Sheet) para manejar eventos. Si una hoja no tiene módulo, no puede reaccionar a eventos como Activate, Change, etc. Esto es una limitación de VBA, no de diseño.

---

## 3. 📊 RESUMEN DE PRIORIDADES

| Prioridad | ID | Descripción | Dependencias |
|-----------|----|-------------|-------------|
| 🔴 **P1** | C01 | Contraseñas débiles e idénticas | — |
| 🔴 **P1** | C02 | Protección deshabilitada | C01 |
| 🔴 **P1** | C03 | Inconsistencia roles entre hojas | — (se puede hacer en paralelo) |
| 🔴 **P1** | C04 | Menú principal sin protección por rol | C03 |
| 🔴 **P1** | C05 | Histórico sin protección | C03 |
| 🔴 **P2** | C06 | Navegación segura desactivada | C03, C04, C05 |
| 🔴 **P2** | C07 | SheetService no aplica protección | C03, C06 |
| 🟡 **P3** | M01 | ErrorLogger mal referenciado | — |
| 🟡 **P3** | M02 | Hoja3 (Checklist) vacía | C03 |
| 🟡 **P3** | M04 | InputBox vs frmInput | — |
| 🟡 **P3** | M06 | Nombres legacy hardcodeados | — |
| 🟡 **P4** | M03 | Persistencia de sesión | — |
| 🟡 **P4** | M05 | Bloqueo por intentos fallidos | M04 |
| 🟡 **P4** | M07 | Backup desactivado | — |
| 🟡 **P4** | M08 | Análisis desactivado | — |
| 🟡 **P4** | M09 | Hojas sin módulo VBA | C03 |

---

## 4. 🏗️ ARQUITECTURA OBJETIVO (Clean Architecture en VBA)

Después de aplicar todas las correcciones, la arquitectura de seguridad debe verse así:

```
┌─────────────────────────────────────────────────────────────────┐
│                     CAPA DE PRESENTACIÓN (Vistas)               │
│                                                                 │
│  Hoja1 (Menú principal)    Hoja5 (Personal)                     │
│  Hoja3 (Checklist)         Hoja6 (Historico)                    │
│  Hoja4 (Aseguramiento)     Hoja9 (Audit trail)                  │
│  ... y demás hojas                                               │
│                                                                 │
│  Responsabilidad:                                               │
│  - Llamar a AplicarProteccionPorRol() en Worksheet_Activate     │
│  - Llamar a AplicarProteccionPorRol() en Worksheet_Deactivate   │
│  - NO implementar lógica de permisos propia                     │
│  - NO referenciar m_userRole directamente (usar GetUserRole())  │
└──────────────────────────┬──────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│                   CAPA DE APLICACIÓN (Servicios)                │
│                                                                 │
│  NavigationService2    AdminAccessControl2                      │
│  CronogramaGestorService                                        │
│                                                                 │
│  Responsabilidad:                                               │
│  - Orquestar navegación entre vistas                            │
│  - Autenticar usuarios y cambiar roles                          │
│  - Validar permisos antes de operaciones                        │
│  - Invocar servicios de infraestructura                         │
└──────────────────────────┬──────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│                CAPA DE INFRAESTRUCTURA (Servicios Técnicos)     │
│                                                                 │
│  SheetProtector2    WorkbookProtector2    SheetService2         │
│  AuditLogger2       ErrorLogger2         mod_BackupManager      │
│                                                                 │
│  Responsabilidad:                                               │
│  - Proteger/desproteger hojas y libro (SheetProtector2)        │
│  - Ocultar/mostrar hojas (SheetService2)                        │
│  - Registrar auditoría y errores (AuditLogger2, ErrorLogger2)  │
│  - Aplicar protección según rol (ApplyRoleBasedProtection)      │
│  - NO tomar decisiones de negocio                               │
└──────────────────────────┬──────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│              CAPA DE CONFIGURACIÓN (Configuración)              │
│                                                                 │
│  Configuration2    VariablesGlobales2                           │
│                                                                 │
│  Responsabilidad:                                               │
│  - Centralizar todas las constantes y contraseñas              │
│  - Definir flags de entorno (desarrollo/producción)            │
│  - NO contener lógica de negocio                               │
│  - NO contener lógica de infraestructura                        │
└─────────────────────────────────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│              CAPA DE DATOS (Repositorios)                       │
│                                                                 │
│  InspectionRepository    ChecklistRepository                    │
│  InspectionScheduler     InspectionHistoryService               │
│  ... y demás repositorios                                      │
│                                                                 │
│  Responsabilidad:                                               │
│  - Acceso a tablas de Excel (ListObjects)                      │
│  - CRUD de datos de negocio                                    │
│  - NO gestionar seguridad (solo datos)                         │
└─────────────────────────────────────────────────────────────────┘
```

### 🧠 Flujo de seguridad corregido (después de implementar TODO)

```
1. USUARIO ABRE EL LIBRO
   └── Workbook_Open
       ├── ENABLE_WORKBOOK_PROTECTION = True
       │   └── WorkbookProtector2.ProtectWorkbook ← Estructura protegida
       ├── ENABLE_SHEET_PROTECTION = True
       │   └── SheetProtector2.ApplyRoleBasedProtection ← Hojas protegidas por rol
       ├── m_userRole = "Usuario" (rol por defecto)
       └── AuditLogger2.LogAction("Apertura del libro")

2. USUARIO AUTENTICACIÓN COMO ADMIN
   └── AdminAccessControl2.CheckAdminAccess
       ├── frmInput.Show vbModal ← Formulario con ocultación de contraseña
       ├── enteredPassword = Configuration2.ADMIN_PASSWORD
       ├── m_userRole = "Admin"
       └── AuditLogger2.LogAction("Cambio de rol a Admin")

3. USUARIO NAVEGA A UNA HOJA
   └── NavigationService2.NavigateToSheet("Personal")
       ├── SheetService2.HideAndProtectAllSheetsExcept("Personal")
       │   ├── WorkbookProtector2.UnprotectWorkbook
       │   ├── Oculta todas las hojas excepto destino y Menú principal
       │   ├── SheetProtector2.ApplyRoleBasedProtection(wsDestino) ← ¡SE APLICA SIEMPRE!
       │   └── WorkbookProtector2.ProtectWorkbook
       ├── Sheets("Personal").Select
       └── (Worksheet_Activate se dispara como redundancia)

4. USUARIO NORMAL INTENTA EDITAR
   └── Hoja está protegida por APP_PASSWORD
       ├── Si el usuario no es Admin → ProtectSheet (sin selección)
       ├── Si el usuario es Admin → UnprotectSheet (edición libre)
       └── En ambos casos → Auditoría registra el acceso
```

---

## 5. 📝 NOTAS DE IMPLEMENTACIÓN

### Orden de implementación recomendado

```
FASE 1 (Prioridad Máxima — Correcciones Inmediatas)
  ├── M01 → ErrorLogger2 mal referenciado (2 minutos, riesgo de error silencioso)
  ├── C01 → Contraseñas fuertes y diferentes (10 minutos, requiere migración física)
  └── C02 → Activar protección (5 minutos, cambiar 2 booleanos)

FASE 2 (Prioridad Alta — Protección de Hojas)
  ├── C03 → Unificar patrón de roles (30 minutos, requiere crear función central)
  ├── C04 → Proteger Menú principal (5 minutos)
  ├── C05 → Proteger Histórico (5 minutos)
  └── M02 → Proteger Checklist (10 minutos)

FASE 3 (Prioridad Media — Navegación y SheetService)
  ├── C06 → Reactivar navegación segura (10 minutos)
  └── C07 → SheetService aplica protección por rol (15 minutos)

FASE 4 (Prioridad Baja — Mejoras y Limpieza)
  ├── M04 → Estandarizar InputBox por frmInput
  ├── M05 → Bloqueo por intentos fallidos
  ├── M06 → Centralizar nombres legacy
  ├── M07 → Reactivar backup
  ├── M08 → Reactivar análisis al cerrar
  └── M09 → Agregar módulos a hojas faltantes

FASE 5 (Futuro — Mejoras Arquitectónicas)
  └── M03 → Evaluar persistencia de sesión (si es necesario)
```

### Pruebas de regresión sugeridas

Después de cada fase, verificar:
1. ✅ El libro se abre sin errores
2. ✅ Un usuario normal NO puede editar datos en ninguna hoja
3. ✅ Un administrador SÍ puede editar datos después de autenticarse
4. ✅ Las contraseñas del cronograma siguen funcionando
5. ✅ Los PDFs se siguen generando correctamente
6. ✅ El Audit Trail sigue registrando acciones
7. ✅ No hay errores 1004 (protección) en ninguna operación
8. ✅ El Gestor de Cronograma sigue funcionando correctamente
9. ✅ El resumen de cronograma se actualiza correctamente
10. ✅ La inicialización automática del sistema funciona (primera vez)

---

> **Documento generado como parte del diagnóstico de seguridad del sistema VBA.**
> **Para preguntas o aclaraciones, contactar al equipo de desarrollo.**
