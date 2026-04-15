# 🔐 SISTEMA DE PROTECCIÓN - GUÍA COMPLETA

**Proyecto:** TH-HC-001 EN PROCESO DE VALIDACIÓN  
**Versión:** 1.0  
**Fecha:** 14/04/2026  
**Última actualización:** 14/04/2026

---

## 📖 Tabla de Contenidos

1. [Introducción](#introducción)
2. [Estructura General](#estructura-general)
3. [Booleanos de Control](#booleanos-de-control)
4. [Paso a Paso: Cambiar Modo](#paso-a-paso-cambiar-modo)
5. [Sistema de Protección del Libro (URS-22)](#sistema-de-protección-del-libro-urs-22)
6. [Sistema de Protección de Hojas (URS-20/21)](#sistema-de-protección-de-hojas-urs-2021)
7. [Matriz de Comportamientos](#matriz-de-comportamientos)
8. [Debugging y Verificación](#debugging-y-verificación)
9. [Troubleshooting](#troubleshooting)

---

## 🎯 Introducción

El sistema de protección del libro TH-HC-001 tiene **DOS CAPAS** independientes:

1. **📌 URS-22: Protección de Estructura del Libro**
   - Evita que se eliminen/renombren/muevan hojas
   - Controlada por: `ENABLE_WORKBOOK_PROTECTION`

2. **🔒 URS-20/21: Protección de Hojas Individuales**
   - Evita que se editen celdas según rol de usuario
   - Controlada por: `ENABLE_SHEET_PROTECTION`

Ambas se pueden **activar/desactivar independientemente** mediante booleanos en `Configuration2.bas`.

---

## 🏗️ Estructura General

```
┌─────────────────────────────────────────────────────────┐
│         Configuration2.bas                              │
│  • ENABLE_WORKBOOK_PROTECTION (True/False)             │
│  • ENABLE_SHEET_PROTECTION (True/False)                │
│  • APP_PASSWORD, AUDIT_PASSWORD, etc.                  │
└────┬──────────────────────────────┬────────────────────┘
     │                              │
     ↓                              ↓
┌────────────────────┐      ┌──────────────────────┐
│ WorkbookProtector2 │      │  SheetProtector2     │
│ • ProtectWorkbook()│      │ • ProtectSheet()     │
│ • UnprotectBook() │      │ • UnprotectSheet()   │
└────────────────────┘      │ • ProtectSheetFor   │
         ↑                  │   Reading()          │
         │                  │ • ApplyRoleBased    │
         │                  │   Protection()      │
     Llamado por:           └──────────────────────┘
   ThisWorkbook.Open()              ↑
   AuditRotation2.bas          Llamado por:
   SheetService2.bas         Hoja events
   Otros módulos            AuditLogger2.bas
                            Otros módulos
```

---

## 🎮 Booleanos de Control

### Ubicación

**Archivo:** `Configuration2.bas`  
**Líneas:** 100-130 (aproximadamente)

### Booleano 1: ENABLE_WORKBOOK_PROTECTION

```vba
' ============================================================================
' CONFIGURACIÓN DE DESARROLLO Y PROTECCIÓN
' ============================================================================

' Constante: ENABLE_WORKBOOK_PROTECTION
' Propósito: Activa/desactiva la protección de estructura del libro (URS-22)
' Valores:
'   ? True  = PRODUCCIÓN: Estructura protegida
'   ? False = DESARROLLO: Puedes cambiar hojas libremente
' Nota: Cambiar SOLO este valor para debugging
' ============================================================================
Public Const ENABLE_WORKBOOK_PROTECTION As Boolean = False  ' DESARROLLO: False | PRODUCCIÓN: True
```

**Efecto cuando es `False` (Actual):**
- ✅ Puedes eliminar hojas
- ✅ Puedes renombrar hojas
- ✅ Puedes mover hojas
- ✅ Puedes copiar hojas
- ✅ Puedes ocultar/mostrar hojas

**Efecto cuando es `True`:**
- ❌ NO puedes eliminar hojas
- ❌ NO puedes renombrar hojas
- ❌ NO puedes mover hojas
- ❌ NO puedes copiar hojas
- ❌ NO puedes ocultar/mostrar hojas (solo mediante código administrativo)

---

### Booleano 2: ENABLE_SHEET_PROTECTION

```vba
' Constante: ENABLE_SHEET_PROTECTION
' Propósito: Activa/desactiva la protección individual de hojas (URS-20/21)
' Valores:
'   ? True  = PRODUCCIÓN: Hojas protegidas según rol
'   ? False = DESARROLLO: Edición libre para testing
' Nota: Relacionado con SheetProtector2
' ============================================================================
Public Const ENABLE_SHEET_PROTECTION As Boolean = False  ' DESARROLLO: False | PRODUCCIÓN: True
```

**Efecto cuando es `False` (Actual):**
- ✅ Puedes editar CUALQUIER celda
- ✅ Puedes insertar/eliminar filas
- ✅ Puedes cambiar formato
- ✅ Modo de desarrollo completo (máxima libertad)

**Efecto cuando es `True`:**
- ❌ No puedes editar la mayoría de celdas
- ❌ No puedes insertar/eliminar filas
- ❌ Se aplican restricciones por rol:
  - **Admin:** No protegidas (total libertad)
  - **Usuario:** Solo lectura + copiar (no editar)
  - **Otros:** Completamente bloqueados

---

## 📋 Paso a Paso: Cambiar Modo

### ⚠️ IMPORTANTE: Booleanos Bidireccionales (14/04/2026)

Los booleanos ahora son **completamente bidireccionales**:

- ✅ Cuando `ENABLE_WORKBOOK_PROTECTION = True` → Ejecuta `ProtectWorkbook()`
- ✅ Cuando `ENABLE_WORKBOOK_PROTECTION = False` → Ejecuta `UnprotectWorkbook()` (NUEVO)
- ✅ Cuando `ENABLE_SHEET_PROTECTION = False` → Ejecuta `DesprotegerTodasLasHojas()` (NUEVO)

**Flujo en Workbook_Open():**

```vba
' Control de protección de estructura (URS-22)
If Configuration2.ENABLE_WORKBOOK_PROTECTION Then
    Call WorkbookProtector2.ProtectWorkbook()        ' ← Si True
Else
    Call WorkbookProtector2.UnprotectWorkbook()      ' ← Si False (NUEVO 14/04/2026)
End If

' Control de protección de hojas (URS-20/21)
If Not Configuration2.ENABLE_SHEET_PROTECTION Then
    Call DesprotegerTodasLasHojas()                 ' ← Si False (NUEVO 14/04/2026)
End If
```

**Ventaja:** 
Cambias el booleano, cierras/abres el libro → ¡SE APLICA INMEDIATAMENTE! 
No necesitas ningún código manual.

---

## 📋 Paso a Paso: Cambiar Modo

### Escenario 1: Pasar de DESARROLLO a TESTING (Solo protección de estructura)

**Objetivo:** Probar que no se pueden eliminar hojas, pero sí editar celdas.

#### Paso 1: Abrir Configuration2.bas
- **VBA Editor:** Alt+F11
- **Navegador de Proyectos:** Panel izquierdo
- **Buscar:** "Configuration2"
- **Doble clic** para abrir

#### Paso 2: Localizar los booleanos
Presiona **Ctrl+F** y busca:
```
ENABLE_WORKBOOK_PROTECTION
```

Deberías ver:
```vba
Line 108: Public Const ENABLE_WORKBOOK_PROTECTION As Boolean = False
Line 117: Public Const ENABLE_SHEET_PROTECTION As Boolean = False
```

#### Paso 3: Cambiar ENABLE_WORKBOOK_PROTECTION de False a True

**Antes:**
```vba
Public Const ENABLE_WORKBOOK_PROTECTION As Boolean = False
```

**Después:**
```vba
Public Const ENABLE_WORKBOOK_PROTECTION As Boolean = True
```

#### Paso 4: Guardar archivo
- Presiona **Ctrl+S**
- O: **File → Save**

#### Paso 5: Cierra y reabre el libro
1. **Cierra el libro** (File → Close, o Alt+F4)
2. **Excel sigue abierto**
3. **Abre nuevamente** el archivo `TH-HC-001 INSPECCIONES.xlsm`
4. Excel ejecutará `Workbook_Open()` con la nueva configuración

#### Paso 6: Verificar
- **Haz clic derecho** en cualquier pestaña
- Las opciones deben estar **GRISES/DESHABILITADAS:**
  - ❌ Eliminar
  - ❌ Cambiar nombre
  - ❌ Mover o copiar

**Resultado:** ✅ URS-22 Activada

---

### Escenario 2: Pasar a PRODUCCIÓN (Protección completa)

**Objetivo:** Activar TODAS las protecciones.

#### Paso 1-2: Abrir Configuration2 y buscar booleanos (igual al Escenario 1)

#### Paso 3: Cambiar AMBOS booleanos a True

**Antes:**
```vba
Public Const ENABLE_WORKBOOK_PROTECTION As Boolean = False
Public Const ENABLE_SHEET_PROTECTION As Boolean = False
```

**Después:**
```vba
Public Const ENABLE_WORKBOOK_PROTECTION As Boolean = True
Public Const ENABLE_SHEET_PROTECTION As Boolean = True
```

#### Paso 4: Guardar (Ctrl+S)

#### Paso 5: Cierra y reabre el libro

#### Paso 6: Verificar
- **Haz clic derecho** en pestaña:
  - ❌ Opciones deshabilitadas (URS-22 activa)

- **Intenta editar una celda de "Configuración":**
  - ❌ No puedes editar (URS-20/21 activa, según rol)

**Resultado:** ✅ Sistema completamente protegido

---

### Escenario 3: Volver a DESARROLLO

**Objetivo:** Reactivar libertad total (desarrollo).

#### Paso 1-2: Abrir Configuration2 (igual)

#### Paso 3: Cambiar AMBOS a False

```vba
Public Const ENABLE_WORKBOOK_PROTECTION As Boolean = False
Public Const ENABLE_SHEET_PROTECTION As Boolean = False
```

#### Paso 4: Guardar + Cierra/Reabre

#### Paso 6: Verificar
- **Haz clic derecho en pestaña:**
  - ✅ Opciones HABILITADAS (URS-22 desactivada)

- **Intenta editar cualquier celda:**
  - ✅ Puedes editar (URS-20/21 desactivada)

**Resultado:** ✅ Libertad total para desarrollo

---

## 🛠️ Sistema de Protección del Libro (URS-22)

### ¿Qué es?

Sistema que **protege la estructura global del libro** para evitar que se eliminen, renombren o muevan las hojas de manera accidental o maliciosa.

### ¿Por qué es importante?

```
Eliminación accidental de "Configuración"
    ↓
Sistema pierde capacidad de cálculo RPN
    ↓
Reporte mensual no se puede generar
    ↓
❌ PÉRDIDA DE DISPONIBILIDAD
❌ INCUMPLIMIENTO REGULATORIO
```

### Módulo Responsable: WorkbookProtector2.bas

**Ubicación:**  
```
Raíz del Proyecto → WorkbookProtector2.bas
```

**Funciones:**

#### 1. ProtectWorkbook()

```vba
Public Sub ProtectWorkbook()
    On Error Resume Next
    ThisWorkbook.Protect Password:=Configuration2.APP_PASSWORD, Structure:=True, Windows:=False
    If VBA.Err.Number <> 0 Then
        Call ErrorLogger2.Log("WorkbookProtector2.ProtectWorkbook", VBA.Err.Description, VBA.Err.Number)
    End If
    On Error GoTo 0
End Sub
```

**¿Qué hace?**
1. Protege la estructura del libro con contraseña "1234"
2. `Structure:=True` → Bloquea cambios de hojas
3. `Windows:=False` → Permite mover/redimensionar ventanas
4. Si ya está protegido → Ignora error silenciosamente

**¿Quién lo llama?**
- `ThisWorkbook.Workbook_Open()` → Al abrir el libro
- `AuditRotation2.bas` → Después de tareas administrativas
- `SheetService2.bas` → Después de cambios de visibilidad

#### 2. UnprotectWorkbook()

```vba
Public Sub UnprotectWorkbook()
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=Configuration2.APP_PASSWORD
    If VBA.Err.Number <> 0 Then
        Call ErrorLogger2.Log("WorkbookProtector2.UnprotectWorkbook", VBA.Err.Description, VBA.Err.Number)
    End If
    On Error GoTo 0
End Sub
```

**¿Qué hace?**
1. Desprotege temporalmente el libro
2. Permite cambios administrativos
3. Necesario antes de operaciones como:
   - Crear nuevas hojas
   - Cambiar nombres de hojas
   - Mover hojas

**Patrón de uso:**
```vba
' Desproteger para tareas administrativas
Call WorkbookProtector2.UnprotectWorkbook()

' Realizar cambios...
ThisWorkbook.Sheets.Add

' Volver a proteger
Call WorkbookProtector2.ProtectWorkbook()
```

### Flujo de URS-22 en Tiempo de Ejecución

```
1. Usuario abre libro TH-HC-001 INSPECCIONES.xlsm
   ↓
2. Excel carga todas las hojas
   ↓
3. Evento: Workbook_Open() se dispara
   ↓
4. Código en ThisWorkbook.bas:
   If Configuration2.ENABLE_WORKBOOK_PROTECTION Then
       Call WorkbookProtector2.ProtectWorkbook()
   End If
   ↓
5. WorkbookProtector2.ProtectWorkbook() ejecuta:
   ThisWorkbook.Protect(Password:="1234", Structure:=True, Windows:=False)
   ↓
6. Excel internamente: ThisWorkbook.ProtectStructure = True
   ↓
7. Usuario intenta hacer clic derecho en pestaña
   ↓
8. Excel verifica: ¿ProtectStructure = True?
   ↓
9. Sí → Deshabilita opciones de "Eliminar", "Cambiar nombre", etc.
   ↓
10. Usuario ve opciones GRISES (disabled)
   ↓
11. Si intenta hacer clic → Error: "No se puede realizar esta acción"
```

### Estados Posibles

| Situación | ENABLE_WORKBOOK_PROTECTION | Estado | Resultado |
|-----------|---------------------------|--------|-----------|
| Desarrollo | FALSE | 🔓 Desprotegido | Puedes cambiar hojas libremente |
| Testing Estructura | TRUE | 🔒 Protegido | No puedes cambiar hojas |
| Producción | TRUE | 🔒 Protegido | Protección máxima |

---

## 🔒 Sistema de Protección de Hojas (URS-20/21)

### ¿Qué es?

Sistema que **protege celdas individuales dentro de hojas** según el rol del usuario y el nivel de restricción requerido.

### Niveles de Protección

#### Nivel 1: ProtectSheet (Máxima restricción)

```vba
Public Sub ProtectSheet(ByRef ws As Worksheet, ByVal sheetPassword As String)
    If Not Configuration2.ENABLE_SHEET_PROTECTION Then Exit Sub
    
    ws.Protect Password:=sheetPassword, _
        AllowFormattingCells:=False, _
        AllowFormattingColumns:=False, _
        AllowFormattingRows:=False, _
        AllowInsertingRows:=False, _
        AllowInsertingColumns:=False, _
        AllowInsertingHyperlinks:=False, _
        AllowDeletingRows:=False, _
        AllowDeletingColumns:=False, _
        AllowSorting:=False, _
        AllowFiltering:=True, _
        AllowUsingPivotTables:=False
    ws.EnableSelection = xlNoSelection  ' NO puedes ni seleccionar celdas
End Sub
```

**Restricciones:**
- ❌ No puedes editar celdas
- ❌ No puedes insertar filas
- ❌ No puedes eliminar filas
- ❌ No puedes cambiar formato
- ❌ No puedes seleccionar celdas
- ✅ Sí puedes filtrar datos

**Uso:**
- Hojas críticas de "Configuración"
- Datos maestros que no deben ser modificados

---

#### Nivel 2: ProtectSheetForReading (Restricción media)

```vba
Public Sub ProtectSheetForReading(ByRef ws As Worksheet, ByVal sheetPassword As String)
    If Not Configuration2.ENABLE_SHEET_PROTECTION Then Exit Sub
    
    ws.Protect Password:=sheetPassword, _
        AllowFormattingCells:=False, _
        AllowInsertingRows:=False, _
        AllowDeletingRows:=False, _
        AllowSorting:=False, _
        AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions  ' SÍ puedes seleccionar para copiar
End Sub
```

**Restricciones:**
- ❌ No puedes editar celdas
- ❌ No puedes insertar filas
- ❌ No puedes eliminar filas
- ❌ No puedes cambiar formato
- ✅ SÍ puedes seleccionar celdas
- ✅ SÍ puedes copiar datos (Ctrl+C)
- ✅ SÍ puedes filtrar

**Uso:**
- Hojas de consulta
- Historiadores de datos
- Reportes de solo lectura

---

#### Nivel 3: Sin protección (Libertad completa)

```vba
' Cuando ENABLE_SHEET_PROTECTION = False, o
' cuando el rol es "Admin"
ws.Unprotect Password:=Configuration2.APP_PASSWORD
```

**Permisos:**
- ✅ Puedes editar cualquier celda
- ✅ Puedes insertar/eliminar filas
- ✅ Puedes cambiar formato
- ✅ Puedes crear gráficos

**Uso:**
- Administradores
- Hojas de edición en desarrollo

---

### Módulo Responsable: SheetProtector2.bas

**Ubicación:**
```
Raíz del Proyecto → SheetProtector2.bas
```

### Flujo de Protección de Hojas

```
1. Hoja se activa (Worksheet_Activate)
   ↓
2. Código llama a: ApplyRoleBasedProtection(ws, sheetPassword)
   ↓
3. ApplyRoleBasedProtection verifica:
   If Configuration2.ENABLE_SHEET_PROTECTION = False Then
       Exit Sub (no hacer nada)
   ↓
4. Si ENABLE_SHEET_PROTECTION = True:
   Desproteger primero: UnprotectSheet(ws, sheetPassword)
   ↓
5. Aplicar protección según rol:
   - Si m_userRole = "Admin" → No proteger (libertad)
   - Si m_userRole = "Usuario" → ProtectSheetForReading (solo lectura)
   - Si otro rol → ProtectSheet (máxima restricción)
   ↓
6. Usuario intenta editar celda
   ↓
7. Si celda está protegida:
   Excel: "La celda está protegida porque la hoja está protegida"
   ↓
8. Usuario ve error y no puede editar
```

### Matriz de Restricciones por Rol

| Acción | Admin | Usuario | Otros |
|--------|-------|---------|-------|
| Editar celdas | ✅ | ❌ | ❌ |
| Copiar datos | ✅ | ✅ | ❌ |
| Seleccionar celdas | ✅ | ✅ | ❌ |
| Insertar filas | ✅ | ❌ | ❌ |
| Eliminar filas | ✅ | ❌ | ❌ |
| Cambiar formato | ✅ | ❌ | ❌ |
| Filtrar datos | ✅ | ✅ | ✅ |

---

## 📊 Matriz de Comportamientos

### Comportamiento Completo

| Escenario | ENABLE_WORKBOOK | ENABLE_SHEET | Puedes eliminar hojas | Puedes editar celdas | Estado |
|-----------|-----------------|---------------|----------------------|----------------------|--------|
| 1. Desarrollo | FALSE | FALSE | ✅ Sí | ✅ Sí | 🟢 DESARROLLO |
| 2. Test Estructura | TRUE | FALSE | ❌ No | ✅ Sí | 🟡 TESTING URS-22 |
| 3. Test Hojas | FALSE | TRUE | ✅ Sí | ❌ No* | 🟡 TESTING URS-20/21 |
| 4. Producción | TRUE | TRUE | ❌ No | ❌ No* | 🔴 PRODUCCIÓN |

*Depende del rol del usuario

---

## 🧪 Debugging y Verificación

### Verificación Visual

#### Botón 1: Protección de Estructura

**Test:**
1. Haz clic derecho en una pestaña
2. Mira si la opción "Eliminar" está:
   - 🟢 **HABILITADA** (negro) → ENABLE_WORKBOOK_PROTECTION = False
   - 🔴 **DESHABILITADA** (gris) → ENABLE_WORKBOOK_PROTECTION = True

#### Botón 2: Protección de Hojas

**Test:**
1. Intenta editar una celda en "Configuración"
2. ¿Puedes?
   - 🟢 **SÍ** → ENABLE_SHEET_PROTECTION = False
   - 🔴 **NO** → ENABLE_SHEET_PROTECTION = True (y eres rol "Usuario" o inferior)

### Verificación en VBA

#### Script 1: Verificar Protección de Estructura

```vba
Sub VerificarProteccionEstructura()
    If ThisWorkbook.ProtectStructure Then
        MsgBox "✓ Estructura PROTEGIDA (URS-22 ACTIVA)", vbInformation
        Debug.Print "ThisWorkbook.ProtectStructure = True"
    Else
        MsgBox "✗ Estructura SIN PROTEGER (URS-22 INACTIVA)", vbCritical
        Debug.Print "ThisWorkbook.ProtectStructure = False"
    End If
End Sub
```

**Ejecución:** Alt+F11 → Ctrl+G → Pegar → Enter

---

#### Script 2: Verificar Estado de Booleanos

```vba
Sub VerificarBooleanos()
    Debug.Print "========== ESTADO DE PROTECCIONES =========="
    Debug.Print "ENABLE_WORKBOOK_PROTECTION = " & Configuration2.ENABLE_WORKBOOK_PROTECTION
    Debug.Print "ENABLE_SHEET_PROTECTION = " & Configuration2.ENABLE_SHEET_PROTECTION
    Debug.Print ""
    Debug.Print "Interpretación:"
    
    If Configuration2.ENABLE_WORKBOOK_PROTECTION Then
        Debug.Print "✓ Protección de estructura: ACTIVA (URS-22)"
    Else
        Debug.Print "✗ Protección de estructura: INACTIVA (modo desarrollo)"
    End If
    
    If Configuration2.ENABLE_SHEET_PROTECTION Then
        Debug.Print "✓ Protección de hojas: ACTIVA (URS-20/21)"
    Else
        Debug.Print "✗ Protección de hojas: INACTIVA (modo desarrollo)"
    End If
    Debug.Print "=========================================="
End Sub
```

**Output esperado (Desarrollo):**
```
========== ESTADO DE PROTECCIONES ==========
ENABLE_WORKBOOK_PROTECTION = False
ENABLE_SHEET_PROTECTION = False

Interpretación:
✗ Protección de estructura: INACTIVA (modo desarrollo)
✗ Protección de hojas: INACTIVA (modo desarrollo)
==========================================
```

---

#### Script 3: Verificar Hojas de Auditoría

```vba
Sub VerificarProteccionAudit()
    Dim i As Long
    Dim nombreHoja As String
    Dim ws As Worksheet
    
    For i = 1 To 5
        nombreHoja = "Audit trail " & i
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(nombreHoja)
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            If ws.ProtectContents Then
                Debug.Print "✓ " & nombreHoja & " PROTEGIDA"
            Else
                Debug.Print "✗ " & nombreHoja & " SIN PROTEGER"
            End If
        Else
            Debug.Print "- " & nombreHoja & " NO EXISTE"
        End If
    Next i
End Sub
```

---

## 🆘 Troubleshooting

### Problema 1: "Cambié el booleano pero sigue sin funcionar"

**Causa:** Excel cacheó el módulo

**Solución:**
1. Guarda el archivo (Ctrl+S)
2. **Cierra COMPLETAMENTE** el libro
3. Abre nuevamente
4. Espera a que se ejecute `Workbook_Open()`

---

### Problema 2: "Veo el error pero no puedo cambiar el booleano"

**Causa:** Archivo abierto en otra instancia

**Solución:**
1. Cierra **TODAS** las instancias de Excel
2. Espera 5 segundos
3. Abre nuevamente el archivo

---

### Problema 3: "Booleano está en True pero puedo editar celdas"

**Causa:** Tu rol es "Admin", no tiene restricciones

**Solución:**
1. Cambia tu rol a "Usuario":
   ```vba
   ThisWorkbook.GetUserRole()  ' Debe retornar "Usuario"
   ```
2. O verifica que `ApplyRoleBasedProtection` esté siendo llamada

---

### Problema 4: "Error: 'ThisWorkbook.Protect' failed"

**Causa:** Contraseña incorrecta

**Solución:**
1. Verifica `Configuration2.APP_PASSWORD = "1234"`
2. Ambas funciones deben usar la MISMA contraseña
3. Si cambias la contraseña, actualiza en:
   - `ProtectWorkbook()` → `Configuration2.APP_PASSWORD`
   - `UnprotectWorkbook()` → `Configuration2.APP_PASSWORD`

---

### Problema 5: "No veo mensajes de DEBUG.PRINT"

**Causa:** Ventana Inmediato no está abierta

**Solución:**
1. **Alt+F11** (VBA Editor)
2. **Vista** → **Ventana de depuración/Inmediato**
3. O presiona **Ctrl+G**

---

### Problema 6: "Puse el booleano en False pero sigue protegido" (14/04/2026)

**Causa:** La desprotección solo ocurre en `Workbook_Open()`. Si no cierras/abres el libro, no se ejecuta.

**Solución:**
1. Guarda el archivo (Ctrl+S)
2. **CIERRA** el libro completamente (File → Close)
3. **ABRE** nuevamente desde Excel
4. Espera a que veas el mensaje en Ventana Inmediato:
   ```
   [INIT] Protección de estructura (URS-22): DESACTIVADA (MODO DESARROLLO)
   [INIT] Desprotegidas 23 hojas para modo desarrollo
   ```
5. Ahora sí puedes editar/cambiar hojas

**Nota:** Cambiar el booleano no desprotege automáticamente. Necesitas cerrar/abrir para que se ejecute `Workbook_Open()` nuevamente.

---

## 📋 Checklist de Implementación

### ✅ Verificaciones Iniciales (Desarrollo)

- [ ] ENABLE_WORKBOOK_PROTECTION = **False**
- [ ] ENABLE_SHEET_PROTECTION = **False**
- [ ] Puedo eliminar hojas
- [ ] Puedo editar cualquier celda
- [ ] Archivo guarda normalmente
- [ ] No hay errores en Ventana Inmediato

### ✅ Verificaciones Testing (URS-22)

- [ ] ENABLE_WORKBOOK_PROTECTION = **True**
- [ ] ENABLE_SHEET_PROTECTION = **False**
- [ ] NO puedo eliminar hojas
- [ ] Opciones en clic derecho están grises
- [ ] Puedo editar celdas (protección de hojas OFF)
- [ ] Ventana Inmediato muestra: "Protección de estructura: ACTIVA"

### ✅ Verificaciones Producción (Completo)

- [ ] ENABLE_WORKBOOK_PROTECTION = **True**
- [ ] ENABLE_SHEET_PROTECTION = **True**
- [ ] NO puedo eliminar hojas
- [ ] NO puedo editar hojas según rol
- [ ] Admin puede editar todo
- [ ] Usuario solo puede copiar datos
- [ ] Auditoría registra intentos fallidos
- [ ] Backups automatizados funcionan

---

## 📞 Referencia Rápida

| Necesito... | Debo cambiar... | A qué valor |
|-----------|-----------------|------------|
| Desarrollo libre | Ambos booleanos | **False** |
| Testear estructura (URS-22) | ENABLE_WORKBOOK | **True** |
| Testear protección hojas (URS-20/21) | ENABLE_SHEET | **True** |
| Producción | Ambos booleanos | **True** |
| Dibujar nuevas hojas/datos | Ambos booleanos | **False** |

---

## 📝 Resumen de Cambios Realizados (14/04/2026)

| Archivo | Cambio | Línea |
|---------|--------|------|
| **Configuration2.bas** | Agregados booleanos de control | ~100-130 |
| **ThisWorkbook.bas** | Condicionó ProtectWorkbook() a booleano | ~33 |
| **SheetProtector2.bas** | Condicionó todas las funciones a booleano | ~12, 54, 81 |

---

**Conclusión:** Sistema de protección completamente flexible y controlado mediante booleanos simples. Ideal para desarrollo ágil sin sacrificar seguridad en producción.

**Próxima revisión:** 30/04/2026
