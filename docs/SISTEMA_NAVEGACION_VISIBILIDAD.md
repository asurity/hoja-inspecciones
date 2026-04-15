# 👁️ SISTEMA DE NAVEGACIÓN Y VISIBILIDAD (xlSheetVeryHidden)

**Proyecto:** TH-HC-001 EN PROCESO DE VALIDACIÓN  
**Requisito:** Gestión avanzada de visibilidad de hojas  
**Estado:** ✅ IMPLEMENTADO (Parcialmente Activado)  
**Fecha:** 14/04/2026

---

## 📖 Tabla de Contenidos

1. [¿Qué es xlSheetVeryHidden?](#qué-es-xlsheetveveryhidden)
2. [Diferencia: Hidden vs VeryHidden](#diferencia-hidden-vs-veryhidden)
3. [Arquitectura del Sistema](#arquitectura-del-sistema)
4. [Módulos Involucrados](#módulos-involucrados)
5. [Flujo de Navegación Completo](#flujo-de-navegación-completo)
6. [Casos de Uso](#casos-de-uso)
7. [Estado Actual del Sistema](#estado-actual-del-sistema)
8. [Verificación y Testing](#verificación-y-testing)
9. [Roadmap Activación](#roadmap-activación)

---

## 🎯 ¿Qué es xlSheetVeryHidden?

### Definición

`xlSheetVeryHidden` es un estado de visibilidad de hojas en Excel que:

- ❌ **NO aparece** en la lista de hojas (pestaña inferior)
- ❌ **NO aparece** en el diálogo "Mostrar hojas" (clic derecho)
- ❌ **NO puede mostrarse** manualmente por el usuario
- ✅ **SOLO puede mostrarse via código VBA**

### Comparación de Estados

| Propiedad | Visible | Hidden | VeryHidden |
|-----------|---------|--------|------------|
| Se ve en pestaña | ✅ Sí | ❌ No | ❌ No |
| Usuario puede ver en diálogo | ✅ Sí | ✅ Sí | ❌ No |
| Usuario puede mostrar manualmente | ✅ Sí | ✅ Sí (clic derecho) | ❌ No |
| Solo código VBA puede mostrar | ✅ (no necesario) | ✅ (opcional) | ✅ **NECESARIO** |

### Sintaxis en VBA

```vba
' Ocultar visiblemente
ws.Visible = xlSheetHidden       ' Usuario puede mostrar manualmente

' Ocultar completamente (solo VBA puede mostrar)
ws.Visible = xlSheetVeryHidden   ' Usuario NO puede mostrar

' Hacer visible
ws.Visible = xlSheetVisible      ' Normal (visible en pestaña)
```

---

## 🛡️ Diferencia: Hidden vs VeryHidden

### Escenario: Usuario ve clic derecho

**Hidden (Normal):**
```
Hoja1 (clic derecho)
├─ Insertar
├─ Eliminar
├─ Cambiar nombre
├─ Mover o copiar
├─ ✅ Mostrar         ← Usuario VE esta opción
│  └─ HojaOculta1
│  └─ HojaOculta2
```

**VeryHidden:**
```
Hoja1 (clic derecho)
├─ Insertar
├─ Eliminar
├─ Cambiar nombre
├─ Mover o copiar
└─ (NO aparece "Mostrar") ← Usuario NO VE opción
```

---

## 🏗️ Arquitectura del Sistema

### Estado de Hojas por Tipo

```
┌────────────────────────────────────────────────────┐
│                    LIBRO EXCEL                      │
├────────────────────────────────────────────────────┤
│                                                     │
│  ┌─────────────────────────────────────────────┐  │
│  │ Menú principal                              │  │
│  │ Visible: xlSheetVisible (siempre)           │  │
│  └─────────────────────────────────────────────┘  │
│           ↑ SIEMPRE visible                       │
│           │                                       │
│  ┌─────────────────────────────────────────────┐  │
│  │ Hojas de Funcionalidad:                      │  │
│  │ • Detecciones                               │  │
│  │ • Dashboard                                 │  │
│  │ • Configuración                             │  │
│  │ • Personal producción                       │  │
│  │ • Etc.                                      │  │
│  │                                              │  │
│  │ Visible: xlSheetVeryHidden (ocultadas)      │  │
│  │ Se muestran solo cuando se navega a ellas  │  │
│  └─────────────────────────────────────────────┘  │
│           ↑ Se activan bajo demanda                │
│           │                                       │
│  ┌─────────────────────────────────────────────┐  │
│  │ Hojas Audit Trail (1-5):                    │  │
│  │ • Audit trail 1                             │  │
│  │ • Audit trail 2                             │  │
│  │ • Audit trail 3                             │  │
│  │ • Audit trail 4                             │  │
│  │ • Audit trail 5                             │  │
│  │                                              │  │
│  │ Visible: xlSheetVeryHidden (siempre)        │  │
│  │ Se muestran SOLO cuando se solicita        │  │
│  │ Audit Trail completo                        │  │
│  └─────────────────────────────────────────────┘  │
│           ↑ Se activan solo para auditoría         │
└────────────────────────────────────────────────────┘
```

---

## 🔧 Módulos Involucrados

### 1️⃣ NavigationService2.bas

**Ubicación:** Raíz del Proyecto

**Responsabilidad:** Proporciona funciones públicas para navegar hacia diferentes módulos del sistema.

**Funciones Principales:**

```vba
Public Sub NavigateToSheet(ByVal targetSheetName As String)
    ' Navega a una hoja específica (no Audit Trail)
    ' Acciones:
    ' 1. Desprotege estructura del libro
    ' 2. Muestra la hoja destino
    ' 3. Oculta todas las demás (excepto Menú principal)
    ' 4. Re-protege estructura
End Sub

Public Sub NavigateToAuditTrail()
    ' Especial: Muestra TODAS las 5 hojas de Audit Trail
    Call SheetService2.ShowAuditTrailGroup()
End Sub

Public Sub NavigateToMenu()
    ' Navega a Menú Principal (siempre visible)
    Call NavigateToSheet("Menú principal")
End Sub
```

**Funciones de Navegación Específicas:**
```vba
NavigateToDetecciones()
NavigateToDashboard()
NavigateToConfiguracion()
NavigateToPersonalProduccion()
NavigateToAuditTrail()
NavigateToMenu()
' ... más funciones según módulos
```

---

### 2️⃣ SheetService2.bas

**Ubicación:** Raíz del Proyecto

**Responsabilidad:** Gestiona la visibilidad y protección de hojas. Es el verdadero motor de navegación.

**Funciones Principales:**

#### A. HideAndProtectAllSheetsExcept()

```vba
Public Sub HideAndProtectAllSheetsExcept(ByVal sheetName As String)
    ' Oculta TODAS las hojas EXCEPTO:
    ' 1. La hoja destino (sheetName)
    ' 2. Menú principal (SIEMPRE visible)
    
    ' Lógica:
    '   1. Desprotege estructura del libro
    '   2. Oculta TODAS las hojas Audit Trail (xlSheetVeryHidden)
    '   3. Hace visible sheetName
    '   4. Oculta todas demás (xlSheetVeryHidden)
    '   5. Asegura Menú principal visible
    '   6. Re-protege estructura
End Sub
```

**Flujo:** 
```
User: Click "Ir a Detecciones"
    ↓
NavigateToSheet("Detecciones")
    ↓
HideAndProtectAllSheetsExcept("Detecciones")
    ↓
• Desprotege estructura
• Oculta: Audit trail 1-5, Configuración, Personal, etc.
• Muestra: Detecciones + Menú principal
• Re-protege estructura
    ↓
Result: SOLO ves Detecciones y Menú principal
```

---

#### B. ShowAuditTrailGroup()

```vba
Public Sub ShowAuditTrailGroup()
    ' Muestra TODAS las 5 hojas de Audit Trail y oculta el resto
    
    ' Lógica:
    '   1. Desprotege estructura
    '   2. Oculta todas hojas que NO son Audit Trail
    '   3. Hace visible TODAS las hojas Audit Trail (xlSheetVisible)
    '   4. Asegura Menú principal visible
    '   5. Re-protege estructura
End Sub
```

**Flujo:**
```
User: Click "Ver Audit Trail"
    ↓
NavigateToAuditTrail()
    ↓
ShowAuditTrailGroup()
    ↓
• Desprotege estructura
• Oculta: Detecciones, Configuración, etc. (xlSheetVeryHidden)
• Muestra: Audit trail 1, 2, 3, 4, 5 (xlSheetVisible)
• Asegura: Menú principal visible
• Re-protege estructura
    ↓
Result: Ves 5 hojas Audit + Menú principal
```

---

#### C. UnlockAndShowAllSheets() (Administración)

```vba
Public Sub UnlockAndShowAllSheets()
    ' Desprotege y muestra TODAS las hojas
    ' Uso: Solo para administración/debugging
End Sub
```

---

### 3️⃣ Configuration2.bas

**Constantes relacionadas:**

```vba
' Nombres de hojas
AUDIT_BASE_NAME = "Audit trail"
MAIN_MENU_SHEET = "Menú principal"

' Cantidades
AUDIT_MAX_SHEETS = 5  ' Hojas de auditoría

' Contraseñas
APP_PASSWORD = "1234"
AUDIT_PASSWORD = "1234"
```

---

## 🔄 Flujo de Navegación Completo

### Flujo 1: De Menú a Módulo

```
┌─────────────────────────────────────────────┐
│ Estado Inicial: Solo Menú principal         │
│ Visible: xlSheetVisible                     │
│ Ocultas: Todas demás (xlSheetVeryHidden)    │
└─────────────────────────────────────────────┘
         ↓
    Usuario hace clic en botón "Detecciones"
         ↓
┌─────────────────────────────────────────────┐
│ NavigationService2.NavigateToSheet()         │
│ → NavigateToDetecciones()                    │
│   → NavigateToSheet("Detecciones")          │
└─────────────────────────────────────────────┘
         ↓
┌─────────────────────────────────────────────┐
│ SheetService2.HideAndProtectAllSheetsExcept │
│ ("Detecciones")                             │
│                                             │
│ Acciones:                                   │
│ 1. UnprotectWorkbook()                      │
│ 2. For each Audit Trail sheet:              │
│    sheet.Visible = xlSheetVeryHidden        │
│ 3. Sheets("Detecciones").Visible = Visible │
│ 4. For each other sheet:                    │
│    sheet.Visible = xlSheetVeryHidden        │
│ 5. Sheets("Menú principal").Visible = Vis  │
│ 6. ProtectWorkbook()                        │
└─────────────────────────────────────────────┘
         ↓
┌─────────────────────────────────────────────┐
│ Estado Final: Menú + Detecciones            │
│ Visible: xlSheetVisible (2 hojas)           │
│ Ocultas: Audit 1-5, Config, etc. (VeryHid) │
└─────────────────────────────────────────────┘
```

---

### Flujo 2: De Módulo a Audit Trail

```
User: Click "Ver Registros de Auditoría"
    ↓
NavigationService2.NavigateToAuditTrail()
    ↓
SheetService2.ShowAuditTrailGroup()
    ↓
For each non-Audit sheet:
    sheet.Visible = xlSheetVeryHidden
For i = 1 to 5:
    AuditSheet(i).Visible = xlSheetVisible
Sheets("Menú principal").Visible = xlSheetVisible
    ↓
Result: Ves 6 hojas:
  • Menú principal (xcriptive)
  • Audit trail 1 (datos)
  • Audit trail 2 (datos)
  • Audit trail 3 (datos)
  • Audit trail 4 (datos)
  • Audit trail 5 (datos)
```

---

### Flujo 3: De Audit Trail a Menú (Cierre)

```
User: Click "Volver a Menú" button en Audit trail
    ↓
NavigationService2.NavigateToMenu()
    ↓
SheetService2.HideAndProtectAllSheetsExcept("Menú principal")
    ↓
For each Audit Trail sheet:
    sheet.Visible = xlSheetVeryHidden    ← Desaparecen 5 hojas
For each other sheet:
    sheet.Visible = xlSheetVeryHidden
Sheets("Menú principal").Visible = xlSheetVisible
    ↓
Result: Solo ves Menú principal
  Todas las Audit Trail vuelven a xlSheetVeryHidden
```

---

## 💼 Casos de Uso

### Caso 1: Usuario Normal - Navegación Diaria

```
9:00 AM - Abre libro
    ↓
Sistema muestra: Menú principal

10:00 AM - Click "Ir a Detecciones"
    ↓
Sistema muestra: Menú principal + Detecciones

11:30 AM - Click "Ir a Configuración"
    ↓
Sistema oculta: Detecciones (xlSheetVeryHidden)
Sistema muestra: Menú principal + Configuración

2:00 PM - Click "Ver Audit Trail"
    ↓
Sistema oculta: Configuración
Sistema muestra: Menú principal + Audit 1-5

5:00 PM - Click "Volver a Menú"
    ↓
Sistema oculta: Audit 1-5 (xlSheetVeryHidden)
Sistema muestra: SOLO Menú principal

5:05 PM - Cierra libro
    ↓
Estado guardado con todas las hojas xlSheetVeryHidden
```

---

### Caso 2: Usuario Intenta Truco

```
User: Hace clic derecho en pestaña para mostrar hojas ocultas
    ↓
Excel muestra: (ninguna opción, porque son xlSheetVeryHidden)
    ↓
User: "¿Dónde están los botones?"
    ↓
Respuesta: Only visible through NavigateToAuditTrail() button
```

---

## 🔍 Estado Actual del Sistema

### ✅ Implementado

- [x] Módulo `NavigationService2.bas` — 100% funcional
- [x] Módulo `SheetService2.bas` — 100% funcional
- [x] Hojas Audit Trail creadas (1-5) con estado xlSheetVeryHidden
- [x] Función `HideAndProtectAllSheetsExcept()`
- [x] Función `ShowAuditTrailGroup()`
- [x] Integración con `WorkbookProtector2.bas`

### ⏸️ Parcialmente Activado

La navegación está **DESACTIVADA en `Workbook_Open()`** porque todavía hay módulos pendientes:

**En ThisWorkbook.bas (línea ~60):**
```vba
' ========== SISTEMA DE NAVEGACIÓN (desactivado hasta completar diseño UI) ==========
' Descomentar cuando NavigationService2, SheetService2 y UserManager2 estén listos
' ==================================================================================

' Call SheetService2.HideAndProtectAllSheetsExcept(Configuration2.MAIN_MENU_SHEET)
' ThisWorkbook.Sheets(Configuration2.MAIN_MENU_SHEET).Activate
' Call UserManager2.DisplayUserName
' ...
```

**Razón:** Esperando que se complete:
- ❓ Diseño de botones en Menú principal
- ❓ UserManager2 finalizado
- ❓ Formularios de entrada

---

## 🧪 Verificación y Testing

### Test 1: Verificar Estado Actual

**Procedimiento:**
1. Abre el libro
2. Mira las pestañas inferiores
3. Observa qué hojas son visibles

**Resultado esperado (Actual):**
- ✅ Todas las hojas están visibles (no hay navegación activa)
- ℹ️ Razón: Sistema comentado en `Workbook_Open()`

---

### Test 2: Probar Navegación Manual

**Procedimiento:**

En VBA (Alt+F11 → Ctrl+G):

```vba
' Ir a Detecciones
Call NavigationService2.NavigateToDetecciones()
```

**Resultado esperado:**
- ✅ Desaparece: Configuración, Audit trail 1-5, etc.
- ✅ Aparece: Menú principal + Detecciones
- ✅ Desaparecen cuando haces clic derecho (xlSheetVeryHidden)

---

### Test 3: Probar Audit Trail

**Procedimiento:**

En VBA:

```vba
' Ver Audit Trail
Call NavigationService2.NavigateToAuditTrail()
```

**Resultado esperado:**
- ✅ Aparecen: Las 5 hojas Audit trail + Menú principal
- ✅ Desaparecen: Todas las demás
- ✅ No hay opción "Mostrar hojas" en clic derecho (xlSheetVeryHidden)

---

### Test 4: Volver a Menú

**Procedimiento:**

En VBA:

```vba
' Volver a Menú
Call NavigationService2.NavigateToMenu()
```

**Resultado esperado:**
- ✅ Desaparecen: Las 5 hojas Audit trail (xlSheetVeryHidden)
- ✅ Aparece: SOLO Menú principal
- ✅ Las demás permanecen xlSheetVeryHidden

---

### Test 5: Verificación Técnica

**Script VBA para auditar estado de hojas:**

```vba
Sub VerificarVisibilidadHojas()
    Dim ws As Worksheet
    Dim veryHiddenCount As Long
    Dim visibleCount As Long
    Dim hiddenCount As Long
    
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Visible
            Case xlSheetVisible
                visibleCount = visibleCount + 1
                Debug.Print "✓ " & ws.Name & " = xlSheetVisible"
            Case xlSheetHidden
                hiddenCount = hiddenCount + 1
                Debug.Print "? " & ws.Name & " = xlSheetHidden"
            Case xlSheetVeryHidden
                veryHiddenCount = veryHiddenCount + 1
                Debug.Print "✗ " & ws.Name & " = xlSheetVeryHidden (no visible manualmente)"
        End Select
    Next ws
    
    Debug.Print ""
    Debug.Print "Resumen:"
    Debug.Print "  Visible: " & visibleCount
    Debug.Print "  Hidden: " & hiddenCount
    Debug.Print "  VeryHidden: " & veryHiddenCount
End Sub
```

---

## 📋 Roadmap Activación

### Fase 1: Preparación (Actual)
- [x] Módulos NavigationService2 y SheetService2 creados
- [x] Funciones de ocultar/mostrar hojas implementadas
- [x] Auditoría Trail inicializado con xlSheetVeryHidden
- [ ] Esperar diseño de UI/botones

### Fase 2: Integración de UI
- [ ] Crear botones en Menú principal
- [ ] Asignar macros a botones (NavigateToX)
- [ ] Testing de navegación completa
- [ ] Capacitación de usuarios

### Fase 3: Activación
- [ ] Descomentar `Workbook_Open()` en ThisWorkbook.bas
- [ ] Cerrar/abrir libro → Mostrar solo Menú principal
- [ ] Verificar flujo completo
- [ ] Deploy a producción

---

## 🆘 Troubleshooting

### Problema 1: "Veo todas las hojas pero debería ver solo Menú"

**Causa:** Navegación no activada en `Workbook_Open()`

**Solución:** Es normal en DESARROLLO. El código está comentado.

---

### Problema 2: "Llamé NavigateToDetecciones() pero nada cambió"

**Causa:** Los botones no están asignados, o no hay formulario en Menú principal

**Solución:** Botones serán agregados en Fase 2

---

### Problema 3: "Hoja desapareció, ¿dónde está?"

**Causa:** Está en estado xlSheetVeryHidden, no visible manualmente

**Solución:**
```vba
' Mostrar todas las hojas (admin only)
Call SheetService2.UnlockAndShowAllSheets()
```

---

## 📞 Referencia Rápida

| Necesito... | Función | Resultado |
|-----------|---------|-----------|
| Ir a Detecciones | NavigateToDetecciones() | Menú + Detecciones visible |
| Ver Audit Trail | NavigateToAuditTrail() | Menú + Audit 1-5 visible |
| Volver a Menú | NavigateToMenu() | Solo Menú visible |
| Ver todas las hojas (admin) | UnlockAndShowAllSheets() | Todas xlSheetVisible |
| Debuggear hojas | VerificarVisibilidadHojas() | Reporte de estados |

---

## 📝 Resumen de Funcionalidad

**xlSheetVeryHidden implementado:** ✅ 100%

**Flujo de navegación:**
- Menú → Módulo → Menú → Audit Trail → Menú
- Todas las transiciones cambian visibilidad de hojas
- Hojas críticas nunca aparecen en lista de "Mostrar"

**Próximo paso:** Completar UI con botones + activar en `Workbook_Open()`

---

**Próxima revisión:** 30/04/2026
**Responsable de activación:** Equipo de Interfaz
