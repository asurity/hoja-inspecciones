# 📋 SISTEMA DE AUDIT TRAIL - CONFIGURACIÓN Y FUNCIONAMIENTO

**Proyecto:** TH-HC-001 EN PROCESO DE VALIDACIÓN  
**Versión:** 1.0  
**Fecha:** 14/04/2026  
**Última actualización:** 14/04/2026

---

## 📖 Tabla de Contenidos

1. [Propósito del Sistema](#propósito-del-sistema)
2. [Arquitectura General](#arquitectura-general)
3. [Componentes Principales](#componentes-principales)
4. [Configuración Inicial](#configuración-inicial)
5. [Cómo Funciona la Rotación](#cómo-funciona-la-rotación)
6. [Capacidad y Límites](#capacidad-y-límites)
7. [Guía de Uso](#guía-de-uso)
8. [Monitoreo y Testing](#monitoreo-y-testing)
9. [Troubleshooting](#troubleshooting)

---

## 🎯 Propósito del Sistema

El **Audit Trail** es un sistema de auditoría centralizado que registra **TODAS** las acciones realizadas en el libro de Excel, incluyendo:

✅ Fecha y hora exacta de cada cambio  
✅ Usuario de Windows que realizó la acción  
✅ Hoja donde ocurrió el cambio  
✅ Tipo de acción (inserción, modificación, eliminación)  
✅ Campo específico modificado  
✅ Valores antes y después del cambio  
✅ Módulo y subrutina que generó el cambio  

**Objetivo:** Garantizar trazabilidad completa, conformidad regulatoria y facilidad de auditoría interna.

---

## 🏗️ Arquitectura General

### Diagrama de Flujo

```
┌─────────────────────────┐
│   Acción en Celda       │  (Usuario modifica datos)
├─────────────────────────┤
│ TableAuditor2.bas       │  Detecta cambio, extrae antes/después
├─────────────────────────┤
│ AuditLogger2.bas        │  Prepara registro, llama rotación
├─────────────────────────┤
│ AuditRotation2.bas      │  ¿Hoja actual llena? → busca siguiente
├─────────────────────────┤
│ Hoja Audit Trail        │  Escribe registro en tabla ListObject
├─────────────────────────┤
│ SheetProtector2.bas     │  Protege hoja con contraseña
└─────────────────────────┘
```

### Capas del Sistema

| Capa | Responsabilidad | Módulo |
|------|-----------------|--------|
| **Detección** | Intercepta cambios en celdas | `TableAuditor2.bas` |
| **Registro** | Formatea y prepara datos | `AuditLogger2.bas` |
| **Rotación** | Distribuye entre 5 hojas | `AuditRotation2.bas` |
| **Almacenamiento** | Escribe en tablas ListObject | Hojas: "Audit trail 1-5" |
| **Protección** | Inmutabilidad de registros | `SheetProtector2.bas` |

---

## 🔧 Componentes Principales

### 1️⃣ **AuditRotation2.bas** (Corazón del Sistema)

**Función:** Gestiona rotación automática entre 5 hojas

**Funciones clave:**
- `ObtenerHojaAuditActiva()` — Retorna la hoja con espacio disponible
- `ObtenerNombreHoja(i)` — Retorna "Audit trail 1", "Audit trail 2", etc.
- `ObtenerNombreTabla(i)` — Retorna "tblAudit1", "tblAudit2", etc.
- `DetectarSiHojaLlena(ws, tabla)` — Verifica si hoja alcanzó límite
- `ContarFilasTabla(ws, tabla)` — Cuenta registros actuales

### 2️⃣ **AuditLogger2.bas** (Escritor de Registros)

**Función:** Escribe registros en la hoja activa con espacio

**Función principal:**
- `LogAction()` — Parámetros:
  - `action` — Tipo de acción (ej: "Cargar detección")
  - `sheetName` — Hoja modificada
  - `dataModified` — Campo modificado
  - `beforeChange` — Valor anterior
  - `afterChange` — Valor nuevo
  - `moduleAndSubroutine` — Origen del cambio

**Ejemplo de uso:**
```vba
Call AuditLogger2.LogAction( _
    action:="Cambio en inspección", _
    sheetName:="Historico", _
    dataModified:="Estado", _
    beforeChange:="Pendiente", _
    afterChange:="Completado", _
    moduleAndSubroutine:="InspectionForm.btnSave_Click" _
)
```

### 3️⃣ **TableAuditor2.bas** (Auditor de Tablas)

**Función:** Agrupa cambios en rangos múltiples

**Función principal:**
- `AuditTableChanges()` — Detecta cambios en tabla y llama a `AuditLogger2`

### 4️⃣ **Configuration2.bas** (Configuración Centralizada)

**Constantes críticas:**
```vba
Public Const AUDIT_MAX_ROWS     As Long = 100        ' Filas por hoja (DEBUG) | 1000000 (PROD)
Public Const AUDIT_MAX_SHEETS   As Long = 5          ' Total de hojas
Public Const AUDIT_BASE_NAME    As String = "Audit trail"
Public Const AUDIT_TABLE_PREFIX As String = "tblAudit"
Public Const AUDIT_PASSWORD     As String = "1234"   ' Protección de hojas
```

---

## ⚙️ Configuración Inicial

### Paso 1: Crear las 5 Hojas de Auditoría

Tu libro debe contenar exactamente 5 hojas:

```
✓ Audit trail 1  (tabla: tblAudit1)  - VISIBLE
✓ Audit trail 2  (tabla: tblAudit2)  - OCULTA (xlSheetVeryHidden)
✓ Audit trail 3  (tabla: tblAudit3)  - OCULTA (xlSheetVeryHidden)
✓ Audit trail 4  (tabla: tblAudit4)  - OCULTA (xlSheetVeryHidden)
✓ Audit trail 5  (tabla: tblAudit5)  - OCULTA (xlSheetVeryHidden)
```

### Paso 2: Crear Tablas ListObject

En cada hoja, crear una tabla con estos encabezados en fila 8:

| B | C | D | E | F | G | H | I | J |
|---|---|---|---|---|---|---|---|---|
| Fecha | Hora | Usuario | Hoja | Acción | Campo | Valor anterior | Valor después | Módulo/Subrutina |

**Rango:** B8:J8 (con datos a partir de B9)

**Nombre de tabla:** `tblAudit1`, `tblAudit2`, etc.

### Paso 3: Actualizar Configuration2.bas

Según tu entorno:

**Para DESARROLLO/TESTING:**
```vba
Public Const AUDIT_MAX_ROWS = 100        ' Filas por hoja (pruebas rápidas)
```

**Para PRODUCCIÓN:**
```vba
Public Const AUDIT_MAX_ROWS = 1000000    ' 1 millón filas por hoja
```

### Paso 4: Verificar Protección de Hojas

Las hojas deben estar protegidas con contraseña `AUDIT_PASSWORD` (1234):

- **Menú:** Tools → Protect Sheet
- **Permitir a usuarios seleccionar celdas** → Sí
- **Permitir editar:** No
- **Contraseña:** 1234

---

## 🔄 Cómo Funciona la Rotación

### Ejemplo: 250 registros con límite de 100 por hoja

```
Registro 1-100   → Audit trail 1 (tblAudit1)
  ↓
[Audit trail 1 LLENA]
  ↓
Registro 101-200 → Audit trail 2 (tblAudit2)  [Se hace visible automáticamente]
  ↓
[Audit trail 2 LLENA]
  ↓
Registro 201-250 → Audit trail 3 (tblAudit3)  [Se hace visible automáticamente]
  ↓
[Audit trail 3 CON ESPACIO]
  ↓
[Sistema listo para próximos registros]
```

### Algoritmo de Detección

```vba
' En AuditRotation2.ObtenerHojaAuditActiva()
For cada hoja de 1 a 5:
    Si hoja no existe:
        Continuar a siguiente
    Si hoja está llena (>= AUDIT_MAX_ROWS):
        Continuar a siguiente
    Si hoja tiene espacio:
        RETORNAR esta hoja  ← Primera con espacio disponible
Next

Si TODAS las hojas están llenas:
    Retornar última hoja (excede límite pero no pierde datos)
```

---

## 📊 Capacidad y Límites

### Por Configuración

| Escenario | AUDIT_MAX_ROWS | Hojas | Capacidad Total | Uso |
|-----------|----------------|-------|-----------------|-----|
| Testing | 100 | 5 | 500 registros | Pruebas rápidas |
| Inspecciones mensuales | 10,000 | 5 | 50,000 registros | Small teams |
| Inspecciones anuales | 100,000 | 5 | 500,000 registros | Medium teams |
| **PRODUCCIÓN** | **1,000,000** | **5** | **5,000,000 registros** | Enterprise |

### Estimación de Crecimiento

Suponiendo **10 registros/día** (actividades rutinarias):

- **50 registros/semana**
- **200 registros/mes**
- **2,400 registros/año**

Con límite de **1,000,000/hoja** → **~416 años** de capacidad por hoja,  
Con **5 hojas** → **~2,000 años** de capacidad total.

---

## 📖 Guía de Uso

### Para Desarrolladores: Registrar una Acción

**Opción 1: Registro Simple**
```vba
Call AuditLogger2.LogAction( _
    action:="Actualización", _
    sheetName:=Me.Name, _
    dataModified:="Estado", _
    beforeChange:=cmbEstado.Value, _
    afterChange:=newState, _
    moduleAndSubroutine:="frmInspection.btnSave_Click" _
)
```

**Opción 2: Desde Evento de Cambio en Tabla**
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim beforeValues As Variant
    beforeValues = oldValues  ' Guardado previamente
    
    Call TableAuditor2.AuditTableChanges( _
        changedSheet:=Me, _
        changedRange:=Target, _
        tablesToAudit:=Array("tblInspecciones", "tblRespuestas"), _
        beforeChange:=beforeValues _
    )
End Sub
```

### Para Auditores: Revisar Registros

1. **Abrir "Audit trail 1"** (Visible siempre)
2. **Filtrar por:**
   - **Fecha:** Rango de fechas específico
   - **Usuario:** Environ("USERNAME")
   - **Acción:** Tipo de cambio
   - **Hoja:** Dónde ocurrió
3. **Si necesitas más registros:** Mostrar "Audit trail 2", "3", etc.
   - Menú: Sheet → Unhide
   - Seleccionar hojas 2-5
   - Click "OK"

### Para Administradores: Mostrar Hojas Ocultas

```vba
Sub MostrarTodasLasHojasAudit()
    Dim i As Long
    Dim ws As Worksheet
    
    For i = 2 To 5
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("Audit trail " & i)
        If Not ws Is Nothing Then ws.Visible = xlSheetVisible
        On Error GoTo 0
    Next i
    
    MsgBox "Todas las hojas Audit Trail son visibles.", vbInformation
End Sub
```

---

## 🧪 Monitoreo y Testing

### Test Automático: Generar 250 Registros

**Paso 1:** Abrir Editor VBA (Alt+F11)

**Paso 2:** Ventana Inmediato (Ctrl+G)

**Paso 3:** Ejecutar en orden:

```vba
' Primero: Cambiar AUDIT_MAX_ROWS a 100 (temporalmente)
' En Configuration2.bas, línea ~64

' Luego generar registros:
TEST_GenerarRegistrosAudit(250)

' Esperar a que termine (~3-5 segundos)

' Verificar distribución:
TEST_VerificarDistribucion()
```

**Salida esperada:**
```
[TEST] ========================================================
[TEST] Generando 250 registros de prueba...
[TEST] AUDIT_MAX_ROWS = 100 | AUDIT_MAX_SHEETS = 5
[AuditRotation2] Buscando hoja Audit Trail activa...
[AuditRotation2] Hoja 1 ('Audit trail 1'): seleccionada...
...
[TEST] ✓ Generación completada en 2.34 segundos.
[TEST] Ejecuta TEST_VerificarDistribucion para ver el resumen.
```

**Resultado de distribución:**
```
[TEST] DISTRIBUCIÓN DE REGISTROS AUDIT TRAIL
[TEST] ================================================
[TEST] Audit Trail 1    : 100 filas (Visible)
[TEST] Audit Trail 2    : 100 filas (Visible)
[TEST] Audit Trail 3    : 50 filas (Visible)
[TEST] Audit Trail 4    : 0 filas (Oculta)
[TEST] Audit Trail 5    : 0 filas (Oculta)
[TEST] ------------------------------------------------
[TEST] TOTAL            : 250 registros
```

### Función: TEST_LimpiarRegistrosPrueba

Borra todos los registros de prueba y restaura el estado inicial:

```vba
TEST_LimpiarRegistrosPrueba()
```

---

## 🔐 Seguridad

### Protección de Hojas

Cada hoja Audit Trail está protegida con:
- **Contraseña:** 1234 (configurable en AUDIT_PASSWORD)
- **Usuarios pueden:** Seleccionar celdas, copiar datos
- **Usuarios NO pueden:** Editar, insertar, eliminar filas

**Para desproteger (administrador):**
```excel
Menú: Tools → Protect Sheet → Ingresar contraseña "1234"
```

### Usuario de Windows

El sistema registra el usuario de Windows **clave** y no configurable:
```vba
auditUser = Environ("USERNAME")  ' No es Application.UserName (manipulable)
```

---

## 🆘 Troubleshooting

### Problema: "Se ha detectado un nombre ambiguo: TEST_GenerarRegistrosAudit"

**Causa:** Existen dos subrutinas con el mismo nombre en diferentes módulos.

**Solución:** 
- Eliminar la duplicada (normalmente en `mod_AuditRotation.bas`)
- Mantener la versión en `AuditRotation2.bas`

### Problema: Hojas "Audit trail 2-5" no existen

**Causa:** No se crearon las hojas pre-configuradas.

**Solución:**
```vba
InicializarHojasAuditTrail()  ' Crea las 5 hojas automáticamente
```

### Problema: Registros no aparecen

**Verificar:**
1. ¿Está la fila de acción llamando a `AuditLogger2.LogAction()`?
2. ¿Están todas las 5 hojas presentes?
3. ¿Existen las tablas con nombres correctos (tblAudit1, tblAudit2, etc.)?
4. ¿Está desactivada la protección de hojas?

**Debug:**
```vba
' Revisar Ventana Inmediato (Ctrl+G en VBA)
' Buscar mensajes [AuditRotation2] o [AuditLogger2]
```

### Problema: Sistema escribiendo en misma hoja indefinidamente

**Causa:** Todas las hojas llenas o no existe la siguiente.

**Mensaje en Inmediato:**
```
[AuditRotation2] *** CRÍTICO: todas las hojas Audit Trail están llenas.
```

**Solución:**
1. Crear nuevas hojas (Audit trail 6, 7, etc.)
2. O aumentar `AUDIT_MAX_ROWS` en `Configuration2.bas`
3. O archivar registros antiguos a base de datos

---

## 📝 Resumen de Archivo Clave

**Ubicación:** `Configuration2.bas` línea ~64

```vba
' ============================================================================
' CONFIGURACIÓN DE AUDITORÍA
' ============================================================================

Public Const AUDIT_MAX_ROWS     As Long = 1000000    ' PRODUCCIÓN: 1000000 | DEBUG: 100
Public Const AUDIT_MAX_SHEETS   As Long = 5
Public Const AUDIT_BASE_NAME    As String = "Audit trail"
Public Const AUDIT_TABLE_PREFIX As String = "tblAudit"
Public Const AUDIT_PASSWORD     As String = "1234"
```

**Para cambiar entre DEBUG y PRODUCCIÓN:**
```vba
' DEBUG (pruebas rápidas):
Public Const AUDIT_MAX_ROWS = 100

' PRODUCCIÓN (entorno real):
Public Const AUDIT_MAX_ROWS = 1000000
```

---

## 📞 Contacto para Soporte

Proyectos relacionados:
- **ErrorLogger2.bas** — Registra errores de aplicación
- **SheetProtector2.bas** — Protege hojas individuales
- **WorkbookProtector2.bas** — Protege estructura del libro

---

**Última revisión:** 14/04/2026  
**Próxima revisión recomendada:** 30/04/2026
