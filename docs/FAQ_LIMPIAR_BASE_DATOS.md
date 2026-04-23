# FAQ: Cómo Limpiar la Base de Datos de Inspecciones

## ❓ PROBLEMA: "Limpié la hoja Histórico pero siguen apareciendo inspecciones en el menú principal"

### 🔍 **Causa Raíz**

El sistema de inspecciones utiliza **MÚLTIPLES tablas** en diferentes hojas de Excel:

```
┌─────────────────────────────────────────────────────────────┐
│  ARQUITECTURA DE DATOS DEL SISTEMA DE INSPECCIONES        │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  1️⃣ Hoja "Historico"                                       │
│     └─ tblInspecciones (Tabla principal de inspecciones)   │
│     └─ tblRespuestas (Respuestas de preguntas)             │
│                                                             │
│  2️⃣ Hoja "Cronograma"                                      │
│     └─ tblCronogramaInspecciones (Programación/Tracking)   │
│                                                             │
│  3️⃣ Hoja "Menu" o "Principal" (Menú Principal)             │
│     └─ tblResumenCronograma (Vista resumen en menú)        │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### 📊 **Flujo de Datos**

```
GUARDAR INSPECCIÓN:
┌──────────────┐      ┌─────────────────┐      ┌──────────────────┐
│ Formulario   │ ───> │ Hoja Historico  │ ───> │ Hoja Cronograma  │
│ Checklist    │      │ - tblInspecciones│      │ - tblCronograma  │
│              │      │ - tblRespuestas  │      │  (actualiza fecha│
│              │      │                  │      │   próxima)       │
└──────────────┘      └─────────────────┘      └──────────────────┘
                                                          │
                                                          ▼
                                                ┌──────────────────┐
                                                │ Hoja Menu        │
                                                │ - tblResumen     │
                                                │  (refresca vista)│
                                                └──────────────────┘

MENÚ PRINCIPAL MUESTRA DATOS DE:
┌─────────────────┐
│ Hoja Cronograma │  ◄─── El menú lee desde AQUÍ
│ tblCronograma   │       NO desde Historico
└─────────────────┘
```

### ✅ **SOLUCIÓN: Limpiar TODAS las Tablas Relacionadas**

Para limpiar completamente la base de datos, debes eliminar datos de **3 hojas**:

#### **Paso 1: Limpiar Hoja "Historico"**

1. Ir a hoja `Historico`
2. Seleccionar toda la tabla `tblInspecciones` (filas de datos)
3. Click derecho → Eliminar filas
4. Seleccionar toda la tabla `tblRespuestas` (filas de datos)
5. Click derecho → Eliminar filas

#### **Paso 2: Limpiar Hoja "Cronograma"** ⚠️ **ESTO ES LO QUE FALTABA**

1. Ir a hoja `Cronograma`
2. Buscar la tabla `tblCronogramaInspecciones`
3. Seleccionar todas las filas de datos de esta tabla
4. Click derecho → Eliminar filas

**💡 IMPORTANTE**: Esta tabla es la que alimenta el menú principal. Si no la limpias, el menú seguirá mostrando inspecciones programadas incluso si limpiaste Historico.

#### **Paso 3: Refrescar el Menú Principal**

1. Ir a hoja `Menu` o `Principal`
2. Ejecutar la macro: `CronogramaResumen.RefrescarResumenCronograma()`
   - O simplemente cerrar y reabrir el archivo
   - O hacer clic en cualquier botón de "Refrescar" en el menú

Ahora sí el menú principal debería aparecer **VACÍO**.

---

## 🛠️ **Código Técnico: Dónde Lee Cada Componente**

### **frmChecklistVirtual → Buscar Historial**
```vba
' Busca en: Hoja "Historico" → tblInspecciones
Set ws = ThisWorkbook.Worksheets(Configuration2.SHEET_HISTORICO)
' Filtra por: Iniciales + Puesto + IDPlantilla
```

### **Menú Principal → Lista de Inspecciones**
```vba
' Lee desde: Hoja "Cronograma" → tblCronogramaInspecciones
Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
' Muestra en: Hoja "Menu" → tblResumenCronograma
```

### **Módulo Responsable: CronogramaResumen.bas**
```vba
Public Sub RefrescarResumenCronograma()
    ' 1. Lee tblCronogramaInspecciones (NO tblInspecciones)
    Set tblCronograma = wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA)
    
    ' 2. Filtra por planta seleccionada
    ' 3. Ordena por criticidad de puesto y urgencia
    ' 4. Escribe resultado en tblResumenCronograma (Menú Principal)
End Sub
```

---

## 📋 **Checklist de Verificación**

Después de limpiar, verifica que estas tablas estén vacías:

- [ ] `tblInspecciones` (Hoja Historico) → Sin filas de datos
- [ ] `tblRespuestas` (Hoja Historico) → Sin filas de datos
- [ ] `tblCronogramaInspecciones` (Hoja Cronograma) → Sin filas de datos ⚠️ **CLAVE**
- [ ] `tblResumenCronograma` (Hoja Menu) → Se limpia automáticamente al refrescar

---

## 🧪 **Prueba de Verificación**

1. Limpia las 3 tablas según los pasos anteriores
2. Ve al Menú Principal
3. Ejecuta: `CronogramaResumen.RefrescarResumenCronograma()`
4. Resultado esperado: **Menú principal vacío, sin inspecciones listadas**
5. Crea una nueva inspección de prueba
6. Resultado esperado: **Aparece en el menú después de guardar**

---

## 🔒 **Consideraciones de Seguridad (URS-22)**

- Si el workbook está protegido, primero debes desproteger:
  ```vba
  Call WorkbookProtector2.UnprotectWorkbook()
  ' ... limpiar tablas ...
  Call WorkbookProtector2.ProtectWorkbook()
  ```

- Las hojas también pueden estar protegidas individualmente:
  ```vba
  Call SheetProtector2.UnprotectSheet(wsHistorico)
  ' ... limpiar tabla ...
  Call SheetProtector2.ProtectSheet(wsHistorico)
  ```

---

## 📚 **Referencias Técnicas**

- **Configuración de tablas**: `Configuration2.bas`
  - `SHEET_HISTORICO = "Historico"`
  - `SHEET_CRONOGRAMA = "Cronograma"`
  - `TABLE_CRONOGRAMA = "tblCronogramaInspecciones"`
  - `TABLE_RESUMEN_CRONOGRAMA = "tblResumenCronograma"`

- **Módulo de cronograma**: `CronogramaResumen.bas` (líneas 1-150)
- **Servicio de historial**: `InspectionHistoryService.bas` (líneas 60-80)

---

## ⚡ **Macro Rápida para Limpiar Todo (Desarrollo)**

```vba
Sub LimpiarTodoElSistema()
    ' ADVERTENCIA: Esto elimina TODAS las inspecciones y respuestas
    ' Solo usar en desarrollo o para resetear sistema de prueba
    
    On Error GoTo ErrorHandler
    
    Dim wsHistorico As Worksheet
    Dim wsCronograma As Worksheet
    Dim wsMenu As Worksheet
    
    ' Desproteger si es necesario
    Call WorkbookProtector2.UnprotectWorkbook()
    
    ' Limpiar Historico
    Set wsHistorico = ThisWorkbook.Sheets(Configuration2.SHEET_HISTORICO)
    Call SheetProtector2.UnprotectSheet(wsHistorico)
    
    On Error Resume Next
    wsHistorico.ListObjects("tblInspecciones").DataBodyRange.Delete
    wsHistorico.ListObjects("tblRespuestas").DataBodyRange.Delete
    On Error GoTo ErrorHandler
    
    Call SheetProtector2.ProtectSheet(wsHistorico)
    
    ' Limpiar Cronograma
    Set wsCronograma = ThisWorkbook.Sheets(Configuration2.SHEET_CRONOGRAMA)
    Call SheetProtector2.UnprotectSheet(wsCronograma)
    
    On Error Resume Next
    wsCronograma.ListObjects(Configuration2.TABLE_CRONOGRAMA).DataBodyRange.Delete
    On Error GoTo ErrorHandler
    
    Call SheetProtector2.ProtectSheet(wsCronograma)
    
    ' Refrescar menú
    Set wsMenu = ThisWorkbook.Sheets(Configuration2.MAIN_MENU_SHEET)
    Call CronogramaResumen.RefrescarResumenCronograma()
    
    ' Reproteger
    Call WorkbookProtector2.ProtectWorkbook()
    
    MsgBox "Sistema limpiado completamente." & vbCrLf & _
           "- Historico: tblInspecciones y tblRespuestas vacias" & vbCrLf & _
           "- Cronograma: tblCronogramaInspecciones vacia" & vbCrLf & _
           "- Menu: tblResumenCronograma refrescado (vacio)", _
           vbInformation, "Limpieza Completa"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al limpiar sistema: " & Err.Description, vbCritical
    On Error Resume Next
    Call WorkbookProtector2.ProtectWorkbook()
End Sub
```

**ADVERTENCIA**: Esta macro es solo para desarrollo. No implementarla en producción sin confirmación del usuario.

---

## 📞 **Contacto y Soporte**

Si después de limpiar las 3 tablas el problema persiste:
1. Verifica que ejecutaste `RefrescarResumenCronograma()`
2. Revisa la ventana Inmediato (Ctrl+G) para ver logs
3. Confirma que la celda de filtro de planta no está causando un filtro incorrecto

**Fecha de creación**: 22/04/2026  
**Autor**: Sistema de Inspecciones - Asurity  
**Versión**: 1.0
