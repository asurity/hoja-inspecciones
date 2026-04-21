# 📋 Instrucciones: Agregar Columna de Auditoría de Procesos

**Fecha:** 20/04/2026  
**Propósito:** Habilitar el guardado y visualización del resultado de Auditoría de Procesos

---

## ✅ Pasos para agregar la columna en Excel

### 1️⃣ Abrir el archivo de Excel
- Abre: `TH-HC-001 INSPECCIONES.xlsm`

### 2️⃣ Ir a la hoja "Historico"
- Navega a la pestaña **"Historico"**
- Localiza la tabla **`tblInspecciones`**

### 3️⃣ Agregar nueva columna
**Ubicación recomendada:** Entre `TA porcentaje` (columna 19) y `RPN calculado` (columna 20)

**Opciones de nombre (usa UNO de estos nombres exactamente):**
- `Auditoria Procesos Resultado` ✅ (Recomendado)
- `Auditoría Procesos Resultado`
- `Resultado Auditoria Procesos`
- `AP Resultado`

**Pasos:**
1. Click derecho en la columna 20 (RPN calculado)
2. Seleccionar "Insertar"
3. Elegir "Columnas de tabla a la izquierda"
4. Renombrar el encabezado de la nueva columna con uno de los nombres de arriba
5. La nueva columna será la 20, y RPN se moverá a la 21

### 4️⃣ Formato de la columna
- **Tipo de datos:** Texto
- **Ancho:** Ajustar a ~15-20 caracteres
- **Valores posibles:** 
  - `Cumple`
  - `No Cumple`
  - (vacío si no se evaluó)

### 5️⃣ Guardar el archivo
- Guardar el archivo Excel
- Cerrar y volver a abrir para que VBA detecte la nueva columna

---

## 🔍 Verificación

Después de agregar la columna:

1. **Ejecuta una inspección de prueba** con el formulario
2. **Revisa en debugging** (Ctrl+G en VBA):
   ```
   [ActualizarCalculos] Auditoría Procesos guardada en columna 'Auditoria Procesos Resultado' (20): Cumple
   ```
3. **Verifica en Excel** que el resultado aparezca en la nueva columna
4. **Genera un PDF** y verifica que muestre:
   ```
   Auditoría Procesos: Cumple
   ```

---

## 📊 Posición de las columnas (después del cambio)

| # | Columna anterior | # nuevo | Columna nueva |
|---|-----------------|---------|---------------|
| 19 | TA porcentaje | 19 | TA porcentaje |
| 20 | RPN calculado | **20** | **Auditoria Procesos Resultado** ⭐ NUEVA |
| 21 | Categoria resultado | 21 | RPN calculado |
| 22 | Requiere accion | 22 | Categoria resultado |
| ... | ... | ... | ... |

---

## ⚠️ Importante

- **Usa el nombre exacto:** `Auditoria Procesos Resultado` (sin tilde en "Auditoria")
- El código busca varios nombres posibles, pero el recomendado es sin tilde
- Si usas otro nombre de la lista, el código lo encontrará automáticamente

---

## 🐛 Troubleshooting

### El dato no se guarda
1. Verifica que la columna existe en `tblInspecciones`
2. Revisa que el nombre sea exacto (sin espacios extras)
3. Mira el Debug Output (Ctrl+G) para ver mensajes de advertencia

### No aparece en el PDF
1. Verifica que el dato esté guardado en Excel
2. Mira el Debug Output al generar el PDF:
   ```
   [CertificadoPDF] Auditoría Procesos encontrada en columna '...': Cumple
   [POBLACIÓN] Auditoría de Procesos: Cumple
   ```

---

## 📝 Código implementado

El sistema ya está listo con:

✅ **Conteo por criticidad** en `InspectionCalculator.ContarRespuestasPorCriticidad`  
✅ **Evaluación de reglas** en `InspectionCalculator.EvaluarAuditoriaProcesos`  
✅ **Guardado robusto** en `InspectionRepository.ActualizarCalculosInspeccion`  
✅ **Lectura desde Excel** en `CertificadoPDFGenerator.ObtenerDatosInspeccion`  
✅ **Visualización en PDF** en `CertificadoPDFGenerator.PoblarPlantillaCertificado` (celda C22)  

Solo falta **agregar la columna en Excel** para que todo funcione.

---

**Cualquier duda, revisa los mensajes de Debug en VBA (Ctrl+G) 🔍**
