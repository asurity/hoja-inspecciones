# 🚀 INSTRUCCIONES: Ejecutar Setup de Plantilla Certificado

**Fecha:** 15/04/2026  
**Fase:** 1 - Creación de Plantilla Excel  
**Archivo:** `PlantillaCertificadoSetup.bas`

---

## 📋 PASOS PARA EJECUTAR EL SETUP

### **1. Abrir el archivo Excel**
- Abrir `TH-HC-001 INSPECCIONES.xlsm` en Excel
- Asegurarse de que las macros estén habilitadas

### **2. Abrir el Editor VBA**
- Presionar `Alt + F11` para abrir el Editor de Visual Basic
- En el panel izquierdo (Explorador de Proyectos), buscar el módulo `PlantillaCertificadoSetup`

### **3. Ejecutar la función de setup**

**Opción A: Desde el Editor VBA**
1. Hacer clic en cualquier parte dentro de la función `InicializarPlantillaCertificado()`
2. Presionar `F5` o hacer clic en el botón ▶️ "Ejecutar"

**Opción B: Desde Excel (recomendado)**
1. Presionar `Alt + F8` para abrir el cuadro de diálogo "Macro"
2. Seleccionar `PlantillaCertificadoSetup.InicializarPlantillaCertificado`
3. Hacer clic en "Ejecutar"

### **4. Confirmar la operación**
- Si es la primera vez: El sistema creará la hoja automáticamente
- Si la hoja ya existe: El sistema preguntará si desea recrearla

### **5. Verificar resultado**
Al completarse exitosamente, verás un mensaje:

```
Plantilla de certificado creada exitosamente.

Hoja: Plantilla Certificado
Estado: Muy oculta (xlSheetVeryHidden)
Rangos nombrados: 20+ definidos
```

---

## ✅ VERIFICACIÓN POST-SETUP

### **Verificar que la hoja existe (está muy oculta)**

Para verificar que la hoja se creó correctamente:

1. **Desde VBA (recomendado):**
   - Presionar `Alt + F11`
   - En la ventana de Inmediato (`Ctrl + G`), escribir:
   ```vba
   ? ThisWorkbook.Sheets(Configuration2.SHEET_PLANTILLA_CERTIFICADO).Name
   ```
   - Debería mostrar: `Plantilla Certificado`

2. **Hacerla visible temporalmente:**
   ```vba
   ThisWorkbook.Sheets("Plantilla Certificado").Visible = xlSheetVisible
   ```
   - Ahora podrás ver la hoja en las pestañas
   - **IMPORTANTE:** Volver a ocultarla después de verificar:
   ```vba
   ThisWorkbook.Sheets("Plantilla Certificado").Visible = xlSheetVeryHidden
   ```

### **Verificar rangos nombrados**

En Excel:
1. Ir a `Fórmulas` → `Administrador de nombres`
2. Buscar rangos que comiencen con `rngCert`
3. Deberías ver al menos 20 rangos:
   - `rngCertFechaInsp`
   - `rngCertHoraInicio`
   - `rngCertHoraFin`
   - `rngCertEvaluadoNombre`
   - ... (etc.)

---

## 🎨 DISEÑO CREADO

La plantilla incluye las siguientes secciones:

### **Encabezado (Filas 1-4)**
- Espacio para logo (placeholder)
- Título: "CERTIFICADO DE INSPECCIÓN"
- Subtítulo: "TÉCNICA ASÉPTICA"
- Proyecto: "TH-HC-001"
- Fondo gris claro

### **Datos de Inspección (Filas 6-13)**
- Fecha inspección y hora (inicio - fin)
- Evaluado (nombre e iniciales)
- Puesto y Planta
- Área y Línea
- Lugar auditoría
- Evaluador
- Personal línea (AY1, AY2, OP)

### **Resultados Generales (Filas 16-21)**
- TA Puntaje obtenido / máximos
- TA Puntos no aplica
- TA Porcentaje
- RPN
- Categoría resultado

### **Tabla de Preguntas (Fila 23+)**
- Encabezado: # | PREGUNTA | RESPUESTA | OBSERVACIÓN
- Filas de datos: Se llenarán dinámicamente por el generador

### **Placeholders**
- Observaciones generales (fila 50 - placeholder)
- Firmas (fila 55 - placeholder)

---

## ⚠️ NOTAS IMPORTANTES

1. **La hoja queda muy oculta:** No aparecerá en las pestañas normales. Solo es accesible por código VBA.

2. **Protección del libro:** Si el libro está protegido, el script lo desprotege temporalmente y lo vuelve a proteger automáticamente.

3. **Ejecución múltiple:** Puedes ejecutar el setup varias veces. Si la hoja ya existe, preguntará si deseas recrearla.

4. **Backup automático:** Se recomienda tener un backup antes de ejecutar (ya se hizo en Fase 0).

---

## 🐛 SOLUCIÓN DE PROBLEMAS

### **Error: "No se puede encontrar el proyecto o biblioteca"**
- **Causa:** Falta referencia a alguna biblioteca
- **Solución:** Esto no debería ocurrir, el código usa solo objetos nativos de Excel

### **Error: "El libro está protegido"**
- **Causa:** El libro tiene protección con contraseña diferente
- **Solución:** Verificar que `Configuration2.APP_PASSWORD` sea correcta

### **La hoja no aparece en las pestañas**
- **Causa:** Está muy oculta (comportamiento esperado)
- **Solución:** Ver sección "Hacerla visible temporalmente" arriba

### **Error al crear rangos nombrados**
- **Causa:** Ya existen rangos con el mismo nombre
- **Solución:** El script elimina automáticamente rangos `rngCert*` antes de crearlos

---

## 📝 DESPUÉS DEL SETUP

Una vez ejecutado exitosamente:

1. ✅ Guardar el archivo Excel
2. ✅ Hacer commit de los cambios (ver abajo)
3. ✅ Continuar con la Fase 2 (Módulo Generador)

### **Commit sugerido:**
```bash
git add "TH-HC-001 INSPECCIONES.xlsm" PlantillaCertificadoSetup.bas
git commit -m "feat(certificado-pdf): [FASE 1] - Plantilla Excel diseñada

- Crear hoja 'Plantilla Certificado' (muy oculta)
- Diseñar layout completo con PlantillaCertificadoSetup.bas
- Configurar PageSetup para exportación A4
- Definir 20+ rangos nombrados para población dinámica
- Encabezado, datos, resultados y tabla de preguntas"

git tag v1.4.0-cert.1
```

---

## ❓ ¿NECESITAS AYUDA?

Si encuentras algún problema durante la ejecución:
1. Verificar que el archivo Configuration2.bas esté presente
2. Revisar que la constante `SHEET_PLANTILLA_CERTIFICADO` esté definida
3. Intentar ejecutar el código paso por paso con `F8` en el Editor VBA
