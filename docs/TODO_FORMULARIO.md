# 📋 TODO: SISTEMA DE FORMULARIOS DE INSPECCIÓN
**Proyecto:** TH-HC-001 EN PROCESO DE VALIDACIÓN  
**Módulo:** Formulario dinámico de inspecciones  
**Fecha creación:** 12/03/2026  
**Arquitectura:** Clean Architecture + SOLID + DRY

---

## 🎯 OBJETIVO GENERAL

Implementar un sistema robusto de formularios dinámicos en la hoja "Formulario Inspeccion" que permita:
1. **Cargar preguntas** dinámicamente desde `tblPreguntas` según plantilla seleccionada
2. **Capturar respuestas** con validación de datos y opciones contextuales
3. **Garantizar integridad** mediante protección de hojas y validaciones
4. **Persistir datos** en `tblInspecciones` y `tblRespuestas` con cálculos automáticos
5. **Escalabilidad** para múltiples plantillas sin duplicar código

---

## 🏗️ ARQUITECTURA DEL SISTEMA

### **Capas de la solución (Clean Architecture)**

```
┌─────────────────────────────────────────────────────────────────┐
│  CAPA DE PRESENTACIÓN (UI)                                      │
│  - Hoja: "Formulario Inspeccion"                                │
│  - Botones: Cargar, Guardar, Limpiar, Cancelar                  │
└──────────────────────┬──────────────────────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────────────────────┐
│  CAPA DE APLICACIÓN (Use Cases)                                 │
│  - InspectionFormManager.bas    → Orquestación del formulario   │
│  - InspectionFormButtons.bas    → Manejadores de eventos        │
└──────────────────────┬──────────────────────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────────────────────┐
│  CAPA DE DOMINIO (Business Logic)                               │
│  - InspectionCore.bas           → Lógica de negocio             │
│  - InspectionValidator.bas      → Validaciones de negocio       │
│  - InspectionCalculator.bas     → Cálculos RPN y scoring        │
└──────────────────────┬──────────────────────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────────────────────┐
│  CAPA DE INFRAESTRUCTURA (Data Access)                          │
│  - InspectionRepository.bas     → CRUD inspecciones             │
│  - QuestionRepository.bas       → Acceso a tblPreguntas         │
│  - AnswerRepository.bas         → Acceso a tblRespuestas        │
└──────────────────────┬──────────────────────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────────────────────┐
│  CAPA DE INFRAESTRUCTURA COMPARTIDA                             │
│  - Configuration2.bas           → Constantes y parámetros       │
│  - SheetProtector2.bas          → Protección de hojas           │
│  - ErrorLogger2.bas             → Manejo de errores             │
│  - AuditLogger2.bas             → Trazabilidad                  │
└─────────────────────────────────────────────────────────────────┘
```

---

## 📊 PIPELINES DEFINIDOS

### **Pipeline 1: CARGA DE FORMULARIO** 🔵

```
┌──────────────┐
│ Usuario hace │
│ clic en      │──────┐
│ "Cronograma" │      │
└──────────────┘      │
                      ▼
            ┌─────────────────────┐
            │ btnNuevaInspeccion_ │
            │ Click()             │
            └──────────┬──────────┘
                       │
                       ▼
        ┌──────────────────────────────┐
        │ InspectionFormManager.       │
        │ InicializarFormulario()      │
        ├──────────────────────────────┤
        │ 1. Limpiar formulario        │
        │ 2. Proteger/Desproteger hoja │
        │ 3. Configurar validaciones   │
        │ 4. Establecer valores por    │
        │    defecto (fechas, auditor) │
        └──────────┬───────────────────┘
                   │
                   ▼
        ┌─────────────────────────┐
        │ SheetService2.          │
        │ ActivateSheet()         │
        │ (Navegar a formulario)  │
        └─────────────────────────┘
```

---

### **Pipeline 2: SELECCIÓN DE PLANTILLA Y CARGA DE PREGUNTAS** 🟢

```
┌──────────────────┐
│ Usuario          │
│ selecciona       │────────┐
│ Puesto (C25)     │        │
└──────────────────┘        │
                            ▼
                ┌────────────────────────┐
                │ Worksheet_Change()     │
                │ (Evento en hoja)       │
                └───────────┬────────────┘
                            │
                            ▼
            ┌───────────────────────────────┐
            │ InspectionFormManager.        │
            │ CargarPreguntasPorPuesto()    │
            ├───────────────────────────────┤
            │ 1. Validar puesto seleccionado│
            │ 2. Desproteger zona preguntas │
            │ 3. Limpiar preguntas previas  │
            └───────────┬───────────────────┘
                        │
                        ▼
        ┌───────────────────────────────────┐
        │ QuestionRepository.               │
        │ ObtenerPreguntasPorPlantilla()    │
        ├───────────────────────────────────┤
        │ - Buscar ID_Plantilla por puesto  │
        │ - Leer tblPreguntas filtrada      │
        │ - Retornar array de preguntas     │
        └───────────┬───────────────────────┘
                    │
                    ▼
        ┌──────────────────────────────────┐
        │ InspectionFormManager.           │
        │ RenderizarPreguntas()            │
        ├──────────────────────────────────┤
        │ 1. Insertar filas necesarias     │
        │ 2. Escribir N°, Texto pregunta   │
        │ 3. Configurar Data Validation    │
        │    por ID_Seccion                │
        │ 4. Bloquear celdas de solo       │
        │    lectura (N°, Texto)           │
        │ 5. Proteger hoja                 │
        └──────────────────────────────────┘
```

---

### **Pipeline 3: CAPTURA DE RESPUESTAS** 🟡

```
┌──────────────────┐
│ Usuario responde │
│ preguntas en     │─────────┐
│ Columna C        │         │
└──────────────────┘         │
                             ▼
                ┌─────────────────────────┐
                │ Validación automática   │
                │ (Data Validation Excel) │
                ├─────────────────────────┤
                │ - Solo opciones válidas │
                │ - Según sección         │
                └─────────────────────────┘
                             │
                             ▼
                ┌─────────────────────────┐
                │ Formato condicional     │
                │ (visual feedback)       │
                ├─────────────────────────┤
                │ - Verde: respondida     │
                │ - Rojo: vacía           │
                └─────────────────────────┘
```

---

### **Pipeline 4: VALIDACIÓN PREVIA A GUARDAR** 🟠

```
┌──────────────────┐
│ Usuario hace     │
│ clic en          │─────────┐
│ "GUARDAR"        │         │
└──────────────────┘         │
                             ▼
            ┌────────────────────────────────┐
            │ btnGuardar_Click()             │
            │ (InspectionFormButtons)        │
            └────────────┬───────────────────┘
                         │
                         ▼
        ┌────────────────────────────────────┐
        │ InspectionValidator.               │
        │ ValidarFormularioCompleto()        │
        ├────────────────────────────────────┤
        │ ✓ Fecha inspección < HOY()         │
        │ ✓ Personal evaluado existe         │
        │ ✓ Puesto seleccionado              │
        │ ✓ Todas preguntas respondidas      │
        │ ✓ Formato horas válido (HH:MM)     │
        │ ✓ Personal involucrado válido      │
        └────────────┬───────────────────────┘
                     │
                     ├─── ❌ Error ──┐
                     │               ▼
                     │    ┌──────────────────┐
                     │    │ Mostrar mensaje  │
                     │    │ específico       │
                     │    │ Focus en campo   │
                     │    └──────────────────┘
                     │
                     ├─── ✅ Válido
                     │
                     ▼
        ┌────────────────────────────────────┐
        │ Continuar a Pipeline 5             │
        └────────────────────────────────────┘
```

---

### **Pipeline 5: PERSISTENCIA Y CÁLCULOS** 🔴

```
┌────────────────────────────────┐
│ Validación exitosa             │
└────────────┬───────────────────┘
             │
             ▼
┌─────────────────────────────────────┐
│ InspectionCore.GuardarInspeccion()  │
├─────────────────────────────────────┤
│ INICIO TRANSACCIÓN                  │
└────────────┬────────────────────────┘
             │
             ▼
┌─────────────────────────────────────┐
│ 1. Generar ID_Inspeccion (UUID)     │
└────────────┬────────────────────────┘
             │
             ▼
┌─────────────────────────────────────┐
│ 2. InspectionRepository.            │
│    CrearRegistroInspeccion()        │
├─────────────────────────────────────┤
│ - Insertar en tblInspecciones       │
│ - Estado: "En progreso"             │
└────────────┬────────────────────────┘
             │
             ▼
┌─────────────────────────────────────┐
│ 3. AnswerRepository.                │
│    GuardarRespuestas()              │
├─────────────────────────────────────┤
│ For cada respuesta:                 │
│   - Generar ID_Respuesta            │
│   - Buscar ID_Opcion seleccionada   │
│   - Copiar Valor numerico           │
│   - Insertar en tblRespuestas       │
└────────────┬────────────────────────┘
             │
             ▼
┌─────────────────────────────────────┐
│ 4. InspectionCalculator.            │
│    CalcularScoringTA()              │
├─────────────────────────────────────┤
│ - Sumar puntajes sección TA         │
│ - Calcular puntos máximos           │
│ - Aplicar lógica "No Aplica"        │
│ - Calcular porcentaje               │
└────────────┬────────────────────────┘
             │
             ▼
┌─────────────────────────────────────┐
│ 5. InspectionCalculator.            │
│    CalcularRPN()                    │
├─────────────────────────────────────┤
│ - Obtener criticidades              │
│ - Aplicar fórmula RPN               │
│ - Determinar categoría (1-5)        │
└────────────┬────────────────────────┘
             │
             ▼
┌─────────────────────────────────────┐
│ 6. InspectionRepository.            │
│    ActualizarCalculos()             │
├─────────────────────────────────────┤
│ - UPDATE tblInspecciones            │
│ - Guardar TA%, RPN, Categoría       │
│ - Estado: "Completado"              │
└────────────┬────────────────────────┘
             │
             ▼
┌─────────────────────────────────────┐
│ 7. InspectionScheduler.             │
│    ActualizarCronograma()           │
├─────────────────────────────────────┤
│ - Actualizar fecha última inspección│
│ - Calcular próxima inspección       │
│ - Recalcular estado cronograma      │
└────────────┬────────────────────────┘
             │
             ▼
┌─────────────────────────────────────┐
│ 8. AuditLogger2.Log()               │
├─────────────────────────────────────┤
│ - Registrar creación inspección     │
│ - Usuario, fecha, ID, plantilla     │
└────────────┬────────────────────────┘
             │
             ▼
┌─────────────────────────────────────┐
│ FIN TRANSACCIÓN                     │
│ Si error → Rollback (no guardar)    │
└────────────┬────────────────────────┘
             │
             ▼
┌─────────────────────────────────────┐
│ Mensaje éxito y limpieza formulario │
└─────────────────────────────────────┘
```

---

### **Pipeline 6: PROTECCIÓN Y SEGURIDAD** 🟣

```
┌────────────────────────────┐
│ En todo momento:           │
├────────────────────────────┤
│ ✓ Hoja protegida           │
│ ✓ Solo celdas C11-C34      │
│   desbloqueadas (datos)    │
│ ✓ Solo columna C (preguntas)│
│   y D (observaciones)      │
│   desbloqueadas            │
│ ✓ Resto bloqueado          │
└────────────────────────────┘
         │
         ▼
┌────────────────────────────┐
│ Al cargar preguntas:       │
├────────────────────────────┤
│ 1. Desproteger hoja        │
│ 2. Limpiar zona preguntas  │
│ 3. Renderizar preguntas    │
│ 4. Bloquear celdas A, B    │
│ 5. Desbloquear C, D        │
│ 6. Proteger hoja           │
└────────────────────────────┘
         │
         ▼
┌────────────────────────────┐
│ Al guardar:                │
├────────────────────────────┤
│ 1. Validar sin desproteger │
│ 2. Persistir datos         │
│ 3. Log auditoría           │
│ 4. Mantener protección     │
└────────────────────────────┘
```

---

## ✅ PLAN DE IMPLEMENTACIÓN

---

## **FASE 0: PREPARACIÓN DE INFRAESTRUCTURA** 🔧

### **0.1 Actualizar Configuration2.bas**
- [ ] Agregar constante `SHEET_FORMULARIO = "Formulario Inspeccion"`
- [ ] Agregar constante `TABLE_PREGUNTAS = "tblPreguntas"`
- [ ] Agregar constante `TABLE_INSPECCIONES = "tblInspecciones"`
- [ ] Agregar constante `TABLE_RESPUESTAS = "tblRespuestas"`
- [ ] Agregar constante `TABLE_PLANTILLAS = "tblPlantillas"`
- [ ] Agregar constante `TABLE_OPCIONES = "tblOpcionesDeRespuesta"`
- [ ] Agregar constante `TABLE_PERSONAL = "tblPersonal"`
- [ ] Agregar constante `FORM_PRIMERA_PREGUNTA_FILA = 44`
- [ ] Agregar constante `FORM_COL_NUMERO = 1` (A)
- [ ] Agregar constante `FORM_COL_PREGUNTA = 2` (B)
- [ ] Agregar constante `FORM_COL_RESPUESTA = 3` (C)
- [ ] Agregar constante `FORM_COL_OBSERVACION = 4` (D)

### **0.2 Crear estructura de hoja "Formulario Inspeccion"**
- [ ] Crear hoja si no existe
- [ ] Configurar diseño visual (filas 1-7: encabezado)
- [ ] Configurar sección datos temporales (filas 11-14)
- [ ] Configurar sección personal y ubicación (filas 19-27)
- [ ] Configurar sección personal involucrado (filas 32-34)
- [ ] Configurar encabezado tabla preguntas (fila 42)
- [ ] Establecer anchos de columna: A=5, B=60, C=15, D=25
- [ ] Aplicar formato visual (colores, bordes, fuentes)

### **0.3 Configurar Data Validations estáticas**
- [ ] C11: Fecha de inspección (tipo Date, >= 01/01/2020)
- [ ] C12: Fecha de ejecución (tipo Date, >= C11)
- [ ] C13: Hora inicio (formato texto "00:00", validación regex)
- [ ] C14: Hora término (formato texto "00:00", validación regex)
- [ ] C19: Personal evaluado (lista desde tblPersonal[Iniciales] donde Activo="Sí")
- [ ] C20: Auditor (texto libre, máximo 50 caracteres)
- [ ] C22: Planta (lista: "Baxter,Electrolitos,Ambas")
- [ ] C23: Área (lista dinámica según C22 - implementar en Fase 2)
- [ ] C24: Etapa (lista: "Producción,Lavado,Sanitizado,Todas")
- [ ] C25: Puesto (lista única desde tblPlantillas[Puesto])
- [ ] C26: Equipo (texto libre, máximo 100 caracteres)
- [ ] C27: Ubicación (lista: "Dentro,Fuera")
- [ ] C32: Iniciales Operador (lista tblPersonal[Iniciales], permite vacío)
- [ ] C33: Iniciales Ayudante 1 (lista tblPersonal[Iniciales], permite vacío)
- [ ] C34: Iniciales Ayudante 2 (lista tblPersonal[Iniciales], permite vacío)

### **0.4 Configurar protección inicial de hoja**
- [ ] Bloquear todas las celdas (formato → proteger celda)
- [ ] Desbloquear rango C11:C14 (datos temporales)
- [ ] Desbloquear rango C19:C27 (personal y ubicación)
- [ ] Desbloquear rango C32:C34 (personal involucrado)
- [ ] Proteger hoja con `Configuration2.APP_PASSWORD`
- [ ] Configurar permisos: AllowFiltering=True, resto=False
- [ ] EnableSelection = xlUnlockedCells

---

## **FASE 1: CAPA DE INFRAESTRUCTURA (Repositories)** 💾

### **1.1 Crear QuestionRepository.bas**
- [ ] Crear módulo `QuestionRepository.bas`
- [ ] Implementar función `ObtenerIDPlantillaPorPuesto(puesto As String) As String`
  - Buscar en tblPlantillas la primera plantilla que coincida con puesto
  - Retornar ID_Plantilla (UUID)
  - Manejar error si no existe plantilla
- [ ] Implementar función `ObtenerPreguntasPorPlantilla(idPlantilla As String) As Variant`
  - Leer tblPreguntas filtrada por ID_Plantilla
  - Ordenar por columna "Orden"
  - Retornar array 2D: (ID_Pregunta, Numero, Texto, ID_Seccion, ID_Criticidad)
  - Manejar caso de plantilla sin preguntas
- [ ] Implementar función `ObtenerSeccionPorID(idSeccion As String) As String`
  - Retornar nombre de sección dado el ID
  - Usar para debug/logs
- [ ] Agregar logging con ErrorLogger2 en todos los GoTo ErrorHandler

### **1.2 Crear AnswerRepository.bas**
- [ ] Crear módulo `AnswerRepository.bas`
- [ ] Implementar función `ObtenerOpcionesPorSeccion(idSeccion As String) As Variant`
  - Leer tblOpcionesDeRespuesta filtrada por ID_Seccion
  - Retornar array 2D: (ID_Opcion, Opcion_texto, Valor_puntaje)
  - Manejar caso sin opciones
- [ ] Implementar función `ObtenerIDOpcionPorTexto(idSeccion As String, textoOpcion As String) As String`
  - Buscar ID_Opcion dado el texto seleccionado por usuario
  - Necesario para guardar respuestas
- [ ] Implementar función `ObtenerValorNumericoPorOpcion(idOpcion As String) As Double`
  - Retornar valor numérico de la opción para cálculos
  - Manejar "-" como NULL o 0 según contexto
- [ ] Implementar sub `GuardarRespuestas(idInspeccion As String, respuestas As Variant)`
  - Recibir array de respuestas: (ID_Pregunta, ID_Opcion, Observacion)
  - Insertar filas en tblRespuestas con IDRespuesta generado (UUID)
  - Transaccional: si falla una, no guardar ninguna
- [ ] Agregar logging con ErrorLogger2

### **1.3 Crear InspectionRepository.bas**
- [ ] Crear módulo `InspectionRepository.bas`
- [ ] Implementar función `CrearInspeccion(datosInspeccion As Object) As String`
  - Generar ID_Inspeccion (UUID usando `Replace(Mid(CreateObject("Scriptlet.TypeLib").GUID, 2, 36), "-", "")`)
  - Insertar nueva fila en tblInspecciones
  - Campos iniciales: ID, Iniciales_personal, ID_Plantilla, Planta, Fecha, Auditor
  - Estado = "En progreso"
  - Retornar ID_Inspeccion
- [ ] Implementar sub `ActualizarCalculos(idInspeccion As String, calculos As Object)`
  - UPDATE fila en tblInspecciones
  - Campos: TA_puntaje, TA_maximos, TA_noaplica, TA_porcentaje, RPN_calculado, Categoria
  - Estado = "Completado"
  - Fecha_completado = Now()
- [ ] Implementar función `ExisteInspeccion(idInspeccion As String) As Boolean`
  - Validar existencia antes de actualizar
- [ ] Implementar función `ObtenerUltimasInspecciones(iniciales As String, idPlantilla As String, cantidad As Long) As Variant`
  - Para validar Categoría 5 (3 inspecciones consecutivas RPN > 20)
  - Retornar array de RPN ordenado DESC por fecha
- [ ] Agregar logging con ErrorLogger2

---

## **FASE 2: CAPA DE DOMINIO (Business Logic)** 🧠

### **2.1 Crear InspectionValidator.bas**
- [ ] Crear módulo `InspectionValidator.bas`
- [ ] Implementar función `ValidarCamposEncabezado(wsForm As Worksheet) As Boolean`
  - Validar C11 (Fecha inspección): no vacío, tipo fecha, <= HOY()
  - Validar C12 (Fecha ejecución): no vacío, >= C11
  - Validar C13, C14 (Horas): formato HH:MM válido (regex: `^\d{2}:\d{2}$`)
  - Validar C19 (Personal evaluado): no vacío, existe en tblPersonal, Activo="Sí"
  - Validar C20 (Auditor): no vacío
  - Validar C22 (Planta): no vacío
  - Validar C24 (Etapa): no vacío
  - Validar C25 (Puesto): no vacío
  - Validar C27 (Ubicación): no vacío
  - Si error: retornar mensaje específico (ByRef msgError As String)
- [ ] Implementar función `ValidarPersonalInvolucrado(wsForm As Worksheet) As Boolean`
  - Si C32, C33, C34 no vacíos → validar que existan en tblPersonal
  - Permitir vacíos (opcional)
- [ ] Implementar función `ValidarRespuestasCompletas(wsForm As Worksheet) As Boolean`
  - Iterar desde fila 44 hasta última pregunta
  - Verificar que columna C (respuesta) no esté vacía
  - Contar preguntas sin responder
  - Si hay faltantes: retornar lista de números de pregunta
- [ ] Implementar función `ValidarCoherenciaFechas(fechaInsp As Date, fechaEjec As Date, horaIni As String, horaFin As String) As Boolean`
  - Si fechaInsp = fechaEjec → horaFin > horaIni
  - Si fechaEjec > fechaInsp → válido sin restricción de horas
- [ ] Agregar logging con ErrorLogger2

### **2.2 Crear InspectionCalculator.bas**
- [ ] Crear módulo `InspectionCalculator.bas`
- [ ] Implementar función `CalcularScoringTA(respuestas As Variant) As Object`
  - Filtrar respuestas de sección "J0Wjpqx8..." (Técnica Aséptica)
  - Obtener configuración: PUNTAJE_MAXIMO_TA_BASE desde tblConfiguracion
  - Calcular: TA_puntaje_obtenido = Σ(valores numéricos)
  - Calcular: TA_puntos_no_aplica = cantidad × valor_de_opcion_NO (4)
  - Calcular: TA_puntos_maximos = 57 (base)
  - Calcular: TA_porcentaje = (obtenido × 100) / (maximos - no_aplica)
  - Retornar objeto con: {puntajeObtenido, puntosMaximos, puntosNoAplica, porcentaje}
- [ ] Implementar función `CalcularRPN(respuestas As Variant, criticidades As Variant) As Object`
  - Aplicar fórmula RPN según arquitectura del sistema
  - Considerar criticidades de cada pregunta
  - Retornar objeto con: {rpnCalculado, detalles}
- [ ] Implementar función `DeterminarCategoria(rpn As Double, historialRPN As Variant) As Integer`
  - Aplicar lógica de tblCategoriasRPN
  - Categoría 1: RPN 0-14
  - Categoría 2: RPN 15-19
  - Categoría 3: RPN 20-40
  - Categoría 4: RPN 40.01-999
  - Categoría 5: RPN 20-999 en 3 inspecciones consecutivas (validar historial)
  - Retornar número de categoría (1-5)
- [ ] Implementar función `RequiereAccion(categoria As Integer) As Boolean`
  - Categoría >= 3 → True
- [ ] Agregar logging con ErrorLogger2

### **2.3 Crear InspectionCore.bas**
- [ ] Crear módulo `InspectionCore.bas`
- [ ] Implementar sub `GuardarInspeccion(wsForm As Worksheet)`
  - **PIPELINE COMPLETO DE PERSISTENCIA**
  - 1. Validar formulario completo (llamar InspectionValidator)
  - 2. Si inválido: mostrar mensaje y salir
  - 3. Extraer datos del formulario (encabezado + respuestas)
  - 4. Crear registro en tblInspecciones (InspectionRepository)
  - 5. Guardar respuestas en tblRespuestas (AnswerRepository)
  - 6. Calcular Scoring TA (InspectionCalculator)
  - 7. Calcular RPN (InspectionCalculator)
  - 8. Determinar categoría (InspectionCalculator)
  - 9. Actualizar tblInspecciones con cálculos
  - 10. Actualizar cronograma (InspectionScheduler.ActualizarCronograma)
  - 11. Auditar operación (AuditLogger2)
  - 12. Mensaje éxito, limpiar formulario
  - **Manejo de errores transaccional: si falla, no guardar nada**
- [ ] Implementar función `ExtraerDatosFormulario(wsForm As Worksheet) As Object`
  - Leer todas las celdas del encabezado (C11-C34)
  - Leer todas las respuestas (desde fila 44)
  - Construir objeto estructurado
- [ ] Implementar función `ExtraerRespuestas(wsForm As Worksheet) As Variant`
  - Iterar filas desde 44 hasta última con contenido
  - Por cada fila: (ID_Pregunta, Texto_Respuesta, Observacion)
  - Retornar array 2D
- [ ] Agregar logging completo con AuditLogger2

---

## **FASE 3: CAPA DE APLICACIÓN (Form Management)** 🎨

### **3.1 Crear InspectionFormManager.bas**
- [ ] Crear módulo `InspectionFormManager.bas`
- [ ] Implementar sub `InicializarFormulario()`
  - Activar hoja "Formulario Inspeccion"
  - Llamar a `LimpiarFormulario()`
  - Establecer fecha inspección = HOY()
  - Establecer auditor = Environ("Username") o solicitar
  - Enfocar celda C11
- [ ] Implementar sub `LimpiarFormulario()`
  - Desproteger hoja
  - Limpiar celdas C11:C14, C19:C27, C32:C34
  - Eliminar preguntas cargadas (desde fila 44 en adelante)
  - Proteger hoja
- [ ] Implementar sub `CargarPreguntasPorPuesto(puesto As String)`
  - Obtener ID_Plantilla (QuestionRepository)
  - Si no existe plantilla: mensaje y salir
  - Obtener preguntas (QuestionRepository)
  - Si no hay preguntas: mensaje y salir
  - Llamar a `RenderizarPreguntas(preguntas)`
- [ ] Implementar sub `RenderizarPreguntas(preguntas As Variant)`
  - Desproteger hoja
  - Limpiar zona preguntas (desde fila 44)
  - Por cada pregunta:
    - Fila = 43 + índice
    - A: Numero
    - B: Texto pregunta
    - C: (vacío, configurar validación)
    - D: (vacío, observación)
  - Configurar Data Validation en columna C por ID_Seccion
  - Bloquear columnas A y B
  - Desbloquear columnas C y D
  - Aplicar formato condicional en columna C (verde si lleno, rojo si vacío)
  - Proteger hoja
- [ ] Implementar sub `ConfigurarValidacionRespuesta(celda As Range, idSeccion As String)`
  - Obtener opciones de respuesta (AnswerRepository)
  - Crear lista de validación a partir de array de opciones
  - Aplicar validación a la celda
- [ ] Implementar función `ObtenerUltimaFilaPreguntas(wsForm As Worksheet) As Long`
  - Detectar última fila con contenido en columna B (preguntas)
  - Usar para iterar respuestas
- [ ] Agregar logging con ErrorLogger2

### **3.2 Crear InspectionFormButtons.bas**
- [ ] Crear módulo `InspectionFormButtons.bas`
- [ ] Implementar sub `btnGuardar_Click()`
  - Confirmar con usuario: "¿Guardar inspección?"
  - Si Sí: llamar a `InspectionCore.GuardarInspeccion()`
  - Si error: mostrar mensaje con detalle
  - Si éxito: mensaje y navegar a Cronograma
- [ ] Implementar sub `btnLimpiar_Click()`
  - Confirmar con usuario: "¿Limpiar formulario? Se perderán datos no guardados"
  - Si Sí: llamar a `InspectionFormManager.LimpiarFormulario()`
- [ ] Implementar sub `btnCancelar_Click()`
  - Confirmar con usuario: "¿Cancelar? Los cambios no se guardarán"
  - Si Sí: limpiar formulario y navegar a Cronograma
- [ ] Vincular botones en la hoja "Formulario Inspeccion"
  - Insertar 3 botones (Shapes o ActiveX)
  - Asignar macros correspondientes
  - Posicionar según diseño (después última pregunta + 2 filas)
- [ ] Agregar logging con AuditLogger2 (acciones de usuario)

### **3.3 Implementar eventos de hoja**
- [ ] Abrir código de hoja "Formulario Inspeccion" (objeto Sheet)
- [ ] Implementar `Worksheet_Change(ByVal Target As Range)`
  - Si Target = C25 (Puesto):
    - Validar que no esté vacío
    - Llamar a `InspectionFormManager.CargarPreguntasPorPuesto(Target.Value)`
  - Si Target en rango de respuestas (C44+):
    - Aplicar formato condicional inmediato (verde)
- [ ] Implementar `Worksheet_Activate()`
  - Si formulario vacío: llamar a `InicializarFormulario()`
  - (Opcional, puede causar conflictos)
- [ ] Agregar logging con ErrorLogger2

---

## **FASE 4: INTEGRACIÓN Y PROTECCIÓN** 🔒

### **4.1 Actualizar CronogramaButtons.bas**
- [ ] Abrir `CronogramaButtons.bas`
- [ ] Reemplazar sub `btnNuevaInspeccion_Click()`
  - Eliminar mensaje "Funcionalidad en desarrollo"
  - Llamar a `InspectionFormManager.InicializarFormulario()`
  - Navegar a hoja "Formulario Inspeccion"
- [ ] Agregar logging con AuditLogger2

### **4.2 Configurar protección robusta**
- [ ] En `InspectionFormManager.RenderizarPreguntas()`:
  - Asegurar que SOLO columnas C y D de preguntas estén desbloqueadas
  - Proteger con `SheetProtector2.ProtectSheet()`
  - Password desde `Configuration2.APP_PASSWORD`
- [ ] En `InspectionFormManager.LimpiarFormulario()`:
  - Desproteger con `SheetProtector2.UnprotectSheet()`
  - Realizar limpieza
  - Re-proteger con `SheetProtector2.ProtectSheet()`
- [ ] Validar que botones funcionen con hoja protegida
  - UseInterfaceOnly = False (botones deben funcionar)
- [ ] Verificar que usuario NO pueda:
  - Eliminar filas de preguntas
  - Modificar numeración o texto de preguntas
  - Cambiar estructura del formulario
  - Modificar fórmulas de formato condicional

### **4.3 Implementar validaciones en tiempo real**
- [ ] Formato condicional en columna C (respuestas):
  - Verde (#C6EFCE): `=Y(C44<>"", CONTARA(C44)>0)`
  - Rojo (#FFC7CE): `=Y(C44="", FILA()>=44, FILA()<=ULTIMA_FILA_PREGUNTAS)`
- [ ] Formato condicional en campos obligatorios del encabezado:
  - Fondo amarillo si vacío: C11, C12, C19, C20, C22, C24, C25, C27
- [ ] Validar en `Worksheet_Change` antes de permitir guardar

---

## **FASE 5: ACTUALIZACIÓN DE CRONOGRAMA** 📅

### **5.1 Actualizar InspectionScheduler.bas**
- [ ] Implementar sub `ActualizarCronogramaDespuesInspeccion(iniciales As String, idPlantilla As String)`
  - Buscar registro en tblCronogramaInspecciones
  - Obtener última inspección (fecha, RPN, categoría)
  - Actualizar campos:
    - Total_inspecciones += 1
    - Fecha_ultima_inspeccion
    - ID_Ultima_inspeccion
    - RPN_ultima_inspeccion
    - Categoria_ultima_inspeccion
  - Calcular próxima inspección: Fecha_ultima + Frecuencia_meses
  - Recalcular días para vencimiento
  - Actualizar estado cronograma (Vigente/Por vencer/Vencido)
  - Fecha_ultima_actualizacion = Now()
- [ ] Implementar función `CalcularEstadoCronograma(diasVencimiento As Long, diasAlerta As Long) As String`
  - Si diasVencimiento < 0 → "Vencido"
  - Si diasVencimiento <= diasAlerta → "Por vencer"
  - Si diasVencimiento > diasAlerta → "Vigente"
- [ ] Validar que el UPDATE se ejecute en transacción con guardado de inspección
- [ ] Agregar logging con AuditLogger2

---

## **FASE 6: TESTING Y VALIDACIÓN** ✅

### **6.1 Pruebas unitarias**
- [ ] **Test 1: Inicialización de formulario**
  - Ejecutar `InicializarFormulario()`
  - Verificar que fecha = HOY(), auditor no vacío
  - Verificar que hoja esté protegida
- [ ] **Test 2: Carga de preguntas**
  - Seleccionar puesto "Operador"
  - Verificar que se carguen preguntas de tblPreguntas
  - Verificar que Data Validation esté configurada
  - Verificar que numeración sea correcta
- [ ] **Test 3: Validación de formulario**
  - Dejar campos obligatorios vacíos
  - Ejecutar validación
  - Verificar mensaje de error específico
- [ ] **Test 4: Guardar inspección completa**
  - Llenar formulario completo
  - Responder todas preguntas
  - Guardar
  - Verificar que se cree registro en tblInspecciones
  - Verificar que se creen N registros en tblRespuestas
  - Verificar cálculos RPN y scoring TA
- [ ] **Test 5: Protección de hoja**
  - Intentar modificar celda bloqueada
  - Verificar que Excel impida modificación
  - Verificar que columnas C y D permitan edición
- [ ] **Test 6: Limpieza de formulario**
  - Llenar formulario
  - Ejecutar limpiar
  - Verificar que todas las celdas estén vacías
  - Verificar que preguntas se eliminen

### **6.2 Pruebas de integración**
- [ ] **Flujo completo: Cronograma → Formulario → Guardar → Cronograma**
  - Iniciar desde botón en Cronograma
  - Completar inspección
  - Guardar
  - Verificar que cronograma se actualice automáticamente
  - Verificar que estado cambie (Vigente, próxima fecha calculada)
- [ ] **Flujo con errores de validación**
  - Intentar guardar con campos vacíos
  - Verificar mensaje de error
  - Corregir errores
  - Guardar exitosamente
- [ ] **Flujo con Categoría 5**
  - Crear 3 inspecciones consecutivas con RPN > 20
  - Verificar que categoría se calcule como 5
  - Verificar flag "Requiere acción"

### **6.3 Pruebas de integridad de datos**
- [ ] Verificar que UUID se generen correctamente (sin guiones, 22 caracteres)
- [ ] Verificar que no se puedan crear inspecciones duplicadas
- [ ] Verificar que relaciones FK sean consistentes:
  - ID_Inspeccion en tblRespuestas existe en tblInspecciones
  - ID_Pregunta existe en tblPreguntas
  - ID_Opcion existe en tblOpcionesDeRespuesta
- [ ] Verificar que valores numéricos se copien correctamente
- [ ] Verificar que fechas sean consistentes (fecha_ejecucion >= fecha_inspeccion)

### **6.4 Pruebas de auditoría**
- [ ] Verificar que cada guardado genere log en Audit Trail
- [ ] Verificar campos registrados: usuario, fecha, acción, tabla, ID_Registro
- [ ] Verificar que errores se registren en ErrorLogger2
- [ ] Verificar integridad de registros históricos (no modificables)

---

## **FASE 7: OPTIMIZACIÓN Y ESCALABILIDAD** 🚀

### **7.1 Optimización de rendimiento**
- [ ] Implementar cache de preguntas frecuentes (Dictionary)
- [ ] Batch insert de respuestas (1 operación vs N operaciones)
- [ ] Deshabilitar `ScreenUpdating` durante renderizado de preguntas
- [ ] Deshabilitar `Calculation` durante guardado masivo
- [ ] Implementar progress bar para guardados largos (>50 preguntas)

### **7.2 Mantenibilidad (DRY)**
- [ ] Refactorizar obtención de rangos de celda a función helper
  - `GetFormCell(campo As String) As Range`
  - Diccionario interno con mapeo campo → celda
- [ ] Crear constantes para mensajes de validación
  - `MSG_CAMPO_VACIO`, `MSG_FECHA_INVALIDA`, etc.
  - Facilitar traducción futura
- [ ] Extraer lógica de UUID a módulo `UUIDGenerator.bas`
  - Reutilizable en otros módulos
- [ ] Documentar todas las funciones públicas con comentarios estándar
  - Propósito, Argumentos, Retorno, Ejemplo

### **7.3 Escalabilidad**
- [ ] **Soporte multi-plantilla en un solo formulario**
  - Permitir seleccionar múltiples plantillas
  - Cargar preguntas de todas en secuencia
  - Identificar sección por plantilla
- [ ] **Versionado de plantillas**
  - Agregar campo `Version` en tblPlantillas
  - Registrar versión usada en inspección
  - Permitir comparación histórica
- [ ] **Preguntas con ramificación condicional**
  - Si respuesta = X → mostrar pregunta adicional
  - Implementar lógica en `RenderizarPreguntas()`
  - (Fase avanzada, opcional)
- [ ] **Formulario en múltiples páginas**
  - Si plantilla tiene >30 preguntas, dividir en páginas
  - Botones "Anterior" / "Siguiente"
  - Guardar parcialmente (estado "En progreso")
  - (Fase avanzada, opcional)

---

## **FASE 8: DOCUMENTACIÓN Y CAPACITACIÓN** 📚

### **8.1 Documentación técnica**
- [ ] Actualizar `SISTEMA_INSPECCIONES_ARQUITECTURA.md`
  - Agregar sección "Módulo de Formularios"
  - Documentar pipelines implementados
  - Diagramas de flujo finales
- [ ] Crear `FORMULARIO_MANUAL_TECNICO.md`
  - Descripción de cada módulo
  - Funciones públicas exportadas
  - Ejemplos de uso
  - Troubleshooting común
- [ ] Documentar estructura de errores
  - Códigos de error personalizados
  - Mensajes estándar
  - Acciones correctivas

### **8.2 Documentación de usuario**
- [ ] Crear `FORMULARIO_MANUAL_USUARIO.md`
  - Cómo crear una nueva inspección
  - Cómo llenar el formulario
  - Interpretación de validaciones
  - Qué hacer si hay errores
- [ ] Crear guía visual con screenshots
  - Paso a paso del flujo completo
  - Identificación de campos obligatorios
  - Interpretación de colores de validación

### **8.3 Capacitación**
- [ ] Preparar sesión de capacitación (1 hora)
  - Demo en vivo del flujo completo
  - Casos de uso reales
  - Sesión de Q&A
- [ ] Crear video tutorial (10-15 min)
  - Publicar en recurso interno
  - Transcripción para accesibilidad

---

## 🎯 CRITERIOS DE ACEPTACIÓN FINAL

### **Funcionalidad**
- ✅ Usuario puede crear inspección desde Cronograma con 1 clic
- ✅ Preguntas se cargan automáticamente según puesto seleccionado
- ✅ Todas las validaciones funcionan correctamente
- ✅ Inspección se guarda en tblInspecciones y tblRespuestas
- ✅ Cálculos RPN y scoring TA son correctos
- ✅ Cronograma se actualiza automáticamente después de guardar
- ✅ Logs de auditoría se generan correctamente

### **Seguridad e Integridad**
- ✅ Hoja protegida, solo celdas de entrada desbloqueadas
- ✅ No se pueden crear inspecciones con datos inválidos
- ✅ No se pueden modificar preguntas cargadas
- ✅ Transaccionalidad: si falla, no se guarda nada
- ✅ Relaciones FK consistentes
- ✅ UUID únicos y válidos
- ✅ Registros de auditoría inmutables

### **Arquitectura**
- ✅ Clean Architecture implementada (4 capas separadas)
- ✅ Principio DRY: sin duplicación de código
- ✅ Escalabilidad: soporta N plantillas sin cambios de código
- ✅ Mantenibilidad: constantes centralizadas, funciones documentadas
- ✅ Pipelines definidos y documentados
- ✅ Separación de responsabilidades (SRP)
- ✅ Dependency inversion (capas superiores no conocen inferiores)

### **Usabilidad**
- ✅ Interfaz intuitiva, campos claramente etiquetados
- ✅ Validaciones en tiempo real con feedback visual
- ✅ Mensajes de error específicos y claros
- ✅ Proceso completo en <5 minutos
- ✅ Documentación clara y accesible

### **Performance**
- ✅ Carga de formulario <2 segundos
- ✅ Renderizado de 50 preguntas <3 segundos
- ✅ Guardado de inspección completa <5 segundos
- ✅ No hay lag en la UI durante operaciones

---

## 📊 MÉTRICAS DE ÉXITO

### **Objetivos cuantitativos**
- [ ] **0 errores críticos** en producción durante primera semana
- [ ] **<5% tasa de validaciones fallidas** por usuario (indica formulario intuitivo)
- [ ] **100% de inspecciones** con datos íntegros (no nulos, FK válidas)
- [ ] **<10 segundos** tiempo promedio para guardar inspección
- [ ] **0 pérdidas de datos** por errores de transacción

### **Objetivos cualitativos**
- [ ] Usuarios reportan que el formulario es **"fácil de usar"**
- [ ] Equipo técnico puede **agregar nuevas validaciones sin modificar múltiples módulos**
- [ ] Nueva plantilla se integra **sin cambios de código** (solo datos)
- [ ] Código es **comprensible** para desarrolladores nuevos con <1 hora de onboarding

---

## 🔄 PLAN DE ROLLOUT

### **Semana 1: Desarrollo**
- Días 1-2: Fase 0, 1
- Días 3-4: Fase 2, 3
- Día 5: Fase 4, 5

### **Semana 2: Testing y Ajustes**
- Días 1-2: Fase 6 (testing completo)
- Días 3-4: Corrección de bugs encontrados
- Día 5: Fase 7 (optimización)

### **Semana 3: Documentación y Deploy**
- Días 1-2: Fase 8 (documentación)
- Día 3: Revisión final, aprobación
- Día 4: Deploy a producción
- Día 5: Monitoreo intensivo, hotfixes si necesario

---

## 📝 NOTAS TÉCNICAS IMPORTANTES

### **Manejo de transacciones en VBA**
VBA no tiene soporte nativo de transacciones. Implementar manualmente:
```vba
' Patrón de transacción simulada
On Error GoTo RollbackTransaction
' 1. Guardar estado previo
' 2. Realizar cambios
' 3. Si éxito: Commit (mantener cambios)
Exit Sub
RollbackTransaction:
' 4. Si error: Rollback (revertir cambios)
' Eliminar filas insertadas, restaurar valores
```

### **UUID Generation**
```vba
Function GenerarUUID() As String
    Dim guid As String
    guid = Mid(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
    GenerarUUID = Replace(guid, "-", "")
End Function
```

### **Protección durante operaciones**
```vba
' Patrón estándar
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Formulario Inspeccion")
Call SheetProtector2.UnprotectSheet(ws, Configuration2.APP_PASSWORD)
' ... operaciones ...
Call SheetProtector2.ProtectSheet(ws, Configuration2.APP_PASSWORD)
```

### **Optimización de rendimiento**
```vba
' Inicio de operaciones pesadas
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
' ... operaciones ...
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
```

---

## ✅ CHECKLIST DE REVISIÓN FINAL

Antes de marcar el proyecto como completo, verificar:

- [ ] ✅ Todas las constantes están en Configuration2.bas
- [ ] ✅ Todas las contraseñas usan constantes (no hardcoded)
- [ ] ✅ Todos los módulos tienen logging con ErrorLogger2
- [ ] ✅ Todas las operaciones críticas tienen logging con AuditLogger2
- [ ] ✅ No hay código duplicado (DRY aplicado)
- [ ] ✅ Nombres de variables descriptivos (no `x`, `i` sin contexto)
- [ ] ✅ Funciones tienen un solo propósito (SRP)
- [ ] ✅ Dependencias en una sola dirección (no circular)
- [ ] ✅ Comentarios actualizados y útiles
- [ ] ✅ No hay `MsgBox` de debug en producción
- [ ] ✅ Manejo de errores en todas las funciones públicas
- [ ] ✅ Testing completo documentado
- [ ] ✅ Manual de usuario disponible
- [ ] ✅ Manual técnico disponible
- [ ] ✅ Aprobación de QA
- [ ] ✅ Backup del libro antes de deploy

---

**FIN DEL TODO**

Este plan garantiza un sistema robusto, escalable y mantenible siguiendo las mejores prácticas de ingeniería de software aplicadas a VBA Excel.
