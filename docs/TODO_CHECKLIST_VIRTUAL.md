# 📋 TODO: CHECKLIST VIRTUAL — CRONOGRAMA + FORMULARIO DE INSPECCIÓN

**Proyecto:** TH-HC-001 EN PROCESO DE VALIDACIÓN  
**Módulo:** Cronograma resumen + Formulario checklist virtual (UserForm)  
**Fecha creación:** 14/04/2026  
**Estado:** PLANIFICACIÓN  
**Referencia:** Reemplaza el enfoque de hoja de TODO_FORMULARIO.md por un UserForm con pestañas

---

## 🎯 OBJETIVO GENERAL

Implementar dos componentes principales:

1. **Cronograma resumen en Menú Principal** — Tabla ListObject que muestra las próximas inspecciones ordenadas por criticidad de puesto, con filtro por planta y doble clic para iniciar inspección.

2. **Formulario de Checklist Virtual (UserForm con tabs)** — Formulario VBA que permite al evaluador completar una inspección con cabecera contextual, 2 pestañas de preguntas (Auditoría de procesos y Técnica aséptica), observaciones por pregunta, y persistencia completa en `tblInspecciones` + `tblRespuestas`.

---

## 📐 ARQUITECTURA GENERAL

```
┌───────────────────────────────────────────────────────────────────┐
│  MENÚ PRINCIPAL (Hoja1)                                           │
│  ┌────────────────────────────────────────────────────────┐       │
│  │ [Combo: Planta]   tblResumenCronograma (ListObject)    │       │
│  │ Iniciales | Puesto | Fecha próxima | Días vencimiento  │       │
│  │ (ordenado por criticidad puesto → urgencia)            │       │
│  │ ══ Doble clic en fila ══                               │       │
│  └──────────────┬─────────────────────────────────────────┘       │
└─────────────────┼─────────────────────────────────────────────────┘
                  │
                  ▼
┌───────────────────────────────────────────────────────────────────┐
│  frmChecklistVirtual (UserForm con MultiPage)                     │
│  ┌─────────────────────────────────────────┐                      │
│  │  CABECERA (campos de contexto)          │                      │
│  │  Evaluado | Planta | Área | Equipo      │                      │
│  │  Fecha | Hora ini/fin | AY1/AY2/OP     │                      │
│  │  Evaluador | Dentro/Fuera               │                      │
│  ├─────────────────────────────────────────┤                      │
│  │  [Tab 1: Auditoría de procesos]         │                      │
│  │   Pregunta | Cumple/NoCumple/NA | Obs   │                      │
│  ├─────────────────────────────────────────┤                      │
│  │  [Tab 2: Técnica aséptica]              │                      │
│  │   Pregunta | Sí/No/NA | Obs             │                      │
│  ├─────────────────────────────────────────┤                      │
│  │  Observación general                    │                      │
│  │  [Guardar] [Cancelar]                   │                      │
│  └─────────────────────────────────────────┘                      │
└───────────────────────────────────────────────────────────────────┘
                  │
                  ▼
┌───────────────────────────────────────────────────────────────────┐
│  PERSISTENCIA                                                     │
│  tblInspecciones ← cabecera + cálculos (Historico)                │
│  tblRespuestas   ← respuestas individuales (Historico)            │
│  tblCronogramaInspecciones ← actualización post-guardado          │
│  tblResumenCronograma ← refresh automático                        │
│  Audit Trail ← log de operación                                  │
└───────────────────────────────────────────────────────────────────┘
```

---

## 📊 TABLAS INVOLUCRADAS

### Tablas de lectura (fuentes de datos)
| Tabla | Hoja | Uso |
|---|---|---|
| `tblCronogramaInspecciones` | Cronograma (Hoja20) | Fuente maestra del resumen |
| `tblPersonal` | Personal (Hoja13) | Evaluados, AY1/AY2/OP por puesto+planta |
| `tblEquipos` | Configuración | Planta → Área → Equipo (cascada) |
| `tblPuestosIniciales` | Configuración | Mapeo: Ayudante 2 → AY2, etc. |
| `tblAseguramientoCalidad` | Aseguramiento de calidad (Hoja4) | Evaluadores autorizados (col 3 = iniciales) |
| `tblPlantillas` | Checklist (Hoja12) | Plantillas por puesto+etapa |
| `tblPreguntas` | Checklist (Hoja12) | Preguntas por plantilla+sección |
| `tblSecciones` | Configuración | Secciones (Auditoría procesos, TA) |
| `tblOpcionesDeRespuesta` | Configuración | Opciones por sección |
| `tblCriticidad` | Configuración | Niveles de criticidad |
| `tblCategoriasRPN` | Configuración | Rangos RPN → categoría |

### Tablas de escritura (destino de datos)
| Tabla | Hoja | Uso |
|---|---|---|
| `tblInspecciones` | Historico (Hoja14) | Registro de inspección (cabecera + cálculos) |
| `tblRespuestas` | Historico (Hoja14) | Respuestas individuales por pregunta |
| `tblCronogramaInspecciones` | Cronograma (Hoja20) | Actualización post-inspección |

### Tabla nueva a crear
| Tabla | Hoja | Uso |
|---|---|---|
| `tblResumenCronograma` | Menú principal (Hoja1) | Vista resumen para doble clic |

---

## 📋 COLUMNAS NUEVAS REQUERIDAS EN `tblInspecciones`

Las siguientes columnas **no existen** actualmente y deben agregarse:

| Columna | Tipo | Descripción |
|---|---|---|
| `Area` | Texto | Área seleccionada de tblEquipos |
| `Linea auditada` | Texto | Equipo/sala seleccionado de tblEquipos |
| `Hora inicio` | Texto (HH:MM) | Hora de inicio de la auditoría |
| `Hora termino` | Texto (HH:MM) | Hora de término de la auditoría |
| `Iniciales AY1` | Texto | Iniciales del Ayudante 1 en la línea |
| `Iniciales AY2` | Texto | Iniciales del Ayudante 2 en la línea |
| `Iniciales OP` | Texto | Iniciales del Operador en la línea |
| `Lugar auditoria` | Texto | "Dentro del área" o "Fuera del área" |

> **⚠️ PREGUNTA PENDIENTE:** Confirmar posición exacta de estas columnas en tblInspecciones (¿al final? ¿después de Auditor?). Verificar que no colisionen con columnas existentes.

---

## ✅ PLAN DE IMPLEMENTACIÓN

---

## **FASE 0: PREPARACIÓN Y CONFIGURACIÓN** 🔧

### 0.1 — Actualizar `Configuration2.bas` con nuevas constantes
- [ ] Agregar constante `SHEET_ASEGURAMIENTO As String = "Aseguramiento de calidad"`
- [ ] Agregar constante `TABLE_EQUIPOS As String = "tblEquipos"`
- [ ] Agregar constante `TABLE_PUESTOS_INICIALES As String = "tblPuestosIniciales"`
- [ ] Agregar constante `TABLE_ASEGURAMIENTO As String = "tblAseguramientoCalidad"`
- [ ] Agregar constante `TABLE_RESUMEN_CRONOGRAMA As String = "tblResumenCronograma"`
- [ ] Agregar constantes de orden de criticidad para el cronograma:
  ```vba
  Public Function GetOrdenCriticidadPuestos() As Variant
      ' Orden de criticidad: índice menor = más crítico
      GetOrdenCriticidadPuestos = Array( _
          "Operador", _
          "Ayudante 1", _
          "Operador Electrolitos", _
          "Ayudante 1 Electrolitos", _
          "Ayudante 2", _
          "Técnico de producción - grado C", _
          "Técnico de producción - grado D", _
          "Muestreador" _
      )
  End Function
  ```
- [ ] Agregar constante para las 2 opciones de lugar de auditoría

> **⚠️ PREGUNTA PENDIENTE:** ¿Los puestos "Quimico", "Digitador" y "Etiquetado" que también están en tblPersonal deben excluirse del orden de criticidad del cronograma? Parecen no aparecer en la lista de prioridad proporcionada.

### 0.2 — Agregar columnas nuevas a `tblInspecciones` (manual en Excel)
- [ ] Abrir el archivo Excel
- [ ] Ir a hoja Historico → tblInspecciones
- [ ] Agregar las 8 columnas nuevas documentadas arriba
- [ ] Verificar que el ListObject las reconozca correctamente
- [ ] Documentar la posición final de cada columna

> **NOTA:** Esto se hace manualmente en Excel, no por VBA, para no romper la estructura existente.

### 0.3 — Crear tabla `tblResumenCronograma` en Menú principal
- [ ] Definir ubicación en Menú principal (debajo del contenido existente)
- [ ] Crear ListObject con columnas: `Iniciales`, `Puesto`, `Fecha proxima`, `Dias vencimiento`
- [ ] Agregar columnas ocultas auxiliares para el doble clic: `ID_Cronograma`, `ID_Plantilla`
- [ ] Aplicar formato visual coherente con el menú principal
- [ ] Agregar celda para combo/dropdown de planta (encima de la tabla)

> **⚠️ PREGUNTA PENDIENTE:** ¿Dónde exactamente en el Menú principal va la tabla? ¿Qué celda? ¿Hay espacio disponible? ¿Hay botones existentes que mover? Necesito ver el layout actual de Hoja1.

---

## **FASE 1: CRONOGRAMA RESUMEN EN MENÚ PRINCIPAL** 📅

### 1.1 — Crear módulo `CronogramaResumen.bas`
- [ ] Crear módulo estándar `CronogramaResumen.bas`
- [ ] Implementar `Public Sub RefrescarResumenCronograma()`
  - Leer planta seleccionada del combo/dropdown
  - Leer `tblCronogramaInspecciones` completa
  - Filtrar por planta (si hay filtro activo)
  - Filtrar solo registros con `Personal activo? = "Sí"` y `Puesto activo? = "Sí"`
  - Obtener orden de criticidad via `GetOrdenCriticidadPuestos()`
  - Ordenar resultados:
    1. **Primario:** Índice de criticidad del puesto (menor = más crítico)
    2. **Secundario:** Días de vencimiento ASC (más vencido/urgente primero)
  - Limpiar `tblResumenCronograma`
  - Escribir resultados filtrados y ordenados
  - Incluir columnas ocultas (`ID_Cronograma`, `ID_Plantilla`) para referencia al doble clic
- [ ] Implementar `Public Sub FiltrarResumenPorPlanta()`
  - Leer valor del combo/dropdown de planta
  - Llamar a `RefrescarResumenCronograma()`
- [ ] Agregar logging con ErrorLogger2

### 1.2 — Implementar evento Doble Clic en `Hoja1.bas` (Menú principal)
- [ ] Agregar evento `Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)`
  - Detectar si el clic es dentro del rango de `tblResumenCronograma`
  - Si no está en la tabla → salir (no cancelar)
  - Si está en encabezado → salir
  - Obtener la fila clickeada
  - Extraer datos de la fila: Iniciales, Puesto, ID_Plantilla, ID_Cronograma
  - `Cancel = True` (evitar modo edición)
  - Llamar a `AbrirChecklistVirtual(iniciales, idPlantilla, puesto, idCronograma)`
- [ ] Agregar confirmación: "¿Iniciar inspección para [Iniciales] - [Puesto]?"
- [ ] Agregar logging con AuditLogger2

### 1.3 — Implementar combo de planta en Menú principal
- [ ] Agregar celda/control dropdown con opciones:
  - "Todas"
  - "Therapia iv Santiago"
  - "Therapia iv Concepción"
- [ ] Vincular cambio del combo a `FiltrarResumenPorPlanta()`
- [ ] Valor por defecto: "Todas"

### 1.4 — Integrar refresh del cronograma resumen
- [ ] Llamar `RefrescarResumenCronograma()` al activar Hoja1 (`Worksheet_Activate`)
- [ ] Llamar `RefrescarResumenCronograma()` después de guardar una inspección
- [ ] Llamar `RefrescarResumenCronograma()` después de recalcular cronograma completo

> **⚠️ PREGUNTA PENDIENTE:** ¿El refresh al activar la hoja puede ser lento si hay muchos registros? ¿Implementar cache o solo refrescar si hay cambios?

---

## **FASE 2: USERFORM — ESTRUCTURA Y CABECERA** 🖼️

### 2.1 — Crear UserForm `frmChecklistVirtual`
- [ ] Crear nuevo UserForm en el proyecto VBA
- [ ] Nombre: `frmChecklistVirtual`
- [ ] Dimensiones iniciales: ~700px ancho × ~600px alto (ajustable)
- [ ] Caption: "Checklist Virtual — Inspección"
- [ ] StartUpPosition: CenterScreen
- [ ] Configurar `ScrollBars` si es necesario por cantidad de preguntas

### 2.2 — Diseñar sección de cabecera (parte superior del form)
- [ ] **Frame "fraCabecera"** — contiene todos los campos de contexto
- [ ] Campos auto-rellenados (solo lectura):
  - `lblEvaluado` + `txtEvaluado` (TextBox, Locked=True) — Iniciales del evaluado
  - `lblPuesto` + `txtPuesto` (TextBox, Locked=True) — Puesto de la inspección
  - `lblPlanta` + `txtPlanta` (TextBox, Locked=True) — Planta del evaluado
- [ ] Campos de selección:
  - `lblArea` + `cboArea` (ComboBox) — Áreas filtradas por planta desde tblEquipos
  - `lblEquipo` + `cboEquipo` (ComboBox) — Equipos filtrados por planta+área desde tblEquipos
  - `lblFecha` + `txtFecha` (TextBox) — Fecha evaluada (formato DD/MM/AAAA)
  - `lblHoraInicio` + `txtHoraInicio` (TextBox) — Formato HH:MM
  - `lblHoraTermino` + `txtHoraTermino` (TextBox) — Formato HH:MM
  - `lblAY2` + `cboAY2` (ComboBox) — Personal con puesto Ayudante 2, misma planta
  - `lblAY1` + `cboAY1` (ComboBox) — Personal con puesto Ayudante 1, misma planta
  - `lblOP` + `cboOP` (ComboBox) — Personal con puesto Operador, misma planta
  - `lblLugar` + `cboLugar` (ComboBox) — "Dentro del área" / "Fuera del área"
  - `lblEvaluador` + `cboEvaluador` (ComboBox) — Iniciales de tblAseguramientoCalidad

### 2.3 — Lógica cascada Área → Equipo
- [ ] Implementar `cboArea_Change()`:
  - Al cambiar área, recargar `cboEquipo` con equipos filtrados por planta+área de tblEquipos
  - Limpiar selección de equipo
- [ ] Implementar carga inicial de `cboArea` al abrir form:
  - Filtrar áreas únicas de tblEquipos donde Planta = planta del evaluado

### 2.4 — Lógica carga de combos de personal (AY1, AY2, OP)
- [ ] Implementar función privada `CargarComboPersonalPorPuesto(cbo As ComboBox, nombrePuesto As String, planta As String)`
  - Leer tblPersonal
  - Filtrar por: columna del puesto = "Si" AND Planta = planta AND Activo = "Si"
  - Cargar iniciales en el ComboBox
- [ ] En `UserForm_Initialize`:
  - Leer `tblPuestosIniciales` para obtener mapeo puesto→sigla
  - Usar sigla como etiqueta del Label (ej: lblAY2.Caption = "AY2")
  - Llamar `CargarComboPersonalPorPuesto` para cada combo

### 2.5 — Lógica carga de evaluadores
- [ ] Implementar carga de `cboEvaluador`:
  - Leer tblAseguramientoCalidad
  - Obtener columna 3 (iniciales)
  - Cargar en combo

---

## **FASE 3: USERFORM — PESTAÑAS DE PREGUNTAS** 📝

### 3.1 — Agregar control MultiPage al UserForm
- [ ] Agregar `MultiPage` control (`mpPreguntas`)
- [ ] Configurar 2 páginas:
  - **Página 0:** "Auditoría de procesos"
  - **Página 1:** "Técnica aséptica"
- [ ] Cada página contendrá un `Frame` con `ScrollBar` vertical para scroll de preguntas

### 3.2 — Diseñar layout dinámico de preguntas por pestaña
- [ ] Dentro de cada página del MultiPage:
  - `fraPreguntas_Etapa[N]` (Frame con ScrollBars=fmScrollBarsVertical)
  - Preguntas se generan dinámicamente al abrir el form
- [ ] Por cada pregunta, generar dinámicamente:
  - `lblPregunta_[ID]` (Label) — Número + Texto de la pregunta
  - `cboRespuesta_[ID]` (ComboBox) — Opciones de respuesta según sección
  - `txtObservacion_[ID]` (TextBox) — Observación individual (opcional)
- [ ] Espaciado vertical: ~50px por pregunta (Label 20px + Combo 20px + Obs 20px + margen)

### 3.3 — Implementar carga dinámica de preguntas
- [ ] Crear módulo privado o función `CargarPreguntasEnTab()`
- [ ] Para cada pestaña/sección:
  1. Identificar la sección (de tblSecciones): "Auditoría de procesos" o "Técnica aséptica"
  2. Buscar la plantilla correspondiente en tblPlantillas (por puesto + etapa/sección)
  3. Obtener preguntas de tblPreguntas filtradas por ID_Plantilla + ID_Seccion
  4. Obtener opciones de respuesta de tblOpcionesDeRespuesta por ID_Seccion
  5. Por cada pregunta activa (Activo=True), crear controles dinámicos en el Frame
  6. Ajustar `ScrollHeight` del frame según cantidad de preguntas

> **⚠️ PREGUNTA PENDIENTE:** ¿Cómo se relaciona Plantilla con Sección? tblPlantillas tiene columna "Etapa". ¿Cada plantilla ya está asociada a una sección/etapa (Auditoría de procesos = Etapa X, TA = Etapa Y)? ¿O una plantilla contiene preguntas de ambas secciones y se filtran por ID_Seccion en tblPreguntas?

> **⚠️ PREGUNTA PENDIENTE:** ¿Cuántas preguntas aproximadas hay por sección? Esto afecta la estrategia de scroll y rendimiento de controles dinámicos.

### 3.4 — Tema de opciones de respuesta por sección
- [ ] **Auditoría de procesos:**
  - Opciones: "Cumple", "No Cumple", "No Aplica"
  - Tipo respuesta: Selección (según tblSecciones)
- [ ] **Técnica aséptica:**
  - Opciones: "Sí", "No", "No Aplica"
  - Tipo respuesta: Selección/Puntaje (según tblSecciones)
- [ ] Cargar opciones dinámicamente desde tblOpcionesDeRespuesta por ID_Seccion

### 3.5 — Almacenamiento temporal de respuestas en memoria
- [ ] Usar `Collection` o `Dictionary` privado para almacenar respuestas en memoria
  - Key: `ID_Pregunta`
  - Value: objeto/array con `ID_Opcion`, `Valor_numerico`, `Observacion`
- [ ] Al cambiar respuesta en combo → actualizar Dictionary
- [ ] Al escribir observación → actualizar Dictionary
- [ ] Al guardar → iterar Dictionary para persistir

---

## **FASE 4: USERFORM — OBSERVACIÓN GENERAL Y BOTONES** 💬

### 4.1 — Sección inferior: Observación general
- [ ] Debajo del MultiPage:
  - `lblObsGeneral` (Label) — "Observación general"
  - `txtObsGeneral` (TextBox) — MultiLine=True, ScrollBars=fmScrollBarsVertical, Height ~60px
- [ ] Se guardará en columna "Observaciones generales" de tblInspecciones

### 4.2 — Botones de acción
- [ ] `btnGuardar` (CommandButton) — "Guardar inspección"
- [ ] `btnCancelar` (CommandButton) — "Cancelar"
- [ ] Posición: parte inferior del formulario, alineados a la derecha
- [ ] Implementar `btnGuardar_Click()`:
  - Validar campos obligatorios
  - Validar que todas las preguntas estén respondidas
  - Confirmar: "¿Guardar la inspección?"
  - Si confirma → llamar pipeline de persistencia
  - Si error → mostrar mensaje específico
  - Si éxito → cerrar formulario + refrescar cronograma
- [ ] Implementar `btnCancelar_Click()`:
  - Confirmar: "¿Cancelar? Los datos no se guardarán."
  - Si confirma → `Unload Me`

---

## **FASE 5: CAPA DE DATOS — REPOSITORIOS** 💾

### 5.1 — Crear módulo `ChecklistRepository.bas`
- [ ] Crear módulo estándar `ChecklistRepository.bas`
- [ ] Implementar `Public Function ObtenerPlantillaPorPuestoYEtapa(puesto As String, etapa As String) As Variant`
  - Buscar en tblPlantillas por Puesto + Etapa
  - Retornar: ID, Nombre, Frecuencia
  - Si no existe → retornar Empty
- [ ] Implementar `Public Function ObtenerPreguntasPorPlantilla(idPlantilla As String) As Collection`
  - Leer tblPreguntas filtrada por ID_Plantilla donde Activo=True
  - Ordenar por columna "Orden"
  - Retornar Collection de arrays: (ID, Numero, Texto, ID_Seccion, ID_Criticidad)
- [ ] Implementar `Public Function ObtenerOpcionesRespuesta(idSeccion As String) As Collection`
  - Leer tblOpcionesDeRespuesta filtrada por ID_Seccion
  - Retornar Collection de arrays: (ID, Texto_opcion, Valor_puntaje)
- [ ] Implementar `Public Function ObtenerEquiposPorPlanta(planta As String) As Collection`
  - Leer tblEquipos filtrada por Planta
  - Retornar Collection con Areas y Equipos
- [ ] Implementar `Public Function ObtenerAreasPorPlanta(planta As String) As Collection`
  - Leer tblEquipos filtrada por Planta
  - Retornar Collection de Áreas únicas
- [ ] Implementar `Public Function ObtenerEquiposPorPlantaYArea(planta As String, area As String) As Collection`
  - Leer tblEquipos filtrada por Planta + Área
  - Retornar Collection de nombres de Equipo
- [ ] Implementar `Public Function ObtenerPersonalPorPuestoYPlanta(nombreColumnaPuesto As String, planta As String) As Collection`
  - Leer tblPersonal filtrada por columna puesto = "Si" AND Planta = planta AND Activo = "Si"
  - Retornar Collection de iniciales
- [ ] Implementar `Public Function ObtenerEvaluadores() As Collection`
  - Leer tblAseguramientoCalidad columna 3
  - Retornar Collection de iniciales
- [ ] Implementar `Public Function ObtenerMapeoPuestosIniciales() As Object` (Dictionary)
  - Leer tblPuestosIniciales
  - Retornar Dictionary: Key=Puesto, Value=Sigla (ej: "Ayudante 2" → "AY2")
- [ ] Agregar logging con ErrorLogger2

### 5.2 — Crear módulo `InspectionRepository.bas`
- [ ] Crear módulo estándar `InspectionRepository.bas`
- [ ] Implementar `Public Function CrearInspeccion(datos As Object) As String`
  - Generar ID_Inspeccion (UUID)
  - Insertar nueva fila en tblInspecciones con todos los campos:
    - ID, Iniciales_personal, ID_Plantilla, Planta_ejecutora
    - Fecha_inspeccion, Auditor (evaluador)
    - Area, Linea_auditada, Hora_inicio, Hora_termino
    - Iniciales_AY1, Iniciales_AY2, Iniciales_OP
    - Lugar_auditoria
    - Estado = "En progreso"
    - Observaciones_generales
  - Retornar ID_Inspeccion generado
- [ ] Implementar `Public Sub GuardarRespuestas(idInspeccion As String, respuestas As Collection)`
  - Por cada respuesta en collection:
    - Generar ID_Respuesta (UUID)
    - Insertar fila en tblRespuestas: ID, ID_Inspeccion, ID_Pregunta, ID_Opcion, Valor_numerico, Observacion, Fecha_respuesta
  - Batch: deshabilitar ScreenUpdating y Calculation durante insert
- [ ] Implementar `Public Sub ActualizarEstadoInspeccion(idInspeccion As String, estado As String)`
  - Buscar fila por ID
  - Actualizar columna Estado
- [ ] Implementar `Public Sub ActualizarCalculosInspeccion(idInspeccion As String, calculos As Object)`
  - Actualizar columnas de scoring: TA_puntaje, TA_maximos, TA_noaplica, TA_porcentaje
  - Actualizar: RPN_calculado, Categoria, Requiere_accion
  - Actualizar: Fecha_proxima, Dias_vencimiento, Estado_programacion
  - Actualizar: Estado = "Completado"
- [ ] Implementar `Public Function ObtenerUltimasNInspecciones(iniciales As String, idPlantilla As String, n As Long) As Collection`
  - Para validación de Categoría 5 (3 inspecciones consecutivas RPN > 20)
  - Retornar últimas N inspecciones ordenadas por fecha DESC
- [ ] Agregar logging con ErrorLogger2

---

## **FASE 6: CAPA DE DOMINIO — LÓGICA DE NEGOCIO** 🧠

### 6.1 — Crear módulo `ChecklistValidator.bas`
- [ ] Crear módulo estándar `ChecklistValidator.bas`
- [ ] Implementar `Public Function ValidarCabecera(frm As frmChecklistVirtual) As String`
  - Validar todos los campos obligatorios no vacíos:
    - Evaluado, Puesto, Planta (auto, siempre llenos)
    - Área (selección requerida)
    - Equipo / Línea auditada (selección requerida)
    - Fecha evaluada (formato válido, <= hoy)
    - Hora inicio (formato HH:MM)
    - Hora término (formato HH:MM, > hora inicio si misma fecha)
    - Lugar auditoría (selección requerida)
    - Evaluador (selección requerida)
  - AY1, AY2, OP → opcionales (pueden estar vacíos)
  - Retornar "" si válido, o mensaje de error específico
- [ ] Implementar `Public Function ValidarRespuestasCompletas(respuestas As Object) As String`
  - Iterar el Dictionary de respuestas
  - Verificar que cada pregunta tenga una opción seleccionada
  - Retornar "" si válido, o "Faltan N preguntas por responder: #X, #Y, #Z"
- [ ] Implementar `Public Function ValidarTodo(frm As frmChecklistVirtual, respuestas As Object) As String`
  - Llamar ValidarCabecera + ValidarRespuestasCompletas
  - Retornar primer error encontrado o ""
- [ ] Agregar logging con ErrorLogger2

### 6.2 — Crear módulo `InspectionCalculator.bas`
- [ ] Crear módulo estándar `InspectionCalculator.bas`
- [ ] Implementar `Public Function CalcularScoringTA(respuestas As Collection, opciones As Collection) As Object`
  - Filtrar solo respuestas de sección Técnica Aséptica
  - TA_puntaje = Σ(Valor_numerico de cada respuesta)
  - TA_maximos = cantidad de preguntas × valor máximo de opción
  - TA_noaplica = cantidad de "No Aplica" × valor de esa opción
  - TA_porcentaje = (TA_puntaje × 100) / (TA_maximos - TA_noaplica)
  - Manejar caso TA_maximos - TA_noaplica = 0 (todas No Aplica)
  - Retornar Dictionary con: puntaje, maximos, noaplica, porcentaje
- [ ] Implementar `Public Function CalcularRPN(taData As Object, ...) As Double`
  - Aplicar fórmula RPN según arquitectura documentada
  - Retornar valor RPN
- [ ] Implementar `Public Function DeterminarCategoria(rpn As Double, historialRPN As Collection) As Long`
  - Buscar en tblCategoriasRPN el rango que contiene el RPN
  - Validar Categoría 5: si RPN en rango [20, ∞) en 3 inspecciones consecutivas
    - Obtener historial via InspectionRepository.ObtenerUltimasNInspecciones
  - Retornar número de categoría (1-5)
- [ ] Implementar `Public Function CalcularFechaProxima(fechaInspeccion As Date, frecuenciaMeses As Long) As Date`
  - Retornar DateAdd("m", frecuenciaMeses, fechaInspeccion)
- [ ] Agregar logging con ErrorLogger2

---

## **FASE 7: ORQUESTACIÓN — PIPELINE DE GUARDADO** 🔄

### 7.1 — Crear módulo `ChecklistOrchestrator.bas`
- [ ] Crear módulo estándar `ChecklistOrchestrator.bas`
- [ ] Implementar `Public Sub AbrirChecklistVirtual(iniciales As String, idPlantilla As String, puesto As String, idCronograma As String)`
  - Punto de entrada (llamado desde doble clic en cronograma)
  - Crear instancia de frmChecklistVirtual
  - Pasar parámetros al form (via propiedades públicas Let/Get):
    - .Evaluado = iniciales
    - .Puesto = puesto
    - .IDPlantilla = idPlantilla
    - .IDCronograma = idCronograma
    - .Planta = buscar planta del evaluado en tblPersonal
  - Mostrar form modal: `frmChecklistVirtual.Show vbModal`
- [ ] Implementar `Public Sub GuardarInspeccionCompleta(frm As frmChecklistVirtual)`
  - **PIPELINE COMPLETO (secuencial, transaccional):**
  1. Validar todo (ChecklistValidator.ValidarTodo)
     - Si inválido → MsgBox con error específico → Exit Sub
  2. Confirmar con usuario
  3. Application.ScreenUpdating = False
  4. Application.EnableEvents = False
  5. Construir objeto datos de inspección desde el form
  6. Crear registro en tblInspecciones (InspectionRepository.CrearInspeccion)
     - Obtener ID_Inspeccion generado
  7. Guardar respuestas en tblRespuestas (InspectionRepository.GuardarRespuestas)
  8. Calcular scoring TA (InspectionCalculator.CalcularScoringTA)
  9. Calcular RPN (InspectionCalculator.CalcularRPN)
  10. Determinar categoría (InspectionCalculator.DeterminarCategoria)
  11. Calcular fecha próxima (InspectionCalculator.CalcularFechaProxima)
  12. Actualizar tblInspecciones con cálculos (InspectionRepository.ActualizarCalculosInspeccion)
  13. Actualizar tblCronogramaInspecciones (InspectionScheduler.ActualizarRegistroCronograma)
  14. Log en Audit Trail (AuditLogger2.LogAction)
  15. Application.ScreenUpdating = True
  16. Application.EnableEvents = True
  17. MsgBox "Inspección guardada exitosamente"
  18. Unload frm
  19. Refrescar cronograma resumen (CronogramaResumen.RefrescarResumenCronograma)
  - **Si error en cualquier paso:**
    - Rollback manual: eliminar fila de tblInspecciones + filas de tblRespuestas
    - Restaurar Application.ScreenUpdating y Application.EnableEvents
    - Log error en ErrorLogger2
    - MsgBox con error detallado
- [ ] Agregar logging con AuditLogger2 y ErrorLogger2

---

## **FASE 8: INTEGRACIÓN CON SISTEMA EXISTENTE** 🔌

### 8.1 — Actualizar `Hoja1.bas` (Menú principal)
- [ ] Agregar evento `Worksheet_BeforeDoubleClick` para tblResumenCronograma
- [ ] Agregar lógica de llamada a `CronogramaResumen.RefrescarResumenCronograma` en `Worksheet_Activate`
- [ ] Si existe combo de planta → vincular evento Change a `FiltrarResumenPorPlanta`

### 8.2 — Actualizar `InspectionScheduler.bas`
- [ ] Verificar que `ActualizarRegistroCronograma()` acepte los nuevos datos
- [ ] Asegurar que actualice: Total inspecciones, Fecha última, ID última, RPN última, Categoría última, Fecha próxima, Días vencimiento, Estado
- [ ] Integrar llamada post-guardado desde ChecklistOrchestrator

### 8.3 — Actualizar `CronogramaButtons.bas`
- [ ] Modificar `btnNuevaInspeccion_Click()` si existe:
  - En lugar de "funcionalidad en desarrollo", llamar a `ChecklistOrchestrator.AbrirChecklistVirtual` con datos de fila seleccionada en cronograma
- [ ] Agregar botón "Refrescar resumen" si aplica

### 8.4 — Actualizar `NavigationService2.bas` si es necesario
- [ ] Si el formulario requiere navegación especial (ej: volver al menú al cerrar)
- [ ] Asegurar que la protección de hojas no bloquee la apertura del form

### 8.5 — Actualizar `Configuration2.bas`
- [ ] Todas las constantes de Fase 0.1

---

## **FASE 9: TESTING Y VALIDACIÓN** ✅

### 9.1 — Pruebas unitarias por módulo
- [ ] **ChecklistRepository:** Verificar que todas las funciones de lectura retornen datos correctos
  - ObtenerPlantillaPorPuestoYEtapa con puestos válidos e inválidos
  - ObtenerPreguntasPorPlantilla con plantilla existente y vacía
  - ObtenerEquiposPorPlanta, ObtenerAreasPorPlanta, cascada Planta→Área→Equipo
  - ObtenerPersonalPorPuestoYPlanta con filtros de planta y puesto
  - ObtenerEvaluadores
  - ObtenerMapeoPuestosIniciales
- [ ] **InspectionRepository:** Verificar CRUD correcto
  - CrearInspeccion → verificar fila creada en tblInspecciones
  - GuardarRespuestas → verificar N filas en tblRespuestas
  - ActualizarCalculosInspeccion → verificar actualización de cálculos
- [ ] **ChecklistValidator:** Verificar validaciones
  - Campos vacíos → error específico
  - Formatos inválidos (fecha, hora)
  - Preguntas sin responder → lista de faltantes
- [ ] **InspectionCalculator:** Verificar cálculos
  - Scoring TA con datos conocidos → resultado esperado
  - RPN con datos conocidos → resultado esperado
  - Categoría por rango RPN → categoría correcta
  - Categoría 5 con historial de 3 RPNs > 20

### 9.2 — Pruebas de integración
- [ ] **Flujo completo: Menú → Doble clic → Form → Responder → Guardar → Cronograma actualizado**
  1. Filtrar cronograma por planta Santiago
  2. Doble clic en primera fila
  3. Verificar que form se abra con datos pre-cargados
  4. Seleccionar área y equipo (cascada funciona)
  5. Ingresar fecha, horas
  6. Seleccionar AY1, AY2, OP
  7. Seleccionar evaluador
  8. Responder todas las preguntas en ambas pestañas
  9. Escribir observación en al menos 1 pregunta
  10. Escribir observación general
  11. Guardar
  12. Verificar tblInspecciones → nueva fila con datos correctos
  13. Verificar tblRespuestas → N filas con respuestas
  14. Verificar tblCronogramaInspecciones → registro actualizado
  15. Verificar tblResumenCronograma → tabla refrescada
- [ ] **Flujo con cancelación:**
  - Abrir form → llenar parcialmente → cancelar → confirmar → verificar que NO se guardó nada
- [ ] **Flujo con errores de validación:**
  - Intentar guardar sin completar → verificar mensaje específico
  - Corregir → guardar exitosamente

### 9.3 — Pruebas de integridad de datos
- [ ] Verificar que UUIDs se generen sin duplicados
- [ ] Verificar relaciones FK: ID_Inspeccion, ID_Pregunta, ID_Opcion
- [ ] Verificar que valores numéricos se calculen correctamente
- [ ] Verificar que columnas nuevas en tblInspecciones se escriban correctamente

### 9.4 — Pruebas de rendimiento
- [ ] Medir tiempo de apertura del form con ~30 preguntas
- [ ] Medir tiempo de guardado (inserción en tblInspecciones + tblRespuestas)
- [ ] Medir tiempo de refresh del cronograma resumen
- [ ] Si alguno > 3 segundos → optimizar con ScreenUpdating/Calculation off

### 9.5 — Pruebas de auditoría
- [ ] Verificar log en Audit Trail al guardar inspección
- [ ] Verificar log de errores en ErrorLogger2
- [ ] Verificar que la protección de hojas siga funcionando post-implementación

---

## **FASE 10: REFINAMIENTO Y EDGE CASES** 🔧

### 10.1 — Manejo de casos borde
- [ ] ¿Qué pasa si el usuario cierra el form con la X? → Confirmar antes de cerrar
  - Implementar `UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)`
- [ ] ¿Qué pasa si no hay preguntas para la plantilla? → Mensaje y no abrir form
- [ ] ¿Qué pasa si tblEquipos no tiene equipos para esa planta? → Combo vacío con mensaje
- [ ] ¿Qué pasa si no hay evaluadores en tblAseguramientoCalidad? → Mensaje de error
- [ ] ¿Qué pasa si la persona ya tiene una inspección "En progreso"? → Decidir política

> **⚠️ PREGUNTA PENDIENTE:** ¿Se permite tener múltiples inspecciones "En progreso" para la misma persona+plantilla? ¿O se debe bloquear?

### 10.2 — Experiencia de usuario
- [ ] Agregar indicador visual de progreso en pestañas (ej: "Tab 1 ✓" si todas respondidas)
- [ ] Agregar contador: "Respondidas: X / Y" por pestaña
- [ ] Formato visual de ComboBox de respuesta:
  - Cumple/Sí → fondo verde al seleccionar
  - No Cumple/No → fondo rojo
  - No Aplica → fondo gris
- [ ] Resaltar preguntas sin responder al intentar guardar
- [ ] Scroll automático a primera pregunta sin responder al fallar validación

### 10.3 — Protección y seguridad
- [ ] Verificar que solo roles autorizados puedan abrir el form
  - ¿Solo Admin? ¿O cualquier usuario puede hacer inspecciones?
- [ ] Verificar que datos no se corrompan si Excel se cierra inesperadamente
  - Sin guardado parcial → se pierde todo (aceptado por el usuario)
- [ ] Verificar que el form no permita inyección de datos maliciosos
  - Sanitizar observaciones (limitar caracteres especiales)

> **⚠️ PREGUNTA PENDIENTE:** ¿Quién puede hacer inspecciones? ¿Solo administradores o cualquier usuario?

---

## 📁 MÓDULOS VBA NUEVOS A CREAR

| Módulo | Tipo | Fase | Responsabilidad |
|---|---|---|---|
| `frmChecklistVirtual.frm` | UserForm | 2-4 | Formulario visual con pestañas |
| `CronogramaResumen.bas` | Módulo estándar | 1 | Refresh y filtrado del resumen en menú |
| `ChecklistRepository.bas` | Módulo estándar | 5 | Lectura de tablas fuente (preguntas, equipos, personal) |
| `InspectionRepository.bas` | Módulo estándar | 5 | CRUD en tblInspecciones + tblRespuestas |
| `ChecklistValidator.bas` | Módulo estándar | 6 | Validación de campos y respuestas |
| `InspectionCalculator.bas` | Módulo estándar | 6 | Cálculos TA scoring, RPN, categoría |
| `ChecklistOrchestrator.bas` | Módulo estándar | 7 | Pipeline de orquestación (abrir form, guardar) |

## 📁 MÓDULOS EXISTENTES A MODIFICAR

| Módulo | Fase | Cambios |
|---|---|---|
| `Configuration2.bas` | 0 | Nuevas constantes (tablas, hojas, orden criticidad) |
| `Hoja1.bas` | 8 | Evento BeforeDoubleClick + Activate refresh |
| `InspectionScheduler.bas` | 8 | Integración post-guardado |
| `CronogramaButtons.bas` | 8 | Reemplazo de botón "funcionalidad en desarrollo" |

---

## ⚠️ PREGUNTAS PENDIENTES (RESOLVER ANTES DE IMPLEMENTAR)

| # | Pregunta | Contexto | Fase afectada |
|---|---|---|---|
| P1 | ¿Dónde exactamente en Menú principal va el cronograma resumen? (celda, fila) | Necesito ver el layout actual | 0, 1 |
| P2 | ¿Los puestos "Químico", "Digitador", "Etiquetado" se excluyen del cronograma? | No aparecen en la lista de criticidad dada | 0, 1 |
| P3 | ¿Posición de las columnas nuevas en tblInspecciones? | ¿Al final? ¿Después de Auditor? | 0 |
| P4 | ¿Cómo se relaciona Plantilla con Sección? ¿Una plantilla = una sección, o una plantilla tiene preguntas de ambas secciones? | Afecta la lógica de carga de pestañas | 3 |
| P5 | ¿Cuántas preguntas aproximadas por sección? | Afecta rendimiento y diseño de scroll | 3 |
| P6 | ¿Se permite múltiples inspecciones "En progreso" para la misma persona? | Política de duplicados | 10 |
| P7 | ¿Quién puede hacer inspecciones? ¿Solo Admin o cualquier usuario? | Control de acceso | 10 |
| P8 | ¿El refresh del cronograma al activar menú puede ser lento? ¿Implementar cache? | Rendimiento con muchos registros | 1 |

---

## 🎯 CRITERIOS DE ACEPTACIÓN FINAL

### Funcionalidad
- [ ] ✅ El cronograma resumen muestra inspecciones pendientes ordenadas por criticidad
- [ ] ✅ El filtro por planta funciona correctamente
- [ ] ✅ El doble clic abre el formulario con datos pre-cargados
- [ ] ✅ La cascada Área → Equipo funciona filtrada por planta
- [ ] ✅ Los combos AY1/AY2/OP cargan personal correcto por puesto y planta
- [ ] ✅ Las 2 pestañas cargan preguntas correctas con opciones por sección
- [ ] ✅ Cada pregunta tiene campo de observación individual
- [ ] ✅ La observación general se guarda correctamente
- [ ] ✅ La validación impide guardar con datos incompletos
- [ ] ✅ El guardado persiste en tblInspecciones + tblRespuestas
- [ ] ✅ Los cálculos TA y RPN se ejecutan correctamente
- [ ] ✅ El cronograma se actualiza automáticamente post-guardado
- [ ] ✅ El Audit Trail registra la operación

### Integridad
- [ ] ✅ No se pueden crear inspecciones sin completar todos los campos obligatorios
- [ ] ✅ Las relaciones FK son consistentes
- [ ] ✅ Los UUIDs son únicos
- [ ] ✅ No se corrompen datos al cancelar

### Rendimiento
- [ ] ✅ Apertura del formulario < 3 segundos
- [ ] ✅ Guardado completo < 5 segundos
- [ ] ✅ Refresh cronograma < 2 segundos

### Seguridad
- [ ] ✅ Solo usuarios autorizados pueden guardar inspecciones
- [ ] ✅ Datos de observaciones sanitizados
- [ ] ✅ Protección de hojas mantiene integridad post-implementación
