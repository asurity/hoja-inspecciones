# 📘 GUÍA MAESTRA - SISTEMA DE INSPECCIONES TH-HC-001

**Proyecto:** TH-HC-001 EN PROCESO DE VALIDACIÓN  
**Propósito:** Documento consolidado maestro para facilitar la creación de instructivos  
**Fecha creación:** 21/04/2026  
**Última actualización:** 21/04/2026

---

## 📋 Índice General

1. [Visión General del Sistema](#visión-general-del-sistema)
2. [Arquitectura de Datos](#arquitectura-de-datos)
3. [Configuración del Sistema](#configuración-del-sistema)
4. [Funcionalidades Principales](#funcionalidades-principales)
5. [Navegación y Seguridad](#navegación-y-seguridad)
6. [Cronogramas e Inspecciones](#cronogramas-e-inspecciones)
7. [Gestión de Datos](#gestión-de-datos)
8. [Auditoría y Trazabilidad](#auditoría-y-trazabilidad)
9. [Certificados y Reportes](#certificados-y-reportes)
10. [Desarrollo y Mantenimiento](#desarrollo-y-mantenimiento)
11. [Planes Futuros](#planes-futuros)
12. [Referencias Rápidas](#referencias-rápidas)

---

## 1. Visión General del Sistema

### 1.1 ¿Qué es TH-HC-001?

Sistema integral en Excel/VBA para **gestión de inspecciones de Técnica Aséptica** en plantas farmacéuticas, cumpliendo normativas GMP.

**Plantas soportadas:**
- Therapia IV Santiago (3 áreas)
- Therapia IV Concepción (2 áreas)

**Funcionalidades clave:**
- ✅ Cronogramas automáticos de inspecciones
- ✅ Formularios virtuales de checklist
- ✅ Cálculo automático de RPN y categorización
- ✅ Generación de certificados PDF
- ✅ Auditoría completa de operaciones
- ✅ Protección de integridad de datos
- ✅ Gestión multi-planta

### 1.2 Usuarios del Sistema

| Rol | Permisos | Funciones |
|-----|----------|-----------|
| **Administrador** | Total | Configuración, gestión de tablas, auditoría |
| **Evaluador** | Lectura + Inspecciones | Realizar inspecciones, generar certificados |
| **Consulta** | Solo lectura | Ver cronogramas, históricos |

### 1.3 Hojas Principales

| Hoja | Nombre | Propósito | Visibilidad |
|------|--------|-----------|-------------|
| 1 | Menú principal | Navegación central | Visible |
| 12 | Checklist | Plantillas y preguntas | Visible |
| 13 | Personal | Personal por planta/puesto | Visible |
| 14 | Historico | Inspecciones completadas | Visible |
| 20 | Cronograma | Planificación inspecciones | Visible |
| 4 | Aseguramiento de calidad | Evaluadores | Visible |
| 2 | Configuración | Parámetros del sistema | Oculta |
| 3-5 | Audit trail 1-5 | Trazabilidad | Muy ocultas |

**Referencias:** 
- [SISTEMA_NAVEGACION_VISIBILIDAD.md](SISTEMA_NAVEGACION_VISIBILIDAD.md)
- [GUIA_PROTECCIONES_SISTEMA.md](GUIA_PROTECCIONES_SISTEMA.md)

---

## 2. Arquitectura de Datos

### 2.1 Base de Datos (10 Tablas Principales)

#### 📊 Tablas de Configuración

**1. tblCriticidad** (Hoja Configuración)
- Define niveles de criticidad (Bajo, Medio, Alto, Crítico)
- Usa valores numéricos (1-4) para cálculos RPN

**2. tblSecciones** (Hoja Configuración)
- Secciones de inspección: Auditoría de Procesos, Técnica Aséptica
- Define tipo de respuesta y opciones

**3. tblCategoriasRPN** (Hoja Configuración)
- 5 categorías según rangos RPN
- Cat 1: Óptimo (0-14)
- Cat 2: Estable (15-19)
- Cat 3: Mejorable (20-40)
- Cat 4: Crítico (40.01-999)
- Cat 5: Recurrente (3 inspecciones consecutivas RPN > 20)

**4. tblOpcionesDeRespuesta** (Hoja Configuración)
- Opciones por sección: Sí/No/No Aplica, Cumple/No Cumple
- Valores de puntaje asociados

#### 📋 Tablas de Checklist

**5. tblPlantillas** (Hoja Checklist)
- Plantillas por puesto/proceso
- Frecuencia de inspección (meses)
- Versiones y estado activo

**6. tblPreguntas** (Hoja Checklist)
- Preguntas por plantilla y sección
- Nivel de criticidad por pregunta
- Orden de presentación

#### 👥 Tablas de Personal

**7. tblPersonal** (Hoja Personal)
- Personal por planta
- Iniciales, nombre, apellido
- Puestos asociados (matriz de columnas 3-13)
- Estado activo/inactivo

**8. tblAseguramientoCalidad** (Hoja Aseguramiento de calidad)
- Evaluadores autorizados
- Iniciales, nombre, rol

#### 📅 Tablas Transaccionales

**9. tblCronogramaInspecciones** (Hoja Cronograma)
- 24 columnas con planificación
- Personal, puesto, planta, plantilla
- Fechas: última inspección, próxima inspección
- RPN y categoría de última inspección
- Estado: Vigente/Por vencer/Vencido

**10. tblInspecciones** (Hoja Historico)
- **40 columnas** (31 originales + 9 nuevas para inspecciones recurrentes)
- Registro completo de inspección
- Cálculos TA: puntaje, máximos, no aplica, porcentaje
- RPN calculado y categoría
- Campos recurrentes: Número inspección, RPN anterior, RPN promedio, RPN total

**11. tblRespuestas** (Hoja Historico)
- Respuestas individuales por pregunta
- FK a tblInspecciones, tblPreguntas, tblOpcionesDeRespuesta
- Observaciones por pregunta

**Diagrama completo:** [SISTEMA_INSPECCIONES_ARQUITECTURA.md](SISTEMA_INSPECCIONES_ARQUITECTURA.md)

### 2.2 Relaciones Entre Tablas

```
tblPlantillas
    ├─→ tblPreguntas (ID Plantilla)
    │       ├─→ tblSecciones (ID Seccion)
    │       ├─→ tblCriticidad (ID Criticidad)
    │       └─→ tblRespuestas (ID Pregunta)
    │
    └─→ tblCronogramaInspecciones (ID Plantilla)
            └─→ tblInspecciones (Iniciales + Planta)
                    └─→ tblRespuestas (ID Inspeccion)

tblPersonal → tblCronogramaInspecciones (Iniciales personal)
tblPersonal → tblInspecciones (Iniciales personal)
```

### 2.3 Columnas Críticas - tblInspecciones (40 columnas)

**Columnas 1-31: Sistema Original**
- [10] Iniciales personal
- [13] Fecha inspeccion
- [16-19] TA puntaje, máximos, no aplica, porcentaje
- [20] Auditoria Procesos Resultado
- [21] RPN calculado
- [22] Categoria resultado

**Columnas 32-40: Sistema Recurrente** (⚠️ Nuevas - Ver plan RPN)
- [32] Numero Inspeccion (1, 2, 3...)
- [33] Es Inspeccion Recurrente (Si/No)
- [34] Puesto Evaluado (puesto específico de esta inspección)
- [35] RPN Anterior Manual
- [36] ID Inspeccion Anterior
- [37] RPN Promedio
- [38] Porcentaje Recuperacion (futuro: microbiología)
- [39] Porcentaje OOL (futuro: microbiología)
- [40] RPN Total

**Referencias:**
- [TODO_RPN_AJUSTADO.md](TODO_RPN_AJUSTADO.md) - Plan completo de inspecciones recurrentes
- Configuration2.bas líneas 260-600 - Documentación de columnas

---

## 3. Configuración del Sistema

### 3.1 Parámetros Centrales (Configuration2.bas)

**Contraseñas:**
```vba
Public Const APP_PASSWORD As String = "1234"
' Usada en:
' - Protección de estructura del libro (URS-22)
' - Protección de hojas (URS-20/21)
' - Desprotección temporal para cambios
```

**Hojas del sistema:**
```vba
Public Const SHEET_CHECKLIST As String = "Checklist"
Public Const SHEET_PERSONAL As String = "Personal"
Public Const SHEET_HISTORICO As String = "Historico"
Public Const SHEET_CONFIGURACION As String = "Configuración"
Public Const SHEET_CRONOGRAMA As String = "Cronograma"
Public Const SHEET_ASEGURAMIENTO As String = "Aseguramiento de calidad"
```

**Nombres de tablas:**
```vba
Public Const TABLE_PLANTILLAS As String = "tblPlantillas"
Public Const TABLE_PREGUNTAS As String = "tblPreguntas"
Public Const TABLE_PERSONAL As String = "tblPersonal"
Public Const TABLE_INSPECCIONES As String = "tblInspecciones"
Public Const TABLE_RESPUESTAS As String = "tblRespuestas"
Public Const TABLE_CRONOGRAMA As String = "tblCronogramaInspecciones"
```

**Audit Trail:**
```vba
Public Const AUDIT_SHEET_PREFIX As String = "Audit trail "
Public Const AUDIT_TABLE_PREFIX As String = "tblAuditTrail"
Public Const AUDIT_SHEETS_COUNT As Long = 5
Public Const AUDIT_MAX_ROWS_PER_SHEET As Long = 1000000
```

### 3.2 Tablas de Parámetros

**tblCriticidad:**
| Nivel | Nombre | Descripción |
|-------|--------|-------------|
| 1 | Bajo | Impacto mínimo en calidad |
| 2 | Medio | Requiere atención |
| 3 | Alto | Afecta calidad directamente |
| 4 | Crítico | Impacto inmediato en producto |

**tblCategoriasRPN:**
| Cat | Rango RPN | Descripción | Acción |
|-----|-----------|-------------|--------|
| 1 | 0-14 | Desempeño óptimo | Ninguna |
| 2 | 15-19 | Desempeño estable | Refuerzo preventivo |
| 3 | 20-40 | Mejorable | Plan de mejora |
| 4 | 40.01-999 | Crítico | Acción correctiva inmediata |
| 5 | 20-999 (3x) | Recurrente | Re-entrenamiento formal |

**Referencias:**
- [SISTEMA_INSPECCIONES_ARQUITECTURA.md](SISTEMA_INSPECCIONES_ARQUITECTURA.md) - Sección 3 y 4
- Configuration2.bas - Líneas 1-250

### 3.3 Controles Booleanos de Seguridad

```vba
' Control de protecciones (Fase de desarrollo vs producción)
Public Const ENABLE_WORKBOOK_PROTECTION As Boolean = True
Public Const ENABLE_SHEET_PROTECTION As Boolean = True
```

**Estado actual:** Ambos en `True` para producción

**Referencias:** [GUIA_PROTECCIONES_SISTEMA.md](GUIA_PROTECCIONES_SISTEMA.md)

---

## 4. Funcionalidades Principales

### 4.1 Cronograma Automático de Inspecciones

**Módulo:** `InspectionScheduler.bas`, `CronogramaButtons.bas`

**¿Qué hace?**
- Calcula automáticamente próximas inspecciones basándose en:
  - Frecuencia de la plantilla (ej: cada 3 meses)
  - Fecha de última inspección
  - RPN y categoría de última inspección

**Cálculo de fecha próxima:**
```vba
FechaProxima = FechaUltimaInspeccion + (FrecuenciaMeses * 30 días)
DiasParaVencimiento = FechaProxima - HOY()
```

**Estados del cronograma:**
- **Vigente:** DiasParaVencimiento > DiasAlerta
- **Por vencer:** 0 < DiasParaVencimiento ≤ DiasAlerta
- **Vencido:** DiasParaVencimiento < 0

**Botones en hoja Cronograma:**
- `btnInicializar` - Crear/actualizar cronograma completo
- `btnNuevaInspeccion` - Iniciar inspección de persona seleccionada
- `btnRefrescar` - Actualizar datos del cronograma

**Resumen en Menú Principal:**
- Tabla `tblResumenCronograma` muestra próximas inspecciones
- Ordenadas por criticidad de puesto y urgencia
- Filtro por planta
- Doble clic → inicia inspección

**Referencias:**
- [TODO_CHECKLIST_VIRTUAL.md](TODO_CHECKLIST_VIRTUAL.md) - Fase 1: Cronograma resumen
- InspectionScheduler.bas
- CronogramaButtons.bas

### 4.2 Formulario Virtual de Inspección (frmChecklistVirtual)

**Estado:** ✅ Implementado (multi-planta con filtros)

**Características:**
- **Cabecera con contexto:**
  - Evaluado (auto-carga desde cronograma)
  - Puesto (auto-carga, locked)
  - Planta (auto-carga, locked)
  - Área y Equipo (combos en cascada filtrados por planta)
  - Fecha, horas inicio/término
  - Personal línea: AY1, AY2, OP (filtrados por planta)
  - Lugar auditoría (Dentro/Fuera del área)
  - Evaluador (auto-selección con Application.UserName)

- **2 Pestañas de preguntas:**
  - Tab 1: Auditoría de Procesos (Cumple/No Cumple/No Aplica)
  - Tab 2: Técnica Aséptica (Sí/No/No Aplica)
  - Preguntas cargadas dinámicamente según plantilla

- **Observaciones:**
  - Por pregunta (campo individual)
  - General (al final del formulario)

- **Validaciones antes de guardar:**
  - Todos los campos obligatorios completos
  - Todas las preguntas respondidas
  - Formato de horas HH:MM válido
  - Hora término > hora inicio

**Flujo:**
```
Usuario doble clic en cronograma
    ↓
Se abre frmChecklistVirtual con datos pre-cargados
    ↓
Usuario completa inspección (responde preguntas)
    ↓
Click [Guardar]
    ↓
Sistema valida → calcula → guarda → actualiza cronograma
    ↓
Mensaje éxito + certificado disponible
```

**Referencias:**
- [TODO_CHECKLIST_VIRTUAL.md](TODO_CHECKLIST_VIRTUAL.md) - Diseño completo
- frmChecklistVirtual.frm
- ChecklistOrchestrator.bas
- ChecklistValidator.bas

### 4.3 Cálculo Automático de Scoring TA

**Módulo:** `InspectionCalculator.bas`

**Fórmula TA Porcentaje:**
```
TA Puntaje Obtenido = Σ(valores de respuestas "Sí")
TA Puntos Máximos = 57 (base configurada)
TA Puntos No Aplica = Cantidad de "No Aplica" × 4

TA Porcentaje = (TA Puntaje Obtenido × 100) / (TA Puntos Máximos - TA Puntos No Aplica)

Ejemplo:
  Puntaje: 45
  Máximos: 57
  No Aplica: 8 (2 preguntas × 4)
  TA% = (45 × 100) / (57 - 8) = 91.84%
```

**Cálculo RPN (1ra inspección):**
```
RPN = TA Porcentaje
```

**Categorización:**
```vba
If RPN <= 14 Then Categoria = 1
ElseIf RPN <= 19 Then Categoria = 2
ElseIf RPN <= 40 Then Categoria = 3
ElseIf RPN > 40 Then Categoria = 4
ElseIf [3 inspecciones consecutivas RPN > 20] Then Categoria = 5
```

**Referencias:**
- [SISTEMA_INSPECCIONES_ARQUITECTURA.md](SISTEMA_INSPECCIONES_ARQUITECTURA.md) - Sección 7: Fórmulas
- InspectionCalculator.bas

### 4.4 Inspecciones Recurrentes (RPN Ajustado)

**Estado:** ⏸️ En desarrollo (ver plan detallado)

**Concepto:**
- 1ra inspección: RPN = TA Porcentaje
- 2da+ inspección: RPN Total = RPN Promedio + factores microbiología

**Controles nuevos en frmChecklistVirtual:**
- CheckBox: "Esta NO es la primera inspección"
- Botón: "Buscar inspecciones previas"
- Campo: Número de inspección (auto-calculado)
- Campo: RPN anterior (automático o manual)
- Info: Resumen de inspecciones previas

**Cálculo RPN Recurrente:**
```
RPN Actual = TA Porcentaje (calculado normal)
RPN Promedio = (RPN Anterior + RPN Actual) / 2
RPN Total = RPN Promedio + %Recuperación + %OOL  (futuro)

Categorización basada en RPN Total
```

**Módulos nuevos:**
- `InspectionHistoryService.bas` - ✅ Creado (búsqueda histórico)
- `RecurrentInspectionCalculator.bas` - ✅ Creado (cálculos RPN recurrente)

**Modificaciones a módulos existentes:**
- ChecklistOrchestrator.bas - ✅ Integración flujo recurrente
- InspectionRepository.bas - ✅ Guardado de columnas 32-40

**Pendiente:**
- Migración de BD (agregar columnas 32-40)
- Testing completo
- Actualización de certificados PDF

**Referencias:**
- [TODO_RPN_AJUSTADO.md](TODO_RPN_AJUSTADO.md) - Plan completo paso a paso
- [INSTRUCCIONES_CONTROLES_RECURRENTES.md](INSTRUCCIONES_CONTROLES_RECURRENTES.md)

### 4.5 Generación de Certificados PDF

**Estado:** ✅ MVP Implementado

**Módulos:**
- `CertificadoPDFGenerator.bas` - Lógica de generación
- `PlantillaCertificadoSetup.bas` - Configuración de plantilla
- Hoja oculta: "Plantilla Certificado"

**¿Qué incluye el certificado?**
- Encabezado con logo y título
- Datos de inspección (fecha, evaluado, puesto, planta, área, línea, evaluador)
- Resultados generales (TA puntaje, porcentaje, RPN, categoría)
- Preguntas y respuestas por sección
- Observaciones generales
- Firmas (Evaluado, Evaluador, Supervisor)
- Metadatos (fecha emisión, ID inspección)

**Nombre de archivo:**
```
CERTIFICADO_[Iniciales]_[FechaISO]_[UUID_corto].pdf
Ejemplo: CERTIFICADO_ABC_20260420_x8fL.pdf
```

**Cómo generar:**
1. Menú Principal → Botón "Generar Certificado PDF"
2. Se abre frmSelectorInspeccion (modo PDF)
3. Filtrar inspecciones completadas
4. Seleccionar inspección
5. Click "Generar PDF"
6. PDF se guarda en escritorio y se abre automáticamente

**Referencias:**
- [TODO_CERTIFICADO_PDF.md](TODO_CERTIFICADO_PDF.md) - Plan de implementación
- [PLAN_MEJORA_CERTIFICADO_MVP.md](PLAN_MEJORA_CERTIFICADO_MVP.md) - Mejoras futuras
- CertificadoPDFGenerator.bas
- PlantillaCertificadoSetup.bas

### 4.6 Gestor de Tablas Maestras

**Módulo:** `TableManager.bas`, Formulario: `frmGestorTablas.frm`

**¿Qué permite gestionar?**
- tblCriticidad (niveles de criticidad)
- tblSecciones (secciones de inspección)
- tblPlantillas (plantillas de checklist)
- tblOpcionesDeRespuesta (opciones de respuesta)
- tblPreguntas (preguntas por plantilla)

**Operaciones CRUD:**
- **Crear:** Agregar nuevos registros
- **Leer:** Listar registros existentes
- **Actualizar:** Modificar registros
- **Eliminar:** Borrar registros (con validaciones de integridad)

**Validaciones:**
- No eliminar si tiene dependencias (ej: plantilla con inspecciones)
- UUIDs únicos
- Campos obligatorios completos
- Relaciones FK válidas

**Referencias:**
- [GESTOR_TABLAS_MANUAL.md](GESTOR_TABLAS_MANUAL.md) - Manual completo de uso
- TableManager.bas
- frmGestorTablas.frm

---

## 5. Navegación y Seguridad

### 5.1 Sistema de Navegación (xlSheetVeryHidden)

**Módulo:** `NavigationService2.bas`, `SheetService2.bas`

**¿Cómo funciona?**
- Todas las hojas excepto "Menú principal" están ocultas (`xlSheetVeryHidden`)
- Usuario no puede hacer visible manualmente (requiere VBA)
- Navegación SOLO a través de botones del menú
- Sistema controla qué hojas son accesibles según rol

**Estados de visibilidad:**
```vba
xlSheetVisible       ' Visible (solo Menú principal)
xlSheetHidden        ' Oculta (accesible si usuario sabe)
xlSheetVeryHidden    ' Muy oculta (requiere VBA)
```

**Flujo de navegación:**
```
Usuario en Menú principal
    ↓
Click en botón (ej: "Cronograma")
    ↓
NavigationService2.ActivateSheet("Cronograma")
    ↓
Sistema hace visible temporalmente la hoja
    ↓
Usuario trabaja en la hoja
    ↓
Al salir → Sistema oculta hoja nuevamente
```

**Referencias:**
- [SISTEMA_NAVEGACION_VISIBILIDAD.md](SISTEMA_NAVEGACION_VISIBILIDAD.md) - Guía completa
- NavigationService2.bas
- SheetService2.bas

### 5.2 Sistema de Protecciones (Dual-Layer)

**Capa 1: Protección de Estructura del Libro (URS-22)**

**Módulo:** `WorkbookProtector2.bas`

**¿Qué protege?**
- No insertar hojas nuevas
- No eliminar hojas existentes
- No cambiar nombre de hojas
- No mover/copiar hojas
- No mostrar/ocultar hojas

**Código:**
```vba
ThisWorkbook.Protect Password:="1234", Structure:=True, Windows:=False
```

**Referencias:**
- [URS-22_PROTECCION_ESTRUCTURA_LIBRO.md](URS-22_PROTECCION_ESTRUCTURA_LIBRO.md)
- WorkbookProtector2.bas

**Capa 2: Protección de Contenido de Hojas (URS-20/21)**

**Módulo:** `SheetProtector2.bas`

**¿Qué protege?**
- Celdas específicas (solo desbloquea celdas de entrada)
- Fórmulas (bloqueadas, usuario no puede modificar)
- Estructura de tablas (headers, columnas)
- Formato condicional

**Control booleano:**
```vba
Public Const ENABLE_SHEET_PROTECTION As Boolean = True
Public Const ENABLE_WORKBOOK_PROTECTION As Boolean = True
```

**Zonas típicas desbloqueadas:**
- Celdas de entrada de datos en formularios
- Columnas de respuestas en inspecciones
- Campos editables en cronogramas

**Referencias:**
- [GUIA_PROTECCIONES_SISTEMA.md](GUIA_PROTECCIONES_SISTEMA.md)
- SheetProtector2.bas

### 5.3 Control de Acceso (RBAC)

**Módulo:** `AdminAccessControl2.bas`, `UserManager2.bas`

**Roles y permisos:**
```vba
' Administrador:
If EsAdministrador(usuario) Then
    ' Acceso total
    ' Configuración, gestión de tablas, auditoría
End If

' Evaluador:
If EsEvaluador(usuario) Then
    ' Lectura + Inspecciones
    ' Realizar inspecciones, generar certificados
End If

' Consulta:
' Solo lectura
' Ver cronogramas, históricos
```

**Validación de acceso:**
```vba
Public Function TieneAcceso(usuario As String, accion As String) As Boolean
    ' Verifica permisos según matriz de roles
End Function
```

**Referencias:**
- AdminAccessControl2.bas
- UserManager2.bas

---

## 6. Cronogramas e Inspecciones

### 6.1 Inicialización del Cronograma

**Módulo:** `InspectionScheduler.bas`

**Función:** `InicializarCronogramaCompleto()`

**¿Qué hace?**
1. Lee tblPersonal (todos los activos)
2. Para cada persona:
   - Identifica puestos (columnas 3-13 con "Si")
   - Para cada puesto:
     - Busca plantilla correspondiente
     - Busca última inspección en tblInspecciones
     - Calcula próxima fecha según frecuencia
     - Calcula días de vencimiento
     - Determina estado (Vigente/Por vencer/Vencido)
     - Inserta/actualiza registro en tblCronogramaInspecciones

**Casos especiales:**
- Si no hay inspección previa → Estado: "Sin inspección inicial"
- Si plantilla no existe → Log de error
- Si persona inactiva → Omitir

**Referencias:**
- InspectionScheduler.bas función InicializarCronogramaCompleto()
- CronogramaButtons.bas btnInicializar_Click()

### 6.2 Flujo Completo de Inspección

**Pipeline de 12 pasos:**

```
PASO 1: Usuario doble clic en cronograma
    ↓
PASO 2: ChecklistOrchestrator.AbrirChecklistVirtual()
    ↓
PASO 3: frmChecklistVirtual se abre con datos pre-cargados
    ↓
PASO 4: Usuario completa inspección
    ↓
PASO 5: Click [Guardar]
    ↓
PASO 6: ChecklistValidator.ValidarFormularioCompleto()
    ↓
PASO 7: InspectionRepository.CrearInspeccion()
    ↓
PASO 8: AnswerRepository.GuardarRespuestas()
    ↓
PASO 9: InspectionCalculator.CalcularScoringTA()
    ↓
PASO 10: InspectionCalculator.CalcularRPN() + DeterminarCategoria()
    ↓
PASO 11: InspectionRepository.ActualizarCalculos()
    ↓
PASO 12: InspectionScheduler.ActualizarCronograma()
    ↓
PASO 13: AuditLogger2.Log() + Mensaje éxito
```

**Módulos involucrados:**
- ChecklistOrchestrator.bas (orquestación)
- ChecklistValidator.bas (validaciones)
- InspectionRepository.bas (persistencia)
- AnswerRepository.bas (respuestas)
- InspectionCalculator.bas (cálculos)
- InspectionScheduler.bas (actualización cronograma)
- AuditLogger2.bas (auditoría)

**Referencias:**
- [TODO_CHECKLIST_VIRTUAL.md](TODO_CHECKLIST_VIRTUAL.md) - Fase 5-7
- [TODO_FORMULARIO.md](TODO_FORMULARIO.md) - Pipeline completo
- ChecklistOrchestrator.bas función GuardarInspeccionCompleta()

### 6.3 Actualización del Cronograma Post-Inspección

**Función:** `InspectionScheduler.ActualizarCronogramaDespuesInspeccion()`

**¿Qué actualiza?**
```
tblCronogramaInspecciones:
  [06] Total inspecciones += 1
  [07] Fecha ultima inspeccion = FechaInspeccionActual
  [08] ID Ultima inspeccion = UUID
  [10] RPN ultima inspeccion = RPN calculado
  [11] Categoria ultima inspeccion = Categoría
  [14] Fecha proxima inspeccion = FechaUltima + Frecuencia
  [15] Dias para vencimiento = FechaProxima - HOY()
  [16] Estado programacion = Vigente/Por vencer/Vencido
  [24] Fecha ultima actualizacion = NOW()
```

**Recálculo de estado:**
```vba
If DiasVencimiento < 0 Then Estado = "Vencido"
ElseIf DiasVencimiento <= DiasAlerta Then Estado = "Por vencer"
Else Estado = "Vigente"
```

**Referencias:**
- InspectionScheduler.bas función ActualizarCronogramaDespuesInspeccion()
- ChecklistOrchestrator.bas (llamada en PASO 12)

---

## 7. Gestión de Datos

### 7.1 Repositorios de Datos

**Patrón de diseño:** Repository Pattern (Clean Architecture)

**InspectionRepository.bas**
- `CrearInspeccion()` - Inserta nueva inspección en tblInspecciones
- `ActualizarCalculos()` - Actualiza scoring, RPN, categoría
- `ObtenerUltimasNInspecciones()` - Histórico para Cat 5
- `ExisteInspeccion()` - Validación de existencia

**ChecklistRepository.bas**
- `ObtenerPlantillaPorPuestoYEtapa()` - Busca plantilla
- `ObtenerPreguntasPorPlantilla()` - Lista preguntas
- `ObtenerOpcionesRespuesta()` - Opciones por sección
- `ObtenerPersonalPorPuestoYPlanta()` - Personal filtrado

**AnswerRepository.bas**
- `GuardarRespuestas()` - Inserta respuestas en batch
- `ObtenerOpcionesPorSeccion()` - Opciones de respuesta
- `ObtenerIDOpcionPorTexto()` - Búsqueda de ID

**Características comunes:**
- Uso de nombres de columna (no índices hardcodeados)
- Manejo robusto de errores con ErrorLogger2
- Queries optimizadas
- Compatibilidad retroactiva

**Referencias:**
- [TODO_FORMULARIO.md](TODO_FORMULARIO.md) - Fase 1: Repositorios
- [TODO_CHECKLIST_VIRTUAL.md](TODO_CHECKLIST_VIRTUAL.md) - Fase 5: Repositorios
- InspectionRepository.bas
- ChecklistRepository.bas

### 7.2 Validadores

**ChecklistValidator.bas**
- `ValidarCabecera()` - Valida campos obligatorios del formulario
- `ValidarRespuestasCompletas()` - Verifica que todas las preguntas tengan respuesta
- `ValidarCoherenciaFechas()` - Valida lógica de fechas y horas
- `ValidarTodo()` - Validación completa pre-guardado

**TableValidator.bas**
- `ValidarIntegridadReferencial()` - Valida FKs antes de eliminar
- `ValidarUUIDUnico()` - Verifica unicidad de IDs
- `ValidarCamposObligatorios()` - Valida campos requeridos

**Ejemplo de validación:**
```vba
Public Function ValidarCabecera(frm As frmChecklistVirtual) As String
    If Trim(frm.txtFecha) = "" Then
        ValidarCabecera = "Fecha de inspección es obligatoria"
        Exit Function
    End If
    
    If Not IsDate(frm.txtFecha) Then
        ValidarCabecera = "Fecha de inspección inválida"
        Exit Function
    End If
    
    If CDate(frm.txtFecha) > Date Then
        ValidarCabecera = "Fecha de inspección no puede ser futura"
        Exit Function
    End If
    
    ' ... más validaciones ...
    
    ValidarCabecera = "" ' OK
End Function
```

**Referencias:**
- [TODO_FORMULARIO.md](TODO_FORMULARIO.md) - Fase 2: Validadores
- ChecklistValidator.bas
- TableValidator.bas

### 7.3 Calculadoras

**InspectionCalculator.bas**
- `CalcularScoringTA()` - Calcula puntaje, máximos, no aplica, porcentaje TA
- `CalcularRPN()` - Calcula RPN (actualmente = TA Porcentaje)
- `DeterminarCategoria()` - Categoriza según RPN y validación Cat 5

**RecurrentInspectionCalculator.bas** (⚠️ Nuevo - En desarrollo)
- `CalcularRPNPromedio()` - (RPN Anterior + RPN Actual) / 2
- `CalcularRPNTotal()` - RPN Promedio + %Recuperación + %OOL
- `DeterminarCategoriaRPNTotal()` - Categoría basada en RPN Total
- `ValidarConsistenciaRPN()` - Detecta cambios bruscos (>50%)

**Referencias:**
- [SISTEMA_INSPECCIONES_ARQUITECTURA.md](SISTEMA_INSPECCIONES_ARQUITECTURA.md) - Sección 7
- [TODO_RPN_AJUSTADO.md](TODO_RPN_AJUSTADO.md) - Fase 4
- InspectionCalculator.bas
- RecurrentInspectionCalculator.bas

---

## 8. Auditoría y Trazabilidad

### 8.1 Sistema de Audit Trail (5 Hojas)

**Arquitectura:**
- 5 hojas: "Audit trail 1" a "Audit trail 5"
- Cada hoja: Capacidad de 1,000,000 de filas
- Capacidad total: 5,000,000 de registros
- Rotación automática entre hojas

**Columnas de auditoría:**
| # | Columna | Tipo | Descripción |
|---|---------|------|-------------|
| 1 | ID | String | UUID único del registro |
| 2 | Fecha hora | DateTime | Timestamp de la acción |
| 3 | Usuario | String | Nombre del usuario (Windows) |
| 4 | Accion | String | Tipo de acción (INSERT, UPDATE, DELETE) |
| 5 | Hoja | String | Hoja afectada |
| 6 | Tabla | String | Tabla afectada |
| 7 | ID Registro | String | ID del registro modificado |
| 8 | Datos modificados | String | Campo modificado |
| 9 | Valor anterior | String | Valor antes del cambio |
| 10 | Valor nuevo | String | Valor después del cambio |
| 11 | Modulo y subrutina | String | Módulo.Función que ejecutó |

**Módulo:** `AuditLogger2.bas`

**Funciones principales:**
```vba
Public Sub LogAction( _
    action As String, _
    sheetName As String, _
    tableName As String, _
    recordID As String, _
    dataModified As String, _
    beforeChange As Variant, _
    afterChange As Variant, _
    moduleAndSubroutine As String _
)
```

**Referencias:**
- [AUDIT_TRAIL_CONFIGURACION.md](AUDIT_TRAIL_CONFIGURACION.md) - Guía completa
- AuditLogger2.bas
- AuditRotation2.bas

### 8.2 Rotación Automática de Audit Trail

**Módulo:** `AuditRotation2.bas`, `mod_AuditRotation.bas`

**¿Cómo funciona?**
1. AuditLogger2 detecta que hoja actual está llena (>=1M registros)
2. Llama a AuditRotation2.RotateToNextSheet()
3. Sistema activa próxima hoja disponible
4. Si todas llenas → vuelve a hoja 1 (overwrite)
5. Logs detallados en Debug.Print

**Función de búsqueda de próxima hoja:**
```vba
Public Function ObtenerProximaHojaDisponible() As Long
    ' Retorna número de hoja con menos registros
    ' O hoja 1 si todas llenas
End Function
```

**Testing:**
```vba
Public Sub TEST_GenerarRegistrosAudit(cantidad As Long)
    ' Genera N registros de prueba
    ' Útil para testing de rotación
End Sub
```

**Referencias:**
- [AUDIT_TRAIL_CONFIGURACION.md](AUDIT_TRAIL_CONFIGURACION.md) - Sección 3 y 4
- AuditRotation2.bas
- mod_AuditRotation.bas

### 8.3 Registro de Errores

**Módulo:** `ErrorLogger2.bas`

**Tabla:** `tblErrores` (Hoja Configuración)

**Columnas:**
- ID (UUID)
- Fecha hora
- Usuario
- Módulo
- Función
- Descripción error
- Número error (Err.Number)
- Línea (si disponible)

**Función principal:**
```vba
Public Sub Log( _
    moduleName As String, _
    errorDescription As String, _
    errorNumber As Long _
)
```

**Uso típico:**
```vba
Public Sub MiFuncion()
    On Error GoTo ErrorHandler
    
    ' ... código ...
    
    Exit Sub
    
ErrorHandler:
    Call ErrorLogger2.Log("MiModulo.MiFuncion", Err.Description, Err.Number)
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
```

**Referencias:**
- ErrorLogger2.bas
- Configuration2.bas (definición de tblErrores)

### 8.4 Trazabilidad de Inspecciones

**Registro completo de trazabilidad:**

**1. Creación de inspección:**
```
AuditLogger2.LogAction(
    action:="INSERT_INSPECCION",
    recordID:=idInspeccion,
    moduleAndSubroutine:="ChecklistOrchestrator.GuardarInspeccionCompleta"
)
```

**2. Guardado de respuestas:**
```
AuditLogger2.LogAction(
    action:="INSERT_RESPUESTAS",
    recordID:=idInspeccion,
    dataModified:="N respuestas guardadas"
)
```

**3. Actualización de cálculos:**
```
AuditLogger2.LogAction(
    action:="UPDATE_CALCULOS",
    recordID:=idInspeccion,
    dataModified:="RPN=" & rpn & ", Cat=" & categoria
)
```

**4. Actualización de cronograma:**
```
AuditLogger2.LogAction(
    action:="UPDATE_CRONOGRAMA",
    recordID:=idCronograma,
    dataModified:="Proxima inspeccion=" & fechaProxima
)
```

**Búsqueda de auditoría por inspección:**
```vba
' En hoja Audit trail, filtrar por:
Column("ID Registro") = [UUID de inspección]
```

**Referencias:**
- ChecklistOrchestrator.bas (múltiples llamadas a AuditLogger2)
- [TODO_CHECKLIST_VIRTUAL.md](TODO_CHECKLIST_VIRTUAL.md) - Fase 7, paso 14

---

## 9. Certificados y Reportes

### 9.1 Certificado PDF - Estructura

**Plantilla:** Hoja oculta "Plantilla Certificado"

**Secciones:**

**1. Encabezado:**
- Logo (espacio reservado 60×60px)
- Título "CERTIFICADO DE INSPECCIÓN"
- Subtítulo "TÉCNICA ASÉPTICA"
- Proyecto "TH-HC-001"

**2. Datos de inspección:**
- Fecha inspección, hora inicio-fin
- Evaluado (nombre + iniciales)
- Puesto, Planta
- Área, Línea auditada
- Lugar auditoría
- Evaluador
- Personal línea (AY1, AY2, OP)

**3. Resultados generales:**
- TA Puntaje obtenido / máximos
- TA Puntos no aplica
- TA Porcentaje
- RPN
- Categoría resultado (con descripción)

**4. Preguntas y respuestas:**
- Sección: Auditoría de Procesos
  - Tabla con #, Pregunta, Respuesta, Observación
- Sección: Técnica Aséptica
  - Tabla con #, Pregunta, Respuesta, Observación

**5. Observaciones generales:**
- Campo de texto libre

**6. Firmas:**
- Evaluado: _______________
- Evaluador: _______________
- Supervisor: _______________

**7. Metadatos:**
- Fecha emisión certificado
- ID Inspección (UUID)
- "Generado por: Sistema TH-HC-001 v1.0"

**Referencias:**
- [TODO_CERTIFICADO_PDF.md](TODO_CERTIFICADO_PDF.md) - Diseño completo
- PlantillaCertificadoSetup.bas

### 9.2 Generación del PDF - Pipeline

**Pasos:**

**1. Usuario selecciona inspección:**
```
Menú Principal → Botón "Generar Certificado PDF"
    ↓
frmSelectorInspeccion (modo PDF)
    ↓
Filtrar inspecciones con Estado="Completado"
    ↓
Seleccionar inspección
    ↓
Click [Generar PDF]
```

**2. Limpiar plantilla:**
```vba
Call CertificadoPDFGenerator.LimpiarPlantillaCertificado()
```

**3. Poblar plantilla con datos:**
```vba
' Cabecera
Call PoblarCabeceraInspeccion(datosInspeccion)

' Resultados
Call PoblarResultadosGenerales(datosInspeccion)

' Preguntas
Call PoblarPreguntasRespuestas(idInspeccion)

' Observaciones
Call PoblarObservaciones(observacionGeneral)
```

**4. Aplicar formato final:**
```vba
Call AplicarFormatoFinal()
```

**5. Generar nombre de archivo:**
```vba
nombrePDF = GenerarNombreArchivoPDF(iniciales, fecha, uuid)
' Resultado: CERTIFICADO_ABC_20260420_x8fL.pdf
```

**6. Exportar a PDF:**
```vba
wsPlantilla.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=rutaCompleta, _
    Quality:=xlQualityStandard, _
    OpenAfterPublish:=True
```

**7. Limpiar plantilla nuevamente**

**8. Registrar en auditoría**

**Referencias:**
- [TODO_CERTIFICADO_PDF.md](TODO_CERTIFICADO_PDF.md) - Fase 2: Módulo generador
- CertificadoPDFGenerator.bas función GenerarCertificadoPDF()

### 9.3 Mejoras Futuras del Certificado (MVP)

**Plan:** [PLAN_MEJORA_CERTIFICADO_MVP.md](PLAN_MEJORA_CERTIFICADO_MVP.md)

**5 cambios de alto impacto:**

**1. Bloque categoría arriba** ⭐⭐⭐⭐⭐
- Celda grande combinada A6:G8
- Texto: "CATEGORÍA N - Descripción"
- Fondo color según categoría:
  - Cat 1-2: Verde claro
  - Cat 3: Amarillo claro
  - Cat 4-5: Rojo claro

**2. Fila "Estado" en resultados** ⭐⭐⭐⭐
- "✓ COMPETENTE" (Cat 1-2, verde)
- "⚠ COMPETENTE CON OBSERVACIONES" (Cat 3, naranja)
- "✗ NO CALIFICADO" (Cat 4-5, rojo)

**3. Resaltar incumplimientos en rojo** ⭐⭐⭐⭐
- Respuestas "No" o "No Cumple" con fondo rojo claro
- Texto en negrita
- Opcional: ícono ❌

**4. Footer con validez** ⭐⭐⭐
- "VÁLIDO HASTA: [Fecha]"
- Cálculo: Fecha inspección + Frecuencia meses
- Advertencia si vencido: "⚠ VENCIDO" en rojo

**5. Nombre de archivo inteligente** ⭐⭐⭐
```
Actual: CERTIFICADO_ABC_20260420_x8fL.pdf
Nuevo:  CERTIFICADO_ACF_OperadorNPT_2026-04-20_CAT1_APROBADO.pdf
```

**Estimado total:** 60-90 minutos

**Referencias:**
- [PLAN_MEJORA_CERTIFICADO_MVP.md](PLAN_MEJORA_CERTIFICADO_MVP.md) - Plan paso a paso

---

## 10. Desarrollo y Mantenimiento

### 10.1 Estructura de Módulos VBA

**Por funcionalidad:**

**Gestión de Inspecciones:**
- ChecklistOrchestrator.bas - Orquestación del flujo
- ChecklistRepository.bas - Lectura de datos de checklist
- ChecklistValidator.bas - Validaciones de formulario
- InspectionRepository.bas - CRUD de inspecciones
- InspectionCalculator.bas - Cálculos TA, RPN, categoría
- InspectionScheduler.bas - Gestión de cronogramas

**Inspecciones Recurrentes (nuevo):**
- InspectionHistoryService.bas - Búsqueda de histórico
- RecurrentInspectionCalculator.bas - Cálculos RPN recurrente

**Gestión de Respuestas:**
- AnswerRepository.bas - CRUD de respuestas

**Navegación y UI:**
- NavigationService2.bas - Control de navegación
- SheetService2.bas - Manejo de visibilidad de hojas
- CronogramaButtons.bas - Botones del cronograma
- CronogramaResumen.bas - Resumen en menú

**Seguridad:**
- WorkbookProtector2.bas - Protección de estructura
- SheetProtector2.bas - Protección de hojas
- AdminAccessControl2.bas - Control de acceso
- UserManager2.bas - Gestión de usuarios

**Auditoría:**
- AuditLogger2.bas - Registro de acciones
- AuditRotation2.bas - Rotación de hojas audit
- ErrorLogger2.bas - Registro de errores

**Certificados:**
- CertificadoPDFGenerator.bas - Generación de PDF
- PlantillaCertificadoSetup.bas - Configuración plantilla

**Gestión de Datos:**
- TableManager.bas - CRUD de tablas maestras
- TableValidator.bas - Validaciones de integridad
- TableAuditor2.bas - Auditoría de cambios en tablas
- TableDiagnostics.bas - Diagnóstico de tablas

**Configuración:**
- Configuration2.bas - Constantes centrales
- SystemInitializer.bas - Inicialización del sistema

**Utilidades:**
- VariablesGlobales2.bas - Variables compartidas
- mod_BackupManager.bas - Gestión de backups
- mod_PasswordMigration.bas - Migración de contraseñas
- DevModeHelper.bas - Herramientas de desarrollo

**Formularios:**
- frmChecklistVirtual.frm - Formulario de inspección
- frmSelectorInspeccion.frm - Selector de inspecciones
- frmGestorTablas.frm - Gestor de tablas maestras
- frmInput.frm - Input genérico

### 10.2 Principios de Arquitectura

**Clean Architecture:**
```
Presentación (UI)
    ↓
Aplicación (Use Cases, Orchestrators)
    ↓
Dominio (Business Logic, Calculators, Validators)
    ↓
Infraestructura (Repositories, Data Access)
    ↓
Infraestructura Compartida (Config, Security, Audit)
```

**SOLID:**
- **S**ingle Responsibility: Cada módulo una responsabilidad
- **O**pen/Closed: Extensión sin modificación
- **L**iskov Substitution: Interfaces consistentes
- **I**nterface Segregation: Interfaces específicas
- **D**ependency Inversion: Dependencias hacia arriba

**DRY (Don't Repeat Yourself):**
- Funciones reutilizables (Configuration2, Utilities)
- Constantes centralizadas (Configuration2.bas)
- Patrones de diseño (Repository, Orchestrator)

**Referencias:**
- [TODO_FORMULARIO.md](TODO_FORMULARIO.md) - Arquitectura completa
- [TODO_CHECKLIST_VIRTUAL.md](TODO_CHECKLIST_VIRTUAL.md) - Arquitectura limpia

### 10.3 Guías de Desarrollo

**Agregar nueva columna a tblInspecciones:**
1. Agregar columna en Excel (después de columna 40)
2. Actualizar Configuration2.bas con documentación
3. Modificar InspectionRepository.CrearInspeccion() si aplica
4. Modificar InspectionRepository.ActualizarCalculos() si aplica
5. Actualizar CertificadoPDFGenerator si debe mostrarse en PDF
6. Testing completo

**Referencia:** [INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md](INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md)

**Agregar nueva columna con auditoría de procesos:**
1. Agregar columna después de "RPN calculado"
2. Actualizar TableAuditor2.bas para auditar cambios
3. Modificar ChecklistOrchestrator para capturar datos
4. Actualizar certificado PDF si aplica

**Referencia:** [INSTRUCCIONES_COLUMNA_AUDITORIA_PROCESOS.md](INSTRUCCIONES_COLUMNA_AUDITORIA_PROCESOS.md)

**Agregar nueva plantilla de checklist:**
1. Usar frmGestorTablas → Pestaña "Plantillas"
2. Click [Agregar nueva plantilla]
3. Llenar: Nombre, Puesto, Etapa, Frecuencia meses
4. Guardar
5. Agregar preguntas en pestaña "Preguntas"
6. Asociar a secciones existentes
7. Testing con inspección de prueba

**Referencia:** [GESTOR_TABLAS_MANUAL.md](GESTOR_TABLAS_MANUAL.md)

**Agregar controles de inspección recurrente al formulario:**
1. Abrir VBA Editor (Alt+F11)
2. Buscar frmChecklistVirtual en explorador de proyecto
3. Abrir diseñador (F7)
4. Agregar controles según lista en INSTRUCCIONES_CONTROLES_RECURRENTES.md
5. No configurar propiedades (se hace por código)
6. Guardar formulario
7. Ejecutar ConfigurarCabecera() para aplicar estilos

**Referencia:** [INSTRUCCIONES_CONTROLES_RECURRENTES.md](INSTRUCCIONES_CONTROLES_RECURRENTES.md)

**Setup de plantilla de certificado:**
1. Seguir paso a paso INSTRUCCIONES_SETUP_PLANTILLA.md
2. Crear/verificar hoja "Plantilla Certificado"
3. Configurar PageSetup (A4, márgenes, orientación)
4. Ejecutar PlantillaCertificadoSetup.InicializarPlantillaCertificado()
5. Verificar estructura visual
6. Testing con generación de PDF

**Referencia:** [INSTRUCCIONES_SETUP_PLANTILLA.md](INSTRUCCIONES_SETUP_PLANTILLA.md)

### 10.4 Debugging y Diagnóstico

**Debug de auditoría de procesos con criticidad:**
1. Seguir guía [DEBUG_AUDITORIA_PROCESOS.md](DEBUG_AUDITORIA_PROCESOS.md)
2. Verificar columnas en tblRespuestas
3. Verificar tblCriticidad tiene valores válidos
4. Usar Debug.Print en InspectionCalculator
5. Revisar logs de ErrorLogger2

**Diagnóstico de tablas:**
```vba
' En Ventana Inmediato
Call TableDiagnostics.DiagnosticarTodasLasTablas()
```

**Output:**
- Existencia de tablas
- Número de registros
- Columnas faltantes
- Registros duplicados
- FKs rotas

**Referencias:**
- [DEBUG_AUDITORIA_PROCESOS.md](DEBUG_AUDITORIA_PROCESOS.md)
- TableDiagnostics.bas

### 10.5 Backups y Recuperación

**Backup automático:**
```vba
' En mod_BackupManager.bas
Public Sub CrearBackupAutomatico()
    Dim rutaBackup As String
    rutaBackup = ThisWorkbook.Path & "\backups\" & _
                 "TH-HC-001_" & Format(Now, "YYYYMMDD_HHMM") & ".xlsm"
    
    ThisWorkbook.SaveCopyAs rutaBackup
End Sub
```

**Backup manual:**
1. Crear carpeta `backups/` en raíz del proyecto
2. Ejecutar `mod_BackupManager.CrearBackupAutomatico()`
3. Verificar archivo en carpeta backups

**Restauración:**
1. Cerrar archivo actual
2. Copiar archivo de backup desde `backups/`
3. Renombrar a `TH-HC-001 INSPECCIONES.xlsm`
4. Abrir y verificar

**Referencias:**
- mod_BackupManager.bas

---

## 11. Planes Futuros

### 11.1 Inspecciones Recurrentes (RPN Ajustado) - En Desarrollo

**Estado:** 80% completado

**Completado:**
- ✅ Análisis de estructura de BD
- ✅ Diseño de 9 columnas nuevas (32-40)
- ✅ UI controles en frmChecklistVirtual
- ✅ InspectionHistoryService.bas creado
- ✅ RecurrentInspectionCalculator.bas creado
- ✅ ChecklistOrchestrator integración completa
- ✅ InspectionRepository guardado de columnas nuevas

**Pendiente:**
- ⏸️ Migración de BD (agregar columnas 32-40 a tblInspecciones)
- ⏸️ Testing completo con datos reales
- ⏸️ Actualización de certificados PDF para mostrar desglose
- ⏸️ Documentación final

**Próximo paso:** Migración de BD

**Referencias:**
- [TODO_RPN_AJUSTADO.md](TODO_RPN_AJUSTADO.md) - Plan completo

### 11.2 Mejoras del Certificado PDF (MVP)

**Prioridad:** Media  
**Estimado:** 60-90 minutos

**Cambios planificados:**
1. Bloque categoría arriba con color (20 min)
2. Fila "Estado" en resultados (10 min)
3. Resaltar incumplimientos en rojo (20 min)
4. Footer con validez (15 min)
5. Nombre archivo inteligente (10 min)

**Referencias:**
- [PLAN_MEJORA_CERTIFICADO_MVP.md](PLAN_MEJORA_CERTIFICADO_MVP.md)

### 11.3 Checklist Virtual Completo

**Estado:** Planificación

**Objetivos:**
1. Cronograma resumen en Menú Principal
2. Formulario UserForm con pestañas (en lugar de hoja)
3. Validaciones en tiempo real
4. Guardado transaccional completo

**Referencias:**
- [TODO_CHECKLIST_VIRTUAL.md](TODO_CHECKLIST_VIRTUAL.md) - Plan completo

### 11.4 Formulario Dinámico en Hoja Excel

**Estado:** Planificación

**Alternativa a UserForm:**
- Hoja "Formulario Inspeccion"
- Preguntas cargadas dinámicamente
- Data Validation por pregunta
- Botones: Cargar, Guardar, Limpiar, Cancelar

**Referencias:**
- [TODO_FORMULARIO.md](TODO_FORMULARIO.md) - Arquitectura completa

### 11.5 Datos de Microbiología (Futuro)

**Concepto:**
- % Recuperación (0-100)
- % OOL (Out of Limits, 0-100)
- Integración en RPN Total

**Fórmula futura:**
```
RPN Total = RPN Promedio + %Recuperación + %OOL
```

**Preparación actual:**
- Columnas [38] y [39] reservadas en tblInspecciones
- RecurrentInspectionCalculator.CalcularRPNTotal() ya acepta parámetros
- Actualmente valores = 0 (no se usan)

**Referencias:**
- [TODO_RPN_AJUSTADO.md](TODO_RPN_AJUSTADO.md) - Q3: Microbiología futura

---

## 12. Referencias Rápidas

### 12.1 Archivos de Documentación

| Archivo | Propósito | Prioridad |
|---------|-----------|-----------|
| SISTEMA_INSPECCIONES_ARQUITECTURA.md | Base de datos y relaciones | ⭐⭐⭐⭐⭐ |
| GESTOR_TABLAS_MANUAL.md | Uso del gestor de tablas | ⭐⭐⭐⭐⭐ |
| AUDIT_TRAIL_CONFIGURACION.md | Sistema de auditoría | ⭐⭐⭐⭐⭐ |
| GUIA_PROTECCIONES_SISTEMA.md | Seguridad y protecciones | ⭐⭐⭐⭐⭐ |
| SISTEMA_NAVEGACION_VISIBILIDAD.md | Navegación xlSheetVeryHidden | ⭐⭐⭐⭐ |
| URS-22_PROTECCION_ESTRUCTURA_LIBRO.md | URS-22 detallado | ⭐⭐⭐⭐ |
| TODO_RPN_AJUSTADO.md | Plan inspecciones recurrentes | ⭐⭐⭐⭐ |
| TODO_CERTIFICADO_PDF.md | Plan certificados PDF | ⭐⭐⭐⭐ |
| PLAN_MEJORA_CERTIFICADO_MVP.md | Mejoras certificado | ⭐⭐⭐ |
| TODO_CHECKLIST_VIRTUAL.md | Plan checklist UserForm | ⭐⭐⭐ |
| TODO_FORMULARIO.md | Plan formulario hoja Excel | ⭐⭐⭐ |
| INSTRUCCIONES_SETUP_PLANTILLA.md | Setup plantilla certificado | ⭐⭐⭐ |
| INSTRUCCIONES_CONTROLES_RECURRENTES.md | Agregar controles UI | ⭐⭐⭐ |
| INSTRUCCIONES_COLUMNA_ID_CRITICIDAD.md | Agregar columna criticidad | ⭐⭐ |
| INSTRUCCIONES_COLUMNA_AUDITORIA_PROCESOS.md | Agregar columna auditoría | ⭐⭐ |
| DEBUG_AUDITORIA_PROCESOS.md | Debug auditoría con criticidad | ⭐⭐ |
| COMO_AGREGAR_CONTROLES_RECURRENTES.md | Guía controles recurrentes | ⭐⭐ |

### 12.2 Módulos VBA Críticos

| Módulo | Función Principal | Criticidad |
|--------|-------------------|------------|
| Configuration2.bas | Constantes centrales | ⭐⭐⭐⭐⭐ |
| ChecklistOrchestrator.bas | Pipeline de guardado | ⭐⭐⭐⭐⭐ |
| InspectionRepository.bas | CRUD inspecciones | ⭐⭐⭐⭐⭐ |
| InspectionCalculator.bas | Cálculos RPN | ⭐⭐⭐⭐⭐ |
| AuditLogger2.bas | Auditoría | ⭐⭐⭐⭐⭐ |
| WorkbookProtector2.bas | Protección estructura | ⭐⭐⭐⭐⭐ |
| SheetProtector2.bas | Protección hojas | ⭐⭐⭐⭐⭐ |
| NavigationService2.bas | Navegación | ⭐⭐⭐⭐ |
| InspectionScheduler.bas | Cronogramas | ⭐⭐⭐⭐ |
| CertificadoPDFGenerator.bas | PDF | ⭐⭐⭐⭐ |
| TableManager.bas | Gestión tablas | ⭐⭐⭐⭐ |
| ChecklistValidator.bas | Validaciones | ⭐⭐⭐ |
| InspectionHistoryService.bas | Histórico (nuevo) | ⭐⭐⭐ |
| RecurrentInspectionCalculator.bas | RPN recurrente (nuevo) | ⭐⭐⭐ |

### 12.3 Tablas Principales

| Tabla | Hoja | Registros Típicos | Criticidad |
|-------|------|-------------------|------------|
| tblInspecciones | Historico | 1000+ | ⭐⭐⭐⭐⭐ |
| tblRespuestas | Historico | 20000+ | ⭐⭐⭐⭐⭐ |
| tblCronogramaInspecciones | Cronograma | 100-200 | ⭐⭐⭐⭐⭐ |
| tblPersonal | Personal | 50-100 | ⭐⭐⭐⭐⭐ |
| tblPlantillas | Checklist | 10-20 | ⭐⭐⭐⭐ |
| tblPreguntas | Checklist | 200-300 | ⭐⭐⭐⭐ |
| tblCriticidad | Configuración | 4 | ⭐⭐⭐⭐ |
| tblCategoriasRPN | Configuración | 5 | ⭐⭐⭐⭐ |
| tblSecciones | Configuración | 2-5 | ⭐⭐⭐ |
| tblOpcionesDeRespuesta | Configuración | 10-15 | ⭐⭐⭐ |

### 12.4 Fórmulas Clave

**TA Porcentaje:**
```
TA% = (TA Puntaje Obtenido × 100) / (TA Puntos Máximos - TA Puntos No Aplica)
```

**RPN (1ra inspección):**
```
RPN = TA Porcentaje
```

**RPN Promedio (2da+ inspección):**
```
RPN Promedio = (RPN Anterior + RPN Actual) / 2
```

**RPN Total (futuro completo):**
```
RPN Total = RPN Promedio + %Recuperación + %OOL
```

**Categorización:**
```
Cat 1: RPN 0-14
Cat 2: RPN 15-19
Cat 3: RPN 20-40
Cat 4: RPN 40.01-999
Cat 5: 3 inspecciones consecutivas con RPN > 20
```

**Días de Vencimiento:**
```
Días = Fecha Próxima Inspección - HOY()
```

**Estado Cronograma:**
```
If Días < 0 Then "Vencido"
ElseIf Días <= Días Alerta Then "Por vencer"
Else "Vigente"
```

### 12.5 Contraseñas y Seguridad

**Contraseña del sistema:**
```vba
Configuration2.APP_PASSWORD = "1234"
```

**Usada en:**
- Protección de estructura del libro
- Protección de hojas individuales
- Desprotección temporal para cambios administrativos

**⚠️ IMPORTANTE:** Cambiar contraseña en producción:
1. Modificar Configuration2.APP_PASSWORD
2. Volver a proteger todas las hojas
3. Volver a proteger estructura del libro
4. Testing completo

### 12.6 Contactos y Soporte

**Sistema desarrollado por:** [Equipo TH-HC-001]  
**Documentación mantenida por:** [Responsable documentación]  
**Fecha última revisión:** 21/04/2026

---

## 📝 Notas Finales

Este documento consolida toda la información de los 17 archivos de documentación del sistema TH-HC-001.

**Uso recomendado:**
- **Para crear instructivos de usuario:** Consultar secciones 4-9
- **Para desarrollo:** Consultar secciones 2, 10
- **Para troubleshooting:** Consultar sección 12
- **Para auditoría:** Consultar sección 8
- **Para planificación:** Consultar sección 11

**Mantenimiento del documento:**
- Actualizar al agregar nueva funcionalidad
- Revisar trimestralmente
- Sincronizar con archivos fuente en docs/

---

**FIN DE LA GUÍA MAESTRA**
