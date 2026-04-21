# 📋 TODO: SISTEMA DE CERTIFICADOS PDF DE INSPECCIÓN

**Proyecto:** TH-HC-001 EN PROCESO DE VALIDACIÓN  
**Módulo:** Generación de certificados PDF de inspecciones completadas  
**Fecha creación:** 15/04/2026  
**Estado:** PLANIFICACIÓN  
**Prioridad:** Media

---

## 🎯 OBJETIVO GENERAL

Implementar un sistema completo de generación de certificados PDF para inspecciones de Técnica Aséptica, que permita:

1. **Generar certificados profesionales en formato PDF** con toda la información de una inspección completada (cabecera, preguntas, respuestas, cálculos, observaciones y firmas)
2. **Exportar sin dependencias externas** usando la funcionalidad nativa de Excel (ExportAsFixedFormat)
3. **Guardar automáticamente en el escritorio** del usuario con nomenclatura consistente
4. **Optimizar diseño para 1 página** (con letra pequeña y formato compacto) o permitir 2 páginas si es necesario
5. **Invocar desde Menú Principal** mediante botón dedicado

---

## 📐 ARQUITECTURA GENERAL

```
┌────────────────────────────────────────────────────────────────┐
│  MENÚ PRINCIPAL (Hoja1)                                        │
│  ┌──────────────────────────────────────────┐                  │
│  │ [Botón: Generar Certificado PDF]         │                  │
│  │      ↓                                    │                  │
│  │ Abre frmSelectorInspeccion (modo PDF)    │                  │
│  └──────────────────────────────────────────┘                  │
└────────────────────────────────────────────────────────────────┘
                       │
                       ▼
┌────────────────────────────────────────────────────────────────┐
│  frmSelectorInspeccion (modo = "PDF")                          │
│  ┌──────────────────────────────────────────┐                  │
│  │ Lista de inspecciones completadas        │                  │
│  │ Filtro por: Planta, Personal, Fecha      │                  │
│  │ Columnas: Fecha | Iniciales | Puesto     │                  │
│  │           | Planta | RPN | Categoría     │                  │
│  │ ══════════════════════════════════════    │                  │
│  │ [Generar PDF] (botón activo solo si      │                  │
│  │               hay fila seleccionada)      │                  │
│  └──────────────────────────────────────────┘                  │
└────────────────────────────────────────────────────────────────┘
                       │
                       ▼
┌────────────────────────────────────────────────────────────────┐
│  CertificadoPDFGenerator.bas                                   │
│  ┌──────────────────────────────────────────┐                  │
│  │ 1. Leer datos de inspección completos    │                  │
│  │    (tblInspecciones + tblRespuestas)     │                  │
│  │ 2. Poblar hoja "Plantilla Certificado"   │                  │
│  │ 3. Aplicar formato y ajustar para PDF    │                  │
│  │ 4. Exportar como PDF con nombre:         │                  │
│  │    CERTIFICADO_[Iniciales]_[FechaISO]_   │                  │
│  │    [UUID_corto].pdf                      │                  │
│  │ 5. Guardar en escritorio                 │                  │
│  │ 6. Limpiar plantilla                     │                  │
│  │ 7. Abrir PDF automáticamente (opcional)  │                  │
│  └──────────────────────────────────────────┘                  │
└────────────────────────────────────────────────────────────────┘
                       │
                       ▼
┌────────────────────────────────────────────────────────────────┐
│  HOJA OCULTA: "Plantilla Certificado"                         │
│  ┌──────────────────────────────────────────┐                  │
│  │  ┌────────────────────────────────────┐  │                  │
│  │  │ [ESPACIO LOGO]  CERTIFICADO DE     │  │                  │
│  │  │                 INSPECCIÓN          │  │                  │
│  │  │                 TÉCNICA ASÉPTICA    │  │                  │
│  │  ├────────────────────────────────────┤  │                  │
│  │  │ DATOS DE INSPECCIÓN                │  │                  │
│  │  │ Fecha: [X]  Hora: [Y]-[Z]          │  │                  │
│  │  │ Evaluado: [Nombre] ([Iniciales])   │  │                  │
│  │  │ Puesto: [X]  Planta: [Y]           │  │                  │
│  │  │ Área: [X]  Línea: [Y]              │  │                  │
│  │  │ Evaluador: [X]  Lugar: [Y]         │  │                  │
│  │  │ AY1: [X]  AY2: [Y]  OP: [Z]        │  │                  │
│  │  ├────────────────────────────────────┤  │                  │
│  │  │ RESULTADOS GENERALES               │  │                  │
│  │  │ TA Puntaje: [X] / [Y]              │  │                  │
│  │  │ TA Porcentaje: [Z]%                │  │                  │
│  │  │ RPN: [X]  Categoría: [Y]           │  │                  │
│  │  ├────────────────────────────────────┤  │                  │
│  │  │ SECCIÓN: AUDITORÍA DE PROCESOS     │  │                  │
│  │  │ # | Pregunta | Respuesta | Obs     │  │                  │
│  │  │ 1 | [texto]  | [opción]  | [texto] │  │                  │
│  │  │ ...                                 │  │                  │
│  │  ├────────────────────────────────────┤  │                  │
│  │  │ SECCIÓN: TÉCNICA ASÉPTICA          │  │                  │
│  │  │ # | Pregunta | Respuesta | Obs     │  │                  │
│  │  │ 1 | [texto]  | [opción]  | [texto] │  │                  │
│  │  │ ...                                 │  │                  │
│  │  ├────────────────────────────────────┤  │                  │
│  │  │ OBSERVACIONES GENERALES:           │  │                  │
│  │  │ [texto observación]                │  │                  │
│  │  ├────────────────────────────────────┤  │                  │
│  │  │ FIRMAS                             │  │                  │
│  │  │ Evaluado:______ Evaluador:______   │  │                  │
│  │  │ Supervisor:_______                 │  │                  │
│  │  │                                     │  │                  │
│  │  │ Fecha emisión: [DD/MM/YYYY HH:MM]  │  │                  │
│  │  │ ID Inspección: [UUID]              │  │                  │
│  │  └────────────────────────────────────┘  │                  │
│  └──────────────────────────────────────────┘                  │
└────────────────────────────────────────────────────────────────┘
                       │
                       ▼
┌────────────────────────────────────────────────────────────────┐
│  ESCRITORIO                                                    │
│  📄 CERTIFICADO_ABC_20260415_x8fL.pdf                          │
└────────────────────────────────────────────────────────────────┘
```

---

## 📋 DISEÑO DEL CERTIFICADO (LAYOUT)

### **Dimensiones y configuración de página**
- **Tamaño:** A4 (210mm × 297mm)
- **Orientación:** Retrato (Portrait)
- **Márgenes:** Superior: 1.5cm, Inferior: 1.5cm, Izquierdo: 1.5cm, Derecho: 1.5cm
- **Escala:** Ajustar al 100% preferiblemente, o "Ajustar a 1 página de ancho × N páginas de alto"
- **Encabezado/Pie:** Sin encabezados ni pies de página adicionales

### **Secciones del certificado (de arriba a abajo)**

#### **1. ENCABEZADO (filas 1-4)**
```
┌─────────────────────────────────────────────────────────────┐
│  [LOGO]              CERTIFICADO DE INSPECCIÓN              │
│  (60×60px)           TÉCNICA ASÉPTICA                       │
│                                                              │
│                      PROYECTO: TH-HC-001                     │
└─────────────────────────────────────────────────────────────┘
```
- **Logo:** Celda A1:B4 (combinada) — espacio reservado para logo (imagen 60×60px)
- **Título:** Celda C1:G2 (combinada) — "CERTIFICADO DE INSPECCIÓN" — Arial 16pt, Negrita, Centrado
- **Subtítulo:** Celda C2:G3 (combinada) — "TÉCNICA ASÉPTICA" — Arial 14pt, Negrita, Centrado
- **Proyecto:** Celda C3:G4 (combinada) — "PROYECTO: TH-HC-001" — Arial 10pt, Centrado
- **Fondo:** Gris claro (#F2F2F2)

#### **2. DATOS DE INSPECCIÓN (filas 6-13)**
```
┌─────────────────────────────────────────────────────────────┐
│  DATOS DE INSPECCIÓN                                        │
├─────────────────────────────────────────────────────────────┤
│  Fecha inspección:  [DD/MM/YYYY]     Hora: [HH:MM] - [HH:MM]│
│  Evaluado:          [Nombre Apellido] ([ABC])               │
│  Puesto:            [Operador]       Planta: [Planta A]     │
│  Área:              [Manufactura]    Línea: [Sala Blanca 1] │
│  Lugar auditoría:   [Dentro del área]                       │
│  Evaluador:         [Nombre Apellido] ([XYZ])               │
│  Personal línea:    AY1: [AAA]  AY2: [BBB]  OP: [CCC]       │
└─────────────────────────────────────────────────────────────┘
```
- **Título sección:** Fila 6, Celdas A6:G6 (combinadas) — Arial 11pt, Negrita, Fondo azul claro (#D6EAF8)
- **Campos:** Filas 7-13, 2 columnas (etiqueta + valor) — Arial 9pt
- **Formato:** Etiquetas en negrita, valores normales

#### **3. RESULTADOS GENERALES (filas 15-18)**
```
┌─────────────────────────────────────────────────────────────┐
│  RESULTADOS GENERALES                                       │
├─────────────────────────────────────────────────────────────┤
│  TA Puntaje obtenido:  [45]  /  [57]  (máximos)             │
│  TA Puntos no aplica:  [8]                                  │
│  TA Porcentaje:        [91.84]%                             │
│  RPN:                  [91.84]                              │
│  Categoría resultado:  [1 - Desempeño óptimo]              │
└─────────────────────────────────────────────────────────────┘
```
- **Título sección:** Fila 15, A15:G15 — Arial 11pt, Negrita, Fondo verde claro (#D5F4E6)
- **Campos:** Filas 16-18, formato tabla — Arial 9pt
- **Destacado:** Categoría en negrita con color según nivel

#### **4. SECCIÓN PREGUNTAS: AUDITORÍA DE PROCESOS (filas 20-X)**
```
┌────┬─────────────────────────────────────┬──────────┬─────────┐
│ #  │ PREGUNTA                            │ RESPUESTA│ OBSERV. │
├────┼─────────────────────────────────────┼──────────┼─────────┤
│ 1  │ [Texto pregunta completo...]        │ Cumple   │ [Obs]   │
│ 2  │ [Texto pregunta completo...]        │ No Cumple│ [Obs]   │
│... │ ...                                 │ ...      │ ...     │
└────┴─────────────────────────────────────┴──────────┴─────────┘
```
- **Título sección:** Fila 20, A20:G20 — Arial 11pt, Negrita, Fondo naranja claro (#FCE5CD)
- **Encabezado tabla:** Fila 21 — Arial 8pt, Negrita, Bordes
- **Anchos columna:**
  - Col A (Nº): 4%
  - Col B (Pregunta): 65%
  - Col C (Respuesta): 15%
  - Col D (Observación): 16%
- **Filas datos:** Arial 7pt, ajuste de altura automático, bordes finos
- **Alternado:** Filas pares con fondo gris muy claro (#FAFAFA)

#### **5. SECCIÓN PREGUNTAS: TÉCNICA ASÉPTICA (filas X+2 - Y)**
```
┌────┬─────────────────────────────────────┬──────────┬─────────┐
│ #  │ PREGUNTA                            │ RESPUESTA│ OBSERV. │
├────┼─────────────────────────────────────┼──────────┼─────────┤
│ 1  │ [Texto pregunta completo...]        │ Sí       │ [Obs]   │
│ 2  │ [Texto pregunta completo...]        │ No       │ [Obs]   │
│... │ ...                                 │ ...      │ ...     │
└────┴─────────────────────────────────────┴──────────┴─────────┘
```
- Igual formato que sección anterior, pero fondo celeste (#D6EAF8)

#### **6. OBSERVACIONES GENERALES (filas Y+2 - Y+5)**
```
┌─────────────────────────────────────────────────────────────┐
│  OBSERVACIONES GENERALES                                    │
├─────────────────────────────────────────────────────────────┤
│  [Texto de observación general de la inspección, puede     │
│   ocupar varias líneas con ajuste automático]               │
└─────────────────────────────────────────────────────────────┘
```
- **Título:** Fila Y+2, A:G — Arial 10pt, Negrita
- **Contenido:** Filas Y+3:Y+5, A:G (combinadas) — Arial 8pt, ajuste texto

#### **7. FIRMAS (filas Y+7 - Y+12)**
```
┌─────────────────────────────────────────────────────────────┐
│  FIRMAS Y VALIDACIÓN                                        │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  Evaluado:                  Evaluador:                      │
│  _________________          _________________               │
│  [Nombre]                   [Nombre]                        │
│                                                              │
│  Supervisor/Responsable:                                    │
│  _________________                                          │
│  [Nombre cargo]                                             │
│                                                              │
├─────────────────────────────────────────────────────────────┤
│  Fecha emisión certificado: [DD/MM/YYYY HH:MM:SS]           │
│  ID Inspección: [UUID completo]                             │
│  Generado por: Sistema TH-HC-001 v1.0                       │
└─────────────────────────────────────────────────────────────┘
```
- **Líneas de firma:** 3 espacios con línea inferior (formato subrayado en celda)
- **Metadatos:** Arial 7pt, gris (#666666), al pie

---

## 📊 ESTRUCTURA DE DATOS

### **Datos de entrada para generar certificado**
El certificado se genera a partir de un **ID Inspección** (UUID) que debe existir en `tblInspecciones` con estado **"Completada"**.

#### **Tablas a consultar**
1. **tblInspecciones** — Datos de cabecera y cálculos
2. **tblRespuestas** — Respuestas individuales por pregunta
3. **tblPreguntas** — Textos de preguntas y orden
4. **tblSecciones** — Nombres de secciones
5. **tblOpcionesDeRespuesta** — Textos de opciones seleccionadas
6. **tblPersonal** — Nombre completo del evaluado
7. **tblAseguramientoCalidad** — Nombre completo del evaluador (opcional)

#### **Campos requeridos del certificado**

| Sección | Campo | Fuente | Observación |
|---------|-------|--------|-------------|
| **Cabecera** | Fecha inspección | tblInspecciones.`Fecha inspeccion` | Formato DD/MM/YYYY |
| | Hora inicio | tblInspecciones.`Hora inicio` | Formato HH:MM |
| | Hora término | tblInspecciones.`Hora termino` | Formato HH:MM |
| | Evaluado (nombre) | JOIN tblPersonal.`Nombre` + `Apellido` | Concatenar |
| | Evaluado (iniciales) | tblInspecciones.`Personal` | Entre paréntesis |
| | Puesto | tblInspecciones.`Puesto` | Texto completo |
| | Planta | tblInspecciones.`Planta` | Texto completo |
| | Área | tblInspecciones.`Area` | Texto completo |
| | Línea auditada | tblInspecciones.`Linea Auditada` | Texto completo |
| | Lugar auditoría | tblInspecciones.`Lugar Auditoria` | Dentro/Fuera del área |
| | Evaluador | tblInspecciones.`Auditor` | Iniciales o nombre completo |
| | AY1 | tblInspecciones.`Iniciales AY1` | |
| | AY2 | tblInspecciones.`Iniciales AY2` | |
| | OP | tblInspecciones.`Iniciales OP` | |
| **Resultados** | TA Puntaje | tblInspecciones.`TA puntaje obtenido` | Número |
| | TA Máximos | tblInspecciones.`TA puntos maximos` | Número |
| | TA No aplica | tblInspecciones.`TA puntos no aplica` | Número |
| | TA Porcentaje | tblInspecciones.`TA porcentaje` | Formato ##.## |
| | RPN | tblInspecciones.`RPN` | Formato ##.## |
| | Categoría | tblInspecciones.`Numero categoria` | Número (1-5) |
| | Descripción categoría | JOIN tblCategoriasRPN.`Nombre categoria` | Texto largo |
| **Preguntas** | Número orden | tblPreguntas.`Numero pregunta` | Ordenar ASC |
| | Texto pregunta | tblPreguntas.`Pregunta texto` | Texto largo |
| | Sección | JOIN tblSecciones.`Nombre de sección` | Agrupar |
| | Respuesta | JOIN tblOpcionesDeRespuesta.`Opción texto` | Sí/No/Cumple/etc. |
| | Observación | tblRespuestas.`Observacion` | Puede estar vacío |
| **Observaciones** | Observación general | tblInspecciones.`Observacion general` | Texto largo |
| **Metadatos** | Fecha emisión | =NOW() | Al generar PDF |
| | ID Inspección | tblInspecciones.`ID Inspeccion` | UUID completo |

---

## 🔧 ARQUITECTURA TÉCNICA

### **Componentes del sistema**

#### **1. Hoja oculta: "Plantilla Certificado"**
- **Propósito:** Hoja de trabajo formateada para exportación PDF
- **Estado:** Hoja muy oculta (`xlSheetVeryHidden`)
- **Protección:** Sin protección (necesita escritura)
- **Formato:** Pre-diseñada con rangos nombrados para cada campo
- **Limpieza:** Limpiar contenido antes y después de cada generación

**Rangos nombrados sugeridos:**
```vba
rngCertLogo              → A1:B4
rngCertTitulo            → C1:G2
rngCertFechaInsp         → C7
rngCertHoraIni           → E7
rngCertHoraFin           → G7
rngCertEvaluadoNombre    → C8
rngCertEvaluadoIniciales → E8
rngCertPuesto            → C9
rngCertPlanta            → E9
... (definir todos)
rngCertTablaPreguntasIni → A21  (primer celda de tabla de preguntas)
rngCertObsGeneral        → A60:G62
rngCertFirmaEvaluado     → B65
rngCertFirmaEvaluador    → E65
rngCertFirmaSupervisor   → B68
rngCertFechaEmision      → E71
rngCertUUID              → E72
```

#### **2. Módulo nuevo: CertificadoPDFGenerator.bas**

**Funciones públicas principales:**
```vba
Public Function GenerarCertificadoPDF(ByVal idInspeccion As String) As Boolean
Public Function ObtenerRutaEscritorio() As String
Public Function ValidarInspeccionCompletada(ByVal idInspeccion As String) As Boolean
```

**Funciones privadas auxiliares:**
```vba
Private Sub LimpiarPlantillaCertificado()
Private Sub PoblarCabeceraInspeccion(ByVal datosInsp As Object)
Private Sub PoblarResultadosGenerales(ByVal datosInsp As Object)
Private Sub PoblarPreguntasRespuestas(ByVal idInspeccion As String)
Private Sub PoblarObservaciones(ByVal observacion As String)
Private Sub AplicarFormatoFinal()
Private Function GenerarNombreArchivoPDF(ByVal iniciales As String, _
                                          ByVal fecha As Date, _
                                          ByVal uuid As String) As String
Private Function ExportarHojaAPDF(ByVal wsPlantilla As Worksheet, _
                                   ByVal rutaCompleta As String) As Boolean
```

#### **3. Modificación: frmSelectorInspeccion**

**Cambios requeridos:**
- Agregar parámetro de modo al abrir: `Public ModoApertura As String` ("EDITAR" o "PDF")
- Si `ModoApertura = "PDF"`:
  - Cambiar título del formulario a "Seleccionar Inspección para Certificado PDF"
  - Filtrar solo inspecciones con estado "Completada"
  - Cambiar botón "Abrir" por "Generar PDF"
  - Al hacer clic:
    ```vba
    Private Sub btnGenerarPDF_Click()
        Dim idInsp As String
        idInsp = Me.lstInspecciones.List(Me.lstInspecciones.ListIndex, COL_ID_OCULTA)
        
        If CertificadoPDFGenerator.GenerarCertificadoPDF(idInsp) Then
            MsgBox "Certificado PDF generado exitosamente en el escritorio.", vbInformation
            Unload Me
        Else
            MsgBox "Error al generar el certificado PDF. Revise los logs.", vbCritical
        End If
    End Sub
    ```

#### **4. Modificación: Hoja1 (Menú Principal)**

**Agregar botón:**
- **Texto:** "Generar Certificado PDF"
- **Posición:** Al lado de los botones existentes
- **Código asignado:**
```vba
Public Sub btnGenerarCertificadoPDF_Click()
    ' Abrir selector en modo PDF
    Dim frm As frmSelectorInspeccion
    Set frm = New frmSelectorInspeccion
    frm.ModoApertura = "PDF"
    frm.Show vbModal
    Set frm = Nothing
End Sub
```

#### **5. Modificación: Configuration2.bas**

**Nuevas constantes:**
```vba
' Hoja plantilla certificado
Public Const SHEET_PLANTILLA_CERTIFICADO As String = "Plantilla Certificado"

' Configuración PDF
Public Const PDF_PREFIJO_NOMBRE As String = "CERTIFICADO"
Public Const PDF_CALIDAD As Long = 0  ' 0 = Standard, 1 = Minimum
Public Const PDF_ABRIR_AUTOMATICO As Boolean = True  ' Abrir PDF tras generarlo
```

---

## ✅ PLAN DE IMPLEMENTACIÓN

### **CONTROL DE VERSIONES Y COMMITS**

Cada fase debe completarse en su totalidad antes de proceder a la siguiente. Al finalizar cada fase:

1. **Commit descriptivo** con mensaje estructurado:
   ```
   feat(certificado-pdf): [FASE X] - Descripción breve
   
   - Detalle 1
   - Detalle 2
   - Detalle 3
   ```

2. **Tag de versión** (versionado semántico):
   - Fase 0: `v1.4.0-cert.0` (preparación)
   - Fase 1: `v1.4.0-cert.1` (plantilla Excel)
   - Fase 2: `v1.4.0-cert.2` (módulo generador)
   - Fase 3: `v1.4.0-cert.3` (integración UI)
   - Fase 4: `v1.4.1` (testing y release)

3. **Backup del archivo `.xlsm`** en carpeta `backups/`:
   ```
   backups/TH-HC-001_CERT_FASE_X_YYYYMMDD_HHMM.xlsm
   ```

---

### **FASE 0: PREPARACIÓN DE INFRAESTRUCTURA** 🔧

**Duración estimada:** 30 minutos  
**Objetivo:** Configurar constantes, agregar columnas faltantes y preparar estructura base.

#### **Tareas**

- [ ] **0.1 — Verificar columnas en tblInspecciones**
  - Abrir Excel → Hoja Historico → tblInspecciones
  - Verificar existencia de columnas documentadas en sección "Estructura de datos"
  - Si faltan columnas de la cabecera (Area, Hora inicio, Hora termino, etc.), agregarlas manualmente
  - Actualizar documentación en Configuration2.bas (sección de documentación de tblInspecciones)

- [ ] **0.2 — Agregar constantes a Configuration2.bas**
  ```vba
  ' ══════════════════════════════════════════════════════════════
  ' CERTIFICADOS PDF
  ' ══════════════════════════════════════════════════════════════
  Public Const SHEET_PLANTILLA_CERTIFICADO As String = "Plantilla Certificado"
  Public Const PDF_PREFIJO_NOMBRE As String = "CERTIFICADO"
  Public Const PDF_CALIDAD As Long = 0  ' 0 = xlQualityStandard
  Public Const PDF_ABRIR_AUTOMATICO As Boolean = True
  
  ' Constantes auxiliares
  Public Const CERT_ANCHO_COL_NUMERO As Double = 20
  Public Const CERT_ANCHO_COL_PREGUNTA As Double = 350
  Public Const CERT_ANCHO_COL_RESPUESTA As Double = 80
  Public Const CERT_ANCHO_COL_OBSERVACION As Double = 90
  ```

- [ ] **0.3 — Crear carpeta backups/ si no existe**
  - Crear carpeta `backups/` en la raíz del proyecto
  - Agregar `backups/` al `.gitignore` (si se usa Git)

- [ ] **0.4 — Respaldar archivo actual**
  - Guardar copia del archivo: `backups/TH-HC-001_PRE_CERT_20260415.xlsm`

#### **Commit y Tag**
```bash
git add Configuration2.bas
git commit -m "feat(certificado-pdf): [FASE 0] - Preparación de infraestructura

- Agregar constantes para certificados PDF en Configuration2
- Verificar columnas requeridas en tblInspecciones
- Crear carpeta de backups"

git tag v1.4.0-cert.0
```

---

### **FASE 1: CREACIÓN DE PLANTILLA EXCEL** 📄

**Duración estimada:** 2-3 horas  
**Objetivo:** Diseñar y formatear la hoja "Plantilla Certificado" con layout profesional.

#### **Tareas**

- [ ] **1.1 — Crear hoja "Plantilla Certificado"**
  - Crear nueva hoja con nombre exacto `Plantilla Certificado`
  - Configurar como `xlSheetVeryHidden` (muy oculta)
  - Configuración de página:
    ```vba
    With wsPlantilla.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False  ' Permitir múltiples páginas si es necesario
        .PrintGridlines = False
        .PrintHeadings = False
    End With
    ```

- [ ] **1.2 — Configurar anchos de columna**
  ```vba
  With wsPlantilla
      .Columns("A:A").ColumnWidth = 3      ' Margen izquierdo
      .Columns("B:B").ColumnWidth = 12     ' Columna 1 (etiquetas)
      .Columns("C:C").ColumnWidth = 20     ' Columna 2 (valores)
      .Columns("D:D").ColumnWidth = 12     ' Columna 3 (etiquetas)
      .Columns("E:E").ColumnWidth = 20     ' Columna 4 (valores)
      .Columns("F:F").ColumnWidth = 12     ' Columna 5 (extra)
      .Columns("G:G").ColumnWidth = 20     ' Columna 6 (valores)
      .Columns("H:H").ColumnWidth = 3      ' Margen derecho
  End With
  ```

- [ ] **1.3 — Diseñar encabezado (filas 1-4)**
  - Combinar A1:B4 (espacio para logo futuro)
  - Insertar placeholder: "LOGO" centrado, Arial 8pt gris
  - Combinar C1:G2, escribir "CERTIFICADO DE INSPECCIÓN"
    - Arial 16pt, Negrita, Centrado verticalmente
  - Combinar C2:G3, escribir "TÉCNICA ASÉPTICA"
    - Arial 14pt, Negrita, Centrado
  - Combinar C3:G4, escribir "PROYECTO: TH-HC-001"
    - Arial 10pt, Centrado
  - Aplicar fondo gris claro (#F2F2F2) a toda la sección A1:G4
  - Bordes exteriores gruesos

- [ ] **1.4 — Diseñar sección "Datos de Inspección" (filas 6-14)**
  - Fila 6: Título sección (combinar A6:G6)
    - Texto: "DATOS DE INSPECCIÓN"
    - Arial 11pt, Negrita, Fondo azul claro (#D6EAF8)
  - Filas 7-14: Labels y placeholders de valores
    - Fila 7: `B7="Fecha inspección:"` `C7="[FECHA]"` `D7="Hora:"` `E7="[INICIO]"` `F7="-"` `G7="[FIN]"`
    - Fila 8: `B8="Evaluado:"` `C8:E8="[NOMBRE COMPLETO]"` `F8="Iniciales:"` `G8="[ABC]"`
    - Fila 9: `B9="Puesto:"` `C9:D9="[PUESTO]"` `E9="Planta:"` `F9:G9="[PLANTA]"`
    - Fila 10: `B10="Área:"` `C10:D10="[ÁREA]"` `E10="Línea:"` `F10:G10="[LÍNEA]"`
    - Fila 11: `B11="Lugar auditoría:"` `C11:G11="[LUGAR]"`
    - Fila 12: `B12="Evaluador:"` `C12:G12="[EVALUADOR]"`
    - Fila 13: `B13="Personal línea:"` `C13="AY1:"` `D13="[---]"` `E13="AY2:"` `F13="[---]"` `G13="OP:"` `H13="[---]"` (mover a G13)
  - Formato:
    - Labels (columnas B, D, F): Arial 9pt, Negrita
    - Valores (columnas C, E, G): Arial 9pt, Regular
    - Bordes finos en toda la sección

- [ ] **1.5 — Diseñar sección "Resultados Generales" (filas 16-21)**
  - Fila 16: Título sección (combinar A16:G16)
    - Texto: "RESULTADOS GENERALES"
    - Arial 11pt, Negrita, Fondo verde claro (#D5F4E6)
  - Filas 17-21:
    - Fila 17: `B17="TA Puntaje obtenido:"` `C17="[X]"` `D17="/ máximos:"` `E17="[Y]"`
    - Fila 18: `B18="TA Puntos no aplica:"` `C18="[Z]"`
    - Fila 19: `B19="TA Porcentaje:"` `C19="[XX.XX]"` `D19="%"`
    - Fila 20: `B20="RPN:"` `C20="[XX.XX]"`
    - Fila 21: `B21="Categoría resultado:"` `C21:G21="[N - Descripción]"`
  - Formato: Igual que sección anterior

- [ ] **1.6 — Diseñar tabla de preguntas (encabezado en fila 23)**
  - Fila 23: Encabezado de tabla
    - `B23="#"` `C23="PREGUNTA"` `E23="RESPUESTA"` `F23="OBSERVACIÓN"`
    - Arial 8pt, Negrita, Centrado, Bordes gruesos
    - Fondo gris oscuro (#CCCCCC)
  - Fila 24 en adelante: Área dinámica (se llenará por código)
    - Formato predefinido para filas de datos:
      - Arial 7pt
      - Ajuste de texto activado en columna pregunta
      - Bordes finos

- [ ] **1.7 — Diseñar sección observaciones (placeholder)**
  - Se ubicará dinámicamente después de preguntas
  - Formato predefinido:
    - Título: Arial 10pt, Negrita
    - Contenido: Arial 8pt, ajuste texto

- [ ] **1.8 — Diseñar sección firmas (placeholder)**
  - Se ubicará al final
  - Predefinir rangos con líneas de subrayado

- [ ] **1.9 — Definir rangos nombrados**
  - Crear nombres de rango para cada campo editable
  - Usar `ThisWorkbook.Names.Add Name:="rngCertFechaInsp", RefersTo:="='Plantilla Certificado'!$C$7"`
  - Repetir para todos los campos documentados arriba

- [ ] **1.10 — Probar impresión manual**
  - Rellenar plantilla con datos de prueba manualmente
  - Vista preliminar → verificar que cabe en 1-2 páginas
  - Ajustar tamaños de fuente si es necesario
  - Exportar manualmente a PDF de prueba
  - Verificar calidad y legibilidad

#### **Commit y Tag**
```bash
git add "TH-HC-001 INSPECCIONES.xlsm"
git commit -m "feat(certificado-pdf): [FASE 1] - Plantilla Excel diseñada

- Crear hoja 'Plantilla Certificado' (muy oculta)
- Diseñar layout completo (encabezado, datos, resultados, tablas, firmas)
- Configurar PageSetup para exportación A4
- Definir rangos nombrados para población dinámica
- Probar exportación manual exitosa"

git tag v1.4.0-cert.1
git push origin v1.4.0-cert.1

# Backup
cp "TH-HC-001 INSPECCIONES.xlsm" "backups/TH-HC-001_CERT_FASE1_20260415.xlsm"
```

---

### **FASE 2: MÓDULO GENERADOR DE PDF** 🔨

**Duración estimada:** 3-4 horas  
**Objetivo:** Implementar módulo VBA `CertificadoPDFGenerator.bas` con toda la lógica de generación.

#### **Tareas**

- [ ] **2.1 — Crear módulo CertificadoPDFGenerator.bas**
  - Crear módulo estándar nuevo
  - Agregar encabezado de documentación:
  ```vba
  ' ══════════════════════════════════════════════════════════════
  ' Módulo: CertificadoPDFGenerator
  ' Descripción: Generación de certificados PDF de inspecciones
  '              completadas. Exporta desde hoja plantilla Excel.
  ' Fecha creación: 15/04/2026
  ' Dependencias: Configuration2, InspectionRepository,
  '               ChecklistRepository, ErrorLogger2, AuditLogger2
  ' ══════════════════════════════════════════════════════════════
  Option Explicit
  ```

- [ ] **2.2 — Implementar función principal: GenerarCertificadoPDF**
  ```vba
  Public Function GenerarCertificadoPDF(ByVal idInspeccion As String) As Boolean
      On Error GoTo ErrorHandler
      
      ' PASO 1: Validar que la inspección existe y está completada
      ' PASO 2: Obtener todos los datos necesarios
      ' PASO 3: Limpiar plantilla
      ' PASO 4: Poblar plantilla con datos
      ' PASO 5: Aplicar formato final
      ' PASO 6: Generar nombre de archivo
      ' PASO 7: Obtener ruta escritorio
      ' PASO 8: Exportar a PDF
      ' PASO 9: Limpiar plantilla nuevamente
      ' PASO 10: Registrar en audit log
      ' PASO 11: Abrir PDF (opcional)
      
      GenerarCertificadoPDF = True
      Exit Function
      
  ErrorHandler:
      GenerarCertificadoPDF = False
      Call ErrorLogger2.Log("CertificadoPDFGenerator.GenerarCertificadoPDF", _
                             Err.Description, Err.Number)
  End Function
  ```

- [ ] **2.3 — Implementar ValidarInspeccionCompletada**
  ```vba
  Public Function ValidarInspeccionCompletada(ByVal idInspeccion As String) As Boolean
      ' Buscar en tblInspecciones
      ' Verificar que existe
      ' Verificar que Estado = "Completada" (si existe columna Estado)
      ' Retornar True/False
  End Function
  ```

- [ ] **2.4 — Implementar ObtenerRutaEscritorio**
  ```vba
  Public Function ObtenerRutaEscritorio() As String
      ' Usar WScript.Shell para obtener ruta del escritorio
      Dim wsh As Object
      Set wsh = CreateObject("WScript.Shell")
      ObtenerRutaEscritorio = wsh.SpecialFolders("Desktop")
      Set wsh = Nothing
  End Function
  ```

- [ ] **2.5 — Implementar LimpiarPlantillaCertificado**
  ```vba
  Private Sub LimpiarPlantillaCertificado()
      ' Obtener referencia a hoja
      ' Limpiar solo celdas de datos (no formato ni estructura)
      ' Usar UsedRange.ClearContents en rangos específicos
      ' NO limpiar títulos, labels ni estructura
  End Function
  ```

- [ ] **2.6 — Implementar PoblarCabeceraInspeccion**
  ```vba
  Private Sub PoblarCabeceraInspeccion(ByVal datosInsp As Object)
      ' datosInsp = Dictionary con todos los campos de tblInspecciones
      ' Escribir en rangos nombrados:
      '   rngCertFechaInsp = Format(datosInsp("Fecha inspeccion"), "DD/MM/YYYY")
      '   rngCertHoraIni = datosInsp("Hora inicio")
      '   rngCertHoraFin = datosInsp("Hora termino")
      '   rngCertEvaluadoNombre = ObtenerNombreCompleto(datosInsp("Personal"))
      '   ... etc
  End Sub
  ```

- [ ] **2.7 — Implementar PoblarResultadosGenerales**
  ```vba
  Private Sub PoblarResultadosGenerales(ByVal datosInsp As Object)
      ' Escribir en rangos:
      '   rngCertTAPuntaje = datosInsp("TA puntaje obtenido")
      '   rngCertTAMaximos = datosInsp("TA puntos maximos")
      '   rngCertTANoAplica = datosInsp("TA puntos no aplica")
      '   rngCertTAPorcentaje = Format(datosInsp("TA porcentaje"), "##.##") & "%"
      '   rngCertRPN = Format(datosInsp("RPN"), "##.##")
      '   rngCertCategoria = ObtenerTextoCategoria(datosInsp("Numero categoria"))
  End Sub
  ```

- [ ] **2.8 — Implementar PoblarPreguntasRespuestas**
  ```vba
  Private Sub PoblarPreguntasRespuestas(ByVal idInspeccion As String)
      ' 1. Obtener todas las respuestas de tblRespuestas (JOIN con tblPreguntas)
      ' 2. Ordenar por sección y número de pregunta
      ' 3. Agrupar por sección
      ' 4. Para cada sección:
      '    a. Escribir título de sección
      '    b. Escribir encabezado de tabla
      '    c. Para cada pregunta:
      '       - Escribir número, pregunta, respuesta, observación
      '       - Aplicar formato alternado (filas pares con fondo gris)
      ' 5. Ajustar altura de filas automáticamente
  End Sub
  ```

- [ ] **2.9 — Implementar PoblarObservaciones**
  ```vba
  Private Sub PoblarObservaciones(ByVal observacion As String)
      ' Escribir observación general en rango correspondiente
      ' Si está vacío, escribir "Sin observaciones"
  End Sub
  ```

- [ ] **2.10 — Implementar AplicarFormatoFinal**
  ```vba
  Private Sub AplicarFormatoFinal()
      ' Ajustar altura de filas
      ' Aplicar bordes finales
      ' Ajustar zoom para vista preliminar (opcional)
      ' Verificar que no hay celdas con #REF o errores
  End Sub
  ```

- [ ] **2.11 — Implementar GenerarNombreArchivoPDF**
  ```vba
  Private Function GenerarNombreArchivoPDF(ByVal iniciales As String, _
                                            ByVal fecha As Date, _
                                            ByVal uuid As String) As String
      ' Formato: CERTIFICADO_ABC_20260415_x8fL.pdf
      ' UUID corto: tomar primeros 4 caracteres del UUID (antes del primer guión)
      Dim fechaISO As String
      Dim uuidCorto As String
      
      fechaISO = Format(fecha, "YYYYMMDD")
      uuidCorto = Left(uuid, 4)
      
      GenerarNombreArchivoPDF = Configuration2.PDF_PREFIJO_NOMBRE & "_" & _
                                 iniciales & "_" & _
                                 fechaISO & "_" & _
                                 uuidCorto & ".pdf"
  End Function
  ```

- [ ] **2.12 — Implementar ExportarHojaAPDF**
  ```vba
  Private Function ExportarHojaAPDF(ByVal wsPlantilla As Worksheet, _
                                     ByVal rutaCompleta As String) As Boolean
      On Error GoTo ErrorHandler
      
      ' Exportar usando ExportAsFixedFormat
      wsPlantilla.ExportAsFixedFormat _
          Type:=xlTypePDF, _
          Filename:=rutaCompleta, _
          Quality:=Configuration2.PDF_CALIDAD, _
          IncludeDocProperties:=True, _
          IgnorePrintAreas:=False, _
          OpenAfterPublish:=Configuration2.PDF_ABRIR_AUTOMATICO
      
      ExportarHojaAPDF = True
      Exit Function
      
  ErrorHandler:
      ExportarHojaAPDF = False
  End Function
  ```

- [ ] **2.13 — Funciones auxiliares**
  ```vba
  Private Function ObtenerNombreCompleto(ByVal iniciales As String) As String
      ' Buscar en tblPersonal, retornar Nombre + Apellido
  End Function
  
  Private Function ObtenerTextoCategoria(ByVal numCategoria As Long) As String
      ' Buscar en tblCategoriasRPN, retornar "N - Descripción"
  End Function
  ```

- [ ] **2.14 — Pruebas unitarias del módulo**
  - Crear inspección de prueba completa
  - Llamar `GenerarCertificadoPDF(idPrueba)` desde Immediate Window
  - Verificar que:
    - No hay errores
    - PDF se genera en escritorio
    - Todos los datos están presentes
    - Formato es correcto
    - PDF se abre automáticamente

#### **Commit y Tag**
```bash
git add CertificadoPDFGenerator.bas
git commit -m "feat(certificado-pdf): [FASE 2] - Módulo generador implementado

- Crear CertificadoPDFGenerator.bas con todas las funciones
- Implementar lógica de población de plantilla Excel
- Implementar exportación nativa a PDF (ExportAsFixedFormat)
- Agregar funciones auxiliares de nombres, rutas y validaciones
- Pruebas unitarias exitosas con inspección de prueba"

git tag v1.4.0-cert.2
git push origin v1.4.0-cert.2

# Backup
cp "TH-HC-001 INSPECCIONES.xlsm" "backups/TH-HC-001_CERT_FASE2_20260415.xlsm"
```

---

### **FASE 3: INTEGRACIÓN CON UI** 🖱️

**Duración estimada:** 1-2 horas  
**Objetivo:** Modificar formularios y menú principal para invocar generación de PDF.

#### **Tareas**

- [ ] **3.1 — Modificar frmSelectorInspeccion**
  
  **3.1.1 — Agregar propiedad ModoApertura**
  ```vba
  ' En la parte superior del código del formulario
  Public ModoApertura As String  ' "EDITAR" o "PDF"
  ```
  
  **3.1.2 — Modificar UserForm_Initialize**
  ```vba
  Private Sub UserForm_Initialize()
      ' Al inicio, verificar modo
      If ModoApertura = "" Then ModoApertura = "EDITAR"  ' Default
      
      ' Configurar según modo
      If ModoApertura = "PDF" Then
          Me.Caption = "Seleccionar Inspección para Certificado PDF"
          Me.btnAbrir.Caption = "Generar PDF"
      Else
          Me.Caption = "Seleccionar Inspección para Editar"
          Me.btnAbrir.Caption = "Abrir"
      End If
      
      ' Cargar inspecciones...
      ' (código existente)
  End Sub
  ```
  
  **3.1.3 — Modificar filtro de inspecciones**
  ```vba
  Private Sub CargarInspecciones()
      ' Si modo PDF, filtrar solo inspecciones completadas
      Dim estadoFiltro As String
      If ModoApertura = "PDF" Then
          estadoFiltro = "Completada"
      Else
          estadoFiltro = ""  ' Todas
      End If
      
      ' Aplicar filtro...
      ' (adaptar código existente)
  End Sub
  ```
  
  **3.1.4 — Modificar evento Click del botón**
  ```vba
  Private Sub btnAbrir_Click()
      If Me.lstInspecciones.ListIndex = -1 Then
          MsgBox "Seleccione una inspección.", vbExclamation
          Exit Sub
      End If
      
      Dim idInsp As String
      idInsp = Me.lstInspecciones.List(Me.lstInspecciones.ListIndex, COL_ID_OCULTA)
      
      If ModoApertura = "PDF" Then
          ' Generar PDF
          If CertificadoPDFGenerator.GenerarCertificadoPDF(idInsp) Then
              MsgBox "Certificado PDF generado exitosamente en el escritorio.", _
                     vbInformation, "PDF Generado"
              Unload Me
          Else
              MsgBox "Error al generar el certificado PDF. Revise los logs de error.", _
                     vbCritical, "Error"
          End If
      Else
          ' Abrir para editar (código existente)
          ' ...
      End If
  End Sub
  ```

- [ ] **3.2 — Modificar Hoja1 (Menú Principal)**
  
  **3.2.1 — Agregar botón en hoja Excel**
  - Insertar botón ActiveX o Shape con texto "Generar Certificado PDF"
  - Posicionar al lado de botones existentes
  - Asignar macro `Hoja1.btnGenerarCertificadoPDF_Click`
  
  **3.2.2 — Implementar código del botón en Hoja1.bas**
  ```vba
  Public Sub btnGenerarCertificadoPDF_Click()
      On Error GoTo ErrorHandler
      
      ' Desproteger hoja si está protegida
      Dim estabaProtegida As Boolean
      estabaProtegida = Me.ProtectContents
      If estabaProtegida Then Me.Unprotect Configuration2.SHEET_PASSWORD
      
      ' Abrir selector en modo PDF
      Dim frm As frmSelectorInspeccion
      Set frm = New frmSelectorInspeccion
      frm.ModoApertura = "PDF"
      frm.Show vbModal
      Set frm = Nothing
      
      ' Reproteger si estaba protegida
      If estabaProtegida Then Me.Protect Configuration2.SHEET_PASSWORD, _
                                           UserInterfaceOnly:=True
      Exit Sub
      
  ErrorHandler:
      MsgBox "Error al abrir el selector de inspecciones: " & Err.Description, _
             vbCritical
      Call ErrorLogger2.Log("Hoja1.btnGenerarCertificadoPDF_Click", _
                             Err.Description, Err.Number)
  End Sub
  ```

- [ ] **3.3 — Agregar registro en AuditLogger**
  - Modificar `GenerarCertificadoPDF` para registrar evento:
  ```vba
  Call AuditLogger2.Log("CERTIFICADO_PDF_GENERADO", _
                         "ID Inspección: " & idInspeccion & _
                         " | Personal: " & iniciales & _
                         " | Archivo: " & nombrePDF, _
                         Application.UserName)
  ```

- [ ] **3.4 — Pruebas de integración**
  - Abrir Menú Principal
  - Hacer clic en "Generar Certificado PDF"
  - Verificar que se abre frmSelectorInspeccion en modo PDF
  - Verificar que solo muestra inspecciones completadas
  - Seleccionar una inspección
  - Hacer clic en "Generar PDF"
  - Verificar que:
    - PDF se genera en escritorio
    - Nombre de archivo es correcto
    - PDF se abre automáticamente (si está configurado)
    - Formulario se cierra tras generación exitosa
    - Audit log registra el evento

#### **Commit y Tag**
```bash
git add frmSelectorInspeccion.frm Hoja1.bas CertificadoPDFGenerator.bas
git commit -m "feat(certificado-pdf): [FASE 3] - Integración con UI completada

- Modificar frmSelectorInspeccion para soportar modo PDF
- Agregar botón en Menú Principal
- Implementar filtro de inspecciones completadas
- Agregar registro de auditoría
- Pruebas de integración exitosas"

git tag v1.4.0-cert.3
git push origin v1.4.0-cert.3

# Backup
cp "TH-HC-001 INSPECCIONES.xlsm" "backups/TH-HC-001_CERT_FASE3_20260415.xlsm"
```

---

### **FASE 4: TESTING, REFINAMIENTO Y RELEASE** ✅

**Duración estimada:** 2-3 horas  
**Objetivo:** Pruebas exhaustivas, ajustes de formato, manejo de errores y documentación final.

#### **Tareas**

- [ ] **4.1 — Pruebas con datos reales**
  - Generar certificados para al menos 5 inspecciones diferentes:
    - Con todas las preguntas respondidas
    - Con observaciones largas
    - Con observaciones vacías
    - Con diferentes categorías (1-5)
    - Con diferentes plantas y puestos
  - Verificar en cada caso:
    - Formato correcto
    - Sin errores #REF! o #VALUE!
    - Texto no cortado
    - Tablas alineadas
    - Firmas visibles
    - Metadatos correctos

- [ ] **4.2 — Optimización de espacio (si no cabe en 1 página)**
  - Reducir tamaños de fuente si es necesario
  - Ajustar márgenes (mínimo 1cm)
  - Considerar orientación horizontal (Landscape) si es crítico
  - Alternativamente: Aceptar 2 páginas si la legibilidad lo requiere
  - **Decisión documentada:** ¿1 página forzado o 2 páginas legibles?

- [ ] **4.3 — Manejo de errores edge cases**
  - ¿Qué pasa si no hay escritorio accesible? → Guardar en Mis Documentos
  - ¿Qué pasa si el archivo ya existe? → Sobrescribir con confirmación o agregar timestamp
  - ¿Qué pasa si faltan datos en tblInspecciones? → Mostrar "N/D" en lugar de errores
  - ¿Qué pasa si no hay respuestas en tblRespuestas? → Mensaje de advertencia
  - Implementar validaciones adicionales

- [ ] **4.4 — Agregar indicador de progreso (opcional)**
  ```vba
  Application.StatusBar = "Generando certificado PDF... Por favor espere."
  ' ... código ...
  Application.StatusBar = False
  ```

- [ ] **4.5 — Documentar en README o manual de usuario**
  - Agregar sección "Generación de Certificados PDF" en documentación
  - Incluir capturas de pantalla del proceso
  - Documentar estructura del PDF
  - Documentar nomenclatura de archivos

- [ ] **4.6 — Actualizar Configuration2 con documentación**
  ```vba
  ' ══════════════════════════════════════════════════════════════
  ' CERTIFICADOS PDF - DOCUMENTACIÓN
  ' ══════════════════════════════════════════════════════════════
  ' El sistema genera certificados PDF de inspecciones completadas.
  ' 
  ' Uso:
  '   1. Desde Menú Principal → Botón "Generar Certificado PDF"
  '   2. Seleccionar inspección completada
  '   3. Clic en "Generar PDF"
  '   4. PDF se guarda en escritorio del usuario
  '
  ' Nombre archivo: CERTIFICADO_[Iniciales]_[FechaISO]_[UUIDcorto].pdf
  ' Ejemplo: CERTIFICADO_ABC_20260415_x8fL.pdf
  '
  ' Secciones del certificado:
  '   - Encabezado (logo + título)
  '   - Datos de inspección (fecha, evaluado, puesto, planta, etc.)
  '   - Resultados generales (TA scoring, RPN, categoría)
  '   - Preguntas y respuestas (por sección)
  '   - Observaciones generales
  '   - Firmas (Evaluado, Evaluador, Supervisor)
  '   - Metadatos (fecha emisión, UUID)
  ' ══════════════════════════════════════════════════════════════
  ```

- [ ] **4.7 — Consideración futura: Logo corporativo**
  - Documentar cómo insertar logo cuando esté disponible:
  ```vba
  ' En PoblarCabeceraInspeccion, agregar:
  ' Dim rutaLogo As String
  ' rutaLogo = ThisWorkbook.Path & "\assets\logo_asurity.png"
  ' If Dir(rutaLogo) <> "" Then
  '     With wsPlantilla.Pictures.Insert(rutaLogo)
  '         .Top = wsPlantilla.Range("A1").Top
  '         .Left = wsPlantilla.Range("A1").Left
  '         .Width = 60
  '         .Height = 60
  '     End With
  ' End If
  ```

- [ ] **4.8 — Testing de rendimiento**
  - Medir tiempo de generación con cronómetro
  - Objetivo: < 3 segundos para inspección típica (30 preguntas)
  - Si es más lento, optimizar:
    - Deshabilitar ScreenUpdating
    - Deshabilitar cálculo automático temporalmente
    - Reducir número de accesos a disco

- [ ] **4.9 — Revisión final de código**
  - Verificar que no quedan Debug.Print
  - Verificar que todos los On Error GoTo están presentes
  - Verificar que se limpian todos los objetos (Set x = Nothing)
  - Verificar indentación y comentarios

- [ ] **4.10 — Prueba en máquina limpia (si es posible)**
  - Copiar archivo a otra máquina
  - Verificar que funciona sin dependencias externas
  - Verificar que PDF se puede abrir sin Adobe (usar visor Windows)

#### **Commit y Tag Final**
```bash
git add .
git commit -m "feat(certificado-pdf): [FASE 4] - Testing y release v1.4.1

- Pruebas exhaustivas con múltiples inspecciones reales
- Optimización de formato para legibilidad
- Manejo robusto de errores y edge cases
- Documentación completa en código y manual
- Rendimiento validado (< 3s por certificado)
- Sistema listo para producción"

git tag v1.4.1
git push origin v1.4.1

# Backup final
cp "TH-HC-001 INSPECCIONES.xlsm" "backups/TH-HC-001_CERT_RELEASE_v1.4.1_20260415.xlsm"
```

---

## 📝 NOTAS IMPORTANTES

### **Limitaciones conocidas**

1. **Cantidad de preguntas:** Con 30 preguntas (19 + 11), es muy difícil caber en 1 página A4 manteniendo legibilidad. Opciones:
   - Reducir fuente a 6pt o 7pt (puede ser ilegible impreso)
   - Usar 2 páginas (preferible para legibilidad)
   - Usar orientación horizontal (Landscape)

2. **Logo:** El espacio está reservado pero el logo debe agregarse manualmente cuando esté disponible. Incluir instrucciones en documentación.

3. **Firmas:** Son campos vacíos para firma manual en papel. No son firmas digitales.

4. **Dependencia de Excel:** El sistema requiere Excel 2010 o superior para `ExportAsFixedFormat`. No funciona en LibreOffice o Google Sheets.

5. **Escritorio multiusuario:** Si el escritorio no es accesible (por permisos o política de red), el sistema fallará. Considerar fallback a carpeta alternativa.

### **Mejoras futuras (post-v1.4.1)**

- [ ] Permitir seleccionar carpeta de destino (en lugar de escritorio fijo)
- [ ] Generar múltiples certificados en lote (selección múltiple)
- [ ] Enviar PDF por correo electrónico automáticamente
- [ ] Firma digital usando certificados digitales
- [ ] Marca de agua "COPIA CONTROLADA" o "COPIA NO CONTROLADA"
- [ ] QR code con UUID para verificación de autenticidad
- [ ] Plantillas personalizadas por planta o tipo de inspección
- [ ] Internacionalización (idiomas múltiples)

### **Referencias técnicas**

- **ExportAsFixedFormat:** https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet.exportasfixedformat
- **PageSetup:** https://docs.microsoft.com/en-us/office/vba/api/excel.pagesetup
- **WScript.Shell SpecialFolders:** https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/windows-scripting/0ea7b5xe(v=vs.84)

---

## ✅ CHECKLIST PRE-RELEASE

Antes de marcar este sistema como completo, verificar:

- [ ] Todas las fases (0-4) completadas
- [ ] Commits y tags creados para cada fase
- [ ] Backups generados en cada fase
- [ ] Pruebas con al menos 5 inspecciones reales
- [ ] PDF generado con formato profesional y legible
- [ ] Sistema funciona sin errores
- [ ] Audit trail registra generaciones de certificados
- [ ] Documentación completa en código
- [ ] Manual de usuario actualizado (si existe)
- [ ] No quedan TODOs o FIXMEs en el código

---

**FIN DEL DOCUMENTO**  
**Proyecto:** TH-HC-001 EN PROCESO DE VALIDACIÓN  
**Versión objetivo:** v1.4.1  
**Fecha:** 15/04/2026
