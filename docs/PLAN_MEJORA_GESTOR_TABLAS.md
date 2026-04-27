# 📋 PLAN MAESTRO DE MEJORA - GESTOR DE TABLAS MAESTRAS

**Sistema de Inspecciones TH-HC-001**  
**Versión:** 2.0  
**Fecha:** Abril 25, 2026  
**Objetivo:** Transformar el gestor en un sistema profesional, intuitivo y fácil de mantener

---

## 📊 RESUMEN EJECUTIVO

### Estado Actual
- **Calificación:** 4/10
- **Tiempo para crear plantilla (20 preguntas):** 47-92 minutos
- **Lugares a modificar para agregar tabla:** 10+
- **Experiencia de usuario:** Confusa y sin ayudas

### Estado Objetivo (Post-Plan)
- **Calificación esperada:** 9/10
- **Tiempo para crear plantilla (20 preguntas):** 10-15 minutos
- **Lugares a modificar para agregar tabla:** 1 (metadata)
- **Experiencia de usuario:** Intuitiva con wizards y ayuda contextual

### ROI Estimado
- **Reducción de tiempo de operación:** 70-80%
- **Reducción de errores de usuario:** 60-70%
- **Reducción de tiempo de mantenimiento:** 80%
- **Adopción de usuario:** De 40% a 90%+

---

## 🎯 ESTRATEGIA DE IMPLEMENTACIÓN

El plan se divide en **4 FASES** progresivas, cada una con valor entregable independiente:

| Fase | Nombre | Duración | Impacto | Prioridad |
|------|--------|----------|---------|-----------|
| **1** | Quick Wins | 2-3 días | Alto | 🔴 Crítica |
| **2** | UX Profesional | 4-5 días | Muy Alto | 🟠 Alta |
| **3** | Refactoring Arquitectónico | 5-6 días | Medio | 🟡 Media |
| **4** | Funcionalidades Avanzadas | 3-4 días | Alto | 🟢 Baja |

**Duración total:** 14-18 días laborables (~3-4 semanas)

---

# 📍 FASE 1: QUICK WINS (2-3 días)

**Objetivo:** Mejoras inmediatas que no requieren cambios arquitectónicos profundos pero aumentan significativamente la usabilidad.

---

## 1.1 Búsqueda y Filtrado en Lista (Prioridad: CRÍTICA)

### Problema que resuelve
❌ Con 100+ registros, imposible encontrar datos  
✅ Búsqueda instantánea mientras se escribe

### Implementación

**Paso 1:** Agregar TextBox de búsqueda en el formulario

```vba
' --- En UserForm_Initialize (después de ConfigurarSelectorTabla) ---
Private Sub ConfigurarBuscador()
    ' Agregar label "Buscar:"
    With Me.lblBuscar
        .Left = MARGIN_LEFT
        .Top = 56
        .Width = 60
        .Height = 18
        .Caption = "Buscar:"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
    End With
    
    ' Agregar TextBox de búsqueda
    With Me.txtBuscar
        .Left = 72
        .Top = 54
        .Width = 200
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Size = 9
    End With
    
    ' Ajustar posición de lstDatos
    Me.lstDatos.Top = 80
    Me.lstDatos.Height = 118  ' Reducido para dar espacio al buscador
End Sub
```

**Paso 2:** Implementar búsqueda en tiempo real

```vba
' --- En frmGestorTablas ---
Private mDatosCompletos As Variant  ' Cache de datos sin filtrar

Private Sub txtBuscar_Change()
    If mTablaActual = "" Then Exit Sub
    If IsEmpty(mDatosCompletos) Then Exit Sub
    
    Dim filtro As String
    filtro = UCase(Trim(Me.txtBuscar.Value))
    
    ' Limpiar lista
    Me.lstDatos.Clear
    
    ' Si no hay filtro, mostrar todo
    If filtro = "" Then
        Call CargarDatosEnLista
        Exit Sub
    End If
    
    ' Filtrar datos
    Dim fila As Long
    For fila = LBound(mDatosCompletos, 1) To UBound(mDatosCompletos, 1)
        ' Buscar en todas las columnas
        Dim encontrado As Boolean: encontrado = False
        Dim col As Long
        For col = 1 To UBound(mDatosCompletos, 2)
            If InStr(1, UCase(CStr(mDatosCompletos(fila, col))), filtro, vbTextCompare) > 0 Then
                encontrado = True
                Exit For
            End If
        Next col
        
        If encontrado Then
            ' Agregar fila a la lista (usar la lógica existente de AddItem)
            Call AgregarFilaALista(fila, mDatosCompletos)
        End If
    Next fila
    
    Call MostrarEstado("Mostrando " & Me.lstDatos.ListCount & " registros filtrados.", False)
End Sub

Private Sub CargarDatosEnLista()
    ' MODIFICAR la función existente para guardar en mDatosCompletos
    Me.lstDatos.Clear
    
    If mTablaActual = "" Then Exit Sub
    
    mDatosCompletos = TableManager.ObtenerDatosTabla(mTablaActual)
    
    If IsEmpty(mDatosCompletos) Then
        Call MostrarEstado("La tabla está vacía.", False)
        Exit Sub
    End If
    
    ' Renderizar todas las filas
    Dim fila As Long
    For fila = LBound(mDatosCompletos, 1) To UBound(mDatosCompletos, 1)
        Call AgregarFilaALista(fila, mDatosCompletos)
    Next fila
End Sub

Private Sub AgregarFilaALista(ByVal filaIdx As Long, ByRef datos As Variant)
    ' Lógica extraída del bucle actual de CargarDatosEnLista
    ' (Código de transformación de IDs a nombres y AddItem)
End Sub
```

**Tiempo:** 3-4 horas  
**Impacto:** ⭐⭐⭐⭐⭐ (Crítico)

---

## 1.2 Tooltips y Ayuda Contextual (Prioridad: ALTA)

### Problema que resuelve
❌ Usuario no sabe qué ingresar en cada campo  
✅ Tooltips explican cada campo al pasar el mouse

### Implementación

```vba
Private Sub ConfigurarTooltips()
    ' Tooltips generales
    Me.btnNuevo.ControlTipText = "Crear un nuevo registro (Ctrl+N)"
    Me.btnGuardar.ControlTipText = "Guardar cambios (Ctrl+S)"
    Me.btnEliminar.ControlTipText = "Eliminar registro seleccionado (Del)"
    Me.btnValidar.ControlTipText = "Validar integridad de todas las tablas"
    Me.txtBuscar.ControlTipText = "Buscar en todos los campos de la tabla"
End Sub

Private Sub ConfigurarCamposParaTabla(ByVal tabla As String)
    ' ... código existente ...
    
    ' AGREGAR tooltips específicos por tabla
    Select Case tabla
        Case "CRITICIDAD"
            Me.txtCampo2.ControlTipText = "Nombre único de la criticidad (ej: Alto, Medio, Bajo)"
            Me.txtCampo3.ControlTipText = "Valor numérico positivo (1-10 recomendado)"
            
        Case "SECCIONES"
            Me.txtCampo2.ControlTipText = "Nombre de la sección (ej: Auditoría de procesos)"
            Me.txtCampo3.ControlTipText = "Tipo: 'Selección' o 'Puntaje'"
            
        Case "PLANTILLAS"
            Me.txtCampo2.ControlTipText = "Nombre descriptivo de la plantilla"
            Me.txtCampo3.ControlTipText = "Etapa del proceso (ej: Inicial, Seguimiento, Final)"
            Me.txtCampo7.ControlTipText = "Puesto responsable (ej: Auditor Senior)"
            Me.txtCampo8.ControlTipText = "Frecuencia en MESES (ej: 1, 3, 6, 12)"
            
        Case "OPCIONES"
            Me.cmbCampo4.ControlTipText = "Seleccione la sección a la que pertenece"
            Me.cmbCampo5.ControlTipText = "Criticidad asociada (opcional)"
            Me.txtCampo2.ControlTipText = "Texto de la opción (ej: Conforme, No conforme)"
            Me.txtCampo3.ControlTipText = "Puntaje numérico (positivo o negativo)"
            
        Case "PREGUNTAS"
            Me.cmbCampo6.ControlTipText = "Plantilla donde aparecerá esta pregunta"
            Me.cmbCampo4.ControlTipText = "Sección de la pregunta"
            Me.cmbCampo5.ControlTipText = "Nivel de criticidad de la pregunta"
            Me.txtCampo2.ControlTipText = "Texto completo de la pregunta (máx 500 caracteres)"
            Me.txtCampo7.ControlTipText = "Número de orden dentro de la sección (1, 2, 3...)"
            Me.txtCampo8.ControlTipText = "Observaciones o guías para el auditor"
    End Select
End Sub
```

**Tiempo:** 1 hora  
**Impacto:** ⭐⭐⭐⭐ (Alto)

---

## 1.3 Validación en Tiempo Real (Prioridad: ALTA)

### Problema que resuelve
❌ Usuario solo ve errores al guardar  
✅ Feedback instantáneo mientras escribe

### Implementación

```vba
' --- Validación de campos numéricos ---
Private Sub txtCampo3_Change()
    If mTablaActual = "" Then Exit Sub
    
    Select Case mTablaActual
        Case "CRITICIDAD", "PLANTILLAS", "OPCIONES"
            ' Validar que sea numérico
            If Trim(Me.txtCampo3.Value) <> "" Then
                If Not IsNumeric(Me.txtCampo3.Value) Then
                    Me.txtCampo3.BackColor = RGB(255, 220, 220)  ' Rojo claro
                    Me.lblEstado.Caption = " ⚠️ Este campo debe ser numérico"
                    Me.lblEstado.ForeColor = COLOR_ESTADO_ERROR
                Else
                    Me.txtCampo3.BackColor = vbWhite
                    Me.lblEstado.Caption = " Listo"
                    Me.lblEstado.ForeColor = COLOR_ESTADO_OK
                End If
            Else
                Me.txtCampo3.BackColor = vbWhite
            End If
    End Select
End Sub

Private Sub txtCampo7_Change()
    ' Validar orden (numérico positivo)
    If mTablaActual = "PREGUNTAS" Then
        If Trim(Me.txtCampo7.Value) <> "" Then
            If Not IsNumeric(Me.txtCampo7.Value) Or CDbl(Me.txtCampo7.Value) <= 0 Then
                Me.txtCampo7.BackColor = RGB(255, 220, 220)
                Me.lblEstado.Caption = " ⚠️ El orden debe ser un número positivo"
                Me.lblEstado.ForeColor = COLOR_ESTADO_ERROR
            Else
                Me.txtCampo7.BackColor = vbWhite
                Me.lblEstado.Caption = " Listo"
                Me.lblEstado.ForeColor = COLOR_ESTADO_OK
            End If
        End If
    End If
End Sub

' --- Validación de campos tipo respuesta ---
Private Sub txtCampo3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If mTablaActual = "SECCIONES" Then
        Dim tipo As String
        tipo = Trim(Me.txtCampo3.Value)
        If tipo <> "" And tipo <> "Selección" And tipo <> "Puntaje" Then
            Cancel = True
            MsgBox "Tipo de respuesta debe ser 'Selección' o 'Puntaje'", vbExclamation
            Me.txtCampo3.SetFocus
        End If
    End If
End Sub

' --- Marcador visual de campos obligatorios ---
Private Sub MostrarCampo(ByVal numCampo As Long, ByVal etiqueta As String, _
                          ByVal tipoCampo As String, ByVal soloLectura As Boolean)
    ' ... código existente ...
    
    ' AGREGAR: Marcar campos obligatorios con asterisco
    If Not soloLectura Then
        Select Case mTablaActual
            Case "CRITICIDAD"
                If numCampo = 2 Then Me.lblCampo2.Caption = etiqueta & " *"
            Case "SECCIONES"
                If numCampo = 2 Then Me.lblCampo2.Caption = etiqueta & " *"
            Case "PLANTILLAS"
                If numCampo = 2 Then Me.lblCampo2.Caption = etiqueta & " *"
            Case "OPCIONES"
                If numCampo = 4 Then Me.lblCampo4.Caption = etiqueta & " *"
                If numCampo = 2 Then Me.lblCampo2.Caption = etiqueta & " *"
            Case "PREGUNTAS"
                If numCampo = 6 Then Me.lblCampo6.Caption = etiqueta & " *"
                If numCampo = 4 Then Me.lblCampo4.Caption = etiqueta & " *"
                If numCampo = 5 Then Me.lblCampo5.Caption = etiqueta & " *"
                If numCampo = 2 Then Me.lblCampo2.Caption = etiqueta & " *"
        End Select
    End If
End Sub
```

**Tiempo:** 2 horas  
**Impacto:** ⭐⭐⭐⭐ (Alto)

---

## 1.4 Atajos de Teclado (Prioridad: MEDIA)

### Problema que resuelve
❌ Todo requiere mouse  
✅ Usuarios expertos trabajan más rápido

### Implementación

```vba
' --- En frmGestorTablas ---
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                              ByVal Shift As Integer)
    ' Ctrl+N = Nuevo
    If Shift = fmCtrlMask And KeyCode = 78 Then  ' N
        If Me.btnNuevo.Enabled Then Call btnNuevo_Click
        KeyCode = 0
    End If
    
    ' Ctrl+S = Guardar
    If Shift = fmCtrlMask And KeyCode = 83 Then  ' S
        If Me.btnGuardar.Enabled Then Call btnGuardar_Click
        KeyCode = 0
    End If
    
    ' Delete = Eliminar
    If KeyCode = 46 Then  ' Del
        If Me.btnEliminar.Enabled Then Call btnEliminar_Click
        KeyCode = 0
    End If
    
    ' Ctrl+F = Focus en búsqueda
    If Shift = fmCtrlMask And KeyCode = 70 Then  ' F
        Me.txtBuscar.SetFocus
        KeyCode = 0
    End If
    
    ' F5 = Recargar datos
    If KeyCode = 116 Then  ' F5
        Call CargarDatosEnLista
        Call MostrarEstado("Datos recargados.", False)
        KeyCode = 0
    End If
End Sub

' Habilitar KeyPreview
Private Sub UserForm_Initialize()
    ' ... código existente ...
    Me.KeyPreview = True  ' AGREGAR
End Sub
```

**Tiempo:** 1 hora  
**Impacto:** ⭐⭐⭐ (Medio)

---

## 1.5 Botón "Duplicar" (Prioridad: ALTA)

### Problema que resuelve
❌ Crear preguntas similares requiere copiar a mano  
✅ Duplicar y modificar solo lo necesario

### Implementación

```vba
' --- Agregar botón en ConfigurarBotones ---
With Me.btnDuplicar
    .Left = MARGIN_LEFT + (btnWidth + btnSpacing) * 3
    .Top = btnTop
    .Width = btnWidth
    .Height = btnHeight
    .Caption = "Duplicar"
    .Font.Name = "Segoe UI"
    .Font.Size = 10
    .BackColor = RGB(200, 230, 255)  ' Azul claro
    .Enabled = False
End With

' --- Evento del botón ---
Private Sub btnDuplicar_Click()
    If mTablaActual = "" Or mFilaSeleccionada < 0 Then
        Call MostrarEstado("ERROR: Seleccione un registro primero.", True)
        Exit Sub
    End If
    
    ' Cargar datos del registro actual
    Call CargarFilaEnCampos(mFilaSeleccionada)
    
    ' Cambiar a modo NUEVO
    mModoEdicion = "NUEVO"
    mFilaSeleccionada = -1
    
    ' Generar nuevo ID
    Me.txtCampo1.Value = TableManager.GenerarNuevoID(mTablaActual)
    
    ' Modificar nombre para evitar duplicados
    Select Case mTablaActual
        Case "CRITICIDAD", "SECCIONES", "PLANTILLAS"
            Me.txtCampo2.Value = Me.txtCampo2.Value & " (Copia)"
        Case "PREGUNTAS"
            Me.txtCampo2.Value = Me.txtCampo2.Value & " (Copia)"
            ' Incrementar orden
            If IsNumeric(Me.txtCampo7.Value) Then
                Me.txtCampo7.Value = CDbl(Me.txtCampo7.Value) + 1
            End If
    End Select
    
    Call HabilitarBotones(True, True, False)
    Call MostrarEstado("Registro duplicado. Modifique y presione Guardar.", False)
    
    ' Focus en campo nombre
    If Me.txtCampo2.Visible Then Me.txtCampo2.SetFocus
End Sub

' --- Habilitar/deshabilitar botón ---
Private Sub lstDatos_Click()
    ' ... código existente ...
    Call HabilitarBotones(True, True, True)
    Me.btnDuplicar.Enabled = True  ' AGREGAR
End Sub
```

**Tiempo:** 1.5 horas  
**Impacto:** ⭐⭐⭐⭐ (Alto)

---

## 1.6 Corregir Inconsistencias (Prioridad: CRÍTICA)

### Problema que resuelve
❌ Documentación dice "días", código usa "meses"  
✅ Consistencia total

### Implementación

**Opción A: Renombrar columna en Excel**
```vba
' Ejecutar una sola vez para renombrar
Sub RenombrarColumnaFrecuencia()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Checklist")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("tblPlantillas")
    
    ' Buscar columna "Frecuencia meses" y renombrar
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name = "Frecuencia meses" Then
            col.Name = "Frecuencia (meses)"
            Exit For
        End If
    Next col
End Sub
```

**Opción B: Actualizar documentación**
```markdown
<!-- En GESTOR_TABLAS_MANUAL.md -->
| Frecuencia (meses) | Número | No | Positivo, ej: 1, 3, 6, 12 |
```

**Actualizar tooltips:**
```vba
Me.txtCampo8.ControlTipText = "Frecuencia en MESES (ej: 1, 3, 6, 12)"
Me.lblCampo8.Caption = "Frecuencia (meses): *"
```

**Tiempo:** 30 minutos  
**Impacto:** ⭐⭐⭐⭐⭐ (Crítico - Previene errores)

---

## 1.7 Indicadores Visuales de Estado (Prioridad: MEDIA)

### Problema que resuelve
❌ No se sabe si hay cambios sin guardar  
✅ Indicador visual claro

### Implementación

```vba
Private mHayCambiosPendientes As Boolean

Private Sub txtCampo2_Change()
    If mModoEdicion <> "" Then
        mHayCambiosPendientes = True
        Me.Caption = "Gestor de Tablas Maestras *"  ' Asterisco = cambios pendientes
    End If
End Sub

' Aplicar a todos los campos editables (txtCampo3, cmbCampo4, etc.)

Private Sub btnGuardar_Click()
    ' ... código existente ...
    
    If resultado Then
        mHayCambiosPendientes = False
        Me.Caption = "Gestor de Tablas Maestras"  ' Quitar asterisco
        ' ... resto del código ...
    End If
End Sub

Private Sub btnCerrar_Click()
    If mHayCambiosPendientes Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("Tiene cambios sin guardar. ¿Desea cerrar de todas formas?", _
                      vbQuestion + vbYesNo, "⚠️ Cambios pendientes")
        If resp = vbNo Then Exit Sub
    End If
    
    Unload Me
End Sub
```

**Tiempo:** 1 hora  
**Impacado:** ⭐⭐⭐ (Medio)

---

## 📊 RESUMEN FASE 1

| Mejora | Tiempo | Impacto | Status |
|--------|--------|---------|--------|
| 1.1 Búsqueda y filtrado | 3-4h | ⭐⭐⭐⭐⭐ | ⬜ Pendiente |
| 1.2 Tooltips y ayuda | 1h | ⭐⭐⭐⭐ | ⬜ Pendiente |
| 1.3 Validación en tiempo real | 2h | ⭐⭐⭐⭐ | ⬜ Pendiente |
| 1.4 Atajos de teclado | 1h | ⭐⭐⭐ | ⬜ Pendiente |
| 1.5 Botón duplicar | 1.5h | ⭐⭐⭐⭐ | ⬜ Pendiente |
| 1.6 Corregir inconsistencias | 0.5h | ⭐⭐⭐⭐⭐ | ⬜ Pendiente |
| 1.7 Indicadores visuales | 1h | ⭐⭐⭐ | ⬜ Pendiente |
| **TOTAL FASE 1** | **10-11h** | **Muy Alto** | **⬜ 0%** |

**Resultado esperado:** Sistema 60% más fácil de usar sin cambios arquitectónicos.

---

# 🎨 FASE 2: UX PROFESIONAL (4-5 días)

**Objetivo:** Transformar la interfaz en una experiencia profesional con wizards, vistas mejoradas y workflows inteligentes.

---

## 2.1 Wizard para Crear Plantilla Completa (Prioridad: CRÍTICA)

### Problema que resuelve
❌ Crear plantilla con 20 preguntas = 47-92 minutos  
✅ Wizard guiado = 10-15 minutos

### Diseño del Wizard

**Nuevo formulario:** `frmWizardPlantilla.frm`

**Paso 1: Datos Básicos**
```
┌─────────────────────────────────────────────────────────┐
│  Paso 1/4: Datos Básicos de la Plantilla               │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  Nombre: ________________________________________       │
│  Etapa: [Inicial ▼]                                     │
│  Puesto responsable: ____________________________       │
│  Frecuencia (meses): [__]                               │
│                                                         │
│  ┌─────────────────────────────────────────────┐       │
│  │ ℹ️ Esta plantilla se usará para crear      │       │
│  │   inspecciones de forma recurrente.        │       │
│  └─────────────────────────────────────────────┘       │
│                                                         │
│               [Cancelar]  [< Atrás]  [Siguiente >]     │
└─────────────────────────────────────────────────────────┘
```

**Paso 2: Seleccionar Secciones**
```
┌─────────────────────────────────────────────────────────┐
│  Paso 2/4: Secciones a Incluir                         │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  Seleccione las secciones que tendrá esta plantilla:   │
│                                                         │
│  ☑ Auditoría de procesos (3 opciones disponibles)      │
│  ☑ Seguridad e higiene (4 opciones disponibles)        │
│  ☐ Cumplimiento regulatorio (2 opciones disponibles)   │
│  ☑ Gestión de recursos (3 opciones disponibles)        │
│  ☐ Documentación (5 opciones disponibles)              │
│                                                         │
│  [+ Nueva Sección]                                      │
│                                                         │
│               [Cancelar]  [< Atrás]  [Siguiente >]     │
└─────────────────────────────────────────────────────────┘
```

**Paso 3: Agregar Preguntas (con grid editable)**
```
┌─────────────────────────────────────────────────────────┐
│  Paso 3/4: Preguntas por Sección                       │
├─────────────────────────────────────────────────────────┤
│  Sección: [Auditoría de procesos ▼]      [+ Pregunta]  │
│                                                         │
│  ┌─────────────────────────────────────────────────┐   │
│  │#│Pregunta                    │Criticidad│Orden │   │
│  ├─┼────────────────────────────┼──────────┼──────┤   │
│  │1│¿Existen procedimientos...  │Alto      │1  [↑│   │
│  │2│¿Se han actualizado los...  │Medio     │2  [↓│   │
│  │3│¿Personal capacitado en...  │Crítico   │3     │   │
│  │ │                            │          │      │   │
│  └─────────────────────────────────────────────────┘   │
│  [Editar] [Duplicar] [Eliminar] [Importar CSV]         │
│                                                         │
│  Preguntas: Auditoría (3), Seguridad (0), Gestión (0)  │
│                                                         │
│               [Cancelar]  [< Atrás]  [Siguiente >]     │
└─────────────────────────────────────────────────────────┘
```

**Paso 4: Preview y Confirmación**
```
┌─────────────────────────────────────────────────────────┐
│  Paso 4/4: Resumen y Confirmación                      │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  📋 Plantilla: Auditoría Inicial                        │
│  📅 Frecuencia: 1 mes                                   │
│  👤 Responsable: Auditor Senior                         │
│                                                         │
│  📊 Contenido:                                          │
│    • Auditoría de procesos: 3 preguntas                │
│    • Seguridad e higiene: 5 preguntas                  │
│    • Gestión de recursos: 2 preguntas                  │
│                                                         │
│  Total: 10 preguntas en 3 secciones                    │
│                                                         │
│  ☑ Activar todas las preguntas al crear                │
│  ☑ Generar audit log de creación                       │
│                                                         │
│               [Cancelar]  [< Atrás]  [✓ Crear Todo]    │
└─────────────────────────────────────────────────────────┘
```

### Implementación Técnica

```vba
' ======================================================================
' UserForm: frmWizardPlantilla
' Descripción: Wizard multi-paso para crear plantillas completas
' ======================================================================
Option Explicit

Private mPasoActual As Integer
Private mDatosPlantilla As Object    ' Dictionary con datos básicos
Private mSeccionesSeleccionadas As Collection
Private mPreguntasPorSeccion As Object  ' Dictionary de Collections

Private Sub UserForm_Initialize()
    mPasoActual = 1
    Set mDatosPlantilla = CreateObject("Scripting.Dictionary")
    Set mSeccionesSeleccionadas = New Collection
    Set mPreguntasPorSeccion = CreateObject("Scripting.Dictionary")
    
    Call MostrarPaso(1)
End Sub

Private Sub MostrarPaso(ByVal paso As Integer)
    ' Ocultar todos los frames
    Me.fraPaso1.Visible = False
    Me.fraPaso2.Visible = False
    Me.fraPaso3.Visible = False
    Me.fraPaso4.Visible = False
    
    ' Mostrar frame correspondiente
    Select Case paso
        Case 1
            Me.fraPaso1.Visible = True
            Me.btnAtras.Enabled = False
            Me.btnSiguiente.Caption = "Siguiente >"
        Case 2
            Me.fraPaso2.Visible = True
            Call CargarSeccionesDisponibles
            Me.btnAtras.Enabled = True
            Me.btnSiguiente.Caption = "Siguiente >"
        Case 3
            Me.fraPaso3.Visible = True
            Call CargarGridPreguntas
            Me.btnAtras.Enabled = True
            Me.btnSiguiente.Caption = "Siguiente >"
        Case 4
            Me.fraPaso4.Visible = True
            Call MostrarPreview
            Me.btnAtras.Enabled = True
            Me.btnSiguiente.Caption = "✓ Crear Todo"
    End Select
    
    mPasoActual = paso
    Me.lblPaso.Caption = "Paso " & paso & "/4"
End Sub

Private Sub btnSiguiente_Click()
    ' Validar paso actual
    If Not ValidarPasoActual() Then Exit Sub
    
    ' Guardar datos del paso
    Call GuardarDatosPaso(mPasoActual)
    
    ' Avanzar o crear
    If mPasoActual = 4 Then
        Call CrearPlantillaCompleta
    Else
        Call MostrarPaso(mPasoActual + 1)
    End If
End Sub

Private Sub btnAtras_Click()
    If mPasoActual > 1 Then
        Call MostrarPaso(mPasoActual - 1)
    End If
End Sub

Private Sub CrearPlantillaCompleta()
    ' Deshabilitar protección
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Configuration2.SHEET_CHECKLIST)
    Call SheetProtector2.UnprotectSheet(ws, Configuration2.APP_PASSWORD)
    
    ' 1. Crear plantilla
    Dim idPlantilla As String
    idPlantilla = TableManager.GenerarNuevoID("PLANTILLAS")
    
    Dim datosPlantilla As Object
    Set datosPlantilla = mDatosPlantilla
    datosPlantilla("ID") = idPlantilla
    
    If Not TableManager.InsertarFila("PLANTILLAS", datosPlantilla) Then
        MsgBox "Error al crear plantilla", vbCritical
        Exit Sub
    End If
    
    ' 2. Crear preguntas por cada sección
    Dim seccion As Variant
    For Each seccion In mSeccionesSeleccionadas
        Dim idSeccion As String
        idSeccion = seccion
        
        If mPreguntasPorSeccion.Exists(idSeccion) Then
            Dim preguntas As Collection
            Set preguntas = mPreguntasPorSeccion(idSeccion)
            
            Dim pregunta As Variant
            For Each pregunta In preguntas
                Dim datosPregunta As Object
                Set datosPregunta = pregunta
                
                ' Agregar ID de plantilla
                datosPregunta("IDPlantilla") = idPlantilla
                datosPregunta("IDPregunta") = TableManager.GenerarNuevoID("PREGUNTAS")
                
                Call TableManager.InsertarFila("PREGUNTAS", datosPregunta)
            Next pregunta
        End If
    Next seccion
    
    ' 3. Re-proteger
    If Configuration2.ENABLE_SHEET_PROTECTION Then
        Call SheetProtector2.ProtectSheet(ws, Configuration2.APP_PASSWORD)
    End If
    
    ' 4. Registrar en audit
    Call AuditLogger2.LogAction("WIZARD_PLANTILLA", "Creada plantilla completa: " & _
                                 datosPlantilla("Nombre"), idPlantilla)
    
    MsgBox "✓ Plantilla creada exitosamente con " & ContarTotalPreguntas() & " preguntas.", _
           vbInformation, "Éxito"
    
    Unload Me
End Sub

' --- Funciones auxiliares ---
Private Function ValidarPasoActual() As Boolean
    Select Case mPasoActual
        Case 1
            If Trim(Me.txtNombrePlantilla.Value) = "" Then
                MsgBox "Debe ingresar un nombre para la plantilla.", vbExclamation
                ValidarPasoActual = False
                Exit Function
            End If
        Case 2
            If mSeccionesSeleccionadas.Count = 0 Then
                MsgBox "Debe seleccionar al menos una sección.", vbExclamation
                ValidarPasoActual = False
                Exit Function
            End If
        Case 3
            If ContarTotalPreguntas() = 0 Then
                Dim resp As VbMsgBoxResult
                resp = MsgBox("No ha agregado ninguna pregunta. ¿Desea continuar de todas formas?", _
                             vbQuestion + vbYesNo)
                If resp = vbNo Then
                    ValidarPasoActual = False
                    Exit Function
                End If
            End If
    End Select
    
    ValidarPasoActual = True
End Function

Private Sub GuardarDatosPaso(ByVal paso As Integer)
    Select Case paso
        Case 1
            mDatosPlantilla("Nombre") = Trim(Me.txtNombrePlantilla.Value)
            mDatosPlantilla("Etapa") = Me.cboEtapa.Value
            mDatosPlantilla("Puesto") = Trim(Me.txtPuesto.Value)
            mDatosPlantilla("Frecuencia") = Me.txtFrecuencia.Value
        Case 2
            ' Ya guardado en eventos de checkboxes
        Case 3
            ' Ya guardado en eventos del grid
    End Select
End Sub

Private Function ContarTotalPreguntas() As Long
    Dim total As Long: total = 0
    Dim key As Variant
    For Each key In mPreguntasPorSeccion.Keys
        total = total + mPreguntasPorSeccion(key).Count
    Next key
    ContarTotalPreguntas = total
End Function
```

**Tiempo:** 2-3 días  
**Impacto:** ⭐⭐⭐⭐⭐ (Crítico - Reduce tiempo de 90 min a 15 min)

---

## 2.2 Importación Masiva desde CSV/Excel (Prioridad: ALTA)

### Problema que resuelve
❌ Copiar 50 preguntas de un documento externo = trabajo manual  
✅ Importar CSV = segundos

### Implementación

```vba
' --- Botón en Paso 3 del wizard o en gestor principal ---
Private Sub btnImportarCSV_Click()
    Dim filePath As String
    filePath = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv", , _
                                          "Importar Preguntas desde CSV")
    
    If filePath = "False" Then Exit Sub  ' Usuario canceló
    
    ' Leer CSV
    Dim preguntas As Collection
    Set preguntas = LeerPreguntasDesdeCSV(filePath)
    
    If preguntas.Count = 0 Then
        MsgBox "No se encontraron preguntas válidas en el archivo.", vbExclamation
        Exit Sub
    End If
    
    ' Mostrar preview
    Dim frmPreview As frmImportPreview
    Set frmPreview = New frmImportPreview
    frmPreview.CargarPreguntas preguntas
    frmPreview.Show
    
    ' Si confirmó, agregar a la colección
    If frmPreview.Confirmado Then
        ' Agregar a mPreguntasPorSeccion
        Dim pregunta As Variant
        For Each pregunta In preguntas
            ' ... lógica de agregar ...
        Next pregunta
        
        Call MostrarEstado("Importadas " & preguntas.Count & " preguntas.", False)
    End If
End Sub

Private Function LeerPreguntasDesdeCSV(ByVal filePath As String) As Collection
    ' Formato esperado del CSV:
    ' Sección,Criticidad,Pregunta,Orden,Observaciones
    
    Set LeerPreguntasDesdeCSV = New Collection
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1, False)  ' 1 = ForReading
    
    ' Saltar header
    If Not ts.AtEndOfStream Then ts.ReadLine
    
    Dim lineNum As Long: lineNum = 1
    Do While Not ts.AtEndOfStream
        lineNum = lineNum + 1
        Dim linea As String
        linea = ts.ReadLine
        
        ' Parsear CSV (simple split por comas)
        Dim campos() As String
        campos = Split(linea, ",")
        
        If UBound(campos) >= 2 Then  ' Al menos Sección, Criticidad, Pregunta
            Dim pregunta As Object
            Set pregunta = CreateObject("Scripting.Dictionary")
            
            ' Resolver ID de sección por nombre
            pregunta("IDSeccion") = ResolverIDSeccion(Trim(campos(0)))
            pregunta("IDCriticidad") = ResolverIDCriticidad(Trim(campos(1)))
            pregunta("TextoPregunta") = Trim(campos(2))
            
            If UBound(campos) >= 3 Then
                pregunta("Orden") = Trim(campos(3))
            Else
                pregunta("Orden") = lineNum - 1
            End If
            
            If UBound(campos) >= 4 Then
                pregunta("Observaciones") = Trim(campos(4))
            Else
                pregunta("Observaciones") = ""
            End If
            
            pregunta("Activo") = "Sí"
            
            LeerPreguntasDesdeCSV.Add pregunta
        End If
    Loop
    
    ts.Close
End Function

Private Function ResolverIDSeccion(ByVal nombreSeccion As String) As String
    ' Buscar en tblSecciones por nombre
    Dim tbl As ListObject
    Set tbl = TableManager.ObtenerListObject("SECCIONES")
    
    If tbl Is Nothing Then
        ResolverIDSeccion = ""
        Exit Function
    End If
    
    Dim fila As Long
    For fila = 1 To tbl.ListRows.Count
        If UCase(Trim(tbl.ListRows(fila).Range.Cells(1, 2).Value)) = UCase(nombreSeccion) Then
            ResolverIDSeccion = tbl.ListRows(fila).Range.Cells(1, 1).Value
            Exit Function
        End If
    Next fila
    
    ResolverIDSeccion = ""
End Function
```

**Plantilla CSV:**
```csv
Sección,Criticidad,Pregunta,Orden,Observaciones
Auditoría de procesos,Alto,¿Existen procedimientos documentados?,1,Verificar vigencia
Auditoría de procesos,Medio,¿Se han actualizado los procedimientos?,2,Revisar últimas 6 meses
Seguridad e higiene,Crítico,¿Personal capacitado en seguridad?,1,Verificar certificados
```

**Tiempo:** 1-1.5 días  
**Impacto:** ⭐⭐⭐⭐⭐ (Crítico para usuarios avanzados)

---

## 2.3 Mejora Visual del ListBox (Prioridad: MEDIA)

### Problema que resuelve
❌ ListBox difícil de leer con 10 columnas comprimidas  
✅ Vista con colores, zebra-striping y headers fijos

### Implementación

**Opción A: UserForm con ListView (requiere referencia MSComctlLib)**
```vba
' Alternativa profesional con ListView
' Requiere: Agregar referencia a "Microsoft Windows Common Controls 6.0"

Private Sub ConfigurarListView()
    With Me.lvwDatos
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .HideSelection = False
        
        ' Agregar columnas según tabla
        Select Case mTablaActual
            Case "PREGUNTAS"
                .ColumnHeaders.Add , "ID", "ID", 0  ' Oculta
                .ColumnHeaders.Add , "Plantilla", "Plantilla", 120
                .ColumnHeaders.Add , "Sección", "Sección", 100
                .ColumnHeaders.Add , "Criticidad", "Criticidad", 100
                .ColumnHeaders.Add , "Pregunta", "Pregunta", 250
                .ColumnHeaders.Add , "Orden", "Orden", 50
                .ColumnHeaders.Add , "Activo", "Activo", 50
        End Select
        
        ' Agregar ítems con colores alternados
        Dim i As Long
        For i = 1 To UBound(datos, 1)
            Dim item As ListItem
            Set item = .ListItems.Add(, , datos(i, 1))  ' ID
            item.SubItems(1) = datos(i, 2)  ' Plantilla
            item.SubItems(2) = datos(i, 3)  ' Sección
            ' ... etc
            
            ' Zebra striping
            If i Mod 2 = 0 Then
                item.BackColor = RGB(245, 245, 245)
            End If
            
            ' Color por criticidad
            Select Case datos(i, 4)
                Case "Crítico", "Alto"
                    item.ForeColor = RGB(200, 0, 0)
                Case "Medio"
                    item.ForeColor = RGB(200, 100, 0)
            End Select
        Next i
    End With
End Sub
```

**Opción B: Mejorar ListBox existente (sin referencias externas)**
```vba
' Si no se puede agregar referencias, mejorar el ListBox actual

Private Sub CargarDatosEnLista()
    ' ... código existente ...
    
    ' Agregar header manual en la primera fila
    Me.lstDatos.AddItem GetHeadersParaTabla(mTablaActual)
    Me.lstDatos.List(0, 0) = "──────"  ' Separador visual
    
    ' Resto de datos ...
End Sub

Private Function GetHeadersParaTabla(ByVal tabla As String) As String
    Select Case tabla
        Case "CRITICIDAD"
            GetHeadersParaTabla = "ID" & vbTab & "Nombre" & vbTab & "Valor"
        Case "PREGUNTAS"
            GetHeadersParaTabla = "Plantilla" & vbTab & "Sección" & vbTab & "Criticidad" & vbTab & _
                                 "Pregunta" & vbTab & "Orden" & vbTab & "Activo"
        ' ... etc
    End Select
End Function

' Deshabilitar selección del header
Private Sub lstDatos_Click()
    If Me.lstDatos.ListIndex = 0 Then
        Me.lstDatos.ListIndex = -1  ' Deseleccionar header
        Exit Sub
    End If
    
    mFilaSeleccionada = Me.lstDatos.ListIndex - 1  ' Ajustar por header
    ' ... resto del código ...
End Sub
```

**Tiempo:** 1 día (Opción A) o 4 horas (Opción B)  
**Impacto:** ⭐⭐⭐⭐ (Alto)

---

## 2.4 Panel de Ayuda Contextual Dinámica (Prioridad: MEDIA)

### Problema que resuelve
❌ Usuario debe consultar manual externo  
✅ Ayuda integrada en el formulario

### Implementación

```vba
' --- Agregar frame de ayuda en el formulario ---
Private Sub ConfigurarPanelAyuda()
    With Me.fraAyuda
        .Left = FORM_WIDTH - 220
        .Top = 204
        .Width = 208
        .Height = 246
        .Caption = " 💡 Ayuda "
        .BackColor = RGB(255, 255, 230)  ' Amarillo claro
    End With
    
    With Me.txtAyuda
        .Left = 8
        .Top = 20
        .Width = 192
        .Height = 218
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .BackColor = RGB(255, 255, 230)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
    End With
    
    ' Ajustar ancho de fraEdicion para dar espacio
    Me.fraEdicion.Width = CONTENT_WIDTH - 216
End Sub

Private Sub ConfigurarCamposParaTabla(ByVal tabla As String)
    ' ... código existente ...
    
    ' AGREGAR: Mostrar ayuda contextual
    Call MostrarAyudaParaTabla(tabla)
End Sub

Private Sub MostrarAyudaParaTabla(ByVal tabla As String)
    Dim ayuda As String
    
    Select Case tabla
        Case "CRITICIDAD"
            ayuda = "CRITICIDAD" & vbCrLf & vbCrLf & _
                   "Define los niveles de severidad para evaluaciones." & vbCrLf & vbCrLf & _
                   "📝 Campos obligatorios:" & vbCrLf & _
                   "• Nombre: Único" & vbCrLf & _
                   "• Valor: Numérico (1-10)" & vbCrLf & vbCrLf & _
                   "💡 Ejemplos:" & vbCrLf & _
                   "• Crítico (10)" & vbCrLf & _
                   "• Alto (7)" & vbCrLf & _
                   "• Medio (5)" & vbCrLf & _
                   "• Bajo (2)" & vbCrLf & vbCrLf & _
                   "⚠️ No puede eliminarse si hay preguntas que la usan."
                   
        Case "SECCIONES"
            ayuda = "SECCIONES" & vbCrLf & vbCrLf & _
                   "Agrupa temas dentro de una auditoría." & vbCrLf & vbCrLf & _
                   "📝 Campos obligatorios:" & vbCrLf & _
                   "• Nombre: Único" & vbCrLf & _
                   "• Tipo: 'Selección' o 'Puntaje'" & vbCrLf & vbCrLf & _
                   "💡 Tipo de Respuesta:" & vbCrLf & _
                   "• Selección: Opciones predefinidas" & vbCrLf & _
                   "• Puntaje: Valor numérico" & vbCrLf & vbCrLf & _
                   "⚠️ Requiere OPCIONES definidas antes de crear preguntas."
                   
        Case "PLANTILLAS"
            ayuda = "PLANTILLAS" & vbCrLf & vbCrLf & _
                   "Modelos reutilizables de inspección." & vbCrLf & vbCrLf & _
                   "📝 Campos obligatorios:" & vbCrLf & _
                   "• Nombre: Único" & vbCrLf & vbCrLf & _
                   "💡 Frecuencia:" & vbCrLf & _
                   "• En MESES (1, 3, 6, 12)" & vbCrLf & _
                   "• Se usa para programación automática" & vbCrLf & vbCrLf & _
                   "🎯 TIP: Use el Wizard para crear plantillas completas más rápido."
                   
        Case "OPCIONES"
            ayuda = "OPCIONES DE RESPUESTA" & vbCrLf & vbCrLf & _
                   "Define respuestas disponibles para cada sección." & vbCrLf & vbCrLf & _
                   "📝 Campos obligatorios:" & vbCrLf & _
                   "• Sección: Debe existir" & vbCrLf & _
                   "• Texto: Único por sección" & vbCrLf & vbCrLf & _
                   "💡 Ejemplos:" & vbCrLf & _
                   "• Conforme (valor: 0)" & vbCrLf & _
                   "• No conforme (valor: -10)" & vbCrLf & vbCrLf & _
                   "⚠️ Debe crear OPCIONES antes de agregar PREGUNTAS."
                   
        Case "PREGUNTAS"
            ayuda = "PREGUNTAS" & vbCrLf & vbCrLf & _
                   "Ítems específicos de cada plantilla." & vbCrLf & vbCrLf & _
                   "📝 Campos obligatorios:" & vbCrLf & _
                   "• Plantilla, Sección, Criticidad" & vbCrLf & _
                   "• Texto de pregunta" & vbCrLf & vbCrLf & _
                   "💡 Orden:" & vbCrLf & _
                   "• Define la secuencia de aparición" & vbCrLf & vbCrLf & _
                   "🎯 TIP:" & vbCrLf & _
                   "• Use 'Duplicar' para preguntas similares" & vbCrLf & _
                   "• Use 'Importar CSV' para agregar muchas a la vez" & vbCrLf & vbCrLf & _
                   "⚠️ Desactive en lugar de eliminar para preservar historial."
    End Select
    
    Me.txtAyuda.Value = ayuda
End Sub

' Actualizar ayuda al hacer foco en campos
Private Sub txtCampo2_Enter()
    If mTablaActual = "CRITICIDAD" Then
        Me.txtAyuda.Value = "NOMBRE DE CRITICIDAD" & vbCrLf & vbCrLf & _
                           "Ingrese un nombre único y descriptivo." & vbCrLf & vbCrLf & _
                           "Ejemplos: Crítico, Alto, Medio-Alto, Medio, Bajo, Mínimo"
    End If
End Sub
```

**Tiempo:** 4-6 horas  
**Impacto:** ⭐⭐⭐⭐ (Alto - Reduce dependencia del manual)

---

## 📊 RESUMEN FASE 2

| Mejora | Tiempo | Impacto | Status |
|--------|--------|---------|--------|
| 2.1 Wizard plantilla completa | 2-3 días | ⭐⭐⭐⭐⭐ | ⬜ Pendiente |
| 2.2 Importación CSV | 1-1.5 días | ⭐⭐⭐⭐⭐ | ⬜ Pendiente |
| 2.3 Mejora visual lista | 1 día | ⭐⭐⭐⭐ | ⬜ Pendiente |
| 2.4 Panel ayuda contextual | 0.5 días | ⭐⭐⭐⭐ | ⬜ Pendiente |
| **TOTAL FASE 2** | **4.5-6 días** | **Crítico** | **⬜ 0%** |

**Resultado esperado:** Sistema 85% más rápido para crear plantillas completas.

---

# 🏗️ FASE 3: REFACTORING ARQUITECTÓNICO (5-6 días)

**Objetivo:** Eliminar switch-cases, implementar metadata-driven design, hacer el sistema extensible.

---

## 3.1 Sistema Metadata-Driven (Prioridad: ALTA)

### Problema que resuelve
❌ Modificar 10+ lugares para agregar una tabla  
✅ Agregar 1 configuración en metadata

### Diseño

**Nuevo módulo:** `TableMetadata.bas`

```vba
' ======================================================================
' Módulo: TableMetadata
' Descripción: Configuración centralizada de todas las tablas
'              (Metadata-driven design)
' ======================================================================
Option Explicit

' Tipo: Configuración de campo
Public Type FieldConfig
    NombreLogico As String      ' "ID", "Nombre", "Valor"
    NombreColumnaExcel As String ' "ID Criticidad", "Nombre de criticidad"
    TipoControl As String       ' "TXT", "CMB", "CHK"
    Obligatorio As Boolean
    SoloLectura As Boolean
    LookupTabla As String       ' Para ComboBox: tabla origen
    Validacion As String        ' "NUMERIC", "TEXT", "OPTION:Selección|Puntaje"
    Tooltip As String
    Orden As Integer
End Type

' Tipo: Configuración de tabla completa
Public Type TableConfig
    NombreLogico As String
    NombreExcel As String
    HojaExcel As String
    Prefijo As String          ' Para GenerarNuevoID
    Campos() As FieldConfig
    NumCampos As Integer
    TieneCampoActivo As Boolean
    PermiteEliminar As Boolean
    ListBoxColumns As String   ' ColumnWidths
    ListBoxColumnCount As Integer
End Type

' ======================================================================
' REGISTRO DE TABLAS
' ======================================================================
Private mTableRegistry As Object  ' Dictionary de TableConfig

Public Sub InicializarMetadata()
    Set mTableRegistry = CreateObject("Scripting.Dictionary")
    
    ' Registrar cada tabla
    Call RegistrarCriticidad
    Call RegistrarSecciones
    Call RegistrarPlantillas
    Call RegistrarOpciones
    Call RegistrarPreguntas
End Sub

' ======================================================================
' CONFIGURACIÓN: CRITICIDAD
' ======================================================================
Private Sub RegistrarCriticidad()
    Dim cfg As TableConfig
    cfg.NombreLogico = "CRITICIDAD"
    cfg.NombreExcel = Configuration2.TABLE_CRITICIDAD
    cfg.HojaExcel = Configuration2.SHEET_CHECKLIST
    cfg.Prefijo = "CRT"
    cfg.TieneCampoActivo = False
    cfg.PermiteEliminar = True
    cfg.ListBoxColumns = "0;280;80"
    cfg.ListBoxColumnCount = 3
    cfg.NumCampos = 3
    
    ReDim cfg.Campos(1 To 3)
    
    ' Campo 1: ID
    With cfg.Campos(1)
        .NombreLogico = "ID"
        .NombreColumnaExcel = "ID Criticidad"
        .TipoControl = "TXT"
        .Obligatorio = True
        .SoloLectura = True
        .Tooltip = "ID único generado automáticamente"
        .Orden = 1
    End With
    
    ' Campo 2: Nombre
    With cfg.Campos(2)
        .NombreLogico = "Nombre"
        .NombreColumnaExcel = "Nombre de criticidad"
        .TipoControl = "TXT"
        .Obligatorio = True
        .SoloLectura = False
        .Validacion = "TEXT"
        .Tooltip = "Nombre único de la criticidad (ej: Alto, Medio, Bajo)"
        .Orden = 2
    End With
    
    ' Campo 3: Valor
    With cfg.Campos(3)
        .NombreLogico = "Valor"
        .NombreColumnaExcel = "Valor"
        .TipoControl = "TXT"
        .Obligatorio = False
        .SoloLectura = False
        .Validacion = "NUMERIC"
        .Tooltip = "Valor numérico positivo (1-10 recomendado)"
        .Orden = 3
    End With
    
    mTableRegistry.Add "CRITICIDAD", cfg
End Sub

' ======================================================================
' CONFIGURACIÓN: SECCIONES
' ======================================================================
Private Sub RegistrarSecciones()
    Dim cfg As TableConfig
    cfg.NombreLogico = "SECCIONES"
    cfg.NombreExcel = Configuration2.TABLE_SECCIONES
    cfg.HojaExcel = Configuration2.SHEET_CHECKLIST
    cfg.Prefijo = "SEC"
    cfg.TieneCampoActivo = False
    cfg.PermiteEliminar = True
    cfg.ListBoxColumns = "0;300;100"
    cfg.ListBoxColumnCount = 3
    cfg.NumCampos = 3
    
    ReDim cfg.Campos(1 To 3)
    
    ' Campo 1: ID
    With cfg.Campos(1)
        .NombreLogico = "ID"
        .NombreColumnaExcel = "ID Seccion"
        .TipoControl = "TXT"
        .Obligatorio = True
        .SoloLectura = True
        .Tooltip = "ID único generado automáticamente"
        .Orden = 1
    End With
    
    ' Campo 2: Nombre
    With cfg.Campos(2)
        .NombreLogico = "Nombre"
        .NombreColumnaExcel = "Nombre de sección"
        .TipoControl = "TXT"
        .Obligatorio = True
        .SoloLectura = False
        .Validacion = "TEXT"
        .Tooltip = "Nombre de la sección (ej: Auditoría de procesos)"
        .Orden = 2
    End With
    
    ' Campo 3: Tipo Respuesta
    With cfg.Campos(3)
        .NombreLogico = "TipoRespuesta"
        .NombreColumnaExcel = "Tipo de respuesta"
        .TipoControl = "TXT"
        .Obligatorio = True
        .SoloLectura = False
        .Validacion = "OPTION:Selección|Puntaje"
        .Tooltip = "Tipo: 'Selección' o 'Puntaje'"
        .Orden = 3
    End With
    
    mTableRegistry.Add "SECCIONES", cfg
End Sub

' (Continuar con RegistrarPlantillas, RegistrarOpciones, RegistrarPreguntas...)

' ======================================================================
' FUNCIONES PÚBLICAS
' ======================================================================

Public Function GetConfig(ByVal nombreLogico As String) As TableConfig
    If mTableRegistry Is Nothing Then Call InicializarMetadata
    
    If mTableRegistry.Exists(nombreLogico) Then
        GetConfig = mTableRegistry(nombreLogico)
    Else
        Err.Raise vbObjectError + 1, "TableMetadata", _
                 "No existe configuración para tabla: " & nombreLogico
    End If
End Function

Public Function GetAllTableNames() As Variant
    If mTableRegistry Is Nothing Then Call InicializarMetadata
    GetAllTableNames = mTableRegistry.Keys
End Function
```

### Refactorizar frmGestorTablas con Metadata

```vba
' ======================================================================
' VERSIÓN 2.0: Metadata-driven
' ======================================================================

Private Sub ConfigurarCamposParaTabla(ByVal tabla As String)
    ' Obtener configuración de metadata
    Dim cfg As TableConfig
    cfg = TableMetadata.GetConfig(tabla)
    
    ' Ocultar todos los campos
    Call OcultarTodosCampos
    
    ' Renderizar campos dinámicamente según metadata
    Dim i As Integer
    Dim campo As FieldConfig
    For i = 1 To cfg.NumCampos
        campo = cfg.Campos(i)
        Call RenderizarCampo(i, campo)
    Next i
    
    ' Configurar ListBox
    Me.lstDatos.ColumnCount = cfg.ListBoxColumnCount
    Me.lstDatos.ColumnWidths = cfg.ListBoxColumns
    
    ' Mostrar checkbox Activo si aplica
    Me.chkActivo.Visible = cfg.TieneCampoActivo
End Sub

Private Sub RenderizarCampo(ByVal numCampo As Integer, ByRef campo As FieldConfig)
    ' Mostrar label
    Dim lbl As MSForms.Label
    Set lbl = GetLabelControl(numCampo)
    lbl.Visible = True
    lbl.Caption = campo.NombreColumnaExcel & IIf(campo.Obligatorio, " *", "") & ":"
    
    ' Mostrar control según tipo
    Select Case campo.TipoControl
        Case "TXT"
            Dim txt As MSForms.TextBox
            Set txt = GetTextBoxControl(numCampo)
            txt.Visible = True
            txt.Locked = campo.SoloLectura
            txt.ControlTipText = campo.Tooltip
            If campo.SoloLectura Then
                txt.BackColor = COLOR_READONLY
            Else
                txt.BackColor = vbWhite
            End If
            
        Case "CMB"
            Dim cmb As MSForms.ComboBox
            Set cmb = GetComboBoxControl(numCampo)
            cmb.Visible = True
            cmb.Enabled = Not campo.SoloLectura
            cmb.ControlTipText = campo.Tooltip
            
            ' Cargar lookup si tiene
            If campo.LookupTabla <> "" Then
                Call CargarLookup(cmb, campo.LookupTabla)
            End If
            
        Case "CHK"
            Me.chkActivo.Visible = True
            Me.chkActivo.ControlTipText = campo.Tooltip
    End Select
End Sub

' Funciones auxiliares para obtener controles genéricos
Private Function GetLabelControl(ByVal num As Integer) As MSForms.Label
    Select Case num
        Case 1: Set GetLabelControl = Me.lblCampo1
        Case 2: Set GetLabelControl = Me.lblCampo2
        Case 3: Set GetLabelControl = Me.lblCampo3
        Case 4: Set GetLabelControl = Me.lblCampo4
        Case 5: Set GetLabelControl = Me.lblCampo5
        Case 6: Set GetLabelControl = Me.lblCampo6
        Case 7: Set GetLabelControl = Me.lblCampo7
        Case 8: Set GetLabelControl = Me.lblCampo8
    End Select
End Function

Private Function GetTextBoxControl(ByVal num As Integer) As MSForms.TextBox
    Select Case num
        Case 1: Set GetTextBoxControl = Me.txtCampo1
        Case 2: Set GetTextBoxControl = Me.txtCampo2
        Case 3: Set GetTextBoxControl = Me.txtCampo3
        Case 7: Set GetTextBoxControl = Me.txtCampo7
        Case 8: Set GetTextBoxControl = Me.txtCampo8
    End Select
End Function

Private Function GetComboBoxControl(ByVal num As Integer) As MSForms.ComboBox
    Select Case num
        Case 4: Set GetComboBoxControl = Me.cmbCampo4
        Case 5: Set GetComboBoxControl = Me.cmbCampo5
        Case 6: Set GetComboBoxControl = Me.cmbCampo6
    End Select
End Function
```

### Beneficios inmediatos

**Antes (agregar nueva tabla):**
- ✏️ Modificar 10+ lugares
- ✏️ 2-3 horas de trabajo
- ⚠️ Alto riesgo de bugs

**Después (agregar nueva tabla):**
- ✏️ Agregar 1 función `RegistrarNuevaTabla()`
- ✏️ 20-30 minutos de trabajo
- ✅ Bajo riesgo (todo centralizado)

**Tiempo:** 3-4 días  
**Impacto:** ⭐⭐⭐⭐⭐ (Crítico para mantenibilidad)

---

## 3.2 Refactorizar TableManager con Metadata (Prioridad: MEDIA)

### Cambios principales

```vba
' ======================================================================
' VERSIÓN 2.0: TableManager con metadata
' ======================================================================

Public Function ObtenerListObject(ByVal nombreLogico As String) As ListObject
    Dim cfg As TableConfig
    cfg = TableMetadata.GetConfig(nombreLogico)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(cfg.HojaExcel)
    
    On Error Resume Next
    Set ObtenerListObject = ws.ListObjects(cfg.NombreExcel)
    On Error GoTo 0
End Function

Public Function GenerarNuevoID(ByVal nombreLogico As String) As String
    Dim cfg As TableConfig
    cfg = TableMetadata.GetConfig(nombreLogico)
    
    GenerarNuevoID = cfg.Prefijo & "-" & Format(Now, "yyyymmdd-hhnnss") & "-" & _
                     Right("000" & Int(Rnd * 999), 3)
End Function

Private Sub MapearDatosAFila(ByVal nombreLogico As String, _
                              ByVal fila As ListRow, _
                              ByVal datos As Object)
    Dim cfg As TableConfig
    cfg = TableMetadata.GetConfig(nombreLogico)
    
    Dim tbl As ListObject
    Set tbl = fila.Parent
    
    ' Iterar por metadata en lugar de switch-case
    Dim i As Integer
    Dim campo As FieldConfig
    For i = 1 To cfg.NumCampos
        campo = cfg.Campos(i)
        
        If datos.Exists(campo.NombreLogico) Then
            Dim colIdx As Long
            colIdx = ObtenerIndiceColumna(tbl, campo.NombreColumnaExcel)
            
            If colIdx > 0 Then
                fila.Range.Cells(1, colIdx).Value = datos(campo.NombreLogico)
            End If
        End If
    Next i
End Sub
```

**Tiempo:** 2 días  
**Impacto:** ⭐⭐⭐⭐ (Alto)

---

## 📊 RESUMEN FASE 3

| Mejora | Tiempo | Impacto | Status |
|--------|--------|---------|--------|
| 3.1 Sistema metadata-driven | 3-4 días | ⭐⭐⭐⭐⭐ | ⬜ Pendiente |
| 3.2 Refactorizar TableManager | 2 días | ⭐⭐⭐⭐ | ⬜ Pendiente |
| **TOTAL FASE 3** | **5-6 días** | **Crítico** | **⬜ 0%** |

**Resultado esperado:** Sistema 100% extensible, agregar tablas en 30 minutos.

---

# 🚀 FASE 4: FUNCIONALIDADES AVANZADAS (3-4 días)

**Objetivo:** Features profesionales que elevan el sistema a nivel enterprise.

---

## 4.1 Exportación a Excel/PDF (Prioridad: MEDIA)

```vba
Private Sub btnExportar_Click()
    Dim exportPath As String
    exportPath = Application.GetSaveAsFilename("Datos_" & mTablaActual & ".xlsx", _
                                               "Excel (*.xlsx), *.xlsx")
    
    If exportPath = "False" Then Exit Sub
    
    ' Exportar a nuevo workbook
    Dim wbExport As Workbook
    Set wbExport = Workbooks.Add
    
    Dim wsExport As Worksheet
    Set wsExport = wbExport.Worksheets(1)
    wsExport.Name = mTablaActual
    
    ' Copiar datos
    Dim tbl As ListObject
    Set tbl = TableManager.ObtenerListObject(mTablaActual)
    tbl.Range.Copy wsExport.Range("A1")
    
    wbExport.SaveAs exportPath
    wbExport.Close
    
    MsgBox "✓ Datos exportados a: " & exportPath, vbInformation
End Sub
```

**Tiempo:** 1 día  
**Impacto:** ⭐⭐⭐

---

## 4.2 Historial de Cambios Visual (Prioridad: BAJA)

```vba
' Formulario: frmAuditHistory
' Muestra el audit trail en formato legible
Private Sub btnVerHistorial_Click()
    Dim frm As frmAuditHistory
    Set frm = New frmAuditHistory
    frm.CargarHistorialParaTabla mTablaActual, Me.txtCampo1.Value
    frm.Show
End Sub
```

**Tiempo:** 1 día  
**Impacto:** ⭐⭐⭐

---

## 4.3 Plantillas de Preguntas (Biblioteca) (Prioridad: MEDIA)

```vba
' Biblioteca de preguntas comunes
' El usuario puede arrastrar y soltar de una biblioteca
Private Sub btnBiblioteca_Click()
    Dim frm As frmBibliotecaPreguntas
    Set frm = New frmBibliotecaPreguntas
    
    If frm.ShowDialog Then
        Dim preguntasSeleccionadas As Collection
        Set preguntasSeleccionadas = frm.GetPreguntasSeleccionadas
        
        ' Agregar a la plantilla actual
        ' ...
    End If
End Sub
```

**Tiempo:** 1-2 días  
**Impacto:** ⭐⭐⭐⭐

---

## 📊 RESUMEN FASE 4

| Mejora | Tiempo | Impacto | Status |
|--------|--------|---------|--------|
| 4.1 Exportación Excel/PDF | 1 día | ⭐⭐⭐ | ⬜ Pendiente |
| 4.2 Historial visual | 1 día | ⭐⭐⭐ | ⬜ Pendiente |
| 4.3 Biblioteca preguntas | 1-2 días | ⭐⭐⭐⭐ | ⬜ Pendiente |
| **TOTAL FASE 4** | **3-4 días** | **Medio** | **⬜ 0%** |

---

# 📈 RESUMEN TOTAL DEL PLAN

## Cronograma General

| Fase | Duración | Complejidad | Prioridad | ROI |
|------|----------|-------------|-----------|-----|
| Fase 1: Quick Wins | 2-3 días | Baja | 🔴 Crítica | Muy Alto |
| Fase 2: UX Profesional | 4-5 días | Media | 🟠 Alta | Alto |
| Fase 3: Refactoring | 5-6 días | Alta | 🟡 Media | Medio |
| Fase 4: Avanzado | 3-4 días | Media | 🟢 Baja | Medio |
| **TOTAL** | **14-18 días** | - | - | - |

## Comparativa Antes/Después

| Métrica | Antes | Después Fase 1 | Después Fase 2 | Después Completo |
|---------|-------|----------------|----------------|------------------|
| Calificación | 4/10 | 6/10 | 8/10 | 9/10 |
| Tiempo crear plantilla (20 preguntas) | 90 min | 60 min | 15 min | 10 min |
| Errores de usuario | Alto | Medio | Bajo | Muy Bajo |
| Tiempo mantenimiento (agregar tabla) | 8 horas | 8 horas | 4 horas | 30 min |
| Curva de aprendizaje | Muy Alta | Alta | Media | Baja |

## Recomendación de Ejecución

### Estrategia Agresiva (Máximo Impacto)
```
Semana 1: Fase 1 (2 días) + Iniciar Fase 2 (3 días)
Semana 2: Completar Fase 2 (2 días) + Iniciar Fase 3 (3 días)
Semana 3: Completar Fase 3 (3 días) + Fase 4 (2 días)
Semana 4: Completar Fase 4 (2 días) + Testing (3 días)
```

### Estrategia Conservadora (Valor Incremental)
```
Mes 1: Fase 1 → Deploy → Recopilar feedback
Mes 2: Fase 2 → Deploy → Recopilar feedback
Mes 3: Fase 3 (si se requiere extensibilidad)
Mes 4: Fase 4 (opcional)
```

### Estrategia Mínima Viable
```
Solo Fase 1 + Item 2.1 (Wizard) de Fase 2
Total: 4-5 días
Mejora inmediata: 70% del beneficio total
```

---

# ✅ CHECKLIST DE IMPLEMENTACIÓN

## Pre-requisitos
- [ ] Backup completo del archivo .xlsm actual
- [ ] Crear rama de testing en control de versiones (si aplica)
- [ ] Documentar configuración actual
- [ ] Preparar datos de prueba

## Fase 1
- [ ] 1.1 Búsqueda y filtrado
- [ ] 1.2 Tooltips y ayuda
- [ ] 1.3 Validación en tiempo real
- [ ] 1.4 Atajos de teclado
- [ ] 1.5 Botón duplicar
- [ ] 1.6 Corregir inconsistencias
- [ ] 1.7 Indicadores visuales
- [ ] Testing Fase 1
- [ ] Deploy Fase 1

## Fase 2
- [ ] 2.1 Wizard plantilla
- [ ] 2.2 Importación CSV
- [ ] 2.3 Mejora visual lista
- [ ] 2.4 Panel ayuda contextual
- [ ] Testing Fase 2
- [ ] Deploy Fase 2

## Fase 3
- [ ] 3.1 Sistema metadata
- [ ] 3.2 Refactorizar TableManager
- [ ] Testing Fase 3
- [ ] Deploy Fase 3

## Fase 4
- [ ] 4.1 Exportación
- [ ] 4.2 Historial visual
- [ ] 4.3 Biblioteca preguntas
- [ ] Testing Fase 4
- [ ] Deploy Fase 4

## Post-implementación
- [ ] Actualizar manual de usuario
- [ ] Capacitar usuarios
- [ ] Monitorear adopción
- [ ] Recopilar feedback
- [ ] Iterar mejoras

---

# 🎯 MÉTRICAS DE ÉXITO

## KPIs a Medir

### Eficiencia Operativa
- **Tiempo promedio para crear plantilla con 20 preguntas**
  - Target: < 15 minutos (vs 90 minutos actual)
- **Número de clicks para operación común**
  - Target: Reducción del 60%

### Calidad
- **Tasa de errores de usuario**
  - Target: < 5% (vs ~25% actual estimado)
- **Registros duplicados por error**
  - Target: 0

### Adopción
- **Porcentaje de usuarios que usan el gestor regularmente**
  - Target: > 80% (vs ~40% estimado)
- **Tiempo de capacitación de nuevo usuario**
  - Target: < 30 minutos (vs 2+ horas)

### Mantenibilidad
- **Tiempo para agregar nueva tabla al sistema**
  - Target: < 1 hora (vs 8 horas)

---

# 📞 SIGUIENTE PASO

**¿Desea que comience la implementación?**

Opciones:
1. **Comenzar con Fase 1 completa** (mejoras inmediatas)
2. **Solo implementar Quick Wins críticos** (1.1, 1.6, 2.1)
3. **Crear prototipo del Wizard** (Demostración de valor)
4. **Diseñar metadata completo** (Planning detallado Fase 3)

**Confirme para proceder con la implementación.**
