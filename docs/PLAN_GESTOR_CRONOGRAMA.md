# PLAN DE IMPLEMENTACIÓN: GESTOR DE CRONOGRAMA

**Fecha de creación:** 27 de abril de 2026  
**Objetivo:** Implementar funcionalidad para pausar/reactivar inspecciones en el cronograma  
**Estimación total:** 7 horas  

---

## PRINCIPIOS DE ARQUITECTURA

Este proyecto debe seguir estrictamente:

- ✅ **Clean Architecture**: Separación de capas (UI → Service → Repository)
- ✅ **DRY (Don't Repeat Yourself)**: Reutilizar código existente, no duplicar lógica
- ✅ **Pipelines**: Flujo de datos unidireccional y predecible
- ✅ **Mantenibilidad**: Código legible, funciones pequeñas, responsabilidades únicas
- ✅ **Escalabilidad**: Diseño que permita agregar nuevos estados sin refactorizar
- ✅ **Desacoplamiento**: Bajo acoplamiento entre módulos, alta cohesión dentro de módulos

---

## VERIFICACIONES CRÍTICAS

Antes de cada commit, verificar que NO se dañen los siguientes flujos:

1. **Flujo de programación de inspecciones** (InspectionScheduler)
2. **Flujo de ejecución de inspecciones** (ChecklistOrchestrator)
3. **Flujo de cálculo de métricas** (InspectionCalculator, RecurrentInspectionCalculator)
4. **Flujo de auditoría** (AuditLogger2, TableAuditor2)
5. **Flujo de generación de certificados** (CertificadoPDFGenerator)
6. **Flujo de búsqueda de historial** (InspectionHistoryService - DEBE encontrar pausadas)
7. **Flujo de resumen de cronograma** (CronogramaResumen)
8. **Flujo de navegación** (NavigationService2)

---

## FASE 1: PREPARACIÓN DE BASE DE DATOS (30 min)

### 1.1 Modificación en Excel (USUARIO)
**⚠️ SOLICITAR AL USUARIO:**
- Abrir `TH-HC-001 INSPECCIONES.xlsm`
- En tabla `tblCronogramaInspecciones`, agregar nueva columna en posición 20:
  - **Nombre:** `Activo en cronograma`
  - **Tipo:** Texto
  - **Valores válidos:** `"Si"` o `"No"`
  - **Valor por defecto:** `"Si"` (para todas las filas existentes)
  - **Ubicación:** Después de columna 19, antes de cualquier columna de auditoría

### 1.2 Actualización de Configuration2.bas (AGENTE)

**Archivo:** `Configuration2.bas`

**Tareas:**
1. Agregar nueva constante de contraseña:
```vba
Public Const CRONOGRAMA_ADMIN_PASSWORD As String = "CronoAdmin2026*"
```

2. Actualizar documentación de columnas de `tblCronogramaInspecciones`:
```vba
' Columna 20: Activo en cronograma (String) - "Si"/"No" - Indica si la inspección está activa en el cronograma
```

3. Actualizar constante de total de columnas:
```vba
Public Const CRONOGRAMA_TOTAL_COLUMNS As Integer = 20  ' Actualizar según nueva columna
```

**Verificaciones:**
- ✅ Compilación sin errores
- ✅ Constante de contraseña no vacía
- ✅ Documentación completa y actualizada

**Commit:**
```
feat(config): agregar soporte para columna Activo en cronograma

- Nueva constante CRONOGRAMA_ADMIN_PASSWORD para gestor
- Documentación de columna 20 en tblCronogramaInspecciones
- Actualización de CRONOGRAMA_TOTAL_COLUMNS
```

---

## FASE 2: CAPA DE SERVICIO - BUSINESS LOGIC (1 hora)

### 2.1 Crear CronogramaGestorService.bas (AGENTE)

**Archivo:** `CronogramaGestorService.bas` (NUEVO)

**Funciones a implementar:**

#### 2.1.1 ValidarContraseñaGestor
```vba
Public Function ValidarContraseñaGestor(ByVal contraseña As String) As Boolean
    ' Validar contraseña contra Configuration2.CRONOGRAMA_ADMIN_PASSWORD
    ' Registrar intento en AuditLogger2 (categoría: CRONOGRAMA_ACCESO)
End Function
```

#### 2.1.2 ObtenerDatosParaGestor
```vba
Public Function ObtenerDatosParaGestor() As Variant
    ' Retornar array 2D con todas las filas de tblCronogramaInspecciones
    ' Columnas necesarias: ID, NombreControl, Planta, Frecuencia, ActivoEnCronograma, Responsable
    ' Ordenar por: Planta ASC, NombreControl ASC
    ' Usar ChecklistRepository para acceso a datos
    ' NO filtrar por estado - mostrar todas
End Function
```

#### 2.1.3 PausarInspecciones
```vba
Public Function PausarInspecciones(ByVal idsInspecciones() As String) As Long
    ' Para cada ID en array:
    '   1. Validar que existe en tblCronogramaInspecciones
    '   2. Leer estado actual de "Activo en cronograma"
    '   3. Si ya está en "No", contar como advertencia pero continuar
    '   4. Actualizar columna "Activo en cronograma" a "No"
    '   5. Registrar en AuditLogger2 (categoría: CRONOGRAMA_PAUSAR)
    ' Retornar: número de filas actualizadas exitosamente
    ' Usar ChecklistRepository.ActualizarCampo() para cambios
End Function
```

#### 2.1.4 ReactivarInspecciones
```vba
Public Function ReactivarInspecciones(ByVal idsInspecciones() As String) As Long
    ' Para cada ID en array:
    '   1. Validar que existe en tblCronogramaInspecciones
    '   2. Leer estado actual de "Activo en cronograma"
    '   3. Si ya está en "Si", contar como advertencia pero continuar
    '   4. Actualizar columna "Activo en cronograma" a "Si"
    '   5. Registrar en AuditLogger2 (categoría: CRONOGRAMA_REACTIVAR)
    ' Retornar: número de filas actualizadas exitosamente
    ' Usar ChecklistRepository.ActualizarCampo() para cambios
End Function
```

#### 2.1.5 ContarInspeccionesPorEstado
```vba
Public Function ContarInspeccionesPorEstado(ByVal estado As String) As Long
    ' Contar filas en tblCronogramaInspecciones donde "Activo en cronograma" = estado
    ' Estado puede ser "Si", "No", o "" (todas)
    ' Usar ChecklistRepository para consulta
End Function
```

**Principios de diseño:**
- ✅ Single Responsibility: cada función una sola responsabilidad
- ✅ Dependency Injection: usar ChecklistRepository, no acceso directo a tablas
- ✅ Error Handling: On Error GoTo con ErrorLogger2
- ✅ Auditoría: todos los cambios registrados en AuditLogger2
- ✅ Transaccionalidad: si falla una actualización, loguear pero continuar con las demás

**Verificaciones:**
- ✅ Compilación sin errores
- ✅ Todas las funciones con manejo de errores
- ✅ Integración con AuditLogger2 (categorías: CRONOGRAMA_ACCESO, CRONOGRAMA_PAUSAR, CRONOGRAMA_REACTIVAR)
- ✅ No hay acceso directo a hojas Excel (usar Repository)
- ✅ Funciones retornan valores útiles para UI

**Commit:**
```
feat(service): implementar CronogramaGestorService para gestión de cronograma

- ValidarContraseñaGestor: autenticación con CRONOGRAMA_ADMIN_PASSWORD
- ObtenerDatosParaGestor: consulta de datos para UI
- PausarInspecciones: cambiar estado a "No" con auditoría
- ReactivarInspecciones: cambiar estado a "Si" con auditoría
- ContarInspeccionesPorEstado: estadísticas de cronograma
- Integración completa con AuditLogger2 y ChecklistRepository
```

---

## FASE 3: CAPA DE PRESENTACIÓN - USERFORM (2 horas)

### 3.1 Crear frmGestorCronograma.frm (AGENTE)

**Archivo:** `frmGestorCronograma.frm` (NUEVO)

**Propiedades del Form:**
- **Caption:** "Gestor de Cronograma de Inspecciones"
- **Width:** 900
- **Height:** 600
- **StartUpPosition:** CenterScreen
- **ShowModal:** True (vbModal)

**Controles:**

#### 3.1.1 Filtros (Superior)
```vba
' Label: lblFiltroPlanta
' ComboBox: cboFiltroPlanta (valores: "Todas", "Planta 1", "Planta 2", etc.)
' Label: lblFiltroEstado
' ComboBox: cboFiltroEstado (valores: "Todos", "Activas", "Pausadas")
```

#### 3.1.2 Lista de Inspecciones (Centro)
```vba
' ListBox: lstCronograma
'   - MultiSelect: fmMultiSelectMulti (selección múltiple ilimitada)
'   - ColumnCount: 7
'   - ColumnHeads: True
'   - Columnas: ID | Nombre Control | Planta | Frecuencia | Estado | Responsable | Última Ejecución
'   - Width: 850
'   - Height: 400
```

#### 3.1.3 Estadísticas (Inferior izquierdo)
```vba
' Label: lblEstadisticas
' TextBox: txtTotalActivas (ReadOnly)
' TextBox: txtTotalPausadas (ReadOnly)
' TextBox: txtSeleccionadas (ReadOnly)
```

#### 3.1.4 Botones de Acción (Inferior derecho)
```vba
' CommandButton: btnPausar (Caption: "Pausar Seleccionadas")
' CommandButton: btnReactivar (Caption: "Reactivar Seleccionadas")
' CommandButton: btnRefrescar (Caption: "Refrescar")
' CommandButton: btnCerrar (Caption: "Cerrar")
```

**Eventos a implementar:**

#### 3.1.5 UserForm_Initialize
```vba
Private Sub UserForm_Initialize()
    ' 1. Configurar lstCronograma (columnas, headers)
    ' 2. Cargar cboFiltroPlanta con plantas únicas de tblCronogramaInspecciones
    ' 3. Cargar cboFiltroEstado con valores predefinidos
    ' 4. Establecer valores por defecto en combos
    ' 5. Llamar a CargarDatos()
    ' 6. Actualizar estadísticas
End Sub
```

#### 3.1.6 CargarDatos
```vba
Private Sub CargarDatos()
    ' 1. Llamar a CronogramaGestorService.ObtenerDatosParaGestor()
    ' 2. Aplicar filtros (planta, estado)
    ' 3. Cargar datos filtrados en lstCronograma
    ' 4. Actualizar txtTotalActivas y txtTotalPausadas
    ' Principio: Separation of Concerns - UI solo renderiza, Service provee datos
End Sub
```

#### 3.1.7 cboFiltroPlanta_Change / cboFiltroEstado_Change
```vba
Private Sub cboFiltroPlanta_Change()
    CargarDatos()
End Sub

Private Sub cboFiltroEstado_Change()
    CargarDatos()
End Sub
```

#### 3.1.8 lstCronograma_Click
```vba
Private Sub lstCronograma_Click()
    ' Actualizar txtSeleccionadas con número de filas seleccionadas
    ' Habilitar/deshabilitar btnPausar y btnReactivar según selección
End Sub
```

#### 3.1.9 btnPausar_Click
```vba
Private Sub btnPausar_Click()
    ' 1. Validar que hay selección
    ' 2. Obtener IDs de filas seleccionadas
    ' 3. Mostrar confirmación: "¿Pausar X inspecciones?"
    ' 4. Si confirma:
    '    a. Llamar a CronogramaGestorService.PausarInspecciones(ids)
    '    b. Mostrar resultado: "X inspecciones pausadas"
    '    c. Llamar a CargarDatos() para refrescar
    '    d. Limpiar selección
    ' Principio: Pipeline - UI → Service → Repository → Audit
End Sub
```

#### 3.1.10 btnReactivar_Click
```vba
Private Sub btnReactivar_Click()
    ' 1. Validar que hay selección
    ' 2. Obtener IDs de filas seleccionadas
    ' 3. Mostrar confirmación: "¿Reactivar X inspecciones?"
    ' 4. Si confirma:
    '    a. Llamar a CronogramaGestorService.ReactivarInspecciones(ids)
    '    b. Mostrar resultado: "X inspecciones reactivadas"
    '    c. Llamar a CargarDatos() para refrescar
    '    d. Limpiar selección
End Sub
```

#### 3.1.11 btnRefrescar_Click
```vba
Private Sub btnRefrescar_Click()
    CargarDatos()
    lstCronograma.ListIndex = -1  ' Limpiar selección
End Sub
```

#### 3.1.12 btnCerrar_Click
```vba
Private Sub btnCerrar_Click()
    Unload Me
End Sub
```

**Principios de diseño:**
- ✅ Separation of Concerns: UI solo maneja presentación, Service maneja lógica
- ✅ Event-Driven: reacciona a cambios de usuario
- ✅ Responsive: filtros y estadísticas en tiempo real
- ✅ User Feedback: confirmaciones y mensajes de resultado
- ✅ Validación: deshabilitar botones cuando no hay selección

**Verificaciones:**
- ✅ Compilación sin errores
- ✅ Todos los controles nombrados correctamente
- ✅ Eventos conectados correctamente
- ✅ No hay lógica de negocio en UI (solo llamadas a Service)
- ✅ Manejo de errores en todos los eventos
- ✅ Mensajes de usuario claros y en español

**Commit:**
```
feat(ui): implementar frmGestorCronograma para gestión visual de cronograma

- UserForm modal 900x600 con filtros de planta y estado
- ListBox multiselección con 7 columnas de información
- Estadísticas en tiempo real (activas/pausadas/seleccionadas)
- Botones para pausar, reactivar, refrescar y cerrar
- Filtrado dinámico por planta y estado
- Integración con CronogramaGestorService
- Confirmaciones y feedback al usuario
```

---

## FASE 4: INTEGRACIÓN CON MENÚ PRINCIPAL (30 min)

### 4.1 Actualización de Hoja1.bas (AGENTE)

**Archivo:** `Hoja1.bas`

**Tareas:**

#### 4.1.1 Agregar procedimiento de apertura
```vba
Public Sub AbrirGestorCronograma()
    ' 1. Mostrar InputBox para contraseña
    ' 2. Validar con CronogramaGestorService.ValidarContraseñaGestor()
    ' 3. Si válida:
    '    a. frmGestorCronograma.Show vbModal
    '    b. Al cerrar form, refrescar tblResumenCronograma si existe
    ' 4. Si inválida:
    '    a. Mostrar error "Contraseña incorrecta"
    '    b. Registrar intento fallido en AuditLogger2
    ' Principio: Seguridad en capas - validación antes de UI
End Sub
```

#### 4.1.2 Documentación
```vba
' Agregar comentario en encabezado de módulo:
' - AbrirGestorCronograma: Punto de entrada para gestión de cronograma (requiere contraseña admin)
```

**Verificaciones:**
- ✅ Validación de contraseña funciona correctamente
- ✅ Form se abre en modo modal
- ✅ Refrescar resumen después de cerrar form
- ✅ Manejo de errores completo
- ✅ Auditoría de intentos de acceso

### 4.2 Modificación en Excel (USUARIO)

**⚠️ SOLICITAR AL USUARIO:**
- Abrir hoja `Menú Principal`
- Agregar botón cerca de `tblResumenCronograma`:
  - **Texto:** "Gestor de Cronograma"
  - **Macro asignada:** `Hoja1.AbrirGestorCronograma`
  - **Ubicación:** Junto a otros botones de gestión
  - **Estilo:** Consistente con botones existentes

**Commit:**
```
feat(integration): integrar gestor de cronograma con menú principal

- Nuevo procedimiento AbrirGestorCronograma en Hoja1
- Validación de contraseña antes de mostrar form
- Refrescar resumen al cerrar gestor
- Auditoría de accesos (exitosos y fallidos)
- Documentación actualizada
```

---

## FASE 5: ACTUALIZACIÓN DE MÓDULOS EXISTENTES (1 hora)

### 5.1 CronogramaResumen.bas (AGENTE)

**Archivo:** `CronogramaResumen.bas`

**Función a modificar:** `RefrescarResumenCronograma()`

**Cambios:**
```vba
' ANTES: Consulta todas las filas de tblCronogramaInspecciones
' DESPUÉS: Filtrar solo donde "Activo en cronograma" = "Si"

' Agregar filtro en consulta:
If ws.ListObjects("tblCronogramaInspecciones").ListColumns("Activo en cronograma").DataBodyRange(i) = "Si" Then
    ' ... procesar fila ...
End If
```

**Principio:**
- ✅ Backward Compatible: filas existentes sin la columna deben funcionar (tratar como "Si")
- ✅ DRY: reutilizar lógica existente, solo agregar filtro
- ✅ Single Responsibility: módulo sigue enfocado en resumen

**Verificaciones:**
- ✅ Resumen solo muestra inspecciones activas
- ✅ Inspecciones pausadas NO aparecen en resumen
- ✅ Flujo existente no se rompe
- ✅ Performance no se degrada

### 5.2 InspectionScheduler.bas (AGENTE)

**Archivo:** `InspectionScheduler.bas`

**Función a modificar:** `ProgramarInspeccion()` o similar

**Cambios:**
```vba
' Al crear nueva fila en tblCronogramaInspecciones:
' Establecer valor por defecto "Si" en columna "Activo en cronograma"

nuevaFila.Range(ws.ListObjects("tblCronogramaInspecciones").ListColumns("Activo en cronograma").Index) = "Si"
```

**Principio:**
- ✅ Explicit Defaults: nuevas inspecciones activas por defecto
- ✅ Consistency: todas las filas tienen valor definido

**Verificaciones:**
- ✅ Nuevas inspecciones se crean con "Activo en cronograma" = "Si"
- ✅ Flujo de programación existente funciona correctamente
- ✅ No hay efectos secundarios

### 5.3 InspectionHistoryService.bas (NO MODIFICAR)

**⚠️ IMPORTANTE: NO TOCAR ESTE MÓDULO**

**Razón:**
- Este módulo debe encontrar inspecciones pausadas para mostrar historial completo
- El filtro solo aplica en resumen y programación, NO en búsqueda de historial
- Mantener funcionalidad de búsqueda intacta

**Verificación:**
- ✅ InspectionHistoryService.BuscarHistorialInspeccion() sigue retornando pausadas

**Commit:**
```
feat(cronograma): filtrar inspecciones pausadas en resumen y programación

- CronogramaResumen: solo mostrar inspecciones con Activo="Si"
- InspectionScheduler: establecer "Si" por defecto en nuevas inspecciones
- InspectionHistoryService: NO modificado (debe encontrar todas)
- Backward compatible: filas sin columna tratadas como activas
```

---

## FASE 6: TESTING INTEGRAL (1.5 horas)

### 6.1 Casos de Prueba

#### Test 1: Autenticación
- ✅ Contraseña correcta → abre form
- ✅ Contraseña incorrecta → error + auditoría
- ✅ Cancelar InputBox → no abre form

#### Test 2: Carga de Datos
- ✅ Filtro "Todas las plantas" → muestra todas
- ✅ Filtro "Planta específica" → muestra solo esa planta
- ✅ Filtro "Activas" → solo muestra "Si"
- ✅ Filtro "Pausadas" → solo muestra "No"
- ✅ Combinar filtros → funciona correctamente

#### Test 3: Pausar Inspecciones
- ✅ Seleccionar 1 inspección activa → pausa correctamente
- ✅ Seleccionar múltiples inspecciones activas → pausa todas
- ✅ Seleccionar inspección ya pausada → advierte pero continúa
- ✅ Cancelar confirmación → no realiza cambios
- ✅ Registro en AuditLogger2 → correcto

#### Test 4: Reactivar Inspecciones
- ✅ Seleccionar 1 inspección pausada → reactiva correctamente
- ✅ Seleccionar múltiples inspecciones pausadas → reactiva todas
- ✅ Seleccionar inspección ya activa → advierte pero continúa
- ✅ Cancelar confirmación → no realiza cambios
- ✅ Registro en AuditLogger2 → correcto

#### Test 5: Estadísticas
- ✅ Totales actualizados al cargar → correcto
- ✅ Totales actualizados después de pausar → correcto
- ✅ Totales actualizados después de reactivar → correcto
- ✅ Contador de seleccionadas → correcto

#### Test 6: Integración con Resumen
- ✅ Pausar inspección → desaparece de tblResumenCronograma
- ✅ Reactivar inspección → aparece en tblResumenCronograma
- ✅ Programar nueva inspección → aparece activa por defecto

#### Test 7: Verificación de Flujos Existentes
- ✅ Programar inspección → funciona
- ✅ Ejecutar inspección → funciona
- ✅ Calcular métricas → funciona
- ✅ Generar certificado → funciona
- ✅ Buscar historial → encuentra pausadas también
- ✅ Auditoría general → funciona

### 6.2 Verificación de Arquitectura

- ✅ Capa UI (frmGestorCronograma) NO accede directamente a tablas
- ✅ Capa Service (CronogramaGestorService) NO accede directamente a hojas
- ✅ Capa Repository (ChecklistRepository) es la única que accede a Excel
- ✅ Auditoría en todas las operaciones críticas
- ✅ Manejo de errores en todas las capas
- ✅ No hay código duplicado (DRY)

### 6.3 Verificación de Performance

- ✅ Cargar 100+ inspecciones → < 2 segundos
- ✅ Filtrar datos → instantáneo
- ✅ Pausar 20+ inspecciones → < 5 segundos
- ✅ Refrescar resumen → < 3 segundos

**Commit:**
```
test: validar funcionalidad completa de gestor de cronograma

- 7 suites de pruebas ejecutadas exitosamente
- Verificación de arquitectura: Clean Architecture cumplida
- Verificación de flujos: ningún flujo existente dañado
- Verificación de performance: todos los umbrales cumplidos
- Auditoría completa funcionando correctamente
```

---

## FASE 7: DOCUMENTACIÓN FINAL Y LIMPIEZA (30 min)

### 7.1 Actualizar README.md (AGENTE)

**Archivo:** `README.md`

**Agregar sección:**
```markdown
## Gestor de Cronograma

Funcionalidad para pausar y reactivar inspecciones programadas.

### Acceso
- Desde Menú Principal → Botón "Gestor de Cronograma"
- Requiere contraseña de administrador

### Funcionalidades
- Pausar inspecciones: oculta del resumen, no afecta historial
- Reactivar inspecciones: vuelve a mostrar en resumen
- Filtros: por planta y estado
- Selección múltiple ilimitada
- Auditoría completa de cambios

### Arquitectura
- **UI:** frmGestorCronograma.frm
- **Service:** CronogramaGestorService.bas
- **Repository:** ChecklistRepository.bas (existente)
- **Audit:** AuditLogger2.bas (existente)
```

### 7.2 Actualizar INDICE_RAPIDO_REFERENCIAS.md (AGENTE)

**Archivo:** `docs/INDICE_RAPIDO_REFERENCIAS.md`

**Agregar entrada:**
```markdown
- **Gestor de Cronograma:** `docs/PLAN_GESTOR_CRONOGRAMA.md`
  - Pausar/reactivar inspecciones programadas
  - Arquitectura, testing y verificaciones
```

### 7.3 Limpiar archivos temporales (AGENTE)

**Verificar que NO se hayan creado:**
- ❌ Resúmenes ejecutivos
- ❌ Guías de usuario adicionales
- ❌ Archivos de explicación
- ❌ Logs de debugging en workspace

**Solo deben existir:**
- ✅ Código fuente (.bas, .frm)
- ✅ Este plan (PLAN_GESTOR_CRONOGRAMA.md)
- ✅ README.md actualizado
- ✅ INDICE_RAPIDO_REFERENCIAS.md actualizado

**Commit final:**
```
docs: actualizar documentación para gestor de cronograma

- README.md: sección de gestor de cronograma
- INDICE_RAPIDO_REFERENCIAS.md: referencia a plan
- Sin archivos adicionales innecesarios
```

---

## RESUMEN DE COMMITS

Total de commits esperados: **7 commits** (uno por fase)

1. `feat(config): agregar soporte para columna Activo en cronograma`
2. `feat(service): implementar CronogramaGestorService para gestión de cronograma`
3. `feat(ui): implementar frmGestorCronograma para gestión visual de cronograma`
4. `feat(integration): integrar gestor de cronograma con menú principal`
5. `feat(cronograma): filtrar inspecciones pausadas en resumen y programación`
6. `test: validar funcionalidad completa de gestor de cronograma`
7. `docs: actualizar documentación para gestor de cronograma`

---

## INTERACCIONES CON USUARIO

El agente debe solicitar al usuario las siguientes acciones en Excel:

### Solicitud 1 (Fase 1.1):
```
Por favor, abre TH-HC-001 INSPECCIONES.xlsm y agrega una nueva columna a la tabla tblCronogramaInspecciones:
- Posición: Columna 20 (después de la columna 19)
- Nombre: "Activo en cronograma"
- Tipo: Texto
- Valores válidos: "Si" o "No"
- Valor por defecto: Llenar todas las filas existentes con "Si"

Cuando esté listo, responde "OK" para continuar.
```

### Solicitud 2 (Fase 4.2):
```
Por favor, abre la hoja "Menú Principal" en el Excel y agrega un botón:
- Texto: "Gestor de Cronograma"
- Macro asignada: Hoja1.AbrirGestorCronograma
- Ubicación: Cerca de tblResumenCronograma
- Estilo: Consistente con los botones existentes

Cuando esté listo, responde "OK" para continuar.
```

---

## CRITERIOS DE ÉXITO

✅ **Funcionalidad:**
- Pausar inspecciones oculta del resumen
- Reactivar inspecciones muestra en resumen
- Historial sigue mostrando inspecciones pausadas
- Auditoría completa de todas las operaciones

✅ **Arquitectura:**
- Clean Architecture respetada
- DRY: sin código duplicado
- Desacoplamiento: UI → Service → Repository
- Escalabilidad: fácil agregar nuevos estados

✅ **Calidad:**
- Sin errores de compilación
- Sin warnings
- Todos los flujos existentes funcionando
- Performance aceptable (<5 seg operaciones)

✅ **Documentación:**
- Plan detallado (este archivo)
- README actualizado
- Índice actualizado
- Sin archivos adicionales

---

**FIN DEL PLAN**
