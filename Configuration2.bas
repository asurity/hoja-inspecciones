
' ----------------------------------------------------------------------
' Módulo: Configuration
' Descripción: Centraliza las configuraciones globales y contraseñas de la aplicación.
'              Permite un mantenimiento sencillo y evita la duplicidad de valores sensibles.
'              Todas las contraseñas y parámetros globales deben declararse aquí.
' Dependencias: Ninguna
' ----------------------------------------------------------------------
Option Explicit

' ============================================================================
' CONFIGURACIÓN DE SEGURIDAD
' Última actualización: 10/03/2026 - FASE 1B Mejoras Finales (Segunda migración)
' ============================================================================

' ----------------------------------------------------------------------
' Constante: APP_PASSWORD
' Propósito: Contraseña para protección de hojas normales y estructura del libro.
'            NO se usa para autenticación de usuarios (ver ADMIN_PASSWORD).
' Actualizada: 10/03/2026 - Segunda migración a contraseñas de proyecto
' ----------------------------------------------------------------------
Public Const APP_PASSWORD As String = "1234"

' ----------------------------------------------------------------------
' Constante: ADMIN_PASSWORD
' Propósito: Contraseña para autenticación de usuarios administradores.
'            Separada de APP_PASSWORD para mayor seguridad.
' Actualizada: 10/03/2026 - Segunda migración a contraseñas de proyecto
' ----------------------------------------------------------------------
Public Const ADMIN_PASSWORD As String = "1234"

' ----------------------------------------------------------------------
' Constante: AUDIT_PASSWORD
' Propósito: Contraseña exclusiva para protección de hojas Audit Trail.
'            Garantiza inmutabilidad de registros históricos de auditoría.
' Actualizada: 10/03/2026 - Segunda migración a contraseñas de proyecto
' ----------------------------------------------------------------------
Public Const AUDIT_PASSWORD As String = "1234"

' ----------------------------------------------------------------------
' Constantes: Sistema de rotación de hojas Audit Trail
' Propósito: Controlan el comportamiento del sistema de múltiples hojas
'            de auditoría (AuditRotation2 + AuditLogger2).
'
' AUDIT_MAX_ROWS:
'   Máximo de filas de datos por hoja antes de rotar automáticamente a la
'   siguiente hoja disponible.
'   ? PRODUCCIÓN : 1000000
'   ? DEBUG/TEST : 100     ? cambiar SOLO en pruebas; revertir antes de desplegar
'
' AUDIT_MAX_SHEETS:
'   Cantidad total de hojas Audit Trail habilitadas en el libro (1–5).
'   Las hojas 2–5 deben existir en el libro (xlSheetVeryHidden hasta que se usen).
'   Ver instrucciones de creación en AuditRotation2.bas.
'
' AUDIT_BASE_NAME:
'   Nombre base de las hojas de auditoría.
'   Hoja 1: "Audit trail"   Hoja 2: "Audit trail 2"  ...  Hoja 5: "Audit trail 5"
'
' AUDIT_TABLE_PREFIX:
'   Prefijo de las tablas ListObject en cada hoja.
'   Hoja 1: "tblAudit"      Hoja 2: "tblAudit2"      ...  Hoja 5: "tblAudit5"
' ----------------------------------------------------------------------
Public Const AUDIT_MAX_ROWS     As Long = 100        ' PRODUCCIÓN: 1000000 | DEBUG: 100 — TEMPORALMENTE EN 100 PARA TEST
Public Const AUDIT_MAX_SHEETS   As Long = 5
Public Const AUDIT_BASE_NAME    As String = "Audit trail"
Public Const AUDIT_TABLE_PREFIX As String = "tblAudit"

' ============================================================================
' CONFIGURACIÓN DE INTERFAZ Y NAVEGACIÓN
' Última actualización: 12/03/2026 - Refactorización para portabilidad
' ============================================================================

' ----------------------------------------------------------------------
' Constante: MAIN_MENU_SHEET
' Propósito: Nombre de la hoja principal del menú.
'            Centraliza el nombre para facilitar portabilidad entre proyectos.
' Uso: ThisWorkbook, SheetService2, NavigationService2, UserManager2
' ----------------------------------------------------------------------
Public Const MAIN_MENU_SHEET As String = "Menú principal"

' ----------------------------------------------------------------------
' Constante: USER_DISPLAY_CELL
' Propósito: Celda donde se muestra el nombre del usuario activo en el menú.
'            Permite cambiar la ubicación sin modificar código de UserManager2.
' Formato: Rango de Excel (ej: "B9", "C5", "A1")
' ----------------------------------------------------------------------
Public Const USER_DISPLAY_CELL As String = "I9"

' ----------------------------------------------------------------------
' Constante: INITIAL_USER_ROLE
' Propósito: Rol por defecto asignado al abrir el libro.
'            Los usuarios pueden elevarse a "Admin" mediante autenticación.
' Valores típicos: "Usuario", "Invitado", "Operador"
' ----------------------------------------------------------------------
Public Const INITIAL_USER_ROLE As String = "Usuario"

' ============================================================================
' CONFIGURACIÓN DE DESARROLLO Y PROTECCIÓN
' Última actualización: 14/04/2026 - Booleanos de control para testing
' ============================================================================

' ----------------------------------------------------------------------
' Constante: ENABLE_WORKBOOK_PROTECTION
' Propósito: Activa/desactiva la protección de estructura del libro (URS-22)
'            en tiempo de ejecución sin modificar código.
' Valores:
'   ? True  = PRODUCCIÓN: Estructura protegida, no se pueden eliminar hojas
'   ? False = DESARROLLO: Puedes cambiar hojas libremente para testing
' Última actualización: 14/04/2026 - Agregado para facilitar desarrollo
' Nota: Cambiar SOLO este valor para debugging
' ----------------------------------------------------------------------
Public Const ENABLE_WORKBOOK_PROTECTION As Boolean = False  ' DESARROLLO: False | PRODUCCIÓN: True

' ----------------------------------------------------------------------
' Constante: ENABLE_SHEET_PROTECTION
' Propósito: Activa/desactiva la protección individual de hojas (URS-20/21)
'            Cuando False, usuarios pueden editar cualquier celda.
' Valores:
'   ? True  = PRODUCCIÓN: Hojas protegidas según reglas
'   ? False = DESARROLLO: Edición libre para testing
' Nota: Relacionado con SheetProtector2.ProtectSheet()
' ----------------------------------------------------------------------
Public Const ENABLE_SHEET_PROTECTION As Boolean = False  ' DESARROLLO: False | PRODUCCIÓN: True

' ============================================================================
' CONFIGURACIÓN DE SISTEMA DE INSPECCIONES
' Última actualización: 12/03/2026 - Creación del sistema de inspecciones
' ============================================================================

' ----------------------------------------------------------------------
' Constantes de nombres de hojas del sistema de inspecciones
' ----------------------------------------------------------------------
Public Const SHEET_PERSONAL As String = "Personal"
Public Const SHEET_CHECKLIST As String = "Checklist"
Public Const SHEET_HISTORICO As String = "Historico"
Public Const SHEET_CONFIGURACION As String = "Configuración"
Public Const SHEET_CRONOGRAMA As String = "Cronograma"
Public Const SHEET_ASEGURAMIENTO As String = "Aseguramiento de calidad"

' ----------------------------------------------------------------------
' Constantes de nombres de tablas del sistema de inspecciones
' ----------------------------------------------------------------------
Public Const TABLE_PERSONAL As String = "tblPersonal"
Public Const TABLE_PLANTILLAS As String = "tblPlantillas"
Public Const TABLE_PREGUNTAS As String = "tblPreguntas"
Public Const TABLE_INSPECCIONES As String = "tblInspecciones"
Public Const TABLE_RESPUESTAS As String = "tblRespuestas"
Public Const TABLE_CRONOGRAMA As String = "tblCronogramaInspecciones"
Public Const TABLE_SECCIONES As String = "tblSecciones"
Public Const TABLE_OPCIONES As String = "tblOpcionesDeRespuesta"
Public Const TABLE_CRITICIDAD As String = "tblCriticidad"
Public Const TABLE_CATEGORIAS_RPN As String = "tblCategoriasRPN"
Public Const TABLE_CONFIGURACION As String = "tblConfiguracion"
Public Const TABLE_EQUIPOS As String = "tblEquipos"
Public Const TABLE_PUESTO As String = "tblPuesto"
Public Const TABLE_PUESTOS_INICIALES As String = "tblPuestosIniciales"
Public Const TABLE_ASEGURAMIENTO As String = "tblAseguramientoCalidad"
Public Const TABLE_RESUMEN_CRONOGRAMA As String = "tblResumenCronograma"
Public Const TABLE_PLANTA As String = "tblPlanta"

' ----------------------------------------------------------------------
' Constantes de ubicación del cronograma resumen en Menú principal
' ----------------------------------------------------------------------
Public Const RESUMEN_CRONOGRAMA_CELDA As String = "B14"
Public Const RESUMEN_FILTRO_PLANTA_CELDA As String = "J15"

' ----------------------------------------------------------------------
' Constantes de lugar de auditoría
' ----------------------------------------------------------------------
Public Const LUGAR_DENTRO_AREA As String = "Dentro del área"
Public Const LUGAR_FUERA_AREA As String = "Fuera del área"

' ----------------------------------------------------------------------
' Constantes de parámetros del sistema de inspecciones
' ----------------------------------------------------------------------
Public Const PARAM_DIAS_ALERTA_VENCIMIENTO As String = "ALERTAR_DIAS_ANTES_VENCIMIENTO"
Public Const PARAM_FRECUENCIA_DEFAULT As String = "FRECUENCIA_DEFAULT_MESES"
Public Const PARAM_VALIDAR_PERSONA_ACTIVA As String = "VALIDAR_PERSONA_ACTIVA"

' ----------------------------------------------------------------------
' Constantes de estados de cronograma
' ----------------------------------------------------------------------
Public Const ESTADO_NUNCA_INSPECCIONADO As String = "Nunca inspeccionado"
Public Const ESTADO_VIGENTE As String = "Vigente"
Public Const ESTADO_POR_VENCER As String = "Por vencer"
Public Const ESTADO_VENCIDO As String = "Vencido"
Public Const ESTADO_PUESTO_INACTIVO As String = "Puesto inactivo"

' ----------------------------------------------------------------------
' Constantes de estados de inspección
' ----------------------------------------------------------------------
Public Const INSPECCION_EN_PROGRESO As String = "En progreso"
Public Const INSPECCION_COMPLETADO As String = "Completado"
Public Const INSPECCION_CANCELADO As String = "Cancelado"

' ============================================================================
' CONFIGURACIÓN DE CERTIFICADOS PDF
' Última actualización: 15/04/2026 - Fase 0: Preparación infraestructura
' ============================================================================

' ----------------------------------------------------------------------
' Constante: SHEET_PLANTILLA_CERTIFICADO
' Propósito: Nombre de la hoja oculta que sirve como plantilla para
'            generar certificados PDF de inspecciones completadas.
' Estado: Hoja muy oculta (xlSheetVeryHidden)
' ----------------------------------------------------------------------
Public Const SHEET_PLANTILLA_CERTIFICADO As String = "Plantilla Certificado"

' ----------------------------------------------------------------------
' Constantes: Configuración de exportación PDF
' ----------------------------------------------------------------------
Public Const PDF_PREFIJO_NOMBRE As String = "CERTIFICADO"
Public Const PDF_CALIDAD As Long = 0  ' 0 = xlQualityStandard, 1 = xlQualityMinimum
Public Const PDF_ABRIR_AUTOMATICO As Boolean = True

' ----------------------------------------------------------------------
' Constantes: Anchos de columnas para tabla de preguntas en certificado
' Valores en puntos de Excel (1 punto ˜ 0.35 mm)
' ----------------------------------------------------------------------
Public Const CERT_ANCHO_COL_NUMERO As Double = 20
Public Const CERT_ANCHO_COL_PREGUNTA As Double = 350
Public Const CERT_ANCHO_COL_RESPUESTA As Double = 80
Public Const CERT_ANCHO_COL_OBSERVACION As Double = 90




' ----------------------------------------------------------------------
' Array de nombres de columnas de puestos en tblPersonal
' Nota: Este array debe mantenerse sincronizado con la estructura de tblPersonal
' ----------------------------------------------------------------------
Public Function GetPuestosColumns() As Variant
    GetPuestosColumns = Array( _
        "Quimico", _
        "Digitador", _
        "Etiquetado", _
        "Ayudante 2", _
        "Ayudante 1", _
        "Ayudante 1 Electrolitos", _
        "Operador", _
        "Operador Electrolitos", _
        "Técnico de producción - grado C", _
        "Técnico de producción - grado D", _
        "Muestreador" _
    )
End Function

' ----------------------------------------------------------------------
' Orden de criticidad de puestos para cronograma resumen
' Índice menor = más crítico. Puestos excluidos: Quimico, Digitador, Etiquetado
' Nota: Mantener sincronizado con GetPuestosColumns
' ----------------------------------------------------------------------
Public Function GetOrdenCriticidadPuestos() As Variant
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

' ============================================================================
' DOCUMENTACIÓN: ESTRUCTURA DE TABLAS VERIFICADAS
' Última verificación: 14/04/2026
' Propósito: Referencia de columnas reales en archivo Excel para mantener
'            consistencia entre código y estructura de datos.
' IMPORTANTE: Si modificas nombres de columnas en Excel, actualiza esta
'             documentación Y todos los módulos que las usen.
' ============================================================================

' ----------------------------------------------------------------------
' TABLA: tblCronogramaInspecciones
' Ubicación: Hoja "Cronograma"
' Propósito: Registro maestro de cronogramas de inspección por persona/puesto
' Total columnas: 24
' ----------------------------------------------------------------------
' COLUMNAS VERIFICADAS (14/04/2026):
'   [01] ID Cronograma                  - String (UUID único)
'   [02] Iniciales personal             - String (FK a tblPersonal)
'   [03] ID Plantilla                   - String (FK a tblPlantillas)
'   [04] Puesto                          - String
'   [05] Planta personal                 - String
'   [06] Total inspecciones              - Long
'   [07] Fecha primera inspeccion        - Date
'   [08] Fecha ultima inspeccion         - Date
'   [09] ID Ultima inspeccion            - String (FK a tblInspecciones)
'   [10] RPN ultima inspeccion           - Double
'   [11] Categoria ultima inspeccion     - String
'   [12] Fecha proxima inspeccion        - Date (calculada)
'   [13] Dias para vencimiento           - Long (calculado)
'   [14] Estado cronograma               - String (ver constantes ESTADO_*)
'   [15] Dias alerta                     - Long
'   [16] Puesto activo en personal       - String ("Si"/"No")
'   [17] Personal activo                 - String ("Si"/"No")
'   [18] Plantilla tiene preguntas       - String ("Si"/"No")
'   [19] Fecha ultima actualizacion      - Date
'   [20] Requiere recalculo              - String ("Si"/"No")
'   [21] Nombre plantilla                - String
'   [22] Frecuencia meses                - Long (1, 3, 6, 12)
'   [23] Activo                          - String ("Si"/"No")
'   [24] Fecha de creacion               - Date
'
' MÓDULOS QUE USAN ESTA TABLA:
'   - InspectionScheduler.bas (InicializarCronograma, RecalcularCronograma)
'   - CronogramaResumen.bas (RefrescarResumenCronograma)
'   - SystemInitializer.bas (InicializarCronogramaSilencioso)
'   - TestDataGenerator.bas (GenerarDatosPrueba)
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' TABLA: tblResumenCronograma
' Ubicación: Hoja "Menú principal"
' Propósito: Vista resumida del cronograma filtrada por planta y ordenada
'            por criticidad de puesto. Se usa para navegación con doble clic.
' Total columnas: 6
' ----------------------------------------------------------------------
' COLUMNAS VERIFICADAS (22/04/2026):
'   [1] Iniciales           - String
'   [2] Puesto              - String
'   [3] Mes proxima inspeccion - String (formato "Mayo", "Julio" - solo mes)
'   [4] Dias vencimiento    - Long (calculado hasta último día del mes)
'   [5] ID Cronograma       - String (para lookup)
'   [6] ID Plantilla        - String (para lookup)
'
' NOTA: El sistema considera TODO EL MES para realizar la inspección.
'       Si la inspección vence en "Mayo", tiene hasta el 31 de Mayo.
'
' MÓDULOS QUE USAN ESTA TABLA:
'   - CronogramaResumen.bas (RefrescarResumenCronograma - ESCRITURA)
'   - Hoja1.bas (Worksheet_BeforeDoubleClick - LECTURA)
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' TABLA: tblPersonal
' Ubicación: Hoja "Personal"
' Propósito: Registro maestro de personal. Cada fila es una persona que
'            puede tener múltiples puestos (columnas de puestos).
' Total columnas: 14
' ----------------------------------------------------------------------
' COLUMNAS VERIFICADAS (14/04/2026):
'   [01] Iniciales                           - String (PK única, FK en tblCronogramaInspecciones)
'   [02] Planta                              - String (lugar de trabajo)
'   [03] Quimico                             - String ("Si"/"No" - indica si tiene puesto)
'   [04] Digitador                           - String ("Si"/"No" - indica si tiene puesto)
'   [05] Etiquetado                          - String ("Si"/"No" - indica si tiene puesto)
'   [06] Ayudante 2                          - String ("Si"/"No" - indica si tiene puesto)
'   [07] Ayudante 1                          - String ("Si"/"No" - indica si tiene puesto)
'   [08] Ayudante 1 Electrolitos             - String ("Si"/"No" - indica si tiene puesto)
'   [09] Operador                            - String ("Si"/"No" - indica si tiene puesto)
'   [10] Operador Electrolitos               - String ("Si"/"No" - indica si tiene puesto)
'   [11] Técnico de producción - grado C     - String ("Si"/"No" - indica si tiene puesto)
'   [12] Técnico de producción - grado D     - String ("Si"/"No" - indica si tiene puesto)
'   [13] Muestreador                         - String ("Si"/"No" - indica si tiene puesto)
'   [14] Activo                              - String ("Si"/"No" - determina si se incluye en cronograma)
'
' VALORES de las columnas [03-13] (Puestos):
'   - "Si" = La persona tiene este puesto asignado
'   - "No" = La persona NO tiene este puesto asignado
'   - Ver GetPuestosColumns() para lista programática de puestos
'
' MÓDULOS QUE USAN ESTA TABLA:
'   - ChecklistRepository.bas (ObtenerPersonalPorPuesto - LECTURA)
'   - InspectionScheduler.bas (InicializarCronograma, RecalcularCronograma, ObtenerValorPersonal - LECTURA)
'   - SystemInitializer.bas (ValidarDatosMaestros, InicializarCronogramaSilencioso - LECTURA)
'   - Configuration2.bas (GetPuestosColumns - DEFINICIÓN de columnas)
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' TABLA: tblPlantillas
' Ubicación: Hoja "Checklist"
' Propósito: Plantillas de inspección. Cada plantilla define qué preguntas
'            se hacen para un puesto específico y con qué frecuencia.
' Total columnas: 5
' Columnas usadas por código: 5 (todas)
' ----------------------------------------------------------------------
' COLUMNAS VERIFICADAS (22/04/2026):
'   [1] ID Plantilla              - String (PK única, FK en tblCronogramaInspecciones)
'   [2] Nombre de plantilla       - String (descripción de la plantilla)
'   [3] Área                      - String (Área de trabajo, precarga cboArea en frmChecklistVirtual)
'                                   VALORES POSIBLES:
'                                   - "" (vacío), "TODAS", "GENERAL" → Plantilla genérica
'                                     * Therapia iv Santiago: Usuario elige entre 3 áreas (Planta Azul NPT, Planta Blanca NPT, Oncología)
'                                     * Therapia iv Concepción: Usuario elige entre 2 áreas (NPT, Oncología)
'                                   - "NPT" → Plantilla NPT específica
'                                     * Therapia iv Santiago: Usuario elige entre 2 opciones NPT (Planta Azul, Planta Blanca)
'                                     * Therapia iv Concepción: Bloqueado en "NPT"
'                                   - "ONCO" → Plantilla Oncología específica
'                                     * Ambas plantas: Bloqueado en "Oncología"
'   [4] Puesto                    - String (FK a GetPuestosColumns)
'   [5] Frecuencia meses          - Long (1, 3, 6, 12 - intervalo de inspección)
'
' MÓDULOS QUE USAN ESTA TABLA:
'   - ChecklistRepository.bas (ObtenerPlantillaPorPuesto - LECTURA)
'   - InspectionScheduler.bas (InicializarCronograma, RecalcularCronograma, ObtenerPlantillaPorIDPlantilla - LECTURA)
'   - SystemInitializer.bas (ValidarDatosMaestros, InicializarCronogramaSilencioso - LECTURA)
'   - TableManager.bas (AddRowToTable - ESCRITURA columna Área, RemoveTableRows - INCLUIDA en tabla_list)
'   - frmSelectorInspeccion.frm (CargarPlantillasDisponibles - LECTURA columna Área)
'   - ChecklistOrchestrator.bas (AbrirChecklistVirtual - PASA área a formulario)
'   - frmChecklistVirtual.frm (AplicarFiltroArea - LÓGICA CONDICIONAL área)
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' TABLA: tblPreguntas
' Ubicación: Hoja "Checklist"
' Propósito: Preguntas de inspección. Define las preguntas específicas
'            para cada plantilla. Las preguntas agrupadas por sección
'            y filtradas por criticidad.
' Total columnas: 10
' Columnas usadas por código: 8
' Columnas no usadas: Observaciones, Fecha de creación (datos adicionales)
' ----------------------------------------------------------------------
' COLUMNAS VERIFICADAS (14/04/2026):
'   [1] ID Plantilla                 - String (FK a tblPlantillas)
'   [2] ID Pregunta                  - String (PK única)
'   [3] ID Seccion                   - String (FK a tblSecciones)
'   [4] Numero                       - Long (número secuencial sin acento)
'   [5] Texto                        - String (texto de la pregunta SIN "pregunta")
'   [6] ID Criticidad                - String (FK a criticidad)
'   [7] Orden                        - Long (orden de aparición, era "Número" antes)
'   [8] Activo                       - String ("Si"/"Sí" - filtrado en búsquedas)
'   [9] Observaciones                - String (NO USADA - datos adicionales)
'  [10] Fecha de creación            - Date (NO USADA - datos adicionales)
'
' NOTAS IMPORTANTES:
'   - SIN ACENTO: "ID Seccion" no "ID Sección"
'   - SIN ACENTO: "Numero" no "Número"
'   - SOLO "Texto": "Texto" no "Texto pregunta"
'   - Columna [8] se valida con UCase(Activo) <> "SI" And UCase(Activo) <> "SÍ"
'
' MÓDULOS QUE USAN ESTA TABLA:
'   - ChecklistRepository.bas (ObtenerPreguntasPlantilla - LECTURA)
'   - TableManager.bas (TableManager.RemoveTableRows - INCLUIDA en tabla_list)
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' TABLA: tblSecciones
' Ubicación: Hoja "Checklist"
' Propósito: Agrupación lógica de preguntas dentro de una plantilla.
'            Define cómo se presentan las preguntas (sección y tipo respuesta).
' Total columnas: 3
' Columnas usadas por código: 3
' Columnas no usadas: ninguna
' ----------------------------------------------------------------------
' COLUMNAS VERIFICADAS (14/04/2026):
'   [1] ID Seccion                - String (PK única, FK en tblPreguntas)
'   [2] Nombre de sección         - String (nombre descriptivo de la sección)
'   [3] Tipo de respuesta         - String (tipo de control UI para respuestas)
'
' MÓDULOS QUE USAN ESTA TABLA:
'   - ChecklistRepository.bas (ObtenerSecciones - LECTURA)
'   - TableManager.bas (TableManager.RemoveTableRows - INCLUIDA en tabla_list)
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' TABLA: tblOpcionesDeRespuesta
' Ubicación: Hoja "Checklist"
' Propósito: Opciones posibles de respuesta para preguntas. Cada opción
'            está asociada a una sección (tipo de respuesta) y tiene un
'            valor de puntaje que se usa para calcular RPN.
' Total columnas: 5
' Columnas usadas por código: 5
' Columnas no usadas: ninguna
' ----------------------------------------------------------------------
' COLUMNAS VERIFICADAS (15/04/2026):
'   [1] ID Opcion                  - String (PK única)
'   [2] ID Seccion                 - String (FK a tblSecciones)
'   [3] ID Criticidad              - String (FK a tblCriticidad - filtrado de opciones)
'   [4] Opción texto               - String (texto de la opción)
'   [5] Valor puntaje              - Long (puntaje para cálculo RPN)
'
' MÓDULOS QUE USAN ESTA TABLA:
'   - ChecklistRepository.bas (ObtenerOpcionesRespuesta, ObtenerIDOpcionPorTexto - LECTURA)
'   - frmChecklistVirtual.frm (Filtra opciones por ID Seccion e ID Criticidad)
'   - TableManager.bas (TableManager.RemoveTableRows - INCLUIDA en tabla_list)
' NOTA: ID Criticidad se usa para filtrar opciones por pregunta según su criticidad.
'       Cada pregunta tiene ID Sección e ID Criticidad, y solo se muestran las
'       opciones que coincidan con ambos valores.
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' TABLA: tblEquipos
' Ubicación: Hoja "Configuración"
' Propósito: Catálogo maestro de equipos/áreas en cada planta.
'            Se usa para filtrar y agrupar inspecciones.
' Total columnas: 3
' Columnas usadas por código: 3
' Columnas no usadas: ninguna
' ----------------------------------------------------------------------
' COLUMNAS VERIFICADAS (14/04/2026):
'   [1] Planta                    - String (nombre de la planta)
'   [2] Área                      - String (nombre del área CON ACENTO)
'   [3] Equipo                    - String (nombre/descripción del equipo)
'
' MÓDULOS QUE USAN ESTA TABLA:
'   - ChecklistRepository.bas (ObtenerEquiposPorPlanta, ObtenerAreasPorPlanta - LECTURA)
'   - SystemInitializer.bas (ValidarDatosMaestros - LECTURA)
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' TABLA: tblAseguramientoCalidad
' Ubicación: Hoja "Aseguramiento"
' Propósito: Registro de evaluadores autorizados para realizar auditorías
'            de aseguramiento de calidad.
' Total columnas: 4
' Columnas usadas por código: 2 (columnas 2 y 3)
' Columnas parcialmente usadas: ValidarDatosMaestros solo verifica existencia
' ----------------------------------------------------------------------
' COLUMNAS VERIFICADAS (22/04/2026):
'   [1] Planta                    - String (planta del evaluador)
'   [2] Nombre                    - String (nombre completo del evaluador, para match con Application.UserName)
'   [3] Iniciales                 - String (COLUMNA 3 - iniciales para cbo/lista)
'   [4] Activo                    - String ("Si"/"No" - activo/inactivo)
'
' NOTA TÉCNICA:
'   - El código accede a Iniciales por posición (columna 3) no por nombre
'   - Línea: iniciales = Trim(asegRow.Range.Cells(1, 3).Value)
'   - ValidarDatosMaestros solo cuenta filas sin usar columnas específicas
'   - ObtenerInicialesEvaluadorPorNombre usa columna 2 (Nombre) para match con Application.UserName (NO Environ("USERNAME"))
'   - IMPORTANTE: Application.UserName retorna nombre completo (ej: "NIEVES CARRERO")
'                 Environ("USERNAME") retorna cuenta de Windows (ej: "carre")
'
' MÓDULOS QUE USAN ESTA TABLA:
'   - ChecklistRepository.bas (ObtenerEvaluadores - LECTURA columna 3)
'   - ChecklistRepository.bas (ObtenerInicialesEvaluadorPorNombre - LECTURA columnas 2 y 3)
'   - SystemInitializer.bas (ValidarDatosMaestros - LECTURA para validación)
'   - frmChecklistVirtual.frm (CargarComboEvaluadores - USA ObtenerInicialesEvaluadorPorNombre para pre-selección)
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' TABLA: tblInspecciones
' Ubicación: Hoja "Historico"
' Propósito: Registro de inspecciones completadas con resultados de scoring,
'            RPN y categorización. Cada inspección está vinculada a un
'            personal, plantilla y tiene múltiples respuestas en tblRespuestas.
' Total columnas: 40
' Última actualización: 21/04/2026 - Migración inspecciones recurrentes (cols 32-40)
' ----------------------------------------------------------------------
' COLUMNAS VERIFICADAS (21/04/2026):
'   [01] ID Inspeccion               - String (UUID único, PK)
'   [02] Area                        - String (área donde se realizó la inspección)
'   [03] Linea Auditada              - String (línea de producción/proceso auditada)
'   [04] Hora inicio                 - String (hora de inicio de inspección)
'   [05] Hora termino                - String (hora de término de inspección)
'   [06] Iniciales AY1               - String (iniciales ayudante 1)
'   [07] Iniciales AY2               - String (iniciales ayudante 2)
'   [08] Iniciales OP                - String (iniciales operador)
'   [09] Lugar Auditoria             - String ("Dentro del área"/"Fuera del área")
'   [10] Iniciales personal          - String (FK a tblPersonal - evaluado)
'   [11] ID Plantilla                - String (FK a tblPlantillas)
'   [12] Planta                      - String (planta ejecutora)
'   [13] Fecha inspeccion            - Date (fecha de realización)
'   [14] Auditor                     - String (quien realizó la auditoría)
'   [15] Estado                      - String (ver INSPECCION_* constantes)
'   [16] TA puntaje obtenido         - Double (puntaje obtenido en sección TA)
'   [17] TA puntos maximos           - Double (máximo posible sección TA)
'   [18] TA puntos no aplica         - Double (puntos no aplicables en TA)
'   [19] TA porcentaje               - Double (porcentaje final TA)
'   [20] Auditoria Procesos Resultado - String (resultado auditoría de procesos - col 20)
'   [21] RPN calculado               - Double (Risk Priority Number calculado)
'   [22] Categoria resultado         - String ("Categoría 1/2/3")
'   [23] Requiere accion             - String ("Si"/"No" - si requiere acción)
'   [24] Fecha proxima inspeccion    - Date (fecha calculada próxima inspección)
'   [25] Dias para vencimiento       - Long (días hasta vencimiento)
'   [26] Estado programacion         - String (ver ESTADO_* constantes)
'   [27] Observaciones generales     - String (observaciones de la inspección)
'   [28] Fecha calculo               - Date (timestamp de cálculos)
'   [29] Usuario calculo             - String (usuario que completó cálculos)
'   [30] Fecha completado            - Date (timestamp de completado)
'   [31] Usuario completado          - String (usuario que completó)
'   
'   NUEVAS COLUMNAS - INSPECCIONES RECURRENTES (21/04/2026):
'   [32] Numero Inspeccion           - Long (secuencia: 1, 2, 3... Default=1)
'   [33] Es Inspeccion Recurrente    - String ("Si"/"No" - Default="No")
'   [34] Puesto Evaluado             - String (puesto específico en esta inspección)
'   [35] RPN Anterior Manual         - Double (RPN ingresado manualmente - Nullable)
'   [36] ID Inspeccion Anterior      - String (UUID inspección previa - Nullable)
'   [37] RPN Promedio                - Double ((RPN Ant + RPN Act)/2 - Nullable)
'   [38] Porcentaje Recuperacion     - Double (futuro: datos microbiología - Default=0)
'   [39] Porcentaje OOL              - Double (futuro: Out Of Limits micro - Default=0)
'   [40] RPN Total                   - Double (RPN Prom + %Rec + %OOL - Nullable)
'
' MÓDULOS QUE USAN ESTA TABLA:
'   - InspectionRepository.bas (CrearInspeccion, ActualizarCalculosInspeccion - ESCRITURA)
'   - InspectionScheduler.bas (ObtenerUltimaInspeccion - LECTURA)
'   - ChecklistOrchestrator.bas (GuardarInspeccionCompleta vía Repository)
'   - CertificadoPDFGenerator.bas (ObtenerDatosInspeccion - LECTURA col 20-22, 27)
'   - InspectionHistoryService.bas (BuscarInspeccionesPrevias - LECTURA cols 32-40)
'   - RecurrentInspectionCalculator.bas (CalcularRPNPromedio - LECTURA/ESCRITURA cols 35-40)
'   - frmChecklistVirtual.frm (Captura datos recurrentes - ESCRITURA cols 32-36)
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' TABLA: tblRespuestas
' Ubicación: Hoja "Historico"
' Propósito: Respuestas individuales de cada pregunta en una inspección.
'            Cada respuesta vincula una inspección con una pregunta y su
'            opción seleccionada, incluyendo observaciones opcionales.
' Total columnas: 8
' ----------------------------------------------------------------------
' COLUMNAS VERIFICADAS (20/04/2026):
'   [1] ID Respuesta            - String (UUID único, PK)
'   [2] ID Inspeccion           - String (FK a tblInspecciones)
'   [3] ID Pregunta             - String (FK a tblPreguntas)
'   [4] ID Opcion               - String (FK a tblOpcionesDeRespuesta)
'   [5] Valor numerico          - Double (valor numérico de la opción para scoring)
'   [6] Observacion             - String (observaciones opcionales de la respuesta)
'   [7] Fecha respuesta         - Date (timestamp de registro de respuesta)
'   [8] ID Criticidad           - String (FK a tblCriticidad - copiado de pregunta para cálculos)
'
' MÓDULOS QUE USAN ESTA TABLA:
'   - InspectionRepository.bas (GuardarRespuestas - ESCRITURA)
'   - frmChecklistVirtual.frm (ObtenerRespuestas, RecopilarRespuestas)
'   - InspectionCalculator.bas (CalcularScoringTA usa IDCriticidad para ajuste "No Aplica")
' ----------------------------------------------------------------------