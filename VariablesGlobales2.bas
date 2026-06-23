'' ----------------------------------------------------------------------
' Módulo: VariablesGlobales
' Descripción: Define variables globales accesibles desde cualquier módulo del proyecto.
'              Facilita la gestión centralizada de información compartida, como el rol del usuario.
'' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Variable global: m_userRole
' Propósito: Almacena el rol actual del usuario (por ejemplo, "Usuario", "Admin").
'            Permite controlar permisos y accesos en todo el libro.
'' ----------------------------------------------------------------------
Public m_userRole As String

'' ----------------------------------------------------------------------
' Variable global: g_AnalisisPendiente
' Propósito: Bandera que indica que al menos un registro fue guardado desde la
'            última ejecución del análisis. Permite diferir EjecutarAnalisis al
'            cierre del libro en lugar de ejecutarlo tras cada guardado individual
'            (FASE 4, 22/02/2026).
'' ----------------------------------------------------------------------
Public g_AnalisisPendiente As Boolean

'' ----------------------------------------------------------------------
' Variable global: g_PreviousSheetName
' Propósito: Almacena el nombre de la hoja anterior para rastreo de navegación.
'            Permite auditar cambios entre hojas sin generar registros duplicados
'            (FASE 2.6, 10/03/2026).
'' ----------------------------------------------------------------------
Public g_PreviousSheetName As String

'' ----------------------------------------------------------------------
' Variable global: g_NavigationInProgress
' Propósito: Bandera que evita recursión entre Workbook_SheetActivate y
'            NavigationService2 durante la navegación programática.
'            Se pone True al iniciar navegación, False al terminar.
'            (FASE 0 - Refactorización Navegación, 08/06/2026)
'' ----------------------------------------------------------------------
Public g_NavigationInProgress As Boolean