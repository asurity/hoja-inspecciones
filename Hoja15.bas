' ----------------------------------------------------------------------
' Módulo: Hoja "Graficos"
' Descripción: Eventos para la hoja Graficos (análisis visual de inspecciones).
'              Contiene tablas dinámicas y gráficos que se actualizan
'              automáticamente cada vez que el usuario navega a esta hoja.
' Fecha: 17/06/2026 — Creado para mantener gráficas en tiempo real.
'
' INSTRUCCIONES DE INSTALACIÓN:
' 1. Abre el VBA Editor (Alt+F11)
' 2. Busca en el árbol de la izquierda: Microsoft Excel Objetos → Hoja "Graficos"
' 3. Haz doble clic en esa hoja para abrir su módulo
' 4. Copia TODO el código de este archivo
' 5. Pégalo en el módulo de la hoja "Graficos"
' 6. Guarda el archivo
' ----------------------------------------------------------------------
Option Explicit

'' ----------------------------------------------------------------------
' Evento: Worksheet_Activate
' Propósito: Se ejecuta cuando se activa la hoja Graficos.
'            Refresca automáticamente todas las tablas dinámicas y gráficos
'            para que el usuario siempre vea datos actualizados.
'            SIN restricción de rol — cualquier usuario dispara el refresh.
' ----------------------------------------------------------------------
Private Sub Worksheet_Activate()
    ' ## NAVEGACIÓN ## Guardia: evitar doble ejecución durante navegación (FASE 4, 08/06/2026)
    If g_NavigationInProgress Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    ' Aplicar protección centralizada según el rol del usuario
    ' Admin → desprotegido (edición libre)
    ' Usuario → solo lectura con copiado
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.APP_PASSWORD)
    
    ' Refrescar tablas dinámicas y gráficos automáticamente
    ' Sin restricción de rol — cualquier usuario ve datos actualizados
    Call UpdatePivotTables.RefrescarGraficosAuto
    
    Exit Sub
ErrorHandler:
    Call ErrorLogger2.Log("Graficos.Worksheet_Activate", VBA.Err.Description, VBA.Err.Number)
End Sub

'' ----------------------------------------------------------------------
' Evento: Worksheet_Deactivate
' Propósito: Se ejecuta al salir de la hoja Graficos.
'            Refuerza la protección según el rol del usuario.
' ----------------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    On Error Resume Next
    Call SheetProtector2.ApplyRoleBasedProtection(Me, Configuration2.APP_PASSWORD)
    On Error GoTo 0
End Sub
