Sub ActualizarTablasDinamicas()

    If m_userRole = "Admin" Then
    
        ActiveWorkbook.RefreshAll
        MsgBox "Gráficos actualizados correctamente", vbApplicationModal, "PROCESO EXITOSO"
        
    Else
    
        MsgBox "No tienes permiso para actualizar los análisis de detecciones", vbCritical, "ACCESO DENEGADO"
        
    End If

End Sub