Option Explicit

Private Sub UserForm_Initialize()
    ' Configurar el TextBox para ocultar la contraseña con asteriscos (***).
    Me.txtContrasena.PasswordChar = "*"
    
    ' Llama al adaptador para configurar el diseño y los elementos del formulario.
    'Call InputAdapter.ConfigureForm(Me)
End Sub

Private Sub cmdAceptar_Click()
    ' Simplemente oculta el formulario para que el código que lo llamó pueda continuar.
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    ' Simplemente oculta el formulario para que el código que lo llamó pueda continuar.
    ' Nota: En el módulo ThisWorkbook se manejará la acción de cancelación.
    Me.Hide
End Sub
