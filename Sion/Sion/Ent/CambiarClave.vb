Public Class CambiarClave
    Public Property contrasenaActual As String
    Public Property contrasenaNueva As String
    Public Property contrasenaConfirmar As String

    Public Property nombreUsuario As Integer

    Public Function verificarContrasena() As Boolean
        Dim params As New List(Of String)
        params.Add(SesionActual.idUsuario)
        params.Add(contrasenaActual)
        params.Add(SesionActual.idEmpresa)
        Return General.getEstadoVF(Consultas.VERIFICAR_CAMBIAR_CLAVE, params)

    End Function


    Public Sub guardarClave()
        CambiarClaveDAL.guardarClave(Me)
    End Sub
End Class
