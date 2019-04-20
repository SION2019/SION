Public Class ConfiguracionEnvioCorreo
    Public Property codigoConfiguracion As Integer
    Public Property smtp As String
    Public Property puerto As Integer
    Sub New(ByVal pCodigoConfiguracion As Integer, ByVal pSmtp As String, ByVal pPuerto As Integer)
        codigoConfiguracion = pCodigoConfiguracion
        smtp = pSmtp
        pPuerto = puerto
    End Sub
    Sub New()

    End Sub
End Class
