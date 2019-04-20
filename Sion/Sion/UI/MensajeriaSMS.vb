Imports System.Net
Public Class MensajeriaSMS
    Property celular As String
    Property mensaje As String
    Public Function enviarMensaje()
        Dim url As String = "http://sinm.smartgrouping.com/API/"
        Dim cliente As New WebClient()
        Dim terminal As String = ""
        Dim tipo As String = "SMS"
        cliente.QueryString.Add("t", terminal)
        cliente.QueryString.Add("d", celular)
        cliente.QueryString.Add("m", mensaje)
        cliente.QueryString.Add("b", tipo)
        Return cliente.DownloadString(url)
    End Function

End Class
