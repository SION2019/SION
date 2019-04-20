Public Class Anexo1BLL
    Dim consulta As New Anexo1DAL


    Public Sub guardar_anexo(ByRef orden As Object,
                                  ByVal registro As String,
                                  ByVal servicio As String,
                                  ByVal solicitud As String,
                                  observacion As String,
                                  ByVal usuario As Integer,
                                  dt_anexo As DataTable)

        consulta.GuardarAnexo(orden, registro, servicio, solicitud, observacion, usuario, dt_anexo)

    End Sub

End Class
