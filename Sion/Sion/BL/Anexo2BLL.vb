Public Class Anexo2BLL
    Dim consulta As New Anexo2DAL

    Public Sub guardaranexo(ByRef N_solicitud As Object, ByVal registro As String, ByVal domicilio As String, ByVal observacion As String, ByVal internacion As String,
                                   ByVal remision As String, ByVal contraremision As String, ByVal otros As String,
                                   ByVal nota As String, ByVal codigo_institucion As String, ByVal codigo_pais As String, ByVal codigo_departamento As String, ByVal codigo_municipio As String,
                                   ByVal fecha_urgencia As DateTime, ByVal usuario As String)

        consulta.GuardarAnexo(N_solicitud, registro, domicilio, observacion, internacion, remision, contraremision, otros, nota, codigo_institucion, codigo_pais, codigo_departamento, codigo_municipio, fecha_urgencia, usuario)


    End Sub
End Class
