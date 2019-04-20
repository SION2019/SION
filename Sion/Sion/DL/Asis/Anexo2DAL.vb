Imports System.Data.SqlClient
Public Class Anexo2DAL
    Public Sub GuardarAnexo(ByVal N_solicitud As Object, ByVal registro As String, ByVal domicilio As String, ByVal observacion As String, ByVal internacion As String,
                                   ByVal remision As String, ByVal contraremision As String, ByVal otros As String,
                                   ByVal nota As String, ByVal codigo_institucion As String, ByVal codigo_pais As String, ByVal codigo_departamento As String, ByVal codigo_municipio As String,
                                   ByVal fecha_urgencia As String, ByVal usuario As String)

        Try
            If N_solicitud.text = "" Then
                Using comando As New SqlCommand("SP_ANEXO2_CREAR")
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Connection = FormPrincipal.cnxion
                    comando.Parameters.Add(New SqlParameter("@N_registro", SqlDbType.NVarChar))
                    comando.Parameters("@N_registro").Value = registro
                    comando.Parameters.Add(New SqlParameter("@domicilio", SqlDbType.Bit))
                    comando.Parameters("@domicilio").Value = domicilio
                    comando.Parameters.Add(New SqlParameter("@observacion", SqlDbType.Bit))
                    comando.Parameters("@observacion").Value = observacion
                    comando.Parameters.Add(New SqlParameter("@internacion", SqlDbType.Bit))
                    comando.Parameters("@internacion").Value = internacion
                    comando.Parameters.Add(New SqlParameter("@remision", SqlDbType.Bit))
                    comando.Parameters("@remision").Value = remision
                    comando.Parameters.Add(New SqlParameter("@contraremision", SqlDbType.Bit))
                    comando.Parameters("@contraremision").Value = contraremision
                    comando.Parameters.Add(New SqlParameter("@otros", SqlDbType.Bit))
                    comando.Parameters("@otros").Value = otros
                    comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar))
                    comando.Parameters("@nota").Value = nota
                    comando.Parameters.Add(New SqlParameter("@codigo_institucion", SqlDbType.NVarChar))
                    comando.Parameters("@codigo_institucion").Value = codigo_institucion
                    comando.Parameters.Add(New SqlParameter("@codigo_pais", SqlDbType.NVarChar))
                    comando.Parameters("@codigo_pais").Value = codigo_pais
                    comando.Parameters.Add(New SqlParameter("@codigo_departamento", SqlDbType.NVarChar))
                    comando.Parameters("@codigo_departamento").Value = codigo_departamento
                    comando.Parameters.Add(New SqlParameter("@codigo_municipio", SqlDbType.NVarChar))
                    comando.Parameters("@codigo_municipio").Value = codigo_municipio
                    comando.Parameters.Add(New SqlParameter("@fecha_urgencia", SqlDbType.DateTime))
                    comando.Parameters("@fecha_urgencia").Value = fecha_urgencia
                    comando.Parameters.Add(New SqlParameter("@usuario_creacion", SqlDbType.NVarChar))
                    comando.Parameters("@usuario_creacion").Value = usuario
                    N_solicitud.text = CType(comando.ExecuteScalar, String)

                End Using
            Else
                Using consulta As New SqlCommand("SP_ANEXO2_ACTUALIZAR")
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Parameters.Add(New SqlParameter("@codigo", SqlDbType.NVarChar))
                    consulta.Parameters("@codigo").Value = N_solicitud.text
                    consulta.Parameters.Add(New SqlParameter("@domicilio", SqlDbType.Bit))
                    consulta.Parameters("@domicilio").Value = domicilio
                    consulta.Parameters.Add(New SqlParameter("@observacion", SqlDbType.Bit))
                    consulta.Parameters("@observacion").Value = observacion
                    consulta.Parameters.Add(New SqlParameter("@internacion", SqlDbType.Bit))
                    consulta.Parameters("@internacion").Value = internacion
                    consulta.Parameters.Add(New SqlParameter("@remision", SqlDbType.Bit))
                    consulta.Parameters("@remision").Value = remision
                    consulta.Parameters.Add(New SqlParameter("@contraremision", SqlDbType.Bit))
                    consulta.Parameters("@contraremision").Value = contraremision
                    consulta.Parameters.Add(New SqlParameter("@otros", SqlDbType.Bit))
                    consulta.Parameters("@otros").Value = otros
                    consulta.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar))
                    consulta.Parameters("@nota").Value = nota
                    consulta.Parameters.Add(New SqlParameter("@codigo_institucion", SqlDbType.NVarChar))
                    consulta.Parameters("@codigo_institucion").Value = codigo_institucion
                    consulta.Parameters.Add(New SqlParameter("@codigo_pais", SqlDbType.NVarChar))
                    consulta.Parameters("@codigo_pais").Value = codigo_pais
                    consulta.Parameters.Add(New SqlParameter("@codigo_departamento", SqlDbType.NVarChar))
                    consulta.Parameters("@codigo_departamento").Value = codigo_departamento
                    consulta.Parameters.Add(New SqlParameter("@codigo_municipio", SqlDbType.NVarChar))
                    consulta.Parameters("@codigo_municipio").Value = codigo_municipio
                    consulta.Parameters.Add(New SqlParameter("@fecha_urgencia", SqlDbType.DateTime))
                    consulta.Parameters("@fecha_urgencia").Value = fecha_urgencia
                    consulta.Parameters.Add(New SqlParameter("@usuario_actualizacion", SqlDbType.NVarChar))
                    consulta.Parameters("@usuario_actualizacion").Value = usuario
                    consulta.ExecuteNonQuery()
                End Using
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
