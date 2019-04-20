Imports System.Data.SqlClient
Public Class HemodialisisDAL
    Public Shared Sub guardarHemodialisis(ByRef objHemodialisis As Hemodialisis)
        Try
            Using comando = New SqlCommand()

                comando.Connection = FormPrincipal.cnxion

                comando.CommandType = CommandType.StoredProcedure
                comando.Parameters.Clear()
                comando.CommandText = objHemodialisis.sqlGuardarRegistro
                comando.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = objHemodialisis.CodigoProcedimiento
                comando.Parameters.Add(New SqlParameter("@IDHEMO", SqlDbType.Int)).Value = objHemodialisis.codigo
                comando.Parameters.Add(New SqlParameter("@Codigo_orden", SqlDbType.Int)).Value = objHemodialisis.CodigoOrden
                comando.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objHemodialisis.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIO_REAL", SqlDbType.Int)).Value = objHemodialisis.usuarioReal
                comando.Parameters.Add(New SqlParameter("@NOTA", SqlDbType.NVarChar)).Value = objHemodialisis.nota
                comando.Parameters.Add(New SqlParameter("@FECHAEX", SqlDbType.DateTime)).Value = objHemodialisis.fechaRegistro
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objHemodialisis.idEmpresa
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Int)).Value = objHemodialisis.editado
                comando.Parameters.Add(New SqlParameter("@RESULTADO", SqlDbType.NVarChar)).Value = objHemodialisis.resultado
                comando.Parameters.Add(New SqlParameter("@tablaMedicamentos", SqlDbType.Structured)).Value = objHemodialisis.dtMedicamento
                comando.Parameters.Add(New SqlParameter("@TABLA_SIGNO", SqlDbType.Structured)).Value = objHemodialisis.dtSigno
                objHemodialisis.codigo = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class