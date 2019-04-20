Imports System.Data.SqlClient
Public Class DestinoDAL
    Public Shared Function guardarDestino(objConfiguracion As ConfiguracionGeneral)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = Consultas.DESTINO_CREAR
                    comando.Parameters.Add(New SqlParameter("@ID", SqlDbType.NVarChar)).Value = objConfiguracion.codigo
                    comando.Parameters.Add(New SqlParameter("@Descripcion", SqlDbType.NVarChar)).Value = objConfiguracion.descripcion
                    comando.Parameters.Add(New SqlParameter("@CodigoEstadoSalida", SqlDbType.NVarChar)).Value = objConfiguracion.codigoComplemento
                    comando.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objConfiguracion.idUsuario
                    objConfiguracion.codigo = CType(comando.ExecuteScalar, String)
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objConfiguracion
    End Function

End Class
