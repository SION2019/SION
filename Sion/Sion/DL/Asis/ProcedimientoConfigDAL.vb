Imports System.Data.SqlClient
Public Class ProcedimientoConfigDAL
    Public Function Guardarconfiguracionprocedimientos(objProcedimiento As ConfiguracionProcedimientoCups)
        Try

            Using comando As New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Connection = FormPrincipal.cnxion
                    comando.CommandText = objProcedimiento.consulta(2)
                    comando.Parameters.Add(New SqlParameter("@CODIGO_GRUP", SqlDbType.Int)).Value = objProcedimiento.CodigoTipo
                    comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar)).Value = objProcedimiento.usuario
                    comando.Parameters.Add(New SqlParameter("@TABLA", SqlDbType.Structured)).Value = objProcedimiento.dtProcedimiento
                    comando.ExecuteNonQuery()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw
        End Try
        Return objProcedimiento
    End Function
End Class
