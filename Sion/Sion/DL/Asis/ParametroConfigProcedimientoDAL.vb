Imports System.Data.SqlClient
Public Class ParametroConfigProcedimientoDAL
    Public Function GuardarconfigParametro(objParametro As ConfigParametroExam)
        Try
            Using comando As New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Connection = FormPrincipal.cnxion
                    comando.CommandText = Consultas.PARAMETRO_EXAM_CONFIGURACION_CREAR
                    comando.Parameters.Add(New SqlParameter("@CODIGO_GRUP", SqlDbType.Int)).Value = objParametro.CodigoTipo
                    comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar)).Value = objParametro.usuario
                    comando.Parameters.Add(New SqlParameter("@TABLA", SqlDbType.Structured)).Value = objParametro.dtParamsExam
                    comando.ExecuteNonQuery()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw
        End Try
        Return objParametro
    End Function
End Class
