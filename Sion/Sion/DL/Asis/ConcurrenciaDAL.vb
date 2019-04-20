Imports System.Data.SqlClient
Public Class ConcurrenciaDAL

    Public Sub guardarConcurrencia(objConcurrencia As ConcurrenciaPaciente, fecha As Date)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "SP_CONCURRENCIA_PACIENTE_CREAR"
                comando.Parameters.Add(New SqlParameter("@detalle", SqlDbType.Structured)).Value = objConcurrencia.dtConcurrencia
                comando.Parameters.Add(New SqlParameter("@fecha", SqlDbType.Date)).Value = fecha
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
