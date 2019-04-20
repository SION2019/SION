Imports System.Data.SqlClient
Public Class ManualTarifarioDAL

    Public Shared Sub guardar(objManual As ManualTarifario)

        Try
            Using consulta As New SqlCommand("SP_MANUAL_TARIFARIO_CREAR")
                consulta.CommandType = CommandType.StoredProcedure
                consulta.Connection = FormPrincipal.cnxion
                consulta.Parameters.Add(New SqlParameter("@CODIGOMANUAL", SqlDbType.Int))
                consulta.Parameters("@CODIGOMANUAL").Value = objManual.codigo
                consulta.Parameters.Add(New SqlParameter("@CODIGOCUPS", SqlDbType.NVarChar))
                consulta.Parameters("@CODIGOCUPS").Value = objManual.codigoCups
                consulta.Parameters.Add(New SqlParameter("@CODIGOSOAT", SqlDbType.NVarChar))
                consulta.Parameters("@CODIGOSOAT").Value = objManual.codigoSoat
                consulta.Parameters.Add(New SqlParameter("@CODIGOISS", SqlDbType.NVarChar))
                consulta.Parameters("@CODIGOISS").Value = objManual.codigoIss
                consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                consulta.Parameters("@USUARIO").Value = objManual.usuario
                consulta.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
