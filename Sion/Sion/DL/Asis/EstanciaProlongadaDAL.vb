Imports System.Data.SqlClient
Public Class EstanciaProlongadaDAL

    Public Shared Sub crearEstancia(ByRef objEstancia As EstanciaProlongada)
        Try
            Using consulta As New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_ESTANCIA_PROLONGADA_CREAR"
                consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objEstancia.registro
                consulta.Parameters.Add(New SqlParameter("@CODIGOEVO", SqlDbType.Int)).Value = objEstancia.codigoevo
                consulta.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.Date)).Value = objEstancia.fecha
                consulta.Parameters.Add(New SqlParameter("@JUSTIFICACION", SqlDbType.NVarChar)).Value = objEstancia.justificacion
                consulta.Parameters.Add(New SqlParameter("@RESULTADO", SqlDbType.NVarChar)).Value = objEstancia.resultado
                consulta.Parameters.Add(New SqlParameter("@CONDUCTA", SqlDbType.NVarChar)).Value = objEstancia.conducta
                consulta.Parameters.Add(New SqlParameter("@EMPLEADO", SqlDbType.Int)).Value = objEstancia.idempleado
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objEstancia.codigoEp
                consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objEstancia.usuario
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub crearEstanciaR(ByRef objEstancia As EstanciaProlongadaR)
        Try
            Using consulta As New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_ESTANCIA_PROLONGADA_CREAR_R"
                consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objEstancia.registro
                consulta.Parameters.Add(New SqlParameter("@CODIGOEVO", SqlDbType.Int)).Value = objEstancia.codigoevo
                consulta.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.Date)).Value = objEstancia.fecha
                consulta.Parameters.Add(New SqlParameter("@JUSTIFICACION", SqlDbType.NVarChar)).Value = objEstancia.justificacion
                consulta.Parameters.Add(New SqlParameter("@RESULTADO", SqlDbType.NVarChar)).Value = objEstancia.resultado
                consulta.Parameters.Add(New SqlParameter("@CONDUCTA", SqlDbType.NVarChar)).Value = objEstancia.conducta
                consulta.Parameters.Add(New SqlParameter("@EMPLEADO", SqlDbType.Int)).Value = objEstancia.idempleado
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objEstancia.codigoEp
                consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objEstancia.usuario
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub crearEstanciaRR(ByRef objEstancia As EstanciaProlongadaRR)
        Try
            Using consulta As New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "SP_ESTANCIA_PROLONGADA_CREAR_RR"
                consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objEstancia.registro
                consulta.Parameters.Add(New SqlParameter("@CODIGOEVO", SqlDbType.Int)).Value = objEstancia.codigoevo
                consulta.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.Date)).Value = objEstancia.fecha
                consulta.Parameters.Add(New SqlParameter("@JUSTIFICACION", SqlDbType.NVarChar)).Value = objEstancia.justificacion
                consulta.Parameters.Add(New SqlParameter("@RESULTADO", SqlDbType.NVarChar)).Value = objEstancia.resultado
                consulta.Parameters.Add(New SqlParameter("@CONDUCTA", SqlDbType.NVarChar)).Value = objEstancia.conducta
                consulta.Parameters.Add(New SqlParameter("@EMPLEADO", SqlDbType.Int)).Value = objEstancia.idempleado
                consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objEstancia.usuario
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
