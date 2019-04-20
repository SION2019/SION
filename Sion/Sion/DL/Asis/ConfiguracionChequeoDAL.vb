Imports System.Data.SqlClient
Public Class ConfiguracionChequeoDAL

    Public Shared Sub guardarChequeo(objConfiguracion As ConfiguracionChequeo)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "SP_ACTUALIZAR_LISTA_CHEQUEO_EXCEPCION"
                comando.Parameters.Add(New SqlParameter("@codigo", SqlDbType.Int)).Value = objConfiguracion.codigoChequeo
                comando.Parameters.Add(New SqlParameter("@exepcion", SqlDbType.Int)).Value = objConfiguracion.exepcion
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub guardarChequeoProcedimiento(objConfiguracion As ConfiguracionChequeo)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "SP_LISTA_CHEQUEO_PROCEDIMIENTOS_GUARDAR"
                comando.Parameters.Add(New SqlParameter("@codigo", SqlDbType.NVarChar)).Value = objConfiguracion.codigoProcedimiento
                comando.Parameters.Add(New SqlParameter("@codigochequeo", SqlDbType.Int)).Value = objConfiguracion.codigoChequeo
                comando.Parameters.Add(New SqlParameter("@descripcion", SqlDbType.NVarChar)).Value = objConfiguracion.descripcionProcedimiento
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub ChequeoQuitar(objConfiguracion As ConfiguracionChequeo)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "SP_LISTA_CHEQUEO_CONF_PROCEDIMIENTO_QUITAR"
                comando.Parameters.Add(New SqlParameter("@codigoprocedimiento", SqlDbType.NVarChar)).Value = objConfiguracion.codigoProcedimiento
                comando.Parameters.Add(New SqlParameter("@codigochequeo", SqlDbType.Int)).Value = objConfiguracion.codigoChequeo
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
