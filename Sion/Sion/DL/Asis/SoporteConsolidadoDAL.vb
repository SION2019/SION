Imports System.Data.SqlClient
Public Class SoporteConsolidadoDAL

    Public Shared Sub guadarConfiguracion(objSoporte As SoporteFacturacion)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "SP_SOPORTE_CONSOLIDADO_CREAR"
                comando.Parameters.Add(New SqlParameter("@CARPETAS", SqlDbType.Structured)).Value = objSoporte.dtFacturacion
                comando.Parameters.Add(New SqlParameter("@DETALLES", SqlDbType.Structured)).Value = objSoporte.dtConfigurado
                comando.Parameters.Add(New SqlParameter("@IDEPS", SqlDbType.Int)).Value = objSoporte.idEps
                comando.Parameters.Add(New SqlParameter("@ACCION", SqlDbType.Int)).Value = objSoporte.accion

                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
