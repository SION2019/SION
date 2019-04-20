Imports System.Data.SqlClient
Public Class ResolucionDAL
    Public Sub actualizarResolucion(objResolucion As Resolucion)

        Try
            Using comando As New SqlCommand("SP_RESOLUCION_ACTUALIZAR")
                comando.CommandType = CommandType.StoredProcedure
                comando.Connection = FormPrincipal.cnxion
                comando.Parameters.Add(New SqlParameter("@IDEMPRESA", SqlDbType.Int))
                comando.Parameters("@IDEMPRESA").Value = objResolucion.idEmpresa
                comando.Parameters.Add(New SqlParameter("@RESOLUCION", SqlDbType.NVarChar))
                comando.Parameters("@RESOLUCION").Value = objResolucion.resolucion
                comando.Parameters.Add(New SqlParameter("@PREFIJO", SqlDbType.NVarChar))
                comando.Parameters("@PREFIJO").Value = objResolucion.prefijo
                comando.Parameters.Add(New SqlParameter("@CONSE_INI", SqlDbType.NVarChar))
                comando.Parameters("@CONSE_INI").Value = objResolucion.conseInic
                comando.Parameters.Add(New SqlParameter("@CONSE_ACTUAL", SqlDbType.NVarChar))
                comando.Parameters("@CONSE_ACTUAL").Value = objResolucion.conseActual
                comando.Parameters.Add(New SqlParameter("@CONSE_FIN", SqlDbType.NVarChar))
                comando.Parameters("@CONSE_FIN").Value = objResolucion.conseFin
                comando.Parameters.Add(New SqlParameter("@FECHAEXP", SqlDbType.Date))
                comando.Parameters("@FECHAEXP").Value = objResolucion.fechaExpedicion
                comando.Parameters.Add(New SqlParameter("@FECHAVENCE", SqlDbType.Date))
                comando.Parameters("@FECHAVENCE").Value = objResolucion.fechaVencimiento
                comando.Parameters.Add(New SqlParameter("@VIGENTE", SqlDbType.Bit))
                comando.Parameters("@VIGENTE").Value = objResolucion.vigente
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                comando.Parameters("@USUARIO").Value = objResolucion.usuario
                comando.Parameters.Add(New SqlParameter("@CODIGORES", SqlDbType.NVarChar))
                comando.Parameters("@CODIGORES").Value = objResolucion.codigo
                comando.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Function guardarResolucion(objResolucion As Resolucion) As String

        Try

            Using comando As New SqlCommand("SP_RESOLUCION_CREAR")
                comando.CommandType = CommandType.StoredProcedure
                comando.Connection = FormPrincipal.cnxion
                comando.Parameters.Add(New SqlParameter("@IDEMPRESA", SqlDbType.Int))
                comando.Parameters("@IDEMPRESA").Value = objResolucion.idEmpresa
                comando.Parameters.Add(New SqlParameter("@RESOLUCION", SqlDbType.NVarChar))
                comando.Parameters("@RESOLUCION").Value = objResolucion.resolucion
                comando.Parameters.Add(New SqlParameter("@PREFIJO", SqlDbType.NVarChar))
                comando.Parameters("@PREFIJO").Value = objResolucion.prefijo
                comando.Parameters.Add(New SqlParameter("@CONSE_INI", SqlDbType.NVarChar))
                comando.Parameters("@CONSE_INI").Value = objResolucion.conseInic
                comando.Parameters.Add(New SqlParameter("@CONSE_ACTUAL", SqlDbType.NVarChar))
                comando.Parameters("@CONSE_ACTUAL").Value = objResolucion.conseActual
                comando.Parameters.Add(New SqlParameter("@CONSE_FIN", SqlDbType.NVarChar))
                comando.Parameters("@CONSE_FIN").Value = objResolucion.conseFin
                comando.Parameters.Add(New SqlParameter("@FECHAEXP", SqlDbType.Date))
                comando.Parameters("@FECHAEXP").Value = objResolucion.fechaExpedicion
                comando.Parameters.Add(New SqlParameter("@FECHAVENCE", SqlDbType.Date))
                comando.Parameters("@FECHAVENCE").Value = objResolucion.fechaVencimiento
                comando.Parameters.Add(New SqlParameter("@VIGENTE", SqlDbType.Bit))
                comando.Parameters("@VIGENTE").Value = objResolucion.vigente
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                comando.Parameters("@USUARIO").Value = objResolucion.usuario
                objResolucion.codigo = CType(comando.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception
            Throw ex
        End Try

        Return objResolucion.codigo
    End Function

End Class
