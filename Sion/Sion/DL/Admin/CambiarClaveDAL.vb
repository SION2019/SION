Imports System.Data.SqlClient
Public Class CambiarClaveDAL


    Public Shared Sub guardarClave(objClave As CambiarClave)
        Try
            Using comando = New SqlCommand
                comando.CommandType = CommandType.StoredProcedure
                comando.Connection = FormPrincipal.cnxion
                comando.CommandText = "SP_CAMBIAR_CLAVE_GUARDAR"
                comando.Parameters.Add(New SqlParameter("@contranueva", SqlDbType.NVarChar)).Value = objClave.contrasenaConfirmar
                comando.Parameters.Add(New SqlParameter("@idEmpresa", SqlDbType.Int)).Value = SesionActual.idEmpresa
                comando.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int)).Value = objClave.nombreUsuario
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
