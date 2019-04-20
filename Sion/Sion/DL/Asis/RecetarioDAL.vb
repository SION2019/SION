Imports System.Data.SqlClient
Public Class RecetarioDAL

    Public Sub crearRecetario(objRecetario As Recetario, codigo As Object)
        Try
            Using comando As New SqlCommand
                If codigo.text = "" Then
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Connection = FormPrincipal.cnxion
                    comando.CommandText = "SP_RECETARIO_CREAR"
                    comando.Parameters.Add(New SqlParameter("@Resultado", SqlDbType.NVarChar))
                    comando.Parameters("@Resultado").Value = objRecetario.resultado
                    comando.Parameters.Add(New SqlParameter("@Registro", SqlDbType.NVarChar))
                    comando.Parameters("@Registro").Value = objRecetario.registro
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.NVarChar))
                    comando.Parameters("@Usuario").Value = objRecetario.usuario
                    comando.Parameters.Add(New SqlParameter("@UsuarioReal", SqlDbType.NVarChar))
                    comando.Parameters("@UsuarioReal").Value = objRecetario.usuarioReal
                    codigo.text = (comando.ExecuteScalar)
                Else
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Connection = FormPrincipal.cnxion
                    comando.CommandText = "SP_RECETARIO_ACTUALIZAR"
                    comando.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar))
                    comando.Parameters("@Codigo").Value = codigo.text
                    comando.Parameters.Add(New SqlParameter("@Resultado", SqlDbType.NVarChar))
                    comando.Parameters("@Resultado").Value = objRecetario.resultado
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.NVarChar))
                    comando.Parameters("@Usuario").Value = objRecetario.usuario
                    comando.Parameters.Add(New SqlParameter("@UsuarioReal", SqlDbType.NVarChar))
                    comando.Parameters("@UsuarioReal").Value = objRecetario.usuarioReal
                    comando.ExecuteNonQuery()
                End If
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
