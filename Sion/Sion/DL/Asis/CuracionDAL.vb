Imports System.Data.SqlClient
Public Class CuracionDAL

    Public Function crearCuracion(objCuracion As Curacion) As String
        Try
            Using comando = New SqlCommand
                comando.CommandType = CommandType.StoredProcedure
                comando.Connection = FormPrincipal.cnxion
                comando.CommandText = "SP_CURACION_CREAR"
                comando.Parameters.Add(New SqlParameter("@Registro", SqlDbType.Int)).Value = objCuracion.Registro
                comando.Parameters.Add(New SqlParameter("@fechaCuracion", SqlDbType.Date)).Value = objCuracion.fechaCuracion
                comando.Parameters.Add(New SqlParameter("@enfermera", SqlDbType.NVarChar)).Value = objCuracion.enfermera
                comando.Parameters.Add(New SqlParameter("@descripcion", SqlDbType.NVarChar)).Value = objCuracion.descripcion
                comando.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int)).Value = objCuracion.usuario
                objCuracion.codigoCuracion = CType(comando.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objCuracion.codigoCuracion
    End Function

    Public Function actualizarCuracion(objCuracion As Curacion) As String
        Try
            Using comando = New SqlCommand
                comando.CommandType = CommandType.StoredProcedure
                comando.Connection = FormPrincipal.cnxion
                comando.CommandText = "SP_CURACION_ACTUALIZAR"
                comando.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.Int)).Value = objCuracion.codigoCuracion
                comando.Parameters.Add(New SqlParameter("@fechaCuracion", SqlDbType.Date)).Value = objCuracion.fechaCuracion
                comando.Parameters.Add(New SqlParameter("@enfermera", SqlDbType.NVarChar)).Value = objCuracion.enfermera
                comando.Parameters.Add(New SqlParameter("@descripcion", SqlDbType.NVarChar)).Value = objCuracion.descripcion
                comando.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int)).Value = objCuracion.usuario
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objCuracion.codigoCuracion
    End Function
End Class
