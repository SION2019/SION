Imports System.Data.SqlClient
Public Class FuripsDAL

    Public Shared Sub guardarRegistro(objFurips As FuRips)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "SP_FURIPS_CREAR"
                comando.Parameters.Add(New SqlParameter("@tabla", SqlDbType.Structured)).Value = objFurips.dtParametro
                comando.Parameters.Add(New SqlParameter("@codigo", SqlDbType.Int)).Value = objFurips.codigo
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objFurips.registro
                comando.Parameters.Add(New SqlParameter("@descripcion", SqlDbType.Text)).Value = objFurips.descripcion
                comando.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int)).Value = objFurips.usuario
                objFurips.codigo = CType(comando.ExecuteScalar(), Integer)
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
