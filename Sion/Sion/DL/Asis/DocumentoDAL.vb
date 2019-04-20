Imports System.Data.SqlClient
Public Class DocumentoDAL
    Public Function guardarDocumento(objDocumento As TipoDocumentos) As String

        Try

            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = "SP_TIPO_DOCUMENTO_CREAR"
                dbCommand.Parameters.Add(New SqlParameter("@DESCRIPCION", SqlDbType.NVarChar))
                dbCommand.Parameters("@DESCRIPCION").Value = objDocumento.descripcion
                dbCommand.Parameters.Add(New SqlParameter("@TIPO", SqlDbType.NChar))
                dbCommand.Parameters("@TIPO").Value = objDocumento.tipoDocumento
                dbCommand.Parameters.Add(New SqlParameter("@SERVICIO", SqlDbType.Int))
                dbCommand.Parameters("@SERVICIO").Value = objDocumento.areaServicio
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                dbCommand.Parameters("@USUARIO").Value = objDocumento.usuario
                objDocumento.codigo = CType(dbCommand.ExecuteScalar, Integer)

            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return objDocumento.codigo
    End Function
    Public Sub actualizarDocumento(objDocumento As TipoDocumentos)
        Try
            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = "SP_TIPO_DOCUMENTO_ACTUALIZAR"
                dbCommand.Parameters.Add(New SqlParameter("@DESCRIPCION", SqlDbType.NVarChar))
                dbCommand.Parameters("@DESCRIPCION").Value = objDocumento.descripcion
                dbCommand.Parameters.Add(New SqlParameter("@TIPO", SqlDbType.NChar))
                dbCommand.Parameters("@TIPO").Value = objDocumento.tipoDocumento
                dbCommand.Parameters.Add(New SqlParameter("@SERVICIO", SqlDbType.Int))
                dbCommand.Parameters("@SERVICIO").Value = objDocumento.areaServicio
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                dbCommand.Parameters("@USUARIO").Value = objDocumento.usuario
                dbCommand.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                dbCommand.Parameters("@CODIGO").Value = objDocumento.codigo
                dbCommand.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
