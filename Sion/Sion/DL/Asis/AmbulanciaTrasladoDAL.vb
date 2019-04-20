Imports System.Data.SqlClient
Public Class Traslado_Ambulancia_C
    Public Function guardarTraslado(objTraslado As TrasladoAmbulancia) As String
        Try

            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = "SP_PROCEDIMIENTO_TRASLADO_CREAR"
                dbCommand.Parameters.Add(New SqlParameter("@COIGO_CUPS", SqlDbType.NVarChar))
                dbCommand.Parameters("@COIGO_CUPS").Value = objTraslado.codigocups
                dbCommand.Parameters.Add(New SqlParameter("@CODIGO_EPS", SqlDbType.Int))
                dbCommand.Parameters("@CODIGO_EPS").Value = objTraslado.idCliente
                dbCommand.Parameters.Add(New SqlParameter("@VALOR", SqlDbType.Int))
                dbCommand.Parameters("@VALOR").Value = objTraslado.valorTraslado
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                dbCommand.Parameters("@USUARIO").Value = objTraslado.usuario
                objTraslado.codigo = CType(dbCommand.ExecuteScalar, Integer)
            End Using
        Catch ex As Exception

            general.manejoErrores(ex) 
        End Try
        Return objTraslado.codigo
    End Function
    Public Sub actualizarTraslado(objTraslado As TrasladoAmbulancia)
        Try

            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = "SP_PROCEDIMIENTO_TRASLADO_ACTUALIZAR"
                dbCommand.Parameters.Add(New SqlParameter("@CODIGO_CUPS", SqlDbType.NVarChar))
                dbCommand.Parameters("@CODIGO_CUPS").Value = objTraslado.codigocups
                dbCommand.Parameters.Add(New SqlParameter("@CODIGO_EPS", SqlDbType.Int))
                dbCommand.Parameters("@CODIGO_EPS").Value = objTraslado.idCliente
                dbCommand.Parameters.Add(New SqlParameter("@VALOR", SqlDbType.Int))
                dbCommand.Parameters("@VALOR").Value = objTraslado.valorTraslado
                dbCommand.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                dbCommand.Parameters("@USUARIO").Value = objTraslado.usuario
                dbCommand.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                dbCommand.Parameters("@CODIGO").Value = objTraslado.codigo
                dbCommand.ExecuteNonQuery()

            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
End Class
