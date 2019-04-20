Imports System.Data.SqlClient

Public Class TrasladoCamaDAL

    Public Shared Sub actualizarRegistroCama(pRegistro As Integer,
                                      pCodigoCama As String)
        Try
            Using dbCommand As New SqlCommand("SP_REGISTRO_CAMA_ACTUALIZAR")
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.Connection = FormPrincipal.cnxion

                dbCommand.Parameters.Add(New SqlParameter("@pREGISTRO", SqlDbType.Int))
                dbCommand.Parameters.Add(New SqlParameter("@pCODIGO_CAMA", SqlDbType.Int))
                dbCommand.Parameters.Add(New SqlParameter("@pUsuario_Actualizacion", SqlDbType.Int))

                dbCommand.Parameters("@pREGISTRO").Value = pRegistro
                dbCommand.Parameters("@pCODIGO_CAMA").Value = pCodigoCama
                dbCommand.Parameters("@pUsuario_Actualizacion").Value = SesionActual.idUsuario

                dbCommand.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
