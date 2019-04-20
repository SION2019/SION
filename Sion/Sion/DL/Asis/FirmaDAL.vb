Imports System.Data.SqlClient
Public Class FirmaDAL

    Public Shared Sub guardarFirma(ByVal codigo As Integer, ByVal arrFile() As Byte, ByVal consulta As String, usuarioEnviar As Boolean)
        Try
            Using dbCommand As New SqlCommand(consulta)
                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.Parameters.Add(New SqlParameter("@N_REGISTRO", SqlDbType.Int))
                dbCommand.Parameters("@N_REGISTRO").Value = codigo
                If usuarioEnviar = True Then
                    dbCommand.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int))
                    dbCommand.Parameters("@usuario").Value = SesionActual.idUsuario
                End If
                dbCommand.Parameters.Add(New SqlParameter("@IMAGEN", SqlDbType.Image))
                dbCommand.Parameters("@IMAGEN").Value = arrFile
                dbCommand.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
