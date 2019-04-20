Imports System.Data.SqlClient
Public Class ManualTarifarioListaDAL
    Public Sub guardar_manual(ByRef txtCodigo As Object, ByVal descripcion As String, opcion As Boolean, manualDuplicar As Integer)
        Dim opcionR As Integer
        If opcion = True Then
            opcionR = 1
        Else
            opcionR = 0
        End If
        Try
            If txtCodigo.TEXT = "" Then

                Using Consulta As New SqlCommand("SP_MANUAL_CREAR")
                    Consulta.CommandType = CommandType.StoredProcedure
                    Consulta.Connection = FormPrincipal.cnxion
                    Consulta.Parameters.Add(New SqlParameter("@NOMBRE", SqlDbType.NVarChar))
                    Consulta.Parameters("@NOMBRE").Value = descripcion
                    Consulta.Parameters.Add(New SqlParameter("@Usuario_creacion", SqlDbType.NVarChar))
                    Consulta.Parameters("@Usuario_creacion").Value = SesionActual.idUsuario
                    Consulta.Parameters.Add(New SqlParameter("@OPCION", SqlDbType.Int))
                    Consulta.Parameters("@OPCION").Value = opcionR
                    Consulta.Parameters.Add(New SqlParameter("@MANUALDUPLICAR", SqlDbType.Int))
                    Consulta.Parameters("@MANUALDUPLICAR").Value = manualDuplicar
                    txtCodigo.TEXT = CType(Consulta.ExecuteScalar, String)
                End Using
            Else
                Using Consulta As New SqlCommand("SP_MANUAL_ACTUALIZAR")
                    Consulta.CommandType = CommandType.StoredProcedure
                    Consulta.Connection = FormPrincipal.cnxion
                    Consulta.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar))
                    Consulta.Parameters("@Codigo").Value = txtCodigo.TEXT
                    Consulta.Parameters.Add(New SqlParameter("@NOMBRE", SqlDbType.NVarChar))
                    Consulta.Parameters("@NOMBRE").Value = descripcion
                    Consulta.Parameters.Add(New SqlParameter("@USUARIO_ACTUALIZACION", SqlDbType.NVarChar))
                    Consulta.Parameters("@USUARIO_ACTUALIZACION").Value = SesionActual.idUsuario
                    Consulta.ExecuteNonQuery()
                End Using
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub anular_MANUAl(ByVal codigo As String)
        Try
            Using Consulta As New SqlCommand("SP_MANUAL_ANULAR")
                Consulta.CommandType = CommandType.StoredProcedure
                Consulta.Connection = FormPrincipal.cnxion
                Consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                Consulta.Parameters("@CODIGO").Value = CInt(codigo)
                Consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                Consulta.Parameters("@USUARIO").Value = SesionActual.idUsuario
                Consulta.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
