Imports System.Data.SqlClient
Public Class HemoderivadoConfigProcedimientoDAL
    Public Function Guardarconfiguracionprocedimientos(ByVal dt As DataTable, ByVal codigoGrupo As String) As Boolean
        Dim respuesta As Integer
        Using Consulta As New SqlCommand()
            Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                Consulta.Transaction = trnsccion
                Consulta.CommandType = CommandType.StoredProcedure
                Consulta.Connection = FormPrincipal.cnxion
                For fila = 0 To dt.Rows.Count - 1
                    If dt.Rows(fila).Item("Estado") = Constantes.ESTADO_PENDIENTE Then
                        Consulta.CommandText = "[SP_CONFIG_PROCED_HEMODERIVADOS_CREAR]"
                        Consulta.Parameters.Clear()
                        Consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar))
                        Consulta.Parameters("@CODIGO").Value = dt.Rows(fila).Item("Codigo")
                        Consulta.Parameters.Add(New SqlParameter("@CODIGO_GRUP", SqlDbType.Int))
                        Consulta.Parameters("@CODIGO_GRUP").Value = codigoGrupo
                        Consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                        Consulta.Parameters("@USUARIO").Value = SesionActual.idUsuario
                        Consulta.ExecuteNonQuery()
                    ElseIf dt.Rows(fila).Item("Estado") = Constantes.ESTADO_QUITADO Then
                        Consulta.CommandText = "[SP_CONFIG_PROCED_HEMODERIVADOS_ACTUALIZAR]"
                        Consulta.Parameters.Clear()
                        Consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar))
                        Consulta.Parameters("@CODIGO").Value = dt.Rows(fila).Item("Codigo")
                        Consulta.Parameters.Add(New SqlParameter("@CODIGO_GRUP", SqlDbType.Int))
                        Consulta.Parameters("@CODIGO_GRUP").Value = codigoGrupo
                        Consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                        Consulta.Parameters("@USUARIO").Value = SesionActual.idUsuario
                        Consulta.ExecuteNonQuery()
                    End If
                Next

                Try
                    respuesta = 1
                    trnsccion.Commit()
                Catch ex As Exception
                    trnsccion.Rollback()
                    general.manejoErrores(ex) 
                End Try
            End Using
        End Using

        If respuesta = 1 Then
            Return True
        Else
            Return False
        End If

    End Function
End Class
