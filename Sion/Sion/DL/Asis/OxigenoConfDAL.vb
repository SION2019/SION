Imports System.Data.SqlClient
Public Class OxigenoConfDAL

    Public Sub guardarOxigeno(ByRef txtCodigo As Object, ByVal descripcion As String,
                                   ByVal factor As Double,
                                   ByVal porcentaje As Double,
                                   ByVal codigoEquivalencia As Integer)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure


                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_OXIGENO_CREAR]"
                    comando.Parameters.Add(New SqlParameter("@pDescripcion", SqlDbType.NVarChar)).Value = descripcion
                    comando.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = SesionActual.idUsuario
                    txtCodigo.text = CType(comando.ExecuteScalar, String)

                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_OXIGENO_DETALLE_CREAR]"
                    comando.Parameters.Add(New SqlParameter("@pCodigo", SqlDbType.Int)).Value = txtCodigo.text
                    comando.Parameters.Add(New SqlParameter("@pCodigoInterno", SqlDbType.Int)).Value = codigoEquivalencia
                    comando.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = SesionActual.idUsuario
                    comando.Parameters.Add(New SqlParameter("@Factor", SqlDbType.Float)).Value = factor
                    comando.Parameters.Add(New SqlParameter("@porcentaje", SqlDbType.Float)).Value = porcentaje
                    comando.ExecuteNonQuery()

                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        trnsccion.Rollback()
                        Throw ex
                    End Try
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub actualizarOxigeno(ByVal codigo As Integer,
                                      ByVal descripcion As String,
                                      ByVal factor As Double,
                                      ByVal porcentaje As Double,
                                      ByVal codigoEquivalencia As Integer)

        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure

                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_OXIGENO_ACTUALIZAR]"
                    comando.Parameters.Add(New SqlParameter("@pCodigo", SqlDbType.NVarChar)).Value = codigo
                    comando.Parameters.Add(New SqlParameter("@pDescripcion", SqlDbType.NVarChar)).Value = descripcion
                    comando.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = SesionActual.idUsuario
                    comando.ExecuteNonQuery()

                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_OXIGENO_DETALLE_ACTUALIZAR]"
                    comando.Parameters.Add(New SqlParameter("@pCodigo", SqlDbType.Int)).Value = codigo
                    comando.Parameters.Add(New SqlParameter("@pCodigoInterno", SqlDbType.Int)).Value = codigoEquivalencia
                    comando.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = SesionActual.idUsuario
                    comando.Parameters.Add(New SqlParameter("@Factor", SqlDbType.Float)).Value = factor
                    comando.Parameters.Add(New SqlParameter("@porcentaje", SqlDbType.Float)).Value = porcentaje
                    comando.ExecuteNonQuery()

                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        trnsccion.Rollback()
                        Throw ex
                    End Try
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
End Class
