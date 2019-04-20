Imports System.Data.SqlClient
Public Class AreaServicioDAL

    Public Sub guardar(ByRef txtCodigo As Object, nombre As String, observaciones As String, codigo_menu As String, NEONATAL As Boolean, color As Integer)
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    If txtCodigo.TEXT <> "" Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_AREA_DE_SERVICIO_ACTUALIZAR"
                        comando.Parameters.Add(New SqlParameter("@Codigo_area", SqlDbType.Int))
                        comando.Parameters("@Codigo_area").Value = txtCodigo.TEXT
                        comando.Parameters.Add(New SqlParameter("@observaciones", SqlDbType.NVarChar))
                        comando.Parameters("@observaciones").Value = observaciones
                        comando.Parameters.Add(New SqlParameter("@Nombre", SqlDbType.NVarChar))
                        comando.Parameters("@Nombre").Value = nombre
                        comando.Parameters.Add(New SqlParameter("@Menu", SqlDbType.NVarChar))
                        comando.Parameters("@Menu").Value = codigo_menu
                        comando.Parameters.Add(New SqlParameter("@nEONATAL", SqlDbType.Bit))
                        comando.Parameters("@nEONATAL").Value = NEONATAL
                        comando.Parameters.Add(New SqlParameter("@Color", SqlDbType.Int)).Value = color
                        comando.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                        comando.Parameters("@Usuario_actualizacion").Value = SesionActual.idUsuario
                        comando.ExecuteNonQuery()

                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_AREA_AGREGAR"
                        comando.Parameters.Add(New SqlParameter("@Codigo_area", SqlDbType.Int))
                        comando.Parameters("@Codigo_area").Value = txtCodigo.TEXT
                        comando.Parameters.Add(New SqlParameter("@Codigo_Ep", SqlDbType.Int))
                        comando.Parameters("@Codigo_Ep").Value = SesionActual.codigoEP
                        comando.ExecuteNonQuery()
                    Else
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_AREA_DE_SERVICIO_CREAR"
                        comando.Parameters.Add(New SqlParameter("@observaciones", SqlDbType.NVarChar))
                        comando.Parameters("@observaciones").Value = observaciones
                        comando.Parameters.Add(New SqlParameter("@Nombre", SqlDbType.NVarChar))
                        comando.Parameters("@Nombre").Value = nombre
                        comando.Parameters.Add(New SqlParameter("@Menu", SqlDbType.NVarChar))
                        comando.Parameters("@Menu").Value = codigo_menu
                        comando.Parameters.Add(New SqlParameter("@nEONATAL", SqlDbType.Bit))
                        comando.Parameters("@nEONATAL").Value = NEONATAL
                        comando.Parameters.Add(New SqlParameter("@Color", SqlDbType.Int)).Value = color
                        comando.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int))
                        comando.Parameters("@Usuario_Creacion").Value = SesionActual.idUsuario
                        txtCodigo.Text = CType(comando.ExecuteScalar, String)
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_AREA_AGREGAR"
                        comando.Parameters.Add(New SqlParameter("@Codigo_area", SqlDbType.Int))
                        comando.Parameters("@Codigo_area").Value = txtCodigo.Text
                        comando.Parameters.Add(New SqlParameter("@Codigo_Ep", SqlDbType.Int))
                        comando.Parameters("@Codigo_Ep").Value = SesionActual.codigoEP
                        comando.ExecuteNonQuery()
                    End If
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Sub anular(ByRef txtCodigo As Object)
        Try
            Dim repuesta As String
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_AREA_DE_SERVICIO_ANULAR"
                    consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.NVarChar))
                    consulta.Parameters("@Usuario_actualizacion").Value = SesionActual.idUsuario
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_area", SqlDbType.Int))
                    consulta.Parameters("@CODIGO_area").Value = txtCodigo.text
                    repuesta = consulta.ExecuteScalar

                    Try
                        trnsccion.Commit()
                        If repuesta = 1 Then
                            MsgBox("Esta área de servicio esta siendo utilizada en empleados y/o admisiones, no se puede anular.", MsgBoxStyle.Information)
                        Else
                            txtCodigo.Text = Nothing
                        End If
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

End Class
