Imports System.Data.SqlClient
Public Class Anexo1DAL


    Public Sub GuardarAnexo(ByRef orden As Object,
                                 ByVal registro As String,
                                 ByVal servicio As String,
                                 ByVal Solicitud As String,
                                observacion As String,
                                 ByVal usuario As String, dt_anexo As DataTable)


        Try

            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()

                    If orden.text = "" Then
                        comando.Parameters.Clear()
                        comando.CommandType = CommandType.StoredProcedure
                        comando.CommandText = "SP_ANEXO_CREAR"
                        comando.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.NVarChar))
                        comando.Parameters("@REGISTRO").Value = registro
                        comando.Parameters.Add(New SqlParameter("@SERVICIO", SqlDbType.Bit))
                        comando.Parameters("@SERVICIO").Value = servicio
                        comando.Parameters.Add(New SqlParameter("@SOLICITUD", SqlDbType.NVarChar))
                        comando.Parameters("@SOLICITUD").Value = Solicitud
                        comando.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        comando.Parameters("@OBSERVACION").Value = observacion
                        comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                        comando.Parameters("@USUARIO").Value = usuario
                        orden.text = CType(comando.ExecuteScalar, String)

                        For i As Int32 = 0 To dt_anexo.Rows.Count - 1
                            comando.Parameters.Clear()
                            comando.CommandText = "SP_ANEXO_DETALLE_CREAR"
                            comando.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                            comando.Parameters("@CODIGO").Value = CInt(orden.text)
                            comando.Parameters.Add(New SqlParameter("@DESCRIPCION", SqlDbType.NVarChar))
                            comando.Parameters("@DESCRIPCION").Value = dt_anexo.Rows(i).Item(0).ToString
                            comando.Parameters.Add(New SqlParameter("@DATOS_DOCUMENTO", SqlDbType.NVarChar))
                            comando.Parameters("@DATOS_DOCUMENTO").Value = dt_anexo.Rows(i).Item(1).ToString
                            comando.Parameters.Add(New SqlParameter("@DATOS_BASE", SqlDbType.NVarChar))
                            comando.Parameters("@DATOS_BASE").Value = dt_anexo.Rows(i).Item(2).ToString
                            comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                            comando.Parameters("@USUARIO").Value = usuario
                            comando.ExecuteNonQuery()
                        Next
                    Else
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_ANEXO_ACTUALIZAR"
                        comando.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar))
                        comando.Parameters("@CODIGO").Value = CInt(orden.text)
                        comando.Parameters.Add(New SqlParameter("@SERVICIO", SqlDbType.Bit))
                        comando.Parameters("@SERVICIO").Value = servicio
                        comando.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        comando.Parameters("@OBSERVACION").Value = observacion
                        comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                        comando.Parameters("@USUARIO").Value = usuario
                        comando.ExecuteNonQuery()


                        comando.Parameters.Clear()
                        For i As Int32 = 0 To dt_anexo.Rows.Count - 1
                            comando.Parameters.Clear()
                            comando.CommandText = "SP_ANEXO_DETALLE_ACTUALIZAR"
                            comando.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar))
                            comando.Parameters("@CODIGO").Value = orden.text
                            comando.Parameters.Add(New SqlParameter("@DESCRIPCION", SqlDbType.NVarChar))
                            comando.Parameters("@DESCRIPCION").Value = dt_anexo.Rows(i).Item(0).ToString
                            comando.Parameters.Add(New SqlParameter("@DATOS_DOCUMENTO", SqlDbType.NVarChar))
                            comando.Parameters("@DATOS_DOCUMENTO").Value = dt_anexo.Rows(i).Item(1).ToString
                            comando.Parameters.Add(New SqlParameter("@DATOS_BASE", SqlDbType.NVarChar))
                            comando.Parameters("@DATOS_BASE").Value = dt_anexo.Rows(i).Item(2).ToString
                            comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                            comando.Parameters("@USUARIO").Value = usuario
                            comando.ExecuteNonQuery()
                        Next
                    End If
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        trnsccion.Rollback()
                        general.manejoErrores(ex) 
                    End Try
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
