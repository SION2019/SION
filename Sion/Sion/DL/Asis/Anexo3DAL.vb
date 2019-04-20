Imports System.Data.SqlClient
Public Class Anexo3DAL

    Public Shared Sub GuardarAnexo3(ByRef anexo3 As Anexo3)
        Try
            Using comando As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    For Each sentencia As String In anexo3.elementoAEliminar
                        If sentencia <> "" Then
                            comando.CommandText = sentencia
                            comando.ExecuteNonQuery()
                        End If
                    Next
                    comando.CommandType = CommandType.StoredProcedure
                    If anexo3.orden = "" Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_ANEXO3_CREAR"
                        comando.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.NVarChar)).Value = anexo3.registro
                        comando.Parameters.Add(New SqlParameter("@NSOLICITUD", SqlDbType.NVarChar)).Value = anexo3.nsolicitud
                        comando.Parameters.Add(New SqlParameter("@TIPOANEXO", SqlDbType.NVarChar)).Value = anexo3.tipoanexo
                        comando.Parameters.Add(New SqlParameter("@MANEJO", SqlDbType.NVarChar)).Value = anexo3.manejo
                        comando.Parameters.Add(New SqlParameter("@MODULO", SqlDbType.NVarChar)).Value = anexo3.servicioNuevo
                        comando.Parameters.Add(New SqlParameter("@REMISION", SqlDbType.Int)).Value = anexo3.remision
                        comando.Parameters.Add(New SqlParameter("@JUSTIFICACION", SqlDbType.NVarChar)).Value = anexo3.justificacion
                        comando.Parameters.Add(New SqlParameter("@TIPOSERVICIO", SqlDbType.Int)).Value = anexo3.tiposervicio
                        comando.Parameters.Add(New SqlParameter("@PRIORIDAD", SqlDbType.Int)).Value = anexo3.prioridad
                        comando.Parameters.Add(New SqlParameter("@INGRESO", SqlDbType.Int)).Value = anexo3.ingreso
                        comando.Parameters.Add(New SqlParameter("@SERVICIO", SqlDbType.NVarChar)).Value = anexo3.servicio
                        comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar)).Value = anexo3.usuario
                        anexo3.orden = CType(comando.ExecuteScalar, String)
                    Else
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_ANEXO3_ACTUALIZAR"
                        comando.Parameters.Add(New SqlParameter("@ORDEN", SqlDbType.NVarChar)).Value = anexo3.orden
                        comando.Parameters.Add(New SqlParameter("@NSOLICITUD", SqlDbType.NVarChar)).Value = anexo3.nsolicitud
                        comando.Parameters.Add(New SqlParameter("@MANEJO", SqlDbType.NVarChar)).Value = anexo3.manejo
                        comando.Parameters.Add(New SqlParameter("@MODULO", SqlDbType.NVarChar)).Value = anexo3.servicioNuevo
                        comando.Parameters.Add(New SqlParameter("@REMISION", SqlDbType.Int)).Value = anexo3.remision
                        comando.Parameters.Add(New SqlParameter("@JUSTIFICACION", SqlDbType.NVarChar)).Value = anexo3.justificacion
                        comando.Parameters.Add(New SqlParameter("@TIPOSERVICIO", SqlDbType.Int)).Value = anexo3.tiposervicio
                        comando.Parameters.Add(New SqlParameter("@PRIORIDAD", SqlDbType.Int)).Value = anexo3.prioridad
                        comando.Parameters.Add(New SqlParameter("@INGRESO", SqlDbType.Int)).Value = anexo3.ingreso
                        comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar)).Value = anexo3.usuario
                        comando.ExecuteNonQuery()

                    End If

                    For i = 0 To anexo3.dtanexo.Rows.Count - 2
                        If anexo3.dtanexo.Rows(i).Item("Estado") = Constantes.ITEM_NUEVO Then
                            comando.Parameters.Clear()
                            comando.CommandText = "SP_ANEXO3_DETALLE_CREAR"
                            comando.Parameters.Add(New SqlParameter("@Cod_Servicio", SqlDbType.NVarChar)).Value = anexo3.dtanexo.Rows(i).Item(0).ToString
                            comando.Parameters.Add(New SqlParameter("@finicio", SqlDbType.Date)).Value = anexo3.dtanexo.Rows(i).Item(2).ToString
                            comando.Parameters.Add(New SqlParameter("@Ffin", SqlDbType.Date)).Value = anexo3.dtanexo.Rows(i).Item(3).ToString
                            comando.Parameters.Add(New SqlParameter("@Dias", SqlDbType.NVarChar)).Value = anexo3.dtanexo.Rows(i).Item(4).ToString
                            comando.Parameters.Add(New SqlParameter("@orden", SqlDbType.NVarChar)).Value = anexo3.orden
                            comando.ExecuteNonQuery()
                            comando.Parameters.Clear()
                        ElseIf anexo3.dtanexo.Rows(i).Item("Estado") = Constantes.ITEM_CARGADO Then
                            comando.Parameters.Clear()
                            comando.CommandText = "SP_ANEXO3_DETALLE_ACTUALIZAR"
                            comando.Parameters.Add(New SqlParameter("@Cod_Servicio", SqlDbType.NVarChar)).Value = anexo3.dtanexo.Rows(i).Item(0).ToString
                            comando.Parameters.Add(New SqlParameter("@orden", SqlDbType.NVarChar)).Value = anexo3.orden
                            comando.Parameters.Add(New SqlParameter("@finicio", SqlDbType.Date)).Value = anexo3.dtanexo.Rows(i).Item(2).ToString
                            comando.Parameters.Add(New SqlParameter("@Ffin", SqlDbType.Date)).Value = anexo3.dtanexo.Rows(i).Item(3).ToString
                            comando.Parameters.Add(New SqlParameter("@Dias", SqlDbType.NVarChar)).Value = anexo3.dtanexo.Rows(i).Item(4).ToString
                            comando.ExecuteNonQuery()
                        End If
                    Next

                    Try
                        trnsccion.Commit()
                    Catch ex As SqlException
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
