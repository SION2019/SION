Imports System.Data.SqlClient
Public Class Clientes_C
    Public Sub Guardarcliente(objCliente As Cliente)
        Try

            Using consulta As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_CLIENTE_CREAR"
                    consulta.Parameters.Add(New SqlParameter("@IDCLIENTE", SqlDbType.Int))
                    consulta.Parameters("@IDCLIENTE").Value = objCliente.IdCliente
                    consulta.Parameters.Add(New SqlParameter("@FormaPago", SqlDbType.NVarChar))
                    consulta.Parameters("@FormaPago").Value = objCliente.formapago
                    consulta.Parameters.Add(New SqlParameter("@Codigo_tipo_regimen", SqlDbType.Int))
                    consulta.Parameters("@Codigo_tipo_regimen").Value = objCliente.codigoregimen
                    consulta.Parameters.Add(New SqlParameter("@Ubicacion", SqlDbType.NVarChar))
                    consulta.Parameters("@Ubicacion").Value = objCliente.ubicacion
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Plazo", SqlDbType.Int))
                    consulta.Parameters("@Codigo_Plazo").Value = objCliente.codigoplazo
                    consulta.Parameters.Add(New SqlParameter("@Codigo_dia_entrega", SqlDbType.Int))
                    consulta.Parameters("@Codigo_dia_entrega").Value = objCliente.codigoentrega
                    consulta.Parameters.Add(New SqlParameter("@Codigo_dia_devolucion", SqlDbType.Int))
                    consulta.Parameters("@Codigo_dia_devolucion").Value = objCliente.codigodevolucion
                    consulta.Parameters.Add(New SqlParameter("@Observaciones", SqlDbType.NVarChar))
                    consulta.Parameters("@Observaciones").Value = objCliente.observacion
                    consulta.Parameters.Add(New SqlParameter("@Suspendido", SqlDbType.Bit))
                    consulta.Parameters("@Suspendido").Value = objCliente.suspendido
                    consulta.Parameters.Add(New SqlParameter("@Bloqueado", SqlDbType.Bit))
                    consulta.Parameters("@Bloqueado").Value = objCliente.bloqueado
                    consulta.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                    consulta.Parameters("@Id_empresa").Value = objCliente.idEmpresa
                    consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@USUARIO").Value = objCliente.usuario
                    consulta.Parameters.Add(New SqlParameter("@PUC", SqlDbType.Int))
                    consulta.Parameters("@PUC").Value = objCliente.codigoPuc
                    consulta.Parameters.Add(New SqlParameter("@CIIU", SqlDbType.NVarChar))
                    consulta.Parameters("@CIIU").Value = objCliente.codigoCiiu
                    consulta.Parameters.Add(New SqlParameter("@CUENTAPUC", SqlDbType.NVarChar))
                    consulta.Parameters("@CUENTAPUC").Value = objCliente.cuentaPuc
                    consulta.Parameters.Add(New SqlParameter("@tipoCliente", SqlDbType.NVarChar))
                    consulta.Parameters("@tipoCliente").Value = objCliente.tipoCliente
                    consulta.ExecuteNonQuery()

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

    Public Sub actualizarCliente(objCliente As Cliente)
        Try

            Using consulta As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion

                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_CLIENTE_ACTUALIZAR"
                    consulta.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                    consulta.Parameters("@Id_empresa").Value = objCliente.idEmpresa
                    consulta.Parameters.Add(New SqlParameter("@FormaPago", SqlDbType.NVarChar))
                    consulta.Parameters("@FormaPago").Value = objCliente.formapago
                    consulta.Parameters.Add(New SqlParameter("@Codigo_tipo_regimen", SqlDbType.Int))
                    consulta.Parameters("@Codigo_tipo_regimen").Value = objCliente.codigoregimen
                    consulta.Parameters.Add(New SqlParameter("@Ubicacion", SqlDbType.NVarChar))
                    consulta.Parameters("@Ubicacion").Value = objCliente.ubicacion
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Plazo", SqlDbType.Int))
                    consulta.Parameters("@Codigo_Plazo").Value = objCliente.codigoplazo
                    consulta.Parameters.Add(New SqlParameter("@Codigo_dia_entrega", SqlDbType.Int))
                    consulta.Parameters("@Codigo_dia_entrega").Value = objCliente.codigoentrega
                    consulta.Parameters.Add(New SqlParameter("@Codigo_dia_devolucion", SqlDbType.Int))
                    consulta.Parameters("@Codigo_dia_devolucion").Value = objCliente.codigodevolucion
                    consulta.Parameters.Add(New SqlParameter("@Observaciones", SqlDbType.NVarChar))
                    consulta.Parameters("@Observaciones").Value = objCliente.observacion
                    consulta.Parameters.Add(New SqlParameter("@Suspendido", SqlDbType.Bit))
                    consulta.Parameters("@Suspendido").Value = objCliente.suspendido
                    consulta.Parameters.Add(New SqlParameter("@Bloqueado", SqlDbType.Bit))
                    consulta.Parameters("@Bloqueado").Value = objCliente.bloqueado
                    consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                    consulta.Parameters("@USUARIO").Value = objCliente.usuario
                    consulta.Parameters.Add(New SqlParameter("@IDCLIENTE", SqlDbType.Int))
                    consulta.Parameters("@IDCLIENTE").Value = objCliente.IdCliente
                    consulta.Parameters.Add(New SqlParameter("@PUC", SqlDbType.Int))
                    consulta.Parameters("@PUC").Value = objCliente.codigoPuc
                    consulta.Parameters.Add(New SqlParameter("@CIIU", SqlDbType.NVarChar))
                    consulta.Parameters("@CIIU").Value = objCliente.codigoCiiu
                    consulta.Parameters.Add(New SqlParameter("@CUENTAPUC", SqlDbType.NVarChar))
                    consulta.Parameters("@CUENTAPUC").Value = objCliente.cuentaPuc
                    consulta.Parameters.Add(New SqlParameter("@tipoCliente", SqlDbType.NVarChar))
                    consulta.Parameters("@tipoCliente").Value = objCliente.tipoCliente
                    consulta.ExecuteNonQuery()
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
