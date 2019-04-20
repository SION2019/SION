Imports System.Data.SqlClient
Public Class ProveedorDAL
    Public Sub guardarProveedor(objProveedor As Proveedor)
        Try
            Using consulta As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_PROVEEDOR_CREAR"
                    consulta.Parameters.Add(New SqlParameter("@IDPROVEEDOR", SqlDbType.Int))
                    consulta.Parameters("@IDPROVEEDOR").Value = objProveedor.codigoProveedor
                    consulta.Parameters.Add(New SqlParameter("@FormaPago", SqlDbType.NVarChar))
                    consulta.Parameters("@FormaPago").Value = objProveedor.formaPago
                    consulta.Parameters.Add(New SqlParameter("@Codigo_tipo_regimen", SqlDbType.Int))
                    consulta.Parameters("@Codigo_tipo_regimen").Value = objProveedor.regimen
                    consulta.Parameters.Add(New SqlParameter("@Ubicacion", SqlDbType.NVarChar))
                    consulta.Parameters("@Ubicacion").Value = objProveedor.ubicacion
                    consulta.Parameters.Add(New SqlParameter("@Codigo_dia_entrega", SqlDbType.Int))
                    consulta.Parameters("@Codigo_dia_entrega").Value = objProveedor.codigoDiaEntrega
                    consulta.Parameters.Add(New SqlParameter("@Codigo_dia_vencimiento", SqlDbType.Int))
                    consulta.Parameters("@Codigo_dia_vencimiento").Value = objProveedor.codigoDiaVencimiento
                    consulta.Parameters.Add(New SqlParameter("@Codigo_dia_devolucion", SqlDbType.Int))
                    consulta.Parameters("@Codigo_dia_devolucion").Value = objProveedor.codigoDiaDevolucion
                    consulta.Parameters.Add(New SqlParameter("@Observaciones", SqlDbType.NVarChar))
                    consulta.Parameters("@Observaciones").Value = objProveedor.observacion
                    consulta.Parameters.Add(New SqlParameter("@Suspendido", SqlDbType.Bit))
                    consulta.Parameters("@Suspendido").Value = objProveedor.suspendido
                    consulta.Parameters.Add(New SqlParameter("@Bloqueado", SqlDbType.Bit))
                    consulta.Parameters("@Bloqueado").Value = objProveedor.bloqueado
                    consulta.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                    consulta.Parameters("@Id_empresa").Value = objProveedor.idEmpresa
                    consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@USUARIO").Value = objProveedor.usuario
                    consulta.Parameters.Add(New SqlParameter("@PUC", SqlDbType.Int))
                    consulta.Parameters("@PUC").Value = objProveedor.codigoPuc
                    consulta.Parameters.Add(New SqlParameter("@CIIU", SqlDbType.NVarChar))
                    consulta.Parameters("@CIIU").Value = objProveedor.codigoCiiu
                    consulta.Parameters.Add(New SqlParameter("@CUENTAPUC", SqlDbType.NVarChar))
                    consulta.Parameters("@CUENTAPUC").Value = objProveedor.cuentaPuc
                    consulta.ExecuteNonQuery()


                    consulta.Parameters.Clear()
                    consulta.CommandText = "SP_PROVEEDOR_D_CREAR"
                    consulta.Parameters.Add(New SqlParameter("@IDPROVEEDOR", SqlDbType.Int))
                    consulta.Parameters("@IDPROVEEDOR").Value = objProveedor.codigoProveedor
                    consulta.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                    consulta.Parameters("@Id_empresa").Value = objProveedor.idEmpresa
                    consulta.Parameters.Add(New SqlParameter("@Tipo_pago", SqlDbType.NVarChar))
                    consulta.Parameters("@Tipo_pago").Value = objProveedor.tipoPago
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Banco", SqlDbType.Int))
                    consulta.Parameters("@Codigo_Banco").Value = objProveedor.banco
                    consulta.Parameters.Add(New SqlParameter("@Tipo_Cuenta", SqlDbType.NVarChar))
                    consulta.Parameters("@Tipo_Cuenta").Value = objProveedor.TipocuentaBancaria
                    consulta.Parameters.Add(New SqlParameter("@Numero_Cuenta", SqlDbType.NVarChar))
                    consulta.Parameters("@Numero_Cuenta").Value = objProveedor.numeroCuenta
                    consulta.Parameters.Add(New SqlParameter("@Cedula_cuenta", SqlDbType.NVarChar))
                    consulta.Parameters("@Cedula_cuenta").Value = objProveedor.identificacionProveedor
                    consulta.Parameters.Add(New SqlParameter("@RetenerIVA", SqlDbType.NVarChar))
                    consulta.Parameters("@RetenerIVA").Value = objProveedor.reteIva
                    consulta.Parameters.Add(New SqlParameter("@RetenerFuente", SqlDbType.NVarChar))
                    consulta.Parameters("@RetenerFuente").Value = objProveedor.reteFuente
                    consulta.Parameters.Add(New SqlParameter("@RetenerICA", SqlDbType.NVarChar))
                    consulta.Parameters("@RetenerICA").Value = objProveedor.reteIca
                    consulta.Parameters.Add(New SqlParameter("@CTAreteIVA", SqlDbType.NVarChar))
                    consulta.Parameters("@CTAreteIVA").Value = objProveedor.ctaReteIva
                    consulta.Parameters.Add(New SqlParameter("@CTAreteICA", SqlDbType.NVarChar))
                    consulta.Parameters("@CTAreteICA").Value = objProveedor.ctaReteIca
                    consulta.Parameters.Add(New SqlParameter("@CTAreteFuente", SqlDbType.NVarChar))
                    consulta.Parameters("@CTAreteFuente").Value = objProveedor.ctaReteFuente
                    consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@USUARIO").Value = objProveedor.usuario
                    consulta.ExecuteNonQuery()


                    For i = 0 To objProveedor.dtPlazo.Rows.Count - 1
                        If objProveedor.dtPlazo.Rows(i).Item("codplazo").ToString <> "" Then
                            consulta.Parameters.Clear()
                            consulta.CommandText = "SP_PROVEEDOR_DETALLE_DESCUENTO_PLAZO_CREAR"
                            consulta.Parameters.Add(New SqlParameter("@IDPROVEEDOR", SqlDbType.Int))
                            consulta.Parameters("@IDPROVEEDOR").Value = objProveedor.codigoProveedor
                            consulta.Parameters.Add(New SqlParameter("@Codigo_EP", SqlDbType.Int))
                            consulta.Parameters("@Codigo_EP").Value = objProveedor.codigoEp
                            consulta.Parameters.Add(New SqlParameter("@Codigo_plazo", SqlDbType.NVarChar))
                            consulta.Parameters("@Codigo_plazo").Value = objProveedor.dtPlazo.Rows(i).Item("codplazo").ToString
                            consulta.Parameters.Add(New SqlParameter("@Descuento", SqlDbType.Decimal))
                            consulta.Parameters("@Descuento").Value = objProveedor.dtPlazo.Rows(i).Item("descuento")
                            consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                            consulta.Parameters("@USUARIO").Value = objProveedor.usuario
                            consulta.ExecuteNonQuery()
                        End If
                    Next

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
    Public Sub actualizarProveedor(objProveedor As Proveedor)


        Try

            Using consulta As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion

                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_PROVEEDOR_ACTUALIZAR"
                    consulta.Parameters.Add(New SqlParameter("@FormaPago", SqlDbType.NVarChar))
                    consulta.Parameters("@FormaPago").Value = objProveedor.formaPago
                    consulta.Parameters.Add(New SqlParameter("@Codigo_tipo_regimen", SqlDbType.Int))
                    consulta.Parameters("@Codigo_tipo_regimen").Value = objProveedor.regimen
                    consulta.Parameters.Add(New SqlParameter("@Ubicacion", SqlDbType.NVarChar))
                    consulta.Parameters("@Ubicacion").Value = objProveedor.ubicacion
                    consulta.Parameters.Add(New SqlParameter("@Codigo_dia_entrega", SqlDbType.Int))
                    consulta.Parameters("@Codigo_dia_entrega").Value = objProveedor.codigoDiaEntrega
                    consulta.Parameters.Add(New SqlParameter("@Codigo_dia_vencimiento", SqlDbType.Int))
                    consulta.Parameters("@Codigo_dia_vencimiento").Value = objProveedor.codigoDiaVencimiento
                    consulta.Parameters.Add(New SqlParameter("@Codigo_dia_devolucion", SqlDbType.Int))
                    consulta.Parameters("@Codigo_dia_devolucion").Value = objProveedor.codigoDiaDevolucion
                    consulta.Parameters.Add(New SqlParameter("@Observaciones", SqlDbType.NVarChar))
                    consulta.Parameters("@Observaciones").Value = objProveedor.observacion
                    consulta.Parameters.Add(New SqlParameter("@Suspendido", SqlDbType.Bit))
                    consulta.Parameters("@Suspendido").Value = objProveedor.suspendido
                    consulta.Parameters.Add(New SqlParameter("@Bloqueado", SqlDbType.Bit))
                    consulta.Parameters("@Bloqueado").Value = objProveedor.bloqueado
                    consulta.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                    consulta.Parameters("@Id_empresa").Value = objProveedor.idEmpresa
                    consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                    consulta.Parameters("@USUARIO").Value = objProveedor.usuario
                    consulta.Parameters.Add(New SqlParameter("@IDPROVEEDOR", SqlDbType.Int))
                    consulta.Parameters("@IDPROVEEDOR").Value = objProveedor.codigoProveedor
                    consulta.Parameters.Add(New SqlParameter("@PUC", SqlDbType.Int))
                    consulta.Parameters("@PUC").Value = objProveedor.codigoPuc
                    consulta.Parameters.Add(New SqlParameter("@CIIU", SqlDbType.NVarChar))
                    consulta.Parameters("@CIIU").Value = objProveedor.codigoCiiu
                    consulta.Parameters.Add(New SqlParameter("@CUENTAPUC", SqlDbType.NVarChar))
                    consulta.Parameters("@CUENTAPUC").Value = objProveedor.cuentaPuc
                    consulta.ExecuteNonQuery()


                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_PROVEEDOR_D_ACTUALIZAR"
                    consulta.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                    consulta.Parameters("@Id_empresa").Value = objProveedor.idEmpresa
                    consulta.Parameters.Add(New SqlParameter("@Tipo_pago", SqlDbType.NVarChar))
                    consulta.Parameters("@Tipo_pago").Value = objProveedor.tipoPago
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Banco", SqlDbType.Int))
                    consulta.Parameters("@Codigo_Banco").Value = objProveedor.banco
                    consulta.Parameters.Add(New SqlParameter("@Tipo_Cuenta", SqlDbType.NVarChar))
                    consulta.Parameters("@Tipo_Cuenta").Value = objProveedor.TipocuentaBancaria
                    consulta.Parameters.Add(New SqlParameter("@Numero_Cuenta", SqlDbType.NVarChar))
                    consulta.Parameters("@Numero_Cuenta").Value = objProveedor.numeroCuenta
                    consulta.Parameters.Add(New SqlParameter("@Cedula_cuenta", SqlDbType.NVarChar))
                    consulta.Parameters("@Cedula_cuenta").Value = objProveedor.identificacionProveedor
                    consulta.Parameters.Add(New SqlParameter("@RetenerIVA", SqlDbType.NVarChar))
                    consulta.Parameters("@RetenerIVA").Value = objProveedor.reteIva
                    consulta.Parameters.Add(New SqlParameter("@RetenerFuente", SqlDbType.NVarChar))
                    consulta.Parameters("@RetenerFuente").Value = objProveedor.reteFuente
                    consulta.Parameters.Add(New SqlParameter("@RetenerICA", SqlDbType.NVarChar))
                    consulta.Parameters("@RetenerICA").Value = objProveedor.reteIca
                    consulta.Parameters.Add(New SqlParameter("@CTAreteIVA", SqlDbType.NVarChar))
                    consulta.Parameters("@CTAreteIVA").Value = objProveedor.ctaReteIva
                    consulta.Parameters.Add(New SqlParameter("@CTAreteICA", SqlDbType.NVarChar))
                    consulta.Parameters("@CTAreteICA").Value = objProveedor.ctaReteIca
                    consulta.Parameters.Add(New SqlParameter("@CTAreteFuente", SqlDbType.NVarChar))
                    consulta.Parameters("@CTAreteFuente").Value = objProveedor.ctaReteFuente
                    consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@USUARIO").Value = objProveedor.usuario
                    consulta.Parameters.Add(New SqlParameter("@IDPROVEEDOR", SqlDbType.Int))
                    consulta.Parameters("@IDPROVEEDOR").Value = objProveedor.codigoProveedor
                    consulta.ExecuteNonQuery()


                    For i = 0 To objProveedor.dtPlazo.Rows.Count - 1
                        If objProveedor.dtPlazo.Rows(i).Item("codplazo").ToString <> "" Then
                            consulta.Parameters.Clear()
                            consulta.CommandText = "SP_PROVEEDOR_DETALLE_DESCUENTO_PLAZO_CREAR"
                            consulta.Parameters.Add(New SqlParameter("@IDPROVEEDOR", SqlDbType.Int))
                            consulta.Parameters("@IDPROVEEDOR").Value = objProveedor.codigoProveedor
                            consulta.Parameters.Add(New SqlParameter("@Codigo_EP", SqlDbType.Int))
                            consulta.Parameters("@Codigo_EP").Value = objProveedor.codigoEp
                            consulta.Parameters.Add(New SqlParameter("@Codigo_plazo", SqlDbType.NVarChar))
                            consulta.Parameters("@Codigo_plazo").Value = objProveedor.dtPlazo.Rows(i).Item("codplazo").ToString
                            consulta.Parameters.Add(New SqlParameter("@Descuento", SqlDbType.Decimal))
                            consulta.Parameters("@Descuento").Value = objProveedor.dtPlazo.Rows(i).Item("descuento")
                            consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                            consulta.Parameters("@USUARIO").Value = objProveedor.usuario
                            consulta.ExecuteNonQuery()
                        End If
                    Next

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
    Public Sub proveedorAnular(objProveedor As Proveedor)
        Try
            Using consulta As New SqlCommand("SP_PROVEEDOR_ANULAR")
                consulta.Parameters.Clear()
                consulta.CommandType = CommandType.StoredProcedure
                consulta.Connection = FormPrincipal.cnxion
                consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar))
                consulta.Parameters("@CODIGO").Value = objProveedor.codigoProveedor
                consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int))
                consulta.Parameters("@Usuario").Value = objProveedor.usuario
                consulta.Parameters.Add(New SqlParameter("@id_empresa", SqlDbType.Int))
                consulta.Parameters("@id_empresa").Value = objProveedor.idEmpresa
                consulta.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
