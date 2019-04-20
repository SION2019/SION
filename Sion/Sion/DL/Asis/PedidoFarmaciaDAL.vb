Public Class PedidoFarmaciaDAL
    Public Sub guardarPedido(ByRef objPedidoFarmacia As PedidoFarmacia,
                             ByVal usuario As Integer,
                             ByVal codigoEP As Integer
                             )
        Try
            Using sentencia = New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    sentencia.Connection = FormPrincipal.cnxion
                    sentencia.Transaction = trnsccion
                    sentencia.CommandType = CommandType.StoredProcedure

                    sentencia.Parameters.Clear()
                    sentencia.CommandText = "[SP_PEDIDO_FARMACIA_CREAR]"
                    sentencia.Parameters.Add(New SqlParameter("@Codigo_EP", SqlDbType.Int)).Value = codigoEP
                    sentencia.Parameters.Add(New SqlParameter("@CodigoOrden", SqlDbType.Int)).Value = objPedidoFarmacia.codigoOrden
                    sentencia.Parameters.Add(New SqlParameter("@CodigoEnfermeria", SqlDbType.Int)).Value = objPedidoFarmacia.codigoOrdenEnfermeria
                    sentencia.Parameters.Add(New SqlParameter("@CodigoFisioterapia", SqlDbType.Int)).Value = objPedidoFarmacia.codigoOrdenFisioterapia
                    sentencia.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int)).Value = usuario
                    sentencia.Parameters.Add(New SqlParameter("@volumenTotal", SqlDbType.Float)).Value = objPedidoFarmacia.nutVolumenesTotales
                    sentencia.Parameters.Add(New SqlParameter("@razon", SqlDbType.Float)).Value = objPedidoFarmacia.nutRazon
                    sentencia.Parameters.Add(New SqlParameter("@TablaMedicamentos", SqlDbType.Structured)).Value = objPedidoFarmacia.tablaMedicamentos
                    sentencia.Parameters.Add(New SqlParameter("@TablaImpregnacion", SqlDbType.Structured)).Value = objPedidoFarmacia.tablaImpregnaciones
                    sentencia.Parameters.Add(New SqlParameter("@TablaNutricion", SqlDbType.Structured)).Value = objPedidoFarmacia.tablaNutricion
                    sentencia.Parameters.Add(New SqlParameter("@TablaInsumosEnfermeria", SqlDbType.Structured)).Value = objPedidoFarmacia.tablaInsumosEnfermeria
                    sentencia.Parameters.Add(New SqlParameter("@TablaInsumosFisioterapia", SqlDbType.Structured)).Value = objPedidoFarmacia.tablaInsumosFisioterapia
                    objPedidoFarmacia.codigoPedidoFarmacia = CType(sentencia.ExecuteScalar, String)

                    For indiceFila = 0 To objPedidoFarmacia.tablaInfusiones.Rows.Count - 1
                        sentencia.Parameters.Clear()
                        sentencia.CommandText = "[SP_PEDIDO_FARMACIA_CREAR_INFUSION]"
                        sentencia.Parameters.Add(New SqlParameter("@CodigoPedidoFarmacia", SqlDbType.Int)).Value = objPedidoFarmacia.codigoPedidoFarmacia
                        sentencia.Parameters.Add(New SqlParameter("@consecutivo_infusion", SqlDbType.Int)).Value = objPedidoFarmacia.tablaInfusiones.Rows(indiceFila).Item("Consecutivo_Infusion")
                        sentencia.Parameters.Add(New SqlParameter("@CodigoInterno", SqlDbType.Int)).Value = objPedidoFarmacia.tablaInfusiones.Rows(indiceFila).Item("Codigo_Interno")
                        sentencia.Parameters.Add(New SqlParameter("@dosis", SqlDbType.Float)).Value = objPedidoFarmacia.tablaInfusiones.Rows(indiceFila).Item("Dosis")
                        sentencia.Parameters.Add(New SqlParameter("@velocidad", SqlDbType.Float)).Value = objPedidoFarmacia.tablaInfusiones.Rows(indiceFila).Item("Velocidad")
                        sentencia.Parameters.Add(New SqlParameter("@Producto_Disolvente", SqlDbType.Int)).Value = objPedidoFarmacia.tablaInfusiones.Rows(indiceFila).Item("Producto_Disolvente")
                        sentencia.ExecuteNonQuery()

                        If objPedidoFarmacia.IsTablaExistente(objPedidoFarmacia.nombrarMezclaInfusion(objPedidoFarmacia.tablaInfusiones.Rows(indiceFila).Item("Consecutivo_Infusion"),
                                                                                                      objPedidoFarmacia.tablaInfusiones.Rows(indiceFila).Item("Codigo_Interno"))) Then
                            sentencia.Parameters.Clear()
                            sentencia.CommandText = "[SP_PEDIDO_FARMACIA_CREAR_INFUSION_MEZCLA]"
                            sentencia.Parameters.Add(New SqlParameter("@CodigoPedidoFarmacia", SqlDbType.Int)).Value = objPedidoFarmacia.codigoPedidoFarmacia
                            sentencia.Parameters.Add(New SqlParameter("@TablaMezcla", SqlDbType.Structured)).Value = objPedidoFarmacia.dtTablaMezclas.Tables(objPedidoFarmacia.nombrarMezclaInfusion(objPedidoFarmacia.tablaInfusiones.Rows(indiceFila).Item("Consecutivo_Infusion"),
                                                                                                                                                                                                     objPedidoFarmacia.tablaInfusiones.Rows(indiceFila).Item("Codigo_Interno")))
                            sentencia.ExecuteNonQuery()
                        End If

                    Next

                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub anularPedido(ByRef objPedidoFarmacia As PedidoFarmacia,
                            ByVal usuario As Integer)
        Try
            Using sentencia = New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    sentencia.Connection = FormPrincipal.cnxion
                    sentencia.Transaction = trnsccion
                    sentencia.CommandType = CommandType.StoredProcedure

                    sentencia.Parameters.Clear()
                    sentencia.CommandText = "[SP_PEDIDO_FARMACIA_ANULAR]"
                    sentencia.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.Int)).Value = objPedidoFarmacia.codigoPedidoFarmacia
                    sentencia.Parameters.Add(New SqlParameter("@CodigoOrden", SqlDbType.Int)).Value = objPedidoFarmacia.codigoOrden
                    sentencia.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = usuario
                    sentencia.ExecuteNonQuery()
                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Function verificarPedidoCargado(ByRef objPedidoFarmacia As PedidoFarmacia) As Boolean
        Dim respuesta As Boolean = False
        Dim params As New List(Of String)
        params.Add(objPedidoFarmacia.codigoPedidoFarmacia)
        Dim fila As DataRow

        fila = General.cargarItem(Consultas.PEDIDO_FARMACIA_VERIFICAR_CARGADO, params)
        If Not IsNothing(fila) Then
            respuesta = True
        End If
        Return respuesta
    End Function

    Public Function verificarEstadoActualPaciente(ByRef nRegistro As Integer,
                                                 ByVal codigoEstadonIniciado As Integer) As Boolean
        Dim respuesta As Boolean = False
        Dim params As New List(Of String)
        params.Add(nRegistro)
        params.Add(codigoEstadonIniciado)
        Dim parametrosArmados As String = Funciones.getParametros(params)
        Try
            Using comandoSelect As New SqlCommand("[SP_VERIFICAR_ESTADO_ACTUAL_PACIENTE]" & parametrosArmados, FormPrincipal.cnxion)
                Using comandoLector = comandoSelect.ExecuteReader
                    If comandoLector.HasRows Then
                        comandoLector.Read()
                        respuesta = CBool(comandoLector.Item("respuesta"))
                    End If
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return respuesta
    End Function
End Class
