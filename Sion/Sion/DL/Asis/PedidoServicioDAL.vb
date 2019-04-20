Imports System.Data.SqlClient
Public Class Pedido_Servicio_C
    Public Sub guardar(obj As Pedido_Servicio)


        Try
            Using sentencia = New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    sentencia.Connection = FormPrincipal.cnxion
                    sentencia.Transaction = trnsccion
                    sentencia.CommandType = CommandType.StoredProcedure

                    sentencia.Parameters.Clear()
                    sentencia.CommandText = "[SP_PEDIDO_SERVICIO_CREAR]"
                    sentencia.Parameters.Add(New SqlParameter("@pArea", SqlDbType.Int)).Value = obj.areaServicio
                    sentencia.Parameters.Add(New SqlParameter("@pFechaPedido", SqlDbType.DateTime)).Value = obj.fechaPedido
                    sentencia.Parameters.Add(New SqlParameter("@pCodigo_EP", SqlDbType.Int)).Value = SesionActual.codigoEP
                    sentencia.Parameters.Add(New SqlParameter("@pEstado", SqlDbType.NVarChar)).Value = obj.estado
                    sentencia.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = SesionActual.idUsuario
                    sentencia.Parameters.Add(New SqlParameter("@tabla", SqlDbType.Structured)).Value = obj.detallePedido
                    obj.codigoPedido = CType(sentencia.ExecuteScalar, String)

                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Sub actualizar(obj As Pedido_Servicio)

        Try
            Using sentencia = New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    sentencia.Connection = FormPrincipal.cnxion
                    sentencia.Transaction = trnsccion
                    sentencia.CommandType = CommandType.StoredProcedure

                    sentencia.Parameters.Clear()
                    sentencia.CommandText = "[SP_PEDIDO_SERVICIO_MODIFICAR]"
                    sentencia.Parameters.Add(New SqlParameter("@pCodigo", SqlDbType.Int)).Value = obj.codigoPedido
                    sentencia.Parameters.Add(New SqlParameter("@pArea", SqlDbType.Int)).Value = obj.areaServicio
                    sentencia.Parameters.Add(New SqlParameter("@pFechaPedido", SqlDbType.DateTime)).Value = obj.fechaPedido
                    sentencia.Parameters.Add(New SqlParameter("@pCodigo_EP", SqlDbType.Int)).Value = SesionActual.codigoEP
                    sentencia.Parameters.Add(New SqlParameter("@pEstado", SqlDbType.NVarChar)).Value = obj.estado
                    sentencia.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = SesionActual.idUsuario
                    sentencia.Parameters.Add(New SqlParameter("@tabla", SqlDbType.Structured)).Value = obj.detallePedido
                    sentencia.ExecuteNonQuery()


                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Sub anular(obj As Pedido_Servicio)

        Try
            Using sentencia = New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    sentencia.Connection = FormPrincipal.cnxion
                    sentencia.Transaction = trnsccion
                    sentencia.CommandType = CommandType.StoredProcedure

                    sentencia.Parameters.Clear()
                    sentencia.CommandText = "[SP_PEDIDO_SERVICIO_ANULAR]"
                    sentencia.Parameters.Add(New SqlParameter("@pcodigo", SqlDbType.Int)).Value = obj.codigoPedido
                    sentencia.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = SesionActual.idUsuario
                    sentencia.ExecuteNonQuery()

                    trnsccion.Commit()
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Sub cargarDatos(ByRef obj As Pedido_Servicio, ByVal codigo As String)

        obj.detallePedido.Clear()

        Dim params As New List(Of String)
        params.Add(codigo)

        Using cmd As New SqlCommand(Consultas.PEDIDO_SERVICIO_CARGAR_INFORMACION & Funciones.getParametros(params), FormPrincipal.cnxion)
            Using lector = cmd.ExecuteReader
                If lector.HasRows Then
                    lector.Read()
                    obj.codigoPedido = lector.Item("codigo_Pedido")
                    obj.areaServicio = lector.Item("Area_Servicio")
                    obj.fechaPedido = lector.Item("Fecha_Pedido")
                    obj.estado = lector.Item("Estado")
                End If
            End Using
        End Using

        Using cmd As New SqlDataAdapter(Consultas.PEDIDO_SERVICIO_CARGAR_INFORMACION_DETALLE & Funciones.getParametros(params), FormPrincipal.cnxion)
            cmd.Fill(obj.detallePedido)
        End Using

    End Sub
End Class
