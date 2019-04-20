Imports System.Data.SqlClient
Public Class ContratoDAL
    Public Function guardarContrato(objContrato As ContratoEps,
                                       elementoAEliminar As List(Of String)) As String

        Try

            Using dbcommand As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    dbcommand.Connection = FormPrincipal.cnxion
                    dbcommand.Transaction = trnsccion
                    For Each sentencia As String In elementoAEliminar
                        If sentencia <> "" Then
                            dbcommand.CommandText = sentencia
                            dbcommand.ExecuteNonQuery()
                        End If
                    Next
                    dbcommand.Parameters.Clear()
                    dbcommand.CommandType = CommandType.StoredProcedure
                    dbcommand.CommandText = "SP_CONTRATO_EPS_CREAR"
                    dbcommand.Parameters.Add(New SqlParameter("@Id_cliente", SqlDbType.Int))
                    dbcommand.Parameters("@Id_cliente").Value = objContrato.idCliente
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_EPS", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_EPS").Value = objContrato.codigoEPS
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_Lista_Precio", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_Lista_Precio").Value = objContrato.codigoListaPrecio
                    dbcommand.Parameters.Add(New SqlParameter("@pCodigoListaPrecioVentaDirecta", SqlDbType.Int)).Value = objContrato.codigoListaPrecioVentaDirecta
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_manual", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_manual").Value = objContrato.codigoManual
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_HAP", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_HAP").Value = objContrato.codigoHAP
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_HAL", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_HAL").Value = objContrato.codigoHAL
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_HAI", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_HAI").Value = objContrato.codigoHAI
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_Tipo_Contrato", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_Tipo_Contrato").Value = objContrato.codigoTipoContrato
                    dbcommand.Parameters.Add(New SqlParameter("@Cedula_representante", SqlDbType.NVarChar))
                    dbcommand.Parameters("@Cedula_representante").Value = objContrato.cedulaRepresentante
                    dbcommand.Parameters.Add(New SqlParameter("@nombre_representante", SqlDbType.NVarChar))
                    dbcommand.Parameters("@nombre_representante").Value = objContrato.nombreRepresentante
                    dbcommand.Parameters.Add(New SqlParameter("@Num_Contrato", SqlDbType.NVarChar))
                    dbcommand.Parameters("@Num_Contrato").Value = objContrato.numContrato
                    dbcommand.Parameters.Add(New SqlParameter("@Valor_Oxigeno", SqlDbType.Int))
                    dbcommand.Parameters("@Valor_Oxigeno").Value = objContrato.valorOxigeno
                    dbcommand.Parameters.Add(New SqlParameter("@Valor_Evento", SqlDbType.BigInt))
                    dbcommand.Parameters("@Valor_Evento").Value = objContrato.valorEvento
                    dbcommand.Parameters.Add(New SqlParameter("@Fecha_Ini_Evento", SqlDbType.Date))
                    dbcommand.Parameters("@Fecha_Ini_Evento").Value = objContrato.fechaIniEvento
                    dbcommand.Parameters.Add(New SqlParameter("@Fecha_final_Evento", SqlDbType.Date))
                    dbcommand.Parameters("@Fecha_final_Evento").Value = objContrato.fechaFinalEvento
                    dbcommand.Parameters.Add(New SqlParameter("@Valor_Capitado", SqlDbType.BigInt))
                    dbcommand.Parameters("@Valor_Capitado").Value = objContrato.valorCapitado
                    dbcommand.Parameters.Add(New SqlParameter("@Fecha_Ini_capit", SqlDbType.Date))
                    dbcommand.Parameters("@Fecha_Ini_capit").Value = objContrato.fechaIniCapit
                    dbcommand.Parameters.Add(New SqlParameter("@Fecha_final_capit", SqlDbType.Date))
                    dbcommand.Parameters("@Fecha_final_capit").Value = objContrato.fechaIniCapit
                    dbcommand.Parameters.Add(New SqlParameter("@Num_Afiliados", SqlDbType.Int))
                    dbcommand.Parameters("@Num_Afiliados").Value = objContrato.numAfiliados
                    dbcommand.Parameters.Add(New SqlParameter("@estado", SqlDbType.Int))
                    dbcommand.Parameters("@estado").Value = objContrato.estado
                    dbcommand.Parameters.Add(New SqlParameter("@codigoEp", SqlDbType.NVarChar))
                    dbcommand.Parameters("@codigoEp").Value = objContrato.codigoEp
                    dbcommand.Parameters.Add(New SqlParameter("@codigoHabilitacion", SqlDbType.NVarChar))
                    dbcommand.Parameters("@codigoHabilitacion").Value = objContrato.codigoHabilitacion
                    dbcommand.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                    dbcommand.Parameters("@USUARIO").Value = objContrato.Usuario
                    objContrato.codigoContrato = CType(dbcommand.ExecuteScalar, Integer)

                    For indice = 0 To objContrato.dtAreasServicio.Rows.Count - 1
                        If objContrato.dtAreasServicio.Rows(indice).Item(0).ToString <> "" And objContrato.dtAreasServicio.Rows(indice).Item(2).ToString = Constantes.ITEM_NUEVO Then
                            dbcommand.Parameters.Clear()
                            dbcommand.CommandText = "SP_CONTRATO_EPS_CREAR_D"
                            dbcommand.Parameters.Add(New SqlParameter("@Codigo_Contrato", SqlDbType.NVarChar))
                            dbcommand.Parameters("@Codigo_Contrato").Value = objContrato.codigoContrato
                            dbcommand.Parameters.Add(New SqlParameter("@Codigo_Area_Servicio", SqlDbType.Int))
                            dbcommand.Parameters("@Codigo_Area_Servicio").Value = objContrato.dtAreasServicio.Rows(indice).Item(0).ToString
                            dbcommand.ExecuteNonQuery()
                        End If
                    Next
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        objContrato.codigoContrato = ""
                        trnsccion.Rollback()
                        Throw ex
                    End Try
                End Using
            End Using
        Catch ex As Exception
            objContrato.codigoContrato = ""
            Throw ex
        End Try
        Return objContrato.codigoContrato
    End Function
    Public Sub actualizarContrato(objContrato As ContratoEps, elementoAEliminar As List(Of String))

        Try

            Using dbcommand As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    dbcommand.Connection = FormPrincipal.cnxion
                    dbcommand.Transaction = trnsccion
                    For Each sentencia As String In elementoAEliminar
                        If sentencia <> "" Then
                            dbcommand.CommandText = sentencia
                            dbcommand.ExecuteNonQuery()
                        End If
                    Next
                    dbcommand.Parameters.Clear()
                    dbcommand.CommandType = CommandType.StoredProcedure
                    dbcommand.CommandText = "SP_CONTRATO_EPS_ACTUALIZAR"
                    dbcommand.Parameters.Add(New SqlParameter("@Id_cliente", SqlDbType.Int))
                    dbcommand.Parameters("@Id_cliente").Value = objContrato.idCliente
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_EPS", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_EPS").Value = objContrato.codigoEPS
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_Lista_Precio", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_Lista_Precio").Value = objContrato.codigoListaPrecio
                    dbcommand.Parameters.Add(New SqlParameter("@pCodigoListaPrecioVentaDirecta", SqlDbType.Int)).Value = objContrato.codigoListaPrecioVentaDirecta
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_manual", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_manual").Value = objContrato.codigoManual
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_HAP", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_HAP").Value = objContrato.codigoHAP
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_HAL", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_HAL").Value = objContrato.codigoHAL
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_HAI", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_HAI").Value = objContrato.codigoHAI
                    dbcommand.Parameters.Add(New SqlParameter("@Codigo_Tipo_Contrato", SqlDbType.Int))
                    dbcommand.Parameters("@Codigo_Tipo_Contrato").Value = objContrato.codigoTipoContrato
                    dbcommand.Parameters.Add(New SqlParameter("@Cedula_representante", SqlDbType.NVarChar))
                    dbcommand.Parameters("@Cedula_representante").Value = objContrato.cedulaRepresentante
                    dbcommand.Parameters.Add(New SqlParameter("@nombre_representante", SqlDbType.NVarChar))
                    dbcommand.Parameters("@nombre_representante").Value = objContrato.nombreRepresentante
                    dbcommand.Parameters.Add(New SqlParameter("@Num_Contrato", SqlDbType.NVarChar))
                    dbcommand.Parameters("@Num_Contrato").Value = objContrato.numContrato
                    dbcommand.Parameters.Add(New SqlParameter("@Valor_Oxigeno", SqlDbType.Int))
                    dbcommand.Parameters("@Valor_Oxigeno").Value = objContrato.valorOxigeno
                    dbcommand.Parameters.Add(New SqlParameter("@Valor_Evento", SqlDbType.BigInt))
                    dbcommand.Parameters("@Valor_Evento").Value = objContrato.valorEvento
                    dbcommand.Parameters.Add(New SqlParameter("@Fecha_Ini_Evento", SqlDbType.Date))
                    dbcommand.Parameters("@Fecha_Ini_Evento").Value = objContrato.fechaIniEvento
                    dbcommand.Parameters.Add(New SqlParameter("@Fecha_final_Evento", SqlDbType.Date))
                    dbcommand.Parameters("@Fecha_final_Evento").Value = objContrato.fechaFinalEvento
                    dbcommand.Parameters.Add(New SqlParameter("@Valor_Capitado", SqlDbType.BigInt))
                    dbcommand.Parameters("@Valor_Capitado").Value = objContrato.valorCapitado
                    dbcommand.Parameters.Add(New SqlParameter("@Fecha_Ini_capit", SqlDbType.Date))
                    dbcommand.Parameters("@Fecha_Ini_capit").Value = objContrato.fechaIniCapit
                    dbcommand.Parameters.Add(New SqlParameter("@Fecha_final_capit", SqlDbType.Date))
                    dbcommand.Parameters("@Fecha_final_capit").Value = objContrato.fechaFinalCapit
                    dbcommand.Parameters.Add(New SqlParameter("@Num_Afiliados", SqlDbType.Int))
                    dbcommand.Parameters("@Num_Afiliados").Value = objContrato.numAfiliados
                    dbcommand.Parameters.Add(New SqlParameter("@estado", SqlDbType.Int))
                    dbcommand.Parameters("@estado").Value = objContrato.estado
                    dbcommand.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.NVarChar))
                    dbcommand.Parameters("@Usuario").Value = objContrato.Usuario
                    dbcommand.Parameters.Add(New SqlParameter("@codigoHabilitacion", SqlDbType.NVarChar))
                    dbcommand.Parameters("@codigoHabilitacion").Value = objContrato.codigoHabilitacion
                    dbcommand.Parameters.Add(New SqlParameter("@id_Contrato", SqlDbType.Int))
                    dbcommand.Parameters("@id_Contrato").Value = objContrato.codigoContrato
                    dbcommand.ExecuteNonQuery()

                    For i = 0 To objContrato.dtAreasServicio.Rows.Count - 1
                        If objContrato.dtAreasServicio.Rows(i).Item(0).ToString <> "" And objContrato.dtAreasServicio.Rows(i).Item(2).ToString = Constantes.ITEM_NUEVO Then
                            dbcommand.Parameters.Clear()
                            dbcommand.CommandText = "SP_CONTRATO_EPS_CREAR_D"
                            dbcommand.Parameters.Add(New SqlParameter("@Codigo_Contrato", SqlDbType.NVarChar))
                            dbcommand.Parameters("@Codigo_Contrato").Value = objContrato.codigoContrato
                            dbcommand.Parameters.Add(New SqlParameter("@Codigo_Area_Servicio", SqlDbType.Int))
                            dbcommand.Parameters("@Codigo_Area_Servicio").Value = objContrato.dtAreasServicio.Rows(i).Item(0).ToString
                            dbcommand.ExecuteNonQuery()
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
End Class
