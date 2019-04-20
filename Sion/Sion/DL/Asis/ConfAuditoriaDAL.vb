Imports System.Data.SqlClient
Public Class ConfAuditoriaDAL
    Public Function guardarConfiguracion(objConfigAuditoria As ConfiguracionAuditoria) As String

        Try

            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion

                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = "SP_CONFIGURACION_AUDITORIA_CREAR"
                dbCommand.Parameters.Add(New SqlParameter("@Id_Config", SqlDbType.NVarChar))
                dbCommand.Parameters("@Id_Config").Value = objConfigAuditoria.idConfig
                dbCommand.Parameters.Add(New SqlParameter("@Codigo_Config", SqlDbType.NVarChar))
                dbCommand.Parameters("@Codigo_Config").Value = objConfigAuditoria.codigoProcedimiento
                dbCommand.Parameters.Add(New SqlParameter("@tipo", SqlDbType.NVarChar))
                dbCommand.Parameters("@tipo").Value = objConfigAuditoria.tipo
                dbCommand.Parameters.Add(New SqlParameter("@Id_tercero_EPS", SqlDbType.Int))
                dbCommand.Parameters("@Id_tercero_EPS").Value = objConfigAuditoria.idTerceroEps
                dbCommand.Parameters.Add(New SqlParameter("@Codigo_Area_Servicio", SqlDbType.Int))
                dbCommand.Parameters("@Codigo_Area_Servicio").Value = objConfigAuditoria.codigoAreaServicio
                dbCommand.Parameters.Add(New SqlParameter("@Historia", SqlDbType.Bit))
                dbCommand.Parameters("@Historia").Value = objConfigAuditoria.historia
                dbCommand.Parameters.Add(New SqlParameter("@Auditoria_M", SqlDbType.Bit))
                dbCommand.Parameters("@Auditoria_M").Value = objConfigAuditoria.auditoriaM
                dbCommand.Parameters.Add(New SqlParameter("@Auditoria_F", SqlDbType.Bit))
                dbCommand.Parameters("@Auditoria_F").Value = objConfigAuditoria.auditoriaF
                dbCommand.Parameters.Add(New SqlParameter("@Paquete", SqlDbType.Bit))
                dbCommand.Parameters("@Paquete").Value = objConfigAuditoria.paquete
                dbCommand.Parameters.Add(New SqlParameter("@Una_Vez", SqlDbType.Bit))
                dbCommand.Parameters("@Una_Vez").Value = objConfigAuditoria.unaVez
                dbCommand.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int))
                dbCommand.Parameters("@Usuario_Creacion").Value = objConfigAuditoria.usuario
                dbCommand.Parameters.Add(New SqlParameter("@detalleInsumos", SqlDbType.Structured))
                dbCommand.Parameters("@detalleInsumos").Value = objConfigAuditoria.dtInsumo
                dbCommand.Parameters.Add(New SqlParameter("@detalleParaclinicos", SqlDbType.Structured))
                dbCommand.Parameters("@detalleParaclinicos").Value = objConfigAuditoria.dtParaclinico
                dbCommand.Parameters.Add(New SqlParameter("@detalleProcedimientos", SqlDbType.Structured))
                dbCommand.Parameters("@detalleProcedimientos").Value = objConfigAuditoria.dtProcedimiento
                objConfigAuditoria.idConfig = CType(dbCommand.ExecuteScalar, Integer)

                '    For i = 0 To objConfigAuditoria.dtProcedimiento.Rows.Count - 2

                '        If objConfigAuditoria.dtProcedimiento.Rows(i).Item(3).ToString = Constantes.ITEM_NUEVO Then
                '            dbCommand.Parameters.Clear()
                '            dbCommand.CommandType = CommandType.StoredProcedure
                '            dbCommand.CommandText = "SP_CONFIGURACION_PROCEDIMIENTOS_AUDITORIA_CREAR"
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Config", SqlDbType.NVarChar))
                '            dbCommand.Parameters("@Codigo_Config").Value = objConfigAuditoria.idConfig
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Procedimiento", SqlDbType.NVarChar))
                '            dbCommand.Parameters("@Codigo_Procedimiento").Value = objConfigAuditoria.dtProcedimiento.Rows(i).Item(0).ToString
                '            dbCommand.Parameters.Add(New SqlParameter("@Cantidad", SqlDbType.Int))
                '            dbCommand.Parameters("@Cantidad").Value = objConfigAuditoria.dtProcedimiento.Rows(i).Item(2).ToString
                '            dbCommand.Parameters.Add(New SqlParameter("@Id_tercero_EPS", SqlDbType.Int))
                '            dbCommand.Parameters("@Id_tercero_EPS").Value = objConfigAuditoria.idTerceroEps
                '            dbCommand.ExecuteNonQuery()
                '        Else
                '            dbCommand.Parameters.Clear()
                '            dbCommand.CommandType = CommandType.StoredProcedure
                '            dbCommand.CommandText = "SP_CONFIGURACION_AUDITORIA_PROCEDIMIENTOS_ACTUALIZAR"
                '            dbCommand.Parameters.Add(New SqlParameter("@Cantidad", SqlDbType.Int))
                '            dbCommand.Parameters("@Cantidad").Value = objConfigAuditoria.dtProcedimiento.Rows(i).Item(2).ToString
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Config", SqlDbType.NVarChar))
                '            dbCommand.Parameters("@Codigo_Config").Value = objConfigAuditoria.idConfig
                '            dbCommand.Parameters.Add(New SqlParameter("@Id_tercero_EPS", SqlDbType.Int))
                '            dbCommand.Parameters("@Id_tercero_EPS").Value = objConfigAuditoria.idTerceroEps
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Area_Servicio", SqlDbType.Int))
                '            dbCommand.Parameters("@Codigo_Area_Servicio").Value = objConfigAuditoria.codigoAreaServicio
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Procedimiento", SqlDbType.NVarChar))
                '            dbCommand.Parameters("@Codigo_Procedimiento").Value = objConfigAuditoria.dtProcedimiento.Rows(i).Item(0).ToString
                '            dbCommand.ExecuteNonQuery()
                '        End If
                '    Next

                '    For i = 0 To objConfigAuditoria.dtParaclinico.Rows.Count - 2

                '        If objConfigAuditoria.dtParaclinico.Rows(i).Item(3).ToString = Constantes.ITEM_NUEVO Then
                '            dbCommand.Parameters.Clear()
                '            dbCommand.CommandType = CommandType.StoredProcedure
                '            dbCommand.CommandText = "SP_CONFIGURACION_AUDITORIA_PARACLINICOS_CREAR"
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Config", SqlDbType.NVarChar))
                '            dbCommand.Parameters("@Codigo_Config").Value = objConfigAuditoria.idConfig
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Procedimiento", SqlDbType.NVarChar))
                '            dbCommand.Parameters("@Codigo_Procedimiento").Value = objConfigAuditoria.dtParaclinico.Rows(i).Item(0).ToString
                '            dbCommand.Parameters.Add(New SqlParameter("@Cantidad", SqlDbType.Int))
                '            dbCommand.Parameters("@Cantidad").Value = objConfigAuditoria.dtParaclinico.Rows(i).Item(2).ToString
                '            dbCommand.Parameters.Add(New SqlParameter("@Id_tercero_EPS", SqlDbType.Int))
                '            dbCommand.Parameters("@Id_tercero_EPS").Value = objConfigAuditoria.idTerceroEps
                '            dbCommand.ExecuteNonQuery()
                '        Else
                '            dbCommand.Parameters.Clear()
                '            dbCommand.CommandType = CommandType.StoredProcedure
                '            dbCommand.CommandText = "SP_CONFIGURACION_AUDITORIA_PARACLINICOS_ACTUALIZAR"
                '            dbCommand.Parameters.Add(New SqlParameter("@Cantidad", SqlDbType.Int))
                '            dbCommand.Parameters("@Cantidad").Value = objConfigAuditoria.dtParaclinico.Rows(i).Item(2).ToString
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Config", SqlDbType.NVarChar))
                '            dbCommand.Parameters("@Codigo_Config").Value = objConfigAuditoria.idConfig
                '            dbCommand.Parameters.Add(New SqlParameter("@Id_tercero_EPS", SqlDbType.Int))
                '            dbCommand.Parameters("@Id_tercero_EPS").Value = objConfigAuditoria.idTerceroEps
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Area_Servicio", SqlDbType.Int))
                '            dbCommand.Parameters("@Codigo_Area_Servicio").Value = objConfigAuditoria.codigoAreaServicio
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Procedimiento", SqlDbType.NVarChar))
                '            dbCommand.Parameters("@Codigo_Procedimiento").Value = objConfigAuditoria.dtParaclinico.Rows(i).Item(0).ToString
                '            dbCommand.ExecuteNonQuery()
                '        End If
                '    Next

                '    For i = 0 To objConfigAuditoria.dtInsumo.Rows.Count - 2

                '        If objConfigAuditoria.dtInsumo.Rows(i).Item(3).ToString = Constantes.ITEM_NUEVO Then
                '            dbCommand.Parameters.Clear()
                '            dbCommand.CommandType = CommandType.StoredProcedure
                '            dbCommand.CommandText = "SP_CONFIGURACION_AUDITORIA_INSUMOS_CREAR"
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Config", SqlDbType.NVarChar))
                '            dbCommand.Parameters("@Codigo_Config").Value = objConfigAuditoria.idConfig
                '            dbCommand.Parameters.Add(New SqlParameter("@Id_tercero_EPS", SqlDbType.Int))
                '            dbCommand.Parameters("@Id_tercero_EPS").Value = objConfigAuditoria.idTerceroEps
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Interno", SqlDbType.Int))
                '            dbCommand.Parameters("@Codigo_Interno").Value = objConfigAuditoria.dtInsumo.Rows(i).Item(0).ToString
                '            dbCommand.Parameters.Add(New SqlParameter("@Cantidad", SqlDbType.Int))
                '            dbCommand.Parameters("@Cantidad").Value = objConfigAuditoria.dtInsumo.Rows(i).Item(2).ToString
                '            dbCommand.ExecuteNonQuery()
                '        Else
                '            dbCommand.Parameters.Clear()
                '            dbCommand.CommandType = CommandType.StoredProcedure
                '            dbCommand.CommandText = "SP_CONFIGURACION_AUDITORIA_INSUMOS_ACTUALIZAR"
                '            dbCommand.Parameters.Add(New SqlParameter("@Cantidad", SqlDbType.Int))
                '            dbCommand.Parameters("@Cantidad").Value = objConfigAuditoria.dtInsumo.Rows(i).Item(2).ToString
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Config", SqlDbType.NVarChar))
                '            dbCommand.Parameters("@Codigo_Config").Value = objConfigAuditoria.idConfig
                '            dbCommand.Parameters.Add(New SqlParameter("@Id_tercero_EPS", SqlDbType.Int))
                '            dbCommand.Parameters("@Id_tercero_EPS").Value = objConfigAuditoria.idTerceroEps
                '            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Interno", SqlDbType.Int))
                '            dbCommand.Parameters("@Codigo_Interno").Value = objConfigAuditoria.dtInsumo.Rows(i).Item(0).ToString
                '            dbCommand.ExecuteNonQuery()
                '        End If

                '    Next
            End Using

        Catch ex As SqlException
            If ex.Errors(0).Number = Constantes.ERROR_LLAVE_DUPLICADA Then
                Throw New Exception("¡Estos procedimientos ya fueron cofigurados para esta eps, este procedimiento y esta area de servicio!")
            Else
                Throw ex
            End If
        End Try
        Return objConfigAuditoria.idConfig
    End Function

End Class
