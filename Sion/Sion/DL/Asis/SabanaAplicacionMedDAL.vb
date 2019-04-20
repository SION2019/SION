Public Class SabanaAplicacionMedDAL
    Public Shared Sub guardarSabanaHC(objSabanaAplicacion As SabanaAplicacionMed)
        Try
            Using comando = New SqlCommand()
                'Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                comando.Connection = FormPrincipal.cnxion
                    comando.CommandType = CommandType.StoredProcedure
                'comando.Transaction = trnsccion
                comando.CommandText = "[SP_SABANA_SIGNO_VITAL_CREAR]"
                    comando.Parameters.Add(New SqlParameter("@Fecha_Real", SqlDbType.DateTime)).Value = objSabanaAplicacion.fechaReal
                    comando.Parameters.Add(New SqlParameter("@Peso", SqlDbType.Float)).Value = objSabanaAplicacion.pesoActual
                    comando.Parameters.Add(New SqlParameter("@Gasto_U", SqlDbType.Float)).Value = objSabanaAplicacion.gastoUrinario
                    comando.Parameters.Add(New SqlParameter("@Ingreso_total", SqlDbType.Float)).Value = objSabanaAplicacion.ingresoTotal
                    comando.Parameters.Add(New SqlParameter("@Egreso_total", SqlDbType.Float)).Value = objSabanaAplicacion.egresoTotal
                    comando.Parameters.Add(New SqlParameter("@balance_total", SqlDbType.Float)).Value = objSabanaAplicacion.balanceActual
                    comando.Parameters.Add(New SqlParameter("@Balance_Acumulado", SqlDbType.Float)).Value = objSabanaAplicacion.balanceAcumulado
                    comando.Parameters.Add(New SqlParameter("@Regitro", SqlDbType.Int)).Value = objSabanaAplicacion.nRegistro
                    comando.Parameters.Add(New SqlParameter("@Usuario_Real", SqlDbType.Int)).Value = objSabanaAplicacion.usuarioReal
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objSabanaAplicacion.usuario
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHorasabana
                    objSabanaAplicacion.codigoSabana = CType(comando.ExecuteScalar, String)
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_SILVERMAN_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaSilverman
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_VALORACION_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaValoracion
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_INGRESO_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaIngreso
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_EGRESO_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaEgreso
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_BALANCE_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaBalance
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_APLICACION_MED_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaMedicamento
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaNPT
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_APLICACION_GOTEO_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHorasabanaGoteo
                    comando.ExecuteNonQuery()
                '    Try
                '        trnsccion.Commit()
                '    Catch ex As Exception
                '        Try
                '            trnsccion.Rollback()
                '            Throw ex
                '        Catch ex1 As Exception
                '            Throw ex
                '        End Try
                '    End Try
                'End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarSabanaHC(objSabanaAplicacion As SabanaAplicacionMed)
        Try
            Using comando = New SqlCommand()
                'Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                comando.Connection = FormPrincipal.cnxion
                    comando.CommandType = CommandType.StoredProcedure
                '  comando.Transaction = trnsccion
                comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_SIGNO_VITAL_ACTUALIZAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@Peso", SqlDbType.Float)).Value = objSabanaAplicacion.pesoActual
                    comando.Parameters.Add(New SqlParameter("@Gasto_U", SqlDbType.Float)).Value = objSabanaAplicacion.gastoUrinario
                    comando.Parameters.Add(New SqlParameter("@Ingreso_total", SqlDbType.Float)).Value = objSabanaAplicacion.ingresoTotal
                    comando.Parameters.Add(New SqlParameter("@Egreso_total", SqlDbType.Float)).Value = objSabanaAplicacion.egresoTotal
                    comando.Parameters.Add(New SqlParameter("@balance_total", SqlDbType.Float)).Value = objSabanaAplicacion.balanceActual
                    comando.Parameters.Add(New SqlParameter("@Balance_Acumulado", SqlDbType.Float)).Value = objSabanaAplicacion.balanceAcumulado
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objSabanaAplicacion.usuario
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHorasabana
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_SILVERMAN_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaSilverman
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_VALORACION_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaValoracion
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_INGRESO_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaIngreso
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_EGRESO_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaEgreso
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_BALANCE_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaBalance
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_APLICACION_MED_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaMedicamento
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaNPT
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_APLICACION_GOTEO_GUARDAR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHorasabanaGoteo
                    comando.ExecuteNonQuery()
                '    Try
                '        trnsccion.Commit()
                '    Catch ex As Exception
                '        Try
                '            trnsccion.Rollback()
                '            Throw ex
                '        Catch ex1 As Exception
                '            Throw ex
                '        End Try
                '    End Try
                'End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarSabanaAM(objSabanaAplicacion As SabanaAplicacionMedR)
        Try
            Using comando = New SqlCommand()
                ' Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                comando.Connection = FormPrincipal.cnxion
                    comando.CommandType = CommandType.StoredProcedure
                ' comando.Transaction = trnsccion
                comando.CommandText = "[SP_SABANA_CREAR_R]"
                    comando.Parameters.Add(New SqlParameter("@Fecha_Real", SqlDbType.DateTime)).Value = objSabanaAplicacion.fechaReal
                    comando.Parameters.Add(New SqlParameter("@Peso", SqlDbType.Float)).Value = objSabanaAplicacion.pesoActual
                    comando.Parameters.Add(New SqlParameter("@Gasto_U", SqlDbType.Float)).Value = objSabanaAplicacion.gastoUrinario
                    comando.Parameters.Add(New SqlParameter("@Ingreso_total", SqlDbType.Float)).Value = objSabanaAplicacion.ingresoTotal
                    comando.Parameters.Add(New SqlParameter("@Egreso_total", SqlDbType.Float)).Value = objSabanaAplicacion.egresoTotal
                    comando.Parameters.Add(New SqlParameter("@balance_total", SqlDbType.Float)).Value = objSabanaAplicacion.balanceActual
                    comando.Parameters.Add(New SqlParameter("@Balance_Acumulado", SqlDbType.Float)).Value = objSabanaAplicacion.balanceAcumulado
                    comando.Parameters.Add(New SqlParameter("@Regitro", SqlDbType.Int)).Value = objSabanaAplicacion.nRegistro
                    comando.Parameters.Add(New SqlParameter("@Usuario_Real", SqlDbType.Int)).Value = objSabanaAplicacion.usuarioReal
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objSabanaAplicacion.usuario
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    objSabanaAplicacion.codigoSabana = CType(comando.ExecuteScalar, String)
                    comando.Parameters.Clear()
                comando.CommandText = "[SP_SABANA_SIGNO_VITAL_CREAR_R]"
                comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                comando.Parameters.Add(New SqlParameter("@Usuario_Real", SqlDbType.Int)).Value = objSabanaAplicacion.usuarioReal
                comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objSabanaAplicacion.usuario
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHorasabana
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_INGRESO_GUARDAR_R]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaIngreso
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_EGRESO_GUARDAR_R]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaEgreso
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_BALANCE_GUARDAR_R]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaBalance
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_APLICACION_MED_GUARDAR_R]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaMedicamento
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaNPT
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_APLICACION_GOTEO_GUARDAR_R]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHorasabanaGoteo
                    comando.ExecuteNonQuery()
                '    Try
                '        trnsccion.Commit()
                '    Catch ex As Exception
                '        Try
                '            trnsccion.Rollback()
                '            Throw ex
                '        Catch ex1 As Exception
                '            Throw ex
                '        End Try
                '    End Try
                'End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Shared Sub actualizarSabanaAM(objSabanaAplicacion As SabanaAplicacionMedR)
        Try
            Using comando = New SqlCommand()
                'Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                comando.Connection = FormPrincipal.cnxion
                    comando.CommandType = CommandType.StoredProcedure
                'comando.Transaction = trnsccion
                comando.CommandText = "[SP_SABANA_ACTUALIZAR_R]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@Peso", SqlDbType.Float)).Value = objSabanaAplicacion.pesoActual
                    comando.Parameters.Add(New SqlParameter("@Gasto_U", SqlDbType.Float)).Value = objSabanaAplicacion.gastoUrinario
                    comando.Parameters.Add(New SqlParameter("@Ingreso_total", SqlDbType.Float)).Value = objSabanaAplicacion.ingresoTotal
                    comando.Parameters.Add(New SqlParameter("@Egreso_total", SqlDbType.Float)).Value = objSabanaAplicacion.egresoTotal
                    comando.Parameters.Add(New SqlParameter("@balance_total", SqlDbType.Float)).Value = objSabanaAplicacion.balanceActual
                    comando.Parameters.Add(New SqlParameter("@Balance_Acumulado", SqlDbType.Float)).Value = objSabanaAplicacion.balanceAcumulado
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objSabanaAplicacion.usuario
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                comando.CommandText = "[SP_SABANA_SIGNO_VITAL_ACTUALIZAR_R]"
                comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objSabanaAplicacion.usuario
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHorasabana
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_INGRESO_GUARDAR_R]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaIngreso
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_EGRESO_GUARDAR_R]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaEgreso
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_BALANCE_GUARDAR_R]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaBalance
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_APLICACION_MED_GUARDAR_R]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaMedicamento
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaNPT
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_APLICACION_GOTEO_GUARDAR_R]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objSabanaAplicacion.codigoEP
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHorasabanaGoteo
                    comando.ExecuteNonQuery()
                '    Try
                '        trnsccion.Commit()
                '    Catch ex As Exception
                '        Try
                '            trnsccion.Rollback()
                '            Throw ex
                '        Catch ex1 As Exception
                '            Throw ex
                '        End Try
                '    End Try
                'End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarSabanaAF(objSabanaAplicacion As SabanaAplicacionMedRR)
        Try
            Using comando = New SqlCommand()
                'Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                comando.Connection = FormPrincipal.cnxion
                    comando.CommandType = CommandType.StoredProcedure
                ' comando.Transaction = trnsccion
                comando.CommandText = "[SP_SABANA_CREAR_RR]"
                    comando.Parameters.Add(New SqlParameter("@Fecha_Real", SqlDbType.DateTime)).Value = objSabanaAplicacion.fechaReal
                    comando.Parameters.Add(New SqlParameter("@Peso", SqlDbType.Float)).Value = objSabanaAplicacion.pesoActual
                    comando.Parameters.Add(New SqlParameter("@Gasto_U", SqlDbType.Float)).Value = objSabanaAplicacion.gastoUrinario
                    comando.Parameters.Add(New SqlParameter("@Ingreso_total", SqlDbType.Float)).Value = objSabanaAplicacion.ingresoTotal
                    comando.Parameters.Add(New SqlParameter("@Egreso_total", SqlDbType.Float)).Value = objSabanaAplicacion.egresoTotal
                    comando.Parameters.Add(New SqlParameter("@balance_total", SqlDbType.Float)).Value = objSabanaAplicacion.balanceActual
                    comando.Parameters.Add(New SqlParameter("@Balance_Acumulado", SqlDbType.Float)).Value = objSabanaAplicacion.balanceAcumulado
                    comando.Parameters.Add(New SqlParameter("@Regitro", SqlDbType.Int)).Value = objSabanaAplicacion.nRegistro
                    comando.Parameters.Add(New SqlParameter("@Usuario_Real", SqlDbType.Int)).Value = objSabanaAplicacion.usuarioReal
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objSabanaAplicacion.usuario
                    objSabanaAplicacion.codigoSabana = CType(comando.ExecuteScalar, String)
                    comando.Parameters.Clear()
                comando.CommandText = "[SP_SABANA_SIGNO_VITAL_CREAR_RR]"
                comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objSabanaAplicacion.usuario
                comando.Parameters.Add(New SqlParameter("@Usuario_Real", SqlDbType.Int)).Value = objSabanaAplicacion.usuarioReal
                comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHorasabana
                comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_INGRESO_GUARDAR_RR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaIngreso
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_EGRESO_GUARDAR_RR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaEgreso
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_BALANCE_GUARDAR_RR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaBalance
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_APLICACION_MED_GUARDAR_RR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaMedicamento
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaNPT
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_APLICACION_GOTEO_GUARDAR_RR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHorasabanaGoteo
                    comando.ExecuteNonQuery()
                '    Try
                '        trnsccion.Commit()
                '    Catch ex As Exception
                '        Try
                '            trnsccion.Rollback()
                '            Throw ex
                '        Catch ex1 As Exception
                '            Throw ex
                '        End Try
                '    End Try
                'End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Shared Sub actualizarSabanaAF(objSabanaAplicacion As SabanaAplicacionMedRR)
        Try
            Using comando = New SqlCommand()
                '   Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                comando.Connection = FormPrincipal.cnxion
                    comando.CommandType = CommandType.StoredProcedure
                ' comando.Transaction = trnsccion
                comando.CommandText = "[SP_SABANA_ACTUALIZAR_RR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@Peso", SqlDbType.Float)).Value = objSabanaAplicacion.pesoActual
                    comando.Parameters.Add(New SqlParameter("@Gasto_U", SqlDbType.Float)).Value = objSabanaAplicacion.gastoUrinario
                    comando.Parameters.Add(New SqlParameter("@Ingreso_total", SqlDbType.Float)).Value = objSabanaAplicacion.ingresoTotal
                    comando.Parameters.Add(New SqlParameter("@Egreso_total", SqlDbType.Float)).Value = objSabanaAplicacion.egresoTotal
                    comando.Parameters.Add(New SqlParameter("@balance_total", SqlDbType.Float)).Value = objSabanaAplicacion.balanceActual
                    comando.Parameters.Add(New SqlParameter("@Balance_Acumulado", SqlDbType.Float)).Value = objSabanaAplicacion.balanceAcumulado
                    comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objSabanaAplicacion.usuario
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                comando.CommandText = "[SP_SABANA_SIGNO_VITAL_ACTUALIZAR_RR]"
                comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                comando.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objSabanaAplicacion.usuario
                comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHorasabana
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_INGRESO_GUARDAR_RR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaIngreso
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_EGRESO_GUARDAR_RR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaEgreso
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_BALANCE_GUARDAR_RR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaBalance
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_APLICACION_MED_GUARDAR_RR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaMedicamento
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHoraSabanaNPT
                    comando.ExecuteNonQuery()
                    comando.Parameters.Clear()
                    comando.CommandText = "[SP_SABANA_APLICACION_GOTEO_GUARDAR_RR]"
                    comando.Parameters.Add(New SqlParameter("@IDSABANA", SqlDbType.Int)).Value = objSabanaAplicacion.codigoSabana
                    comando.Parameters.Add(New SqlParameter("@TablaSign", SqlDbType.Structured)).Value = objSabanaAplicacion.dtHorasabanaGoteo
                    comando.ExecuteNonQuery()
                '    Try
                '        trnsccion.Commit()
                '    Catch ex As Exception
                '        Try
                '            trnsccion.Rollback()
                '            Throw ex
                '        Catch ex1 As Exception
                '            Throw ex
                '        End Try
                '    End Try
                'End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarSabanaAutomaticaHC(objOrdenMedica As OrdenMedica)
        Try
            Using comando = New SqlCommand()
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "[SP_SABANA_AUTOMATICA_GUARDAR]"
                comando.Parameters.Add(New SqlParameter("@nregistro", SqlDbType.Int)).Value = objOrdenMedica.registro
                comando.Parameters.Add(New SqlParameter("@fecha", SqlDbType.Date)).Value = objOrdenMedica.fechaOrden.Date
                comando.Parameters.Add(New SqlParameter("@usuario_real", SqlDbType.Int)).Value = objOrdenMedica.usuarioReal
                comando.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                comando.Parameters.Add(New SqlParameter("@valorpredeterminado", SqlDbType.NVarChar)).Value = ConstantesHC.IDENTIFICADOR_SABANA_NO_APLICADO
                comando.Parameters.Add(New SqlParameter("@IDINFUSIONIV", SqlDbType.Int)).Value = ConstantesHC.CODIGO_INFUSION_IV
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_INGRESO", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_INGRESO
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_INGRESOBALANCE", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_INGRESO_BALANCE
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_EGRESO", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_EGRESO
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_EGRESOBALANCE", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_EGRESO_BALANCE
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_GU", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_GASTO_URINARIO
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_BALANCE", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_BALANCE
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Shared Sub guardarSabanaAutomaticaAM(objOrdenMedica As OrdenMedica)
        Try
            Using comando = New SqlCommand()
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "[SP_SABANA_AUTOMATICA_GUARDAR_R]"
                comando.Parameters.Add(New SqlParameter("@nregistro", SqlDbType.Int)).Value = objOrdenMedica.registro
                comando.Parameters.Add(New SqlParameter("@fecha", SqlDbType.Date)).Value = objOrdenMedica.fechaOrden.Date
                comando.Parameters.Add(New SqlParameter("@usuario_real", SqlDbType.Int)).Value = objOrdenMedica.usuarioReal
                comando.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                comando.Parameters.Add(New SqlParameter("@valorpredeterminado", SqlDbType.NVarChar)).Value = ConstantesHC.IDENTIFICADOR_SABANA_NO_APLICADO
                comando.Parameters.Add(New SqlParameter("@IDINFUSIONIV", SqlDbType.Int)).Value = ConstantesHC.CODIGO_INFUSION_IV
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_INGRESO", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_INGRESO
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_INGRESOBALANCE", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_INGRESO_BALANCE
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_EGRESO", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_EGRESO
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_EGRESOBALANCE", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_EGRESO_BALANCE
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_GU", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_GASTO_URINARIO
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_BALANCE", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_BALANCE
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Shared Sub guardarSabanaAutomaticaAF(objOrdenMedica As OrdenMedica)
        Try
            Using comando = New SqlCommand()
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = "[SP_SABANA_AUTOMATICA_GUARDAR_RR]"
                comando.Parameters.Add(New SqlParameter("@nregistro", SqlDbType.Int)).Value = objOrdenMedica.registro
                comando.Parameters.Add(New SqlParameter("@fecha", SqlDbType.Date)).Value = objOrdenMedica.fechaOrden.Date
                comando.Parameters.Add(New SqlParameter("@usuario_real", SqlDbType.Int)).Value = objOrdenMedica.usuarioReal
                comando.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                comando.Parameters.Add(New SqlParameter("@valorpredeterminado", SqlDbType.NVarChar)).Value = ConstantesHC.IDENTIFICADOR_SABANA_NO_APLICADO
                comando.Parameters.Add(New SqlParameter("@IDINFUSIONIV", SqlDbType.Int)).Value = ConstantesHC.CODIGO_INFUSION_IV
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_INGRESO", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_INGRESO
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_INGRESOBALANCE", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_INGRESO_BALANCE
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_EGRESO", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_EGRESO
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_EGRESOBALANCE", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_EGRESO_BALANCE
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_GU", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_GASTO_URINARIO
                comando.Parameters.Add(New SqlParameter("@IDTOTAL_BALANCE", SqlDbType.Int)).Value = ConstantesHC.CODIGO_TOTAL_BALANCE
                comando.ExecuteNonQuery()
            End Using
        Catch ex As SqlException
            Throw ex
        End Try

    End Sub
End Class
