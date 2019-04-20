Imports System.Data.SqlClient
Public Class MedicamentoNoPosDAL

    Public Shared Sub guardarMedicamentosNOPOSHc(objMedicamentos As SolicitudMedicamentosNoPos, dtMedicamento As DataTable)


        Try

            Using consulta As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_SOLICITUD_MEDICAMENTOS_CREAR"
                    consulta.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.NVarChar))
                    consulta.Parameters("@N_Registro").Value = objMedicamentos.registro
                    consulta.Parameters.Add(New SqlParameter("@Fecha_Solicitud", SqlDbType.DateTime))
                    consulta.Parameters("@Fecha_Solicitud").Value = objMedicamentos.fechaSolicitud
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Interno1", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Interno1").Value = objMedicamentos.codigoEquivalenciaUno
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Interno2", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Interno2").Value = objMedicamentos.codigoEquivalenciaDos
                    consulta.Parameters.Add(New SqlParameter("@Resumen_hc", SqlDbType.NVarChar))
                    consulta.Parameters("@Resumen_hc").Value = objMedicamentos.resumen
                    consulta.Parameters.Add(New SqlParameter("@Si", SqlDbType.Bit))
                    consulta.Parameters("@Si").Value = objMedicamentos.rdsi
                    consulta.Parameters.Add(New SqlParameter("@No", SqlDbType.Bit))
                    consulta.Parameters("@No").Value = objMedicamentos.rdno
                    consulta.Parameters.Add(New SqlParameter("@RespuestaPOS", SqlDbType.NVarChar))
                    consulta.Parameters("@RespuestaPOS").Value = objMedicamentos.respuesta
                    consulta.Parameters.Add(New SqlParameter("@Efecto_Deseado", SqlDbType.NVarChar))
                    consulta.Parameters("@Efecto_Deseado").Value = objMedicamentos.efecto
                    consulta.Parameters.Add(New SqlParameter("@Iniciacion_Efecto", SqlDbType.NVarChar))
                    consulta.Parameters("@Iniciacion_Efecto").Value = objMedicamentos.inicioefecto
                    consulta.Parameters.Add(New SqlParameter("@Explicar_Razon", SqlDbType.NVarChar))
                    consulta.Parameters("@Explicar_Razon").Value = objMedicamentos.explicarazon
                    consulta.Parameters.Add(New SqlParameter("@Criterio1", SqlDbType.NVarChar))
                    consulta.Parameters("@Criterio1").Value = objMedicamentos.cbocriterio
                    consulta.Parameters.Add(New SqlParameter("@Criterio2", SqlDbType.NVarChar))
                    consulta.Parameters("@Criterio2").Value = objMedicamentos.cbocriterio2
                    consulta.Parameters.Add(New SqlParameter("@PreCoEsTox", SqlDbType.NVarChar))
                    consulta.Parameters("@PreCoEsTox").Value = objMedicamentos.precaucion
                    consulta.Parameters.Add(New SqlParameter("@PeligroPaciente", SqlDbType.NVarChar))
                    consulta.Parameters("@PeligroPaciente").Value = objMedicamentos.pelipaciente
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeutica", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeutica").Value = objMedicamentos.cboposibilidad
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeuticaPOS", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeuticaPOS").Value = objMedicamentos.reacciones
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeuticaPOScontra", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeuticaPOScontra").Value = objMedicamentos.contraindicaciones
                    consulta.Parameters.Add(New SqlParameter("@Duracion_Tto", SqlDbType.NVarChar))
                    consulta.Parameters("@Duracion_Tto").Value = objMedicamentos.duraciontratamiento
                    consulta.Parameters.Add(New SqlParameter("@Dosis", SqlDbType.NVarChar))
                    consulta.Parameters("@Dosis").Value = objMedicamentos.dosis
                    consulta.Parameters.Add(New SqlParameter("@Frecuencia", SqlDbType.NVarChar))
                    consulta.Parameters("@Frecuencia").Value = objMedicamentos.frecuencia
                    consulta.Parameters.Add(New SqlParameter("@Num_Dosis", SqlDbType.NVarChar))
                    consulta.Parameters("@Num_Dosis").Value = objMedicamentos.ndosis
                    consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.NVarChar))
                    consulta.Parameters("@Usuario_Creacion").Value = objMedicamentos.usuarioInforme
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.NVarChar))
                    consulta.Parameters("@CODIGO_EP").Value = objMedicamentos.codigoPuntoActual
                    objMedicamentos.Codigo = CType(consulta.ExecuteScalar, Integer)
                    consulta.Parameters.Clear()
                    consulta.CommandText = "[SP_SOLICITUD_MEDICAMENTOS_DETALLE_ELIMINAR]"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                    consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.NVarChar))
                    consulta.Parameters("@CODIGO_EP").Value = objMedicamentos.codigoPuntoActual
                    consulta.ExecuteNonQuery()
                    For i As Int32 = 0 To dtMedicamento.Rows.Count - 1
                        consulta.Parameters.Clear()
                        consulta.CommandText = "SP_SOLICITUD_MEDICAMENTOS_DETALLE_CREAR"
                        consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EP").Value = objMedicamentos.codigoPuntoActual
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = dtMedicamento.Rows(i).Item(0).ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EVO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EVO").Value = dtMedicamento.Rows(i).Item(2).ToString
                        consulta.ExecuteNonQuery()
                    Next

                    consulta.Parameters.Clear()
                    consulta.CommandText = "SP_CTC_AUTOMATICO"
                    consulta.Parameters.Add(New SqlParameter("@pSolicitud", SqlDbType.Int)).Value = objMedicamentos.Codigo
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

    Public Shared Sub actualizarMedicamentosNOPOSHc(objMedicamentos As SolicitudMedicamentosNoPos, dtMedicamento As DataTable)


        Try

            Using consulta As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_SOLICITUD_MEDICAMENTOS_ACTUALIZAR"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                    consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                    consulta.Parameters.Add(New SqlParameter("@Fecha_Solicitud", SqlDbType.DateTime))
                    consulta.Parameters("@Fecha_Solicitud").Value = objMedicamentos.fechaSolicitud
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Interno1", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Interno1").Value = objMedicamentos.codigoEquivalenciaUno
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Interno2", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Interno2").Value = objMedicamentos.codigoEquivalenciaDos
                    consulta.Parameters.Add(New SqlParameter("@Resumen_hc", SqlDbType.NVarChar))
                    consulta.Parameters("@Resumen_hc").Value = objMedicamentos.resumen
                    consulta.Parameters.Add(New SqlParameter("@Si", SqlDbType.Bit))
                    consulta.Parameters("@Si").Value = objMedicamentos.rdsi
                    consulta.Parameters.Add(New SqlParameter("@No", SqlDbType.Bit))
                    consulta.Parameters("@No").Value = objMedicamentos.rdno
                    consulta.Parameters.Add(New SqlParameter("@RespuestaPOS", SqlDbType.NVarChar))
                    consulta.Parameters("@RespuestaPOS").Value = objMedicamentos.respuesta
                    consulta.Parameters.Add(New SqlParameter("@Efecto_Deseado", SqlDbType.NVarChar))
                    consulta.Parameters("@Efecto_Deseado").Value = objMedicamentos.efecto
                    consulta.Parameters.Add(New SqlParameter("@Iniciacion_Efecto", SqlDbType.NVarChar))
                    consulta.Parameters("@Iniciacion_Efecto").Value = objMedicamentos.inicioefecto
                    consulta.Parameters.Add(New SqlParameter("@Explicar_Razon", SqlDbType.NVarChar))
                    consulta.Parameters("@Explicar_Razon").Value = objMedicamentos.explicarazon
                    consulta.Parameters.Add(New SqlParameter("@Criterio1", SqlDbType.NVarChar))
                    consulta.Parameters("@Criterio1").Value = objMedicamentos.cbocriterio
                    consulta.Parameters.Add(New SqlParameter("@Criterio2", SqlDbType.NVarChar))
                    consulta.Parameters("@Criterio2").Value = objMedicamentos.cbocriterio2
                    consulta.Parameters.Add(New SqlParameter("@PreCoEsTox", SqlDbType.NVarChar))
                    consulta.Parameters("@PreCoEsTox").Value = objMedicamentos.precaucion
                    consulta.Parameters.Add(New SqlParameter("@PeligroPaciente", SqlDbType.NVarChar))
                    consulta.Parameters("@PeligroPaciente").Value = objMedicamentos.pelipaciente
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeutica", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeutica").Value = objMedicamentos.cboposibilidad
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeuticaPOS", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeuticaPOS").Value = objMedicamentos.reacciones
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeuticaPOScontra", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeuticaPOScontra").Value = objMedicamentos.contraindicaciones
                    consulta.Parameters.Add(New SqlParameter("@Duracion_Tto", SqlDbType.NVarChar))
                    consulta.Parameters("@Duracion_Tto").Value = objMedicamentos.duraciontratamiento
                    consulta.Parameters.Add(New SqlParameter("@Dosis", SqlDbType.NVarChar))
                    consulta.Parameters("@Dosis").Value = objMedicamentos.dosis
                    consulta.Parameters.Add(New SqlParameter("@Frecuencia", SqlDbType.NVarChar))
                    consulta.Parameters("@Frecuencia").Value = objMedicamentos.frecuencia
                    consulta.Parameters.Add(New SqlParameter("@Num_Dosis", SqlDbType.NVarChar))
                    consulta.Parameters("@Num_Dosis").Value = objMedicamentos.ndosis
                    consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.NVarChar))
                    consulta.Parameters("@Usuario").Value = objMedicamentos.usuarioInforme
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.NVarChar))
                    consulta.Parameters("@CODIGO_EP").Value = objMedicamentos.codigoPuntoActual
                    consulta.ExecuteNonQuery()
                    consulta.Parameters.Clear()
                    consulta.CommandText = "[SP_SOLICITUD_MEDICAMENTOS_DETALLE_ELIMINAR]"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                    consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                    consulta.Parameters("@CODIGO_EP").Value = objMedicamentos.codigoPuntoActual
                    consulta.ExecuteNonQuery()
                    For i As Int32 = 0 To dtMedicamento.Rows.Count - 1
                        consulta.Parameters.Clear()
                        consulta.CommandText = "SP_SOLICITUD_MEDICAMENTOS_DETALLE_CREAR"
                        consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EP").Value = objMedicamentos.codigoPuntoActual
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = dtMedicamento.Rows(i).Item(0).ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EVO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EVO").Value = dtMedicamento.Rows(i).Item(3).ToString
                        consulta.ExecuteNonQuery()
                    Next
                    consulta.Parameters.Clear()
                    consulta.CommandText = "SP_CTC_AUTOMATICO"
                    consulta.Parameters.Add(New SqlParameter("@pSolicitud", SqlDbType.Int)).Value = objMedicamentos.Codigo
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


    Public Shared Sub guardarMedicamentosNOPOSAM(objMedicamentos As SolicitudMedicamentosNoPos, dtMedicamento As DataTable)


        Try

            Using consulta As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_SOLICITUD_MEDICAMENTOS_CREAR_R"
                    consulta.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.NVarChar))
                    consulta.Parameters("@N_Registro").Value = objMedicamentos.registro
                    consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar))
                    consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                    consulta.Parameters.Add(New SqlParameter("@Fecha_Solicitud", SqlDbType.DateTime))
                    consulta.Parameters("@Fecha_Solicitud").Value = objMedicamentos.fechaSolicitud
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Interno1", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Interno1").Value = objMedicamentos.codigoEquivalenciaUno
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Interno2", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Interno2").Value = objMedicamentos.codigoEquivalenciaDos
                    consulta.Parameters.Add(New SqlParameter("@Resumen_hc", SqlDbType.NVarChar))
                    consulta.Parameters("@Resumen_hc").Value = objMedicamentos.resumen
                    consulta.Parameters.Add(New SqlParameter("@Si", SqlDbType.Bit))
                    consulta.Parameters("@Si").Value = objMedicamentos.rdsi
                    consulta.Parameters.Add(New SqlParameter("@No", SqlDbType.Bit))
                    consulta.Parameters("@No").Value = objMedicamentos.rdno
                    consulta.Parameters.Add(New SqlParameter("@RespuestaPOS", SqlDbType.NVarChar))
                    consulta.Parameters("@RespuestaPOS").Value = objMedicamentos.respuesta
                    consulta.Parameters.Add(New SqlParameter("@Efecto_Deseado", SqlDbType.NVarChar))
                    consulta.Parameters("@Efecto_Deseado").Value = objMedicamentos.efecto
                    consulta.Parameters.Add(New SqlParameter("@Iniciacion_Efecto", SqlDbType.NVarChar))
                    consulta.Parameters("@Iniciacion_Efecto").Value = objMedicamentos.inicioefecto
                    consulta.Parameters.Add(New SqlParameter("@Explicar_Razon", SqlDbType.NVarChar))
                    consulta.Parameters("@Explicar_Razon").Value = objMedicamentos.explicarazon
                    consulta.Parameters.Add(New SqlParameter("@Criterio1", SqlDbType.NVarChar))
                    consulta.Parameters("@Criterio1").Value = objMedicamentos.cbocriterio
                    consulta.Parameters.Add(New SqlParameter("@Criterio2", SqlDbType.NVarChar))
                    consulta.Parameters("@Criterio2").Value = objMedicamentos.cbocriterio2
                    consulta.Parameters.Add(New SqlParameter("@PreCoEsTox", SqlDbType.NVarChar))
                    consulta.Parameters("@PreCoEsTox").Value = objMedicamentos.precaucion
                    consulta.Parameters.Add(New SqlParameter("@PeligroPaciente", SqlDbType.NVarChar))
                    consulta.Parameters("@PeligroPaciente").Value = objMedicamentos.pelipaciente
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeutica", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeutica").Value = objMedicamentos.cboposibilidad
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeuticaPOS", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeuticaPOS").Value = objMedicamentos.reacciones
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeuticaPOScontra", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeuticaPOScontra").Value = objMedicamentos.contraindicaciones
                    consulta.Parameters.Add(New SqlParameter("@Duracion_Tto", SqlDbType.NVarChar))
                    consulta.Parameters("@Duracion_Tto").Value = objMedicamentos.duraciontratamiento
                    consulta.Parameters.Add(New SqlParameter("@Dosis", SqlDbType.NVarChar))
                    consulta.Parameters("@Dosis").Value = objMedicamentos.dosis
                    consulta.Parameters.Add(New SqlParameter("@Frecuencia", SqlDbType.NVarChar))
                    consulta.Parameters("@Frecuencia").Value = objMedicamentos.frecuencia
                    consulta.Parameters.Add(New SqlParameter("@Num_Dosis", SqlDbType.NVarChar))
                    consulta.Parameters("@Num_Dosis").Value = objMedicamentos.ndosis
                    consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.NVarChar))
                    consulta.Parameters("@Usuario_Creacion").Value = objMedicamentos.usuarioInforme
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.NVarChar))
                    consulta.Parameters("@CODIGO_EP").Value = objMedicamentos.codigoPuntoActual
                    objMedicamentos.Codigo = CType(consulta.ExecuteScalar, Integer)
                    consulta.Parameters.Clear()
                    consulta.CommandText = "[SP_SOLICITUD_MEDICAMENTOS_DETALLE_ELIMINAR_R]"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                    consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.NVarChar))
                    consulta.Parameters("@CODIGO_EP").Value = objMedicamentos.codigoPuntoActual
                    consulta.ExecuteNonQuery()
                    For i As Int32 = 0 To dtMedicamento.Rows.Count - 1
                        consulta.Parameters.Clear()
                        consulta.CommandText = "SP_SOLICITUD_MEDICAMENTOS_DETALLE_CREAR_R"
                        consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EP").Value = objMedicamentos.codigoPuntoActual
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = dtMedicamento.Rows(i).Item(0).ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EVO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EVO").Value = dtMedicamento.Rows(i).Item(2).ToString
                        consulta.ExecuteNonQuery()
                    Next


                    consulta.Parameters.Clear()
                    consulta.CommandText = "SP_CTC_AUTOMATICO_R"
                    consulta.Parameters.Add(New SqlParameter("@pSolicitud", SqlDbType.Int)).Value = objMedicamentos.Codigo
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

    Public Shared Sub actualizarMedicamentosNOPOSAM(objMedicamentos As SolicitudMedicamentosNoPos, dtMedicamento As DataTable)


        Try

            Using consulta As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_SOLICITUD_MEDICAMENTOS_ACTUALIZAR_R"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                    consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                    consulta.Parameters.Add(New SqlParameter("@Fecha_Solicitud", SqlDbType.DateTime))
                    consulta.Parameters("@Fecha_Solicitud").Value = objMedicamentos.fechaSolicitud
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Interno1", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Interno1").Value = objMedicamentos.codigoEquivalenciaUno
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Interno2", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Interno2").Value = objMedicamentos.codigoEquivalenciaDos
                    consulta.Parameters.Add(New SqlParameter("@Resumen_hc", SqlDbType.NVarChar))
                    consulta.Parameters("@Resumen_hc").Value = objMedicamentos.resumen
                    consulta.Parameters.Add(New SqlParameter("@Si", SqlDbType.Bit))
                    consulta.Parameters("@Si").Value = objMedicamentos.rdsi
                    consulta.Parameters.Add(New SqlParameter("@No", SqlDbType.Bit))
                    consulta.Parameters("@No").Value = objMedicamentos.rdno
                    consulta.Parameters.Add(New SqlParameter("@RespuestaPOS", SqlDbType.NVarChar))
                    consulta.Parameters("@RespuestaPOS").Value = objMedicamentos.respuesta
                    consulta.Parameters.Add(New SqlParameter("@Efecto_Deseado", SqlDbType.NVarChar))
                    consulta.Parameters("@Efecto_Deseado").Value = objMedicamentos.efecto
                    consulta.Parameters.Add(New SqlParameter("@Iniciacion_Efecto", SqlDbType.NVarChar))
                    consulta.Parameters("@Iniciacion_Efecto").Value = objMedicamentos.inicioefecto
                    consulta.Parameters.Add(New SqlParameter("@Explicar_Razon", SqlDbType.NVarChar))
                    consulta.Parameters("@Explicar_Razon").Value = objMedicamentos.explicarazon
                    consulta.Parameters.Add(New SqlParameter("@Criterio1", SqlDbType.NVarChar))
                    consulta.Parameters("@Criterio1").Value = objMedicamentos.cbocriterio
                    consulta.Parameters.Add(New SqlParameter("@Criterio2", SqlDbType.NVarChar))
                    consulta.Parameters("@Criterio2").Value = objMedicamentos.cbocriterio2
                    consulta.Parameters.Add(New SqlParameter("@PreCoEsTox", SqlDbType.NVarChar))
                    consulta.Parameters("@PreCoEsTox").Value = objMedicamentos.precaucion
                    consulta.Parameters.Add(New SqlParameter("@PeligroPaciente", SqlDbType.NVarChar))
                    consulta.Parameters("@PeligroPaciente").Value = objMedicamentos.pelipaciente
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeutica", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeutica").Value = objMedicamentos.cboposibilidad
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeuticaPOS", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeuticaPOS").Value = objMedicamentos.reacciones
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeuticaPOScontra", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeuticaPOScontra").Value = objMedicamentos.contraindicaciones
                    consulta.Parameters.Add(New SqlParameter("@Duracion_Tto", SqlDbType.NVarChar))
                    consulta.Parameters("@Duracion_Tto").Value = objMedicamentos.duraciontratamiento
                    consulta.Parameters.Add(New SqlParameter("@Dosis", SqlDbType.NVarChar))
                    consulta.Parameters("@Dosis").Value = objMedicamentos.dosis
                    consulta.Parameters.Add(New SqlParameter("@Frecuencia", SqlDbType.NVarChar))
                    consulta.Parameters("@Frecuencia").Value = objMedicamentos.frecuencia
                    consulta.Parameters.Add(New SqlParameter("@Num_Dosis", SqlDbType.NVarChar))
                    consulta.Parameters("@Num_Dosis").Value = objMedicamentos.ndosis
                    consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.NVarChar))
                    consulta.Parameters("@Usuario").Value = objMedicamentos.usuarioInforme
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.NVarChar))
                    consulta.Parameters("@CODIGO_EP").Value = objMedicamentos.codigoPuntoActual
                    consulta.ExecuteNonQuery()
                    consulta.Parameters.Clear()
                    consulta.CommandText = "[SP_SOLICITUD_MEDICAMENTOS_DETALLE_ELIMINAR_R]"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                    consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                    consulta.Parameters("@CODIGO_EP").Value = objMedicamentos.codigoPuntoActual
                    consulta.ExecuteNonQuery()
                    For i As Int32 = 0 To dtMedicamento.Rows.Count - 1
                        consulta.Parameters.Clear()
                        consulta.CommandText = "SP_SOLICITUD_MEDICAMENTOS_DETALLE_CREAR_R"
                        consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EP").Value = objMedicamentos.codigoPuntoActual
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = dtMedicamento.Rows(i).Item(0).ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EVO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EVO").Value = dtMedicamento.Rows(i).Item(3).ToString
                        consulta.ExecuteNonQuery()
                    Next

                    consulta.Parameters.Clear()
                    consulta.CommandText = "SP_CTC_AUTOMATICO_R"
                    consulta.Parameters.Add(New SqlParameter("@pSolicitud", SqlDbType.Int)).Value = objMedicamentos.Codigo
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


    Public Shared Sub guardarMedicamentosNOPOSAF(objMedicamentos As SolicitudMedicamentosNoPos, dtMedicamento As DataTable)


        Try

            Using consulta As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_SOLICITUD_MEDICAMENTOS_CREAR_RR"
                    consulta.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.NVarChar))
                    consulta.Parameters("@N_Registro").Value = objMedicamentos.registro
                    consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.NVarChar))
                    consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                    consulta.Parameters.Add(New SqlParameter("@Fecha_Solicitud", SqlDbType.DateTime))
                    consulta.Parameters("@Fecha_Solicitud").Value = objMedicamentos.fechaSolicitud
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Interno1", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Interno1").Value = objMedicamentos.codigoEquivalenciaUno
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Interno2", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Interno2").Value = objMedicamentos.codigoEquivalenciaDos
                    consulta.Parameters.Add(New SqlParameter("@Resumen_hc", SqlDbType.NVarChar))
                    consulta.Parameters("@Resumen_hc").Value = objMedicamentos.resumen
                    consulta.Parameters.Add(New SqlParameter("@Si", SqlDbType.Bit))
                    consulta.Parameters("@Si").Value = objMedicamentos.rdsi
                    consulta.Parameters.Add(New SqlParameter("@No", SqlDbType.Bit))
                    consulta.Parameters("@No").Value = objMedicamentos.rdno
                    consulta.Parameters.Add(New SqlParameter("@RespuestaPOS", SqlDbType.NVarChar))
                    consulta.Parameters("@RespuestaPOS").Value = objMedicamentos.respuesta
                    consulta.Parameters.Add(New SqlParameter("@Efecto_Deseado", SqlDbType.NVarChar))
                    consulta.Parameters("@Efecto_Deseado").Value = objMedicamentos.efecto
                    consulta.Parameters.Add(New SqlParameter("@Iniciacion_Efecto", SqlDbType.NVarChar))
                    consulta.Parameters("@Iniciacion_Efecto").Value = objMedicamentos.inicioefecto
                    consulta.Parameters.Add(New SqlParameter("@Explicar_Razon", SqlDbType.NVarChar))
                    consulta.Parameters("@Explicar_Razon").Value = objMedicamentos.explicarazon
                    consulta.Parameters.Add(New SqlParameter("@Criterio1", SqlDbType.NVarChar))
                    consulta.Parameters("@Criterio1").Value = objMedicamentos.cbocriterio
                    consulta.Parameters.Add(New SqlParameter("@Criterio2", SqlDbType.NVarChar))
                    consulta.Parameters("@Criterio2").Value = objMedicamentos.cbocriterio2
                    consulta.Parameters.Add(New SqlParameter("@PreCoEsTox", SqlDbType.NVarChar))
                    consulta.Parameters("@PreCoEsTox").Value = objMedicamentos.precaucion
                    consulta.Parameters.Add(New SqlParameter("@PeligroPaciente", SqlDbType.NVarChar))
                    consulta.Parameters("@PeligroPaciente").Value = objMedicamentos.pelipaciente
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeutica", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeutica").Value = objMedicamentos.cboposibilidad
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeuticaPOS", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeuticaPOS").Value = objMedicamentos.reacciones
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeuticaPOScontra", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeuticaPOScontra").Value = objMedicamentos.contraindicaciones
                    consulta.Parameters.Add(New SqlParameter("@Duracion_Tto", SqlDbType.NVarChar))
                    consulta.Parameters("@Duracion_Tto").Value = objMedicamentos.duraciontratamiento
                    consulta.Parameters.Add(New SqlParameter("@Dosis", SqlDbType.NVarChar))
                    consulta.Parameters("@Dosis").Value = objMedicamentos.dosis
                    consulta.Parameters.Add(New SqlParameter("@Frecuencia", SqlDbType.NVarChar))
                    consulta.Parameters("@Frecuencia").Value = objMedicamentos.frecuencia
                    consulta.Parameters.Add(New SqlParameter("@Num_Dosis", SqlDbType.NVarChar))
                    consulta.Parameters("@Num_Dosis").Value = objMedicamentos.ndosis
                    consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.NVarChar))
                    consulta.Parameters("@Usuario_Creacion").Value = objMedicamentos.usuarioInforme
                    objMedicamentos.Codigo = CType(consulta.ExecuteScalar, Integer)
                    consulta.Parameters.Clear()
                    consulta.CommandText = "[SP_SOLICITUD_MEDICAMENTOS_DETALLE_ELIMINAR_RR]"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                    consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo

                    consulta.ExecuteNonQuery()
                    For i As Int32 = 0 To dtMedicamento.Rows.Count - 1
                        consulta.Parameters.Clear()
                        consulta.CommandText = "SP_SOLICITUD_MEDICAMENTOS_DETALLE_CREAR_RR"
                        consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = dtMedicamento.Rows(i).Item(0).ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EVO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EVO").Value = dtMedicamento.Rows(i).Item(2).ToString
                        consulta.ExecuteNonQuery()
                    Next

                    consulta.Parameters.Clear()
                    consulta.CommandText = "SP_CTC_AUTOMATICO_RR"
                    consulta.Parameters.Add(New SqlParameter("@pSolicitud", SqlDbType.Int)).Value = objMedicamentos.Codigo
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

    Public Shared Sub actualizarMedicamentosNOPOSAF(objMedicamentos As SolicitudMedicamentosNoPos, dtMedicamento As DataTable)


        Try

            Using consulta As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Parameters.Clear()
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.CommandText = "SP_SOLICITUD_MEDICAMENTOS_ACTUALIZAR_RR"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                    consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                    consulta.Parameters.Add(New SqlParameter("@Fecha_Solicitud", SqlDbType.DateTime))
                    consulta.Parameters("@Fecha_Solicitud").Value = objMedicamentos.fechaSolicitud
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Interno1", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Interno1").Value = objMedicamentos.codigoEquivalenciaUno
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Interno2", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Interno2").Value = objMedicamentos.codigoEquivalenciaDos
                    consulta.Parameters.Add(New SqlParameter("@Resumen_hc", SqlDbType.NVarChar))
                    consulta.Parameters("@Resumen_hc").Value = objMedicamentos.resumen
                    consulta.Parameters.Add(New SqlParameter("@Si", SqlDbType.Bit))
                    consulta.Parameters("@Si").Value = objMedicamentos.rdsi
                    consulta.Parameters.Add(New SqlParameter("@No", SqlDbType.Bit))
                    consulta.Parameters("@No").Value = objMedicamentos.rdno
                    consulta.Parameters.Add(New SqlParameter("@RespuestaPOS", SqlDbType.NVarChar))
                    consulta.Parameters("@RespuestaPOS").Value = objMedicamentos.respuesta
                    consulta.Parameters.Add(New SqlParameter("@Efecto_Deseado", SqlDbType.NVarChar))
                    consulta.Parameters("@Efecto_Deseado").Value = objMedicamentos.efecto
                    consulta.Parameters.Add(New SqlParameter("@Iniciacion_Efecto", SqlDbType.NVarChar))
                    consulta.Parameters("@Iniciacion_Efecto").Value = objMedicamentos.inicioefecto
                    consulta.Parameters.Add(New SqlParameter("@Explicar_Razon", SqlDbType.NVarChar))
                    consulta.Parameters("@Explicar_Razon").Value = objMedicamentos.explicarazon
                    consulta.Parameters.Add(New SqlParameter("@Criterio1", SqlDbType.NVarChar))
                    consulta.Parameters("@Criterio1").Value = objMedicamentos.cbocriterio
                    consulta.Parameters.Add(New SqlParameter("@Criterio2", SqlDbType.NVarChar))
                    consulta.Parameters("@Criterio2").Value = objMedicamentos.cbocriterio2
                    consulta.Parameters.Add(New SqlParameter("@PreCoEsTox", SqlDbType.NVarChar))
                    consulta.Parameters("@PreCoEsTox").Value = objMedicamentos.precaucion
                    consulta.Parameters.Add(New SqlParameter("@PeligroPaciente", SqlDbType.NVarChar))
                    consulta.Parameters("@PeligroPaciente").Value = objMedicamentos.pelipaciente
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeutica", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeutica").Value = objMedicamentos.cboposibilidad
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeuticaPOS", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeuticaPOS").Value = objMedicamentos.reacciones
                    consulta.Parameters.Add(New SqlParameter("@PosibilidadTerapeuticaPOScontra", SqlDbType.NVarChar))
                    consulta.Parameters("@PosibilidadTerapeuticaPOScontra").Value = objMedicamentos.contraindicaciones
                    consulta.Parameters.Add(New SqlParameter("@Duracion_Tto", SqlDbType.NVarChar))
                    consulta.Parameters("@Duracion_Tto").Value = objMedicamentos.duraciontratamiento
                    consulta.Parameters.Add(New SqlParameter("@Dosis", SqlDbType.NVarChar))
                    consulta.Parameters("@Dosis").Value = objMedicamentos.dosis
                    consulta.Parameters.Add(New SqlParameter("@Frecuencia", SqlDbType.NVarChar))
                    consulta.Parameters("@Frecuencia").Value = objMedicamentos.frecuencia
                    consulta.Parameters.Add(New SqlParameter("@Num_Dosis", SqlDbType.NVarChar))
                    consulta.Parameters("@Num_Dosis").Value = objMedicamentos.ndosis
                    consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.NVarChar))
                    consulta.Parameters("@Usuario").Value = objMedicamentos.usuarioInforme

                    consulta.ExecuteNonQuery()
                    consulta.Parameters.Clear()
                    consulta.CommandText = "[SP_SOLICITUD_MEDICAMENTOS_DETALLE_ELIMINAR_RR]"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                    consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                    consulta.ExecuteNonQuery()
                    For i As Int32 = 0 To dtMedicamento.Rows.Count - 1
                        consulta.Parameters.Clear()
                        consulta.CommandText = "SP_SOLICITUD_MEDICAMENTOS_DETALLE_CREAR_RR"
                        consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO").Value = objMedicamentos.Codigo
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = dtMedicamento.Rows(i).Item(0).ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EVO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EVO").Value = dtMedicamento.Rows(i).Item(3).ToString
                        consulta.ExecuteNonQuery()
                    Next

                    consulta.Parameters.Clear()
                    consulta.CommandText = "SP_CTC_AUTOMATICO_RR"
                    consulta.Parameters.Add(New SqlParameter("@pSolicitud", SqlDbType.Int)).Value = objMedicamentos.Codigo
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
