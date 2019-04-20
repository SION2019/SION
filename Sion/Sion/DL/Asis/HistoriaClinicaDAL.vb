Imports System.Data.SqlClient
Public Class HistoriaClinicaDAL
    Public Shared Sub cargarMedicamentosAutomatico(ByRef objOrden As OrdenMedica)
        Try
            ' Create the SelectCommand.
            Dim command As SqlCommand = New SqlCommand()
            command.Connection = FormPrincipal.cnxion
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = objOrden.medicamentoAutoCargar
            command.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrden.registro
            command.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.DateTime)).Value = objOrden.fechaOrden
            command.Parameters.Add(New SqlParameter("@Tabla", SqlDbType.Structured)).Value = objOrden.dtMedicamentos

            Dim adapter As SqlDataAdapter = New SqlDataAdapter(command)

            adapter.Fill(objOrden.dtMedicamentosAuto)

        Catch ex As SqlException
            Throw ex
        End Try
    End Sub
    Public Shared Sub cargarProcedimientosAutomatico(ByRef objOrden As OrdenMedica)
        Try
            ' Create the SelectCommand.
            Dim command As SqlCommand = New SqlCommand()
            command.Connection = FormPrincipal.cnxion
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = objOrden.procedimientoAutoCargar
            command.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrden.registro
            command.Parameters.Add(New SqlParameter("@ORDEN", SqlDbType.NVarChar)).Value = objOrden.codigoOrden
            command.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.DateTime)).Value = objOrden.fechaOrden
            command.Parameters.Add(New SqlParameter("@Tabla", SqlDbType.Structured)).Value = objOrden.dtProcedimientos

            Dim adapter As SqlDataAdapter = New SqlDataAdapter(command)

            adapter.Fill(objOrden.dtProcedimientosAuto)

        Catch ex As SqlException
            Throw ex
        End Try
    End Sub
    Public Shared Sub cargarMedicamentosAuditoria(ByRef objOrden As OrdenMedica)
        Try
            ' Create the SelectCommand.
            Dim command As SqlCommand = New SqlCommand()
            command.Connection = FormPrincipal.cnxion
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = objOrden.medicamentoAutoCargar
            command.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrden.registro
            command.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.DateTime)).Value = objOrden.fechaOrden
            command.Parameters.Add(New SqlParameter("@Tabla", SqlDbType.Structured)).Value = objOrden.dtMedicamentos

            Dim adapter As SqlDataAdapter = New SqlDataAdapter(command)

            adapter.Fill(objOrden.dtMedicamentosAuto)

        Catch ex As SqlException
            Throw ex
        End Try
    End Sub
    Public Shared Function guardarOrdenMedica(objOrdenMedica As OrdenMedica, ByRef paramsDgv As List(Of DataGridView), paramsDt As List(Of DataTable)) As String
        Dim respuesta As String ''esto lo uso para poder identificar el error actual en produccion
        Try

            Dim dtEstancias, dtComidas, dtOxigenos, dtParaclinicos, dtHemoderivados,
            dtGlucometrias, dtProcedimientos, dtMedicamentos, dtImpregnaciones As New DataTable
            dtEstancias = paramsDt(0).Copy
            dtComidas = paramsDt(1).Copy
            dtOxigenos = paramsDt(2).Copy
            dtParaclinicos = paramsDt(3).Copy
            dtHemoderivados = paramsDt(4).Copy
            dtGlucometrias = paramsDt(5).Copy
            dtProcedimientos = paramsDt(6).Copy
            dtMedicamentos = paramsDt(7).Copy
            dtImpregnaciones = paramsDt(8).Copy
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objOrdenMedica.elementoAEliminar
                        If sentencia <> "" Then
                            Try
                                consulta.CommandText = sentencia
                                consulta.ExecuteNonQuery()
                            Catch ex As Exception
                                respuesta = sentencia
                                Throw ex
                            End Try

                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure
                    '--------------------------------------------
                    ''---------------GUARDAR O ACTUALIZA ORDEN MEDICA
                    '--------------------------------------------
                    If objOrdenMedica.codigoOrden <> "" Then
                        consulta.Parameters.Clear()
                        If objOrdenMedica.modulo = Constantes.HC Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_ACTUALIZAR]"
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        ElseIf objOrdenMedica.modulo = Constantes.AM Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_ACTUALIZAR_R]"
                            consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = 1
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        ElseIf objOrdenMedica.modulo = Constantes.AF Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_ACTUALIZAR_RR]"
                            consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = 1
                        End If
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                        consulta.Parameters.Add(New SqlParameter("@FECHA_ORDEN", SqlDbType.DateTime)).Value = objOrdenMedica.fechaOrden
                        consulta.Parameters.Add(New SqlParameter("@Usuario_r", SqlDbType.Int)).Value = objOrdenMedica.usuarioReal
                        consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int)).Value = objOrdenMedica.usuario
                        consulta.ExecuteNonQuery()
                    Else
                        consulta.Parameters.Clear()
                        If objOrdenMedica.modulo = Constantes.HC Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_CREAR]"
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        ElseIf objOrdenMedica.modulo = Constantes.AM Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_CREAR_R]"
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = -1
                            consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = 0
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        ElseIf objOrdenMedica.modulo = Constantes.AF Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_CREAR_RR]"
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = -1
                            consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = 0
                        End If
                        consulta.Parameters.Add(New SqlParameter("@FECHA_ORDEN", SqlDbType.DateTime)).Value = objOrdenMedica.fechaOrden
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrdenMedica.registro
                        consulta.Parameters.Add(New SqlParameter("@Usuario_r", SqlDbType.Int)).Value = objOrdenMedica.usuarioReal
                        consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int)).Value = objOrdenMedica.usuario
                        objOrdenMedica.codigoOrden = CType(consulta.ExecuteScalar, String)
                    End If
                    '--------------------------------------------
                    ''---------------GUARDAR DETALLE ESTANCIAS
                    '--------------------------------------------
                    For I = 0 To dtEstancias.Rows.Count - 1
                        If dtEstancias.Rows(I).Item("N_Registro") = objOrdenMedica.registro Then
                            consulta.Parameters.Clear()
                            If objOrdenMedica.modulo = Constantes.HC Then
                                consulta.CommandText = "[SP_ORDEN_MEDICA_ESTANCIA_GUARDAR]"
                            ElseIf objOrdenMedica.modulo = Constantes.AM Then
                                consulta.CommandText = "[SP_ORDEN_MEDICA_ESTANCIA_GUARDAR_R]"
                            ElseIf objOrdenMedica.modulo = Constantes.AF Then
                                consulta.CommandText = "[SP_ORDEN_MEDICA_ESTANCIA_GUARDAR_RR]"
                            End If
                            Dim FECHA As Date
                            FECHA = CDate(dtEstancias.Rows(I).Item(2).ToString).Date
                            consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrdenMedica.registro
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_CUPS", SqlDbType.NVarChar)).Value = dtEstancias.Rows(I).Item("Código").ToString
                            consulta.Parameters.Add(New SqlParameter("@DIA", SqlDbType.Date)).Value = CDate(dtEstancias.Rows(I).Item("Dia").ToString).Date
                            consulta.Parameters.Add(New SqlParameter("@JUSTIFICACION", SqlDbType.NVarChar)).Value = dtEstancias.Rows(I).Item("Justificación").ToString
                            consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                            If objOrdenMedica.modulo = Constantes.HC Or objOrdenMedica.modulo = Constantes.AM Then
                                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                            End If
                            consulta.ExecuteNonQuery()
                        End If

                    Next
                    '--------------------------------------------
                    ''---------------GUARDAR DETALLE INDICACIONES
                    '--------------------------------------------
                    If objOrdenMedica.indicacion <> "" Then
                        consulta.Parameters.Clear()
                        If objOrdenMedica.modulo = Constantes.HC Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_INDICACION_GUARDAR]"
                        ElseIf objOrdenMedica.modulo = Constantes.AM Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_INDICACION_GUARDAR_R]"
                        ElseIf objOrdenMedica.modulo = Constantes.AF Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_INDICACION_GUARDAR_RR]"
                        End If
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar)).Value = objOrdenMedica.indicacion
                        consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                        If objOrdenMedica.modulo = Constantes.HC Or objOrdenMedica.modulo = Constantes.AM Then
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        End If
                        consulta.ExecuteNonQuery()
                    End If
                    '--------------------------------------------
                    ''---------------GUARDAR DETALLE COMIDAS
                    '--------------------------------------------
                    For I = 0 To dtComidas.Rows.Count - 1
                        consulta.Parameters.Clear()
                        If objOrdenMedica.modulo = Constantes.HC Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_COMIDA_GUARDAR]"
                        ElseIf objOrdenMedica.modulo = Constantes.AM Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_COMIDA_GUARDAR_R]"
                        ElseIf objOrdenMedica.modulo = Constantes.AF Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_COMIDA_GUARDAR_RR]"
                        End If
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                        consulta.Parameters.Add(New SqlParameter("@DESAYUNO", SqlDbType.NVarChar)).Value = If(dtComidas.Rows(I).Item("Desayuno") = False, "", dtComidas.Rows(I).Item("Código").ToString)
                        consulta.Parameters.Add(New SqlParameter("@ALMUERZO", SqlDbType.NVarChar)).Value = If(dtComidas.Rows(I).Item("Almuerzo") = False, "", dtComidas.Rows(I).Item("Código").ToString)
                        consulta.Parameters.Add(New SqlParameter("@CENA", SqlDbType.NVarChar)).Value = If(dtComidas.Rows(I).Item("Cena") = False, "", dtComidas.Rows(I).Item("Código").ToString)
                        consulta.Parameters.Add(New SqlParameter("@JUSTIFICACION", SqlDbType.NVarChar)).Value = dtComidas.Rows(I).Item("Justificación").ToString
                        consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                        If objOrdenMedica.modulo = Constantes.HC Or objOrdenMedica.modulo = Constantes.AM Then
                            If objOrdenMedica.modulo = Constantes.HC Then
                                consulta.Parameters.Add(New SqlParameter("@VIRTUAL", SqlDbType.Bit)).Value = dtComidas.Rows(I).Item("Virtual")
                            End If
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        End If
                        consulta.ExecuteNonQuery()
                    Next
                    '--------------------------------------------
                    ''---------------GUARDAR DETALLE OXIGENOS
                    '--------------------------------------------

                    consulta.Parameters.Clear()
                    If objOrdenMedica.modulo = Constantes.HC Then
                        consulta.CommandText = "[SP_ORDEN_MEDICA_OXIGENO_GUARDAR]"
                    ElseIf objOrdenMedica.modulo = Constantes.AM Then
                        consulta.CommandText = "[SP_ORDEN_MEDICA_OXIGENO_GUARDAR_R]"
                    ElseIf objOrdenMedica.modulo = Constantes.AF Then
                        consulta.CommandText = "[SP_ORDEN_MEDICA_OXIGENO_GUARDAR_RR]"
                    End If
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                    consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                    consulta.Parameters.Add(New SqlParameter("@pTabla", SqlDbType.Structured)).Value = objOrdenMedica.dtOxigenos
                    consulta.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.Date)).Value = objOrdenMedica.fechaOrden
                    If objOrdenMedica.modulo = Constantes.HC Or objOrdenMedica.modulo = Constantes.AM Then
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                    End If
                    consulta.ExecuteNonQuery()

                    '--------------------------------------------
                    ''---------------GUARDAR DETALLE PARACLINICOS
                    '--------------------------------------------
                    For I = 0 To dtParaclinicos.Rows.Count - 1
                        consulta.Parameters.Clear()
                        If objOrdenMedica.modulo = Constantes.HC Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_PARACLINICO_GUARDAR]"
                        ElseIf objOrdenMedica.modulo = Constantes.AM Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_PARACLINICO_GUARDAR_R]"
                        ElseIf objOrdenMedica.modulo = Constantes.AF Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_PARACLINICO_GUARDAR_RR]"
                        End If
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrdenMedica.registro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CUPS", SqlDbType.NVarChar)).Value = dtParaclinicos.Rows(I).Item("Código").ToString
                        consulta.Parameters.Add(New SqlParameter("@CANTIDAD", SqlDbType.Int)).Value = dtParaclinicos.Rows(I).Item("Cantidad").ToString
                        consulta.Parameters.Add(New SqlParameter("@INDICACION", SqlDbType.NVarChar)).Value = dtParaclinicos.Rows(I).Item("Indicación").ToString
                        consulta.Parameters.Add(New SqlParameter("@JUSTIFICACION", SqlDbType.NVarChar)).Value = dtParaclinicos.Rows(I).Item("Justificación").ToString
                        consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                        If objOrdenMedica.modulo = Constantes.HC Or objOrdenMedica.modulo = Constantes.AM Then
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        End If
                        consulta.ExecuteNonQuery()
                    Next
                    '--------------------------------------------
                    ''---------------GUARDAR DETALLE HEMODERIVADOS
                    '--------------------------------------------
                    For I = 0 To dtHemoderivados.Rows.Count - 1
                        consulta.Parameters.Clear()
                        If objOrdenMedica.modulo = Constantes.HC Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_HEMODERIVADO_GUARDAR]"
                        ElseIf objOrdenMedica.modulo = Constantes.AM Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_HEMODERIVADO_GUARDAR_R]"
                        ElseIf objOrdenMedica.modulo = Constantes.AF Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_HEMODERIVADO_GUARDAR_RR]"
                        End If
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrdenMedica.registro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CUPS", SqlDbType.NVarChar)).Value = dtHemoderivados.Rows(I).Item("Código").ToString
                        consulta.Parameters.Add(New SqlParameter("@CANTIDAD", SqlDbType.Int)).Value = dtHemoderivados.Rows(I).Item("Cantidad").ToString
                        consulta.Parameters.Add(New SqlParameter("@INDICACION", SqlDbType.NVarChar)).Value = dtHemoderivados.Rows(I).Item("Indicación").ToString
                        consulta.Parameters.Add(New SqlParameter("@JUSTIFICACION", SqlDbType.NVarChar)).Value = dtHemoderivados.Rows(I).Item("Justificación").ToString
                        consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                        If objOrdenMedica.modulo = Constantes.HC Or objOrdenMedica.modulo = Constantes.AM Then
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        End If
                        consulta.ExecuteNonQuery()
                    Next
                    '--------------------------------------------
                    ''---------------GUARDAR DETALLE GLUCOMETRIAS
                    '--------------------------------------------
                    For I = 0 To dtGlucometrias.Rows.Count - 1
                        consulta.Parameters.Clear()
                        If objOrdenMedica.modulo = Constantes.HC Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_GLUCOMETRIA_GUARDAR]"
                        ElseIf objOrdenMedica.modulo = Constantes.AM Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_GLUCOMETRIA_GUARDAR_R]"
                        ElseIf objOrdenMedica.modulo = Constantes.AF Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_GLUCOMETRIA_GUARDAR_RR]"
                        End If
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrdenMedica.registro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CUPS", SqlDbType.NVarChar)).Value = dtGlucometrias.Rows(I).Item("Código").ToString
                        consulta.Parameters.Add(New SqlParameter("@CANTIDAD", SqlDbType.Int)).Value = CInt(dtGlucometrias.Rows(I).Item("Cantidad").ToString)
                        consulta.Parameters.Add(New SqlParameter("@INICIO", SqlDbType.Int)).Value = CInt(IIf(dtGlucometrias.Rows(I).Item("Inicio").ToString = Constantes.ITEM_dgv_SELECCIONE, Constantes.ITEM_DGV_SELECCIONE_VALOR, dtGlucometrias.Rows(I).Item("Inicio").ToString))
                        consulta.Parameters.Add(New SqlParameter("@FRECUENCIA", SqlDbType.Int)).Value = CInt(dtGlucometrias.Rows(I).Item("Frecuencia").ToString)
                        consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                        If objOrdenMedica.modulo = Constantes.HC Or objOrdenMedica.modulo = Constantes.AM Then
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        End If
                        consulta.ExecuteNonQuery()
                    Next
                    '--------------------------------------------
                    ''---------------GUARDAR DETALLE PROCEDIMIENTOS
                    '--------------------------------------------
                    For I = 0 To dtProcedimientos.Rows.Count - 1
                        consulta.Parameters.Clear()
                        If objOrdenMedica.modulo = Constantes.HC Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_PROCEDIMIENTO_GUARDAR]"
                        ElseIf objOrdenMedica.modulo = Constantes.AM Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_PROCEDIMIENTO_GUARDAR_R]"
                        ElseIf objOrdenMedica.modulo = Constantes.AF Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_PROCEDIMIENTO_GUARDAR_RR]"
                        End If
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrdenMedica.registro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CUPS", SqlDbType.NVarChar)).Value = dtProcedimientos.Rows(I).Item("Código").ToString
                        consulta.Parameters.Add(New SqlParameter("@CANTIDAD", SqlDbType.Int)).Value = dtProcedimientos.Rows(I).Item("Cantidad").ToString
                        consulta.Parameters.Add(New SqlParameter("@INDICACION", SqlDbType.NVarChar)).Value = dtProcedimientos.Rows(I).Item("Indicación").ToString
                        consulta.Parameters.Add(New SqlParameter("@JUSTIFICACION", SqlDbType.NVarChar)).Value = dtProcedimientos.Rows(I).Item("Justificación").ToString
                        consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                        If objOrdenMedica.modulo = Constantes.HC Or objOrdenMedica.modulo = Constantes.AM Then
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        End If
                        consulta.ExecuteNonQuery()
                    Next
                    '--------------------------------------------
                    ''---------------GUARDAR DETALLE MEDICAMENTOS
                    '--------------------------------------------
                    For I = 0 To dtMedicamentos.Rows.Count - 1
                        consulta.Parameters.Clear()
                        If objOrdenMedica.modulo = Constantes.HC Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_MEDICAMENTO_GUARDAR]"
                        ElseIf objOrdenMedica.modulo = Constantes.AM Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_MEDICAMENTO_GUARDAR_R]"
                        ElseIf objOrdenMedica.modulo = Constantes.AF Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_MEDICAMENTO_GUARDAR_RR]"
                        End If
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrdenMedica.registro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_INTERNO", SqlDbType.Int)).Value = CInt(dtMedicamentos.Rows(I).Item("Código").ToString)
                        consulta.Parameters.Add(New SqlParameter("@DOSIS", SqlDbType.Float)).Value = CDbl(dtMedicamentos.Rows(I).Item("Dosis").ToString)
                        consulta.Parameters.Add(New SqlParameter("@VIA", SqlDbType.Int)).Value = CInt(dtMedicamentos.Rows(I).Item("CódigoVia").ToString)
                        consulta.Parameters.Add(New SqlParameter("@HORARIO", SqlDbType.Int)).Value = CInt(IIf(dtMedicamentos.Rows(I).Item("Horario").ToString = Constantes.ITEM_dgv_SELECCIONE,
                                                                                                              Constantes.ITEM_DGV_SELECCIONE_VALOR,
                                                                                                              IIf(dtMedicamentos.Rows(I).Item("Horario").ToString = Constantes.ITEM_dgv_AH,
                                                                                                              Constantes.ITEM_dgv_AH_VALOR,
                                                                                                                  IIf(dtMedicamentos.Rows(I).Item("Horario").ToString = Constantes.ITEM_dgv_RN,
                                                                                                                  Constantes.ITEM_dgv_RN_VALOR,
                                                                                                                    IIf(dtMedicamentos.Rows(I).Item("Horario").ToString = Constantes.ITEM_dgv_CA,
                                                                                                                    Constantes.ITEM_dgv_CA_VALOR,
                                                                                                                    dtMedicamentos.Rows(I).Item("Horario").ToString)))))
                        consulta.Parameters.Add(New SqlParameter("@INICIO", SqlDbType.Int)).Value = CInt(IIf(dtMedicamentos.Rows(I).Item("Inicio").ToString = Constantes.ITEM_dgv_SELECCIONE, Constantes.ITEM_DGV_SELECCIONE_VALOR, dtMedicamentos.Rows(I).Item("Inicio").ToString))
                        consulta.Parameters.Add(New SqlParameter("@FRECUENCIA", SqlDbType.Int)).Value = CInt(dtMedicamentos.Rows(I).Item("Frecuencia").ToString)
                        consulta.Parameters.Add(New SqlParameter("@SUSPENDIDO", SqlDbType.Bit)).Value = dtMedicamentos.Rows(I).Item("Suspender")
                        consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                        If objOrdenMedica.modulo = Constantes.HC Or objOrdenMedica.modulo = Constantes.AM Then
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        End If
                        consulta.ExecuteNonQuery()
                    Next
                    '--------------------------------------------
                    ''---------------GUARDAR DETALLE IMPREGNACIONES
                    '--------------------------------------------
                    For I = 0 To dtImpregnaciones.Rows.Count - 1
                        consulta.Parameters.Clear()
                        If objOrdenMedica.modulo = Constantes.HC Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_IMPREGNACION_GUARDAR]"
                        ElseIf objOrdenMedica.modulo = Constantes.AM Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_IMPREGNACION_GUARDAR_R]"
                        ElseIf objOrdenMedica.modulo = Constantes.AF Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_IMPREGNACION_GUARDAR_RR]"
                        End If
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrdenMedica.registro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_INTERNO", SqlDbType.Int)).Value = CInt(dtImpregnaciones.Rows(I).Item("Código").ToString)
                        consulta.Parameters.Add(New SqlParameter("@DOSIS", SqlDbType.Float)).Value = CDbl(dtImpregnaciones.Rows(I).Item("Dosis").ToString)
                        consulta.Parameters.Add(New SqlParameter("@VELOCIDAD", SqlDbType.Float)).Value = CDbl(dtImpregnaciones.Rows(I).Item("Velocidad").ToString)
                        consulta.Parameters.Add(New SqlParameter("@DISOLVENTE", SqlDbType.Int)).Value = CInt(dtImpregnaciones.Rows(I).Item("CódigoDisolvente").ToString)
                        consulta.Parameters.Add(New SqlParameter("@INICIO", SqlDbType.Int)).Value = CInt(IIf(dtImpregnaciones.Rows(I).Item("Inicio").ToString = Constantes.ITEM_dgv_SELECCIONE, Constantes.ITEM_DGV_SELECCIONE_VALOR, dtImpregnaciones.Rows(I).Item("Inicio").ToString))
                        consulta.Parameters.Add(New SqlParameter("@CANTIDAD", SqlDbType.NVarChar)).Value = dtImpregnaciones.Rows(I).Item("Cantidad").ToString
                        consulta.Parameters.Add(New SqlParameter("@SUSPENDIDO", SqlDbType.Bit)).Value = dtImpregnaciones.Rows(I).Item("Suspender")
                        consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                        If objOrdenMedica.modulo = Constantes.HC Or objOrdenMedica.modulo = Constantes.AM Then
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        End If
                        consulta.ExecuteNonQuery()
                    Next
                    '--------------------------------------------
                    ''---------------GUARDAR DETALLE INFUSIONES
                    '--------------------------------------------
                    Dim pConsecutivo As Integer
                    Dim NombreTablaMezcla As String
                    For I = 0 To paramsDgv(0).Rows.Count - 2
                        consulta.Parameters.Clear()
                        If objOrdenMedica.modulo = Constantes.HC Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_INFUSION_GUARDAR]"
                        ElseIf objOrdenMedica.modulo = Constantes.AM Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_INFUSION_GUARDAR_R]"
                        ElseIf objOrdenMedica.modulo = Constantes.AF Then
                            consulta.CommandText = "[SP_ORDEN_MEDICA_INFUSION_GUARDAR_RR]"
                        End If
                        If paramsDgv(0).DataSource.Rows(I).Item("Consecutivo").ToString = "" Then
                            pConsecutivo = -1
                        Else
                            pConsecutivo = paramsDgv(0).DataSource.Rows(I).Item("Consecutivo").ToString
                        End If
                        consulta.Parameters.Add(New SqlParameter("@CONSECUTIVO", SqlDbType.Int)).Value = pConsecutivo
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrdenMedica.registro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_INTERNO", SqlDbType.Int)).Value = CInt(paramsDgv(0).DataSource.Rows(I).Item("Código").ToString)
                        consulta.Parameters.Add(New SqlParameter("@DOSIS", SqlDbType.Float)).Value = CDbl(paramsDgv(0).DataSource.Rows(I).Item("Dosis").ToString)
                        consulta.Parameters.Add(New SqlParameter("@VELOCIDAD", SqlDbType.Float)).Value = CDbl(paramsDgv(0).DataSource.Rows(I).Item("Velocidad").ToString)
                        consulta.Parameters.Add(New SqlParameter("@DISOLVENTE", SqlDbType.Int)).Value = CInt(paramsDgv(0).DataSource.Rows(I).Item("CódigoDisolvente").ToString)
                        consulta.Parameters.Add(New SqlParameter("@INICIO", SqlDbType.Int)).Value = CInt(IIf(paramsDgv(0).DataSource.Rows(I).Item("Inicio").ToString = Constantes.ITEM_dgv_SELECCIONE, Constantes.ITEM_DGV_SELECCIONE_VALOR, paramsDgv(0).DataSource.Rows(I).Item("Inicio").ToString))
                        consulta.Parameters.Add(New SqlParameter("@TOTAL_DOSIS", SqlDbType.NVarChar)).Value = paramsDgv(0).DataSource.Rows(I).Item("Dosis").ToString
                        consulta.Parameters.Add(New SqlParameter("@SUSPENDIDO", SqlDbType.Bit)).Value = paramsDgv(0).DataSource.Rows(I).Item("Suspender")
                        consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                        If objOrdenMedica.modulo = Constantes.HC Or objOrdenMedica.modulo = Constantes.AM Then
                            consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                        End If
                        pConsecutivo = CType(consulta.ExecuteScalar, String)
                        '--------------------------------------------
                        ''---------------GUARDAR DETALLE MEZCLAS SI EXISTEN
                        '--------------------------------------------
                        NombreTablaMezcla = HistoriaClinicaBLL.nombreTabla(paramsDgv(0), paramsDgv(0).DataSource, I)
                        If objOrdenMedica.dsMezcla.Tables.Contains(NombreTablaMezcla) = True Then
                            For J = 0 To objOrdenMedica.dsMezcla.Tables(NombreTablaMezcla).Rows.Count - 1
                                consulta.Parameters.Clear()
                                If objOrdenMedica.modulo = Constantes.HC Then
                                    consulta.CommandText = "[SP_ORDEN_MEDICA_MEZCLA_GUARDAR]"
                                ElseIf objOrdenMedica.modulo = Constantes.AM Then
                                    consulta.CommandText = "[SP_ORDEN_MEDICA_MEZCLA_GUARDAR_R]"
                                ElseIf objOrdenMedica.modulo = Constantes.AF Then
                                    consulta.CommandText = "[SP_ORDEN_MEDICA_MEZCLA_GUARDAR_RR]"
                                End If
                                consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                                consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrdenMedica.registro
                                consulta.Parameters.Add(New SqlParameter("@CONSECUTIVO", SqlDbType.Int)).Value = pConsecutivo
                                consulta.Parameters.Add(New SqlParameter("@CODIGO_INTERNO", SqlDbType.Int)).Value = objOrdenMedica.dsMezcla.Tables(NombreTablaMezcla).Rows(J).Item("Código").ToString
                                consulta.Parameters.Add(New SqlParameter("@DOSIS", SqlDbType.Float)).Value = objOrdenMedica.dsMezcla.Tables(NombreTablaMezcla).Rows(J).Item("Dosis").ToString
                                consulta.Parameters.Add(New SqlParameter("@Usuario", SqlDbType.Int)).Value = objOrdenMedica.usuario
                                If objOrdenMedica.modulo = Constantes.HC Or objOrdenMedica.modulo = Constantes.AM Then
                                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOrdenMedica.codigoEP
                                End If
                                consulta.ExecuteNonQuery()
                            Next
                        End If

                    Next
                    'If HistoriaClinicaBLL.GetNumeroMedicamentosOrdenadosNutricion(objOrdenMedica.objetoNutricion.dtNutricion) > 0 Then
                    consulta.Parameters.Clear()
                    If objOrdenMedica.modulo = Constantes.HC Then
                        consulta.CommandText = "[SP_ORDEN_MEDICA_NUTRICION_GUARDAR]"
                    ElseIf objOrdenMedica.modulo = Constantes.AM Then
                        consulta.CommandText = "[SP_ORDEN_MEDICA_NUTRICION_GUARDAR_R]"
                    ElseIf objOrdenMedica.modulo = Constantes.AF Then
                        consulta.CommandText = "[SP_ORDEN_MEDICA_NUTRICION_GUARDAR_RR]"
                    End If

                    consulta.Parameters.Add(New SqlParameter("@pCodigo_Orden", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                    consulta.Parameters.Add(New SqlParameter("@pPESO", SqlDbType.Float)).Value = objOrdenMedica.objetoNutricion.peso
                    consulta.Parameters.Add(New SqlParameter("@pHora_inicial", SqlDbType.Float)).Value = objOrdenMedica.objetoNutricion.horaInicial
                    consulta.Parameters.Add(New SqlParameter("@pTasa_Hidrica", SqlDbType.Float)).Value = objOrdenMedica.objetoNutricion.TasaHidrica
                    consulta.Parameters.Add(New SqlParameter("@pFator_ajustes", SqlDbType.Float)).Value = objOrdenMedica.objetoNutricion.FactorAjustes
                    consulta.Parameters.Add(New SqlParameter("@pAlimentacion", SqlDbType.Float)).Value = objOrdenMedica.objetoNutricion.alimentacion
                    consulta.Parameters.Add(New SqlParameter("@pMedicacion", SqlDbType.Float)).Value = objOrdenMedica.objetoNutricion.medicacion
                    consulta.Parameters.Add(New SqlParameter("@potros", SqlDbType.Float)).Value = objOrdenMedica.objetoNutricion.otros
                    consulta.Parameters.Add(New SqlParameter("@pTHT", SqlDbType.Float)).Value = objOrdenMedica.objetoNutricion.tht
                    consulta.Parameters.Add(New SqlParameter("@pVolumen", SqlDbType.Float)).Value = objOrdenMedica.objetoNutricion.volumen
                    consulta.Parameters.Add(New SqlParameter("@pRazon", SqlDbType.Float)).Value = objOrdenMedica.objetoNutricion.razon
                    consulta.Parameters.Add(New SqlParameter("@pUsuario_Creacion", SqlDbType.Int)).Value = objOrdenMedica.usuario
                    consulta.Parameters.Add(New SqlParameter("@pCodigoInternoAguaDestilada", SqlDbType.NVarChar)).Value = objOrdenMedica.objetoNutricion.codigoAgua
                    consulta.Parameters.Add(New SqlParameter("@pCantidadAguaDestilada", SqlDbType.Float)).Value = objOrdenMedica.objetoNutricion.cantidadAgua
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Float)).Value = objOrdenMedica.codigoEP
                    consulta.Parameters.Add(New SqlParameter("@pTabla", SqlDbType.Structured)).Value = objOrdenMedica.objetoNutricion.dtNutricion
                    consulta.ExecuteNonQuery()
                    ' End If
                    consulta.Parameters.Clear()
                    If objOrdenMedica.modulo = Constantes.HC Then
                        consulta.CommandText = "[SP_ORDEN_MEDICA_CONFIG_AUDITORIA_GUARDAR]"
                    ElseIf objOrdenMedica.modulo = Constantes.AM Then
                        consulta.CommandText = "[SP_ORDEN_MEDICA_CONFIG_AUDITORIA_GUARDAR_R]"
                    ElseIf objOrdenMedica.modulo = Constantes.AF Then
                        consulta.CommandText = "[SP_ORDEN_MEDICA_CONFIG_AUDITORIA_GUARDAR_RR]"
                    End If
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Orden", SqlDbType.Int)).Value = objOrdenMedica.codigoOrden
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrdenMedica.registro
                    consulta.ExecuteNonQuery()
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception

                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

            guardarSolicitudes(objOrdenMedica, paramsDgv)
        Catch ex As Exception
            Return respuesta
            Throw ex

        End Try
    End Function
    Public Shared Sub agregaMedicamentoAutomaticoAuditoria(objOrden As OrdenMedica, dtMezcla As DataTable)
        Try
            If objOrden.dtParaclinicos.Columns.Count < 8 Then
                objOrden.dtParaclinicos.Columns.Add("CodigoTipoExamen")
            End If
            ' Create the SelectCommand.
            Dim command As SqlCommand = New SqlCommand()
            command.Connection = FormPrincipal.cnxion
            command.CommandType = CommandType.StoredProcedure
            command.CommandText = objOrden.medicamentoAuditoriaCargar
            command.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int)).Value = objOrden.registro
            command.Parameters.Add(New SqlParameter("@FECHA", SqlDbType.DateTime)).Value = objOrden.fechaOrden
            command.Parameters.Add(New SqlParameter("@H_ORDEN_MEDICA_HEMODERIVADO", SqlDbType.Structured)).Value = objOrden.dtHemoderivados
            command.Parameters.Add(New SqlParameter("@H_ORDEN_MEDICA_EXAMEN", SqlDbType.Structured)).Value = objOrden.dtParaclinicos
            command.Parameters.Add(New SqlParameter("@H_ORDEN_MEDICA_D_PROCEDIMIENTO", SqlDbType.Structured)).Value = objOrden.dtProcedimientos
            command.Parameters.Add(New SqlParameter("@H_ORDEN_MEDICA_D_OXIGENO", SqlDbType.Structured)).Value = objOrden.dtOxigenos
            command.Parameters.Add(New SqlParameter("@H_ORDEN_MEDICA_D", SqlDbType.Structured)).Value = objOrden.dtMedicamentos
            command.Parameters.Add(New SqlParameter("@H_ORDEN_MEDICA_D_IMPREGNACION", SqlDbType.Structured)).Value = objOrden.dtImpregnaciones
            command.Parameters.Add(New SqlParameter("@H_ORDEN_MEDICA_D_INFUSION", SqlDbType.Structured)).Value = objOrden.dtInfusiones
            command.Parameters.Add(New SqlParameter("@H_ORDEN_MEDICA_D_MEZCLA", SqlDbType.Structured)).Value = dtMezcla
            Dim adapter As SqlDataAdapter = New SqlDataAdapter(command)

            adapter.Fill(objOrden.dtMedicamentosAuto)

        Catch ex As SqlException
            Throw ex
        End Try
    End Sub

    Private Shared Sub guardarSolicitudes(ByRef objOrdenMedica As OrdenMedica, ByRef paramsDgv As List(Of DataGridView))
        '-----------------------------------------------------------------------------------
        '-----GUARDA LOS FORMATOS DE TRANSFUCION, SOLICTUD NO POS, MEDICAMENTOS ESPECIALES, INFORMES QX Y HEMODIALISIS
        '------------------------------------------------------------------------------------
        For I = 0 To paramsDgv(1).RowCount - 2
            If Not IsNothing(paramsDgv(1).Rows(I).Cells("dgFormatoHemo").Tag) Then
                paramsDgv(1).Rows(I).Cells("dgFormatoHemo").Tag.guardarMedico(True, objOrdenMedica.codigoOrden)
            End If
        Next
        '-----SOLICITUDES DE MEDICAMENTOS NO POS Y SOLICITUDES DE JUSTIFICACION DE ANTIBIOTICOS
        For I = 0 To paramsDgv(3).RowCount - 2
            If Not IsNothing(paramsDgv(3).Rows(I).Cells("dgCodigoMed").Tag) Then
                paramsDgv(3).Rows(I).Cells("dgCodigoMed").Tag.guardarMedicamentosNoPos(True)
            End If
            If Not IsNothing(paramsDgv(3).Rows(I).Cells("dgDescripcionMed").Tag) Then
                paramsDgv(3).Rows(I).Cells("dgDescripcionMed").Tag.guardarAntibioticos(True)
            End If
        Next
        For I = 0 To paramsDgv(4).RowCount - 2
            If Not IsNothing(paramsDgv(4).Rows(I).Cells("dgCodigoImpre").Tag) Then
                paramsDgv(4).Rows(I).Cells("dgCodigoImpre").Tag.guardarMedicamentosNoPos(True)
            End If
            If Not IsNothing(paramsDgv(4).Rows(I).Cells("dgDescripcionImpre").Tag) Then
                paramsDgv(4).Rows(I).Cells("dgDescripcionImpre").Tag.guardarAntibioticos(True)
            End If
        Next
        For I = 0 To paramsDgv(0).RowCount - 2
            If Not IsNothing(paramsDgv(0).Rows(I).Cells("dgCodigoInfu").Tag) Then
                paramsDgv(0).Rows(I).Cells("dgCodigoInfu").Tag.guardarMedicamentosNoPos(True)
            End If
            If Not IsNothing(paramsDgv(0).Rows(I).Cells("dgDescripcionInfu").Tag) Then
                paramsDgv(0).Rows(I).Cells("dgDescripcionInfu").Tag.guardarAntibioticos(True)
            End If
        Next
        For I = 0 To paramsDgv(5).RowCount - 2
            If Not IsNothing(paramsDgv(5).Rows(I).Cells("dgCodigoMezcla").Tag) Then
                paramsDgv(5).Rows(I).Cells("dgCodigoMezcla").Tag.guardarMedicamentosNoPos(True)
            End If
            If Not IsNothing(paramsDgv(5).Rows(I).Cells("dgDescripcionMezcla").Tag) Then
                paramsDgv(5).Rows(I).Cells("dgDescripcionMezcla").Tag.guardarAntibioticos(True)
            End If
        Next

        '-----SOLICITUDES DE HEMODIALISIS E INFORMES QUIRÚRGICOS
        For I = 0 To paramsDgv(2).RowCount - 2
            If Not IsNothing(paramsDgv(2).Rows(I).Cells("dgCodigoProce").Tag) Then
                paramsDgv(2).Rows(I).Cells("dgCodigoProce").Tag.guardarInforme(True, objOrdenMedica.codigoOrden)
            End If
        Next
        '-----SOLICITUDES DE PLANTILLAS ESPECIALES
        For I = 0 To paramsDgv(6).RowCount - 2
            If Not IsNothing(paramsDgv(6).Rows(I).Cells("dgCodigoPara").Tag) Then
                paramsDgv(6).Rows(I).Cells("dgCodigoPara").Tag.guardarInforme(True, objOrdenMedica.codigoOrden)
            End If
        Next
    End Sub

    Public Shared Sub guardarEvolucionMedica(ByRef objEvolucionMedica As EvolucionMedica)
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objEvolucionMedica.elementoAEliminar
                        If sentencia <> "" Then
                            consulta.CommandText = sentencia
                            consulta.ExecuteNonQuery()
                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()

                    consulta.CommandText = "[SP_EVOLUCION_MEDICA_CREAR]"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                    consulta.Parameters("@CODIGO_EP").Value = objEvolucionMedica.codigoEP

                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = objEvolucionMedica.nRegistro
                    consulta.Parameters.Add(New SqlParameter("@ANALISIS", SqlDbType.NVarChar))
                    consulta.Parameters("@ANALISIS").Value = objEvolucionMedica.analisis
                    consulta.Parameters.Add(New SqlParameter("@PLAN", SqlDbType.NVarChar))
                    consulta.Parameters("@PLAN").Value = objEvolucionMedica.planTratamiento
                    consulta.Parameters.Add(New SqlParameter("@FECHA_EVO", SqlDbType.DateTime))
                    consulta.Parameters("@FECHA_EVO").Value = objEvolucionMedica.fechaEvolucion
                    consulta.Parameters.Add(New SqlParameter("@SUBJETIVO", SqlDbType.NVarChar))
                    consulta.Parameters("@SUBJETIVO").Value = objEvolucionMedica.subjetivo
                    consulta.Parameters.Add(New SqlParameter("@SIG_VITALES", SqlDbType.NVarChar))
                    consulta.Parameters("@SIG_VITALES").Value = objEvolucionMedica.signoVital
                    consulta.Parameters.Add(New SqlParameter("@CAB_CUELLO", SqlDbType.NVarChar))
                    consulta.Parameters("@CAB_CUELLO").Value = objEvolucionMedica.cabezaCuello
                    consulta.Parameters.Add(New SqlParameter("@TORAX", SqlDbType.NVarChar))
                    consulta.Parameters("@TORAX").Value = objEvolucionMedica.torax
                    consulta.Parameters.Add(New SqlParameter("@CARDIO", SqlDbType.NVarChar))
                    consulta.Parameters("@CARDIO").Value = objEvolucionMedica.cardioPulmonar
                    consulta.Parameters.Add(New SqlParameter("@ABDOMEN", SqlDbType.NVarChar))
                    consulta.Parameters("@ABDOMEN").Value = objEvolucionMedica.abdomen
                    consulta.Parameters.Add(New SqlParameter("@GENTURINARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@GENTURINARIO").Value = objEvolucionMedica.genturinario
                    consulta.Parameters.Add(New SqlParameter("@EXTREM", SqlDbType.NVarChar))
                    consulta.Parameters("@EXTREM").Value = objEvolucionMedica.extremidades
                    consulta.Parameters.Add(New SqlParameter("@S_NERV", SqlDbType.NVarChar))
                    consulta.Parameters("@S_NERV").Value = objEvolucionMedica.sistemaNervioso
                    consulta.Parameters.Add(New SqlParameter("@Usuario_r", SqlDbType.Int))
                    consulta.Parameters("@Usuario_r").Value = objEvolucionMedica.usuarioReal
                    consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int))
                    consulta.Parameters("@Usuario_Creacion").Value = objEvolucionMedica.usuario
                    consulta.Parameters.Add(New SqlParameter("@tabla", SqlDbType.Structured))
                    consulta.Parameters("@tabla").Value = objEvolucionMedica.dtParaclinicosAGuardar
                    objEvolucionMedica.codigoEvolucion = CType(consulta.ExecuteScalar, String)

                    For I = 0 To objEvolucionMedica.dtDiagnosticos.Rows.Count - 2
                        consulta.Parameters.Clear()
                        consulta.CommandText = "[SP_EVOLUCION_MEDICA_DIAGNOSTICO_CREAR]"
                        consulta.Parameters.Add(New SqlParameter("@ORDEN", SqlDbType.Int))
                        consulta.Parameters("@ORDEN").Value = I
                        consulta.Parameters.Add(New SqlParameter("@ID_EVO", SqlDbType.Int))
                        consulta.Parameters("@ID_EVO").Value = objEvolucionMedica.codigoEvolucion
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = objEvolucionMedica.dtDiagnosticos.Rows(I).Item("Código").ToString
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        consulta.Parameters("@OBSERVACION").Value = objEvolucionMedica.dtDiagnosticos.Rows(I).Item("Observacion").ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EP").Value = objEvolucionMedica.codigoEP
                        consulta.ExecuteNonQuery()
                    Next
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarEvolucionMedicaR(ByRef objEvolucionMedica As EvolucionMedicaR)
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objEvolucionMedica.elementoAEliminar
                        If sentencia <> "" Then
                            consulta.CommandText = sentencia
                            consulta.ExecuteNonQuery()
                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()

                    consulta.CommandText = "[SP_EVOLUCION_MEDICA_CREAR_R]"
                    consulta.Parameters.Add(New SqlParameter("@IDEVO", SqlDbType.Int))
                    consulta.Parameters("@IDEVO").Value = -1
                    consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                    consulta.Parameters("@EDITADO").Value = 0
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                    consulta.Parameters("@CODIGO_EP").Value = objEvolucionMedica.codigoEP

                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = objEvolucionMedica.nRegistro
                    consulta.Parameters.Add(New SqlParameter("@ANALISIS", SqlDbType.NVarChar))
                    consulta.Parameters("@ANALISIS").Value = objEvolucionMedica.analisis
                    consulta.Parameters.Add(New SqlParameter("@PLAN", SqlDbType.NVarChar))
                    consulta.Parameters("@PLAN").Value = objEvolucionMedica.planTratamiento
                    consulta.Parameters.Add(New SqlParameter("@FECHA_EVO", SqlDbType.DateTime))
                    consulta.Parameters("@FECHA_EVO").Value = objEvolucionMedica.fechaEvolucion
                    consulta.Parameters.Add(New SqlParameter("@SUBJETIVO", SqlDbType.NVarChar))
                    consulta.Parameters("@SUBJETIVO").Value = objEvolucionMedica.subjetivo
                    consulta.Parameters.Add(New SqlParameter("@SIG_VITALES", SqlDbType.NVarChar))
                    consulta.Parameters("@SIG_VITALES").Value = objEvolucionMedica.signoVital
                    consulta.Parameters.Add(New SqlParameter("@CAB_CUELLO", SqlDbType.NVarChar))
                    consulta.Parameters("@CAB_CUELLO").Value = objEvolucionMedica.cabezaCuello
                    consulta.Parameters.Add(New SqlParameter("@TORAX", SqlDbType.NVarChar))
                    consulta.Parameters("@TORAX").Value = objEvolucionMedica.torax
                    consulta.Parameters.Add(New SqlParameter("@CARDIO", SqlDbType.NVarChar))
                    consulta.Parameters("@CARDIO").Value = objEvolucionMedica.cardioPulmonar
                    consulta.Parameters.Add(New SqlParameter("@ABDOMEN", SqlDbType.NVarChar))
                    consulta.Parameters("@ABDOMEN").Value = objEvolucionMedica.abdomen
                    consulta.Parameters.Add(New SqlParameter("@GENTURINARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@GENTURINARIO").Value = objEvolucionMedica.genturinario
                    consulta.Parameters.Add(New SqlParameter("@EXTREM", SqlDbType.NVarChar))
                    consulta.Parameters("@EXTREM").Value = objEvolucionMedica.extremidades
                    consulta.Parameters.Add(New SqlParameter("@S_NERV", SqlDbType.NVarChar))
                    consulta.Parameters("@S_NERV").Value = objEvolucionMedica.sistemaNervioso
                    consulta.Parameters.Add(New SqlParameter("@Usuario_r", SqlDbType.Int))
                    consulta.Parameters("@Usuario_r").Value = objEvolucionMedica.usuarioReal
                    consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int))
                    consulta.Parameters("@Usuario_Creacion").Value = objEvolucionMedica.usuario
                    consulta.Parameters.Add(New SqlParameter("@tabla", SqlDbType.Structured))
                    consulta.Parameters("@tabla").Value = objEvolucionMedica.dtParaclinicosAGuardar
                    objEvolucionMedica.codigoEvolucion = CType(consulta.ExecuteScalar, String)

                    For I = 0 To objEvolucionMedica.dtDiagnosticos.Rows.Count - 2
                        consulta.Parameters.Clear()
                        consulta.CommandText = "[SP_EVOLUCION_MEDICA_DIAGNOSTICO_CREAR_R]"
                        consulta.Parameters.Add(New SqlParameter("@ID_EVO", SqlDbType.Int))
                        consulta.Parameters("@ID_EVO").Value = objEvolucionMedica.codigoEvolucion
                        consulta.Parameters.Add(New SqlParameter("@ORDEN", SqlDbType.Int))
                        consulta.Parameters("@ORDEN").Value = I
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = objEvolucionMedica.dtDiagnosticos.Rows(I).Item("Código").ToString
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        consulta.Parameters("@OBSERVACION").Value = objEvolucionMedica.dtDiagnosticos.Rows(I).Item("Observacion").ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EP").Value = objEvolucionMedica.codigoEP
                        consulta.ExecuteNonQuery()
                    Next
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub guardarEvolucionMedicaRR(ByRef objEvolucionMedica As EvolucionMedicaRR)
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objEvolucionMedica.elementoAEliminar
                        If sentencia <> "" Then
                            consulta.CommandText = sentencia
                            consulta.ExecuteNonQuery()
                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()

                    consulta.CommandText = "[SP_EVOLUCION_MEDICA_CREAR_RR]"
                    consulta.Parameters.Add(New SqlParameter("@IDEVO", SqlDbType.Int))
                    consulta.Parameters("@IDEVO").Value = -1
                    consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                    consulta.Parameters("@EDITADO").Value = 0

                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = objEvolucionMedica.nRegistro
                    consulta.Parameters.Add(New SqlParameter("@ANALISIS", SqlDbType.NVarChar))
                    consulta.Parameters("@ANALISIS").Value = objEvolucionMedica.analisis
                    consulta.Parameters.Add(New SqlParameter("@PLAN", SqlDbType.NVarChar))
                    consulta.Parameters("@PLAN").Value = objEvolucionMedica.planTratamiento
                    consulta.Parameters.Add(New SqlParameter("@FECHA_EVO", SqlDbType.DateTime))
                    consulta.Parameters("@FECHA_EVO").Value = objEvolucionMedica.fechaEvolucion
                    consulta.Parameters.Add(New SqlParameter("@SUBJETIVO", SqlDbType.NVarChar))
                    consulta.Parameters("@SUBJETIVO").Value = objEvolucionMedica.subjetivo
                    consulta.Parameters.Add(New SqlParameter("@SIG_VITALES", SqlDbType.NVarChar))
                    consulta.Parameters("@SIG_VITALES").Value = objEvolucionMedica.signoVital
                    consulta.Parameters.Add(New SqlParameter("@CAB_CUELLO", SqlDbType.NVarChar))
                    consulta.Parameters("@CAB_CUELLO").Value = objEvolucionMedica.cabezaCuello
                    consulta.Parameters.Add(New SqlParameter("@TORAX", SqlDbType.NVarChar))
                    consulta.Parameters("@TORAX").Value = objEvolucionMedica.torax
                    consulta.Parameters.Add(New SqlParameter("@CARDIO", SqlDbType.NVarChar))
                    consulta.Parameters("@CARDIO").Value = objEvolucionMedica.cardioPulmonar
                    consulta.Parameters.Add(New SqlParameter("@ABDOMEN", SqlDbType.NVarChar))
                    consulta.Parameters("@ABDOMEN").Value = objEvolucionMedica.abdomen
                    consulta.Parameters.Add(New SqlParameter("@GENTURINARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@GENTURINARIO").Value = objEvolucionMedica.genturinario
                    consulta.Parameters.Add(New SqlParameter("@EXTREM", SqlDbType.NVarChar))
                    consulta.Parameters("@EXTREM").Value = objEvolucionMedica.extremidades
                    consulta.Parameters.Add(New SqlParameter("@S_NERV", SqlDbType.NVarChar))
                    consulta.Parameters("@S_NERV").Value = objEvolucionMedica.sistemaNervioso
                    consulta.Parameters.Add(New SqlParameter("@Usuario_r", SqlDbType.Int))
                    consulta.Parameters("@Usuario_r").Value = objEvolucionMedica.usuarioReal
                    consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int))
                    consulta.Parameters("@Usuario_Creacion").Value = objEvolucionMedica.usuario
                    consulta.Parameters.Add(New SqlParameter("@tabla", SqlDbType.Structured))
                    consulta.Parameters("@tabla").Value = objEvolucionMedica.dtParaclinicosAGuardar
                    objEvolucionMedica.codigoEvolucion = CType(consulta.ExecuteScalar, String)

                    For I = 0 To objEvolucionMedica.dtDiagnosticos.Rows.Count - 2
                        consulta.Parameters.Clear()
                        consulta.CommandText = "[SP_EVOLUCION_MEDICA_DIAGNOSTICO_CREAR_RR]"
                        consulta.Parameters.Add(New SqlParameter("@ID_EVO", SqlDbType.Int))
                        consulta.Parameters("@ID_EVO").Value = objEvolucionMedica.codigoEvolucion
                        consulta.Parameters.Add(New SqlParameter("@ORDEN", SqlDbType.Int))
                        consulta.Parameters("@ORDEN").Value = I
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = objEvolucionMedica.dtDiagnosticos.Rows(I).Item("Código").ToString
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        consulta.Parameters("@OBSERVACION").Value = objEvolucionMedica.dtDiagnosticos.Rows(I).Item("Observacion").ToString
                        consulta.ExecuteNonQuery()
                    Next
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub actualizarEvolucionMedica(ByRef objEvolucionMedica As EvolucionMedica)
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objEvolucionMedica.elementoAEliminar
                        If sentencia <> "" Then
                            consulta.CommandText = sentencia
                            consulta.ExecuteNonQuery()
                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure

                    consulta.Parameters.Clear()

                    consulta.CommandText = "[SP_EVOLUCION_MEDICA_ACTUALIZAR]"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                    consulta.Parameters("@CODIGO_EP").Value = objEvolucionMedica.codigoEP
                    consulta.Parameters.Add(New SqlParameter("@IDEVO", SqlDbType.Int))
                    consulta.Parameters("@IDEVO").Value = objEvolucionMedica.codigoEvolucion
                    consulta.Parameters.Add(New SqlParameter("@ANALISIS", SqlDbType.NVarChar))
                    consulta.Parameters("@ANALISIS").Value = objEvolucionMedica.analisis
                    consulta.Parameters.Add(New SqlParameter("@PLAN", SqlDbType.NVarChar))
                    consulta.Parameters("@PLAN").Value = objEvolucionMedica.planTratamiento
                    consulta.Parameters.Add(New SqlParameter("@FECHA_EVO", SqlDbType.DateTime))
                    consulta.Parameters("@FECHA_EVO").Value = objEvolucionMedica.fechaEvolucion
                    consulta.Parameters.Add(New SqlParameter("@SUBJETIVO", SqlDbType.NVarChar))
                    consulta.Parameters("@SUBJETIVO").Value = objEvolucionMedica.subjetivo
                    consulta.Parameters.Add(New SqlParameter("@SIG_VITALES", SqlDbType.NVarChar))
                    consulta.Parameters("@SIG_VITALES").Value = objEvolucionMedica.signoVital
                    consulta.Parameters.Add(New SqlParameter("@CAB_CUELLO", SqlDbType.NVarChar))
                    consulta.Parameters("@CAB_CUELLO").Value = objEvolucionMedica.cabezaCuello
                    consulta.Parameters.Add(New SqlParameter("@TORAX", SqlDbType.NVarChar))
                    consulta.Parameters("@TORAX").Value = objEvolucionMedica.torax
                    consulta.Parameters.Add(New SqlParameter("@CARDIO", SqlDbType.NVarChar))
                    consulta.Parameters("@CARDIO").Value = objEvolucionMedica.cardioPulmonar
                    consulta.Parameters.Add(New SqlParameter("@ABDOMEN", SqlDbType.NVarChar))
                    consulta.Parameters("@ABDOMEN").Value = objEvolucionMedica.abdomen
                    consulta.Parameters.Add(New SqlParameter("@GENTURINARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@GENTURINARIO").Value = objEvolucionMedica.genturinario
                    consulta.Parameters.Add(New SqlParameter("@EXTREM", SqlDbType.NVarChar))
                    consulta.Parameters("@EXTREM").Value = objEvolucionMedica.extremidades
                    consulta.Parameters.Add(New SqlParameter("@S_NERV", SqlDbType.NVarChar))
                    consulta.Parameters("@S_NERV").Value = objEvolucionMedica.sistemaNervioso
                    consulta.Parameters.Add(New SqlParameter("@Usuario_r", SqlDbType.Int))
                    consulta.Parameters("@Usuario_r").Value = objEvolucionMedica.usuarioReal
                    consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                    consulta.Parameters("@Usuario_actualizacion").Value = objEvolucionMedica.usuario
                    consulta.Parameters.Add(New SqlParameter("@tabla", SqlDbType.Structured))
                    consulta.Parameters("@tabla").Value = objEvolucionMedica.dtParaclinicosAGuardar
                    consulta.ExecuteNonQuery()
                    For I = 0 To objEvolucionMedica.dtDiagnosticos.Rows.Count - 2
                        consulta.Parameters.Clear()
                        consulta.CommandText = "[SP_EVOLUCION_MEDICA_DIAGNOSTICO_CREAR]"
                        consulta.Parameters.Add(New SqlParameter("@ID_EVO", SqlDbType.Int))
                        consulta.Parameters("@ID_EVO").Value = objEvolucionMedica.codigoEvolucion
                        consulta.Parameters.Add(New SqlParameter("@ORDEN", SqlDbType.Int))
                        consulta.Parameters("@ORDEN").Value = I
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = objEvolucionMedica.dtDiagnosticos.Rows(I).Item("Código").ToString
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        consulta.Parameters("@OBSERVACION").Value = objEvolucionMedica.dtDiagnosticos.Rows(I).Item("Observacion").ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EP").Value = objEvolucionMedica.codigoEP
                        consulta.ExecuteNonQuery()
                    Next

                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw ex

        End Try
    End Sub
    Public Shared Sub actualizarEvolucionMedicaR(ByRef objEvolucionMedica As EvolucionMedicaR)
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objEvolucionMedica.elementoAEliminar
                        If sentencia <> "" Then
                            consulta.CommandText = sentencia
                            consulta.ExecuteNonQuery()
                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()

                    consulta.CommandText = "[SP_EVOLUCION_MEDICA_ACTUALIZAR_R]"
                    consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                    consulta.Parameters("@EDITADO").Value = 1
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                    consulta.Parameters("@CODIGO_EP").Value = objEvolucionMedica.codigoEP

                    consulta.Parameters.Add(New SqlParameter("@IDEVO", SqlDbType.Int))
                    consulta.Parameters("@IDEVO").Value = objEvolucionMedica.codigoEvolucion
                    consulta.Parameters.Add(New SqlParameter("@ANALISIS", SqlDbType.NVarChar))
                    consulta.Parameters("@ANALISIS").Value = objEvolucionMedica.analisis
                    consulta.Parameters.Add(New SqlParameter("@PLAN", SqlDbType.NVarChar))
                    consulta.Parameters("@PLAN").Value = objEvolucionMedica.planTratamiento
                    consulta.Parameters.Add(New SqlParameter("@FECHA_EVO", SqlDbType.DateTime))
                    consulta.Parameters("@FECHA_EVO").Value = objEvolucionMedica.fechaEvolucion
                    consulta.Parameters.Add(New SqlParameter("@SUBJETIVO", SqlDbType.NVarChar))
                    consulta.Parameters("@SUBJETIVO").Value = objEvolucionMedica.subjetivo
                    consulta.Parameters.Add(New SqlParameter("@SIG_VITALES", SqlDbType.NVarChar))
                    consulta.Parameters("@SIG_VITALES").Value = objEvolucionMedica.signoVital
                    consulta.Parameters.Add(New SqlParameter("@CAB_CUELLO", SqlDbType.NVarChar))
                    consulta.Parameters("@CAB_CUELLO").Value = objEvolucionMedica.cabezaCuello
                    consulta.Parameters.Add(New SqlParameter("@TORAX", SqlDbType.NVarChar))
                    consulta.Parameters("@TORAX").Value = objEvolucionMedica.torax
                    consulta.Parameters.Add(New SqlParameter("@CARDIO", SqlDbType.NVarChar))
                    consulta.Parameters("@CARDIO").Value = objEvolucionMedica.cardioPulmonar
                    consulta.Parameters.Add(New SqlParameter("@ABDOMEN", SqlDbType.NVarChar))
                    consulta.Parameters("@ABDOMEN").Value = objEvolucionMedica.abdomen
                    consulta.Parameters.Add(New SqlParameter("@GENTURINARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@GENTURINARIO").Value = objEvolucionMedica.genturinario
                    consulta.Parameters.Add(New SqlParameter("@EXTREM", SqlDbType.NVarChar))
                    consulta.Parameters("@EXTREM").Value = objEvolucionMedica.extremidades
                    consulta.Parameters.Add(New SqlParameter("@S_NERV", SqlDbType.NVarChar))
                    consulta.Parameters("@S_NERV").Value = objEvolucionMedica.sistemaNervioso
                    consulta.Parameters.Add(New SqlParameter("@Usuario_r", SqlDbType.Int))
                    consulta.Parameters("@Usuario_r").Value = objEvolucionMedica.usuarioReal
                    consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                    consulta.Parameters("@Usuario_actualizacion").Value = objEvolucionMedica.usuario
                    consulta.Parameters.Add(New SqlParameter("@tabla", SqlDbType.Structured))
                    consulta.Parameters("@tabla").Value = objEvolucionMedica.dtParaclinicosAGuardar
                    consulta.ExecuteNonQuery()
                    For I = 0 To objEvolucionMedica.dtDiagnosticos.Rows.Count - 2
                        consulta.Parameters.Clear()
                        consulta.CommandText = "[SP_EVOLUCION_MEDICA_DIAGNOSTICO_CREAR_R]"
                        consulta.Parameters.Add(New SqlParameter("@ID_EVO", SqlDbType.Int))
                        consulta.Parameters("@ID_EVO").Value = objEvolucionMedica.codigoEvolucion
                        consulta.Parameters.Add(New SqlParameter("@ORDEN", SqlDbType.Int))
                        consulta.Parameters("@ORDEN").Value = I
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = objEvolucionMedica.dtDiagnosticos.Rows(I).Item("Código").ToString
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        consulta.Parameters("@OBSERVACION").Value = objEvolucionMedica.dtDiagnosticos.Rows(I).Item("Observacion").ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EP").Value = objEvolucionMedica.codigoEP
                        consulta.ExecuteNonQuery()
                    Next

                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw ex

        End Try
    End Sub

    Public Shared Sub actualizarEvolucionMedicaRR(ByRef objEvolucionMedica As EvolucionMedicaRR)
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objEvolucionMedica.elementoAEliminar
                        If sentencia <> "" Then
                            consulta.CommandText = sentencia
                            consulta.ExecuteNonQuery()
                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure

                    consulta.Parameters.Clear()
                    consulta.CommandText = "[SP_EVOLUCION_MEDICA_ACTUALIZAR_RR]"
                    consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                    consulta.Parameters("@EDITADO").Value = 1

                    consulta.Parameters.Add(New SqlParameter("@IDEVO", SqlDbType.Int))
                    consulta.Parameters("@IDEVO").Value = objEvolucionMedica.codigoEvolucion
                    consulta.Parameters.Add(New SqlParameter("@ANALISIS", SqlDbType.NVarChar))
                    consulta.Parameters("@ANALISIS").Value = objEvolucionMedica.analisis
                    consulta.Parameters.Add(New SqlParameter("@PLAN", SqlDbType.NVarChar))
                    consulta.Parameters("@PLAN").Value = objEvolucionMedica.planTratamiento
                    consulta.Parameters.Add(New SqlParameter("@FECHA_EVO", SqlDbType.DateTime))
                    consulta.Parameters("@FECHA_EVO").Value = objEvolucionMedica.fechaEvolucion
                    consulta.Parameters.Add(New SqlParameter("@SUBJETIVO", SqlDbType.NVarChar))
                    consulta.Parameters("@SUBJETIVO").Value = objEvolucionMedica.subjetivo
                    consulta.Parameters.Add(New SqlParameter("@SIG_VITALES", SqlDbType.NVarChar))
                    consulta.Parameters("@SIG_VITALES").Value = objEvolucionMedica.signoVital
                    consulta.Parameters.Add(New SqlParameter("@CAB_CUELLO", SqlDbType.NVarChar))
                    consulta.Parameters("@CAB_CUELLO").Value = objEvolucionMedica.cabezaCuello
                    consulta.Parameters.Add(New SqlParameter("@TORAX", SqlDbType.NVarChar))
                    consulta.Parameters("@TORAX").Value = objEvolucionMedica.torax
                    consulta.Parameters.Add(New SqlParameter("@CARDIO", SqlDbType.NVarChar))
                    consulta.Parameters("@CARDIO").Value = objEvolucionMedica.cardioPulmonar
                    consulta.Parameters.Add(New SqlParameter("@ABDOMEN", SqlDbType.NVarChar))
                    consulta.Parameters("@ABDOMEN").Value = objEvolucionMedica.abdomen
                    consulta.Parameters.Add(New SqlParameter("@GENTURINARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@GENTURINARIO").Value = objEvolucionMedica.genturinario
                    consulta.Parameters.Add(New SqlParameter("@EXTREM", SqlDbType.NVarChar))
                    consulta.Parameters("@EXTREM").Value = objEvolucionMedica.extremidades
                    consulta.Parameters.Add(New SqlParameter("@S_NERV", SqlDbType.NVarChar))
                    consulta.Parameters("@S_NERV").Value = objEvolucionMedica.sistemaNervioso
                    consulta.Parameters.Add(New SqlParameter("@Usuario_r", SqlDbType.Int))
                    consulta.Parameters("@Usuario_r").Value = objEvolucionMedica.usuarioReal
                    consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                    consulta.Parameters("@Usuario_actualizacion").Value = objEvolucionMedica.usuario
                    consulta.Parameters.Add(New SqlParameter("@tabla", SqlDbType.Structured))
                    consulta.Parameters("@tabla").Value = objEvolucionMedica.dtParaclinicosAGuardar
                    consulta.ExecuteNonQuery()
                    For I = 0 To objEvolucionMedica.dtDiagnosticos.Rows.Count - 2
                        consulta.Parameters.Clear()
                        consulta.CommandText = "[SP_EVOLUCION_MEDICA_DIAGNOSTICO_CREAR_RR]"
                        consulta.Parameters.Add(New SqlParameter("@ID_EVO", SqlDbType.Int))
                        consulta.Parameters("@ID_EVO").Value = objEvolucionMedica.codigoEvolucion
                        consulta.Parameters.Add(New SqlParameter("@ORDEN", SqlDbType.Int))
                        consulta.Parameters("@ORDEN").Value = I
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = objEvolucionMedica.dtDiagnosticos.Rows(I).Item("Código").ToString
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        consulta.Parameters("@OBSERVACION").Value = objEvolucionMedica.dtDiagnosticos.Rows(I).Item("Observacion").ToString
                        consulta.ExecuteNonQuery()
                    Next

                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw ex

        End Try
    End Sub

    Public Shared Sub guardarInterconsulta(ByRef objInterconsulta As InterconsultaMedica)
        Try
            Using consulta = New SqlCommand()
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "[SP_INTERCONSULTA_GUARDAR]"
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                consulta.Parameters("@CODIGO_EP").Value = objInterconsulta.codigoEP

                consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int))
                consulta.Parameters("@CODIGO_ORDEN").Value = objInterconsulta.codigoOrden
                consulta.Parameters.Add(New SqlParameter("@PROCEDIMIENTO", SqlDbType.NVarChar))
                consulta.Parameters("@PROCEDIMIENTO").Value = objInterconsulta.codigoProcedimiento
                consulta.Parameters.Add(New SqlParameter("@RESPUESTA", SqlDbType.NVarChar))
                consulta.Parameters("@RESPUESTA").Value = objInterconsulta.Respuesta
                consulta.Parameters.Add(New SqlParameter("@FECHA_INT", SqlDbType.DateTime))
                consulta.Parameters("@FECHA_INT").Value = objInterconsulta.fechaInterconsulta
                consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int))
                consulta.Parameters("@Usuario_Creacion").Value = SesionActual.idUsuario
                consulta.Parameters.Add(New SqlParameter("@Usuario_Real", SqlDbType.Int))
                consulta.Parameters("@Usuario_Real").Value = objInterconsulta.usuario
                consulta.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarInterconsultaR(ByRef objInterconsulta As InterconsultaMedicaR)
        Try
            Using consulta = New SqlCommand()
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "[SP_INTERCONSULTA_GUARDAR_R]"
                consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                consulta.Parameters("@EDITADO").Value = 0
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                consulta.Parameters("@CODIGO_EP").Value = objInterconsulta.codigoEP

                consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int))
                consulta.Parameters("@CODIGO_ORDEN").Value = objInterconsulta.codigoOrden
                consulta.Parameters.Add(New SqlParameter("@PROCEDIMIENTO", SqlDbType.NVarChar))
                consulta.Parameters("@PROCEDIMIENTO").Value = objInterconsulta.codigoProcedimiento
                consulta.Parameters.Add(New SqlParameter("@RESPUESTA", SqlDbType.NVarChar))
                consulta.Parameters("@RESPUESTA").Value = objInterconsulta.Respuesta
                consulta.Parameters.Add(New SqlParameter("@FECHA_INT", SqlDbType.DateTime))
                consulta.Parameters("@FECHA_INT").Value = objInterconsulta.fechaInterconsulta
                consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int))
                consulta.Parameters("@Usuario_Creacion").Value = SesionActual.idUsuario
                consulta.Parameters.Add(New SqlParameter("@Usuario_Real", SqlDbType.Int))
                consulta.Parameters("@Usuario_Real").Value = objInterconsulta.usuario
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarInterconsultaRR(ByRef objInterconsulta As InterconsultaMedicaRR)
        Try
            Using consulta = New SqlCommand()
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "[SP_INTERCONSULTA_GUARDAR_RR]"
                consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                consulta.Parameters("@EDITADO").Value = 0

                consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int))
                consulta.Parameters("@CODIGO_ORDEN").Value = objInterconsulta.codigoOrden
                consulta.Parameters.Add(New SqlParameter("@PROCEDIMIENTO", SqlDbType.NVarChar))
                consulta.Parameters("@PROCEDIMIENTO").Value = objInterconsulta.codigoProcedimiento
                consulta.Parameters.Add(New SqlParameter("@RESPUESTA", SqlDbType.NVarChar))
                consulta.Parameters("@RESPUESTA").Value = objInterconsulta.Respuesta
                consulta.Parameters.Add(New SqlParameter("@FECHA_INT", SqlDbType.DateTime))
                consulta.Parameters("@FECHA_INT").Value = objInterconsulta.fechaInterconsulta
                consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int))
                consulta.Parameters("@Usuario_Creacion").Value = SesionActual.idUsuario
                consulta.Parameters.Add(New SqlParameter("@Usuario_Real", SqlDbType.Int))
                consulta.Parameters("@Usuario_Real").Value = objInterconsulta.usuario
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub guardarRemision(ByRef pRemision As Remision)
        Try
            Using consulta = New SqlCommand()
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure
                consulta.CommandText = "[SP_REMISION_CREAR]"
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                consulta.Parameters("@CODIGO_EP").Value = pRemision.codigoEP

                consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int))
                consulta.Parameters("@CODIGO_ORDEN").Value = pRemision.codigoOrden
                consulta.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int))
                consulta.Parameters("@N_Registro").Value = pRemision.nRegistro
                consulta.Parameters.Add(New SqlParameter("@Codigo_Motivo_observacion", SqlDbType.Int))
                consulta.Parameters("@Codigo_Motivo_observacion").Value = pRemision.modalidad
                consulta.Parameters.Add(New SqlParameter("@Datos_medicos", SqlDbType.NVarChar))
                consulta.Parameters("@Datos_medicos").Value = pRemision.datosMedicos
                consulta.Parameters.Add(New SqlParameter("@Fecha_Remision", SqlDbType.DateTime))
                consulta.Parameters("@Fecha_Remision").Value = pRemision.fechaRemision
                consulta.Parameters.Add(New SqlParameter("@Codigo_Prioridad", SqlDbType.Int))
                consulta.Parameters("@Codigo_Prioridad").Value = pRemision.prioridad
                consulta.Parameters.Add(New SqlParameter("@Codigo_complejidad", SqlDbType.Int))
                consulta.Parameters("@Codigo_complejidad").Value = pRemision.complejidad
                consulta.Parameters.Add(New SqlParameter("@Otras", SqlDbType.NVarChar))
                consulta.Parameters("@Otras").Value = pRemision.otras
                consulta.Parameters.Add(New SqlParameter("@Antecedentes", SqlDbType.NVarChar))
                consulta.Parameters("@Antecedentes").Value = pRemision.antecedentes
                consulta.Parameters.Add(New SqlParameter("@Glasgow", SqlDbType.NVarChar))
                consulta.Parameters("@Glasgow").Value = pRemision.glasgow
                consulta.Parameters.Add(New SqlParameter("@Descripcion_Glasgow", SqlDbType.NVarChar))
                consulta.Parameters("@Descripcion_Glasgow").Value = pRemision.descripglasgow
                consulta.Parameters.Add(New SqlParameter("@Presion_Sistolica", SqlDbType.NVarChar))
                consulta.Parameters("@Presion_Sistolica").Value = pRemision.presionsis
                consulta.Parameters.Add(New SqlParameter("@Presion_Diastolica", SqlDbType.NVarChar))
                consulta.Parameters("@Presion_Diastolica").Value = pRemision.presiondias
                consulta.Parameters.Add(New SqlParameter("@Frecuencia_Cardiaca", SqlDbType.NVarChar))
                consulta.Parameters("@Frecuencia_Cardiaca").Value = pRemision.freccar
                consulta.Parameters.Add(New SqlParameter("@Frecuencia_Respiratoria", SqlDbType.NVarChar))
                consulta.Parameters("@Frecuencia_Respiratoria").Value = pRemision.frecresp
                consulta.Parameters.Add(New SqlParameter("@Ambulancia", SqlDbType.NVarChar))
                consulta.Parameters("@Ambulancia").Value = pRemision.ambulancia
                consulta.Parameters.Add(New SqlParameter("@Codigo_traslado", SqlDbType.NVarChar))
                consulta.Parameters("@Codigo_traslado").Value = pRemision.traslados
                consulta.Parameters.Add(New SqlParameter("@oxigeno", SqlDbType.NVarChar))
                consulta.Parameters("@oxigeno").Value = pRemision.oxigeno
                consulta.Parameters.Add(New SqlParameter("@codigo_especialidad", SqlDbType.Int))
                consulta.Parameters("@codigo_especialidad").Value = pRemision.especialidad
                consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int))
                consulta.Parameters("@Usuario_Creacion").Value = pRemision.usuarioReal
                pRemision.codigoRemision = CType(consulta.ExecuteScalar, String)

            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarRemisionR(ByRef pRemision As RemisionR)
        Try
            Using consulta = New SqlCommand()
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure

                consulta.CommandText = "[SP_REMISION_CREAR_R]"
                consulta.Parameters.Add(New SqlParameter("@Codigo_remision", SqlDbType.Int))
                consulta.Parameters("@Codigo_remision").Value = -1
                consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                consulta.Parameters("@EDITADO").Value = 0
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                consulta.Parameters("@CODIGO_EP").Value = pRemision.codigoEP
                consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int))
                consulta.Parameters("@CODIGO_ORDEN").Value = pRemision.codigoOrden
                consulta.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int))
                consulta.Parameters("@N_Registro").Value = pRemision.nRegistro
                consulta.Parameters.Add(New SqlParameter("@Codigo_Motivo_observacion", SqlDbType.Int))
                consulta.Parameters("@Codigo_Motivo_observacion").Value = pRemision.modalidad
                consulta.Parameters.Add(New SqlParameter("@Datos_medicos", SqlDbType.NVarChar))
                consulta.Parameters("@Datos_medicos").Value = pRemision.datosMedicos
                consulta.Parameters.Add(New SqlParameter("@Fecha_Remision", SqlDbType.DateTime))
                consulta.Parameters("@Fecha_Remision").Value = pRemision.fechaRemision
                consulta.Parameters.Add(New SqlParameter("@Codigo_Prioridad", SqlDbType.Int))
                consulta.Parameters("@Codigo_Prioridad").Value = pRemision.prioridad
                consulta.Parameters.Add(New SqlParameter("@Codigo_complejidad", SqlDbType.Int))
                consulta.Parameters("@Codigo_complejidad").Value = pRemision.complejidad
                consulta.Parameters.Add(New SqlParameter("@Otras", SqlDbType.NVarChar))
                consulta.Parameters("@Otras").Value = pRemision.otras
                consulta.Parameters.Add(New SqlParameter("@Antecedentes", SqlDbType.NVarChar))
                consulta.Parameters("@Antecedentes").Value = pRemision.antecedentes
                consulta.Parameters.Add(New SqlParameter("@Glasgow", SqlDbType.NVarChar))
                consulta.Parameters("@Glasgow").Value = pRemision.glasgow
                consulta.Parameters.Add(New SqlParameter("@Descripcion_Glasgow", SqlDbType.NVarChar))
                consulta.Parameters("@Descripcion_Glasgow").Value = pRemision.descripglasgow
                consulta.Parameters.Add(New SqlParameter("@Presion_Sistolica", SqlDbType.NVarChar))
                consulta.Parameters("@Presion_Sistolica").Value = pRemision.presionsis
                consulta.Parameters.Add(New SqlParameter("@Presion_Diastolica", SqlDbType.NVarChar))
                consulta.Parameters("@Presion_Diastolica").Value = pRemision.presiondias
                consulta.Parameters.Add(New SqlParameter("@Frecuencia_Cardiaca", SqlDbType.NVarChar))
                consulta.Parameters("@Frecuencia_Cardiaca").Value = pRemision.freccar
                consulta.Parameters.Add(New SqlParameter("@Frecuencia_Respiratoria", SqlDbType.NVarChar))
                consulta.Parameters("@Frecuencia_Respiratoria").Value = pRemision.frecresp
                consulta.Parameters.Add(New SqlParameter("@Ambulancia", SqlDbType.NVarChar))
                consulta.Parameters("@Ambulancia").Value = pRemision.ambulancia
                consulta.Parameters.Add(New SqlParameter("@Codigo_traslado", SqlDbType.NVarChar))
                consulta.Parameters("@Codigo_traslado").Value = pRemision.traslados
                consulta.Parameters.Add(New SqlParameter("@oxigeno", SqlDbType.NVarChar))
                consulta.Parameters("@oxigeno").Value = pRemision.oxigeno
                consulta.Parameters.Add(New SqlParameter("@codigo_especialidad", SqlDbType.Int))
                consulta.Parameters("@codigo_especialidad").Value = pRemision.especialidad
                consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int))
                consulta.Parameters("@Usuario_Creacion").Value = pRemision.usuarioReal
                pRemision.codigoRemision = CType(consulta.ExecuteScalar, String)

            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarRemisionRR(ByRef pRemision As RemisionRR)
        Try
            Using consulta = New SqlCommand()
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure

                consulta.CommandText = "[SP_REMISION_CREAR_RR]"
                consulta.Parameters.Add(New SqlParameter("@Codigo_remision", SqlDbType.Int))
                consulta.Parameters("@Codigo_remision").Value = -1
                consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                consulta.Parameters("@EDITADO").Value = 0
                consulta.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int))
                consulta.Parameters("@CODIGO_ORDEN").Value = pRemision.codigoOrden
                consulta.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int))
                consulta.Parameters("@N_Registro").Value = pRemision.nRegistro
                consulta.Parameters.Add(New SqlParameter("@Codigo_Motivo_observacion", SqlDbType.Int))
                consulta.Parameters("@Codigo_Motivo_observacion").Value = pRemision.modalidad
                consulta.Parameters.Add(New SqlParameter("@Datos_medicos", SqlDbType.NVarChar))
                consulta.Parameters("@Datos_medicos").Value = pRemision.datosMedicos
                consulta.Parameters.Add(New SqlParameter("@Fecha_Remision", SqlDbType.DateTime))
                consulta.Parameters("@Fecha_Remision").Value = pRemision.fechaRemision
                consulta.Parameters.Add(New SqlParameter("@Codigo_Prioridad", SqlDbType.Int))
                consulta.Parameters("@Codigo_Prioridad").Value = pRemision.prioridad
                consulta.Parameters.Add(New SqlParameter("@Codigo_complejidad", SqlDbType.Int))
                consulta.Parameters("@Codigo_complejidad").Value = pRemision.complejidad
                consulta.Parameters.Add(New SqlParameter("@Otras", SqlDbType.NVarChar))
                consulta.Parameters("@Otras").Value = pRemision.otras
                consulta.Parameters.Add(New SqlParameter("@Antecedentes", SqlDbType.NVarChar))
                consulta.Parameters("@Antecedentes").Value = pRemision.antecedentes
                consulta.Parameters.Add(New SqlParameter("@Glasgow", SqlDbType.NVarChar))
                consulta.Parameters("@Glasgow").Value = pRemision.glasgow
                consulta.Parameters.Add(New SqlParameter("@Descripcion_Glasgow", SqlDbType.NVarChar))
                consulta.Parameters("@Descripcion_Glasgow").Value = pRemision.descripglasgow
                consulta.Parameters.Add(New SqlParameter("@Presion_Sistolica", SqlDbType.NVarChar))
                consulta.Parameters("@Presion_Sistolica").Value = pRemision.presionsis
                consulta.Parameters.Add(New SqlParameter("@Presion_Diastolica", SqlDbType.NVarChar))
                consulta.Parameters("@Presion_Diastolica").Value = pRemision.presiondias
                consulta.Parameters.Add(New SqlParameter("@Frecuencia_Cardiaca", SqlDbType.NVarChar))
                consulta.Parameters("@Frecuencia_Cardiaca").Value = pRemision.freccar
                consulta.Parameters.Add(New SqlParameter("@Frecuencia_Respiratoria", SqlDbType.NVarChar))
                consulta.Parameters("@Frecuencia_Respiratoria").Value = pRemision.frecresp
                consulta.Parameters.Add(New SqlParameter("@Ambulancia", SqlDbType.NVarChar))
                consulta.Parameters("@Ambulancia").Value = pRemision.ambulancia
                consulta.Parameters.Add(New SqlParameter("@Codigo_traslado", SqlDbType.NVarChar))
                consulta.Parameters("@Codigo_traslado").Value = pRemision.traslados
                consulta.Parameters.Add(New SqlParameter("@oxigeno", SqlDbType.NVarChar))
                consulta.Parameters("@oxigeno").Value = pRemision.oxigeno
                consulta.Parameters.Add(New SqlParameter("@codigo_especialidad", SqlDbType.Int))
                consulta.Parameters("@codigo_especialidad").Value = pRemision.especialidad
                consulta.Parameters.Add(New SqlParameter("@Usuario_Creacion", SqlDbType.Int))
                consulta.Parameters("@Usuario_Creacion").Value = pRemision.usuarioReal
                pRemision.codigoRemision = CType(consulta.ExecuteScalar, String)

            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarRemision(ByRef pRemision As Remision)
        Try
            Using consulta = New SqlCommand()
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure

                consulta.CommandText = "[SP_REMISION_ACTUALIZAR]"
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                consulta.Parameters("@CODIGO_EP").Value = pRemision.codigoEP

                consulta.Parameters.Add(New SqlParameter("@CODIGO_REM", SqlDbType.Int))
                consulta.Parameters("@CODIGO_REM").Value = pRemision.codigoRemision
                consulta.Parameters.Add(New SqlParameter("@Codigo_Motivo_observacion", SqlDbType.Int))
                consulta.Parameters("@Codigo_Motivo_observacion").Value = pRemision.modalidad
                consulta.Parameters.Add(New SqlParameter("@Datos_medicos", SqlDbType.NVarChar))
                consulta.Parameters("@Datos_medicos").Value = pRemision.datosMedicos
                consulta.Parameters.Add(New SqlParameter("@Fecha_Remision", SqlDbType.DateTime))
                consulta.Parameters("@Fecha_Remision").Value = pRemision.fechaRemision
                consulta.Parameters.Add(New SqlParameter("@Codigo_Prioridad", SqlDbType.Int))
                consulta.Parameters("@Codigo_Prioridad").Value = pRemision.prioridad
                consulta.Parameters.Add(New SqlParameter("@Codigo_complejidad", SqlDbType.Int))
                consulta.Parameters("@Codigo_complejidad").Value = pRemision.complejidad
                consulta.Parameters.Add(New SqlParameter("@Otras", SqlDbType.NVarChar))
                consulta.Parameters("@Otras").Value = pRemision.otras
                consulta.Parameters.Add(New SqlParameter("@Antecedentes", SqlDbType.NVarChar))
                consulta.Parameters("@Antecedentes").Value = pRemision.antecedentes
                consulta.Parameters.Add(New SqlParameter("@Glasgow", SqlDbType.NVarChar))
                consulta.Parameters("@Glasgow").Value = pRemision.glasgow
                consulta.Parameters.Add(New SqlParameter("@Descripcion_Glasgow", SqlDbType.NVarChar))
                consulta.Parameters("@Descripcion_Glasgow").Value = pRemision.descripglasgow
                consulta.Parameters.Add(New SqlParameter("@Presion_Sistolica", SqlDbType.NVarChar))
                consulta.Parameters("@Presion_Sistolica").Value = pRemision.presionsis
                consulta.Parameters.Add(New SqlParameter("@Presion_Diastolica", SqlDbType.NVarChar))
                consulta.Parameters("@Presion_Diastolica").Value = pRemision.presiondias
                consulta.Parameters.Add(New SqlParameter("@Frecuencia_Cardiaca", SqlDbType.NVarChar))
                consulta.Parameters("@Frecuencia_Cardiaca").Value = pRemision.freccar
                consulta.Parameters.Add(New SqlParameter("@Frecuencia_Respiratoria", SqlDbType.NVarChar))
                consulta.Parameters("@Frecuencia_Respiratoria").Value = pRemision.frecresp
                consulta.Parameters.Add(New SqlParameter("@Ambulancia", SqlDbType.NVarChar))
                consulta.Parameters("@Ambulancia").Value = pRemision.ambulancia
                consulta.Parameters.Add(New SqlParameter("@Codigo_traslado", SqlDbType.NVarChar))
                consulta.Parameters("@Codigo_traslado").Value = pRemision.traslados
                consulta.Parameters.Add(New SqlParameter("@oxigeno", SqlDbType.NVarChar))
                consulta.Parameters("@oxigeno").Value = pRemision.oxigeno
                consulta.Parameters.Add(New SqlParameter("@codigo_especialidad", SqlDbType.Int))
                consulta.Parameters("@codigo_especialidad").Value = pRemision.especialidad
                consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                consulta.Parameters("@Usuario_actualizacion").Value = pRemision.usuarioReal
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarRemisionR(ByRef pRemision As RemisionR)
        Try
            Using consulta = New SqlCommand()
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure

                consulta.CommandText = "[SP_REMISION_ACTUALIZAR_R]"
                consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                consulta.Parameters("@EDITADO").Value = pRemision.editado
                consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                consulta.Parameters("@CODIGO_EP").Value = pRemision.codigoEP

                consulta.Parameters.Add(New SqlParameter("@CODIGO_REM", SqlDbType.Int))
                consulta.Parameters("@CODIGO_REM").Value = pRemision.codigoRemision
                consulta.Parameters.Add(New SqlParameter("@Codigo_Motivo_observacion", SqlDbType.Int))
                consulta.Parameters("@Codigo_Motivo_observacion").Value = pRemision.modalidad
                consulta.Parameters.Add(New SqlParameter("@Datos_medicos", SqlDbType.NVarChar))
                consulta.Parameters("@Datos_medicos").Value = pRemision.datosMedicos
                consulta.Parameters.Add(New SqlParameter("@Fecha_Remision", SqlDbType.DateTime))
                consulta.Parameters("@Fecha_Remision").Value = pRemision.fechaRemision
                consulta.Parameters.Add(New SqlParameter("@Codigo_Prioridad", SqlDbType.Int))
                consulta.Parameters("@Codigo_Prioridad").Value = pRemision.prioridad
                consulta.Parameters.Add(New SqlParameter("@Codigo_complejidad", SqlDbType.Int))
                consulta.Parameters("@Codigo_complejidad").Value = pRemision.complejidad
                consulta.Parameters.Add(New SqlParameter("@Otras", SqlDbType.NVarChar))
                consulta.Parameters("@Otras").Value = pRemision.otras
                consulta.Parameters.Add(New SqlParameter("@Antecedentes", SqlDbType.NVarChar))
                consulta.Parameters("@Antecedentes").Value = pRemision.antecedentes
                consulta.Parameters.Add(New SqlParameter("@Glasgow", SqlDbType.NVarChar))
                consulta.Parameters("@Glasgow").Value = pRemision.glasgow
                consulta.Parameters.Add(New SqlParameter("@Descripcion_Glasgow", SqlDbType.NVarChar))
                consulta.Parameters("@Descripcion_Glasgow").Value = pRemision.descripglasgow
                consulta.Parameters.Add(New SqlParameter("@Presion_Sistolica", SqlDbType.NVarChar))
                consulta.Parameters("@Presion_Sistolica").Value = pRemision.presionsis
                consulta.Parameters.Add(New SqlParameter("@Presion_Diastolica", SqlDbType.NVarChar))
                consulta.Parameters("@Presion_Diastolica").Value = pRemision.presiondias
                consulta.Parameters.Add(New SqlParameter("@Frecuencia_Cardiaca", SqlDbType.NVarChar))
                consulta.Parameters("@Frecuencia_Cardiaca").Value = pRemision.freccar
                consulta.Parameters.Add(New SqlParameter("@Frecuencia_Respiratoria", SqlDbType.NVarChar))
                consulta.Parameters("@Frecuencia_Respiratoria").Value = pRemision.frecresp
                consulta.Parameters.Add(New SqlParameter("@Ambulancia", SqlDbType.NVarChar))
                consulta.Parameters("@Ambulancia").Value = pRemision.ambulancia
                consulta.Parameters.Add(New SqlParameter("@Codigo_traslado", SqlDbType.NVarChar))
                consulta.Parameters("@Codigo_traslado").Value = pRemision.traslados
                consulta.Parameters.Add(New SqlParameter("@oxigeno", SqlDbType.NVarChar))
                consulta.Parameters("@oxigeno").Value = pRemision.oxigeno
                consulta.Parameters.Add(New SqlParameter("@codigo_especialidad", SqlDbType.Int))
                consulta.Parameters("@codigo_especialidad").Value = pRemision.especialidad
                consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                consulta.Parameters("@Usuario_actualizacion").Value = pRemision.usuarioReal
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarRemisionRR(ByRef pRemision As RemisionRR)
        Try
            Using consulta = New SqlCommand()
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandType = CommandType.StoredProcedure

                consulta.CommandText = "[SP_REMISION_ACTUALIZAR_RR]"
                consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                consulta.Parameters("@EDITADO").Value = pRemision.editado

                consulta.Parameters.Add(New SqlParameter("@CODIGO_REM", SqlDbType.Int))
                consulta.Parameters("@CODIGO_REM").Value = pRemision.codigoRemision
                consulta.Parameters.Add(New SqlParameter("@Codigo_Motivo_observacion", SqlDbType.Int))
                consulta.Parameters("@Codigo_Motivo_observacion").Value = pRemision.modalidad
                consulta.Parameters.Add(New SqlParameter("@Datos_medicos", SqlDbType.NVarChar))
                consulta.Parameters("@Datos_medicos").Value = pRemision.datosMedicos
                consulta.Parameters.Add(New SqlParameter("@Fecha_Remision", SqlDbType.DateTime))
                consulta.Parameters("@Fecha_Remision").Value = pRemision.fechaRemision
                consulta.Parameters.Add(New SqlParameter("@Codigo_Prioridad", SqlDbType.Int))
                consulta.Parameters("@Codigo_Prioridad").Value = pRemision.prioridad
                consulta.Parameters.Add(New SqlParameter("@Codigo_complejidad", SqlDbType.Int))
                consulta.Parameters("@Codigo_complejidad").Value = pRemision.complejidad
                consulta.Parameters.Add(New SqlParameter("@Otras", SqlDbType.NVarChar))
                consulta.Parameters("@Otras").Value = pRemision.otras
                consulta.Parameters.Add(New SqlParameter("@Antecedentes", SqlDbType.NVarChar))
                consulta.Parameters("@Antecedentes").Value = pRemision.antecedentes
                consulta.Parameters.Add(New SqlParameter("@Glasgow", SqlDbType.NVarChar))
                consulta.Parameters("@Glasgow").Value = pRemision.glasgow
                consulta.Parameters.Add(New SqlParameter("@Descripcion_Glasgow", SqlDbType.NVarChar))
                consulta.Parameters("@Descripcion_Glasgow").Value = pRemision.descripglasgow
                consulta.Parameters.Add(New SqlParameter("@Presion_Sistolica", SqlDbType.NVarChar))
                consulta.Parameters("@Presion_Sistolica").Value = pRemision.presionsis
                consulta.Parameters.Add(New SqlParameter("@Presion_Diastolica", SqlDbType.NVarChar))
                consulta.Parameters("@Presion_Diastolica").Value = pRemision.presiondias
                consulta.Parameters.Add(New SqlParameter("@Frecuencia_Cardiaca", SqlDbType.NVarChar))
                consulta.Parameters("@Frecuencia_Cardiaca").Value = pRemision.freccar
                consulta.Parameters.Add(New SqlParameter("@Frecuencia_Respiratoria", SqlDbType.NVarChar))
                consulta.Parameters("@Frecuencia_Respiratoria").Value = pRemision.frecresp
                consulta.Parameters.Add(New SqlParameter("@Ambulancia", SqlDbType.NVarChar))
                consulta.Parameters("@Ambulancia").Value = pRemision.ambulancia
                consulta.Parameters.Add(New SqlParameter("@Codigo_traslado", SqlDbType.NVarChar))
                consulta.Parameters("@Codigo_traslado").Value = pRemision.traslados
                consulta.Parameters.Add(New SqlParameter("@oxigeno", SqlDbType.NVarChar))
                consulta.Parameters("@oxigeno").Value = pRemision.oxigeno
                consulta.Parameters.Add(New SqlParameter("@codigo_especialidad", SqlDbType.Int))
                consulta.Parameters("@codigo_especialidad").Value = pRemision.especialidad
                consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                consulta.Parameters("@Usuario_actualizacion").Value = pRemision.usuarioReal
                consulta.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarInfoIngreso(ByRef objInfoIngreso As InfoIngresoAdulto, dtDiagImpresion As DataTable)
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objInfoIngreso.elementoAEliminar
                        If sentencia <> "" Then
                            consulta.CommandText = sentencia
                            consulta.ExecuteNonQuery()
                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()

                    consulta.CommandText = "[SP_INFO_INGRESO_ADULTO_ACTUALIZAR]"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                    consulta.Parameters("@CODIGO_EP").Value = objInfoIngreso.codigoEP

                    consulta.Parameters.Add(New SqlParameter("@PESO", SqlDbType.Float))
                    consulta.Parameters("@PESO").Value = objInfoIngreso.peso
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = objInfoIngreso.nRegistro
                    consulta.Parameters.Add(New SqlParameter("@MOTIVO", SqlDbType.NVarChar))
                    consulta.Parameters("@MOTIVO").Value = objInfoIngreso.motivo
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES").Value = objInfoIngreso.antecedentes
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_QUI", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_QUI").Value = objInfoIngreso.antecedentesQuirurgicos
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_TRANS", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_TRANS").Value = objInfoIngreso.antecedentesTransfusionales
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_ALER", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_ALER").Value = objInfoIngreso.antecedentesAlergicos
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_TRAU", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_TRAU").Value = objInfoIngreso.antecedentesTraumaticos
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_TOX", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_TOX").Value = objInfoIngreso.antecedentesToxicos
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_FAM", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_FAM").Value = objInfoIngreso.antecedentesFamiliares
                    consulta.Parameters.Add(New SqlParameter("@REVISION_SIST", SqlDbType.NVarChar))
                    consulta.Parameters("@REVISION_SIST").Value = objInfoIngreso.revisionSistema
                    consulta.Parameters.Add(New SqlParameter("@SIG_VITALES", SqlDbType.NVarChar))
                    consulta.Parameters("@SIG_VITALES").Value = objInfoIngreso.signosVitales
                    consulta.Parameters.Add(New SqlParameter("@CAB_CUELLO", SqlDbType.NVarChar))
                    consulta.Parameters("@CAB_CUELLO").Value = objInfoIngreso.cabezaYCuello
                    consulta.Parameters.Add(New SqlParameter("@TORAX", SqlDbType.NVarChar))
                    consulta.Parameters("@TORAX").Value = objInfoIngreso.torax
                    consulta.Parameters.Add(New SqlParameter("@CARDIO", SqlDbType.NVarChar))
                    consulta.Parameters("@CARDIO").Value = objInfoIngreso.cardio
                    consulta.Parameters.Add(New SqlParameter("@ABDOMEN", SqlDbType.NVarChar))
                    consulta.Parameters("@ABDOMEN").Value = objInfoIngreso.abdomen
                    consulta.Parameters.Add(New SqlParameter("@GENTURINARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@GENTURINARIO").Value = objInfoIngreso.gentoUrinario
                    consulta.Parameters.Add(New SqlParameter("@EXTREM", SqlDbType.NVarChar))
                    consulta.Parameters("@EXTREM").Value = objInfoIngreso.extremidades
                    consulta.Parameters.Add(New SqlParameter("@S_NERV", SqlDbType.NVarChar))
                    consulta.Parameters("@S_NERV").Value = objInfoIngreso.sistemaNervioso
                    consulta.Parameters.Add(New SqlParameter("@PARACLINICO", SqlDbType.NVarChar))
                    consulta.Parameters("@PARACLINICO").Value = objInfoIngreso.paraclinico
                    consulta.Parameters.Add(New SqlParameter("@ANALISIS", SqlDbType.NVarChar))
                    consulta.Parameters("@ANALISIS").Value = objInfoIngreso.analisis
                    consulta.Parameters.Add(New SqlParameter("@PRONOSTICO", SqlDbType.NVarChar))
                    consulta.Parameters("@PRONOSTICO").Value = objInfoIngreso.pronostico
                    consulta.Parameters.Add(New SqlParameter("@USUARIO_R", SqlDbType.Int))
                    consulta.Parameters("@USUARIO_R").Value = objInfoIngreso.usuarioReal
                    consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                    consulta.Parameters("@Usuario_actualizacion").Value = objInfoIngreso.usuario
                    consulta.ExecuteNonQuery()
                    For I = 0 To dtDiagImpresion.Rows.Count - 1
                        consulta.Parameters.Clear()
                        consulta.CommandText = "[SP_IMPRESION_DIAGNOSTICA_CREAR]"
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                        consulta.Parameters("@REGISTRO").Value = objInfoIngreso.nRegistro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = dtDiagImpresion.Rows(I).Item(0).ToString
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        consulta.Parameters("@OBSERVACION").Value = dtDiagImpresion.Rows(I).Item("Observacion").ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EP").Value = objInfoIngreso.codigoEP
                        consulta.ExecuteNonQuery()
                    Next
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarInfoIngresoR(ByRef objInfoIngreso As InfoIngresoAdultoR, dtDiagImpresion As DataTable)
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objInfoIngreso.elementoAEliminar
                        If sentencia <> "" Then
                            consulta.CommandText = sentencia
                            consulta.ExecuteNonQuery()
                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()
                    consulta.CommandText = "[SP_INFO_INGRESO_ADULTO_ACTUALIZAR_R]"
                    consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                    consulta.Parameters("@EDITADO").Value = 1
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                    consulta.Parameters("@CODIGO_EP").Value = objInfoIngreso.codigoEP
                    consulta.Parameters.Add(New SqlParameter("@PESO", SqlDbType.Float))
                    consulta.Parameters("@PESO").Value = objInfoIngreso.peso
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = objInfoIngreso.nRegistro
                    consulta.Parameters.Add(New SqlParameter("@MOTIVO", SqlDbType.NVarChar))
                    consulta.Parameters("@MOTIVO").Value = objInfoIngreso.motivo
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES").Value = objInfoIngreso.antecedentes
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_QUI", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_QUI").Value = objInfoIngreso.antecedentesQuirurgicos
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_TRANS", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_TRANS").Value = objInfoIngreso.antecedentesTransfusionales
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_ALER", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_ALER").Value = objInfoIngreso.antecedentesAlergicos
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_TRAU", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_TRAU").Value = objInfoIngreso.antecedentesTraumaticos
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_TOX", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_TOX").Value = objInfoIngreso.antecedentesToxicos
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_FAM", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_FAM").Value = objInfoIngreso.antecedentesFamiliares
                    consulta.Parameters.Add(New SqlParameter("@REVISION_SIST", SqlDbType.NVarChar))
                    consulta.Parameters("@REVISION_SIST").Value = objInfoIngreso.revisionSistema
                    consulta.Parameters.Add(New SqlParameter("@SIG_VITALES", SqlDbType.NVarChar))
                    consulta.Parameters("@SIG_VITALES").Value = objInfoIngreso.signosVitales
                    consulta.Parameters.Add(New SqlParameter("@CAB_CUELLO", SqlDbType.NVarChar))
                    consulta.Parameters("@CAB_CUELLO").Value = objInfoIngreso.cabezaYCuello
                    consulta.Parameters.Add(New SqlParameter("@TORAX", SqlDbType.NVarChar))
                    consulta.Parameters("@TORAX").Value = objInfoIngreso.torax
                    consulta.Parameters.Add(New SqlParameter("@CARDIO", SqlDbType.NVarChar))
                    consulta.Parameters("@CARDIO").Value = objInfoIngreso.cardio
                    consulta.Parameters.Add(New SqlParameter("@ABDOMEN", SqlDbType.NVarChar))
                    consulta.Parameters("@ABDOMEN").Value = objInfoIngreso.abdomen
                    consulta.Parameters.Add(New SqlParameter("@GENTURINARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@GENTURINARIO").Value = objInfoIngreso.gentoUrinario
                    consulta.Parameters.Add(New SqlParameter("@EXTREM", SqlDbType.NVarChar))
                    consulta.Parameters("@EXTREM").Value = objInfoIngreso.extremidades
                    consulta.Parameters.Add(New SqlParameter("@S_NERV", SqlDbType.NVarChar))
                    consulta.Parameters("@S_NERV").Value = objInfoIngreso.sistemaNervioso
                    consulta.Parameters.Add(New SqlParameter("@PARACLINICO", SqlDbType.NVarChar))
                    consulta.Parameters("@PARACLINICO").Value = objInfoIngreso.paraclinico
                    consulta.Parameters.Add(New SqlParameter("@ANALISIS", SqlDbType.NVarChar))
                    consulta.Parameters("@ANALISIS").Value = objInfoIngreso.analisis
                    consulta.Parameters.Add(New SqlParameter("@PRONOSTICO", SqlDbType.NVarChar))
                    consulta.Parameters("@PRONOSTICO").Value = objInfoIngreso.pronostico
                    consulta.Parameters.Add(New SqlParameter("@USUARIO_R", SqlDbType.Int))
                    consulta.Parameters("@USUARIO_R").Value = objInfoIngreso.usuarioReal
                    consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                    consulta.Parameters("@Usuario_actualizacion").Value = objInfoIngreso.usuario
                    consulta.ExecuteNonQuery()
                    For I = 0 To dtDiagImpresion.Rows.Count - 1
                        consulta.Parameters.Clear()
                        consulta.CommandText = "[SP_IMPRESION_DIAGNOSTICA_CREAR_R]"
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                        consulta.Parameters("@REGISTRO").Value = objInfoIngreso.nRegistro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = dtDiagImpresion.Rows(I).Item(0).ToString
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        consulta.Parameters("@OBSERVACION").Value = dtDiagImpresion.Rows(I).Item("Observacion").ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EP").Value = objInfoIngreso.codigoEP
                        consulta.ExecuteNonQuery()
                    Next
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub guardarInfoIngresoRR(ByRef objInfoIngreso As InfoIngresoAdultoRR, dtDiagImpresion As DataTable)
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objInfoIngreso.elementoAEliminar
                        If sentencia <> "" Then
                            consulta.CommandText = sentencia
                            consulta.ExecuteNonQuery()
                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()

                    consulta.CommandText = "[SP_INFO_INGRESO_ADULTO_ACTUALIZAR_RR]"
                    consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                    consulta.Parameters("@EDITADO").Value = 1

                    consulta.Parameters.Add(New SqlParameter("@PESO", SqlDbType.Float))
                    consulta.Parameters("@PESO").Value = objInfoIngreso.peso
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = objInfoIngreso.nRegistro
                    consulta.Parameters.Add(New SqlParameter("@MOTIVO", SqlDbType.NVarChar))
                    consulta.Parameters("@MOTIVO").Value = objInfoIngreso.motivo
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES").Value = objInfoIngreso.antecedentes
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_QUI", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_QUI").Value = objInfoIngreso.antecedentesQuirurgicos
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_TRANS", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_TRANS").Value = objInfoIngreso.antecedentesTransfusionales
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_ALER", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_ALER").Value = objInfoIngreso.antecedentesAlergicos
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_TRAU", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_TRAU").Value = objInfoIngreso.antecedentesTraumaticos
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_TOX", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_TOX").Value = objInfoIngreso.antecedentesToxicos
                    consulta.Parameters.Add(New SqlParameter("@ANTECEDENTES_FAM", SqlDbType.NVarChar))
                    consulta.Parameters("@ANTECEDENTES_FAM").Value = objInfoIngreso.antecedentesFamiliares
                    consulta.Parameters.Add(New SqlParameter("@REVISION_SIST", SqlDbType.NVarChar))
                    consulta.Parameters("@REVISION_SIST").Value = objInfoIngreso.revisionSistema
                    consulta.Parameters.Add(New SqlParameter("@SIG_VITALES", SqlDbType.NVarChar))
                    consulta.Parameters("@SIG_VITALES").Value = objInfoIngreso.signosVitales
                    consulta.Parameters.Add(New SqlParameter("@CAB_CUELLO", SqlDbType.NVarChar))
                    consulta.Parameters("@CAB_CUELLO").Value = objInfoIngreso.cabezaYCuello
                    consulta.Parameters.Add(New SqlParameter("@TORAX", SqlDbType.NVarChar))
                    consulta.Parameters("@TORAX").Value = objInfoIngreso.torax
                    consulta.Parameters.Add(New SqlParameter("@CARDIO", SqlDbType.NVarChar))
                    consulta.Parameters("@CARDIO").Value = objInfoIngreso.cardio
                    consulta.Parameters.Add(New SqlParameter("@ABDOMEN", SqlDbType.NVarChar))
                    consulta.Parameters("@ABDOMEN").Value = objInfoIngreso.abdomen
                    consulta.Parameters.Add(New SqlParameter("@GENTURINARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@GENTURINARIO").Value = objInfoIngreso.gentoUrinario
                    consulta.Parameters.Add(New SqlParameter("@EXTREM", SqlDbType.NVarChar))
                    consulta.Parameters("@EXTREM").Value = objInfoIngreso.extremidades
                    consulta.Parameters.Add(New SqlParameter("@S_NERV", SqlDbType.NVarChar))
                    consulta.Parameters("@S_NERV").Value = objInfoIngreso.sistemaNervioso
                    consulta.Parameters.Add(New SqlParameter("@PARACLINICO", SqlDbType.NVarChar))
                    consulta.Parameters("@PARACLINICO").Value = objInfoIngreso.paraclinico
                    consulta.Parameters.Add(New SqlParameter("@ANALISIS", SqlDbType.NVarChar))
                    consulta.Parameters("@ANALISIS").Value = objInfoIngreso.analisis
                    consulta.Parameters.Add(New SqlParameter("@PRONOSTICO", SqlDbType.NVarChar))
                    consulta.Parameters("@PRONOSTICO").Value = objInfoIngreso.pronostico
                    consulta.Parameters.Add(New SqlParameter("@USUARIO_R", SqlDbType.Int))
                    consulta.Parameters("@USUARIO_R").Value = objInfoIngreso.usuarioReal
                    consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                    consulta.Parameters("@Usuario_actualizacion").Value = objInfoIngreso.usuario
                    consulta.ExecuteNonQuery()
                    For I = 0 To dtDiagImpresion.Rows.Count - 1
                        consulta.Parameters.Clear()
                        consulta.CommandText = "[SP_IMPRESION_DIAGNOSTICA_CREAR_RR]"
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                        consulta.Parameters("@REGISTRO").Value = objInfoIngreso.nRegistro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = dtDiagImpresion.Rows(I).Item(0).ToString
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        consulta.Parameters("@OBSERVACION").Value = dtDiagImpresion.Rows(I).Item("Observacion").ToString
                        consulta.ExecuteNonQuery()
                    Next
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarInfoIngresoN(ByRef objInfoIngreso As InfoIngresoNeonato, dtDiagImpresion As DataTable)
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objInfoIngreso.elementoAEliminar
                        If sentencia <> "" Then
                            consulta.CommandText = sentencia
                            consulta.ExecuteNonQuery()
                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()
                    consulta.CommandText = "[SP_INFO_INGRESO_NEONATO_ACTUALIZAR]"
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                    consulta.Parameters("@CODIGO_EP").Value = objInfoIngreso.codigoEP
                    consulta.Parameters.Add(New SqlParameter("@PESO", SqlDbType.Float))
                    consulta.Parameters("@PESO").Value = objInfoIngreso.peso
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = objInfoIngreso.nRegistro
                    consulta.Parameters.Add(New SqlParameter("@MOTIVO", SqlDbType.NVarChar))
                    consulta.Parameters("@MOTIVO").Value = objInfoIngreso.motivo
                    consulta.Parameters.Add(New SqlParameter("@GENERALES", SqlDbType.NVarChar))
                    consulta.Parameters("@GENERALES").Value = objInfoIngreso.generales
                    consulta.Parameters.Add(New SqlParameter("@SIG_VITALES", SqlDbType.NVarChar))
                    consulta.Parameters("@SIG_VITALES").Value = objInfoIngreso.signosVitales
                    consulta.Parameters.Add(New SqlParameter("@CAB_CUELLO", SqlDbType.NVarChar))
                    consulta.Parameters("@CAB_CUELLO").Value = objInfoIngreso.cabezaYCuello
                    consulta.Parameters.Add(New SqlParameter("@TORAX", SqlDbType.NVarChar))
                    consulta.Parameters("@TORAX").Value = objInfoIngreso.torax
                    consulta.Parameters.Add(New SqlParameter("@CARDIO", SqlDbType.NVarChar))
                    consulta.Parameters("@CARDIO").Value = objInfoIngreso.cardio
                    consulta.Parameters.Add(New SqlParameter("@abdomen", SqlDbType.NVarChar))
                    consulta.Parameters("@abdomen").Value = objInfoIngreso.abdomen
                    consulta.Parameters.Add(New SqlParameter("@GENTURINARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@GENTURINARIO").Value = objInfoIngreso.gentoUrinario
                    consulta.Parameters.Add(New SqlParameter("@extrem", SqlDbType.NVarChar))
                    consulta.Parameters("@extrem").Value = objInfoIngreso.extremidades
                    consulta.Parameters.Add(New SqlParameter("@s_nerv", SqlDbType.NVarChar))
                    consulta.Parameters("@s_nerv").Value = objInfoIngreso.sistemaNervioso
                    consulta.Parameters.Add(New SqlParameter("@t_parto", SqlDbType.NVarChar))
                    consulta.Parameters("@t_parto").Value = objInfoIngreso.tParto
                    consulta.Parameters.Add(New SqlParameter("@t_rupt_m", SqlDbType.NVarChar))
                    consulta.Parameters("@t_rupt_m").Value = objInfoIngreso.tRupturaM
                    consulta.Parameters.Add(New SqlParameter("@indu_parto", SqlDbType.NVarChar))
                    consulta.Parameters("@indu_parto").Value = objInfoIngreso.induccionParto
                    consulta.Parameters.Add(New SqlParameter("@caract_liquidas", SqlDbType.NVarChar))
                    consulta.Parameters("@caract_liquidas").Value = objInfoIngreso.caracteristicasLiquidas
                    consulta.Parameters.Add(New SqlParameter("@apgar1", SqlDbType.NVarChar))
                    consulta.Parameters("@apgar1").Value = objInfoIngreso.apgar1
                    consulta.Parameters.Add(New SqlParameter("@apgar5", SqlDbType.NVarChar))
                    consulta.Parameters("@apgar5").Value = objInfoIngreso.apgar5
                    consulta.Parameters.Add(New SqlParameter("@reanim", SqlDbType.NVarChar))
                    consulta.Parameters("@reanim").Value = objInfoIngreso.reanimacion
                    consulta.Parameters.Add(New SqlParameter("@edad_madre", SqlDbType.NVarChar))
                    consulta.Parameters("@edad_madre").Value = objInfoIngreso.edadMadre
                    consulta.Parameters.Add(New SqlParameter("@ant_obst", SqlDbType.NVarChar))
                    consulta.Parameters("@ant_obst").Value = objInfoIngreso.antecedentesObstetricos
                    consulta.Parameters.Add(New SqlParameter("@fum", SqlDbType.NVarChar))
                    consulta.Parameters("@fum").Value = objInfoIngreso.fumador
                    consulta.Parameters.Add(New SqlParameter("@edad_gesta", SqlDbType.NVarChar))
                    consulta.Parameters("@edad_gesta").Value = objInfoIngreso.edadGestacional
                    consulta.Parameters.Add(New SqlParameter("@hemoc_m", SqlDbType.NVarChar))
                    consulta.Parameters("@hemoc_m").Value = objInfoIngreso.hemoclasificacionMadre
                    consulta.Parameters.Add(New SqlParameter("@hemoc_p", SqlDbType.NVarChar))
                    consulta.Parameters("@hemoc_p").Value = objInfoIngreso.hemoclasificacionPadre
                    consulta.Parameters.Add(New SqlParameter("@control", SqlDbType.NVarChar))
                    consulta.Parameters("@control").Value = objInfoIngreso.control
                    consulta.Parameters.Add(New SqlParameter("@medc", SqlDbType.NVarChar))
                    consulta.Parameters("@medc").Value = objInfoIngreso.medicamentos
                    consulta.Parameters.Add(New SqlParameter("@habito", SqlDbType.NVarChar))
                    consulta.Parameters("@habito").Value = objInfoIngreso.habito
                    consulta.Parameters.Add(New SqlParameter("@infecc", SqlDbType.NVarChar))
                    consulta.Parameters("@infecc").Value = objInfoIngreso.infeccion
                    consulta.Parameters.Add(New SqlParameter("@diabete_g", SqlDbType.NVarChar))
                    consulta.Parameters("@diabete_g").Value = objInfoIngreso.diabeteG
                    consulta.Parameters.Add(New SqlParameter("@diabete_m", SqlDbType.NVarChar))
                    consulta.Parameters("@diabete_m").Value = objInfoIngreso.diabeteM
                    consulta.Parameters.Add(New SqlParameter("@hipertencion", SqlDbType.NVarChar))
                    consulta.Parameters("@hipertencion").Value = objInfoIngreso.hipertencion
                    consulta.Parameters.Add(New SqlParameter("@preclampcia", SqlDbType.NVarChar))
                    consulta.Parameters("@preclampcia").Value = objInfoIngreso.preclampcia
                    consulta.Parameters.Add(New SqlParameter("@enfer", SqlDbType.NVarChar))
                    consulta.Parameters("@enfer").Value = objInfoIngreso.enfermedad
                    consulta.Parameters.Add(New SqlParameter("@vacunacion", SqlDbType.NVarChar))
                    consulta.Parameters("@vacunacion").Value = objInfoIngreso.vacunacion
                    consulta.Parameters.Add(New SqlParameter("@torch", SqlDbType.NVarChar))
                    consulta.Parameters("@torch").Value = objInfoIngreso.torch
                    consulta.Parameters.Add(New SqlParameter("@hemoclasificacion", SqlDbType.NVarChar))
                    consulta.Parameters("@hemoclasificacion").Value = objInfoIngreso.hemoclasificacion
                    consulta.Parameters.Add(New SqlParameter("@vdrl", SqlDbType.NVarChar))
                    consulta.Parameters("@vdrl").Value = objInfoIngreso.vdrl
                    consulta.Parameters.Add(New SqlParameter("@tsh", SqlDbType.NVarChar))
                    consulta.Parameters("@tsh").Value = objInfoIngreso.tsh
                    consulta.Parameters.Add(New SqlParameter("@glucometria", SqlDbType.NVarChar))
                    consulta.Parameters("@glucometria").Value = objInfoIngreso.glucometria
                    consulta.Parameters.Add(New SqlParameter("@analisis", SqlDbType.NVarChar))
                    consulta.Parameters("@analisis").Value = objInfoIngreso.analisis
                    consulta.Parameters.Add(New SqlParameter("@pronostico", SqlDbType.NVarChar))
                    consulta.Parameters("@pronostico").Value = objInfoIngreso.pronostico
                    consulta.Parameters.Add(New SqlParameter("@USUARIO_R", SqlDbType.Int))
                    consulta.Parameters("@USUARIO_R").Value = objInfoIngreso.usuarioReal
                    consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                    consulta.Parameters("@Usuario_actualizacion").Value = objInfoIngreso.usuario
                    consulta.ExecuteNonQuery()
                    For I = 0 To dtDiagImpresion.Rows.Count - 1
                        consulta.Parameters.Clear()

                        consulta.CommandText = "[SP_IMPRESION_DIAGNOSTICA_CREAR]"
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                        consulta.Parameters("@REGISTRO").Value = objInfoIngreso.nRegistro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = dtDiagImpresion.Rows(I).Item(0).ToString
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        consulta.Parameters("@OBSERVACION").Value = dtDiagImpresion.Rows(I).Item("Observacion").ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EP").Value = objInfoIngreso.codigoEP

                        consulta.ExecuteNonQuery()
                    Next
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub guardarInfoIngresoNR(ByRef objInfoIngreso As InfoIngresoNeonatoR, dtDiagImpresion As DataTable)
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objInfoIngreso.elementoAEliminar
                        If sentencia <> "" Then
                            consulta.CommandText = sentencia
                            consulta.ExecuteNonQuery()
                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()
                    consulta.CommandText = "[SP_INFO_INGRESO_NEONATO_ACTUALIZAR_R]"
                    consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                    consulta.Parameters("@EDITADO").Value = 1
                    consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                    consulta.Parameters("@CODIGO_EP").Value = objInfoIngreso.codigoEP
                    consulta.Parameters.Add(New SqlParameter("@PESO", SqlDbType.Float))
                    consulta.Parameters("@PESO").Value = objInfoIngreso.peso
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = objInfoIngreso.nRegistro
                    consulta.Parameters.Add(New SqlParameter("@MOTIVO", SqlDbType.NVarChar))
                    consulta.Parameters("@MOTIVO").Value = objInfoIngreso.motivo
                    consulta.Parameters.Add(New SqlParameter("@GENERALES", SqlDbType.NVarChar))
                    consulta.Parameters("@GENERALES").Value = objInfoIngreso.generales
                    consulta.Parameters.Add(New SqlParameter("@SIG_VITALES", SqlDbType.NVarChar))
                    consulta.Parameters("@SIG_VITALES").Value = objInfoIngreso.signosVitales
                    consulta.Parameters.Add(New SqlParameter("@CAB_CUELLO", SqlDbType.NVarChar))
                    consulta.Parameters("@CAB_CUELLO").Value = objInfoIngreso.cabezaYCuello
                    consulta.Parameters.Add(New SqlParameter("@TORAX", SqlDbType.NVarChar))
                    consulta.Parameters("@TORAX").Value = objInfoIngreso.torax
                    consulta.Parameters.Add(New SqlParameter("@CARDIO", SqlDbType.NVarChar))
                    consulta.Parameters("@CARDIO").Value = objInfoIngreso.cardio
                    consulta.Parameters.Add(New SqlParameter("@abdomen", SqlDbType.NVarChar))
                    consulta.Parameters("@abdomen").Value = objInfoIngreso.abdomen
                    consulta.Parameters.Add(New SqlParameter("@GENTURINARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@GENTURINARIO").Value = objInfoIngreso.gentoUrinario
                    consulta.Parameters.Add(New SqlParameter("@extrem", SqlDbType.NVarChar))
                    consulta.Parameters("@extrem").Value = objInfoIngreso.extremidades
                    consulta.Parameters.Add(New SqlParameter("@s_nerv", SqlDbType.NVarChar))
                    consulta.Parameters("@s_nerv").Value = objInfoIngreso.sistemaNervioso
                    consulta.Parameters.Add(New SqlParameter("@t_parto", SqlDbType.NVarChar))
                    consulta.Parameters("@t_parto").Value = objInfoIngreso.tParto
                    consulta.Parameters.Add(New SqlParameter("@t_rupt_m", SqlDbType.NVarChar))
                    consulta.Parameters("@t_rupt_m").Value = objInfoIngreso.tRupturaM
                    consulta.Parameters.Add(New SqlParameter("@indu_parto", SqlDbType.NVarChar))
                    consulta.Parameters("@indu_parto").Value = objInfoIngreso.induccionParto
                    consulta.Parameters.Add(New SqlParameter("@caract_liquidas", SqlDbType.NVarChar))
                    consulta.Parameters("@caract_liquidas").Value = objInfoIngreso.caracteristicasLiquidas
                    consulta.Parameters.Add(New SqlParameter("@apgar1", SqlDbType.NVarChar))
                    consulta.Parameters("@apgar1").Value = objInfoIngreso.apgar1
                    consulta.Parameters.Add(New SqlParameter("@apgar5", SqlDbType.NVarChar))
                    consulta.Parameters("@apgar5").Value = objInfoIngreso.apgar5
                    consulta.Parameters.Add(New SqlParameter("@reanim", SqlDbType.NVarChar))
                    consulta.Parameters("@reanim").Value = objInfoIngreso.reanimacion
                    consulta.Parameters.Add(New SqlParameter("@edad_madre", SqlDbType.NVarChar))
                    consulta.Parameters("@edad_madre").Value = objInfoIngreso.edadMadre
                    consulta.Parameters.Add(New SqlParameter("@ant_obst", SqlDbType.NVarChar))
                    consulta.Parameters("@ant_obst").Value = objInfoIngreso.antecedentesObstetricos
                    consulta.Parameters.Add(New SqlParameter("@fum", SqlDbType.NVarChar))
                    consulta.Parameters("@fum").Value = objInfoIngreso.fumador
                    consulta.Parameters.Add(New SqlParameter("@edad_gesta", SqlDbType.NVarChar))
                    consulta.Parameters("@edad_gesta").Value = objInfoIngreso.edadGestacional
                    consulta.Parameters.Add(New SqlParameter("@hemoc_m", SqlDbType.NVarChar))
                    consulta.Parameters("@hemoc_m").Value = objInfoIngreso.hemoclasificacionMadre
                    consulta.Parameters.Add(New SqlParameter("@hemoc_p", SqlDbType.NVarChar))
                    consulta.Parameters("@hemoc_p").Value = objInfoIngreso.hemoclasificacionPadre
                    consulta.Parameters.Add(New SqlParameter("@control", SqlDbType.NVarChar))
                    consulta.Parameters("@control").Value = objInfoIngreso.control
                    consulta.Parameters.Add(New SqlParameter("@medc", SqlDbType.NVarChar))
                    consulta.Parameters("@medc").Value = objInfoIngreso.medicamentos
                    consulta.Parameters.Add(New SqlParameter("@habito", SqlDbType.NVarChar))
                    consulta.Parameters("@habito").Value = objInfoIngreso.habito
                    consulta.Parameters.Add(New SqlParameter("@infecc", SqlDbType.NVarChar))
                    consulta.Parameters("@infecc").Value = objInfoIngreso.infeccion
                    consulta.Parameters.Add(New SqlParameter("@diabete_g", SqlDbType.NVarChar))
                    consulta.Parameters("@diabete_g").Value = objInfoIngreso.diabeteG
                    consulta.Parameters.Add(New SqlParameter("@diabete_m", SqlDbType.NVarChar))
                    consulta.Parameters("@diabete_m").Value = objInfoIngreso.diabeteM
                    consulta.Parameters.Add(New SqlParameter("@hipertencion", SqlDbType.NVarChar))
                    consulta.Parameters("@hipertencion").Value = objInfoIngreso.hipertencion
                    consulta.Parameters.Add(New SqlParameter("@preclampcia", SqlDbType.NVarChar))
                    consulta.Parameters("@preclampcia").Value = objInfoIngreso.preclampcia
                    consulta.Parameters.Add(New SqlParameter("@enfer", SqlDbType.NVarChar))
                    consulta.Parameters("@enfer").Value = objInfoIngreso.enfermedad
                    consulta.Parameters.Add(New SqlParameter("@vacunacion", SqlDbType.NVarChar))
                    consulta.Parameters("@vacunacion").Value = objInfoIngreso.vacunacion
                    consulta.Parameters.Add(New SqlParameter("@torch", SqlDbType.NVarChar))
                    consulta.Parameters("@torch").Value = objInfoIngreso.torch
                    consulta.Parameters.Add(New SqlParameter("@hemoclasificacion", SqlDbType.NVarChar))
                    consulta.Parameters("@hemoclasificacion").Value = objInfoIngreso.hemoclasificacion
                    consulta.Parameters.Add(New SqlParameter("@vdrl", SqlDbType.NVarChar))
                    consulta.Parameters("@vdrl").Value = objInfoIngreso.vdrl
                    consulta.Parameters.Add(New SqlParameter("@tsh", SqlDbType.NVarChar))
                    consulta.Parameters("@tsh").Value = objInfoIngreso.tsh
                    consulta.Parameters.Add(New SqlParameter("@glucometria", SqlDbType.NVarChar))
                    consulta.Parameters("@glucometria").Value = objInfoIngreso.glucometria
                    consulta.Parameters.Add(New SqlParameter("@analisis", SqlDbType.NVarChar))
                    consulta.Parameters("@analisis").Value = objInfoIngreso.analisis
                    consulta.Parameters.Add(New SqlParameter("@pronostico", SqlDbType.NVarChar))
                    consulta.Parameters("@pronostico").Value = objInfoIngreso.pronostico
                    consulta.Parameters.Add(New SqlParameter("@USUARIO_R", SqlDbType.Int))
                    consulta.Parameters("@USUARIO_R").Value = objInfoIngreso.usuarioReal
                    consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                    consulta.Parameters("@Usuario_actualizacion").Value = objInfoIngreso.usuario
                    consulta.ExecuteNonQuery()
                    For I = 0 To dtDiagImpresion.Rows.Count - 1
                        consulta.Parameters.Clear()

                        consulta.CommandText = "[SP_IMPRESION_DIAGNOSTICA_CREAR]"
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                        consulta.Parameters("@REGISTRO").Value = objInfoIngreso.nRegistro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = dtDiagImpresion.Rows(I).Item(0).ToString
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        consulta.Parameters("@OBSERVACION").Value = dtDiagImpresion.Rows(I).Item("Observacion").ToString
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_EP").Value = objInfoIngreso.codigoEP

                        consulta.ExecuteNonQuery()
                    Next
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub guardarInfoIngresoNRR(ByRef objInfoIngreso As InfoIngresoNeonatoRR, dtDiagImpresion As DataTable)
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For Each sentencia As String In objInfoIngreso.elementoAEliminar
                        If sentencia <> "" Then
                            consulta.CommandText = sentencia
                            consulta.ExecuteNonQuery()
                        End If
                    Next
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()
                    consulta.CommandText = "[SP_INFO_INGRESO_NEONATO_ACTUALIZAR_RR]"
                    consulta.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit))
                    consulta.Parameters("@EDITADO").Value = 1
                    consulta.Parameters.Add(New SqlParameter("@PESO", SqlDbType.Float))
                    consulta.Parameters("@PESO").Value = objInfoIngreso.peso
                    consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                    consulta.Parameters("@REGISTRO").Value = objInfoIngreso.nRegistro
                    consulta.Parameters.Add(New SqlParameter("@MOTIVO", SqlDbType.NVarChar))
                    consulta.Parameters("@MOTIVO").Value = objInfoIngreso.motivo
                    consulta.Parameters.Add(New SqlParameter("@GENERALES", SqlDbType.NVarChar))
                    consulta.Parameters("@GENERALES").Value = objInfoIngreso.generales
                    consulta.Parameters.Add(New SqlParameter("@SIG_VITALES", SqlDbType.NVarChar))
                    consulta.Parameters("@SIG_VITALES").Value = objInfoIngreso.signosVitales
                    consulta.Parameters.Add(New SqlParameter("@CAB_CUELLO", SqlDbType.NVarChar))
                    consulta.Parameters("@CAB_CUELLO").Value = objInfoIngreso.cabezaYCuello
                    consulta.Parameters.Add(New SqlParameter("@TORAX", SqlDbType.NVarChar))
                    consulta.Parameters("@TORAX").Value = objInfoIngreso.torax
                    consulta.Parameters.Add(New SqlParameter("@CARDIO", SqlDbType.NVarChar))
                    consulta.Parameters("@CARDIO").Value = objInfoIngreso.cardio
                    consulta.Parameters.Add(New SqlParameter("@abdomen", SqlDbType.NVarChar))
                    consulta.Parameters("@abdomen").Value = objInfoIngreso.abdomen
                    consulta.Parameters.Add(New SqlParameter("@GENTURINARIO", SqlDbType.NVarChar))
                    consulta.Parameters("@GENTURINARIO").Value = objInfoIngreso.gentoUrinario
                    consulta.Parameters.Add(New SqlParameter("@extrem", SqlDbType.NVarChar))
                    consulta.Parameters("@extrem").Value = objInfoIngreso.extremidades
                    consulta.Parameters.Add(New SqlParameter("@s_nerv", SqlDbType.NVarChar))
                    consulta.Parameters("@s_nerv").Value = objInfoIngreso.sistemaNervioso
                    consulta.Parameters.Add(New SqlParameter("@t_parto", SqlDbType.NVarChar))
                    consulta.Parameters("@t_parto").Value = objInfoIngreso.tParto
                    consulta.Parameters.Add(New SqlParameter("@t_rupt_m", SqlDbType.NVarChar))
                    consulta.Parameters("@t_rupt_m").Value = objInfoIngreso.tRupturaM
                    consulta.Parameters.Add(New SqlParameter("@indu_parto", SqlDbType.NVarChar))
                    consulta.Parameters("@indu_parto").Value = objInfoIngreso.induccionParto
                    consulta.Parameters.Add(New SqlParameter("@caract_liquidas", SqlDbType.NVarChar))
                    consulta.Parameters("@caract_liquidas").Value = objInfoIngreso.caracteristicasLiquidas
                    consulta.Parameters.Add(New SqlParameter("@apgar1", SqlDbType.NVarChar))
                    consulta.Parameters("@apgar1").Value = objInfoIngreso.apgar1
                    consulta.Parameters.Add(New SqlParameter("@apgar5", SqlDbType.NVarChar))
                    consulta.Parameters("@apgar5").Value = objInfoIngreso.apgar5
                    consulta.Parameters.Add(New SqlParameter("@reanim", SqlDbType.NVarChar))
                    consulta.Parameters("@reanim").Value = objInfoIngreso.reanimacion
                    consulta.Parameters.Add(New SqlParameter("@edad_madre", SqlDbType.NVarChar))
                    consulta.Parameters("@edad_madre").Value = objInfoIngreso.edadMadre
                    consulta.Parameters.Add(New SqlParameter("@ant_obst", SqlDbType.NVarChar))
                    consulta.Parameters("@ant_obst").Value = objInfoIngreso.antecedentesObstetricos
                    consulta.Parameters.Add(New SqlParameter("@fum", SqlDbType.NVarChar))
                    consulta.Parameters("@fum").Value = objInfoIngreso.fumador
                    consulta.Parameters.Add(New SqlParameter("@edad_gesta", SqlDbType.NVarChar))
                    consulta.Parameters("@edad_gesta").Value = objInfoIngreso.edadGestacional
                    consulta.Parameters.Add(New SqlParameter("@hemoc_m", SqlDbType.NVarChar))
                    consulta.Parameters("@hemoc_m").Value = objInfoIngreso.hemoclasificacionMadre
                    consulta.Parameters.Add(New SqlParameter("@hemoc_p", SqlDbType.NVarChar))
                    consulta.Parameters("@hemoc_p").Value = objInfoIngreso.hemoclasificacionPadre
                    consulta.Parameters.Add(New SqlParameter("@control", SqlDbType.NVarChar))
                    consulta.Parameters("@control").Value = objInfoIngreso.control
                    consulta.Parameters.Add(New SqlParameter("@medc", SqlDbType.NVarChar))
                    consulta.Parameters("@medc").Value = objInfoIngreso.medicamentos
                    consulta.Parameters.Add(New SqlParameter("@habito", SqlDbType.NVarChar))
                    consulta.Parameters("@habito").Value = objInfoIngreso.habito
                    consulta.Parameters.Add(New SqlParameter("@infecc", SqlDbType.NVarChar))
                    consulta.Parameters("@infecc").Value = objInfoIngreso.infeccion
                    consulta.Parameters.Add(New SqlParameter("@diabete_g", SqlDbType.NVarChar))
                    consulta.Parameters("@diabete_g").Value = objInfoIngreso.diabeteG
                    consulta.Parameters.Add(New SqlParameter("@diabete_m", SqlDbType.NVarChar))
                    consulta.Parameters("@diabete_m").Value = objInfoIngreso.diabeteM
                    consulta.Parameters.Add(New SqlParameter("@hipertencion", SqlDbType.NVarChar))
                    consulta.Parameters("@hipertencion").Value = objInfoIngreso.hipertencion
                    consulta.Parameters.Add(New SqlParameter("@preclampcia", SqlDbType.NVarChar))
                    consulta.Parameters("@preclampcia").Value = objInfoIngreso.preclampcia
                    consulta.Parameters.Add(New SqlParameter("@enfer", SqlDbType.NVarChar))
                    consulta.Parameters("@enfer").Value = objInfoIngreso.enfermedad
                    consulta.Parameters.Add(New SqlParameter("@vacunacion", SqlDbType.NVarChar))
                    consulta.Parameters("@vacunacion").Value = objInfoIngreso.vacunacion
                    consulta.Parameters.Add(New SqlParameter("@torch", SqlDbType.NVarChar))
                    consulta.Parameters("@torch").Value = objInfoIngreso.torch
                    consulta.Parameters.Add(New SqlParameter("@hemoclasificacion", SqlDbType.NVarChar))
                    consulta.Parameters("@hemoclasificacion").Value = objInfoIngreso.hemoclasificacion
                    consulta.Parameters.Add(New SqlParameter("@vdrl", SqlDbType.NVarChar))
                    consulta.Parameters("@vdrl").Value = objInfoIngreso.vdrl
                    consulta.Parameters.Add(New SqlParameter("@tsh", SqlDbType.NVarChar))
                    consulta.Parameters("@tsh").Value = objInfoIngreso.tsh
                    consulta.Parameters.Add(New SqlParameter("@glucometria", SqlDbType.NVarChar))
                    consulta.Parameters("@glucometria").Value = objInfoIngreso.glucometria
                    consulta.Parameters.Add(New SqlParameter("@analisis", SqlDbType.NVarChar))
                    consulta.Parameters("@analisis").Value = objInfoIngreso.analisis
                    consulta.Parameters.Add(New SqlParameter("@pronostico", SqlDbType.NVarChar))
                    consulta.Parameters("@pronostico").Value = objInfoIngreso.pronostico
                    consulta.Parameters.Add(New SqlParameter("@USUARIO_R", SqlDbType.Int))
                    consulta.Parameters("@USUARIO_R").Value = objInfoIngreso.usuarioReal
                    consulta.Parameters.Add(New SqlParameter("@Usuario_actualizacion", SqlDbType.Int))
                    consulta.Parameters("@Usuario_actualizacion").Value = objInfoIngreso.usuario
                    consulta.ExecuteNonQuery()
                    For I = 0 To dtDiagImpresion.Rows.Count - 1
                        consulta.Parameters.Clear()
                        consulta.CommandText = "[SP_IMPRESION_DIAGNOSTICA_CREAR_RR]"
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                        consulta.Parameters("@REGISTRO").Value = objInfoIngreso.nRegistro
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_CIE", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_CIE").Value = dtDiagImpresion.Rows(I).Item(0).ToString
                        consulta.Parameters.Add(New SqlParameter("@OBSERVACION", SqlDbType.NVarChar))
                        consulta.Parameters("@OBSERVACION").Value = dtDiagImpresion.Rows(I).Item("Observacion").ToString
                        consulta.ExecuteNonQuery()
                    Next
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                    End Try
                End Using
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub guardarDetalleTerapia(ByRef objTerapiaFyR As TerapiaFisicaYRespiratoria)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_FISICA_Y_RESPIRATORIA_GUARDAR
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objTerapiaFyR.codigoOrden
                comando.Parameters.Add(New SqlParameter("@CODIGOPROCEDIMIENTO", SqlDbType.Int)).Value = objTerapiaFyR.codigoProcedimiento
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objTerapiaFyR.usuario
                comando.Parameters.Add(New SqlParameter("@DETALLE_TERAPIA_FYR", SqlDbType.Structured)).Value = objTerapiaFyR.dtTerapiaFyR
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objTerapiaFyR.codigoEP
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarDetalleTerapiaR(ByRef objTerapiaFyR As TerapiaFisicaYRespiratoriaR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_FISICA_Y_RESPIRATORIA_GUARDAR_R
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objTerapiaFyR.codigoOrden
                comando.Parameters.Add(New SqlParameter("@CODIGOPROCEDIMIENTO", SqlDbType.Int)).Value = objTerapiaFyR.codigoProcedimiento
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objTerapiaFyR.usuario
                comando.Parameters.Add(New SqlParameter("@DETALLE_TERAPIA_FYR", SqlDbType.Structured)).Value = objTerapiaFyR.dtTerapiaFyR
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objTerapiaFyR.codigoEP
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarDetalleTerapiaRR(ByRef objTerapiaFyR As TerapiaFisicaYRespiratoriaRR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_FISICA_Y_RESPIRATORIA_GUARDAR_RR
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objTerapiaFyR.codigoOrden
                comando.Parameters.Add(New SqlParameter("@CODIGOPROCEDIMIENTO", SqlDbType.Int)).Value = objTerapiaFyR.codigoProcedimiento
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objTerapiaFyR.usuario
                comando.Parameters.Add(New SqlParameter("@DETALLE_TERAPIA_FYR", SqlDbType.Structured)).Value = objTerapiaFyR.dtTerapiaFyR
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub guardarDetalleInsumosFisio(ByRef objInsumosFisio As InsumoFisioterapia)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_INSUMO_CREAR
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objInsumosFisio.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objInsumosFisio.usuarioReal
                comando.Parameters.Add(New SqlParameter("@fechaOrden", SqlDbType.DateTime)).Value = objInsumosFisio.fechaOrden
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objInsumosFisio.nRegistro
                comando.Parameters.Add(New SqlParameter("@DETALLE_INSUMO_FISIO", SqlDbType.Structured)).Value = objInsumosFisio.dtInsumosFisio
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objInsumosFisio.codigoEP
                objInsumosFisio.codigoOrden = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub guardarDetalleInsumosFisioR(ByRef objInsumosFisio As InsumoFisioterapiaR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_INSUMO_CREAR_R
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objInsumosFisio.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objInsumosFisio.usuarioReal
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objInsumosFisio.codigoOrden
                comando.Parameters.Add(New SqlParameter("@fechaOrden", SqlDbType.DateTime)).Value = objInsumosFisio.fechaOrden
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objInsumosFisio.nRegistro
                comando.Parameters.Add(New SqlParameter("@DETALLE_INSUMO_FISIO", SqlDbType.Structured)).Value = objInsumosFisio.dtInsumosFisio
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objInsumosFisio.codigoEP
                objInsumosFisio.codigoOrden = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarDetalleInsumosFisioRR(ByRef objInsumosFisio As InsumoFisioterapiaRR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_INSUMO_CREAR_RR
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objInsumosFisio.usuario
                comando.Parameters.Add(New SqlParameter("@fechaOrden", SqlDbType.DateTime)).Value = objInsumosFisio.fechaOrden
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objInsumosFisio.usuarioReal
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objInsumosFisio.codigoOrden
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objInsumosFisio.nRegistro
                comando.Parameters.Add(New SqlParameter("@DETALLE_INSUMO_FISIO", SqlDbType.Structured)).Value = objInsumosFisio.dtInsumosFisio
                objInsumosFisio.codigoOrden = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarDetalleInsumosFisio(ByRef objInsumosFisio As InsumoFisioterapia)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_INSUMO_ACTUALIZAR
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objInsumosFisio.nRegistro
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objInsumosFisio.codigoOrden
                comando.Parameters.Add(New SqlParameter("@fechaOrden", SqlDbType.DateTime)).Value = objInsumosFisio.fechaOrden
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objInsumosFisio.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objInsumosFisio.usuarioReal
                comando.Parameters.Add(New SqlParameter("@DETALLE_INSUMO_FISIO", SqlDbType.Structured)).Value = objInsumosFisio.dtInsumosFisio
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objInsumosFisio.codigoEP
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarDetalleInsumosFisioR(ByRef objInsumosFisio As InsumoFisioterapiaR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_INSUMO_ACTUALIZAR_R
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objInsumosFisio.nRegistro
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = objInsumosFisio.editado
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objInsumosFisio.codigoOrden
                comando.Parameters.Add(New SqlParameter("@fechaOrden", SqlDbType.DateTime)).Value = objInsumosFisio.fechaOrden
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objInsumosFisio.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objInsumosFisio.usuarioReal
                comando.Parameters.Add(New SqlParameter("@DETALLE_INSUMO_FISIO", SqlDbType.Structured)).Value = objInsumosFisio.dtInsumosFisio
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objInsumosFisio.codigoEP
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarDetalleInsumosFisioRR(ByRef objInsumosFisio As InsumoFisioterapiaRR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_INSUMO_ACTUALIZAR_RR
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objInsumosFisio.nRegistro
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = objInsumosFisio.editado
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objInsumosFisio.codigoOrden
                comando.Parameters.Add(New SqlParameter("@fechaOrden", SqlDbType.DateTime)).Value = objInsumosFisio.fechaOrden
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objInsumosFisio.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objInsumosFisio.usuarioReal
                comando.Parameters.Add(New SqlParameter("@DETALLE_INSUMO_FISIO", SqlDbType.Structured)).Value = objInsumosFisio.dtInsumosFisio
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarNotaFisio(ByRef objNotaFisio As NotaFisioterapia)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_NOTA_CREAR
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objNotaFisio.usuario
                comando.Parameters.Add(New SqlParameter("@fechaNota", SqlDbType.DateTime)).Value = objNotaFisio.fechaNota
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objNotaFisio.usuarioReal
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objNotaFisio.nRegistro
                comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar)).Value = objNotaFisio.nota
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objNotaFisio.codigoEP
                objNotaFisio.codigoNota = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarNotaFisioR(ByRef objNotaFisio As NotaFisioterapiaR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_NOTA_CREAR_R
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objNotaFisio.usuario
                comando.Parameters.Add(New SqlParameter("@fechaNota", SqlDbType.DateTime)).Value = objNotaFisio.fechaNota
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objNotaFisio.usuarioReal
                comando.Parameters.Add(New SqlParameter("@CODIGONOTA", SqlDbType.Int)).Value = objNotaFisio.codigoNota
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objNotaFisio.nRegistro
                comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar)).Value = objNotaFisio.nota
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objNotaFisio.codigoEP
                objNotaFisio.codigoNota = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarNotaFisioRR(ByRef objNotaFisio As NotaFisioterapiaRR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_NOTA_CREAR_RR
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objNotaFisio.usuario
                comando.Parameters.Add(New SqlParameter("@fechaNota", SqlDbType.DateTime)).Value = objNotaFisio.fechaNota
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objNotaFisio.usuarioReal
                comando.Parameters.Add(New SqlParameter("@CODIGONOTA", SqlDbType.Int)).Value = objNotaFisio.codigoNota
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objNotaFisio.nRegistro
                comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar)).Value = objNotaFisio.nota
                objNotaFisio.codigoNota = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarNotaFisio(ByRef objNotaFisio As NotaFisioterapia)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_NOTA_ACTUALIZAR
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objNotaFisio.nRegistro
                comando.Parameters.Add(New SqlParameter("@CODIGOnota", SqlDbType.Int)).Value = objNotaFisio.codigoNota
                comando.Parameters.Add(New SqlParameter("@fechaNota", SqlDbType.DateTime)).Value = objNotaFisio.fechaNota
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objNotaFisio.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objNotaFisio.usuarioReal
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objNotaFisio.codigoEP
                comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar)).Value = objNotaFisio.nota
                comando.ExecuteNonQuery()

            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarNotaFisioR(ByRef objNotaFisio As NotaFisioterapiaR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_NOTA_ACTUALIZAR_R
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objNotaFisio.nRegistro
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = objNotaFisio.editado
                comando.Parameters.Add(New SqlParameter("@CODIGOnota", SqlDbType.Int)).Value = objNotaFisio.codigoNota
                comando.Parameters.Add(New SqlParameter("@fechaNota", SqlDbType.DateTime)).Value = objNotaFisio.fechaNota
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objNotaFisio.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objNotaFisio.usuarioReal
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objNotaFisio.codigoEP
                comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar)).Value = objNotaFisio.nota
                comando.ExecuteNonQuery()

            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarNotaFisioRR(ByRef objNotaFisio As NotaFisioterapiaRR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.FISIOTERAPIA_NOTA_ACTUALIZAR_RR
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objNotaFisio.nRegistro
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = objNotaFisio.editado
                comando.Parameters.Add(New SqlParameter("@CODIGOnota", SqlDbType.Int)).Value = objNotaFisio.codigoNota
                comando.Parameters.Add(New SqlParameter("@fechaNota", SqlDbType.DateTime)).Value = objNotaFisio.fechaNota
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objNotaFisio.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objNotaFisio.usuarioReal
                comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar)).Value = objNotaFisio.nota
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''pesataña de enfermeria

    Public Shared Sub guardarDetalleInsumosEnfer(ByRef objInsumosEnfer As InsumoEnfermeria)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_INSUMO_CREAR
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objInsumosEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objInsumosEnfer.usuarioReal
                comando.Parameters.Add(New SqlParameter("@fechaOrden", SqlDbType.DateTime)).Value = objInsumosEnfer.fechaOrden
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objInsumosEnfer.nRegistro
                comando.Parameters.Add(New SqlParameter("@DETALLE_INSUMO_Enfer", SqlDbType.Structured)).Value = objInsumosEnfer.dtInsumosEnfer
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objInsumosEnfer.codigoEP
                objInsumosEnfer.codigoOrden = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub guardarDetalleInsumosEnferR(ByRef objInsumosEnfer As InsumoEnfermeriaR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_INSUMO_CREAR_R
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objInsumosEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objInsumosEnfer.usuarioReal
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objInsumosEnfer.codigoOrden
                comando.Parameters.Add(New SqlParameter("@fechaOrden", SqlDbType.DateTime)).Value = objInsumosEnfer.fechaOrden
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objInsumosEnfer.nRegistro
                comando.Parameters.Add(New SqlParameter("@DETALLE_INSUMO_Enfer", SqlDbType.Structured)).Value = objInsumosEnfer.dtInsumosEnfer
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objInsumosEnfer.codigoEP
                objInsumosEnfer.codigoOrden = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarDetalleInsumosEnferRR(ByRef objInsumosEnfer As InsumoEnfermeriaRR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_INSUMO_CREAR_RR
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objInsumosEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@fechaOrden", SqlDbType.DateTime)).Value = objInsumosEnfer.fechaOrden
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objInsumosEnfer.usuarioReal
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objInsumosEnfer.codigoOrden
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objInsumosEnfer.nRegistro
                comando.Parameters.Add(New SqlParameter("@DETALLE_INSUMO_Enfer", SqlDbType.Structured)).Value = objInsumosEnfer.dtInsumosEnfer
                objInsumosEnfer.codigoOrden = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarDetalleInsumosEnfer(ByRef objInsumosEnfer As InsumoEnfermeria)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_INSUMO_ACTUALIZAR
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objInsumosEnfer.nRegistro
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objInsumosEnfer.codigoOrden
                comando.Parameters.Add(New SqlParameter("@fechaOrden", SqlDbType.DateTime)).Value = objInsumosEnfer.fechaOrden
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objInsumosEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objInsumosEnfer.usuarioReal
                comando.Parameters.Add(New SqlParameter("@DETALLE_INSUMO_Enfer", SqlDbType.Structured)).Value = objInsumosEnfer.dtInsumosEnfer
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objInsumosEnfer.codigoEP
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarDetalleInsumosEnferR(ByRef objInsumosEnfer As InsumoEnfermeriaR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_INSUMO_ACTUALIZAR_R
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objInsumosEnfer.nRegistro
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = objInsumosEnfer.editado
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objInsumosEnfer.codigoOrden
                comando.Parameters.Add(New SqlParameter("@fechaOrden", SqlDbType.DateTime)).Value = objInsumosEnfer.fechaOrden
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objInsumosEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objInsumosEnfer.usuarioReal
                comando.Parameters.Add(New SqlParameter("@DETALLE_INSUMO_Enfer", SqlDbType.Structured)).Value = objInsumosEnfer.dtInsumosEnfer
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objInsumosEnfer.codigoEP
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarDetalleInsumosEnferRR(ByRef objInsumosEnfer As InsumoEnfermeriaRR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_INSUMO_ACTUALIZAR_RR
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objInsumosEnfer.nRegistro
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = objInsumosEnfer.editado
                comando.Parameters.Add(New SqlParameter("@CODIGOORDEN", SqlDbType.Int)).Value = objInsumosEnfer.codigoOrden
                comando.Parameters.Add(New SqlParameter("@fechaOrden", SqlDbType.DateTime)).Value = objInsumosEnfer.fechaOrden
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objInsumosEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objInsumosEnfer.usuarioReal
                comando.Parameters.Add(New SqlParameter("@DETALLE_INSUMO_Enfer", SqlDbType.Structured)).Value = objInsumosEnfer.dtInsumosEnfer
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarNotaEnfer(ByRef objNotaEnfer As NotaEnfermeria)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_NOTA_CREAR
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objNotaEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@fechaNota", SqlDbType.DateTime)).Value = objNotaEnfer.fechaNota
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objNotaEnfer.usuarioReal
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objNotaEnfer.nRegistro
                comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar)).Value = objNotaEnfer.nota
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objNotaEnfer.codigoEP
                objNotaEnfer.codigoNota = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarNotaEnferR(ByRef objNotaEnfer As NotaEnfermeriaR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_NOTA_CREAR_R
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objNotaEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@fechaNota", SqlDbType.DateTime)).Value = objNotaEnfer.fechaNota
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objNotaEnfer.usuarioReal
                comando.Parameters.Add(New SqlParameter("@CODIGONOTA", SqlDbType.Int)).Value = objNotaEnfer.codigoNota
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objNotaEnfer.nRegistro
                comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar)).Value = objNotaEnfer.nota
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objNotaEnfer.codigoEP
                objNotaEnfer.codigoNota = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarNotaEnferRR(ByRef objNotaEnfer As NotaEnfermeriaRR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_NOTA_CREAR_RR
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objNotaEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@fechaNota", SqlDbType.DateTime)).Value = objNotaEnfer.fechaNota
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objNotaEnfer.usuarioReal
                comando.Parameters.Add(New SqlParameter("@CODIGONOTA", SqlDbType.Int)).Value = objNotaEnfer.codigoNota
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objNotaEnfer.nRegistro
                comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar)).Value = objNotaEnfer.nota
                objNotaEnfer.codigoNota = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarNotaEnfer(ByRef objNotaEnfer As NotaEnfermeria)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_NOTA_ACTUALIZAR
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objNotaEnfer.nRegistro
                comando.Parameters.Add(New SqlParameter("@CODIGOnota", SqlDbType.Int)).Value = objNotaEnfer.codigoNota
                comando.Parameters.Add(New SqlParameter("@fechaNota", SqlDbType.DateTime)).Value = objNotaEnfer.fechaNota
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objNotaEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objNotaEnfer.usuarioReal
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objNotaEnfer.codigoEP
                comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar)).Value = objNotaEnfer.nota
                comando.ExecuteNonQuery()

            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarNotaEnferR(ByRef objNotaEnfer As NotaEnfermeriaR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_NOTA_ACTUALIZAR_R
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objNotaEnfer.nRegistro
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = objNotaEnfer.editado
                comando.Parameters.Add(New SqlParameter("@CODIGOnota", SqlDbType.Int)).Value = objNotaEnfer.codigoNota
                comando.Parameters.Add(New SqlParameter("@fechaNota", SqlDbType.DateTime)).Value = objNotaEnfer.fechaNota
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objNotaEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objNotaEnfer.usuarioReal
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objNotaEnfer.codigoEP
                comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar)).Value = objNotaEnfer.nota
                comando.ExecuteNonQuery()

            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub actualizarNotaEnferRR(ByRef objNotaEnfer As NotaEnfermeriaRR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_NOTA_ACTUALIZAR_RR
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objNotaEnfer.nRegistro
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = objNotaEnfer.editado
                comando.Parameters.Add(New SqlParameter("@CODIGOnota", SqlDbType.Int)).Value = objNotaEnfer.codigoNota
                comando.Parameters.Add(New SqlParameter("@fechaNota", SqlDbType.DateTime)).Value = objNotaEnfer.fechaNota
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objNotaEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@USUARIOREAL", SqlDbType.Int)).Value = objNotaEnfer.usuarioReal
                comando.Parameters.Add(New SqlParameter("@nota", SqlDbType.NVarChar)).Value = objNotaEnfer.nota
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#Region "Glucometria"
    Public Shared Sub guardarGlucometriaHC(ByRef objGlucometria As Enfermeria)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_GLUCOMETRIA_CREAR
                comando.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objGlucometria.codigoOrden
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objGlucometria.objGlucomEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objGlucometria.objGlucomEnfer.codigoEP
                comando.Parameters.Add(New SqlParameter("@TABLAGLUCOM", SqlDbType.Structured)).Value = objGlucometria.objGlucomEnfer.dtGlucomEnfer
                comando.ExecuteNonQuery()

            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarGlucometriaAM(ByRef objGlucometria As Enfermeria)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_GLUCOMETRIA_CREAR_R
                comando.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objGlucometria.codigoOrden
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objGlucometria.objGlucomEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objGlucometria.objGlucomEnfer.codigoEP
                comando.Parameters.Add(New SqlParameter("@EDITAR", SqlDbType.Bit)).Value = objGlucometria.objGlucomEnfer.editado
                comando.Parameters.Add(New SqlParameter("@TABLAGLUCOM", SqlDbType.Structured)).Value = objGlucometria.objGlucomEnfer.dtGlucomEnfer
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarGlucometriaAF(ByRef objGlucometria As Enfermeria)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.ENFERMERIA_GLUCOMETRIA_CREAR_RR
                comando.Parameters.Add(New SqlParameter("@CODIGO_ORDEN", SqlDbType.Int)).Value = objGlucometria.codigoOrden
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int)).Value = objGlucometria.objGlucomEnfer.usuario
                comando.Parameters.Add(New SqlParameter("@EDITAR", SqlDbType.Bit)).Value = objGlucometria.objGlucomEnfer.editado
                comando.Parameters.Add(New SqlParameter("@TABLAGLUCOM", SqlDbType.Structured)).Value = objGlucometria.objGlucomEnfer.dtGlucomEnfer
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Oxigeno"
    Public Shared Sub guardarFisioterapiaOxigeno(ByRef objOxigeno As Oxigeno)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion2doPlano
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.OXIGENO_FISIOTERAPIA_GUARDAR
                comando.Parameters.Add(New SqlParameter("@pCodigoOrden", SqlDbType.Int)).Value = objOxigeno.codigoOrden
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOxigeno.codigoEP
                comando.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = objOxigeno.usuario
                comando.Parameters.Add(New SqlParameter("@pTabla", SqlDbType.Structured)).Value = objOxigeno.tablaOxigeno
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = ConstantesHC.VALOR_EDITADO
                comando.ExecuteNonQuery()
            End Using
            FormPrincipal.cnxion2doPlano.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarFisioterapiaOxigenoR(ByRef objOxigeno As OxigenoR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.OXIGENO_FISIOTERAPIA_GUARDAR_R
                comando.Parameters.Add(New SqlParameter("@pCodigoOrden", SqlDbType.Int)).Value = objOxigeno.codigoOrden
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOxigeno.codigoEP
                comando.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = objOxigeno.usuario
                comando.Parameters.Add(New SqlParameter("@pTabla", SqlDbType.Structured)).Value = objOxigeno.tablaOxigeno
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = ConstantesHC.VALOR_EDITADO
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarFisioterapiaOxigenoRR(ByRef objOxigeno As OxigenoRR)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.OXIGENO_FISIOTERAPIA_GUARDAR_RR
                comando.Parameters.Add(New SqlParameter("@pCodigoOrden", SqlDbType.Int)).Value = objOxigeno.codigoOrden
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objOxigeno.codigoEP
                comando.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = objOxigeno.usuario
                comando.Parameters.Add(New SqlParameter("@pTabla", SqlDbType.Structured)).Value = objOxigeno.tablaOxigeno
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Bit)).Value = ConstantesHC.VALOR_EDITADO
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Nebulizacion"
    Public Shared Sub guardarFisioterapiaNebulizacion(ByRef objNebulizacion As Nebulizacion)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.NEBULIZACION_FISIOTERAPIA_GUARDAR
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objNebulizacion.codigoEP
                comando.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = objNebulizacion.usuario
                comando.Parameters.Add(New SqlParameter("@pTabla", SqlDbType.Structured)).Value = objNebulizacion.tablaNebulacion
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Int)).Value = Constantes.EDITADO
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarFisioterapiaNebulizacionR(ByRef objNebulizacion As Nebulizacion)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.NEBULIZACION_FISIOTERAPIA_GUARDAR_R
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objNebulizacion.codigoEP
                comando.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = objNebulizacion.usuario
                comando.Parameters.Add(New SqlParameter("@pTabla", SqlDbType.Structured)).Value = objNebulizacion.tablaNebulacion
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Int)).Value = Constantes.EDITADO
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub guardarFisioterapiaNebulizacionRR(ByRef objNebulizacion As Nebulizacion)
        Try
            Using comando = New SqlCommand
                comando.Connection = FormPrincipal.cnxion
                comando.CommandType = CommandType.StoredProcedure
                comando.CommandText = ConsultasHC.NEBULIZACION_FISIOTERAPIA_GUARDAR_RR
                comando.Parameters.Add(New SqlParameter("@CODIGO_EP", SqlDbType.Int)).Value = objNebulizacion.codigoEP
                comando.Parameters.Add(New SqlParameter("@pUsuario", SqlDbType.Int)).Value = objNebulizacion.usuario
                comando.Parameters.Add(New SqlParameter("@pTabla", SqlDbType.Structured)).Value = objNebulizacion.tablaNebulacion
                comando.Parameters.Add(New SqlParameter("@EDITADO", SqlDbType.Int)).Value = Constantes.EDITADO
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
End Class
