Imports System.Data.SqlClient
Public Class AdmisionDAL

    Public Function guardarAdmision(ByVal objAdmision As Admision)
        Dim codigo As String = String.Empty

        Try

            Using dbCommand As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    dbCommand.Connection = FormPrincipal.cnxion
                    dbCommand.Transaction = trnsccion

                    dbCommand.CommandType = CommandType.StoredProcedure
                    dbCommand.CommandText = "SP_ADMISIONES_CREAR"
                    dbCommand.Parameters.Add(New SqlParameter("@Identi_Paciente", SqlDbType.Int))
                    dbCommand.Parameters("@Identi_Paciente").Value = objAdmision.identiPaciente
                    dbCommand.Parameters.Add(New SqlParameter("@solicitud", SqlDbType.Int))
                    dbCommand.Parameters("@solicitud").Value = objAdmision.solicitud
                    dbCommand.Parameters.Add(New SqlParameter("@Fecha_Admision", SqlDbType.DateTime))
                    dbCommand.Parameters("@Fecha_Admision").Value = objAdmision.fechaAdmision
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_EP", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_EP").Value = objAdmision.CodigoEp
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_Especialidad", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_Especialidad").Value = objAdmision.codigoEspecialidad
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_Triage", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_Triage").Value = objAdmision.codigoTriage
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_Contrato", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_Contrato").Value = objAdmision.codigoContrato
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_Area_Servicio", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_Area_Servicio").Value = objAdmision.codigoArea
                    dbCommand.Parameters.Add(New SqlParameter("@Id_contacto", SqlDbType.Int))
                    dbCommand.Parameters("@Id_contacto").Value = objAdmision.idContacto
                    dbCommand.Parameters.Add(New SqlParameter("@USUARIO_Creacion", SqlDbType.NVarChar))
                    dbCommand.Parameters("@USUARIO_Creacion").Value = objAdmision.usuario
                    dbCommand.Parameters.Add(New SqlParameter("@observacion", SqlDbType.NVarChar))
                    dbCommand.Parameters("@observacion").Value = objAdmision.observacion
                    dbCommand.Parameters.Add(New SqlParameter("@valorSOAT", SqlDbType.Int))
                    dbCommand.Parameters("@valorSOAT").Value = objAdmision.valorSOAT
                    dbCommand.Parameters.Add(New SqlParameter("@Documento_Paciente", SqlDbType.NVarChar))
                    dbCommand.Parameters("@Documento_Paciente").Value = objAdmision.documentoPaciente
                    objAdmision.nRegistro = CType(dbCommand.ExecuteScalar, Integer)


                    If objAdmision.opcionAcomp = True Then
                        dbCommand.Parameters.Clear()
                        dbCommand.CommandText = "SP_ADMISION_ACOMPAÑANTE_CREAR"
                        dbCommand.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int))
                        dbCommand.Parameters("@N_Registro").Value = objAdmision.nRegistro
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Identificacion_A", SqlDbType.Int))
                        dbCommand.Parameters("@Codigo_Identificacion_A").Value = objAdmision.codigoIdAcompanate
                        dbCommand.Parameters.Add(New SqlParameter("@Identificacion_A", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Identificacion_A").Value = objAdmision.idAcompanate
                        dbCommand.Parameters.Add(New SqlParameter("@Acompañante", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Acompañante").Value = objAdmision.acompanante
                        dbCommand.Parameters.Add(New SqlParameter("@Direccion", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Direccion").Value = objAdmision.Direccion
                        dbCommand.Parameters.Add(New SqlParameter("@Telefono_A", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Telefono_A").Value = objAdmision.telAcompanate
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Pais", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Codigo_Pais").Value = objAdmision.codigoPais
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Departamento", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Codigo_Departamento").Value = objAdmision.codigoDepartamento
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Municipio", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Codigo_Municipio").Value = objAdmision.codigoMunicipio
                        dbCommand.ExecuteNonQuery()
                    End If
                    If objAdmision.opcionRespon = True Then
                        dbCommand.Parameters.Clear()
                        dbCommand.CommandText = "SP_ADMISION_RESPONSABLE_CREAR"
                        dbCommand.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int))
                        dbCommand.Parameters("@N_Registro").Value = objAdmision.nRegistro
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Identificacion_R", SqlDbType.Int))
                        dbCommand.Parameters("@Codigo_Identificacion_R").Value = objAdmision.codigoIdResponsable
                        dbCommand.Parameters.Add(New SqlParameter("@Identificacion_R", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Identificacion_R").Value = objAdmision.idResponsable
                        dbCommand.Parameters.Add(New SqlParameter("@Responsable", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Responsable").Value = objAdmision.responsable
                        dbCommand.Parameters.Add(New SqlParameter("@Direccion_R", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Direccion_R").Value = objAdmision.direccionResponsable
                        dbCommand.Parameters.Add(New SqlParameter("@Telefono_R", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Telefono_R").Value = objAdmision.telResponsable
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Pais_R", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Codigo_Pais_R").Value = objAdmision.codigoPaisResponsable
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Departamento_R", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Codigo_Departamento_R").Value = objAdmision.codigoDepartamentoResponsable
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Municipio_R", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Codigo_Municipio_R").Value = objAdmision.codigoMunicipioResponsable
                        dbCommand.ExecuteNonQuery()
                    End If

                    dbCommand.Parameters.Clear()
                    dbCommand.CommandText = "SP_INGRESO_CREAR"
                    dbCommand.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int))
                    dbCommand.Parameters("@N_Registro").Value = objAdmision.nRegistro
                    dbCommand.Parameters.Add(New SqlParameter("@N_Autorizacion", SqlDbType.NVarChar))
                    dbCommand.Parameters("@N_Autorizacion").Value = objAdmision.numAutorizacion
                    dbCommand.Parameters.Add(New SqlParameter("@Via_Ingreso", SqlDbType.Int))
                    dbCommand.Parameters("@Via_Ingreso").Value = objAdmision.viaIngreso
                    dbCommand.Parameters.Add(New SqlParameter("@Institucion", SqlDbType.Int))
                    dbCommand.Parameters("@Institucion").Value = objAdmision.institucion
                    dbCommand.Parameters.Add(New SqlParameter("@Causas_Externa", SqlDbType.Int))
                    dbCommand.Parameters("@Causas_Externa").Value = objAdmision.causaExterna
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_Estado_Atencion", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_Estado_Atencion").Value = objAdmision.codigoEstado
                    dbCommand.Parameters.Add(New SqlParameter("@USUARIO_Creacion", SqlDbType.NVarChar))
                    dbCommand.Parameters("@USUARIO_Creacion").Value = objAdmision.usuario
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_ep", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_ep").Value = objAdmision.CodigoEp
                    dbCommand.ExecuteNonQuery()

                    dbCommand.Parameters.Clear()
                    dbCommand.CommandText = "SP_REGISTRO_CAMA_CREAR"
                    dbCommand.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int))
                    dbCommand.Parameters("@N_Registro").Value = objAdmision.nRegistro
                    dbCommand.Parameters.Add(New SqlParameter("@codigo_cama", SqlDbType.Int))
                    dbCommand.Parameters("@codigo_cama").Value = objAdmision.codigoCama
                    dbCommand.Parameters.Add(New SqlParameter("@USUARIO_Creacion", SqlDbType.NVarChar))
                    dbCommand.Parameters("@USUARIO_Creacion").Value = objAdmision.usuario
                    dbCommand.ExecuteNonQuery()

                    For i = 0 To objAdmision.dtDiagnostico.Rows.Count - 1
                        If objAdmision.dtDiagnostico.Rows(i).Item("Codigo_CIE").ToString <> "" And objAdmision.dtDiagnostico.Rows(i).Item("Estado").ToString = Constantes.ITEM_NUEVO Then
                            dbCommand.Parameters.Clear()
                            dbCommand.CommandText = "SP_INGRESO_DIAG_CREAR"
                            dbCommand.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int))
                            dbCommand.Parameters("@N_Registro").Value = objAdmision.nRegistro
                            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Cie", SqlDbType.NVarChar))
                            dbCommand.Parameters("@Codigo_Cie").Value = objAdmision.dtDiagnostico.Rows(i).Item("Codigo_CIE").ToString
                            dbCommand.Parameters.Add(New SqlParameter("@Observacion", SqlDbType.NVarChar))
                            dbCommand.Parameters("@Observacion").Value = objAdmision.dtDiagnostico.Rows(i).Item("Observacion").ToString
                            dbCommand.Parameters.Add(New SqlParameter("@Codigo_ep", SqlDbType.Int))
                            dbCommand.Parameters("@Codigo_ep").Value = objAdmision.CodigoEp
                            dbCommand.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                            dbCommand.Parameters("@USUARIO").Value = objAdmision.usuario
                            dbCommand.ExecuteNonQuery()
                        End If
                    Next

                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        codigo = ""
                        trnsccion.Rollback()
                        Throw ex
                    End Try
                End Using
            End Using
        Catch ex As Exception
            codigo = ""
            Throw ex
        End Try
        Return objAdmision.nRegistro
    End Function

    Public Function guardarTraslado(ByVal objAdmision As Admision, elementoAEliminar As List(Of String)) As String

        Try
            For indice = 0 To objAdmision.dtTraslado.Rows.Count - 2
                Using dbCommand As New SqlCommand
                    dbCommand.Connection = FormPrincipal.cnxion
                    For Each sentencia As String In elementoAEliminar
                        If sentencia <> "" Then
                            dbCommand.CommandText = sentencia
                            dbCommand.ExecuteNonQuery()
                        End If
                    Next
                    If objAdmision.dtTraslado.Rows(indice).Item(0).ToString = "" Then
                        dbCommand.Parameters.Clear()
                        dbCommand.CommandType = CommandType.StoredProcedure
                        dbCommand.CommandText = "SP_SERVICIO_TRASLADO_CREAR"
                        dbCommand.Parameters.Add(New SqlParameter("@Registro", SqlDbType.Int))
                        dbCommand.Parameters("@Registro").Value = objAdmision.nRegistro
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_TRAS", SqlDbType.Int))
                        dbCommand.Parameters("@Codigo_TRAS").Value = objAdmision.dtTraslado.Rows(indice).Item(5).ToString
                        dbCommand.Parameters.Add(New SqlParameter("@OBSERVACIONES", SqlDbType.NVarChar))
                        dbCommand.Parameters("@OBSERVACIONES").Value = objAdmision.dtTraslado.Rows(indice).Item(4).ToString
                        dbCommand.ExecuteNonQuery()
                    Else
                        dbCommand.Parameters.Clear()
                        dbCommand.CommandType = CommandType.StoredProcedure
                        dbCommand.CommandText = "SP_SERVICIO_TRASLADO_ACTUALIZAR"
                        dbCommand.Parameters.Add(New SqlParameter("@OBSERVACIONES", SqlDbType.NVarChar))
                        dbCommand.Parameters("@OBSERVACIONES").Value = objAdmision.dtTraslado.Rows(indice).Item(4).ToString
                        dbCommand.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                        dbCommand.Parameters("@CODIGO").Value = objAdmision.dtTraslado.Rows(indice).Item(0).ToString
                        dbCommand.ExecuteNonQuery()
                    End If
                End Using
            Next
        Catch ex As Exception

            Throw ex
        End Try
        Return objAdmision.nRegistro
    End Function

    Public Sub actualizarAdmision(ByVal objAdmision As Admision)

        Try

            Using dbCommand As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    dbCommand.Connection = FormPrincipal.cnxion
                    dbCommand.Transaction = trnsccion

                    dbCommand.CommandType = CommandType.StoredProcedure
                    dbCommand.CommandText = "SP_ADMISIONES_ACTUALIZAR"
                    dbCommand.Parameters.Add(New SqlParameter("@Identi_Paciente", SqlDbType.Int))
                    dbCommand.Parameters("@Identi_Paciente").Value = objAdmision.identiPaciente
                    dbCommand.Parameters.Add(New SqlParameter("@Fecha_Admision", SqlDbType.DateTime))
                    dbCommand.Parameters("@Fecha_Admision").Value = objAdmision.fechaAdmision
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_EP", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_EP").Value = objAdmision.CodigoEp
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_Especialidad", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_Especialidad").Value = objAdmision.codigoEspecialidad
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_Triage", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_Triage").Value = objAdmision.codigoTriage
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_Contrato", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_Contrato").Value = objAdmision.codigoContrato
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_Area_Servicio", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_Area_Servicio").Value = objAdmision.codigoArea
                    dbCommand.Parameters.Add(New SqlParameter("@Id_contacto", SqlDbType.Int))
                    dbCommand.Parameters("@Id_contacto").Value = objAdmision.idContacto
                    dbCommand.Parameters.Add(New SqlParameter("@USUARIO_Creacion", SqlDbType.NVarChar))
                    dbCommand.Parameters("@USUARIO_Creacion").Value = objAdmision.usuario
                    dbCommand.Parameters.Add(New SqlParameter("@observacion", SqlDbType.NVarChar))
                    dbCommand.Parameters("@observacion").Value = objAdmision.observacion
                    dbCommand.Parameters.Add(New SqlParameter("@N_registro", SqlDbType.Int))
                    dbCommand.Parameters("@N_registro").Value = objAdmision.nRegistro
                    dbCommand.Parameters.Add(New SqlParameter("@valorSOAT", SqlDbType.Int))
                    dbCommand.Parameters("@valorSOAT").Value = objAdmision.valorSOAT
                    dbCommand.ExecuteNonQuery()

                    If objAdmision.opcionAcomp = True Then
                        dbCommand.Parameters.Clear()
                        dbCommand.CommandText = "SP_ADMISION_ACOMPAÑANTE_CREAR"
                        dbCommand.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int))
                        dbCommand.Parameters("@N_Registro").Value = objAdmision.nRegistro
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Identificacion_A", SqlDbType.Int))
                        dbCommand.Parameters("@Codigo_Identificacion_A").Value = objAdmision.codigoIdAcompanate
                        dbCommand.Parameters.Add(New SqlParameter("@Identificacion_A", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Identificacion_A").Value = objAdmision.idAcompanate
                        dbCommand.Parameters.Add(New SqlParameter("@Acompañante", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Acompañante").Value = objAdmision.acompanante
                        dbCommand.Parameters.Add(New SqlParameter("@Direccion", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Direccion").Value = objAdmision.Direccion
                        dbCommand.Parameters.Add(New SqlParameter("@Telefono_A", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Telefono_A").Value = objAdmision.telAcompanate
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Pais", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Codigo_Pais").Value = objAdmision.codigoPais
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Departamento", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Codigo_Departamento").Value = objAdmision.codigoDepartamento
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Municipio", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Codigo_Municipio").Value = objAdmision.codigoMunicipio
                        dbCommand.ExecuteNonQuery()
                    End If
                    If objAdmision.opcionRespon = True Then
                        dbCommand.Parameters.Clear()
                        dbCommand.CommandText = "SP_ADMISION_RESPONSABLE_CREAR"
                        dbCommand.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int))
                        dbCommand.Parameters("@N_Registro").Value = objAdmision.nRegistro
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_Identificacion_R", SqlDbType.Int))
                        dbCommand.Parameters("@Codigo_Identificacion_R").Value = objAdmision.codigoIdResponsable
                        dbCommand.Parameters.Add(New SqlParameter("@Identificacion_R", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Identificacion_R").Value = objAdmision.idResponsable
                        dbCommand.Parameters.Add(New SqlParameter("@Responsable", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Responsable").Value = objAdmision.responsable
                        dbCommand.Parameters.Add(New SqlParameter("@direccion_R", SqlDbType.NVarChar))
                        dbCommand.Parameters("@direccion_R").Value = objAdmision.direccionResponsable
                        dbCommand.Parameters.Add(New SqlParameter("@Telefono_R", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Telefono_R").Value = objAdmision.telResponsable
                        dbCommand.Parameters.Add(New SqlParameter("@codigo_pais_R", SqlDbType.NVarChar))
                        dbCommand.Parameters("@codigo_pais_R").Value = objAdmision.codigoPaisResponsable
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_departamento_R", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Codigo_departamento_R").Value = objAdmision.codigoDepartamentoResponsable
                        dbCommand.Parameters.Add(New SqlParameter("@Codigo_municipio_R", SqlDbType.NVarChar))
                        dbCommand.Parameters("@Codigo_municipio_R").Value = objAdmision.codigoMunicipioResponsable
                        dbCommand.ExecuteNonQuery()
                    End If

                    dbCommand.Parameters.Clear()
                    dbCommand.CommandText = "SP_INGRESO_ACTUALIZAR"
                    dbCommand.Parameters.Add(New SqlParameter("@N_Autorizacion", SqlDbType.NVarChar))
                    dbCommand.Parameters("@N_Autorizacion").Value = objAdmision.numAutorizacion
                    dbCommand.Parameters.Add(New SqlParameter("@Via_Ingreso", SqlDbType.Int))
                    dbCommand.Parameters("@Via_Ingreso").Value = objAdmision.viaIngreso
                    dbCommand.Parameters.Add(New SqlParameter("@Institucion", SqlDbType.Int))
                    dbCommand.Parameters("@Institucion").Value = objAdmision.institucion
                    dbCommand.Parameters.Add(New SqlParameter("@Causas_Externa", SqlDbType.Int))
                    dbCommand.Parameters("@Causas_Externa").Value = objAdmision.causaExterna
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_Estado_Atencion", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_Estado_Atencion").Value = objAdmision.codigoEstado
                    dbCommand.Parameters.Add(New SqlParameter("@USUARIO_Creacion", SqlDbType.NVarChar))
                    dbCommand.Parameters("@USUARIO_Creacion").Value = objAdmision.usuario
                    dbCommand.Parameters.Add(New SqlParameter("@Codigo_ep", SqlDbType.Int))
                    dbCommand.Parameters("@Codigo_ep").Value = objAdmision.CodigoEp
                    dbCommand.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int))
                    dbCommand.Parameters("@N_Registro").Value = objAdmision.nRegistro
                    dbCommand.ExecuteNonQuery()

                    dbCommand.Parameters.Clear()
                    dbCommand.CommandText = "SP_REGISTRO_CAMA_ACTUALIZAR"
                    dbCommand.Parameters.Add(New SqlParameter("@pRegistro", SqlDbType.Int))
                    dbCommand.Parameters("@pRegistro").Value = objAdmision.nRegistro
                    dbCommand.Parameters.Add(New SqlParameter("@pcodigo_cama", SqlDbType.Int))
                    dbCommand.Parameters("@pcodigo_cama").Value = objAdmision.codigoCama
                    dbCommand.Parameters.Add(New SqlParameter("@pUsuario_Actualizacion", SqlDbType.NVarChar))
                    dbCommand.Parameters("@pUsuario_Actualizacion").Value = objAdmision.usuario
                    dbCommand.ExecuteNonQuery()

                    For i = 0 To objAdmision.dtDiagnostico.Rows.Count - 1
                        If objAdmision.dtDiagnostico.Rows(i).Item("Codigo_CIE").ToString <> "" And objAdmision.dtDiagnostico.Rows(i).Item("Estado").ToString = Constantes.ITEM_NUEVO Then
                            dbCommand.Parameters.Clear()
                            dbCommand.CommandText = "SP_INGRESO_DIAG_CREAR"
                            dbCommand.Parameters.Add(New SqlParameter("@N_Registro", SqlDbType.Int))
                            dbCommand.Parameters("@N_Registro").Value = objAdmision.nRegistro
                            dbCommand.Parameters.Add(New SqlParameter("@Codigo_Cie", SqlDbType.NVarChar))
                            dbCommand.Parameters("@Codigo_Cie").Value = objAdmision.dtDiagnostico.Rows(i).Item("Codigo_CIE").ToString
                            dbCommand.Parameters.Add(New SqlParameter("@Observacion", SqlDbType.NVarChar))
                            dbCommand.Parameters("@Observacion").Value = objAdmision.dtDiagnostico.Rows(i).Item("Observacion").ToString
                            dbCommand.Parameters.Add(New SqlParameter("@Codigo_ep", SqlDbType.Int))
                            dbCommand.Parameters("@Codigo_ep").Value = objAdmision.CodigoEp
                            dbCommand.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                            dbCommand.Parameters("@USUARIO").Value = objAdmision.usuario
                            dbCommand.ExecuteNonQuery()
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

    Public Sub guardarDocumento(objAdmision As Admision, consulta As String)
        Try
            Using dbCommand As New SqlCommand(consulta)
                dbCommand.Parameters.Clear()
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.Parameters.Add(New SqlParameter("@N_REGISTRO", SqlDbType.Int))
                dbCommand.Parameters("@N_REGISTRO").Value = objAdmision.nRegistro
                dbCommand.Parameters.Add(New SqlParameter("@CODIGO_DOC", SqlDbType.Int))
                dbCommand.Parameters("@CODIGO_DOC").Value = objAdmision.tipoDocumento
                dbCommand.Parameters.Add(New SqlParameter("@IMAGEN", SqlDbType.VarBinary))
                dbCommand.Parameters("@IMAGEN").Value = objAdmision.imagen
                dbCommand.Parameters.Add(New SqlParameter("@EXTENSION", SqlDbType.NVarChar))
                dbCommand.Parameters("@EXTENSION").Value = objAdmision.extensionDoc
                dbCommand.Parameters.Add(New SqlParameter("@TIPO", SqlDbType.NVarChar))
                dbCommand.Parameters("@TIPO").Value = objAdmision.tipo
                dbCommand.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
