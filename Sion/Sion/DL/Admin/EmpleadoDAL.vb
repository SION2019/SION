Public Class EmpleadoDAL

    Public Shared Sub guardarEmpleado(empleados As EmpleadoCollection,
                                      cuenta As CuentaBancariaCollection,
                                      salario As SalarioCollection,
                                      turno As TurnoCollection,
                                      foto As FotoCollection,
                                      firma As FirmaCollection,
                                      viatico As ViaticoCollection,
                                      auxilioTrans As AuxilioTransporteCollection,
                                      usuario As TerceroUsuarioCollection,
                                      terceroCuenta As TerceroDetalleCollection,
                                      puntosServicio As EmpresaPuntoCollection,
                                      areasServicio As AreaServicioCollection,
                                      adminEmpleado As AdminEmpleado)

        Using dbCommand As New SqlCommand("SP_EMPLEADO_ACTUALIZAR")
            dbCommand.Connection = FormPrincipal.cnxion
            dbCommand.CommandType = CommandType.StoredProcedure

            dbCommand.Parameters.Add("@tblEmpleado", SqlDbType.Structured).Value = empleados
            dbCommand.Parameters.Add("@tblCuentaBancaria", SqlDbType.Structured).Value = cuenta
            dbCommand.Parameters.Add("@tblSalario", SqlDbType.Structured).Value = salario
            dbCommand.Parameters.Add("@tblTurno", SqlDbType.Structured).Value = turno
            dbCommand.Parameters.Add("@tblFoto", SqlDbType.Structured).Value = foto
            dbCommand.Parameters.Add("@tblFirma", SqlDbType.Structured).Value = firma
            dbCommand.Parameters.Add("@tblViatico", SqlDbType.Structured).Value = viatico
            dbCommand.Parameters.Add("@tblAuxilio", SqlDbType.Structured).Value = auxilioTrans
            dbCommand.Parameters.Add("@tblTerceroUsuario", SqlDbType.Structured).Value = usuario
            dbCommand.Parameters.Add("@tblTerceroDetalle", SqlDbType.Structured).Value = terceroCuenta
            dbCommand.Parameters.Add("@tblEmpleadoPuntoS", SqlDbType.Structured).Value = puntosServicio
            dbCommand.Parameters.Add("@tblEmpleadoAreaS", SqlDbType.Structured).Value = areasServicio
            dbCommand.Parameters.Add("@esNuevaFirma", SqlDbType.Bit).Value = adminEmpleado.esOtraFirma
            dbCommand.Parameters.Add("@esNuevaCuenta", SqlDbType.Bit).Value = adminEmpleado.esOtraCuenta
            dbCommand.Parameters.Add("@esNuevoSalario", SqlDbType.Bit).Value = adminEmpleado.esOtroSalario
            dbCommand.Parameters.Add("@esNuevoTurno", SqlDbType.Bit).Value = adminEmpleado.esOtroTurno
            dbCommand.Parameters.Add("@esNuevaFoto", SqlDbType.Bit).Value = adminEmpleado.esOtraFoto
            dbCommand.Parameters.Add("@esNuevoViatico", SqlDbType.Bit).Value = adminEmpleado.esOtroViatico
            dbCommand.Parameters.Add("@esNuevoAuxilio", SqlDbType.Bit).Value = adminEmpleado.esOtroAuxilio
            dbCommand.Parameters.Add("@esNuevoUsuario", SqlDbType.Bit).Value = adminEmpleado.esOtroUsuario
            dbCommand.Parameters.Add("@pUsuario", SqlDbType.Int).Value = SesionActual.idUsuario
            dbCommand.ExecuteNonQuery()

        End Using
    End Sub

    Public Shared Function tomarHuella(txtcodigo As String, ByRef huella1 As DPFP.Template) As Boolean
        Try
            Using cmdDelete As New SqlCommand("DELETE FROM D_EMPLEADO_HUELLA WHERE Id_empleado=@Id_empleado ", FormPrincipal.cnxion)
                Using cmdInsert As New SqlCommand("INSERT INTO D_EMPLEADO_HUELLA (Huella,id_empleado) values (@Huella,@Id_empleado) ", FormPrincipal.cnxion)
                    Using trans = FormPrincipal.cnxion.BeginTransaction

                        With cmdDelete
                            .Transaction = trans
                            .Parameters.Add(New SqlParameter("@Id_empleado", SqlDbType.Int)).Value = CInt(txtcodigo)
                        End With
                        cmdDelete.ExecuteNonQuery()

                        With cmdInsert
                            .Transaction = trans
                            .Parameters.Add(New SqlParameter("@Id_empleado", SqlDbType.Int)).Value = CInt(txtcodigo)
                            .Parameters.Add(New SqlParameter("@Huella", SqlDbType.VarBinary, huella1.Bytes.Length)).Value = huella1.Bytes
                        End With
                        cmdInsert.ExecuteNonQuery()

                        trans.Commit()

                    End Using
                End Using
            End Using

            Return True

        Catch ex As Exception
            Throw ex
        End Try

        Return False
    End Function

    Public Shared Function guardarDocumento(txtcodigo As String, descripcion1 As Object, ByRef documento1 As Byte()) As Boolean
        Try
            Using cmdDelete As New SqlCommand
                Using cmdInsert As New SqlCommand
                    Using trans = FormPrincipal.cnxion.BeginTransaction

                        cmdDelete.Connection = FormPrincipal.cnxion
                        cmdDelete.Transaction = trans
                        cmdDelete.CommandText = "Delete From D_EMPLEADO_DOCUMENTO Where Codigo_Doc = '" & descripcion1 & "' and Id_empleado='" & txtcodigo & "'"
                        cmdDelete.ExecuteNonQuery()

                        cmdInsert.Connection = FormPrincipal.cnxion
                        cmdInsert.Transaction = trans
                        cmdInsert.CommandText = "Insert Into D_EMPLEADO_DOCUMENTO(Id_empleado,Codigo_Doc,Imagen)Values(@Id_empleado,@Descripcion,@Imagen)"

                        cmdInsert.Parameters.Add(New SqlParameter("@Id_empleado", SqlDbType.Int)).Value = CInt(txtcodigo)
                        cmdInsert.Parameters.Add(New SqlParameter("@Descripcion", SqlDbType.Int)).Value = descripcion1
                        cmdInsert.Parameters.Add(New SqlParameter("@Imagen", SqlDbType.Image)).Value = documento1
                        cmdInsert.ExecuteNonQuery()

                        trans.Commit()

                        Return True
                    End Using
                End Using
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Function



End Class
