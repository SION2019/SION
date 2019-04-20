Imports System.Data.SqlClient
Public Class EmpresaModuloDAL

    Public Function cargar(Codigo_punto As String, empresa As String) As DataTable
        Dim Dt As New DataTable
        Dt.Clear()
        Try
            Using rsltdo = Dt.CreateDataReader()
                Using adpter = New SqlDataAdapter("EXEC [SP_EMPRESA_MODULO_CARGAR] @Codigo_punto=" & Codigo_punto & ",@id_empresa=" & empresa & "", FormPrincipal.cnxion)
                    adpter.Fill(Dt)
                End Using
            End Using

        Catch ex As Exception
            general.manejoErrores(ex) 

        End Try
        Return Dt
    End Function

    Public Function guardar(Codigo_punto As String, empresa As String) As Boolean
        Try
            Using comando = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    comando.Connection = FormPrincipal.cnxion
                    comando.Transaction = trnsccion
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.Clear()
                    comando.CommandText = "SP_EMPRESA_MODULO_ELIMINAR"
                    comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                    comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                    comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                    comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                    comando.ExecuteNonQuery()
                    If FormModulosEmpresa.R1.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_ADMON
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r2.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_AMBU
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r3.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_ASIS
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r4.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_CONTA
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r5.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_FARMA
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r6.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_HEMO
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r7.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_IMAG
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r8.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_INVEN
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r9.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_LAB
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r10.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_NOM
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r11.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_HISTC
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r12.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_AUDM
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r13.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_AUDF
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r14.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_CEXT
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.R17.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_QUIR
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r15.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_HOSP
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.r16.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_URG
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.R19.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.CODIGO_MENU_SOCG
                        comando.ExecuteNonQuery()
                    End If
                    If FormModulosEmpresa.rEspecial.Checked = True Then
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_EMPRESA_MODULO"
                        comando.Parameters.Add(New SqlParameter("@Codigo_Punto", SqlDbType.Int))
                        comando.Parameters("@Codigo_Punto").Value = Codigo_punto
                        comando.Parameters.Add(New SqlParameter("@Id_empresa", SqlDbType.Int))
                        comando.Parameters("@Id_empresa").Value = SesionActual.idEmpresa
                        comando.Parameters.Add(New SqlParameter("@CodigoModulo", SqlDbType.NVarChar))
                        comando.Parameters("@CodigoModulo").Value = Constantes.OPCION_ESPECIAL
                        comando.ExecuteNonQuery()
                    End If
                    Try
                        trnsccion.Commit()
                        Return True
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            Throw ex
                        Catch ex1 As Exception
                            Throw ex
                        End Try
                        Return False
                    End Try
                End Using
            End Using
        Catch ex As Exception
            Throw ex
            Return False
        End Try
    End Function
End Class
