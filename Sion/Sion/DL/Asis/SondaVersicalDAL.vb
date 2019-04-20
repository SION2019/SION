Imports System.Data.SqlClient
Public Class SondaVersicalDAL

    Public Shared Sub guardarSonda(Codigo_Sonda As Object, N_Registro As String, Fecha_Colocacion As Date, Microorganismo As String, Fecha_retiro As Date, usuario As Integer, dtsonda As DataTable)

        Try

            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()

                    If Codigo_Sonda.text = "" Then
                        consulta.Parameters.Clear()
                        consulta.CommandType = CommandType.StoredProcedure
                        consulta.CommandText = "SP_SONDA_VESICAL_CREAR"
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                        consulta.Parameters("@REGISTRO").Value = N_Registro
                        consulta.Parameters.Add(New SqlParameter("@FECHACOLO", SqlDbType.Date))
                        consulta.Parameters("@FECHACOLO").Value = Fecha_Colocacion
                        consulta.Parameters.Add(New SqlParameter("@MICROORGANISMO", SqlDbType.NVarChar))
                        consulta.Parameters("@MICROORGANISMO").Value = Microorganismo
                        consulta.Parameters.Add(New SqlParameter("@FECHARETIRO", SqlDbType.Date))
                        consulta.Parameters("@FECHARETIRO").Value = Fecha_retiro
                        consulta.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int))
                        consulta.Parameters("@usuario").Value = usuario
                        Codigo_Sonda.text = CType(consulta.ExecuteScalar, String)

                        For i = 0 To dtsonda.Rows.Count - 1
                            consulta.Parameters.Clear()
                            consulta.CommandText = "SP_SONDA_VESICAL_DETALLE_CREAR"
                            consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                            consulta.Parameters("@CODIGO").Value = Codigo_Sonda.text
                            consulta.Parameters.Add(New SqlParameter("@CODIGOCRITERIO", SqlDbType.Int))
                            consulta.Parameters("@CODIGOCRITERIO").Value = dtsonda.Rows(i).Item(0).ToString
                            consulta.Parameters.Add(New SqlParameter("@RESPUESTA", SqlDbType.NVarChar))
                            consulta.Parameters("@RESPUESTA").Value = dtsonda.Rows(i).Item(2).ToString
                            consulta.ExecuteNonQuery()
                        Next
                    Else
                        consulta.Parameters.Clear()
                        consulta.CommandType = CommandType.StoredProcedure
                        consulta.CommandText = "SP_SONDA_VESICAL_ACTUALIZAR"
                        consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO").Value = Codigo_Sonda.text
                        consulta.Parameters.Add(New SqlParameter("@FECHACOLO", SqlDbType.Date))
                        consulta.Parameters("@FECHACOLO").Value = Fecha_Colocacion
                        consulta.Parameters.Add(New SqlParameter("@MICROORGANISMO", SqlDbType.NVarChar))
                        consulta.Parameters("@MICROORGANISMO").Value = Microorganismo
                        consulta.Parameters.Add(New SqlParameter("@FECHARETIRO", SqlDbType.Date))
                        consulta.Parameters("@FECHARETIRO").Value = Fecha_retiro
                        consulta.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int))
                        consulta.Parameters("@usuario").Value = usuario
                        consulta.ExecuteNonQuery()
                        For i = 0 To dtsonda.Rows.Count - 1
                            consulta.Parameters.Clear()
                            consulta.CommandText = "SP_SONDA_VESICAL_DETALLE_ACTUALIZAR"
                            consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                            consulta.Parameters("@CODIGO").Value = Codigo_Sonda.text
                            consulta.Parameters.Add(New SqlParameter("@CODIGOCRITERIO", SqlDbType.Int))
                            consulta.Parameters("@CODIGOCRITERIO").Value = dtsonda.Rows(i).Item(0).ToString
                            consulta.Parameters.Add(New SqlParameter("@RESPUESTA", SqlDbType.NVarChar))
                            consulta.Parameters("@RESPUESTA").Value = dtsonda.Rows(i).Item(2).ToString
                            consulta.ExecuteNonQuery()
                        Next
                    End If
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        trnsccion.Rollback()
                        general.manejoErrores(ex) 
                    End Try
                End Using
            End Using
        Catch ex As Exception
            general.manejoErrores(ex) 
        End Try
    End Sub


    Public Shared Sub guardarVeno(Codigo_veno As Object, N_Registro As String, Fecha_Colocacion As Date, usuario As Integer, dtsonda As DataTable)
        Try

            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Parameters.Clear()

                    If Codigo_veno.text = "" Then
                        consulta.Parameters.Clear()
                        consulta.CommandType = CommandType.StoredProcedure
                        consulta.CommandText = "SP_VENOPUNCION_CREAR"
                        consulta.Parameters.Add(New SqlParameter("@REGISTRO", SqlDbType.Int))
                        consulta.Parameters("@REGISTRO").Value = N_Registro
                        consulta.Parameters.Add(New SqlParameter("@FECHACOLO", SqlDbType.Date))
                        consulta.Parameters("@FECHACOLO").Value = Fecha_Colocacion
                        consulta.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int))
                        consulta.Parameters("@usuario").Value = usuario
                        Codigo_veno.text = CType(consulta.ExecuteScalar, String)
                        For i = 0 To dtsonda.Rows.Count - 1
                            consulta.Parameters.Clear()
                            consulta.CommandText = "SP_VENOPUNCION_DETALLE_CREAR"
                            consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                            consulta.Parameters("@CODIGO").Value = Codigo_veno.text
                            consulta.Parameters.Add(New SqlParameter("@CODIGOCRITERIO", SqlDbType.Int))
                            consulta.Parameters("@CODIGOCRITERIO").Value = dtsonda.Rows(i).Item(0).ToString
                            consulta.Parameters.Add(New SqlParameter("@RESPUESTA", SqlDbType.NVarChar))
                            consulta.Parameters("@RESPUESTA").Value = dtsonda.Rows(i).Item(2).ToString
                            consulta.ExecuteNonQuery()
                        Next
                    Else
                        consulta.Parameters.Clear()
                        consulta.CommandType = CommandType.StoredProcedure
                        consulta.CommandText = "SP_VENOPUNCION_ACTUALIZAR"
                        consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                        consulta.Parameters("@CODIGO").Value = Codigo_veno.text
                        consulta.Parameters.Add(New SqlParameter("@FECHACOLO", SqlDbType.Date))
                        consulta.Parameters("@FECHACOLO").Value = Fecha_Colocacion
                        consulta.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int))
                        consulta.Parameters("@usuario").Value = usuario
                        consulta.ExecuteNonQuery()
                        For i = 0 To dtsonda.Rows.Count - 1
                            consulta.Parameters.Clear()
                            consulta.CommandText = "SP_VENOPUNCION_DETALLE_ACTUALIZAR"
                            consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                            consulta.Parameters("@CODIGO").Value = Codigo_veno.text
                            consulta.Parameters.Add(New SqlParameter("@CODIGOCRITERIO", SqlDbType.Int))
                            consulta.Parameters("@CODIGOCRITERIO").Value = dtsonda.Rows(i).Item(0).ToString
                            consulta.Parameters.Add(New SqlParameter("@RESPUESTA", SqlDbType.NVarChar))
                            consulta.Parameters("@RESPUESTA").Value = dtsonda.Rows(i).Item(2).ToString
                            consulta.ExecuteNonQuery()
                        Next
                    End If
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        trnsccion.Rollback()
                        general.manejoErrores(ex) 
                    End Try
                End Using
            End Using
        Catch ex As Exception
            general.manejoErrores(ex) 
        End Try
    End Sub
End Class
