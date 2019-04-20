Imports System.Data.SqlClient
Public Class ListaPrecioMedicamentoDAL
    Public Function CARGar_dgv(LISTA As String, todo As Integer) As DataTable
        Dim dt As New DataTable
        Dim cdna As String
        Dim adpter As SqlDataAdapter
        Dim enlce_dta As BindingSource = New BindingSource
        cdna = "EXEC [SP_LISTA_PRECIO_MEDICAMENTO_CARGAR] '" & LISTA & "','" & todo & "'"
        Try
            dt.Clear()
            adpter = New SqlDataAdapter(cdna, FormPrincipal.cnxion)
            adpter.Fill(dt) : adpter.Dispose()
        Catch ex As Exception : MsgBox("Error " + ex.Message)
        Finally

        End Try
        Return dt
    End Function

    Public Function guardar(ByVal codigo_L As String, ByVal descripcion As String, tabla As DataTable, ventaDirecta As Boolean) As String
        Dim codigo As String = ""
        Try
            If codigo_L = "" Then
                Using comando As New SqlCommand("SP_LISTA_PRECIO_MEDICAMENTO_CREAR")
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Connection = FormPrincipal.cnxion
                    comando.Parameters.Add(New SqlParameter("@NOMBRE", SqlDbType.NVarChar)).Value = descripcion
                    comando.Parameters.Add(New SqlParameter("@CODIGO_PUNTO", SqlDbType.NVarChar)).Value = SesionActual.codigoEP
                    comando.Parameters.Add(New SqlParameter("@Usuario_creacion", SqlDbType.NVarChar)).Value = SesionActual.idUsuario
                    comando.Parameters.Add(New SqlParameter("@pVentaDirecta", SqlDbType.Bit)).Value = ventaDirecta
                    codigo = CType(comando.ExecuteScalar, String)
                    For i As Int32 = 0 To tabla.Rows.Count - 1
                        comando.Parameters.Clear()
                        comando.CommandText = "SP_LISTA_PRECIO_MEDICAMENTO_D_GUARDAR"
                        comando.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int)).Value = CInt(codigo)
                        comando.Parameters.Add(New SqlParameter("@CODIGOint", SqlDbType.Int)).Value = CInt(tabla.Rows(i).Item(0).ToString)
                        comando.Parameters.Add(New SqlParameter("@categoria", SqlDbType.NVarChar)).Value = tabla.Rows(i).Item(2).ToString
                        comando.Parameters.Add(New SqlParameter("@PRECIO", SqlDbType.Float)).Value = CDbl(IIf(String.IsNullOrEmpty(tabla.Rows(i).Item(3).ToString), 0, tabla.Rows(i).Item(3).ToString))
                        comando.ExecuteNonQuery()
                    Next
                End Using
            Else
                Using Consulta As New SqlCommand("SP_LISTA_PRECIO_MEDICAMENTO_ACTUALIZAR")
                    Consulta.CommandType = CommandType.StoredProcedure
                    Consulta.Connection = FormPrincipal.cnxion
                    Consulta.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar)).Value = codigo_L
                    Consulta.Parameters.Add(New SqlParameter("@NOMBRE", SqlDbType.NVarChar)).Value = descripcion
                    Consulta.Parameters.Add(New SqlParameter("@USUARIO_ACTUALIZACION", SqlDbType.NVarChar)).Value = SesionActual.idUsuario
                    Consulta.Parameters.Add(New SqlParameter("@pVentaDirecta", SqlDbType.Bit)).Value = ventaDirecta
                    Consulta.ExecuteNonQuery()
                    codigo = codigo_L
                    For i As Int32 = 0 To tabla.Rows.Count - 1
                        Consulta.Parameters.Clear()
                        Consulta.CommandText = "SP_LISTA_PRECIO_MEDICAMENTO_D_GUARDAR"
                        Consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int)).Value = CInt(codigo)
                        Consulta.Parameters.Add(New SqlParameter("@CODIGOint", SqlDbType.Int)).Value = CInt(tabla.Rows(i).Item(0).ToString)
                        Consulta.Parameters.Add(New SqlParameter("@categoria", SqlDbType.NVarChar)).Value = tabla.Rows(i).Item(2).ToString
                        Consulta.Parameters.Add(New SqlParameter("@PRECIO", SqlDbType.Float)).Value = CDbl(IIf(String.IsNullOrEmpty(tabla.Rows(i).Item(3).ToString), 0, tabla.Rows(i).Item(3).ToString))
                        Consulta.ExecuteNonQuery()
                    Next
                End Using
            End If
        Catch ex As Exception
            General.manejoErrores(ex)
            codigo = ""
        End Try

        Return codigo
    End Function
    Public Function anular_LISTA(ByVal codigo As String) As Boolean
        Try
            Using Consulta As New SqlCommand("SP_LISTA_PRECIO_MEDICAMENTO_ANULAR")
                Consulta.CommandType = CommandType.StoredProcedure
                Consulta.Connection = FormPrincipal.cnxion
                Consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                Consulta.Parameters("@CODIGO").Value = CInt(codigo)
                Consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                Consulta.Parameters("@USUARIO").Value = SesionActual.idUsuario
                Consulta.ExecuteNonQuery()
            End Using
            Return True
        Catch ex As Exception
            general.manejoErrores(ex) 
            Return False
        End Try
    End Function

End Class
