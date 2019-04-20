Imports System.Data.SqlClient
Public Class EpsDAL

    Public Function guardar_eps(ByVal codigo_in As String, ByVal descripcion As String, ByVal codigo_eps As String, comboFormato As Integer, estancia As Integer, dtConfiguracion As DataTable, codigoGrupo As Integer) As String
        Dim codigo As String = ""
        Dim respuesta As String = ""
        Try
            If codigo_in = "" Then

                Using Consulta As New SqlCommand("SP_EPS_CREAR")
                    Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                        Consulta.Connection = FormPrincipal.cnxion
                        Consulta.Transaction = trnsccion
                        Consulta.CommandType = CommandType.StoredProcedure
                        Consulta.Connection = FormPrincipal.cnxion
                        Consulta.Parameters.Add(New SqlParameter("@NOMBRE", SqlDbType.NVarChar))
                        Consulta.Parameters("@NOMBRE").Value = descripcion
                        Consulta.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar))
                        Consulta.Parameters("@Codigo").Value = codigo_eps
                        Consulta.Parameters.Add(New SqlParameter("@Estancia", SqlDbType.Int))
                        Consulta.Parameters("@Estancia").Value = estancia
                        Consulta.Parameters.Add(New SqlParameter("@CodigoReporte", SqlDbType.Int))
                        Consulta.Parameters("@CodigoReporte").Value = comboFormato
                        Consulta.Parameters.Add(New SqlParameter("@Usuario_creacion", SqlDbType.NVarChar))
                        Consulta.Parameters("@Usuario_creacion").Value = SesionActual.idUsuario
                        codigo = CType(Consulta.ExecuteScalar, String)
                        trnsccion.Commit()
                    End Using
                End Using
            Else
                Using Consulta As New SqlCommand("SP_EPS_ACTUALIZAR")
                    Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                        Consulta.Connection = FormPrincipal.cnxion
                        Consulta.Transaction = trnsccion
                        Consulta.CommandType = CommandType.StoredProcedure
                        Consulta.Connection = FormPrincipal.cnxion
                        Consulta.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.Int))
                        Consulta.Parameters("@Codigo").Value = codigo_in
                        Consulta.Parameters.Add(New SqlParameter("@CODIGO_EPS", SqlDbType.NVarChar))
                        Consulta.Parameters("@CODIGO_EPS").Value = codigo_eps
                        Consulta.Parameters.Add(New SqlParameter("@NOMBRE", SqlDbType.NVarChar))
                        Consulta.Parameters("@NOMBRE").Value = descripcion
                        Consulta.Parameters.Add(New SqlParameter("@Estancia", SqlDbType.Int))
                        Consulta.Parameters("@Estancia").Value = estancia
                        Consulta.Parameters.Add(New SqlParameter("@CodigoReporte", SqlDbType.Int))
                        Consulta.Parameters("@CodigoReporte").Value = comboFormato
                        Consulta.Parameters.Add(New SqlParameter("@USUARIO_ACTUALIZACION", SqlDbType.NVarChar))
                        Consulta.Parameters("@USUARIO_ACTUALIZACION").Value = SesionActual.idUsuario
                        Consulta.Parameters.Add(New SqlParameter("@TABLA", SqlDbType.Structured)).Value = dtConfiguracion
                        Consulta.Parameters.Add(New SqlParameter("@codigogrupo", SqlDbType.Int)).Value = codigoGrupo
                        Consulta.ExecuteNonQuery()
                        codigo = codigo_in
                        trnsccion.Commit()
                    End Using
                End Using
            End If
        Catch ex As Exception
            Throw ex
        End Try

        Return codigo
    End Function
    Public Function anular_eps(ByVal codigo As String) As Boolean
        Try
            Using Consulta As New SqlCommand("SP_EPS_ANULAR")
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
            Throw ex
            Return False
        End Try
    End Function

End Class
