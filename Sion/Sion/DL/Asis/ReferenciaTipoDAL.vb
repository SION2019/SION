Imports System.Data.SqlClient
Public Class ReferenciaTipoDAL
    Public Function guardar_tiporeferencias(ByVal codigo_in As String, ByVal descripcion As String) As String
        Dim codigo As String = ""
        Try
            If codigo_in = "" Then

                Using comando As New SqlCommand("SP_TIPOREFERENCIAS_CREAR")
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Connection = FormPrincipal.cnxion
                    comando.Parameters.Add(New SqlParameter("@NOMBRE", SqlDbType.NVarChar))
                    comando.Parameters("@NOMBRE").Value = descripcion
                    comando.Parameters.Add(New SqlParameter("@Usuario_creacion", SqlDbType.NVarChar))
                    comando.Parameters("@Usuario_creacion").Value = SesionActual.idUsuario
                    codigo = CType(comando.ExecuteScalar, String)
                End Using
            Else
                Using Consulta As New SqlCommand("SP_TIPOREFERENCIAS_ACTUALIZAR")
                    Consulta.CommandType = CommandType.StoredProcedure
                    Consulta.Connection = FormPrincipal.cnxion
                    Consulta.Parameters.Add(New SqlParameter("@Codigo", SqlDbType.NVarChar))
                    Consulta.Parameters("@Codigo").Value = codigo_in
                    Consulta.Parameters.Add(New SqlParameter("@NOMBRE", SqlDbType.NVarChar))
                    Consulta.Parameters("@NOMBRE").Value = descripcion
                    Consulta.Parameters.Add(New SqlParameter("@USUARIO_ACTUALIZACION", SqlDbType.NVarChar))
                    Consulta.Parameters("@USUARIO_ACTUALIZACION").Value = SesionActual.idUsuario
                    Consulta.ExecuteNonQuery()
                    codigo = codigo_in
                End Using
            End If
            'If verificar_formulario(Equivalencia) = True Then
            '    Equivalencia.Cargarcualidades()
            'End If
        Catch ex As Exception
            general.manejoErrores(ex) 
        End Try

        Return codigo
    End Function
    Public Function anular_tiporeferencia(ByVal codigo As String) As Boolean
        Try
            Using Consulta As New SqlCommand("SP_TIPOREFERENCIAS_ANULAR")
                Consulta.CommandType = CommandType.StoredProcedure
                Consulta.Connection = FormPrincipal.cnxion
                Consulta.Parameters.Add(New SqlParameter("@CODIGO", SqlDbType.Int))
                Consulta.Parameters("@CODIGO").Value = CInt(codigo)
                Consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                Consulta.Parameters("@USUARIO").Value = SesionActual.idUsuario
                Consulta.ExecuteNonQuery()
            End Using
            'If verificar_formulario(Equivalencia) = True Then
            '    Equivalencia.Cargarcualidades()
            'End If

            Return True
        Catch ex As Exception
            general.manejoErrores(ex) 
            Return False
        End Try
    End Function
    Private Function verificar_formulario(form As Windows.Forms.Form) As Boolean
        Dim valor As Boolean = False
        For Each f As Windows.Forms.Form In Application.OpenForms
            If f.Name = form.Name Then
                valor = True
            End If
        Next
        Return valor
    End Function

End Class
