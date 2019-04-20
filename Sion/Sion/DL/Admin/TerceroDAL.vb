Imports System.Data.SqlClient
Public Class TerceroDAL

    Public Shared Function getIdTercero(ByVal nit As String) As Integer

        Dim idTercero As Integer

        Using dbCommand As New SqlCommand("SELECT dbo.FN_GET_ID_TERCERO_BY_NIT(@nit)")
            dbCommand.Connection = FormPrincipal.cnxion
            dbCommand.Parameters.Add(New SqlParameter("@nit", SqlDbType.NVarChar)).Value = nit

            idTercero = IIf(dbCommand.ExecuteScalar() IsNot DBNull.Value, dbCommand.ExecuteScalar(), Nothing)
        End Using

        Return idTercero

    End Function

    Public Shared Sub GuardarTercero(objTercero As TerceroI)
        Try
            If Not String.IsNullOrEmpty(objTercero.codigo) Then
                Using consulta As New SqlCommand("SP_TERCERO_ACTUALIZAR")
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Parameters.Add(New SqlParameter("@Codigo_tipo_identificacion", SqlDbType.Int))
                    consulta.Parameters("@Codigo_tipo_identificacion").Value = objTercero.cboidentificacion
                    consulta.Parameters.Add(New SqlParameter("@Nit", SqlDbType.NVarChar))
                    consulta.Parameters("@Nit").Value = objTercero.identificacion
                    consulta.Parameters.Add(New SqlParameter("@DV", SqlDbType.NVarChar))
                    consulta.Parameters("@DV").Value = objTercero.dv
                    consulta.Parameters.Add(New SqlParameter("@RazonSocial", SqlDbType.NVarChar))
                    consulta.Parameters("@RazonSocial").Value = objTercero.razon
                    consulta.Parameters.Add(New SqlParameter("@Nombre", SqlDbType.NVarChar))
                    consulta.Parameters("@Nombre").Value = objTercero.pnombre
                    consulta.Parameters.Add(New SqlParameter("@sNombre", SqlDbType.NVarChar))
                    consulta.Parameters("@sNombre").Value = objTercero.nombre
                    consulta.Parameters.Add(New SqlParameter("@Apellido", SqlDbType.NVarChar))
                    consulta.Parameters("@Apellido").Value = objTercero.papellido
                    consulta.Parameters.Add(New SqlParameter("@sApellido", SqlDbType.NVarChar))
                    consulta.Parameters("@sApellido").Value = objTercero.sapellido
                    consulta.Parameters.Add(New SqlParameter("@Telefono1", SqlDbType.NVarChar))
                    consulta.Parameters("@Telefono1").Value = objTercero.telefono
                    consulta.Parameters.Add(New SqlParameter("@Telefono2", SqlDbType.NVarChar))
                    consulta.Parameters("@Telefono2").Value = objTercero.telefono2
                    consulta.Parameters.Add(New SqlParameter("@whatsapp", SqlDbType.NVarChar))
                    consulta.Parameters("@whatsapp").Value = objTercero.whatsapp
                    consulta.Parameters.Add(New SqlParameter("@email", SqlDbType.NVarChar))
                    consulta.Parameters("@email").Value = objTercero.email
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Pais", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Pais").Value = objTercero.Combopais
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Departamento", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Departamento").Value = objTercero.Combodepartamento
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Municipio", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Municipio").Value = objTercero.Combociudad
                    consulta.Parameters.Add(New SqlParameter("@Direccion", SqlDbType.NVarChar))
                    consulta.Parameters("@Direccion").Value = objTercero.direccion
                    consulta.Parameters.Add(New SqlParameter("@Usuario_ACTUALIZACION", SqlDbType.NVarChar))
                    consulta.Parameters("@Usuario_ACTUALIZACION").Value = objTercero.usuario
                    consulta.Parameters.Add(New SqlParameter("@Id_tercero", SqlDbType.NVarChar))
                    consulta.Parameters("@Id_tercero").Value = objTercero.codigo
                    consulta.ExecuteNonQuery()

                End Using
            Else
                Using consulta As New SqlCommand("SP_TERCERO_CREAR")
                    consulta.CommandType = CommandType.StoredProcedure
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Parameters.Add(New SqlParameter("@Codigo_tipo_identificacion", SqlDbType.Int))
                    consulta.Parameters("@Codigo_tipo_identificacion").Value = objTercero.cboidentificacion
                    consulta.Parameters.Add(New SqlParameter("@Nit", SqlDbType.NVarChar))
                    consulta.Parameters("@Nit").Value = objTercero.identificacion
                    consulta.Parameters.Add(New SqlParameter("@DV", SqlDbType.NVarChar))
                    consulta.Parameters("@DV").Value = objTercero.dv
                    consulta.Parameters.Add(New SqlParameter("@RazonSocial", SqlDbType.NVarChar))
                    consulta.Parameters("@RazonSocial").Value = objTercero.razon
                    consulta.Parameters.Add(New SqlParameter("@Nombre", SqlDbType.NVarChar))
                    consulta.Parameters("@Nombre").Value = objTercero.pnombre
                    consulta.Parameters.Add(New SqlParameter("@sNombre", SqlDbType.NVarChar))
                    consulta.Parameters("@sNombre").Value = objTercero.nombre
                    consulta.Parameters.Add(New SqlParameter("@Apellido", SqlDbType.NVarChar))
                    consulta.Parameters("@Apellido").Value = objTercero.papellido
                    consulta.Parameters.Add(New SqlParameter("@sApellido", SqlDbType.NVarChar))
                    consulta.Parameters("@sApellido").Value = objTercero.sapellido
                    consulta.Parameters.Add(New SqlParameter("@Telefono1", SqlDbType.NVarChar))
                    consulta.Parameters("@Telefono1").Value = objTercero.telefono
                    consulta.Parameters.Add(New SqlParameter("@Telefono2", SqlDbType.NVarChar))
                    consulta.Parameters("@Telefono2").Value = objTercero.telefono2
                    consulta.Parameters.Add(New SqlParameter("@whatsapp", SqlDbType.NVarChar))
                    consulta.Parameters("@whatsapp").Value = objTercero.whatsapp
                    consulta.Parameters.Add(New SqlParameter("@email", SqlDbType.NVarChar))
                    consulta.Parameters("@email").Value = objTercero.email
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Pais", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Pais").Value = objTercero.Combopais
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Departamento", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Departamento").Value = objTercero.Combodepartamento
                    consulta.Parameters.Add(New SqlParameter("@Codigo_Municipio", SqlDbType.NVarChar))
                    consulta.Parameters("@Codigo_Municipio").Value = objTercero.Combociudad
                    consulta.Parameters.Add(New SqlParameter("@Direccion", SqlDbType.NVarChar))
                    consulta.Parameters("@Direccion").Value = objTercero.direccion
                    consulta.Parameters.Add(New SqlParameter("@Usuario_creacion", SqlDbType.NVarChar))
                    consulta.Parameters("@Usuario_creacion").Value = objTercero.usuario
                    objTercero.codigo = CType(consulta.ExecuteScalar, String)

                End Using
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
