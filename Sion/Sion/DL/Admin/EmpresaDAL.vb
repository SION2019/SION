Imports System.Data.SqlClient
Imports System.IO
Public Class EmpresaDAL
    Public Function guardarEmpresa(objEmpresa As Empresa) As String

        Try

            Using comando As New SqlCommand("SP_EMPRESA_CREAR")
                comando.CommandType = CommandType.StoredProcedure
                comando.Connection = FormPrincipal.cnxion
                comando.Parameters.Add(New SqlParameter("@NIT", SqlDbType.NVarChar)).Value = objEmpresa.nit
                comando.Parameters.Add(New SqlParameter("@DV", SqlDbType.NVarChar)).Value = objEmpresa.dv
                comando.Parameters.Add(New SqlParameter("@CODIGOH", SqlDbType.NVarChar)).Value = objEmpresa.codHabilitacion
                comando.Parameters.Add(New SqlParameter("@RAZONS", SqlDbType.NVarChar)).Value = objEmpresa.razonSocial
                comando.Parameters.Add(New SqlParameter("@DIRECCION", SqlDbType.NVarChar)).Value = objEmpresa.direccion
                comando.Parameters.Add(New SqlParameter("@TELEFONO", SqlDbType.NVarChar)).Value = objEmpresa.telefono
                comando.Parameters.Add(New SqlParameter("@CELULAR", SqlDbType.NVarChar)).Value = objEmpresa.celular
                comando.Parameters.Add(New SqlParameter("@PAIS", SqlDbType.NVarChar)).Value = objEmpresa.codPais
                comando.Parameters.Add(New SqlParameter("@DEPARTAMENTO", SqlDbType.NVarChar)).Value = objEmpresa.codDepartamento
                comando.Parameters.Add(New SqlParameter("@MUNICIPIO", SqlDbType.NVarChar)).Value = objEmpresa.codMunicipio
                comando.Parameters.Add(New SqlParameter("@EMAIL", SqlDbType.NVarChar)).Value = objEmpresa.email
                comando.Parameters.Add(New SqlParameter("@WEB", SqlDbType.NVarChar)).Value = objEmpresa.portalWeb
                comando.Parameters.Add(New SqlParameter("@SIGLA", SqlDbType.NVarChar)).Value = objEmpresa.sigla
                comando.Parameters.Add(New SqlParameter("@ENCABEZADO", SqlDbType.NVarChar)).Value = objEmpresa.encabezadoFactura
                comando.Parameters.Add(New SqlParameter("@PIE", SqlDbType.NVarChar)).Value = objEmpresa.pieFactura
                comando.Parameters.Add(New SqlParameter("@REPRESENTANTE", SqlDbType.NVarChar)).Value = objEmpresa.idRepresentante
                comando.Parameters.Add(New SqlParameter("@LOGO", SqlDbType.Image)).Value = objEmpresa.logoFactura
                comando.Parameters.Add(New SqlParameter("@LOGOPRINCIPAL", SqlDbType.Image)).Value = objEmpresa.logoPrincipal
                comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar)).Value = objEmpresa.usuario
                comando.Parameters.Add(New SqlParameter("@RESOLUCION", SqlDbType.NVarChar)).Value = objEmpresa.resolucion
                comando.Parameters.Add(New SqlParameter("@PREFIJO", SqlDbType.NVarChar)).Value = objEmpresa.prefijo
                comando.Parameters.Add(New SqlParameter("@CONSE_INI", SqlDbType.NVarChar)).Value = objEmpresa.conseInic
                comando.Parameters.Add(New SqlParameter("@CONSE_ACTUAL", SqlDbType.NVarChar)).Value = objEmpresa.conseActual
                comando.Parameters.Add(New SqlParameter("@CONSE_FIN", SqlDbType.NVarChar)).Value = objEmpresa.conseFin
                comando.Parameters.Add(New SqlParameter("@FECHAEXP", SqlDbType.Date)).Value = objEmpresa.fechaExpedicion
                comando.Parameters.Add(New SqlParameter("@FECHAVENCE", SqlDbType.Date)).Value = objEmpresa.fechaVencimiento
                comando.Parameters.Add(New SqlParameter("@VIGENTE", SqlDbType.NVarChar)).Value = 1
                comando.Parameters.Add(New SqlParameter("@pParametoItem", SqlDbType.NVarChar)).Value = Constantes.COMERCIALIZADORA
                comando.Parameters.Add(New SqlParameter("@valor", SqlDbType.NVarChar)).Value = objEmpresa.comercializadora
                objEmpresa.idEmpresa = CType(comando.ExecuteScalar, Integer)
            End Using

        Catch ex As Exception
            Throw ex
        End Try
        Return objEmpresa.idEmpresa
    End Function
    Public Sub actualizarEmpresa(objEmpresa As Empresa)

        Try
            Using Comando As New SqlCommand("SP_EMPRESA_ACTUALIZAR")
                Comando.CommandType = CommandType.StoredProcedure
                Comando.Connection = FormPrincipal.cnxion
                Comando.Parameters.Add(New SqlParameter("@NIT", SqlDbType.NVarChar))
                Comando.Parameters("@NIT").Value = objEmpresa.nit
                Comando.Parameters.Add(New SqlParameter("@DV", SqlDbType.NVarChar))
                Comando.Parameters("@DV").Value = objEmpresa.dv
                Comando.Parameters.Add(New SqlParameter("@CODIGOH", SqlDbType.NVarChar))
                Comando.Parameters("@CODIGOH").Value = objEmpresa.codHabilitacion
                Comando.Parameters.Add(New SqlParameter("@RAZONS", SqlDbType.NVarChar))
                Comando.Parameters("@RAZONS").Value = objEmpresa.razonSocial
                Comando.Parameters.Add(New SqlParameter("@DIRECCION", SqlDbType.NVarChar))
                Comando.Parameters("@DIRECCION").Value = objEmpresa.direccion
                Comando.Parameters.Add(New SqlParameter("@TELEFONO", SqlDbType.NVarChar))
                Comando.Parameters("@TELEFONO").Value = objEmpresa.telefono
                Comando.Parameters.Add(New SqlParameter("@CELULAR", SqlDbType.NVarChar))
                Comando.Parameters("@CELULAR").Value = objEmpresa.celular
                Comando.Parameters.Add(New SqlParameter("@PAIS", SqlDbType.NVarChar))
                Comando.Parameters("@PAIS").Value = objEmpresa.codPais
                Comando.Parameters.Add(New SqlParameter("@DEPARTAMENTO", SqlDbType.NVarChar))
                Comando.Parameters("@DEPARTAMENTO").Value = objEmpresa.codDepartamento
                Comando.Parameters.Add(New SqlParameter("@MUNICIPIO", SqlDbType.NVarChar))
                Comando.Parameters("@MUNICIPIO").Value = objEmpresa.codMunicipio
                Comando.Parameters.Add(New SqlParameter("@EMAIL", SqlDbType.NVarChar))
                Comando.Parameters("@EMAIL").Value = objEmpresa.email
                Comando.Parameters.Add(New SqlParameter("@WEB", SqlDbType.NVarChar))
                Comando.Parameters("@WEB").Value = objEmpresa.portalWeb
                Comando.Parameters.Add(New SqlParameter("@SIGLA", SqlDbType.NVarChar))
                Comando.Parameters("@SIGLA").Value = objEmpresa.sigla
                Comando.Parameters.Add(New SqlParameter("@ENCABEZADO", SqlDbType.NVarChar))
                Comando.Parameters("@ENCABEZADO").Value = objEmpresa.encabezadoFactura
                Comando.Parameters.Add(New SqlParameter("@PIE", SqlDbType.NVarChar))
                Comando.Parameters("@PIE").Value = objEmpresa.pieFactura
                Comando.Parameters.Add(New SqlParameter("@REPRESENTANTE", SqlDbType.NVarChar))
                Comando.Parameters("@REPRESENTANTE").Value = objEmpresa.idRepresentante
                Comando.Parameters.Add(New SqlParameter("@LOGO", SqlDbType.Image))
                Comando.Parameters("@LOGO").Value = objEmpresa.logoFactura
                Comando.Parameters.Add(New SqlParameter("@LOGOPRINCIPAL", SqlDbType.Image)).Value = objEmpresa.logoPrincipal
                Comando.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.NVarChar))
                Comando.Parameters("@USUARIO").Value = objEmpresa.usuario
                Comando.Parameters.Add(New SqlParameter("@IDEMPRESA", SqlDbType.NVarChar))
                Comando.Parameters("@IDEMPRESA").Value = objEmpresa.idEmpresa
                Comando.Parameters.Add(New SqlParameter("@pParametoItem", SqlDbType.NVarChar)).Value = Constantes.COMERCIALIZADORA
                Comando.Parameters.Add(New SqlParameter("@valor", SqlDbType.NVarChar)).Value = objEmpresa.comercializadora
                Comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

End Class

