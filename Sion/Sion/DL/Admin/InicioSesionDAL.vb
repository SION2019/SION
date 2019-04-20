Imports System.Data.SqlClient
Imports System.IO
Public Class InicioSesionDAL
    Public Function buscar_empresas_Empleado(USUARIO As String) As DataTable
        Dim dt As DataTable
        Dim dw As DataRow
        Dim cadena As String

        cadena = "execute SP_EMPRESA_USUARIO_CONSULTAR '" & USUARIO & "'"
        dt = New DataTable("Empresa")
        dt.Columns.Add("Id_empresa")
        dt.Columns.Add("Razon_Social")

        dw = dt.NewRow()
        dw.Item(0) = "-1"
        dw.Item(1) = "- Seleccione empresa -"

        dt.Rows.Add(dw)
        Try
            Using da = New SqlDataAdapter(cadena, FormPrincipal.cnxion)
                da.Fill(dt)
            End Using

        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
        Return dt
    End Function

    Public Function buscar_empresas() As DataTable
        Dim cadena As String
        Dim dt3 As New DataTable
        dt3.Clear()
        dt3.Columns.Add("Id_empresa")
        dt3.Rows.Add()

        Try
            cadena = "execute SP_EMPRESA_BUSCAR ''"

            Using da = New SqlDataAdapter(cadena, FormPrincipal.cnxion)
                da.Fill(dt3)
            End Using
        Catch ex As Exception
            dt3.Clear()
            General.manejoErrores(ex)
        End Try
        Return dt3
    End Function

    Public Function llenar_punto(valor As String, USUARIO As String) As DataTable
        'Dim da As SqlDataAdapter
        Dim dt2 As DataTable
        Dim dw2 As DataRow
        Dim cadena As String


        cadena = "exec SP_PUNTO_SERVICIO_EMPRESA_BUSCAR " & valor & ",'" & USUARIO & "'"
        dt2 = New DataTable("Punto")
        dt2.Columns.Add("Codigo_EP")
        dt2.Columns.Add("Nombre")

        dw2 = dt2.NewRow()
        dw2.Item(0) = "-1"
        dw2.Item(1) = "- Seleccione punto -"

        dt2.Rows.Add(dw2)
        Try
            Using da = New SqlDataAdapter(cadena, FormPrincipal.cnxion)
                da.Fill(dt2)
            End Using

        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
        Return dt2
    End Function


    Public Sub iniciarSesion(pUsuario As String, clave As String)
        Dim respuesta As Boolean
        Dim drDatosSesion As DataRow

        Try
            Dim params As New List(Of String)
            params.Add(pUsuario)
            params.Add(clave)
            params.Add(InicioSesion.ComboEmpresa.SelectedValue.ToString)
            params.Add(InicioSesion.Combopunto.SelectedValue.ToString.Trim)

            drDatosSesion = General.cargarItem("SP_EMPLEADO_INICIAR_SESION", params)

            If drDatosSesion IsNot Nothing Then
                SesionActual.idUsuario = drDatosSesion.Item("Id_empleado")
                SesionActual.idEmpresa = InicioSesion.ComboEmpresa.SelectedValue.ToString.Trim
                SesionActual.nombreEmpresa = InicioSesion.ComboEmpresa.Text.Trim
                SesionActual.nombreSede = InicioSesion.Combopunto.Text.ToUpper
                SesionActual.codigoEP = InicioSesion.Combopunto.SelectedValue.ToString.Trim
                SesionActual.usuario = InicioSesion.Textusuario.Text.Trim
                SesionActual.nombreCompleto = drDatosSesion.Item("Nombre")
                SesionActual.codigoPerfil = CInt(drDatosSesion.Item("Codigo_Perfil"))
                SesionActual.codigoEnlace = drDatosSesion.Item("Codigo_Enlace")

                FormPrincipal.Textnombre_completo.Text = drDatosSesion.Item("Nombre")
                FormPrincipal.Textusuario_actual.Text = drDatosSesion.Item("Nit")
                respuesta = True
            Else
                respuesta = False
            End If
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
        If respuesta = True Then
            FormPrincipal.Text = "Celer " & String.Format("Versión {0}", My.Application.Info.Version.ToString)
            FormPrincipal.cnxion.Close()
            logo.MdiParent = FormPrincipal
            logo.Show()
            FormPrincipal.Show()
            InicioSesion.Close()

        Else
            Exec.SacudirCrtl(InicioSesion)
            MsgBox("Contraseña incorrecta!!", MsgBoxStyle.Information)
            InicioSesion.Textcontraseña.Clear() : InicioSesion.Textcontraseña.Focus()
        End If

    End Sub

End Class
