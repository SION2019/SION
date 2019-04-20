Imports System.Data.SqlClient
Imports System.IO
Imports System.Reflection
Public Class PrincipalDAL
    Dim dtPermisos As New DataTable
    Dim i As Integer
    Public Shared hashtableImagen As New Hashtable
    Public Function cargarOpciones(codigo_perfil As Integer, id_empresa As String) As DataTable
        Dim cadena As String
        Try
            cadena = "EXEC [SP_PERFIL_MODULO_CARGAR] " & codigo_perfil & ", " & SesionActual.codigoEP & "," & id_empresa & ""
            dtPermisos.Clear()

            Using da = New SqlDataAdapter(cadena, FormPrincipal.cnxion)
                da.Fill(dtPermisos)
            End Using

        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
        Return dtPermisos
        dtPermisos.Dispose()
    End Function


    Public Sub cargarImagenEmpresa(id_empresa As String)
        Try
            Dim bites() As Byte = Funciones.empresaActual.logoPrincipal
            Dim ms As New MemoryStream(bites)
            FormPrincipal.BackgroundImage = Image.FromStream(ms)

        Catch ex As Exception
            General.manejoErrores(ex)
        End Try

    End Sub

    Public Sub CreaOpciones(id_empresa As String, perfil_actual As String)
        Dim dsMenu As New DataSet

        Try
            Using consulta = New SqlCommand
                consulta.Connection = FormPrincipal.cnxion
                consulta.CommandText = " SP_EMPRESA_MODULO_CARGAR_MENU " & SesionActual.codigoEP & ", " & perfil_actual & "," & id_empresa & ""
                Dim da = New SqlDataAdapter(consulta)
                'FormPrincipal.dsDatos = New DataSet()
                da.Fill(dsMenu, "UsuariosMenu")

                ' obtener los elementos para crear las opciones de menú de primer nivel (barra de menú)
                Dim aDrFilas As DataRow()
                aDrFilas = dsMenu.Tables(0).Select("Padre_Menu is null", "Codigo_Menu")

                ' instanciar el menú
                ' recorrer las filas e ir creando la estructura de menú
                Dim mnuOpcion As ToolStripMenuItem

                For Each drFila As DataRow In aDrFilas
                    mnuOpcion = New ToolStripMenuItem(drFila("Descripcion_Menu").ToString())
                    mnuOpcion.Tag = New ElementoMenu(drFila("Codigo_Menu").ToString(),
                                                     drFila("Formulario").ToString,
                                                     Nothing,
                                                     drFila("Descripcion_menu").ToString)

                    ' añadir este menú desplegable a la barra de menú
                    mnuOpcion.BackColor = Color.SandyBrown
                    mnuOpcion.BackgroundImage = My.Resources.menu
                    FormPrincipal.mnuMenuUsuario.Items.Add(mnuOpcion)

                    ' recorrer si hubiera las opciones dependientes de este menú
                    Dim dtFilas As New DataTable
                    dtFilas = dsMenu.Tables(0)
                    CrearSubopciones(mnuOpcion, dtFilas)
                Next
                ' añadir el menú al formulario
                FormPrincipal.Controls.Add(FormPrincipal.mnuMenuUsuario)
                FormPrincipal.mnuMenuUsuario.LayoutStyle = ToolStripLayoutStyle.Flow
                FormPrincipal.mnuMenuUsuario.AutoSize = True
                FormPrincipal.mnuMenuUsuario.BackColor = Color.White
                FormPrincipal.mnuMenuUsuario.GripStyle = ToolStripGripStyle.Visible
                FormPrincipal.mnuMenuUsuario.BackgroundImage = My.Resources.menu
                FormPrincipal.mnuMenuUsuario.Font = New Font(Constantes.TIPO_LETRA_ELEMENTO_ARIAL, 8)
                FormPrincipal.mnuMenuUsuario.Renderer = New MyRenderer()

            End Using

        Catch ex As Exception
            General.manejoErrores(ex)
        End Try

    End Sub
    Public Sub eliminarMenu()
        For t = 0 To FormPrincipal.mnuMenuUsuario.Items.Count - 1
            FormPrincipal.mnuMenuUsuario.Items.RemoveAt(0)
        Next
        Dim con As Control
        For controlIndex As Integer = FormPrincipal.Controls.Count - 1 To 0 Step -1
            con = FormPrincipal.Controls(controlIndex)
            If con.Name = "mnuMenuUsuario" Then
                FormPrincipal.Controls.Remove(con)
            End If
        Next
        Dim con1 As Control
        For controlIndex As Integer = FormPrincipal.Controls.Count - 1 To 0 Step -1
            con1 = FormPrincipal.Controls(controlIndex)
            If con1.Name = "mnuOpcion" Then
                FormPrincipal.Controls.Remove(con1)
            End If
        Next
    End Sub
    ' crear las opciones de un menú desplegable
    Private Sub CrearSubopciones(ByVal mnuOpcion As ToolStripMenuItem, dtFilas As DataTable)
        Dim mnuSubOpcion As ToolStripMenuItem

        Dim drFilas = dtFilas.Select("Padre_menu =" & "'" & mnuOpcion.Tag.codigo & "'", "Codigo_Menu")
        For Each drFila As DataRow In drFilas
            mnuSubOpcion = New ToolStripMenuItem(drFila("Descripcion_Menu").ToString)
            mnuSubOpcion.Tag = New ElementoMenu(drFila("Codigo_Menu").ToString(),
                                                drFila("Formulario").ToString,
                                                (mnuOpcion.Tag.codigo.ToString.Substring(0, 2)),
                                                (mnuOpcion.Tag.nombrePadre.ToString))
            agregarMenuItem(drFila("Formulario").ToString, mnuSubOpcion)
            If Not String.IsNullOrEmpty(mnuSubOpcion.Tag.Nombre) Then
                mnuSubOpcion.Image = hashtableImagen(mnuSubOpcion.Tag.Nombre)

            End If
            mnuOpcion.DropDownItems.Add(mnuSubOpcion)
            ' llamar recursivamente a este método por si a partir de esta subopción se desplegaran
            ' más opciones
            CrearSubopciones(mnuSubOpcion, dtFilas)
        Next
    End Sub

    Public Function Cargar_alertas() As DataTable
        Dim dsAlertas As New DataSet
        Dim params As New List(Of String)
        params.Add(SesionActual.codigoPerfil)
        params.Add(SesionActual.codigoEP)
        General.llenarDataSet(Consultas.NOTIFICACION_LISTAR, params, dsAlertas)
        Dim tablacombinada As New DataTable
        For Each tabla As DataTable In dsAlertas.Tables
            tablacombinada.Merge(tabla)
        Next
        Return tablacombinada
    End Function

    Public Sub agregarMenuItem(ByVal form As String, ByVal mnuSubOpcion As ToolStripMenuItem)

        Select Case form
            Case "CerrarS"
                AddHandler mnuSubOpcion.Click, AddressOf CerrarS
            Case "Salir"
                AddHandler mnuSubOpcion.Click, AddressOf salir
            Case Else
                AddHandler mnuSubOpcion.Click, AddressOf cargarFormulario
        End Select


        If form.Trim <> String.Empty Then
            AddHandler mnuSubOpcion.Click, AddressOf FormPrincipal.click_Global
        End If

    End Sub

    Public Function estaAbierto(ByVal Myform As String) As Boolean
        Dim objForm As Windows.Forms.Form

        For Each objForm In My.Application.OpenForms
            Dim objName = Trim(objForm.Name)
            If (objName = Trim(Myform)) Then
                Return True
            End If
        Next
        Return False

    End Function

    Public Function estaAbiertoTipo(clase As Form, Optional cerrar As Boolean = False) As Boolean
        Dim objForm As Windows.Forms.Form

        For Each objForm In My.Application.OpenForms
            If objForm.GetType = clase.GetType Then
                Return True
            End If
        Next
        Return False

    End Function



    Public Function estaAbierto(ByVal elemento As ElementoMenu)
        Dim objForm As Windows.Forms.Form

        For Each objForm In My.Application.OpenForms
            Dim objName = Trim(objForm.Name)

            If (objName = Trim(elemento.nombre) AndAlso
                (IsNothing(objForm.Tag) OrElse (objForm.Tag.modulo = elemento.modulo AndAlso
                objForm.Tag.codigo = elemento.codigo))) Then
                Return True
            End If
        Next
        Return False

    End Function

    Public Sub traerAlFrente(ByVal Myform As String)
        Dim objForm As Windows.Forms.Form
        For Each objForm In My.Application.OpenForms
            If (Trim(objForm.Name) = Trim(Myform)) Then
                objForm.Focus()
                If objForm.WindowState = FormWindowState.Minimized Then
                    objForm.WindowState = FormWindowState.Normal
                End If
                Exit For
            End If
        Next
    End Sub

    Public Sub traerAlFrente(ByVal elemento As ElementoMenu)
        Dim objForm As Windows.Forms.Form
        For Each objForm In My.Application.OpenForms
            If (Trim(objForm.Name) = Trim(elemento.nombre) AndAlso objForm.Tag.Modulo = elemento.modulo AndAlso objForm.Tag.codigo = elemento.codigo) Then
                objForm.Focus()
                If objForm.WindowState = FormWindowState.Minimized Then
                    objForm.WindowState = FormWindowState.Normal
                End If
                Exit For
            End If
        Next
    End Sub
    Public Sub cerrarFormulario(ByVal elemento As ElementoMenu)
        Dim objForm As Windows.Forms.Form
        For Each objForm In My.Application.OpenForms
            If (Trim(objForm.Name) = Trim(elemento.nombre) AndAlso objForm.Tag.Modulo = elemento.modulo AndAlso objForm.Tag.codigo = elemento.codigo) Then
                objForm.Close()
                Exit For
            End If
        Next
    End Sub
    Public Sub cerrarFormulario(ByVal Myform As String)
        Dim objForm As Windows.Forms.Form
        For Each objForm In My.Application.OpenForms
            If (Trim(objForm.Name) = Trim(Myform)) Then
                objForm.Close()
                Exit For
            End If
        Next
    End Sub

    'Falta por implementar
    Public Sub guardar_cama()
    End Sub

    Public Sub guardarVersion(objPrincipa As Principal)
        Try
            Using comando = New SqlCommand
                comando.CommandType = CommandType.StoredProcedure
                comando.Connection = FormPrincipal.cnxion
                comando.CommandText = "SP_EQUIPOS_GUARDAR_VERSION"
                comando.Parameters.Add(New SqlParameter("@nombre", SqlDbType.NVarChar)).Value = objPrincipa.nombre
                comando.Parameters.Add(New SqlParameter("@fecha", SqlDbType.Date)).Value = objPrincipa.fecha
                comando.Parameters.Add(New SqlParameter("@version", SqlDbType.NVarChar)).Value = objPrincipa.version
                comando.Parameters.Add(New SqlParameter("@observacion", SqlDbType.NVarChar)).Value = objPrincipa.observacion
                comando.Parameters.Add(New SqlParameter("@usuario", SqlDbType.Int)).Value = objPrincipa.usuario
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub registroAuditorExterno(objPrincipa As Principal)
        Try
            Using comando = New SqlCommand
                comando.CommandType = CommandType.StoredProcedure
                comando.Connection = FormPrincipal.cnxion
                comando.CommandText = "SP_AUDITOR_EXTERNO_CREAR"
                comando.Parameters.Add(New SqlParameter("@idempresa", SqlDbType.Int)).Value = objPrincipa.idEmpresa
                comando.Parameters.Add(New SqlParameter("@idempleado", SqlDbType.Int)).Value = objPrincipa.idEmpleado
                comando.Parameters.Add(New SqlParameter("@codigoperfil", SqlDbType.Int)).Value = objPrincipa.codigoPerfil
                objPrincipa.codigoEntrada = comando.ExecuteScalar()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub salidaAuditorExterno(objPrincipa As Principal)
        Try
            Using comando = New SqlCommand
                comando.CommandType = CommandType.StoredProcedure
                comando.Connection = FormPrincipal.cnxion
                comando.CommandText = "SP_AUDITOR_EXTERNO_SALIDA"
                comando.Parameters.Add(New SqlParameter("@codigo", SqlDbType.Int)).Value = objPrincipa.codigoEntrada
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub AuditorExternoRegistroPaciente(objPrincipa As Principal)
        Try
            Using comando = New SqlCommand
                comando.CommandType = CommandType.StoredProcedure
                comando.Connection = FormPrincipal.cnxion
                comando.CommandText = "SP_AUDITOR_EXTERNO_REGISTRO_PACIENTE"
                comando.Parameters.Add(New SqlParameter("@codigo", SqlDbType.Int)).Value = objPrincipa.codigoEntrada
                comando.Parameters.Add(New SqlParameter("@registro", SqlDbType.Int)).Value = objPrincipa.registro
                comando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub guardar_empleados()
        Dim cadena As String

        cadena = " (getdate(),'" & SesionActual.idUsuario & "','" & My.Computer.Name & "') "
        Try

            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.CommandText = "INSERT INTO A_INGRESO_SISTEMA(Fecha,Usuario_creacion,Nombre_Maquina) values " & cadena
                    consulta.ExecuteNonQuery()
                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        Try
                            trnsccion.Rollback()
                            General.manejoErrores(ex)
                        Catch ex1 As Exception
                            MsgBox(ex1.Message)
                        End Try
                    End Try
                End Using

            End Using
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
    End Sub

    Private Sub CerrarS(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If estaAbiertoTipo(Form_Historia_clinica) Then
            MsgBox(Mensajes.VENTANA_ACTIVA, MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        If (SesionActual.codigoPerfil = Constantes.AMBUQ Or SesionActual.codigoPerfil = Constantes.COMFACOR Or SesionActual.codigoPerfil = Constantes.NUEVA_EPS) Then
            Dim objPrincipal As New Principal
            Dim lista As New List(Of String)
            lista.Add(SesionActual.codigoPerfil)
            objPrincipal.codigoEntrada = General.getValorConsulta(Consultas.AUDITOR_EXTERNO_GET_CODIGO_ENTRADA, lista)
            salidaAuditorExterno(objPrincipal)
        End If
        InicioSesion.Textusuario.Text = SesionActual.usuario
        InicioSesion.Show()
        InicioSesion.textusuario_Leave(Nothing, Nothing)
        InicioSesion.Textcontraseña.Focus()
        FormPrincipal.Dispose()
    End Sub

    Private Sub CerrarSesionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        InicioSesion.Textusuario.Text = SesionActual.usuario
        InicioSesion.textusuario_Leave(Nothing, Nothing)
        InicioSesion.Show()
        InicioSesion.Textcontraseña.Focus()
        FormPrincipal.Dispose()
    End Sub
    Private Sub BuscarActualizacionesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim version_actu, num_version, cadena, nombre_servidor As String
        FormPrincipal.Cursor = Cursors.WaitCursor
        version_actu = ""
        Try

            Using consulta As New SqlCommand("SELECT AV.Version FROM A_VERSION AS AV", FormPrincipal.cnxion)
                Using resultado = consulta.ExecuteReader()
                    If resultado.HasRows = True Then
                        resultado.Read()
                        version_actu = resultado.Item("Version")
                    End If
                End Using
            End Using
            num_version = My.Application.Info.Version.ToString

        Catch ex As Exception
            General.manejoErrores(ex)
            Exit Sub
        Finally
            FormPrincipal.Cursor = Cursors.Default
        End Try

        If (version_actu > num_version AndAlso MsgBox("Existe una version de Celer mas reciente, ¿Desea actualizar el Celer ahora mismo?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Atención") = MsgBoxResult.Yes) OrElse
          (version_actu <= num_version AndAlso MsgBox("Usted tiene la version mas reciente de Celer, ¿Desea actualizar el Celer de todas maneras?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Atención") = MsgBoxResult.Yes) Then
            FormPrincipal.Timer_abrir.Stop()
            nombre_servidor = Inicio.nombre_servidor
            If SalirToolStripMenuItem_Click(sender, e) Then
                cadena = Application.StartupPath & "\Sifaho_Actualizacion.exe"
                If My.Computer.FileSystem.FileExists(cadena) Then Process.Start(cadena, nombre_servidor)
            Else
                FormPrincipal.Timer_abrir.Start()
                MsgBox("No fue posible instalar la nueva version de Celer, Puede volverlo a intentar.", 64, "Error al Actualizar.")
            End If
        End If
    End Sub

    Private Function SalirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) As Boolean
        If MsgBox("¿Realmente desea salir de Celer?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, TitulosForm.SALIR) = MsgBoxResult.Yes Then
            Inicio.Close()
            FormPrincipal.Dispose()
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub cargarFormulario(sender As Object, e As EventArgs)
        Try
            Dim menuItem = DirectCast(sender, ToolStripMenuItem)
            Dim elemMenu As ElementoMenu = menuItem.Tag

            If estaAbierto(elemMenu) = True Then
                traerAlFrente(elemMenu)
            Else
                Dim nombreTipo = "Celer." & elemMenu.nombre
                Dim vTipo As Type = Assembly.GetExecutingAssembly.GetType(nombreTipo)

                If vTipo IsNot Nothing Then
                    Dim vFormulario = Activator.CreateInstance(vTipo)
                    vFormulario.tag = elemMenu
                    General.cargarForm(vFormulario)
                End If
            End If
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try

    End Sub

    Public Sub salir(sender As Object, e As EventArgs)
        If estaAbiertoTipo(Form_Historia_clinica) Then
            MsgBox(Mensajes.VENTANA_ACTIVA, MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        If MsgBox("¿Realmente desea salir de CELER?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Salir") = MsgBoxResult.Yes Then
            FormPrincipal.Close()
        End If
    End Sub

    Public Shared Sub cargarLogoFormulario()

        Dim dt As New DataTable
        Try
            General.llenarTabla2doPlanoConExcepcion("EXEC SP_MENU_NOMBRE_FORMULARIOS_LISTAR", Nothing, dt)
            For i As Int32 = 0 To dt.Rows.Count - 1
                Dim nombreTipo = "Celer." & dt.Rows(i).Item(0)
                Dim vTipo As Type = Assembly.GetExecutingAssembly.GetType(nombreTipo)
                If vTipo IsNot Nothing Then
                    Dim vFormulario = Activator.CreateInstance(vTipo)
                    hashtableImagen.Add(dt.Rows(i).Item(0), vFormulario.muestraImagen())
                End If
            Next
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
    End Sub

End Class
