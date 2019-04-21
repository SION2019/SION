Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading
Imports System.Reflection

Public Class General

    'general de busquedas sin que se cierre el formualrio de busqueda
    Public Delegate Sub cargaInfoForm(ByVal codigo As String)
    Public Delegate Sub cargaInfoFormObj(ByVal fila As DataRow)
    Public Delegate Sub subRutina()
    Public Delegate Sub subProceso(ByRef obj As Correo)
    Shared hilo As Thread
    Public Shared Sub busquedaItems(nombreProc As String,
                                    params As List(Of String),
                                    titulo As String,
                                    datagrid As DataGridView,
                                    datatable As DataTable,
                                    indColumnaInicio As Integer,
                                    indColumnaFin As Integer,
                                    indColumnaFoco As Integer,
                                    Optional ocultarCodigo As Boolean = False,
                                    Optional noRepetir As Boolean = False,
                                    Optional indColumnaNoRepetible As Integer = 0,
                                    Optional cerrarBusqueda As Boolean = False,
                                    Optional pMetodo As subRutina = Nothing,
                                    Optional pTamaño As Integer = Constantes.TAMANO_MEDIANO)
        Dim seguir As Boolean = False
        Dim seleccionados As String = ""
        Do
            Dim tbl As DataRow = Nothing
            If noRepetir = True Then
                seleccionados = getCadenaSeleccionados(datatable, seleccionados, indColumnaNoRepetible)
                params.Add(seleccionados)
            End If

            Dim formBusqueda As Form_BusquedaGeneral = buscarElemento(nombreProc, params, titulo, ocultarCodigo, pTamaño, True)
            If noRepetir = True Then
                params.RemoveAt(params.Count - 1)
            End If

            tbl = formBusqueda.rowResultados
            Dim ultimaFila = datatable.Rows.Count - 1
            If tbl IsNot Nothing Then
                If datatable.Rows.Count = 0 OrElse datatable.Rows(ultimaFila).Item(indColumnaInicio).ToString <> "" Then
                    datatable.Rows.Add()
                    ultimaFila = datatable.Rows.Count - 1
                End If

                For indColumna = indColumnaInicio To indColumnaFin
                    datatable.Rows(ultimaFila).Item(indColumna) = tbl.Item(indColumna - indColumnaInicio).ToString
                Next
                If Not IsNothing(datatable.Columns("Estado")) Then
                    datatable.Rows(ultimaFila).Item("Estado") = Constantes.ITEM_NUEVO
                End If
                datatable.Rows.Add()
                ultimaFila = datatable.Rows.Count - 1
                'datatable.AcceptChanges()
                datagrid.EndEdit()
                Try
                    datagrid.Rows(ultimaFila - 1).Cells(indColumnaFoco).Selected = True
                Catch ex As Exception

                End Try
                seguir = True
                If Not IsNothing(pMetodo) Then
                    pMetodo.Invoke()
                End If
            Else
                seguir = False
            End If

        Loop While (seguir = True AndAlso cerrarBusqueda = False)
    End Sub

    'Dado un objeto habilita todos los elementos contenidos en el de forma recursiva
    Public Shared Sub habilitarControles(ByRef pElemento As Object)
        For Each vItem In pElemento.Controls
            If ((TypeOf vItem Is TextBox) Or (TypeOf vItem Is RichTextBox) Or (TypeOf vItem Is MaskedTextBox) Or (TypeOf vItem Is DataGridView)) And
                   Not (vItem.name.ToString.ToLower.Contains(Constantes.TEXTBOX_CODIGO)) Then
                vItem.readonly = False
            ElseIf (TypeOf vItem Is CheckBox) Or (TypeOf vItem Is RadioButton) Or (TypeOf vItem Is ComboBox) Or
                (TypeOf vItem Is Button) Or (TypeOf vItem Is TreeView) Or (TypeOf vItem Is DateTimePicker) Or (TypeOf vItem Is NumericUpDown) Then
                vItem.enabled = True
            ElseIf (TypeOf vItem Is GroupBox) Or (vItem.hasChildren) Then
                habilitarControles(vItem)
            End If
        Next
    End Sub
    'Dado un objeto deshabilita todos los elementos contenidos en el de forma recursiva
    Public Shared Sub deshabilitarControles(ByRef pElemento As Object)
        Dim vItem As Object

        For Each vItem In pElemento.Controls
            If (TypeOf vItem Is TextBox) Or (TypeOf vItem Is RichTextBox) Or (TypeOf vItem Is MaskedTextBox) Or (TypeOf vItem Is DataGridView) Then
                vItem.readonly = True
            ElseIf (TypeOf vItem Is CheckBox) Or (TypeOf vItem Is RadioButton) Or (TypeOf vItem Is ComboBox) Or
                   ((TypeOf vItem Is Button) And Not vItem.name.ToString.ToLower.Contains(Constantes.BOTON_OPCION)) Or (TypeOf vItem Is TreeView) Or (TypeOf vItem Is DateTimePicker) Or (TypeOf vItem Is NumericUpDown) Then
                vItem.enabled = False
            ElseIf (TypeOf vItem Is GroupBox) Or (vItem.hasChildren) Then
                deshabilitarControles(vItem)
            End If
        Next

    End Sub
    '-----ARBOL----------------------------------------------------------------
    Public Shared Sub cargarOpcionesArbol(ByVal consulta As String,
                                    ByVal params As List(Of String),
                                    ByVal arbol As TreeView,
                                    ByVal vlrName As String,
                                    ByVal vlrText As String)
        Dim dtArbol As New DataTable

        Dim listaParams As String = Funciones.getParametros(params)
        Try
            Using daAdapter = New SqlDataAdapter(consulta & listaParams, FormPrincipal.cnxion)
                daAdapter.Fill(dtArbol)
            End Using
            Dim nodo As TreeNode
            Dim dFila As DataRow()
            arbol.Nodes.Clear()
            arbol.ExpandAll()

            dFila = dtArbol.Select()
            For Each Fila As DataRow In dFila
                nodo = New TreeNode
                nodo.Name = Fila.Item(vlrName)
                nodo.Text = Fila.Item(vlrText)
                arbol.Nodes.Add(nodo)
            Next

        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
    End Sub


    'Abre la ruta para seleccionar el archivo de imagen y lo carga al formulario
    Public Shared Sub subirimagen(ByVal objeto As PictureBox, ByVal componente As OpenFileDialog, Optional soloPDF As Boolean = False, Optional objetoDestino As PictureBox = Nothing)
        Try

            With componente
                .InitialDirectory = ""
                If soloPDF Then
                    .Filter = "Todos los archivos de imagen|*.pdf"
                Else
                    .Filter = "Todos los archivos de imagen|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.pdf;*|PDF|*.pdf|JPEG|*.jpeg;*.jpg|BMP|*.bmp|GIF|*.gif|PNG|*.png"
                End If
                .Title = "Seleccionar Archivo"
            End With
            If componente.ShowDialog() = DialogResult.OK Then
                objeto.Image = Nothing
                objeto.SizeMode = PictureBoxSizeMode.StretchImage
                Dim documento = componente.FileName
                Dim aux = Path.GetExtension(componente.FileName).ToLower
                With objeto
                    If aux = ".pdf" Then
                        .Image = My.Resources.picpdf
                    Else
                        documento = Nothing
                        .Image = Image.FromFile(componente.FileName)
                    End If
                End With
                If IsNothing(objetoDestino) Then Exit Sub
                redimencionarImagen(objeto, objetoDestino)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Shared Sub redimencionarImagen(pImagen As PictureBox, pImagenDestino As PictureBox)
        ' Cree un mapa de bits del contenido del control de fileUpload en la memoria
        Dim originalBMP As Bitmap = New Bitmap(pImagen.Image)
        ' Calcule las nuevas dimensiones de imagen
        Dim origWidth As Integer = originalBMP.Width
        Dim origHeight As Integer = originalBMP.Height
        Dim sngRatio As Double
        Dim newHeight As Integer
        Dim newWidth As Integer
        If origHeight > origWidth AndAlso origHeight > Constantes.TAMANO_VERTICAL_IMAGEN_PIX Then
            sngRatio = origHeight / origWidth
            newHeight = Constantes.TAMANO_VERTICAL_IMAGEN_PIX
            newWidth = newHeight / sngRatio
        ElseIf origHeight < origWidth AndAlso origWidth > Constantes.TAMANO_HORIZONTAL_IMAGEN_PIX Then
            sngRatio = origWidth / origHeight
            newWidth = Constantes.TAMANO_HORIZONTAL_IMAGEN_PIX
            newHeight = newWidth / sngRatio
        Else
            pImagenDestino.Image = pImagen.Image
            Exit Sub
        End If
        ' Cree un nuevo mapa de bits que sostendrá el mapa de bits anterior redimensionado
        Dim newBMP As New Bitmap(originalBMP, newWidth, newHeight)
        ' Cree un gráfico basado en el nuevo mapa de bits
        Dim oGraphics As Graphics = Graphics.FromImage(newBMP)
        ' Ponga las propiedades para el nuevo archivo gráfico
        oGraphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        oGraphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        ' Dibuje el nuevo gráfico basado en el mapa de bits redimensionado
        oGraphics.DrawImage(originalBMP, 0, 0, newWidth, newHeight)
        ' Guearde el nuevo archivo gráfico al servidor
        Dim direccion As String
        direccion = Path.GetTempPath & "imagen" & Format(Now, "HH'h'mm'm'ss's'") & ".png"
        Try
            newBMP.Save(direccion)
            pImagenDestino.Image = Image.FromFile(direccion)
            'Una vez terminado con los objetos de mapa de bits, los desasignamos.

            originalBMP.Dispose()
            newBMP.Dispose()
            oGraphics.Dispose()
        Catch ex As Exception
            MsgBox("ha ocurrido un error, por favor vuelva a intentarlo.", MsgBoxStyle.Critical)
        End Try

    End Sub
    Public Shared Sub abrirJustificacion(dgv As DataGridView, dt As DataTable, panel As Panel, txtJustificacion As TextBox, nombreColumna As String, soloLectura As Boolean, Optional teclaPresionada As String = "", Optional ultimaFila As Boolean = False)
        If dgv.RowCount > 0 AndAlso dgv.Rows(dgv.CurrentRow.Index).Cells(nombreColumna).Selected = True AndAlso (String.IsNullOrEmpty(dt.Rows(dgv.CurrentCell.RowIndex).Item(0).ToString) = False OrElse ultimaFila = True) Then
            If String.IsNullOrEmpty(teclaPresionada) = False Then
                Dim inte As Integer = Asc(teclaPresionada) ''esto solo es para cuando quiera hacer debug
            End If

            If (String.IsNullOrEmpty(teclaPresionada) = False) AndAlso Asc(teclaPresionada) = 13 Then Exit Sub
            panel.Visible = True
            If (String.IsNullOrEmpty(teclaPresionada) = False) AndAlso ((Asc(teclaPresionada) >= 65 And Asc(teclaPresionada) <= 90) _
                  Or (Asc(teclaPresionada) >= 97 And Asc(teclaPresionada) <= 122) _
                  Or (teclaPresionada = "ñ") Or (teclaPresionada = "Ñ")) AndAlso
                  String.IsNullOrEmpty(dgv.Rows(dgv.CurrentRow.Index).Cells(nombreColumna).Value.ToString) Then
                If soloLectura Then teclaPresionada = ""
                txtJustificacion.Text = teclaPresionada
            Else
                If (String.IsNullOrEmpty(teclaPresionada) = False) AndAlso ((Asc(teclaPresionada) >= 65 And Asc(teclaPresionada) <= 90) _
                  Or (Asc(teclaPresionada) >= 97 And Asc(teclaPresionada) <= 122) _
                  Or (Asc(teclaPresionada) >= 48 And Asc(teclaPresionada) <= 57) _
                  Or (teclaPresionada = "ñ") Or (teclaPresionada = "Ñ")) Then
                    txtJustificacion.Text = dgv.Rows(dgv.CurrentRow.Index).Cells(nombreColumna).Value.ToString & teclaPresionada
                Else
                    txtJustificacion.Text = dgv.Rows(dgv.CurrentRow.Index).Cells(nombreColumna).Value.ToString
                End If
            End If

            If Not IsNothing(dgv.Columns("Estado")) Then
                dt.Rows(dgv.CurrentCell.RowIndex).Item("Estado") = Constantes.ITEM_NUEVO
            End If
            txtJustificacion.ReadOnly = soloLectura
            txtJustificacion.Focus()
            txtJustificacion.SelectionStart = txtJustificacion.TextLength


        End If
    End Sub
    Public Shared Function subirPdf(ByVal objeto As Object, ByVal componente As OpenFileDialog) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("ruta")
        dt.Columns.Add("nombre")
        With componente
            .InitialDirectory = ""
            .Filter = "Todos los archivos de pdf|*.pdf|PDF|*.pdf"
            .Title = "Seleccionar Resultado"
            .Multiselect = False
        End With
        If componente.ShowDialog() = DialogResult.OK Then
            dt.Rows.Add()
            With objeto
                objeto.LoadFile(componente.FileName)
                dt.Rows(0).Item(0) = componente.FileName
                dt.Rows(0).Item(1) = componente.SafeFileName
            End With
            objeto.Visible = True
        End If
        Return dt
    End Function
    Public Shared Sub cargarImagen(usuario As String, id_empresa As Integer, ByRef pictbox As PictureBox)
        Try
            Using consulta = New SqlCommand("exec SP_EMPLEADO_FOTO_BUSCAR '" & usuario & "'," & id_empresa & "", FormPrincipal.cnxion)
                Using resultado = consulta.ExecuteReader()
                    If resultado.HasRows = True Then
                        resultado.Read()
                        Dim bites() As Byte
                        If Not IsDBNull(resultado.Item("Foto")) Then
                            bites = resultado.Item("Foto")
                            Dim ms As New MemoryStream(bites)
                            pictbox.Image = Image.FromStream(ms)
                            ms.Dispose()
                            bites = Nothing
                        Else
                            pictbox.Image = My.Resources.usuario
                        End If

                    Else
                        pictbox.Image = My.Resources.usuario
                    End If
                End Using
            End Using
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
    End Sub
    Public Shared Function subirimagenDiagnostica(ByVal componente As OpenFileDialog)
        Dim dt As New DataTable
        Dim arrFile() As Byte = Nothing
        Dim contador As Integer
        dt.Columns.Add("nombreArchivo")
        dt.Columns.Add("Archivo", Type.GetType("System.Byte[]"))
        dt.Columns.Add("ruta")
        dt.Columns.Add("idArchivo")

        With componente
            .InitialDirectory = ""
            .Multiselect = True
            .Filter = "Todos los archivos de imagen|*.jpg;*.dcm;*.IMA|JPG|;*.jpg|Dicom|*.dcm;*.IMA|Video|*.mp4"
            .Title = "Seleccionar la Imagen"
        End With

        If componente.ShowDialog() = DialogResult.OK Then
            contador = 0
            While (componente.FileNames.Count - 1) >= contador
                arrFile = File.ReadAllBytes(componente.FileNames.ElementAt(contador))
                dt.Rows.Add()
                dt.Rows(contador).Item("ruta") = componente.FileNames.ElementAt(contador)
                dt.Rows(contador).Item("nombreArchivo") = Replace(componente.SafeFileNames.ElementAt(contador), "JPEG", "jpg")
                dt.Rows(contador).Item("Archivo") = arrFile
                contador = contador + 1
            End While
        End If

        Return dt

    End Function
    Public Shared Sub verImagen(ByVal objeto As PictureBox)
        If objeto.Image IsNot Nothing Then
            Dim tempfileurl As String = System.IO.Path.GetTempPath + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "-tempimg"
            Try
                Using bmp As New Bitmap(objeto.Image)
                    bmp.Save(tempfileurl & Constantes.FORMATO_PNG, objeto.Image.RawFormat)
                End Using
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Process.Start("file://" + tempfileurl)
        End If
    End Sub
    Public Shared Sub verArchivo(ByVal datos As Byte())
        Dim tempfileurl As String = System.IO.Path.GetTempPath + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "-tempimg"

        Try

            'tempfileurl = tempfileurl & Constantes.FORMATO_PNG
            'Using bmp As New Bitmap(objeto.Image)
            '    bmp.Save(tempfileurl, objeto.Image.RawFormat)
            'End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Process.Start("file://" + tempfileurl)
    End Sub
    'Habilita los botones de un ToolStrip recibido por parametro
    Public Shared Sub habilitarBotones(ByRef pToolStrip As ToolStrip)

        'Recorre y habilita cada elemento
        For Each oToolStripButton In pToolStrip.Items
            If TypeOf oToolStripButton Is ToolStripButton Then
                oToolStripButton.enabled = True
            ElseIf TypeOf oToolStripButton Is ToolStripDropDown Then
                oToolStripButton.enabled = True
            End If
        Next
    End Sub

    Public Shared Sub deshabilitarBotones(ByRef pToolStrip As ToolStrip)
        'Recorre y deshabilita cada elemento
        For Each oToolStripButton In pToolStrip.Items
            If TypeOf oToolStripButton Is ToolStripButton Then
                oToolStripButton.enabled = False
            ElseIf TypeOf oToolStripButton Is ToolStripDropDownButton Then
                oToolStripButton.enabled = False
            End If
        Next
    End Sub
    Public Shared Function getEstadoInventario(ByVal est_inv As Boolean) As String
        If est_inv = Constantes.PENDIENTE Then
            Return "Pendiente"
        Else
            Return "Terminada"
        End If
    End Function
    Public Shared Function getEstadoInventarioString(ByVal est_inv As String) As Integer
        If est_inv = "Pendiente" Then
            Return Constantes.PENDIENTE
        Else
            Return Constantes.TERMINADO
        End If
    End Function
    Public Shared Sub deshabilitarBotones_HC(ByRef pToolStrip As ToolStrip)
        'Recorre y habilita cada elemento
        For Each oToolStripButton In pToolStrip.Items
            If TypeOf oToolStripButton Is ToolStripButton And Not oToolStripButton.name.ToString.ToLower.Contains(Constantes.VISTA_PREVIA) Then
                oToolStripButton.enabled = False
            End If
        Next
    End Sub

    Public Shared Sub getConnReporte(Itblas As Tables)
        Dim connReporte As New ConnectionInfo

        connReporte.ServerName = FormPrincipal.cnxion.DataSource
        connReporte.DatabaseName = FormPrincipal.cnxion.Database
        connReporte.UserID = Constantes.SIFAHO_REPORTES_USER
        connReporte.Password = Constantes.SIFAHO_REPORTES_PASS

        connReporte.Type = ConnectionInfoType.SQL

        For Each tbla As Table In Itblas
            Dim boTableLogOnInfo As TableLogOnInfo = tbla.LogOnInfo
            boTableLogOnInfo.ConnectionInfo = connReporte
            tbla.ApplyLogOnInfo(boTableLogOnInfo)
        Next
    End Sub
    Public Shared Function consultarPacienteFacturado(registro As Integer) As Boolean
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(registro)

        General.llenarTabla(Consultas.CONSULTAR_PACIENTE_FACTURADO, params, dt)
        If dt.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Shared Sub cargarCombo(ByVal consulta As String,
                                  ByVal vlrDisplayMember As String,
                                  ByVal vlrValueMember As String,
                                  ByVal cbCombo As ComboBox)
        Dim dtTabla As New DataTable
        Try

            Dim drFila As DataRow = dtTabla.NewRow()
            dtTabla.Columns.Add(vlrValueMember)
            dtTabla.Columns.Add(vlrDisplayMember)
            drFila.Item(0) = "-1"
            drFila.Item(1) = " - - - Seleccione - - - "
            dtTabla.Rows.Add(drFila)
            Using da = New SqlDataAdapter(consulta, FormPrincipal.cnxion)
                da.Fill(dtTabla)
            End Using
            cbCombo.DataSource = dtTabla
            cbCombo.DisplayMember = vlrDisplayMember
            cbCombo.ValueMember = vlrValueMember
            'If TypeOf cbCombo Is ComboBox Then
            '    If dtTabla.Rows.Count > 10 Then
            '        cbCombo.DropDownStyle = ComboBoxStyle.DropDown
            '        cbCombo.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            '        cbCombo.AutoCompleteSource = AutoCompleteSource.ListItems
            '    Else
            cbCombo.DropDownHeight = 150
            cbCombo.AutoCompleteMode = AutoCompleteMode.None
            cbCombo.AutoCompleteSource = AutoCompleteSource.None
            cbCombo.DropDownStyle = ComboBoxStyle.DropDownList
            '    End If
            'End If
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
    End Sub

    Public Shared Sub cargarCombo(ByVal consulta As String,
                                  ByVal params As List(Of String),
                                  ByVal vlrDisplayMember As String,
                                  ByVal vlrValueMember As String,
                                  ByVal cbCombo As ComboBox)
        Dim dtTabla As New DataTable
        Try

            Dim drFila As DataRow = dtTabla.NewRow()
            dtTabla.Columns.Add(vlrValueMember)
            dtTabla.Columns.Add(vlrDisplayMember)
            drFila.Item(0) = "-1"
            drFila.Item(1) = " - - - Seleccione - - - "
            dtTabla.Rows.Add(drFila)
            Using da = New SqlDataAdapter(consulta & Funciones.getParametros(params), FormPrincipal.cnxion)
                da.Fill(dtTabla)
            End Using
            cbCombo.DataSource = dtTabla
            cbCombo.DisplayMember = vlrDisplayMember
            cbCombo.ValueMember = vlrValueMember

            If cbCombo IsNot Nothing Then
                'If dtTabla.Rows.Count > 10 Then
                '    cbCombo.DropDownStyle = ComboBoxStyle.DropDown
                '    cbCombo.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                '    cbCombo.AutoCompleteSource = AutoCompleteSource.ListItems
                'Else

                cbCombo.AutoCompleteMode = AutoCompleteMode.None
                cbCombo.AutoCompleteSource = AutoCompleteSource.None
                cbCombo.DropDownStyle = ComboBoxStyle.DropDownList
                'End If
            End If
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
    End Sub

    Public Shared Sub mensajeProximante()
        MsgBox(Mensajes.EN_CONSTRUCCION, MsgBoxStyle.Information)
    End Sub
    Public Shared Sub mensajeValidacion(ByVal msm As String)
        MsgBox(msm, MsgBoxStyle.Exclamation)
    End Sub
    Public Shared Sub mensajeError(ByVal msm As String)
        MsgBox(msm, MsgBoxStyle.Critical)
    End Sub

    Public Shared Sub cargarLista(ByVal consulta As String,
                                  ByVal params As List(Of String),
                                  ByVal vlrDisplayMember As String,
                                  ByVal vlrValueMember As String,
                                  ByVal lista As ListBox)
        Dim dtTabla As New DataTable
        Try

            Dim drFila As DataRow = dtTabla.NewRow()
            dtTabla.Columns.Add(vlrValueMember)
            dtTabla.Columns.Add(vlrDisplayMember)
            drFila.Item(0) = "-1"
            drFila.Item(1) = " - - - Seleccione - - - "
            dtTabla.Rows.Add(drFila)
            Using da = New SqlDataAdapter(consulta & Funciones.getParametros(params), FormPrincipal.cnxion)
                da.Fill(dtTabla)
            End Using
            lista.DataSource = dtTabla
            lista.DisplayMember = vlrDisplayMember
            lista.ValueMember = vlrValueMember
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
    End Sub

    Public Shared Sub agregarDiagnosticosCIE(ByVal datagrid As DataGridView, ByVal datatable As DataTable)
        Dim seguir As Boolean = False
        Dim seleccionados As String = ""
        Do
            seleccionados = getCadenaSeleccionados(datatable, seleccionados, 0)
            Dim tbl As DataRow = Nothing
            Dim params As New List(Of String)
            params.Add(seleccionados)
            Dim formBusqueda As Form_BusquedaGeneral = General.buscarElemento(Consultas.BUSQUEDA_DIAGNOSTICO_CIE, params, TitulosForm.BUSQUEDA_DIAGNOSTICO_CIE, False)
            tbl = formBusqueda.rowResultados
            If tbl IsNot Nothing Then
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(0) = tbl(0)
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(1) = tbl(1)
                If datatable.Columns.Count > 2 Then
                    datatable.Rows(datagrid.CurrentCell.RowIndex).Item(2) = Constantes.ITEM_NUEVO
                End If

                datatable.Rows.Add()
                datagrid.Rows(datagrid.RowCount - 1).Cells(1).Selected = True
                seguir = True
            Else
                seguir = False
            End If
        Loop While (seguir = True)
    End Sub
    Public Shared Sub agregarParametroDatatable(ByVal codigoColumna As String,
                                                ByVal valorColumna As String,
                                                ByVal codigoParametro As String,
                                                ByVal valorParametro As String,
                                                ByRef dtParametro As DataTable)

        Dim drRespuestas As DataRow = dtParametro.NewRow
        drRespuestas.Item(codigoColumna) = codigoParametro
        drRespuestas.Item(valorColumna) = valorParametro
        dtParametro.Rows.Add(drRespuestas)
    End Sub
    Public Shared Sub actualizarParametroValorDatatable(ByVal codigoColumna As String,
                                                        ByVal valorColumna As String,
                                                        ByVal codigoParametro As String,
                                                        ByVal valorParametro As String,
                                                        ByVal dtParametro As DataTable)
        For i = 0 To dtParametro.Rows.Count - 1
            If dtParametro.Rows(i).Item(codigoColumna) = codigoParametro And
                dtParametro.Rows(i).Item(valorColumna) <> valorParametro Then
                dtParametro.Rows(i).Item(valorColumna) = valorParametro
                Exit Sub
            End If
        Next
    End Sub

    Public Shared Sub eliminarPametroDatatable(ByVal codigoColumna As String,
                                               ByVal codigoParametro As String,
                                               ByVal dtParametro As DataTable)

        For i = 0 To dtParametro.Rows.Count - 1
            If dtParametro.Rows(i).Item(codigoColumna) = codigoParametro Then
                dtParametro.Rows.RemoveAt(i)
                Exit Sub
            End If
        Next
    End Sub

    Public Shared Sub agregarProductoCotizacion(ByVal datagrid As DataGridView, ByVal datatable As DataTable, ByVal Cliente As String)
        Dim seguir As Boolean = False
        Dim seleccionados As String = ""
        Do
            seleccionados = getCadenaSeleccionados(datatable, seleccionados, 0)
            Dim tbl As DataRow = Nothing
            Dim params As New List(Of String)
            params.Add(seleccionados)
            params.Add(Cliente)
            params.Add(SesionActual.codigoEP)
            Dim formBusqueda As Form_BusquedaGeneral = General.buscarElemento(Consultas.BUSQUEDA_COTIZACION_PRODUCTOS, params, TitulosForm.BUSQUEDA_PRODUCTOS_COTIZACION, False)
            tbl = formBusqueda.rowResultados
            If tbl IsNot Nothing Then
                datatable.Rows(datatable.Rows.Count - 1).Item("Codigo") = tbl(0)
                datatable.Rows(datatable.Rows.Count - 1).Item("Descripcion") = tbl(1)
                datatable.Rows(datatable.Rows.Count - 1).Item("Marca") = tbl(2)
                datatable.Rows(datatable.Rows.Count - 1).Item("Stock") = tbl(3)
                datatable.Rows(datatable.Rows.Count - 1).Item("Cantidad") = 0
                datatable.Rows(datatable.Rows.Count - 1).Item("Iva") = (4)
                datatable.Rows(datatable.Rows.Count - 1).Item("Precio") = tbl(5)

                datatable.Rows.Add()
                datagrid.Rows(datagrid.RowCount - 1).Cells("Codigo").Selected = True
                seguir = True
            Else
                seguir = False
            End If
        Loop While (seguir = True)
    End Sub
    Public Shared Sub agregarItems(ByVal consulta As String,
                                   ByVal titulo As String,
                                   ByVal datagrid As DataGridView,
                                   ByVal datatable As DataTable)
        Dim seguir As Boolean = False
        Dim seleccionados As String = ""
        Do
            seleccionados = getCadenaSeleccionados(datatable, seleccionados, 0)
            Dim tbl As DataRow = Nothing
            Dim params As New List(Of String)
            params.Add(seleccionados)
            Dim formBusqueda As Form_BusquedaGeneral = General.buscarElemento(consulta, params, titulo, False)
            tbl = formBusqueda.rowResultados
            If tbl IsNot Nothing Then
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(0) = tbl(0)
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(1) = tbl(1)
                If datatable.Columns.Count > 2 Then
                    datatable.Rows(datagrid.CurrentCell.RowIndex).Item(2) = Constantes.ITEM_NUEVO
                End If
                datatable.Rows.Add()
                datagrid.Rows(datagrid.RowCount - 1).Cells(1).Selected = True
                seguir = True
            Else
                seguir = False
            End If
        Loop While (seguir = True)
    End Sub

    Public Shared Sub agregarMiembrosComite(ByVal consulta As String, ByVal titulo As String, ByVal datagrid As Object, ByVal datatable As Object, empresa As Integer)
        Dim seleccionados As String = ""
        Dim tbl As DataRow = Nothing
        Dim params As New List(Of String)
        params.Add(empresa)
        Dim formBusqueda As Form_BusquedaGeneral = General.buscarElemento(consulta, params, titulo, False)
        tbl = formBusqueda.rowResultados
        If tbl IsNot Nothing Then
            datatable.Rows(datagrid.CurrentCell.RowIndex).Item(2) = tbl(0)
            datatable.Rows(datagrid.CurrentCell.RowIndex).Item(3) = tbl(1)
            datatable.Rows(datagrid.CurrentCell.RowIndex).Item(4) = tbl(2)
            datagrid.Rows(datagrid.RowCount - 1).Cells(3).Selected = True
        End If
    End Sub
    Public Shared Sub agregarDiagnosticosUltimaEvo(ByVal datagrid As DataGridView, ByVal datatable As DataTable, ByVal Codigo As Object)
        Dim seguir As Boolean = False
        Dim seleccionados As String = ""
        Do
            seleccionados = getCadenaSeleccionados(datatable, seleccionados, 1)
            Dim tbl As DataRow = Nothing
            Dim params As New List(Of String)
            params.Add(Codigo)
            params.Add(seleccionados)
            Dim formBusqueda As Form_BusquedaGeneral = General.buscarElemento(Consultas.DIAGNOSTICOS_ULTIMA_EVO_LISTAR, params, TitulosForm.BUSQUEDA_DIAGNOSTICOS_ULTIMA_EVO, False)
            tbl = formBusqueda.rowResultados
            If tbl IsNot Nothing Then
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(0) = tbl(0)
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(1) = tbl(1)
                datatable.Rows.Add()
                datagrid.Rows(datagrid.RowCount - 1).Cells(1).Selected = True
                seguir = True
            Else
                seguir = False
            End If
        Loop While (seguir = True)
    End Sub

    Public Shared Sub agregarTraslados(ByVal datagrid As DataGridView, ByVal datatable As DataTable, ByVal eps As Object, ByVal registro As Integer)
        Dim objAdmision As New Admision
        Try
            Dim seguir As Boolean = False
            Dim seleccionados As String = ""
            Do
                seleccionados = getCadenaSeleccionados(datatable, seleccionados, 1)
                Dim tbl As DataRow = Nothing
                Dim params As New List(Of String)
                params.Add(eps)
                params.Add(registro)
                params.Add(seleccionados)
                Dim formBusqueda As Form_BusquedaGeneral = General.buscarElemento(Consultas.BUSQUEDA_PROCEDIMIENTO_TRASLADOS, params, TitulosForm.BUSQUEDA_TRASLADOS, False)
                tbl = formBusqueda.rowResultados
                If tbl IsNot Nothing Then
                    datatable.Rows(datagrid.CurrentCell.RowIndex).Item(1) = tbl(1)
                    datatable.Rows(datagrid.CurrentCell.RowIndex).Item(2) = tbl(2)
                    datatable.Rows(datagrid.CurrentCell.RowIndex).Item(3) = tbl(3)
                    datatable.Rows(datagrid.CurrentCell.RowIndex).Item(5) = tbl(0)
                    datatable.Rows.Add()
                    datagrid.Rows(datagrid.RowCount - 1).Cells(1).Selected = True
                    seguir = True
                Else
                    seguir = False
                End If
            Loop While (seguir = True)
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try

    End Sub

    Public Shared Sub agregarAreasDeServicios(ByVal datagrid As DataGridView, ByVal datatable As DataTable, ByVal codigoEp As Object)
        Try
            Dim seguir As Boolean = False
            Dim seleccionados As String = ""
            Do
                seleccionados = getCadenaSeleccionados(datatable, seleccionados, 0)
                Dim tbl As DataRow = Nothing
                Dim params As New List(Of String)
                params.Add(codigoEp)
                params.Add(Nothing)
                params.Add(Constantes.VALOR_PREDETERMINADO)
                params.Add(seleccionados)
                Dim formBusqueda As Form_BusquedaGeneral = General.buscarElemento(Consultas.AREA_SERVICIO_LISTAR_CONTRATO, params, TitulosForm.BUSQUEDA_AREA_SERVICIO, False)
                tbl = formBusqueda.rowResultados
                If tbl IsNot Nothing Then
                    datatable.Rows(datagrid.CurrentCell.RowIndex).Item("Codigo") = tbl(0)
                    datatable.Rows(datagrid.CurrentCell.RowIndex).Item("Descripcion") = tbl(1)
                    datatable.Rows(datagrid.CurrentCell.RowIndex).Item("Estado") = Constantes.ITEM_NUEVO
                    datatable.Rows.Add()
                    datagrid.Rows(datagrid.RowCount - 1).Cells("Descripcion").Selected = True
                    seguir = True
                Else
                    seguir = False
                End If
            Loop While (seguir = True)
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
    End Sub
    Public Shared Sub agregarItemsQx(ByVal consulta As String, ByVal titulo As String, ByVal datagrid As DataGridView, ByVal datatable As DataTable)
        Dim seguir As Boolean = False
        Dim seleccionados As String = ""
        Do
            seleccionados = getCadenaSeleccionados(datatable, seleccionados, 0)
            Dim tbl As DataRow = Nothing
            Dim params As New List(Of String)
            params.Add(seleccionados)
            Dim formBusqueda As Form_BusquedaGeneral = General.buscarElemento(consulta, params, titulo, False)
            tbl = formBusqueda.rowResultados
            If tbl IsNot Nothing Then
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(0) = tbl(0)
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(1) = tbl(1)
                datatable.Rows.Add()
                datagrid.Rows(datagrid.RowCount - 1).Cells(1).Selected = True
                seguir = True
            Else
                seguir = False
            End If
        Loop While (seguir = True)
    End Sub
    Public Shared Sub agregarItemsFormatIngreso(ByVal consulta As String, ByVal titulo As String, ByVal datagrid As DataGridView, ByVal datatable As DataTable)
        Dim seguir As Boolean = False
        Dim seleccionados As String = ""
        Do
            seleccionados = getCadenaSeleccionados(datatable, seleccionados, 0)
            Dim tbl As DataRow = Nothing
            Dim params As New List(Of String)
            params.Add(seleccionados)
            Dim formBusqueda As Form_BusquedaGeneral = General.buscarElemento(consulta, params, titulo, False)
            tbl = formBusqueda.rowResultados
            If tbl IsNot Nothing Then
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(0) = tbl(0)
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(1) = tbl(1)
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(2) = Constantes.ITEM_NUEVO
                datatable.Rows.Add()
                datagrid.Rows(datagrid.RowCount - 1).Cells(1).Selected = True
                seguir = True
            Else
                seguir = False
            End If
        Loop While (seguir = True)
    End Sub
    Public Shared Sub agregarItemsHojaRuta(ByVal consulta As String,
                                   ByVal titulo As String,
                                   ByVal datagrid As DataGridView,
                                   ByVal datatable As DataTable)
        Dim tbl As DataRow = Nothing
        Dim formBusqueda As Form_BusquedaGeneral = General.buscarElemento(consulta, Nothing, titulo, False)
        tbl = formBusqueda.rowResultados
        If tbl IsNot Nothing Then
            datatable.Rows(datagrid.CurrentCell.RowIndex).Item(3) = tbl(0)
            datatable.Rows(datagrid.CurrentCell.RowIndex).Item(4) = tbl(1)
        End If
    End Sub
    Public Shared Function getCadenaSeleccionados(datatable As DataTable, seleccionados As String, posicionCodigo As Integer) As String

        For I = 0 To datatable.Rows.Count - 2
            If I = 0 Then
                seleccionados = "" & datatable.Rows(I).Item(posicionCodigo).ToString & ""
            ElseIf I <> datatable.Rows.Count - 2 Then
                seleccionados = seleccionados & "'" & Constantes.COMODIN_DE_BUSQUEDA & "','" & Constantes.COMODIN_DE_BUSQUEDA & "'" & datatable.Rows(I).Item(posicionCodigo).ToString & ""
            Else
                seleccionados = seleccionados & "'" & Constantes.COMODIN_DE_BUSQUEDA & "','" & Constantes.COMODIN_DE_BUSQUEDA & "'" & datatable.Rows(I).Item(posicionCodigo).ToString
            End If
        Next
        If String.IsNullOrEmpty(seleccionados) Then
            seleccionados = Constantes.COMODIN_DE_BUSQUEDA
        End If
        Return seleccionados
    End Function

    Public Shared Sub llenarTablaYdgv(ByVal cadena As String, ByVal nmbreDT As Object)
        Try

            If TypeOf nmbreDT Is DataTable Then
                nmbreDT.Clear()
                Using da = New SqlDataAdapter(cadena, FormPrincipal.cnxion)
                    da.Fill(nmbreDT)
                End Using
            Else
                Dim dt As New DataTable
                dt.Clear()
                Using da = New SqlDataAdapter(cadena, FormPrincipal.cnxion)
                    da.Fill(dt)
                End Using
                nmbreDT.DataSource = dt
                nmbreDT.DefaultCellStyle.Font = New Font(Constantes.TIPO_LETRA_ELEMENTO, 9)
            End If

        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
    End Sub

    Public Shared Sub llenarTablaConExcepcion(ByVal consulta As String,
                                  ByVal params As List(Of String),
                                  ByVal dtTabla As DataTable,
                                  Optional pLimpiarDT As Boolean = True)

        Dim listaParams As String = Funciones.getParametros(params)

        Try
            If pLimpiarDT Then dtTabla.Clear()
            Using daAdapter = New SqlDataAdapter(consulta & listaParams, FormPrincipal.cnxion)
                daAdapter.SelectCommand.CommandTimeout = 120
                daAdapter.Fill(dtTabla)
            End Using

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Shared Sub llenarTabla(ByVal consulta As String,
                                  ByVal params As List(Of String),
                                  ByVal dtTabla As DataTable,
                                  Optional pLimpiarDT As Boolean = True)

        Dim listaParams As String = Funciones.getParametros(params)

        Try
            If pLimpiarDT Then dtTabla.Clear()
            Using daAdapter = New SqlDataAdapter(consulta & listaParams, FormPrincipal.cnxion)
                daAdapter.SelectCommand.CommandTimeout = 120
                daAdapter.Fill(dtTabla)
            End Using

        Catch ex As Exception
            General.manejoErrores(ex)
        End Try

    End Sub
    Public Shared Sub llenarTablaParametroTabla(ByVal consulta As String,
                                                ByVal params As Hashtable,
                                                ByVal tblParametro As DataTable,
                                                ByVal dtTabla As DataTable,
                                                Optional pLimpiarDT As Boolean = True)
        Try
            If pLimpiarDT Then dtTabla.Clear()
            Using Query As New SqlCommand
                Query.CommandType = CommandType.StoredProcedure
                Query.CommandText = consulta
                For Each elemento As DictionaryEntry In params
                    Query.Parameters.Add(elemento.Key, SqlDbType.NVarChar).Value = elemento.Value
                Next

                Query.Parameters.Add("@pTabla", SqlDbType.Structured).Value = tblParametro
                Query.CommandTimeout = 120
                Query.Connection = FormPrincipal.cnxion

                Using daAdapter = New SqlDataAdapter(Query)
                    daAdapter.Fill(dtTabla)
                End Using
            End Using


        Catch ex As Exception
            General.manejoErrores(ex)

        End Try

    End Sub

    Public Shared Sub llenarTabla2doPlano(ByVal consulta As String,
                                          ByVal params As List(Of String),
                                          ByVal dtTabla As DataTable,
                                          Optional pLimpiarDT As Boolean = True)

        Dim listaParams As String = Funciones.getParametros(params)

        Try
            If pLimpiarDT Then dtTabla.Clear()
            Using daAdapter = New SqlDataAdapter(consulta & listaParams, FormPrincipal.cnxion2doPlano)
                daAdapter.Fill(dtTabla)
            End Using
            FormPrincipal.cnxion2doPlano.Close()

        Catch ex As Exception
            General.manejoErrores(ex)
        End Try

    End Sub
    Public Shared Sub llenarTabla2doPlanoConExcepcion(ByVal consulta As String,
                                          ByVal params As List(Of String),
                                          ByVal dtTabla As DataTable,
                                          Optional pLimpiarDT As Boolean = True)

        Dim listaParams As String = Funciones.getParametros(params)

        Try
            If pLimpiarDT Then dtTabla.Clear()
            Using daAdapter = New SqlDataAdapter(consulta & listaParams, FormPrincipal.cnxion2doPlano)
                daAdapter.Fill(dtTabla)
            End Using
            FormPrincipal.cnxion2doPlano.Close()

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Shared Sub manejoErrores(ex As Object)
        Select Case True
            Case TypeOf ex Is ArgumentNullException
                mensajeCritico(ConstantesError.ARGUMENTO_NULL & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is ArgumentOutOfRangeException
                mensajeCritico(ConstantesError.ARGUMENTO_NO_VALIDO & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is IndexOutOfRangeException
                mensajeCritico(ConstantesError.INDICE_NO_VALIDO & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is NullReferenceException
                mensajeCritico(ConstantesError.DESREFERENCIAR & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is DivideByZeroException
                mensajeCritico(ConstantesError.DIVISION_ENTRE_0 & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is StackOverflowException
                mensajeCritico(ConstantesError.LLAMADO_EXCESIVO & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is OverflowException
                mensajeCritico(ConstantesError.DESBORDAMIENTO & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is FileNotFoundException
                mensajeCritico(ConstantesError.DIRECTORIO_NO_ENCONTRADO & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is EndOfStreamException
                mensajeCritico(ConstantesError.LECTURA_SOBREPASADA & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is DirectoryNotFoundException
                mensajeCritico(ConstantesError.DIRECTORIO_NO_ENCONTRADO & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is CrystalReportsException
                mensajeCritico(ConstantesError.REPORTE & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is InternalException
                mensajeCritico(ConstantesError.REPORTE & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is DataSourceException
                mensajeCritico(ConstantesError.ERROR_SIN_CONEXION & ConstantesError.COMUNICAR_SISTEMA, ex.message)
            Case TypeOf ex Is SqlException
                manejoErroresSql(ex.number, ex.message)
            Case Else
                Dim a As String
                'If TypeOf ex Is SqlException Then
                '    a = ex.GetType().FullName & "Error n°: " & ex.number
                'Else
                a = ex.GetType().FullName
                'End If
                Dim b = ex.data.ToString
                mensajeCritico(ex.message & ConstantesError.COMUNICAR_SISTEMA, a & vbCrLf & b & vbCrLf & ex.message)
        End Select

    End Sub
    Public Shared Sub mensajeCritico(msg As String, errorMsg As String, Optional enviar As Boolean = True)
        Dim rutaCarpeta, rutaArchivo As String
        rutaCarpeta = Path.GetTempPath
        rutaArchivo = rutaCarpeta & "\Incidencia_" & Format(Now, "HH'h'mm'm'ss's'") & ConstantesHC.EXTENSION_ARCHIVO_IMAGEN_PNG
        MsgBox(msg, MsgBoxStyle.Critical)
        If enviar Then
            Using pantallazo As Image = capturarImagen()
                pantallazo.Save(rutaArchivo, Imaging.ImageFormat.Png)
            End Using

            Dim correoConfigurado As New Correo
            Dim listaCorreos As New List(Of Correo)
            Try
                listaCorreos = General.cargarInformacionCorreo(Constantes.CODIGO_CORREO_ERROR)
                correoConfigurado = listaCorreos.First()
                correoConfigurado.asunto = "Incidencia en PC: " & Environment.MachineName
                correoConfigurado.cuerpo = errorMsg
                correoConfigurado.correoDestinatario = Constantes.CORREO_SISTEMAS
                Funciones.enviarCorreo(correoConfigurado, rutaArchivo)
            Catch ex As Exception

            End Try

        End If

    End Sub

    Public Shared Function capturarImagen() As Image

        ' Tamaño de lo que queremos copiar
        ' En este caso el tamaño de la ventana principal
        Dim fSize As Size = Screen.PrimaryScreen.Bounds.Size
        ' Creamos el bitmap con el área que vamos a capturar
        Dim bm As New Bitmap(fSize.Width, fSize.Height)
        ' Un objeto Graphics a partir del bitmap
        Dim gr As Graphics = Graphics.FromImage(bm)
        ' Copiar todo el área de la pantalla
        gr.CopyFromScreen(0, 0, 0, 0, fSize)

        Return bm
    End Function

    Public Shared Sub manejoErroresSql(errorNum As Double, errorMsg As String)
        Dim mensaje As String
        mensaje = SesionActual.mostrarError(errorNum)
        If String.IsNullOrEmpty(mensaje) Then
            mensaje = ConstantesError.ERROR_GENERAL
        End If
        Select Case errorNum
            Case ConstantesError.CODIGO_ERROR_CONEXION,
                 ConstantesError.CODIGO_ERROR_CONEXION2,
                 ConstantesError.CODIGO_ERROR_TIEMPO_ESPERA,
                 ConstantesError.CODIGO_ERROR_SIN_CONEXION,
                 ConstantesError.CODIGO_ERROR_SIN_CONEXION2,
                 ConstantesError.CODIGO_ERROR_SIN_CONEXION3,
                 ConstantesError.CODIGO_ERROR_SIN_CONEXION4,
                 ConstantesError.CODIGO_ERROR_SIN_CONEXION5
                mensajeCritico(mensaje, errorMsg, False)
            Case ConstantesError.CODIGO_ERROR_INTERBLOQUEO2,
                 ConstantesError.CODIGO_ERROR_INTERBLOQUEO

            Case Else
                mensajeCritico(errorMsg, "Código error: " & errorNum & ". " & errorMsg)
        End Select

    End Sub
    Public Shared Function getCadenaSplit(cadena As String, divisor As String, posicion As Integer) As String
        Dim codigos As String()
        codigos = Split(cadena, divisor)
        Return codigos(posicion)
    End Function

    Public Shared Function llenarTabla(ByVal consulta As String,
                                       ByVal params As List(Of SqlParameter)) As DataTable

        Dim dtTabla As New DataTable

        Try
            Using dbCommand As New SqlCommand
                dbCommand.Connection = FormPrincipal.cnxion
                dbCommand.CommandType = CommandType.StoredProcedure
                dbCommand.CommandText = consulta

                For Each param As SqlParameter In params

                    If param.Value = Nothing Then
                        param.Value = DBNull.Value
                    End If

                    dbCommand.Parameters.Add(param)
                Next

                Using daAdapter = New SqlDataAdapter(dbCommand)
                    daAdapter.Fill(dtTabla)
                End Using


            End Using

        Catch ex As Exception
            Throw
        End Try

        Return dtTabla

    End Function

    Public Shared Sub concatenacionComas(ByRef dt As DataTable, ByRef csv As String)

        For Each dw As DataRow In dt.Rows
            Dim dw_csv As String = ""
            For Each i In dw.ItemArray
                If dw_csv <> "" Then
                    dw_csv += ","
                End If
                If i.ToString.Contains("vbCrLf") Then
                    i = Replace(i, vbCrLf, "")
                End If
                dw_csv += i.ToString().ToUpper()
            Next
            If csv <> "" Then
                csv += vbCrLf
            End If
            csv += dw_csv
        Next
        dt.Clear()
        dt.Columns.Clear()
    End Sub

    Public Shared Function cargarItem(ByVal consulta As String,
                                      ByVal params As List(Of String)) As DataRow

        Dim listaParams As String = Funciones.getParametros(params)
        Dim dtTabla As New DataTable

        Try
            Using daAdapter = New SqlDataAdapter(consulta & listaParams, FormPrincipal.cnxion)
                daAdapter.Fill(dtTabla)
            End Using

        Catch ex As Exception
            Throw ex
        End Try

        If dtTabla.Rows.Count > 0 Then
            Return dtTabla.Rows(0)
        Else
            Return Nothing
        End If


    End Function
    Public Shared Function cargarItem(ByVal consulta As String) As DataRow
        Dim dtTabla As New DataTable

        Try
            Using daAdapter = New SqlDataAdapter(consulta, FormPrincipal.cnxion)
                daAdapter.Fill(dtTabla)
            End Using

        Catch ex As Exception
            Throw ex
        End Try

        If dtTabla.Rows.Count > 0 Then
            Return dtTabla.Rows(0)
        Else
            Return Nothing
        End If


    End Function



    Public Shared Sub llenardgv(ByVal consulta As String,
                                   ByVal dgdgv As DataGridView,
                                   ByVal params As List(Of String))

        Dim listaParams As String = Funciones.getParametros(params)
        Dim dtTabla As New DataTable

        Try
            Using daAdapter = New SqlDataAdapter(consulta & listaParams, FormPrincipal.cnxion)
                daAdapter.Fill(dtTabla)
            End Using
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try

        dgdgv.DataSource = dtTabla
        dgdgv.DefaultCellStyle.Font = New Font(Constantes.TIPO_LETRA_ELEMENTO, 9)

    End Sub

    Public Shared Sub llenarDataSet(ByVal cadena As String, ByVal objDs As Object,
       Optional ByVal eliminartablas As Boolean = False)
        Try

            If TypeOf objDs Is DataSet Then
                If objDs Is Nothing Then
                    objDs = New DataSet()
                ElseIf objDs.Tables IsNot Nothing Then

                    If eliminartablas Then
                        objDs.Tables.Clear()
                    Else
                        For Each dt In objDs.Tables
                            dt.Clear()
                        Next
                    End If
                End If
                Using da = New SqlDataAdapter(cadena, FormPrincipal.cnxion)
                    da.Fill(objDs)
                End Using
            End If

        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
    End Sub
    Public Shared Sub llenarDataSet(ByVal cadena As String, params As List(Of String), ByVal objDs As DataSet)
        Try
            Dim listaParams As String = Funciones.getParametros(params)
            For Each dt In objDs.Tables
                dt.Clear()
            Next
            Using da = New SqlDataAdapter(cadena & listaParams, FormPrincipal.cnxion)
                da.Fill(objDs)
            End Using
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
    End Sub

    Public Shared Function anularRegistro(ByVal sentencia As String) As Boolean
        Try
            Using dbCommand = New SqlCommand(sentencia, FormPrincipal.cnxion)
                dbCommand.ExecuteNonQuery()
                Return True
            End Using
        Catch ex As Exception
            General.manejoErrores(ex)
            Return False
        End Try
    End Function
    Public Shared Function anularConValidacion(ByVal cadena As String) As Boolean
        Dim respuesta As Integer
        Dim bdra As Boolean
        Try
            Using consulta = New SqlCommand(cadena, FormPrincipal.cnxion)
                respuesta = consulta.ExecuteScalar
                If respuesta > 0 Then
                    bdra = True
                End If
                Return bdra
            End Using
        Catch ex As Exception
            General.manejoErrores(ex)
            Return False
        End Try
    End Function
    Public Shared Sub ejecutarSQL(ByVal nombreProc As String, params As List(Of String))
        Dim listaParams As String = Funciones.getParametros(params)

        Try
            Using dbCommand = New SqlCommand(nombreProc & listaParams, FormPrincipal.cnxion)
                dbCommand.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Shared Function getEstadoVF(ByVal cadena As String) As Boolean
        Dim respuesta As Boolean
        Try
            Using consulta As New SqlCommand(cadena)
                consulta.Connection = FormPrincipal.cnxion
                respuesta = consulta.ExecuteScalar()
            End Using
            Return respuesta
        Catch ex As Exception
            General.manejoErrores(ex)
            Return False
        End Try
    End Function
    Public Shared Function getEstadoVF(pConsultaSQL As String,
                                       plistaParam As List(Of String)) As Boolean
        Dim respuesta As Boolean
        If Not IsNothing(plistaParam) Then
            pConsultaSQL = pConsultaSQL & Funciones.getParametros(plistaParam)
        End If
        Try
            Using consulta As New SqlCommand(pConsultaSQL)
                consulta.Connection = FormPrincipal.cnxion
                respuesta = consulta.ExecuteScalar()
            End Using
            Return respuesta
        Catch ex As Exception
            General.manejoErrores(ex)
            Return False
        End Try
    End Function

    Public Shared Function getValorConsulta(pConsultaSQL As String,
                                     plistaParam As List(Of String)) As String
        Dim respuesta As String
        If Not IsNothing(plistaParam) Then
            pConsultaSQL = pConsultaSQL & Funciones.getParametros(plistaParam)
        End If
        Try
            Using consulta As New SqlCommand(pConsultaSQL)
                consulta.Connection = FormPrincipal.cnxion
                respuesta = consulta.ExecuteScalar().ToString
            End Using
            Return respuesta
        Catch ex As Exception
            General.manejoErrores(ex)
            Return ""
        End Try
    End Function
    Public Shared Sub buscarElemento(pConsultaSQL As String,
                                     plistaParam As List(Of String),
                                     pMetodo As cargaInfoForm,
                                     pTitulo As String,
                                     pOcultaCol As Boolean,
                                     Optional pTamanoForm As Integer = 0,
                                     Optional pBuscarDarEnter As Boolean = False
                                     )
        Dim vForm As New Form_BusquedaGeneral()
        vForm.Text = pTitulo
        If Not IsNothing(plistaParam) Then
            vForm.consulta = pConsultaSQL & Funciones.getParametros(plistaParam)
        Else
            vForm.consulta = pConsultaSQL
        End If
        vForm.metodoCarga = pMetodo
        vForm.isOcultaCol = pOcultaCol
        vForm.buscarAlDarEnter = pBuscarDarEnter
        Select Case pTamanoForm
            Case Constantes.TAMANO_MEDIANO
                vForm.Size = New Size(vForm.Width + 300, vForm.Height)

            Case Constantes.TAMANO_GRANDE
                vForm.WindowState = FormWindowState.Maximized
        End Select
        vForm.ShowDialog()
    End Sub

    Public Shared Function buscarElemento(pConsultaSQL As String,
                                          plistaParam As List(Of String),
                                          pTitulo As String,
                                          pOcultaCol As Boolean,
                                          Optional pTamanoForm As Integer = 0,
                                          Optional pBuscarDarEnter As Boolean = False)

        Dim vForm As New Form_BusquedaGeneral()
        vForm.Text = pTitulo
        If Not IsNothing(plistaParam) Then
            vForm.consulta = pConsultaSQL & Funciones.getParametros(plistaParam)
        Else
            vForm.consulta = pConsultaSQL
        End If
        vForm.isOcultaCol = pOcultaCol
        vForm.buscarAlDarEnter = pBuscarDarEnter
        Select Case pTamanoForm
            Case Constantes.TAMANO_MEDIANO
                vForm.Size = New Size(vForm.Width + 350, vForm.Height)

            Case Constantes.TAMANO_GRANDE
                vForm.WindowState = FormWindowState.Maximized
        End Select
        vForm.ShowDialog()
        Return vForm

    End Function

    ''USAR ESTE METODO PARA BUSQUEDAS BASICAS, SIN QUE SE DEBAN ESCONDER OTRAS COLUMNAS DIFERENTES A LA DEL CODIGO.

    Public Shared Sub buscarItem(pConsultaSQL As String,
                                     plistaParam As List(Of String),
                                     pMetodo As cargaInfoFormObj,
                                     pTitulo As String,
                                     pOcultaCol As Boolean,
                                     Optional pBuscarDarEnter As Boolean = False,
                                     Optional pTamanoForm As Integer = 0
                                     )

        Dim vForm As New Form_BusquedaGeneral()

        vForm.Text = pTitulo
        vForm.isRetornaObj = True
        If Not IsNothing(plistaParam) Then
            vForm.consulta = pConsultaSQL & Funciones.getParametros(plistaParam)
        Else
            vForm.consulta = pConsultaSQL
        End If
        vForm.buscarAlDarEnter = pBuscarDarEnter
        vForm.metodoCargaObj = pMetodo
        vForm.isOcultaCol = pOcultaCol
        Select Case pTamanoForm
            Case Constantes.TAMANO_MEDIANO
                vForm.Size = New Size(vForm.Width + 350, vForm.Height)

            Case Constantes.TAMANO_GRANDE
                vForm.WindowState = FormWindowState.Maximized
        End Select
        vForm.ShowDialog()

    End Sub


    ' limpia control del formulario
    Public Shared Sub limpiarControles(ByRef pFormulario As Object)
        Dim vFrtRB As Integer = pFormulario.Width + pFormulario.Height
        If (TypeOf pFormulario Is Form) Then
            aplicarIcono(pFormulario)
        End If
        For Each vControl In pFormulario.Controls

            If (TypeOf vControl Is TextBox) OrElse (TypeOf vControl Is RichTextBox) Then
                vControl.Clear()
            ElseIf (TypeOf vControl Is Label) AndAlso vControl.text.contains(Constantes.LINEA_TITULO) Then
                vControl.forecolor = Color.LimeGreen
            ElseIf (TypeOf vControl Is MaskedTextBox) Then
                vControl.ResetText()
            ElseIf (TypeOf vControl Is ComboBox) Then
                If (vControl.Items.Count > 0) Then vControl.SelectedIndex = 0 _
                    Else vControl.Text = ""
            ElseIf (TypeOf vControl Is DateTimePicker) Then
                vControl.Value = Funciones.Fecha(Constantes.FORMATO_FECHA_HORA_NUMERICA)
            ElseIf (TypeOf vControl Is NumericUpDown) Then
                vControl.Value = vControl.Minimum
            ElseIf (TypeOf vControl Is CheckBox) Then
                vControl.Checked = False
            ElseIf (TypeOf vControl Is PictureBox) And (vControl.name <> Constantes.FORMULARIO_LOGO) Then
                vControl.Image = Nothing
            ElseIf (TypeOf vControl Is CheckedListBox) Then
                vControl.Items.Clear()
            ElseIf (TypeOf vControl Is RadioButton) Then
                Dim vCurrentRB = vControl.Location.X + vControl.Location.Y
                If (vFrtRB > vCurrentRB) Then
                    vFrtRB = vCurrentRB
                Else
                    vControl.Checked = False
                End If
            ElseIf (TypeOf vControl Is TreeView) Then
                vControl.Nodes.Clear()
            ElseIf (TypeOf vControl Is DataGridView) Then
                vControl.EndEdit()
                If TypeOf vControl.DataSource Is BindingSource Then
                    vControl.DataSource.DataSource.Clear()
                Else
                    If vControl.Datasource IsNot Nothing Then
                        vControl.Datasource.Clear()
                    End If
                End If
                personalizarDatagrid(vControl)
            Else
                ' mira a ve si es un contenedor
                limpiarControles(vControl)
            End If
        Next

    End Sub

    ' limpia control del formulario


    Private Shared Sub aplicarIcono(pFormulario As Object)
        ''Dim myBitmap As New Bitmap(My.Resources.CHFinalIcono1)
        ''Dim HIcon As IntPtr = myBitmap.GetHicon()
        ''Dim newIcon As Icon = Icon.FromHandle(HIcon)
        'pFormulario.icon = New Icon(newIcon, New Size(48, 48))
    End Sub

    Public Shared Sub personalizarDatagrid(ByRef dgv As DataGridView)
        dgv.BackgroundColor = Color.White
        dgv.DefaultCellStyle.BackColor = Color.White
        dgv.DefaultCellStyle.ForeColor = Color.Black
        dgv.DefaultCellStyle.SelectionBackColor = Color.LightSteelBlue
        dgv.DefaultCellStyle.SelectionForeColor = Color.Black
        dgv.DefaultCellStyle.Font = New Font(Constantes.TIPO_LETRA_ELEMENTO, 9)
        dgv.EnableHeadersVisualStyles = False
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.AlternatingRowsDefaultCellStyle.Font = New Font(Constantes.TIPO_LETRA_ELEMENTO, 9)
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font(Constantes.TIPO_LETRA_ELEMENTO, 9)
        dgv.RowHeadersDefaultCellStyle.SelectionBackColor = Color.LightSteelBlue
        For indiceColumna = 0 To dgv.Columns.Count - 1
            dgv.Columns(indiceColumna).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
        AddHandler dgv.CellMouseEnter, AddressOf resaltarFila
        AddHandler dgv.CellMouseLeave, AddressOf quitarResaltadoFila
    End Sub
    Public Shared Sub resaltarFila(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then
            sender.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.FromArgb(226, 234, 243)
        End If
    End Sub
    Public Shared Sub quitarResaltadoFila(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then
            sender.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
        End If
    End Sub

    Public Shared Sub formNuevo(ByRef pForm As Object,
                                ByRef pToolStrip As Windows.Forms.ToolStrip,
                                ByRef pBoton1 As ToolStripButton,
                                ByRef pBoton2 As ToolStripButton)

        Try
            General.limpiarControles(pForm)
            General.habilitarControles(pForm)
            General.deshabilitarBotones(pToolStrip)

            'Parametros pBoton1 y pBoton2 son opcionales, se valida si vienen vacios
            If pBoton1 IsNot Nothing Then
                pBoton1.Enabled = True
            End If

            If pBoton2 IsNot Nothing Then
                pBoton2.Enabled = True
            End If
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try

    End Sub

    'Metodo general para las acciones del boton cancelar
    Public Shared Function formCancelar(ByRef pForm As Object,
                                        ByRef pToolStrip As Windows.Forms.ToolStrip,
                                              pBoton1 As Windows.Forms.ToolStripButton,
                                              pBoton2 As Windows.Forms.ToolStripButton) As Boolean
        If MsgBox(Mensajes.CANCELAR, MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Return False
            Exit Function
        End If

        Try
            limpiarControles(pForm)
            deshabilitarBotones(pToolStrip)
            deshabilitarControles(pForm)

            'Parametros pBoton1 y pBoton2 son opcionales, se valida si vienen vacios
            If pBoton1 IsNot Nothing Then
                pBoton1.Enabled = True
            End If

            If pBoton2 IsNot Nothing Then
                pBoton2.Enabled = True
            End If
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
        Return True
    End Function
    'Metodo general para las acciones del boton editar
    Public Shared Sub formEditar(ByRef pForm As Object,
                                 ByRef pToolStrip As Windows.Forms.ToolStrip,
                                 pBoton1 As Windows.Forms.ToolStripButton,
                                 pBoton2 As Windows.Forms.ToolStripButton)

        If MsgBox(Mensajes.EDITAR, MsgBoxStyle.Question + MsgBoxStyle.YesNo, TitulosForm.EDITAR) = MsgBoxResult.No Then
            Exit Sub
        End If

        Try
            deshabilitarBotones(pToolStrip)
            habilitarControles(pForm)

            'Parametros pBoton1 y pBoton2 son opcionales, se valida si vienen vacios
            If pBoton1 IsNot Nothing Then
                pBoton1.Enabled = True
            End If

            If pBoton2 IsNot Nothing Then
                pBoton2.Enabled = True
            End If

        Catch ex As Exception
            General.manejoErrores(ex)
            Throw ex
        End Try

    End Sub
    'Metodo general para las acciones del boton Duplicar
    Public Shared Sub formDuplicar(ByRef pForm As Object,
                                 ByRef pToolStrip As Windows.Forms.ToolStrip,
                                 pBoton1 As Windows.Forms.ToolStripButton,
                                 pBoton2 As Windows.Forms.ToolStripButton)
        Try
            deshabilitarBotones(pToolStrip)
            habilitarControles(pForm)

            'Parametros pBoton1 y pBoton2 son opcionales, se valida si vienen vacios
            If pBoton1 IsNot Nothing Then
                pBoton1.Enabled = True
            End If

            If pBoton2 IsNot Nothing Then
                pBoton2.Enabled = True
            End If

        Catch ex As Exception
            General.manejoErrores(ex)
            Throw ex
        End Try

    End Sub
    Public Shared Function fnFormEditar(ByRef pForm As Object,
                                     ByRef pToolStrip As Windows.Forms.ToolStrip,
                                           pBoton1 As Windows.Forms.ToolStripButton,
                                           pBoton2 As Windows.Forms.ToolStripButton) As Boolean

        If MsgBox(Mensajes.EDITAR, MsgBoxStyle.Question + MsgBoxStyle.YesNo, TitulosForm.EDITAR) = MsgBoxResult.No Then
            Return False
        End If

        Try
            deshabilitarBotones(pToolStrip)
            habilitarControles(pForm)

            'Parametros pBoton1 y pBoton2 son opcionales, se valida si vienen vacios
            If pBoton1 IsNot Nothing Then
                pBoton1.Enabled = True
            End If

            If pBoton2 IsNot Nothing Then
                pBoton2.Enabled = True
            End If

        Catch ex As Exception
            General.manejoErrores(ex)
            Throw ex
        End Try
        Return True
    End Function
    Public Shared Function verificar_formulario(form As Windows.Forms.Form) As Boolean
        Dim valor As Boolean = False
        For Each f As Windows.Forms.Form In Application.OpenForms
            If f.Name = form.Name Then
                valor = True
            End If
        Next
        Return valor
    End Function
    Public Shared Sub cargarForm(ByVal form As Windows.Forms.Form)
        FormPrincipal.Cursor = Cursors.WaitCursor
        form.MdiParent = FormPrincipal
        General.limpiarControles(form)
        form.Show()
        form.Focus()
        If form.WindowState = FormWindowState.Minimized Then
            form.WindowState = FormWindowState.Normal
        End If
        FormPrincipal.Cursor = Cursors.Default
    End Sub

    'Metodo general para configurar controles luego de cargar un formulario
    Public Shared Sub posLoadForm(ByRef pForm As Object,
                                  ByRef pToolStrip As Windows.Forms.ToolStrip,
                                  pBoton1 As Windows.Forms.ToolStripButton,
                                  pBoton2 As Windows.Forms.ToolStripButton)

        General.deshabilitarControles(pForm)
        General.deshabilitarBotones(pToolStrip)

        If pBoton1 IsNot Nothing Then
            pBoton1.Enabled = True
        End If

        If pBoton2 IsNot Nothing Then
            pBoton2.Enabled = True
        End If

    End Sub

    'Metodo general para configurar controles luego de elegir un item desde 
    'el formulario de busqueda
    Public Shared Sub posBuscarForm(ByRef pForm As Object,
                                    ByRef pToolStrip As Windows.Forms.ToolStrip,
                                    pBoton1 As Windows.Forms.ToolStripButton,
                                    pBoton2 As Windows.Forms.ToolStripButton,
                                    pBoton3 As Windows.Forms.ToolStripButton,
                                    pBoton4 As Windows.Forms.ToolStripButton)
        General.deshabilitarControles(pForm)
        General.deshabilitarBotones(pToolStrip)

        If pBoton1 IsNot Nothing Then
            pBoton1.Enabled = True
        End If

        If pBoton2 IsNot Nothing Then
            pBoton2.Enabled = True
        End If

        If pBoton3 IsNot Nothing Then
            pBoton3.Enabled = True
        End If

        If pBoton4 IsNot Nothing Then
            pBoton4.Enabled = True
        End If

    End Sub

    Public Shared Sub posGuardarForm(ByRef pForm As Object,
                                     ByRef pToolStrip As Windows.Forms.ToolStrip,
                                     pBoton1 As Windows.Forms.ToolStripButton,
                                     pBoton2 As Windows.Forms.ToolStripButton,
                                     pBoton3 As Windows.Forms.ToolStripButton,
                                     pBoton4 As Windows.Forms.ToolStripButton,
                                     Optional pBoton5 As Windows.Forms.ToolStripButton = Nothing)

        General.deshabilitarControles(pForm)
        General.deshabilitarBotones(pToolStrip)

        If pBoton1 IsNot Nothing Then
            pBoton1.Enabled = True
        End If
        If pBoton2 IsNot Nothing Then
            pBoton2.Enabled = True
        End If
        If pBoton3 IsNot Nothing Then
            pBoton3.Enabled = True
        End If
        If pBoton4 IsNot Nothing Then
            pBoton4.Enabled = True
        End If
        If pBoton5 IsNot Nothing Then
            pBoton4.Enabled = True
        End If

        MsgBox(Mensajes.GUARDADO, MsgBoxStyle.Information)
    End Sub

    Public Shared Sub posAnularForm(ByRef pForm As Object,
                                    ByRef pToolStrip As Windows.Forms.ToolStrip,
                                    pBoton1 As Windows.Forms.ToolStripButton,
                                    pBoton2 As Windows.Forms.ToolStripButton)

        General.limpiarControles(pForm)
        General.deshabilitarControles(pForm)
        General.deshabilitarBotones(pToolStrip)

        If pBoton1 IsNot Nothing Then
            pBoton1.Enabled = True
        End If

        If pBoton2 IsNot Nothing Then
            pBoton2.Enabled = True
        End If
        MsgBox(Mensajes.ANULADO, MsgBoxStyle.Information)
    End Sub

    Public Shared Sub calculosProcesoInventario(ByRef dt As DataTable,
                                                ByVal columnaCantidad As String,
                                                ByVal columnaPrecio As String,
                                                ByVal columnaTotal As String,
                                                ByVal columnaIva As String,
                                                ByVal columnaValorIva As String,
                                                ByVal columnaDescuento As String,
                                                ByVal columnaValorDescuento As String,
                                                ByVal tipo_calculo As Boolean)

        For i = 0 To If(tipo_calculo = True, dt.Rows.Count - 1, dt.Rows.Count - 2)
            dt.Rows(i).Item(columnaTotal) = dt.Rows(i).Item(columnaCantidad) * dt.Rows(i).Item(columnaPrecio)
            dt.Rows(i).Item(columnaValorIva) = dt.Rows(i).Item(columnaTotal) * (dt.Rows(i).Item(columnaIva) / 100)
            dt.Rows(i).Item(columnaValorDescuento) = dt.Rows(i).Item(columnaTotal) * (dt.Rows(i).Item(columnaDescuento) / 100)
        Next
    End Sub
    Public Shared Function fechaActualServidor() As DateTime
        Try
            Dim fechaServidor As DateTime
            Dim dr As DataRow
            dr = cargarItem(Consultas.CONSULTAR_FECHA_SERVIDOR, Nothing)
            fechaServidor = dr(0)
            Return fechaServidor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

    End Function
    Public Shared Sub enumerarFilas(ByRef tabla As DataTable,
                                    ByVal nombreColumna As String)
        For indiceFila = 0 To tabla.Rows.Count - 1
            tabla.Rows(indiceFila).Item(nombreColumna) = (indiceFila + 1)
        Next
    End Sub
    'Public Shared Sub generarSecuencial(dgv As DataGridView)
    '    If dgv IsNot Nothing Then
    '        For Each r As DataGridViewRow In dgv.Rows
    '            r.HeaderCell.Value = (r.Index + 1).ToString()
    '        Next
    '    End If

    'End Sub
    Public Shared Sub generarSecuencial(dgv As DataGridView, e As DataGridViewRowPostPaintEventArgs)

        Using Pinceles As New SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString, dgv.Rows(e.RowIndex).Cells(0).InheritedStyle.Font, Pinceles, e.RowBounds.Location.X + 12, e.RowBounds.Location.Y + 2)
        End Using
    End Sub
    Public Shared Sub IsExistente(ByRef dgv As DataGridView,
                                  ByVal nombreColumna As String,
                                  ByVal fila As Integer)
        dgv.Rows(fila).Cells(nombreColumna).Value = My.Resources.OK
        dgv.Rows(fila).Cells(nombreColumna).Style.BackColor = Color.LightGreen
    End Sub
    Public Shared Sub IsNotExistente(ByRef dgv As DataGridView,
                                     ByVal nombreColumna As String,
                                     ByVal fila As Integer)
        dgv.Rows(fila).Cells(nombreColumna).Value = My.Resources._new
        dgv.Rows(fila).Cells(nombreColumna).Style.BackColor = Color.White
    End Sub
#Region "Inventario"
    Public Shared Sub quitarTabla(ByRef dts As DataSet, ByVal nombreTabla As String)
        If MsgBox(Mensajes.QUITAR, MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.Yes AndAlso dts.Tables.Contains(nombreTabla) = True Then
            dts.Tables.Remove(nombreTabla)
        End If
    End Sub
    Public Shared Sub cantidadValidaCelda(ByRef dgv As DataGridView,
                                          ByVal e As DataGridViewCellEventArgs)
        If Funciones.filaValida(e.RowIndex) Then
            If dgv.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString = "" Then
                dgv.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
            End If
        End If

    End Sub
#End Region
    Public Shared Function ExportarDataGridViewADataTable(ByVal dataGrid As DataGridView, ByRef dt As DataTable)
        Try
            Dim filaNueva As System.Data.DataRow
            Dim numCols As Integer

            numCols = dataGrid.ColumnCount

            For Each filaDatos As DataGridViewRow In dataGrid.Rows
                filaNueva = dt.NewRow()
                For i As Integer = 0 To numCols - 1
                    filaNueva(i) = filaDatos.Cells(i).Value
                Next
                dt.Rows.Add(filaNueva)
            Next

        Catch ex As Exception
            dt = New DataTable
        End Try
        Return dt
    End Function
    Public Shared Sub agregarDiagnosticosFormatoIngreso(ByVal datagrid As Object,
                                                        ByVal datatable As Object,
                                                        codigoArea As Integer,
                                                        codigoGenero As Integer)
        Dim seguir As Boolean = False
        Dim seleccionados As String = ""
        Do
            Dim tbl As DataRow = Nothing
            Dim params As New List(Of String)
            seleccionados = getCadenaSeleccionados(datatable, seleccionados, 0)
            params.Add(codigoArea)
            params.Add(codigoGenero)
            params.Add(String.Empty)
            params.Add(seleccionados)
            Dim formBusqueda As Form_BusquedaGeneral = General.buscarElemento(Consultas.CONSULTAR_DIAGNOSTICO_FORMATO, params, TitulosForm.BUSQUEDA_DIAGNOSTICO_CIE, False)
            tbl = formBusqueda.rowResultados
            If tbl IsNot Nothing Then
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(0) = tbl(0)
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(1) = tbl(1)
                datatable.Rows(datagrid.CurrentCell.RowIndex).Item(2) = Constantes.ITEM_NUEVO
                datatable.Rows.Add()
                datagrid.Rows(datagrid.RowCount - 1).Cells(1).Selected = True
                seguir = True
            Else
                seguir = False
            End If
        Loop While (seguir = True)
    End Sub

    Public Shared Sub formatDatePicker(dtTimePicker As DateTimePicker, fecha As Date)
        If fecha = Nothing Then
            'dtTimePicker.CustomFormat = Constantes.COMBO_FORMATO_PREDETERMINADO
            'dtTimePicker.Format = DateTimePickerFormat.Custom
            dtTimePicker.Value = Today
        Else
            dtTimePicker.Value = fecha
        End If

    End Sub

    Public Shared Sub setDgvColumnsReadonly(ByRef dgvGeneral As DataGridView)

        dgvGeneral.ReadOnly = False
        For Each dgvColumn As DataGridViewColumn In dgvGeneral.Columns
            dgvColumn.ReadOnly = True
        Next

    End Sub

    Public Shared Sub visualizarPanel(dgvGeneral As DataGridView,
                                       pnlTexto As Panel,
                                       esEditable As Boolean,
                                       txtMensaje As TextBox,
                                       nombreColumna As String)
        If dgvGeneral.Rows(dgvGeneral.CurrentRow.Index).Cells(nombreColumna).Selected = True Then
            pnlTexto.Visible = True
            txtMensaje.ReadOnly = esEditable
            txtMensaje.Text = dgvGeneral.Rows(dgvGeneral.CurrentRow.Index).Cells(nombreColumna).Value.ToString
            txtMensaje.Focus()
            txtMensaje.SelectionStart = txtMensaje.TextLength
        End If
    End Sub
    Public Shared Sub comboSoloEleccion(ByRef cbCombo As ComboBox)
        cbCombo.AutoCompleteMode = AutoCompleteMode.None
        cbCombo.AutoCompleteSource = AutoCompleteSource.None
        cbCombo.DropDownStyle = ComboBoxStyle.DropDownList
    End Sub

    Public Shared Sub deshabilitarOrdenarDgv(ByRef dgv As DataGridView)

        For Each columna As DataGridViewColumn In dgv.Columns
            columna.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub

    Public Shared Sub deshabilitarResizeDgv(ByRef dgv As DataGridView)
        For Each columna As DataGridViewColumn In dgv.Columns
            columna.Resizable = DataGridViewTriState.False
        Next
    End Sub

    'Public Shared Sub clonarObj(objOriginal As Object, objClon As Object)

    '    For Each p As System.Reflection.PropertyInfo In objOriginal.GetType().GetProperties()
    '        If p.CanRead Then
    '            objClon.GetType().GetProperty(p.Name).SetValue(objClon, p.GetValue(objOriginal, Nothing))
    '        End If
    '    Next
    'End Sub
    Public Shared Function cargarInformacionCorreo(ByVal codigoForm As String, Optional rutaReporte As String = "") As List(Of Correo)
        'Dim correo As New Correo
        'Dim fila As DataRow
        Dim tblTemp As New DataTable
        Dim listaCorreo As New List(Of Correo)
        Dim params As New List(Of String)
        params.Add(codigoForm)
        params.Add(SesionActual.codigoEP)
        params.Add(SesionActual.idUsuario)
        Try
            General.llenarTablaConExcepcion(Consultas.CORREO_CONFIGURACION_OBTENER_CORREO, params, tblTemp)
            If tblTemp.Rows.Count > 0 Then
                Dim correo As Correo
                For Each filaEnTabla As DataRow In tblTemp.Rows
                    correo = New Correo
                    correo.correo = filaEnTabla.Item("correo")
                    correo.pass = filaEnTabla.Item("pass")
                    correo.configuracionSMTP = filaEnTabla.Item("smtp")
                    correo.puerto = filaEnTabla.Item("puerto")
                    correo.asunto = filaEnTabla.Item("asunto").ToString
                    listaCorreo.Add(correo)
                Next
            Else
                listaCorreo = Nothing
            End If
            Return listaCorreo
        Catch ex As Exception
            listaCorreo = Nothing
            Throw ex
        End Try
        Return listaCorreo
    End Function
    Public Shared Sub ConfigurarCorreo(ByVal correoConfigurado As Correo, ByVal tipo As Integer, ByVal cuerpo As String, ByVal destinatario As String)
        correoConfigurado.rutaArchivo = Funciones.crearCarpeta(tipo)
        correoConfigurado.cuerpo = cuerpo
        correoConfigurado.correoDestinatario = destinatario
    End Sub

    Public Shared Sub abrirOpcionNotificacion(drFila As DataRow)
        Dim opcion As ElementoMenu
        opcion = New ElementoMenu(drFila("Modulo").ToString,
                                  drFila("Formulario").ToString,
                                  Nothing,
                                  drFila("Descripcion_menu").ToString)
        cargarFormulario(opcion, drFila)

    End Sub
    Private Shared Sub cargarFormulario(elemMenu As ElementoMenu, drFila As DataRow)
        Try
            Dim principal As New PrincipalDAL
            If principal.estaAbierto(elemMenu) = True Then
                principal.cerrarFormulario(elemMenu)
                If Not principal.estaAbierto(elemMenu) Then
                    cargarFormulario(elemMenu, drFila)
                End If
            Else
                Dim nombreTipo = "Celer." & elemMenu.nombre
                Dim vTipo As Type = Assembly.GetExecutingAssembly.GetType(nombreTipo)

                If vTipo IsNot Nothing Then
                    Dim vFormulario = Activator.CreateInstance(vTipo)
                    vFormulario.tag = elemMenu
                    cargarForm(vFormulario)
                    vFormulario.cargarNotificacion(drFila)
                End If
            End If

        Catch ex As Exception
            General.manejoErrores(ex)
        End Try

    End Sub
    Public Shared Sub enviarMensajeriaMovil(puertoConexion As String,
                                            mensaje As String,
                                            numeroDestino As String)
        Dim puertoUSB As New System.IO.Ports.SerialPort
        Try
            puertoUSB.PortName = puertoConexion 'DEFINE EL PUERTO DE CONEXION
            puertoUSB.Open() 'ABRE EL PUERTO
            puertoUSB.Write("AT" & vbCrLf) 'COMPRUEBA LA CONEXION CON EL MODEM
            Threading.Thread.Sleep(100)
            puertoUSB.Write("AT+CMGF=1" & vbCrLf) 'INDICA QUE VA A ENVIAR SMS
            Threading.Thread.Sleep(100)
            puertoUSB.Write("AT+CMGS=" & Chr(34) & numeroDestino & Chr(34) & vbCrLf) 'ENVIARA EL MENSAJE SIN ALMACENAMIENTO
            Threading.Thread.Sleep(100)
            puertoUSB.Write(mensaje & Chr(26)) ' ENVIA EL MENSAJE
        Catch ex As Exception
            MsgBox("Mensaje no enviado", MsgBoxStyle.Exclamation)
        End Try
    End Sub
    Public Shared Function validarCorreo(correo As String) As Boolean
        Dim re As New RegexStringValidator("^[\w._%-]+@[\w.-]+\.[a-zA-Z]{2,4}$")
        Dim respuesta As Boolean
        Try
            re.Validate(correo)
            respuesta = True
        Catch ex As Exception
            respuesta = False
        End Try
        Return respuesta
    End Function
    'Public Shared Sub abrirAreaTexto(ByVal sender As System.Object, ByVal e As EventArgs)
    '    If sender.enabled = True AndAlso sender.readonly = False Then
    '        Dim localizacion = sender.location
    '        Dim tamano = sender.size
    '        Dim area = sender.tag
    '        area.Size = tamano
    '        Dim padre = sender.parent
    '        area.Location = localizacion
    '        area.Location = New Point(area.Location.X, area.Location.Y)
    '        area.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Style3D
    '        Dim teclaPresionada As String
    '        If (TypeOf e Is KeyPressEventArgs) Then
    '            teclaPresionada = CType(e, KeyPressEventArgs).KeyChar
    '            If (String.IsNullOrEmpty(teclaPresionada) = False) AndAlso ((Asc(teclaPresionada) >= 65 And Asc(teclaPresionada) <= 90) _
    '                  Or (Asc(teclaPresionada) >= 97 And Asc(teclaPresionada) <= 122) _
    '                  Or (Asc(teclaPresionada) >= 48 And Asc(teclaPresionada) <= 57) _
    '                  Or (teclaPresionada = "ñ") Or (teclaPresionada = "Ñ")) Then
    '            Else
    '                teclaPresionada = ""
    '            End If
    '        End If

    '        sender.text = sender.text & teclaPresionada
    '        area.Text = sender.text
    '            area.Font = sender.Font
    '            sender.readonly = True
    '            If padre.name.ToString.ToLower.Contains("justifica") Then
    '                Dim localizacionPadre = padre.location
    '                Dim tamanoPadre = padre.size
    '                Dim contenedorPadre As New Panel
    '                contenedorPadre.Name = "panelNuevoDev"
    '                contenedorPadre.Location = localizacionPadre
    '                contenedorPadre.Size = tamanoPadre
    '                contenedorPadre.Parent = padre.parent
    '                contenedorPadre.Visible = True
    '                area.parent = contenedorPadre
    '                contenedorPadre.BringToFront()
    '            Else
    '                area.parent = padre
    '                area.BringToFront()
    '            End If
    '            area.Tag = sender
    '            area.visible = True
    '            area.Focus()
    '            area.SelectionStart = area.text.length
    '            AddHandler CType(area, DevExpress.XtraEditors.MemoEdit).Leave, AddressOf cerrarAreaTexto
    '        End If

    'End Sub
    'Public Shared Sub agregarEventoDevExp(ByRef pFormulario As Object, memoEdit1 As Object)

    '    For Each vControl In pFormulario.Controls

    '        If (TypeOf vControl Is TextBox) OrElse (TypeOf vControl Is RichTextBox) Then
    '            agregarEvento(vControl, memoEdit1)
    '        ElseIf (TypeOf vControl Is Label) OrElse vControl.text.contains(Constantes.LINEA_TITULO) OrElse (TypeOf vControl Is MaskedTextBox) OrElse
    '            (TypeOf vControl Is ComboBox) OrElse (TypeOf vControl Is DateTimePicker) OrElse (TypeOf vControl Is NumericUpDown) OrElse
    '            (TypeOf vControl Is CheckBox) OrElse (TypeOf vControl Is PictureBox) OrElse (TypeOf vControl Is CheckedListBox) OrElse (TypeOf vControl Is RadioButton) OrElse
    '            (TypeOf vControl Is TreeView) OrElse (TypeOf vControl Is DataGridView) Then
    '        Else
    '            ' mira a ve si es un contenedor
    '            agregarEventoDevExp(vControl, memoEdit1)
    '        End If
    '    Next
    'End Sub

    'Public Shared Sub agregarEvento(ByRef vControl As Object, memoEdit1 As Object)
    '    vControl.tag = memoEdit1
    '    If (vControl.productname = "DevExpress.XtraEditors") Then Exit Sub
    '    If (TypeOf vControl Is TextBox) Then
    '        AddHandler CType(vControl, TextBox).Click, AddressOf abrirAreaTexto
    '        AddHandler CType(vControl, TextBox).KeyPress, AddressOf abrirAreaTexto
    '    ElseIf (TypeOf vControl Is RichTextBox) Then
    '        AddHandler CType(vControl, RichTextBox).Click, AddressOf abrirAreaTexto
    '        AddHandler CType(vControl, RichTextBox).KeyPress, AddressOf abrirAreaTexto
    '    End If

    'End Sub
    'Public Shared Sub agregarDicc(SpellChecker1 As DevExpress.XtraSpellChecker.SpellChecker)
    '    Try
    '        Dim ruta, rutaTxt, rutaDic, rutaAff As String
    '        ruta = Directory.GetCurrentDirectory() & "\Resources"
    '        rutaTxt = ruta & "\palabrasDiccionario.txt"
    '        rutaAff = ruta & "\es_ANY.aff"
    '        rutaDic = ruta & "\es_ANY.dic"
    '        Dim dic_es_ES As New DevExpress.XtraSpellChecker.HunspellDictionary
    '        dic_es_ES.DictionaryPath = rutaDic
    '        dic_es_ES.GrammarPath = rutaAff
    '        dic_es_ES.Culture = New Globalization.CultureInfo("es-CO")
    '        SpellChecker1.Dictionaries.Add(dic_es_ES)
    '        SpellChecker1.OptionsSpelling.IgnoreUpperCaseWords = DevExpress.Utils.DefaultBoolean.False
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub
    'Public Shared Sub cerrarAreaTexto(ByVal memoArea As Object, ByVal e As EventArgs)
    '    If memoArea.Visible Then
    '        memoArea.Tag.text = Replace(memoArea.Text, vbCrLf, "")
    '        If memoArea.parent.name.ToString.ToLower.Contains("panelnuevodev") Then
    '            memoArea.parent.visible = False
    '            memoArea.Tag.parent.visible = True
    '            memoArea.tag.focus()
    '        Else
    '            memoArea.Visible = False
    '        End If
    '        memoArea.Tag.readonly = False
    '        memoArea.Tag.parent.Focus()
    '        memoArea.Tag.SelectionStart = memoArea.Tag.text.length
    '    End If


    'End Sub

End Class
