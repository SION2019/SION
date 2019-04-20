Public Class ViaAdministracionBLL
    Dim objViaAdministracionC As New ViaAdministracionDAL
    Public Function guardarViaAdministracion(ByVal objViaAdministracion As ViaAdministracion) As String
        If String.IsNullOrEmpty(objViaAdministracion.codigo) Then
            Return objViaAdministracionC.crearViaAdministracion(objViaAdministracion)

        Else
            objViaAdministracionC.actualizarViaAdministracion(objViaAdministracion)
        End If
        Return Nothing
    End Function

    Public Sub cargarComboTipoVia(cbCombo As ComboBox)
        Dim dtTabla As New DataTable
        Try

            Dim drFila As DataRow = dtTabla.NewRow()
            Dim drFila2 As DataRow = dtTabla.NewRow()
            Dim drFila3 As DataRow = dtTabla.NewRow()

            dtTabla.Columns.Add("codigo")
            dtTabla.Columns.Add("tipo_via")

            drFila.Item(0) = "-1"
            drFila.Item(1) = " - - - Seleccione - - - "
            dtTabla.Rows.Add(drFila)

            drFila2.Item(0) = "QX"
            drFila2.Item(1) = "Quirurgico"
            dtTabla.Rows.Add(drFila2)

            drFila3.Item(0) = "MED"
            drFila3.Item(1) = "Medicamento"
            dtTabla.Rows.Add(drFila3)

            cbCombo.DataSource = dtTabla
            cbCombo.DisplayMember = "tipo_via"
            cbCombo.ValueMember = "codigo"
            cbCombo.AutoCompleteMode = AutoCompleteMode.None
            cbCombo.AutoCompleteSource = AutoCompleteSource.None
            cbCombo.DropDownStyle = ComboBoxStyle.DropDownList

        Catch ex As Exception
            general.manejoErrores(ex) 
        End Try
    End Sub

End Class
