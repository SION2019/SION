Imports System.Data.SqlClient
Public Class CupsDAL

    Public Function llenardgv(TextBox1 As String, codigo As String) As DataTable
        Dim dt As New DataTable
        Dim cdna, cadena As String
        Dim adpter As SqlDataAdapter
        Dim enlce_dta As BindingSource = New BindingSource

        cdna = "" : cadena = ""
        cdna = "EXEC [SP_CUPS_CARGAR] '" & TextBox1 & "','" & codigo & "'"

        Try
            dt.Clear()
            adpter = New SqlDataAdapter(cdna, FormPrincipal.cnxion)
            adpter.Fill(dt) : adpter.Dispose()
        Catch ex As Exception : MsgBox("Error " + ex.Message)
        Finally

        End Try
        Return dt
    End Function

    Public Function guardar(ByRef dgvProcedimiento As DataGridView, ByVal codigoRef As String) As Boolean
        Dim i, j, l, indcdor, clmnaInicio, clmnaFin, clmnaCategoria As Integer
        Dim cdna, cadena, cdnaInsert, cdnaInsertNuevo, cdnaUpdate, cdnaDelete, valor As String
        Dim dtUpdate, dtDelete As New DataTable
        Dim bndra As Boolean = True
        clmnaInicio = 4 : clmnaFin = 10 : clmnaCategoria = clmnaInicio + 2
        cdnaInsertNuevo = ConsultasAsis.MANUAL_CUPS_GUARDAR
        cdnaInsert = "INSERT INTO M_LISTA_MANUAL_D(Codigo_Manual,Codigo_Procedimiento,Descripcion,Facturable,Equipo,Visible,Usuario_creacion,Fecha_creacion) VALUES(Codigo_Manu4l,"
        cdnaUpdate = "UPDATE M_LISTA_MANUAL_D SET Codigo_Procedimiento='_valor1',Descripcion='_valor2',Facturable='_valor5',Equipo='_valor6',Visible='_valor7',Usuario_actualizacion=" & SesionActual.idUsuario & ",Fecha_actualizacion=getdate() WHERE Codigo_Procedimiento='_condicion' and Codigo_Manual=" & codigoRef & ""
        cdnaDelete = "DELETE M_LISTA_MANUAL_D WHERE Codigo_Procedimiento='_condicion' and Codigo_Manual=" & codigoRef & ""
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For i = 0 To dgvProcedimiento.RowCount - 1

                        If (IsDBNull(dgvProcedimiento.Rows(i).Cells("Inserción").Value) = True Or dgvProcedimiento.Rows(i).Cells("Código").Value.ToString = "") And i <> dgvProcedimiento.RowCount - 1 Then
                            MsgBox("error al validar dgv, falta información", MsgBoxStyle.Critical)
                            Return False
                            Exit For
                        ElseIf (IsDBNull(dgvProcedimiento.Rows(i).Cells("Inserción").Value) = True Or dgvProcedimiento.Rows(i).Cells("Código").Value.ToString = "") And i = dgvProcedimiento.RowCount - 1 Then
                            Exit For
                        End If
                        If dgvProcedimiento.Rows(i).Cells("Inserción").Value = True Then
                            bndra = validardgv(dgvProcedimiento, i, clmnaInicio, clmnaFin)
                            If bndra Then

                                Return False
                            Else
                                If indcdor <> 0 Then
                                    cdnaInsert = cdnaInsert & ",(Codigo_Manu4l,"
                                End If
                                cdnaInsertNuevo = cdnaInsertNuevo & "'" & Replace(dgvProcedimiento.Rows(i).Cells(clmnaInicio).Value.ToString(), "'", "") & "','" &
                                                    Replace(dgvProcedimiento.Rows(i).Cells(clmnaCategoria).Value.ToString(), "'", "") & "'," &
                                                    SesionActual.idUsuario & "; "
                                For j = clmnaInicio To clmnaFin
                                    If dgvProcedimiento.Rows(i).Cells(j).ValueType.ToString = "System.String" Then
                                        valor = Replace(dgvProcedimiento.Rows(i).Cells(j).Value.ToString(), "'", "")
                                    Else
                                        If IsDBNull(dgvProcedimiento.Rows(i).Cells(j).Value) Then
                                            valor = "False"
                                        Else
                                            valor = dgvProcedimiento.Rows(i).Cells(j).Value
                                        End If
                                    End If
                                    If j <> 6 And j <> 7 Then
                                        If j = clmnaFin Then
                                            cdnaInsert = cdnaInsert & "'" & valor & "'"
                                        Else
                                            cdnaInsert = cdnaInsert & "'" & valor & "'"
                                        End If

                                        If j <> clmnaFin Then
                                            cdnaInsert = cdnaInsert & ","
                                        End If
                                    End If
                                Next
                                cdnaInsert = cdnaInsert & "," & SesionActual.idUsuario & ",getdate())"
                                cdnaInsert = Replace(cdnaInsert, "Codigo_Manu4l", codigoRef)
                                indcdor = indcdor + 1
                            End If
                        ElseIf dgvProcedimiento.Rows(i).Cells("Edición").Value = True Then
                            bndra = validardgv(dgvProcedimiento, i, clmnaInicio, clmnaFin)
                            If bndra Then

                                Return False
                            Else
                                l = 1
                                cdna = cdnaUpdate
                                For j = clmnaInicio To clmnaFin
                                    cadena = "_valor" & CStr(l)
                                    If dgvProcedimiento.Rows(i).Cells(j).ValueType.ToString = "System.String" Then
                                        valor = Replace(dgvProcedimiento.Rows(i).Cells(j).Value.ToString(), "'", "")
                                    Else
                                        If IsDBNull(dgvProcedimiento.Rows(i).Cells(j).Value) Then
                                            valor = "False"
                                        Else
                                            valor = dgvProcedimiento.Rows(i).Cells(j).Value
                                        End If
                                    End If

                                    If j = clmnaFin Then
                                        cdna = Replace(cdna, cadena, valor)
                                    Else
                                        cdna = Replace(cdna, cadena, valor)
                                    End If
                                    l = l + 1
                                Next
                                cdna = Replace(cdna, "_condicion", dgvProcedimiento.Rows(i).Cells(clmnaInicio - 1).Value.ToString())
                                If dtUpdate.Columns.Count <= 0 Then
                                    dtUpdate.Columns.Add("Consulta")
                                End If
                                dtUpdate.Rows.Add()
                                dtUpdate.Rows(dtUpdate.Rows.Count - 1).Item(0) = cdna
                            End If
                        ElseIf dgvProcedimiento.Rows(i).Cells("Eliminación").Value = True Then
                            cdnaDelete = Replace(cdnaDelete, "_condicion", dgvProcedimiento.Rows(i).Cells(clmnaInicio - 1).Value.ToString())
                            If dtDelete.Columns.Count <= 0 Then
                                dtDelete.Columns.Add("Consulta")
                            End If
                            dtDelete.Rows.Add()
                            dtDelete.Rows(dtDelete.Rows.Count - 1).Item(0) = cdnaDelete
                        End If
                    Next

                    If dtUpdate.Rows.Count > 0 Then
                        For k = 0 To dtUpdate.Rows.Count - 1
                            consulta.CommandText = dtUpdate.Rows(k).Item(0).ToString
                            consulta.ExecuteNonQuery()
                        Next
                    ElseIf indcdor > 0 Then
                        consulta.CommandText = cdnaInsertNuevo
                        consulta.ExecuteNonQuery()
                        consulta.CommandText = cdnaInsert
                        consulta.ExecuteNonQuery()

                    ElseIf dtDelete.Rows.Count > 0 Then
                        For k = 0 To dtDelete.Rows.Count - 1
                            consulta.CommandText = dtDelete.Rows(k).Item(0).ToString
                            consulta.ExecuteNonQuery()
                        Next
                    End If
                    If dtUpdate.Rows.Count > 0 Or indcdor > 0 Or dtDelete.Rows.Count > 0 Then
                        trnsccion.Commit()
                        Return True
                    Else
                        MsgBox("revise bien antes guardar, falta información", MsgBoxStyle.Exclamation)
                        Return False
                    End If
                End Using
            End Using
        Catch ex As Exception
            Try
                MsgBox("Error al realizar la transacción. Se revertirán los cambios. " & ex.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Error")
            Catch ex1 As SqlException
            End Try
            Return False
        Finally
        End Try
    End Function
    Private Function validardgv(ByRef dgvprocedimiento As DataGridView, fila_actual As Integer, clmna_inicio As Integer, clmna_fin As Integer) As Boolean
        Dim bndra As Boolean
        Dim i As Integer
        Dim fila As String

        bndra = False
        fila = dgvprocedimiento.Rows(fila_actual).Cells(8).Value.ToString

        For i = clmna_inicio To clmna_fin
            If dgvprocedimiento.Rows(fila_actual).Cells(i).ValueType = Type.GetType("System.String") Then
                If String.IsNullOrEmpty(dgvprocedimiento.Rows(fila_actual).Cells(i).Value.ToString) Or String.IsNullOrWhiteSpace(dgvprocedimiento.Rows(fila_actual).Cells(i).Value.ToString) Then
                    bndra = True
                    MsgBox("Hace falta un dato", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Advertencia")
                    dgvprocedimiento.Rows(fila_actual).Cells(i).Selected = True
                    Exit For
                End If
            End If
        Next
        Return bndra
    End Function

    'Public Sub buscar()
    '    Dim form_busqueda As New Form_Busqueda
    '    form_busqueda.Text = "Busqueda de categorias"
    '    form_busqueda.FormPadre = Name
    '    form_busqueda.Cadena = ""
    '    form_busqueda.Cadena = "EXEC [SP_CATEGORIA_CUPS_BUSCAR] ''"
    '    form_busqueda.ShowDialog()
    'End Sub


End Class
