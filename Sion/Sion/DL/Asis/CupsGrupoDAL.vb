Imports System.Data.SqlClient
Public Class CupsGrupoDAL

    Public Function llenar_dgv(TextBox1 As String) As DataTable
        Dim dt As New DataTable
        Dim cdna, cadena As String
        Dim adpter As SqlDataAdapter
        Dim enlce_dta As BindingSource = New BindingSource

        cdna = "" : cadena = ""
        cdna = "EXEC [SP_GCUPS_CARGAR] '" & TextBox1 & "'"

        Try
            dt.Clear()
            adpter = New SqlDataAdapter(cdna, FormPrincipal.cnxion)
            adpter.Fill(dt) : adpter.Dispose()
        Catch ex As Exception : MsgBox("Error " + ex.Message)
        Finally

        End Try
        Return dt
    End Function

    Public Function guardar(ByRef dgvprocedimiento As DataGridView, usuario As Integer) As Boolean
        Dim i, j, contadorColumna, indcdor, clmna_inicio, clmna_fin, cmpo_inicio, cmpo_fin As Integer
        Dim cdna, cadena, cdna_insert, cdna_update, cdna_delete, valor As String
        Dim dt_update, dt_delete As New DataTable
        Dim bndra As Boolean

        cdna_insert = "" : cdna_update = "" : cdna_delete = "" : indcdor = 0 : bndra = True

        clmna_inicio = 4 : clmna_fin = 6 : cmpo_inicio = 4 : cmpo_fin = 9
        cdna_insert = "INSERT INTO M_GRUPOS_CUPS(Codigo_Grupo,Descripcion ,Codigo_Capitulo,usuario_creacion,fecha_creacion) VALUES ("
        cdna_update = "UPDATE M_GRUPOS_CUPS SET Codigo_Grupo='_valor1',Descripcion='_valor2',Codigo_Capitulo='_valor3',usuario_actualizacion=" & usuario & " 
                        ,fecha_actualizacion=GETDATE() WHERE Codigo_Grupo='_condicion'"
        cdna_delete = "DELETE M_GRUPOS_CUPS WHERE Codigo_Grupo='_condicion' "
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For i = 0 To dgvprocedimiento.RowCount - 1

                        If (IsDBNull(dgvprocedimiento.Rows(i).Cells("Inserción").Value) = True Or dgvprocedimiento.Rows(i).Cells("Código").Value.ToString = "") And i <> dgvprocedimiento.RowCount - 1 Then
                            MsgBox("error al validar dgv, falta información", MsgBoxStyle.Critical)
                            Return False
                            Exit For
                        ElseIf (IsDBNull(dgvprocedimiento.Rows(i).Cells("Inserción").Value) = True Or dgvprocedimiento.Rows(i).Cells("Código").Value.ToString = "") And i = dgvprocedimiento.RowCount - 1 Then
                            Exit For
                        End If
                        If dgvprocedimiento.Rows(i).Cells("Inserción").Value = True Then
                            bndra = Validar_dgv(dgvprocedimiento, i, clmna_inicio, clmna_fin)
                            If bndra Then

                                Return False
                            Else
                                If indcdor <> 0 Then
                                    cdna_insert = cdna_insert & ",("
                                End If
                                For j = clmna_inicio To clmna_fin
                                    If dgvprocedimiento.Rows(i).Cells(j).ValueType.ToString = "System.String" Then
                                        valor = dgvprocedimiento.Rows(i).Cells(j).Value.ToString()
                                    Else
                                        If IsDBNull(dgvprocedimiento.Rows(i).Cells(j).Value) Then
                                            valor = "False"
                                        Else
                                            valor = "True"
                                        End If
                                    End If
                                    If j <> 7 Then
                                        If j = clmna_fin Then
                                            cdna_insert = cdna_insert & "'" & Replace(valor, "'", "´") & "'"
                                        Else
                                            cdna_insert = cdna_insert & "'" & Replace(valor, "'", "´") & "'"
                                        End If

                                        If j <> clmna_fin Then
                                            cdna_insert = cdna_insert & ","
                                        End If
                                    End If
                                Next
                                cdna_insert = cdna_insert & "," & usuario & ",GETDATE())"
                                indcdor = indcdor + 1
                            End If
                        ElseIf dgvprocedimiento.Rows(i).Cells("Edición").Value = True Then
                            bndra = Validar_dgv(dgvprocedimiento, i, clmna_inicio, clmna_fin)
                            If bndra Then

                                Return False
                            Else
                                contadorColumna = 1
                                cdna = cdna_update
                                For j = clmna_inicio To clmna_fin
                                    cadena = "_valor" & CStr(contadorColumna)
                                    If dgvprocedimiento.Rows(i).Cells(j).ValueType.ToString = "System.String" Then
                                        valor = dgvprocedimiento.Rows(i).Cells(j).Value.ToString()
                                    Else
                                        valor = dgvprocedimiento.Rows(i).Cells(j).Value.ToString()
                                    End If

                                    If j = clmna_fin Then
                                        cdna = Replace(cdna, cadena, Replace(valor, "'", "´"))
                                    Else
                                        cdna = Replace(cdna, cadena, Replace(valor, "'", "´"))
                                    End If
                                    contadorColumna = contadorColumna + 1
                                Next
                                cdna = Replace(cdna, "_condicion", dgvprocedimiento.Rows(i).Cells(clmna_inicio - 1).Value.ToString())
                                If dt_update.Columns.Count <= 0 Then
                                    dt_update.Columns.Add("Consulta")
                                End If
                                dt_update.Rows.Add()
                                dt_update.Rows(dt_update.Rows.Count - 1).Item(0) = cdna
                            End If
                        ElseIf dgvprocedimiento.Rows(i).Cells("Eliminación").Value = True Then
                            cdna = Replace(cdna_delete, "_condicion", Replace(dgvprocedimiento.Rows(i).Cells(clmna_inicio - 1).Value.ToString(), "'", "´"))
                            If dt_delete.Columns.Count <= 0 Then
                                dt_delete.Columns.Add("Consulta")
                            End If
                            dt_delete.Rows.Add()
                            dt_delete.Rows(dt_delete.Rows.Count - 1).Item(0) = cdna
                        End If
                    Next

                    If dt_update.Rows.Count > 0 Then
                        For k = 0 To dt_update.Rows.Count - 1
                            consulta.CommandText = dt_update.Rows(k).Item(0).ToString
                            consulta.ExecuteNonQuery()
                        Next
                    ElseIf indcdor > 0 Then

                        consulta.CommandText = cdna_insert
                        consulta.ExecuteNonQuery()
                    ElseIf dt_delete.Rows.Count > 0 Then
                        For k = 0 To dt_delete.Rows.Count - 1
                            consulta.CommandText = dt_delete.Rows(k).Item(0).ToString
                            consulta.ExecuteNonQuery()
                        Next
                    End If
                    If dt_update.Rows.Count > 0 Or indcdor > 0 Or dt_delete.Rows.Count > 0 Then
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
    Private Function Validar_dgv(ByRef dgvprocedimiento As DataGridView, fila_actual As Integer, clmna_inicio As Integer, clmna_fin As Integer) As Boolean
        Dim bndra As Boolean
        Dim i As Integer

        bndra = False

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
    '    form_busqueda.Text = "Busqueda de CAPITULOS"
    '    form_busqueda.formPadre = Name
    '    form_busqueda.Cadena = ""
    '    form_busqueda.Cadena = "EXEC [SP_CAPITULO_CUPS_BUSCAR] ''"
    '    form_busqueda.ShowDialog()
    'End Sub
End Class
