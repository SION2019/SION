Imports System.Data.SqlClient
Public Class CupsSeccionDAL
    Public Function llenar_dgv(TextBox1 As String) As DataTable
        Dim dt As New DataTable
        Dim cdna, cadena As String
        Dim adpter As SqlDataAdapter
        Dim enlce_dta As BindingSource = New BindingSource

        cdna = "" : cadena = ""
        cdna = "EXEC [SP_SCUPS_CARGAR] '" & TextBox1 & "'"

        Try
            dt.Clear()
            adpter = New SqlDataAdapter(cdna, FormPrincipal.cnxion)
            adpter.Fill(dt) : adpter.Dispose()
        Catch ex As Exception : MsgBox("Error " + ex.Message)
        Finally

        End Try
        Return dt
    End Function

    Public Function guardar(form_seccion_cups As Form_SECCION_CUPS) As Boolean
        Dim i, j, l, indcdor, clmna_inicio, clmna_fin, cmpo_inicio, cmpo_fin As Integer
        Dim cdna, cadena, cdna_insert, cdna_update, cdna_delete, valor As String
        Dim dt_update, dt_delete As New DataTable
        Dim bndra As Boolean

        cdna_insert = "" : cdna_update = "" : cdna_delete = "" : indcdor = 0 : bndra = True

        clmna_inicio = 4 : clmna_fin = 5 : cmpo_inicio = 4 : cmpo_fin = 9
        cdna_insert = "INSERT INTO M_SECCIONES_CUPS(Codigo_Seccion,Descripcion) VALUES("
        cdna_update = "UPDATE M_SECCIONES_CUPS SET Codigo_Seccion='_valor1',Descripcion='_valor2' WHERE Codigo_Seccion='_condicion'"
        cdna_delete = "DELETE M_SECCIONES_CUPS WHERE Codigo_Seccion='_condicion'"
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For i = 0 To form_seccion_cups.dgvprocedimiento.RowCount - 1

                        If (IsDBNull(form_seccion_cups.dgvprocedimiento.Rows(i).Cells("Inserción").Value) = True Or form_seccion_cups.dgvprocedimiento.Rows(i).Cells("Código").Value.ToString = "") And i <> form_seccion_cups.dgvprocedimiento.RowCount - 1 Then
                            MsgBox("error al validar dgv, falta información", MsgBoxStyle.Critical)
                            Return False
                            Exit For
                        ElseIf (IsDBNull(form_seccion_cups.dgvprocedimiento.Rows(i).Cells("Inserción").Value) = True Or form_seccion_cups.dgvprocedimiento.Rows(i).Cells("Código").Value.ToString = "") And i = form_seccion_cups.dgvprocedimiento.RowCount - 1 Then
                            Exit For
                        End If
                        If form_seccion_cups.dgvprocedimiento.Rows(i).Cells("Inserción").Value = True Then
                            bndra = Validar_dgv(i, clmna_inicio, clmna_fin, form_seccion_cups)
                            If bndra Then

                                Return False
                            Else
                                If indcdor <> 0 Then
                                    cdna_insert = cdna_insert & ",("
                                End If
                                For j = clmna_inicio To clmna_fin
                                    If form_seccion_cups.dgvprocedimiento.Rows(i).Cells(j).ValueType.ToString = "System.String" Then
                                        valor = form_seccion_cups.dgvprocedimiento.Rows(i).Cells(j).Value.ToString()
                                    Else
                                        If IsDBNull(form_seccion_cups.dgvprocedimiento.Rows(i).Cells(j).Value) Then
                                            valor = "False"
                                        Else
                                            valor = "True"
                                        End If
                                    End If
                                    If j <> 7 Then
                                        If j = clmna_fin Then
                                            cdna_insert = cdna_insert & "'" & valor & "'"
                                        Else
                                            cdna_insert = cdna_insert & "'" & valor & "'"
                                        End If

                                        If j <> clmna_fin Then
                                            cdna_insert = cdna_insert & ","
                                        End If
                                    End If
                                Next
                                cdna_insert = cdna_insert & ")"
                                indcdor = indcdor + 1
                            End If
                        ElseIf form_seccion_cups.dgvprocedimiento.Rows(i).Cells("Edición").Value = True Then
                            bndra = Validar_dgv(i, clmna_inicio, clmna_fin, form_seccion_cups)
                            If bndra Then

                                Return False
                            Else
                                l = 1
                                cdna = cdna_update
                                For j = clmna_inicio To clmna_fin
                                    cadena = "_valor" & CStr(l)
                                    If form_seccion_cups.dgvprocedimiento.Rows(i).Cells(j).ValueType.ToString = "System.String" Then
                                        valor = form_seccion_cups.dgvprocedimiento.Rows(i).Cells(j).Value.ToString()
                                    Else
                                        valor = form_seccion_cups.dgvprocedimiento.Rows(i).Cells(j).Value.ToString()
                                    End If

                                    If j = clmna_fin Then
                                        cdna = Replace(cdna, cadena, valor)
                                    Else
                                        cdna = Replace(cdna, cadena, valor)
                                    End If
                                    l = l + 1
                                Next
                                cdna = Replace(cdna, "_condicion", form_seccion_cups.dgvprocedimiento.Rows(i).Cells(clmna_inicio - 1).Value.ToString())
                                If dt_update.Columns.Count <= 0 Then
                                    dt_update.Columns.Add("Consulta")
                                End If
                                dt_update.Rows.Add()
                                dt_update.Rows(dt_update.Rows.Count - 1).Item(0) = cdna
                            End If
                        ElseIf form_seccion_cups.dgvprocedimiento.Rows(i).Cells("Eliminación").Value = True Then
                            cdna_delete = Replace(cdna_delete, "_condicion", form_seccion_cups.dgvprocedimiento.Rows(i).Cells(clmna_inicio - 1).Value.ToString())
                            If dt_delete.Columns.Count <= 0 Then
                                dt_delete.Columns.Add("Consulta")
                            End If
                            dt_delete.Rows.Add()
                            dt_delete.Rows(dt_delete.Rows.Count - 1).Item(0) = cdna_delete
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
    Private Function Validar_dgv(fila_actual As Integer, clmna_inicio As Integer, clmna_fin As Integer, form_seccion_cups As Form_SECCION_CUPS) As Boolean
        Dim bndra As Boolean
        Dim i As Integer

        bndra = False

        For i = clmna_inicio To clmna_fin
            If form_seccion_cups.dgvprocedimiento.Rows(fila_actual).Cells(i).ValueType = Type.GetType("System.String") Then
                If String.IsNullOrEmpty(form_seccion_cups.dgvprocedimiento.Rows(fila_actual).Cells(i).Value.ToString) Or String.IsNullOrWhiteSpace(form_seccion_cups.dgvprocedimiento.Rows(fila_actual).Cells(i).Value.ToString) Then
                    bndra = True
                    MsgBox("Hace falta un dato", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Advertencia")
                    form_seccion_cups.dgvprocedimiento.Rows(fila_actual).Cells(i).Selected = True
                    Exit For
                End If
            End If
        Next
        Return bndra
    End Function


End Class
