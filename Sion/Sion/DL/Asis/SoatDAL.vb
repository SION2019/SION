Imports System.Data.SqlClient
Public Class SoatDAL

    Public Function guardar(ByRef dgvprocedimiento As DataGridView, codigoManual As Integer) As Boolean
        Dim i, j, l, indcdor, clmna_inicio, clmna_fin As Integer
        Dim cdna, cadena, cdna_insert, cdna_update, cdna_delete, valor As String
        Dim dt_update, dt_delete As New DataTable
        Dim bndra As Boolean = True

        clmna_inicio = 4 : clmna_fin = 16
        cdna_insert = "INSERT INTO M_PROCEDIMIENTOS_SOAT(Codigo_Manual,Codigo_SOAT,Descripcion,Codigo_GRUPOQX,FactOR,Valor_Personalizado,CODIGO_GRUPO,CODIGO_CAPITULO,CODIGO_ARTICULO,FACTURABLE,Usuario_creacion,Fecha_creacion) VALUES(Codigo_Manu4l,"
        cdna_update = "UPDATE M_PROCEDIMIENTOS_SOAT Set Codigo_SOAT='_valor1',Descripcion='_valor2',Codigo_GRUPOQX='_valor3',FactOR='_valor5',Valor_Personalizado='_valor6',CODIGO_GRUPO='_valor7',CODIGO_CAPITULO='_valor9',CODIGO_ARTICULO='_valor01',FACTURABLE='_valor03',Usuario_actualizacion=" & SesionActual.idUsuario & ",Fecha_actualizacion=getdate()  WHERE Codigo_SOAT='_condicion' and Codigo_Manual=" & codigoManual & ""
        cdna_delete = "DELETE M_PROCEDIMIENTOS_SOAT WHERE Codigo_SOAT='_condicion' and Codigo_Manual=" & codigoManual & ""
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
                            If dgvprocedimiento.Rows(i).Cells("Factor").Value.ToString = "" Then
                                dgvprocedimiento.Rows(i).Cells("Factor").Value = "0"
                                dgvprocedimiento.EndEdit()
                            End If
                            bndra = Validar_dgv(dgvprocedimiento, i, clmna_inicio, clmna_fin)
                            If bndra Then

                                Return False
                            Else
                                If indcdor <> 0 Then
                                    cdna_insert = cdna_insert & ",(Codigo_Manu4l,"
                                End If
                                For j = clmna_inicio To clmna_fin
                                    If dgvprocedimiento.Rows(i).Cells(j).ValueType.ToString = "System.String" Or dgvprocedimiento.Rows(i).Cells(j).ValueType.ToString = "System.Double" Or dgvprocedimiento.Rows(i).Cells(j).ValueType.ToString = "System.Int32" Then
                                        valor = Replace(dgvprocedimiento.Rows(i).Cells(j).Value.ToString(), "'", "")
                                    Else
                                        If IsDBNull(dgvprocedimiento.Rows(i).Cells(j).Value) Then
                                            valor = "False"
                                        Else
                                            valor = dgvprocedimiento.Rows(i).Cells(j).Value
                                        End If
                                    End If
                                    If j <> 7 And j <> 11 And j <> 13 And j <> 15 Then
                                        If j = clmna_fin Then
                                            cdna_insert = cdna_insert & "'" & Replace(valor, ",", ".") & "'"
                                        Else
                                            cdna_insert = cdna_insert & "'" & Replace(valor, ",", ".") & "'"
                                        End If

                                        If j <> clmna_fin Then
                                            cdna_insert = cdna_insert & ","
                                        End If
                                    End If
                                Next
                                cdna_insert = cdna_insert & "," & SesionActual.idUsuario & ",getdate())"
                                cdna_insert = Replace(cdna_insert, "Codigo_Manu4l", codigoManual)
                                indcdor = indcdor + 1
                            End If
                        ElseIf dgvprocedimiento.Rows(i).Cells("Edición").Value = True Then
                            If dgvprocedimiento.Rows(i).Cells("Factor").Value.ToString = "" Then
                                dgvprocedimiento.Rows(i).Cells("Factor").Value = 0
                            End If
                            bndra = Validar_dgv(dgvprocedimiento, i, clmna_inicio, clmna_fin)
                            If bndra Then

                                Return False
                            Else
                                l = 1
                                cdna = cdna_update
                                For j = clmna_inicio To clmna_fin
                                    cadena = "_valor" & CStr(l)
                                    If dgvprocedimiento.Rows(i).Cells(j).ValueType.ToString = "System.String" Or dgvprocedimiento.Rows(i).Cells(j).ValueType.ToString = "System.Double" Or dgvprocedimiento.Rows(i).Cells(j).ValueType.ToString = "System.Int32" Then
                                        valor = Replace(dgvprocedimiento.Rows(i).Cells(j).Value.ToString(), "'", "")
                                    Else
                                        If IsDBNull(dgvprocedimiento.Rows(i).Cells(j).Value) Then
                                            valor = "False"
                                        Else
                                            valor = dgvprocedimiento.Rows(i).Cells(j).Value
                                        End If
                                    End If
                                    If cadena = "_valor11" Then
                                        cadena = "_valor01"
                                    ElseIf cadena = "_valor13" Then
                                        cadena = "_valor03"
                                    End If
                                    If j = clmna_fin Then
                                        cdna = Replace(cdna, cadena, Replace(valor, ",", "."))
                                    Else
                                        cdna = Replace(cdna, cadena, Replace(valor, ",", "."))
                                    End If
                                    l = l + 1
                                Next
                                cdna = Replace(cdna, "_condicion", dgvprocedimiento.Rows(i).Cells(clmna_inicio - 1).Value.ToString())
                                If dt_update.Columns.Count <= 0 Then
                                    dt_update.Columns.Add("Consulta")
                                End If
                                dt_update.Rows.Add()
                                dt_update.Rows(dt_update.Rows.Count - 1).Item(0) = cdna
                            End If
                        ElseIf dgvprocedimiento.Rows(i).Cells("Eliminación").Value = True Then
                            cdna_delete = Replace(cdna_delete, "_condicion", dgvprocedimiento.Rows(i).Cells(clmna_inicio - 1).Value.ToString())
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

End Class
