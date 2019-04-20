Imports System.Data.SqlClient
Public Class IssDAL

    Public Function guardar(ByRef dgv As DataGridView, ByVal codigoRef As String) As Boolean
        Dim i, j, l, indcdor, clmna_inicio, clmna_fin As Integer
        Dim cdna, cadena, cdna_insert, cdna_update, cdna_delete, valor As String
        Dim dt_update, dt_delete As New DataTable
        Dim bndra As Boolean = True

        clmna_inicio = 4 : clmna_fin = 14
        cdna_insert = "INSERT INTO M_PROCEDIMIENTOS_ISS(Codigo_Manual,Codigo_ISS,Descripcion,UVR,VALOR,Valor_Personalizado,CODIGO_GRUPO,CODIGO_CAPITULO,CODIGO_EQUIPO,FACTURABLE,Usuario_creacion,Fecha_creacion) VALUES(Codigo_Manu4l,"
        cdna_update = "UPDATE M_PROCEDIMIENTOS_ISS SET Codigo_ISS='_valor1',Descripcion='_valor2',UVR='_valor3',VALOR='_valor4',Valor_Personalizado='_valor5',CODIGO_GRUPO='_valor6',CODIGO_CAPITULO='_valor8',CODIGO_EQUIPO='_valor00',FACTURABLE='_valor01',Usuario_actualizacion=" & SesionActual.idUsuario & ",Fecha_actualizacion=getdate()  WHERE Codigo_ISS='_condicion' and Codigo_Manual=" & codigoRef & ""
        cdna_delete = "DELETE M_PROCEDIMIENTOS_ISS WHERE Codigo_ISS='_condicion' and Codigo_Manual='" & codigoRef & "'"
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    For i = 0 To dgv.RowCount - 1

                        If (IsDBNull(dgv.Rows(i).Cells("Inserción").Value) = True Or dgv.Rows(i).Cells("Código").Value.ToString = "") And i <> dgv.RowCount - 1 Then
                            MsgBox("error al validar dgv, falta información", MsgBoxStyle.Critical)
                            Return False
                            Exit For
                        ElseIf (IsDBNull(dgv.Rows(i).Cells("Inserción").Value) = True Or dgv.Rows(i).Cells("Código").Value.ToString = "") And i = dgv.RowCount - 1 Then
                            Exit For
                        End If
                        If dgv.Rows(i).Cells("Inserción").Value = True Then
                            If dgv.Rows(i).Cells("UVR").Value.ToString = "" Then
                                dgv.Rows(i).Cells("UVR").Value = "0"
                                dgv.EndEdit()
                            End If
                            If dgv.Rows(i).Cells("Valor").Value.ToString = "" Then
                                dgv.Rows(i).Cells("Valor").Value = "0"
                                dgv.EndEdit()
                            End If
                            If dgv.Rows(i).Cells("Equipo").Value.ToString = "" Then
                                dgv.Rows(i).Cells("Equipo").Value = "0"
                                dgv.EndEdit()
                            End If
                            bndra = Validar_dgv(i, clmna_inicio, clmna_fin, dgv)
                            If bndra Then

                                Return False
                            Else
                                If indcdor <> 0 Then
                                    cdna_insert = cdna_insert & ",(Codigo_Manu4l,"
                                End If
                                For j = clmna_inicio To clmna_fin
                                    If dgv.Rows(i).Cells(j).ValueType.ToString = "System.String" Or dgv.Rows(i).Cells(j).ValueType.ToString = "System.Int32" Then
                                        valor = Replace(dgv.Rows(i).Cells(j).Value.ToString(), "'", "")
                                    Else
                                        If IsDBNull(dgv.Rows(i).Cells(j).Value) Then
                                            valor = "False"
                                        Else
                                            valor = dgv.Rows(i).Cells(j).Value
                                        End If
                                    End If
                                    If j <> 10 And j <> 12 Then
                                        If j = 6 Or j = 7 Or j = 8 Then
                                            If valor = "False" Then
                                                valor = 0
                                            End If
                                        End If
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
                                cdna_insert = cdna_insert & "," & SesionActual.idUsuario & ",getdate())"
                                cdna_insert = Replace(cdna_insert, "Codigo_Manu4l", codigoRef)
                                indcdor = indcdor + 1
                            End If
                        ElseIf dgv.Rows(i).Cells("Edición").Value = True Then
                            If dgv.Rows(i).Cells("UVR").Value.ToString = "" Then
                                dgv.Rows(i).Cells("Factor").Value = "0"
                                dgv.EndEdit()
                            End If
                            If dgv.Rows(i).Cells("Valor").Value.ToString = "" Then
                                dgv.Rows(i).Cells("Valor").Value = "0"
                                dgv.EndEdit()
                            End If
                            If dgv.Rows(i).Cells("Equipo").Value.ToString = "" Then
                                dgv.Rows(i).Cells("Equipo").Value = "0"
                                dgv.EndEdit()
                            End If
                            bndra = Validar_dgv(i, clmna_inicio, clmna_fin, dgv)
                            If bndra Then

                                Return False
                            Else
                                l = 1
                                cdna = cdna_update
                                For j = clmna_inicio To clmna_fin
                                    If dgv.Rows(i).Cells(j).ValueType.ToString = "System.String" Or dgv.Rows(i).Cells(j).ValueType.ToString = "System.Int32" Then
                                        valor = Replace(dgv.Rows(i).Cells(j).Value.ToString(), "'", "")
                                    Else
                                        If IsDBNull(dgv.Rows(i).Cells(j).Value) Then
                                            valor = "False"
                                        Else
                                            valor = dgv.Rows(i).Cells(j).Value
                                        End If
                                    End If
                                    cadena = "_valor" & CStr(l)
                                    If cadena = "_valor10" Then
                                        cadena = "_valor00"
                                    ElseIf cadena = "_valor11" Then
                                        cadena = "_valor01"
                                    End If
                                    valor = dgv.Rows(i).Cells(j).Value.ToString().Trim
                                    cdna = Replace(cdna, cadena, valor)

                                    l = l + 1
                                Next
                                cdna = Replace(cdna, "_condicion", dgv.Rows(i).Cells(clmna_inicio - 1).Value.ToString().Trim)
                                If dt_update.Columns.Count <= 0 Then
                                    dt_update.Columns.Add("Consulta")
                                End If
                                dt_update.Rows.Add()
                                dt_update.Rows(dt_update.Rows.Count - 1).Item(0) = cdna
                            End If
                        ElseIf dgv.Rows(i).Cells("Eliminación").Value = True Then
                            cdna_delete = Replace(cdna_delete, "_condicion", dgv.Rows(i).Cells(clmna_inicio - 1).Value.ToString().Trim)
                            If dt_delete.Columns.Count <= 0 Then
                                dt_delete.Columns.Add("Consulta")
                            End If
                            dt_delete.Rows.Add()
                            dt_delete.Rows(dt_delete.Rows.Count - 1).Item(0) = cdna_delete
                        End If
                    Next
                    Try
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
                    Catch ex As Exception
                        General.manejoErrores(ex)
                    End Try

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
    Private Function Validar_dgv(fila_actual As Integer, clmna_inicio As Integer, clmna_fin As Integer, ByRef dgv As DataGridView) As Boolean
        Dim bndra As Boolean
        Dim i As Integer
        Dim fila As String

        bndra = False
        fila = dgv.Rows(fila_actual).Cells(8).Value.ToString

        For i = clmna_inicio To clmna_fin
            If dgv.Rows(fila_actual).Cells(i).ValueType = Type.GetType("System.String") Then
                If String.IsNullOrEmpty(dgv.Rows(fila_actual).Cells(i).Value.ToString) Or String.IsNullOrWhiteSpace(dgv.Rows(fila_actual).Cells(i).Value.ToString) Then
                    bndra = True
                    MsgBox("Hace falta un dato", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Advertencia")
                    dgv.Rows(fila_actual).Cells(i).Selected = True
                    Exit For
                End If
            End If
        Next
        Return bndra
    End Function


End Class
