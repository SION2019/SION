Imports System.Data.SqlClient
Public Class ArticuloDAL
    Public Function llenar_dgv(TextBox1 As String) As DataTable
        Dim dt As New DataTable
        Dim cdna, cadena As String
        Dim adpter As SqlDataAdapter
        Dim enlce_dta As BindingSource = New BindingSource

        cdna = "" : cadena = ""
        cdna = "EXEC [SP_ARTICULO_CARGAR] '" & TextBox1 & "'"

        Try
            dt.Clear()
            adpter = New SqlDataAdapter(cdna, FormPrincipal.cnxion)
            adpter.Fill(dt) : adpter.Dispose()
        Catch ex As Exception : MsgBox("Error " + ex.Message)
        Finally

        End Try
        Return dt
    End Function

    Public Function guardar(ByRef articulosoat As ArticuloSoat) As Boolean
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    If articulosoat.dt_update.Rows.Count > 0 Then
                        For k = 0 To articulosoat.dt_update.Rows.Count - 1
                            consulta.CommandText = articulosoat.dt_update.Rows(k).Item(0).ToString
                            consulta.ExecuteNonQuery()
                        Next
                    ElseIf articulosoat.indcdor > 0 Then

                        consulta.CommandText = articulosoat.cdna_insert
                        consulta.ExecuteNonQuery()
                    ElseIf articulosoat.dt_delete.Rows.Count > 0 Then
                        For k = 0 To articulosoat.dt_delete.Rows.Count - 1
                            consulta.CommandText = articulosoat.dt_delete.Rows(k).Item(0).ToString
                            consulta.ExecuteNonQuery()
                        Next
                    End If
                    If articulosoat.dt_update.Rows.Count > 0 Or articulosoat.indcdor > 0 Or articulosoat.dt_delete.Rows.Count > 0 Then
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

End Class
