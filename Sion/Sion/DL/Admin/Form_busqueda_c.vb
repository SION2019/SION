Imports System.Data.SqlClient
Public Class Form_busqueda_c

    Public Function llenar(cadena As String, busqueda As String) As DataTable
        Dim dt As New DataTable
        Dim enlce_dta As BindingSource = New BindingSource
        Dim aux As String
        Try
            dt.Clear()

            Using rsltdo = dt.CreateDataReader()
                aux = Replace(cadena, "''", "'" & busqueda & "'")
                aux = Replace(aux, Constantes.COMODIN_DE_BUSQUEDA, "")
                Using adpter = New SqlDataAdapter(aux, FormPrincipal.cnxion)
                    adpter.Fill(dt)
                End Using
            End Using

        Catch ex As Exception
            general.manejoErrores(ex)
            'FormPrincipal.cnxion.Close()

        End Try
        Return dt

    End Function

End Class
