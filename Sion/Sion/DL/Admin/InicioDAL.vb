Imports System.Data.SqlClient
Public Class InicioDAL

    Public Sub buscar_actualizaciones()
        Try
            Using consulta = New SqlCommand("SELECT AV.Version FROM A_VERSION AS AV", FormPrincipal.cnxion)
                Using resultado = consulta.ExecuteReader()
                    'If resultado.HasRows = True Then
                    Do While resultado.Read()
                        Inicio.version_actu = CType(resultado.Item("Version"), String)
                    Loop
                    'End If
                End Using
            End Using
        Catch ex As Exception
            general.manejoErrores(ex) 
        End Try
    End Sub

End Class
