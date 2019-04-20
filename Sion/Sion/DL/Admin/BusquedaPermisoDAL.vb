Imports System.Data.SqlClient
Public Class BusquedaPermisoDAL
    Public Function buscar_permiso_general(ByVal nombre_form As String, Optional modulo As String = "") As String
        Dim codigo As String = ""
        Dim consulta As String = ""

        If String.IsNullOrEmpty(modulo) Then
            consulta = "SELECT Codigo_Menu FROM  A_MENU WHERE Formulario ='" & nombre_form & "'"
        Else
            consulta = "SELECT Codigo_Menu FROM  A_MENU WHERE Formulario ='" & nombre_form & "' and codigo_menu like '" & modulo & "%'"
        End If
        Using consultor As New SqlCommand(consulta, DBSifaho.Connection.getConn(SesionActual.codigoEnlace))
            Using lector = consultor.ExecuteReader
                If lector.HasRows = True Then
                    lector.Read()
                    codigo = lector.Item("Codigo_Menu")
                End If
            End Using
        End Using
        Return codigo
    End Function
End Class
