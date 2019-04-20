Imports System.Data.SqlClient
Public Class ProductoCualidadBLL
    Dim cmd As New ProductoCualidadDAL
    Public Function guardar_cualidad(ByVal codigo As String, ByVal descripcion As String) As Boolean
        Dim codigo_ent As String = ""

        codigo_ent = cmd.guardar_cualidad(codigo, descripcion)
        If codigo_ent = "" Then
            Return False
        Else
            Cualidades.txtcodigo.Text = codigo_ent
            Return True
        End If
        codigo_ent = Nothing

    End Function
    Public Function anular_cualidad(ByVal codigo As String) As Boolean
        If cmd.anular_cualidad(codigo) = True Then
            Return True
        Else
            Return False
        End If
    End Function

End Class