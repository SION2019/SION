Public Class ProductoLineaBLL
    Dim lineasCmd As New ProductoLineaDAL
    Public Sub guardar(ByRef objlinea As Linea,
                       ByVal usuario As Integer)
        If objlinea.codigo = "" Then
            lineasCmd.guardar(objlinea, usuario)
        Else
            lineasCmd.actualizar(objlinea, usuario)
        End If
    End Sub
    Public Sub anular(ByRef objlinea As Linea,
                      ByVal usuario As Integer)
        lineasCmd.anular(objlinea, usuario)
    End Sub
End Class
