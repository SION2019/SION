Public Class ProductoPresentacionBLL
    Dim presentacionCmd As New ProductoPresentacionDAL
    Public Sub guardar(ByRef present As Presentacion,
                       ByVal usuario As Integer)
        If present.Codigo = "" Then
            presentacionCmd.guardar(present, usuario)
        Else
            presentacionCmd.actualizar(present, usuario)
        End If
    End Sub
    Public Sub anular(ByRef present As Presentacion,
                       ByVal usuario As Integer)
        presentacionCmd.anular(present, usuario)
    End Sub
End Class
