Public Class Lista_Manual_D
    Dim cmd As New ManualTarifarioListaDAL
    Public Sub guardar_MANUAL(ByRef txtCodigo As Object, ByVal descripcion As String, opcion As Boolean, manualDuplicar As Integer)
        cmd.guardar_manual(txtCodigo, descripcion, opcion, manualDuplicar)
    End Sub
    Public Sub anular_MANUAL(ByVal codigo As String)
        cmd.anular_MANUAl(codigo)
    End Sub
End Class
