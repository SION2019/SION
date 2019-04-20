Public Class BodegaBLL
    Sub establecerGuardado(ByRef objBodega As BodegaE, ByVal usuario As Integer)
        Dim objBodegaC As New Bodegas_C
        If objBodega.codigo = "" Then
            objBodegaC.guardarBodega(objBodega, usuario)
        Else
            objBodegaC.actualizarBodega(objBodega, usuario)
        End If
    End Sub
End Class
