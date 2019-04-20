Public Class ProductoMarcaBLL
    Dim objMarcaDAL As New ProductoMarcaDAL
    Public Sub guardarMarca(ByRef objProductoMarca As Marca,
                             ByRef usuario As Integer)

        If objProductoMarca.codigo = -1 Then
            objMarcaDAL.guardarMarca(objProductoMarca, usuario)
        Else
            objMarcaDAL.actualizarMarca(objProductoMarca, usuario)
        End If
    End Sub
    Public Sub anularMarca(ByRef objProductoMarca As Marca,
                           ByRef usuario As Integer)
        objMarcaDAL.anularMarca(objProductoMarca, usuario)
    End Sub

End Class
