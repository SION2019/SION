Public Class ProductoBLL
    Dim objProductoDal As New ProductoDAL
    Public Sub guardar(ByRef objProducto As Producto, ByVal usuario As Integer)
        If objProducto.codigo = "" Then
            objProductoDal.guardar(objProducto, usuario)
        Else
            objProductoDal.actualizar(objProducto, usuario)
        End If
    End Sub
    Public Function verificar_asignacion(ByVal codigo_producto As String, ByVal codigo_bodega As Integer) As Boolean
        Return objProductoDal.verificar_asignacion(codigo_producto, codigo_bodega)
    End Function
    Public Function anular_Producto(ByVal obj As Producto, ByVal usuario As Integer) As Boolean
        Return objProductoDal.anular_Producto(obj, usuario)
    End Function
    Public Function verificar_existencia_producto(ByVal Codigo_producto As String, ByVal Codigo_Bodega As String) As Boolean
        Return objProductoDal.verificar_existencia_producto(Codigo_producto, Codigo_Bodega)
    End Function
End Class
