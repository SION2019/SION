Public Class Pedido_Servicio
    Public Property codigoPedidoPunto As String
    Public Property codigoPedido As String
    Public Property fechaPedido As DateTime
    Public Property areaServicio As String
    Public Property estado As Boolean

    Public Property detallePedido As DataTable
    Sub New()
        detallePedido = New DataTable()
        detallePedido.Columns.Add("Codigo", Type.GetType("System.String"))
        detallePedido.Columns.Add("Descripcion", Type.GetType("System.String"))
        detallePedido.Columns.Add("Cantidad", Type.GetType("System.Int64"))
        detallePedido.Columns.Add("Estado", Type.GetType("System.String"))
        detallePedido.Columns("Cantidad").DefaultValue = 0
    End Sub
End Class
