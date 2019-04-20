Public Class GeneralAsis
    Public Shared Sub crearTablaFactura(ByRef dtDetalleFactura As DataTable)
        dtDetalleFactura = New DataTable
        dtDetalleFactura.Columns.Add("Código", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Descripción", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Cantidad", Type.GetType("System.Int32")).DefaultValue = 0
        dtDetalleFactura.Columns.Add("Precio unitario", Type.GetType("System.Double")).DefaultValue = 0
        dtDetalleFactura.Columns.Add("Total", Type.GetType("System.Double")).DefaultValue = 0

    End Sub
    Public Shared Sub crearTablaFacturaMed(ByRef dtDetalleFactura As DataTable)
        dtDetalleFactura = New DataTable
        dtDetalleFactura.Columns.Add("Código", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Código Equivalencia", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Descripción", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Cantidad", Type.GetType("System.Int32")).DefaultValue = 0
        dtDetalleFactura.Columns.Add("Precio unitario", Type.GetType("System.Double")).DefaultValue = 0
        dtDetalleFactura.Columns.Add("Total", Type.GetType("System.Double")).DefaultValue = 0
    End Sub
    Public Shared Sub crearTablaFacturaMedRem(ByRef dtDetalleFactura As DataTable)
        dtDetalleFactura = New DataTable
        dtDetalleFactura.Columns.Add("Código", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Código Producto", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Descripción", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Laboratorio", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Cantidad", Type.GetType("System.Int32")).DefaultValue = 0
        dtDetalleFactura.Columns.Add("IVA", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("ValorIva", Type.GetType("System.Double")).DefaultValue = 0
        dtDetalleFactura.Columns.Add("Precio unitario", Type.GetType("System.Double")).DefaultValue = 0
        dtDetalleFactura.Columns.Add("Total", Type.GetType("System.Double")).DefaultValue = 0
    End Sub
    Public Shared Sub crearTablaFacturaRem(ByRef dtDetalleFactura As DataTable)
        dtDetalleFactura = New DataTable
        dtDetalleFactura.Columns.Add("Código", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Bodega", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Orden", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Creacion", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("ValorBruto", Type.GetType("System.Double"))
        dtDetalleFactura.Columns.Add("IVA", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("ValorTotal", Type.GetType("System.Double"))
        dtDetalleFactura.PrimaryKey = New DataColumn() {dtDetalleFactura.Columns("Código")}
    End Sub
    Public Shared Sub crearTablaConsolidado(ByRef dtDetalleFactura As DataTable)
        dtDetalleFactura = New DataTable
        dtDetalleFactura.Columns.Add("Tipo", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Paciente", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Fecha", Type.GetType("System.DateTime"))
        dtDetalleFactura.Columns.Add("Producto", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Cantidad", Type.GetType("System.String"))
    End Sub

    Public Shared Sub crearTablaFacturaExt(ByRef dtDetalleFactura As DataTable)
        dtDetalleFactura = New DataTable
        dtDetalleFactura.Columns.Add("Código", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Paciente", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Fecha Atención", Type.GetType("System.DateTime"))
        dtDetalleFactura.Columns.Add("Observacion", Type.GetType("System.String"))
        dtDetalleFactura.PrimaryKey = New DataColumn() {dtDetalleFactura.Columns("Código")}
    End Sub

    Public Shared Sub crearTablaFacturaAmb(ByRef dtDetalleFactura As DataTable)
        dtDetalleFactura = New DataTable
        dtDetalleFactura.Columns.Add("PaisO", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("DptoO", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("MpioO", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("EntidadO", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("PaisD", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("DptoD", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("MpioD", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("EntidadD", Type.GetType("System.String"))
        dtDetalleFactura.Columns.Add("Descripcion", Type.GetType("System.String"))
    End Sub

End Class
