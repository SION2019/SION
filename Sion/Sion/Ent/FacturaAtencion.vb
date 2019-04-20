Public MustInherit Class FacturaAtencion
    Public Property tipo As Integer
    Public Property CodigoContrato As Integer
    Public Property registroAFacturar As Integer
    Public Property codigoFactura As String
    Public Property fechaFactura As DateTime
    Public Property fechaVencimiento As DateTime
    Public Property totalFactura As Double
    Public Property usuario As Integer
    Public Property nombrePDF As String
    Public Property vistaReporte As String
    Public Property objReporte As Object
    Public Property sqlBuscarFactura As String
    Public Property sqlCargarFactura As String
    Public Property sqlAnularFactura As String
    Public Property sqlBuscarContrato As String
    Public Property sqlCargarContrato As String
    Public Property sqlbuscarRegistroAFacturar As String
    Public Property sqlCargarRegistroAFacturar As String
    Public Property visado As Boolean
    Public Property CTC As Boolean
    Public MustOverride Sub cargarDetalle()
    Public MustOverride Sub cargarFactura()
    Public MustOverride Sub guardarFactura()
    Public MustOverride Sub imprimirFactura()
    Public MustOverride Sub calcularFechas()
    Public MustOverride Sub anularFactura()

End Class
