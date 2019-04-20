Public Class Producto
    Public Property codigo As String
    Public Property nombre As String
    Public Property regSanitario As String
    Public Property codigoMarca As Integer
    Public Property codigoInterno As Integer
    Public Property codigoBarras As String
    Public Property codigoCum As String
    Public Property iva As Double
    Public Property Presentacion As String
    Public Property tblBodegas As DataTable
    Sub New()
        tblBodegas = New DataTable
        tblBodegas.Columns.Add("Id", Type.GetType("System.Int32"))
        tblBodegas.Columns.Add("Bodegas", Type.GetType("System.String"))
        tblBodegas.Columns.Add("OK", Type.GetType("System.Boolean"))
    End Sub
End Class
