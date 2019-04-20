Public Class Marca
    Public Property codigo As Integer
    Public Property nombre As String

    Sub New()
        codigo = -1
        nombre = ""
    End Sub
    Public Sub cargarMarca(objMarca As Marca)
        Dim fila As DataRow
        Dim params As New List(Of String)
        params.Add(objMarca.codigo)
        fila = General.cargarItem(Consultas.BUSQUEDA_MARCA_INDIVIDUAL, params)
        objMarca.nombre = fila.Item(1)
    End Sub
End Class
