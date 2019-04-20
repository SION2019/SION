Public Class OxigenoCPAP
    Public Property codigoOxigeno As Integer
    Public Property descripcion As String
    Public Property porcentaje As Double
    Public Property factor As Double
    Public Property Valor As Double

    Sub New()
        codigoOxigeno = -1
        descripcion = ""
        porcentaje = 0
        factor = 0
    End Sub

    Sub cargarMetodo(ByRef codigo As Integer)
        Dim params As New List(Of String)
        params.Add(codigo)
        codigoOxigeno = codigo
        Dim fila As DataRow = General.cargarItem(Consultas.BUSQUEDA_OXIGENO_CPAP, params)
        If Not IsNothing(fila) Then
            descripcion = fila.Item("Descripcion_Oxigeno")
            porcentaje = fila.Item("Porcentaje")
            factor = fila.Item("Factor")
            Valor = fila.Item("Valor")
        End If
    End Sub
End Class
