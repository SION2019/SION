Public MustInherit Class ManualProcedimientos

    Public Property codigoManual As Integer
    Public Property busqueda As String
    Public Property dtProcedimiento As New DataTable
    Public Property enlceData As New BindingSource
    Public Property usuario As Integer
    Public Property sqlCargarProcedimiento As String

    Public Sub New()

    End Sub
    Public Sub cargarDetalle()
        Dim params As New List(Of String)
        params.Add(busqueda)
        params.Add(codigoManual)
        General.llenarTabla(sqlCargarProcedimiento, params, dtProcedimiento)
        enlceData.DataSource = dtProcedimiento
    End Sub
    Public MustOverride Sub guardarDetalle()



End Class
