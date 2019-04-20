Public Class ConfiguracionAuditoria
    Public Property idConfig As String
    Public Property codigoProcedimiento As String
    Public Property idTerceroEps As Integer
    Public Property codigoAreaServicio As Integer
    Public Property historia As Boolean
    Public Property auditoriaM As Boolean
    Public Property auditoriaF As Boolean
    Public Property paquete As Boolean
    Public Property unaVez As Boolean
    Public Property tipo As String
    Public Property dtInsumo As DataTable
    Public Property dtProcedimiento As DataTable
    Public Property dtParaclinico As DataTable
    Public Property usuario As Integer

    Sub New()
        dtInsumo = New DataTable
        dtInsumo.Columns.Add("Codigo", Type.GetType("System.Int32"))
        dtInsumo.Columns.Add("Cantidad", Type.GetType("System.Int32"))
        dtParaclinico = New DataTable
        dtParaclinico.Columns.Add("Codigo", Type.GetType("System.String"))
        dtParaclinico.Columns.Add("Cantidad", Type.GetType("System.Int32"))
        dtProcedimiento = New DataTable
        dtProcedimiento.Columns.Add("Codigo", Type.GetType("System.String"))
        dtProcedimiento.Columns.Add("Cantidad", Type.GetType("System.Int32"))
    End Sub
End Class
