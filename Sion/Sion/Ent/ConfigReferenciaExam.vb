Public Class ConfigReferenciaExam
    Public Property CodigoTipo As String
    Public Property CodigoGenero As Integer
    Public Property dtReferenciaExam As DataTable
    Public Property usuario As Integer
    Public Sub New()
        dtReferenciaExam = New DataTable
        dtReferenciaExam.Columns.Add("Codigo", Type.GetType("System.String"))
        dtReferenciaExam.Columns.Add("Descripcion", Type.GetType("System.String"))
        dtReferenciaExam.Columns.Add("Referencia", Type.GetType("System.String"))
        dtReferenciaExam.Columns.Add("Unidad", Type.GetType("System.String"))
    End Sub
    Public Sub llenarTabla()
        Dim params As New List(Of String)
        params.Add(CodigoTipo)
        params.Add(CodigoGenero)
        General.llenarTabla(Consultas.REFERENCIA_EXAMEN_CONF_CARGAR, params, dtReferenciaExam)
    End Sub

End Class
