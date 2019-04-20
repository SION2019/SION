Public Class ConfigParametroExam
    Public Property CodigoTipo As String
    Public Property dtParamsExam As DataTable
    Public Property usuario As Integer
    Public Sub New()
        dtParamsExam = New DataTable
        dtParamsExam.Columns.Add("Codigo", Type.GetType("System.String"))
        dtParamsExam.Columns.Add("Descripcion", Type.GetType("System.String"))
        dtParamsExam.Columns.Add("Seleccionar", Type.GetType("System.Boolean"))
    End Sub
    Public Sub llenarTabla()
        Dim params As New List(Of String)
        params.Add(CodigoTipo)
        General.llenarTabla(Consultas.PARAMETRO_EXAM_CONFIGURACION_CARGAR, params, dtParamsExam)
    End Sub

End Class
