Public Class TerapiaFisicaYRespiratoria
    Public Property codigoOrden As Integer
    Public Property codigoProcedimiento As String
    Public Property empresa As String
    Public Property usuario As Integer
    Public Property dtTerapiaFyR As DataTable
    Public Property codigoEP As Integer
    Public Property terapiaFyRCarga As String
    Sub New()
        terapiaFyRCarga = ConsultasHC.FISIOTERAPIA_FISICA_Y_RESPIRATORIA_CARGAR
        dtTerapiaFyR = New DataTable
        dtTerapiaFyR.Columns.Add("Codigo_Procedimiento", Type.GetType("System.String"))
        dtTerapiaFyR.Columns.Add("Descripcion", Type.GetType("System.String"))
        dtTerapiaFyR.Columns.Add("Realizado", Type.GetType("System.Boolean"))
    End Sub
    Public Sub cargarTerapiaFyR()
        Dim params As New List(Of String)
        params.Add(codigoOrden)
        params.Add(codigoProcedimiento)
        General.llenarTabla(terapiaFyRCarga, params, dtTerapiaFyR)
    End Sub

End Class
