Public Class ParaclinicoEnfermeria
    Public Property codigoOrden As String
    Public Property dtParaclinicoEnfer As DataTable
    Public Property sqlParacEnfercarga As String

    Public Sub New()
        sqlParacEnfercarga = ConsultasHC.ENFERMERIA_PARACLICOS_CARGAR
        dtParaclinicoEnfer = New DataTable
        dtParaclinicoEnfer.Columns.Add("CodigoTipoExamen", Type.GetType("System.Int32"))
        dtParaclinicoEnfer.Columns.Add("Codigo", Type.GetType("System.String"))
        dtParaclinicoEnfer.Columns.Add("Examen", Type.GetType("System.String"))
        dtParaclinicoEnfer.Columns.Add("Estado", Type.GetType("System.Int32"))
    End Sub

    Public Sub cargarParaclEnfer()
        Dim params As New List(Of String)
        params.Add(codigoOrden)
        General.llenarTabla(sqlParacEnfercarga, params, dtParaclinicoEnfer)
    End Sub
End Class
