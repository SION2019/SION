Public Class HemoderivadoEnfermeria
    Public Property codigoOrden As Integer
    Public Property dtHemoderivadoEnfer As DataTable
    Public Property sqlHemodEnfercarga As String

    Public Sub New()
        sqlHemodEnfercarga = ConsultasHC.ENFERMERIA_HEMODERIVADOS_CARGAR
        dtHemoderivadoEnfer = New DataTable
        dtHemoderivadoEnfer.Columns.Add("CodigoTipoExamenH", Type.GetType("System.Int32"))
        dtHemoderivadoEnfer.Columns.Add("CodigoH", Type.GetType("System.String"))
        dtHemoderivadoEnfer.Columns.Add("ExamenH", Type.GetType("System.String"))
        dtHemoderivadoEnfer.Columns.Add("EstadoH", Type.GetType("System.Int32"))
    End Sub

    Public Sub cargarHemodEnfer()
        Dim params As New List(Of String)
        params.Add(codigoOrden)
        General.llenarTabla(sqlHemodEnfercarga, params, dtHemoderivadoEnfer)
    End Sub
End Class
