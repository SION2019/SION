Public Class GlucometriaEnfermeria
    Public Property codigoOrden As String
    Public Property fechaOrden As DateTime
    Public Property empresa As String
    Public Property editado As Boolean
    Public Property usuario As Integer
    Public Property usuarioReal As Integer
    Public Property nRegistro As Integer
    Public Property dtGlucomEnfer As DataTable
    Public Property codigoEP As Integer
    Public Property sqlGlucomEnfercarga As String
    Public Property sqlGlucomEnferAnular As String


    Public Sub New()
        sqlGlucomEnfercarga = ConsultasHC.ENFERMERIA_GLUCOMETRIA_CARGAR

        dtGlucomEnfer = New DataTable
        dtGlucomEnfer.Columns.Add("Consecutivo", Type.GetType("System.Int32"))
        dtGlucomEnfer.Columns.Add("HoraDia", Type.GetType("System.String"))
        dtGlucomEnfer.Columns.Add("Glicemia", Type.GetType("System.Int32"))
        dtGlucomEnfer.Columns.Add("UnidadMedida", Type.GetType("System.String"))
        dtGlucomEnfer.Columns.Add("DosisInsulina", Type.GetType("System.String"))
        dtGlucomEnfer.Columns.Add("Iniciales", Type.GetType("System.String"))
        dtGlucomEnfer.Columns.Add("usuario", Type.GetType("System.Int32"))
        dtGlucomEnfer.Columns.Add("Descripcion", Type.GetType("System.String"))
    End Sub

    Public Sub cargarGlucomEnfer()
        Dim params As New List(Of String)
        params.Add(codigoOrden)
        params.Add(empresa)
        General.llenarTabla(sqlGlucomEnfercarga, params, dtGlucomEnfer)
    End Sub
    Public Function cargarIniciales() As String
        Dim params As New List(Of String)
        Dim dfila As DataRow
        params.Add(usuarioReal)
        params.Add(empresa)
        dfila = General.cargarItem(ConsultasHC.EMPLEADO_INICIALES, params)
        Return dfila.Item(0).ToString
    End Function
End Class
