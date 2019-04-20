Public Class InsumoFisioterapia
    Public Property codigoOrden As String
    Public Property fechaOrden As DateTime
    Public Property empresa As String
    Public Property editado As Boolean
    Public Property usuario As Integer
    Public Property usuarioNombre As String
    Public Property usuarioReal As Integer
    Public Property nRegistro As Integer
    Public Property dtInsumosFisio As DataTable
    Public Property codigoEP As Integer
    Public Property insumosFisiocarga As String
    Public Property insumosFisioAnular As String
    Public Property insumosFisioPeriodicidad As String
    Public Property insumosFisioConfigAud As String

    Sub New()
        insumosFisiocarga = ConsultasHC.FISIOTERAPIA_INSUMO_CARGAR
        insumosFisioAnular = ConsultasHC.FISIOTERAPIA_INSUMO_ANULAR
        insumosFisioPeriodicidad = ConsultasHC.FISIOTERAPIA_INSUMO_PERIODICIDAD
        insumosFisioConfigAud = ConsultasHC.FISIOTERAPIA_INSUMO_CONFIGURACION
        dtInsumosFisio = New DataTable
        dtInsumosFisio.Columns.Add("Código", Type.GetType("System.String"))
        dtInsumosFisio.Columns.Add("Descripción", Type.GetType("System.String"))
        dtInsumosFisio.Columns.Add("Cantidad", Type.GetType("System.Int32"))
    End Sub
    Public Sub cargarInsumoFisio()
        Dim params As New List(Of String)
        params.Add(codigoOrden)
        General.llenarTabla(insumosFisiocarga, params, dtInsumosFisio)
    End Sub
    Public Sub cargarInsumoFisioConfigAud()
        Dim params As New List(Of String)
        params.Add(nRegistro)
        General.llenarTabla(insumosFisioConfigAud, params, dtInsumosFisio)
    End Sub
    Public Sub cargarInsumoFisioPeriodicidad()
        Dim params As New List(Of String)
        params.Add(nRegistro)
        params.Add(fechaOrden.Date)
        General.llenarTabla(insumosFisioPeriodicidad, params, dtInsumosFisio, False)
    End Sub
    Public Overridable Sub anularOrden()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoOrden)
        params.Add(codigoEP)
        General.ejecutarSQL(insumosFisioAnular, params)
    End Sub
End Class
