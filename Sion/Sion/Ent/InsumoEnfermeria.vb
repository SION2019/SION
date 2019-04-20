Public Class InsumoEnfermeria
    Public Property codigoOrden As String
    Public Property fechaOrden As DateTime
    Public Property empresa As String
    Public Property editado As Boolean
    Public Property usuario As Integer
    Public Property usuarioReal As Integer
    Public Property usuarioNombre As String
    Public Property nRegistro As Integer
    Public Property dtInsumosEnfer As DataTable
    Public Property codigoEP As Integer
    Public Property insumosEnfercarga As String
    Public Property insumosEnferAnular As String
    Public Property insumosEnferPeriodicidad As String
    Public Property insumosEnferConfigAud As String
    Sub New()
        insumosEnfercarga = ConsultasHC.ENFERMERIA_INSUMO_CARGAR
        insumosEnferAnular = ConsultasHC.ENFERMERIA_INSUMO_ANULAR
        insumosEnferPeriodicidad = ConsultasHC.ENFERMERIA_INSUMO_PERIODICIDAD
        insumosEnferConfigAud = ConsultasHC.ENFERMERIA_INSUMO_CONFIGURACION
        dtInsumosEnfer = New DataTable
        dtInsumosEnfer.Columns.Add("Código", Type.GetType("System.String"))
        dtInsumosEnfer.Columns.Add("Descripción", Type.GetType("System.String"))
        dtInsumosEnfer.Columns.Add("Cantidad", Type.GetType("System.Int32"))
    End Sub
    Public Sub cargarInsumoEnfer()
        Dim params As New List(Of String)
        params.Add(codigoOrden)
        General.llenarTabla(insumosEnfercarga, params, dtInsumosEnfer)
    End Sub
    Public Sub cargarInsumoFisioConfigAud()
        Dim params As New List(Of String)
        params.Add(nRegistro)
        General.llenarTabla(insumosEnferConfigAud, params, dtInsumosEnfer)
    End Sub
    Public Sub cargarInsumoEnferPeriodicidad()
        Dim params As New List(Of String)
        params.Add(nRegistro)
        params.Add(fechaOrden.Date)
        General.llenarTabla(insumosEnferPeriodicidad, params, dtInsumosEnfer, False)
    End Sub
    Public Overridable Sub anularOrden()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoOrden)
        params.Add(codigoEP)
        General.ejecutarSQL(insumosEnferAnular, params)
    End Sub
End Class
