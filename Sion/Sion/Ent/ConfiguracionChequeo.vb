Public Class ConfiguracionChequeo

    Public Property arbol As New TreeView
    Public Property codigoChequeo As Integer
    Public Property exepcion As Boolean
    Public Property dtConfiguracion As DataTable
    Public Property dsLista As New DataSet
    Public Property dtProcedimientos As DataTable
    Public Property dtExepciones As DataTable
    Public Property codigoProcedimiento As String
    Public Property descripcionProcedimiento As String
    Public Property codigoPadre As String
    Public Property busqueda As String
    Public Sub cargarConfiguracion()
        General.llenarTabla(Consultas.CONFIGURACION_CHEQUEO_CARGAR, Nothing, dtConfiguracion)
    End Sub
    Public Sub cargarProcedimientos()
        Dim params As New List(Of String)
        params.Add(codigoChequeo)
        params.Add(busqueda)
        General.llenarTabla(Consultas.CARGAR_PROCEDIMIENTOS_LISTA, params, dtProcedimientos)
    End Sub

    Public Sub exepcionQuitar()
        ConfiguracionChequeoDAL.ChequeoQuitar(Me)
    End Sub

    Public Sub cargarProcedimientosExepcion()
        Dim params As New List(Of String)
        params.Add(codigoChequeo)
        params.Add(busqueda)
        General.llenarTabla(Consultas.CARGAR_PROCEDIMIENTO_EXEPCION, params, dtExepciones)
    End Sub
    Public Sub guardarCheck()
        ConfiguracionChequeoDAL.guardarChequeo(Me)
    End Sub
    Public Sub guardarProcedimientos()
        ConfiguracionChequeoDAL.guardarChequeoProcedimiento(Me)
    End Sub
    Sub New()
        dtConfiguracion = New DataTable
        dtConfiguracion.Columns.Add("Código", Type.GetType("System.String"))
        dtConfiguracion.Columns.Add("Descripcion", Type.GetType("System.String"))
        dtConfiguracion.Columns.Add("Ex Global", Type.GetType("System.Boolean"))
        dtConfiguracion.Columns.Add("Exepciones", Type.GetType("System.String"))

        dtProcedimientos = New DataTable
        dtProcedimientos.Columns.Add("Codigo Procedimiento", Type.GetType("System.String"))
        dtProcedimientos.Columns.Add("Descripcion", Type.GetType("System.String"))
        dtProcedimientos.Columns.Add("Codigo Padre", Type.GetType("System.String"))

        dtExepciones = New DataTable
        dtExepciones.Columns.Add("Codigo Procedimiento", Type.GetType("System.String"))
        dtExepciones.Columns.Add("Descripción", Type.GetType("System.String"))

    End Sub
End Class
