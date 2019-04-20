Public Class SeguimientoFacturacion
    Public Property modulo As String
    Public Property dtFacturacion As DataTable
    Public Property fechadesde As Date
    Public Property fechahasta As Date
    Public Property condicion As Integer
    Public Property cadenaFiltro As String
    Public Property Navegador As BindingSource
    Public Sub cargarPacienteFactura()
        Dim params As New List(Of String)
        Dim respuestas As String = Nothing
        If modulo = Constantes.HC Then
            respuestas = Consultas.SEGUIMIENTO_FACTURACION
        ElseIf modulo = Constantes.AM Then
            respuestas = Consultas.SEGUIMIENTO_FACTURACION_R
        ElseIf modulo = Constantes.AF Then
            respuestas = Consultas.SEGUIMIENTO_FACTURACION_RR
        End If
        If modulo <> Nothing Then
            params.Add(Format(fechadesde, Constantes.FORMATO_FECHA_GEN))
            params.Add(Format(fechahasta, Constantes.FORMATO_FECHA_GEN))
            params.Add(condicion)
            General.llenarTabla(respuestas, params, dtFacturacion)
        End If
    End Sub
    Sub New()
        dtFacturacion = New DataTable
        Navegador = New BindingSource
        'dtFacturacion.Columns.Add("No", Type.GetType("System.Int32"))
        ''dtFacturacion.Columns.Add("N_Registro", Type.GetType("System.String"))
        'dtFacturacion.Columns.Add("Paciente", Type.GetType("System.String"))
        'dtFacturacion.Columns.Add("Identificación", Type.GetType("System.String"))
        'dtFacturacion.Columns.Add("Area", Type.GetType("System.String"))
        'dtFacturacion.Columns.Add("EPS", Type.GetType("System.String"))
        'dtFacturacion.Columns.Add("Factura", Type.GetType("System.String"))
        'dtFacturacion.Columns.Add("Valor", Type.GetType("System.Double"))
        'dtFacturacion.Columns.Add("Fecha_Factura", Type.GetType("System.DateTime"))
        'dtFacturacion.Columns.Add("Fecha_Ingreso", Type.GetType("System.DateTime"))
        'dtFacturacion.Columns.Add("Fecha_Egreso", Type.GetType("System.DateTime"))
        'dtFacturacion.Columns.Add("Fecha_Visado", Type.GetType("System.DateTime"))
        'dtFacturacion.Columns.Add("Fecha_Radicación", Type.GetType("System.DateTime"))
        'dtFacturacion.Columns.Add("Días_Radicación", Type.GetType("System.Int32"))
        'dtFacturacion.Columns.Add("Días_Visado", Type.GetType("System.Int32"))
        'dtFacturacion.Columns.Add("Días_Facturado", Type.GetType("System.Int32"))
        'dtFacturacion.Columns.Add("Días_Estancias", Type.GetType("System.Int32"))
        'dtFacturacion.Columns.Add("Total_Días", Type.GetType("System.Int32"))
        'dtFacturacion.Columns.Add("Atención", Type.GetType("System.String"))
        'dtFacturacion.Columns.Add("Glosa_Objetada", Type.GetType("System.DateTime"))
        'dtFacturacion.Columns.Add("Valor_Glosa", Type.GetType("System.Double"))
        Navegador.DataSource = dtFacturacion
    End Sub
End Class
