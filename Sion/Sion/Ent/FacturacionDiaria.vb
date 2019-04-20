Public Class FacturacionDiaria
    Public Property modulo As String
    Public Property moduloConf As String
    Public Property bandera As Boolean
    Public Property moduloDetalle As String
    Public Property sqlGuardarMeta As String
    Public Property sqlGuardarConfigDiaFactura As String
    Public Property dtFacturacion As DataTable
    Public Property dtFacturacionDetalle As DataTable
    Public Property dtTotales As DataTable
    Public Property dtTotalesG As DataTable
    Public Property dtAsignarMetas As DataTable
    Public Property dtPromedio As DataTable
    Public Property dtPromedioTG As DataTable
    Public Property dtConfigDiaFactura As DataTable
    Public Property fechaDesde As Date
    Public Property fechaHasta As Date
    Public Property codigoArea As Integer
    Public Property codigoEstado As Integer
    Public Property codigoMeta As Integer
    Public Property usuario As Integer
    Public Property codigoPaciente As Integer
    Public Property condicion As Integer
    Public Property ctc As Boolean
    Public Property cadenaFiltro As Integer
    Public Property navegador As BindingSource
    Public Property navegadorDetalle As BindingSource
    Public Property navegadorTotalesG As BindingSource
    Public Property navegadorPromedio As BindingSource
    Public Property navegadorPromedioTG As BindingSource
    Public Sub cargarPacienteFactura()
        Dim params As New List(Of String)
        params.Add(Format(fechaDesde, Constantes.FORMATO_FECHA_GEN))
        params.Add(Format(fechaHasta, Constantes.FORMATO_FECHA_GEN))
        params.Add(codigoEstado)
        params.Add(condicion)
        params.Add(ctc)
        params.Add(modulo)
        params.Add(SesionActual.codigoEP)
        General.llenarTabla(Consultas.FACTURACION_DIARIA, params, dtFacturacion)
    End Sub
    Public Sub cargarPacienteFacturaDetallado()
        Dim params As New List(Of String)
        params.Add(Format(fechaDesde, Constantes.FORMATO_FECHA_GEN))
        params.Add(Format(fechaHasta, Constantes.FORMATO_FECHA_GEN))
        params.Add(codigoPaciente)
        params.Add(moduloDetalle)
        General.llenarTabla(Consultas.FACTURACION_DIARIA_DETALLADO, params, dtFacturacionDetalle)
    End Sub
    Public Sub cargarMetas()
        Dim params As New List(Of String)
        params.Add(SesionActual.codigoEP)
        General.llenarTabla(Consultas.META_AREA_SERVICIO_CARGAR, params, dtAsignarMetas)
    End Sub
    Public Sub cargarEPSFactura()
        Dim params As New List(Of String)
        params.Add(Format(fechaDesde, Constantes.FORMATO_FECHA_GEN))
        params.Add(Format(fechaHasta, Constantes.FORMATO_FECHA_GEN))
        params.Add(codigoEstado)
        params.Add(condicion)
        params.Add(ctc)
        params.Add(modulo)
        params.Add(SesionActual.codigoEP)
        General.llenarTabla(Consultas.FACTURACION_DIARIA_X_EPS, params, dtFacturacion)
    End Sub
    Public Sub cargarTotales()
        Dim params As New List(Of String)
        params.Add(Format(fechaDesde, Constantes.FORMATO_FECHA_GEN))
        params.Add(Format(fechaHasta, Constantes.FORMATO_FECHA_GEN))
        params.Add(codigoEstado)
        params.Add(ctc)
        General.llenarTabla(Consultas.FACTURACION_DIARIA_TOTAL, params, dtTotales)
        calcularTotales()
    End Sub
    Private Sub calcularTotales()
        If dtTotales.Rows(0).Item("TOTAL") IsNot DBNull.Value Then
            dtTotales.Rows(1).Item("Diferencia") = (dtTotales.Rows(1).Item("Total") - dtTotales.Rows(0).Item("Total"))
            dtTotales.Rows(1).Item("Porcentaje") = Math.Round(((dtTotales.Rows(1).Item("Total") * 100) / dtTotales.Rows(0).Item("Total")) - 100)
            dtTotales.Rows(2).Item("Diferencia") = (dtTotales.Rows(2).Item("Total") - dtTotales.Rows(0).Item("Total"))
            dtTotales.Rows(2).Item("Porcentaje") = Math.Round(((dtTotales.Rows(2).Item("Total") * 100) / dtTotales.Rows(0).Item("Total")) - 100)
        End If
    End Sub
    Public Sub cargarPromedioxEPS()
        Dim params As New List(Of String)
        params.Add(Format(fechaDesde, Constantes.FORMATO_FECHA_GEN))
        params.Add(Format(fechaHasta, Constantes.FORMATO_FECHA_GEN))
        params.Add(codigoArea)
        params.Add(SesionActual.codigoEP)
        General.llenarTabla(Consultas.FACTURACION_PROMEDIO_X_EPS, params, dtPromedio)
    End Sub
    Public Sub cargarDia()
        Dim params As New List(Of String)
        params.Add(Format(fechaDesde, Constantes.FORMATO_FECHA_GEN))
        params.Add(Format(fechaHasta, Constantes.FORMATO_FECHA_GEN))
        params.Add(codigoArea)
        params.Add(SesionActual.codigoEP)
        params.Add(modulo)
        params.Add(cadenaFiltro)
        General.llenarTabla(Consultas.FACTURACION_DIA, params, dtPromedio)
    End Sub
    Public Sub cargarDiaxEPS()
        Dim params As New List(Of String)
        params.Add(Format(fechaDesde, Constantes.FORMATO_FECHA_GEN))
        params.Add(Format(fechaHasta, Constantes.FORMATO_FECHA_GEN))
        params.Add(codigoArea)
        params.Add(SesionActual.codigoEP)
        General.llenarTabla(Consultas.FACTURACION_DIA_X_EPS, params, dtPromedio)
    End Sub
    Public Sub cargarConfigDiaFact()
        Dim params As New List(Of String)
        params.Add("DF")
        params.Add("-" & SesionActual.idUsuario)
        General.llenarTabla(Consultas.CONFIGURACION_DIA_FACTURACION_BUSCAR, params, dtConfigDiaFactura)
    End Sub
    Sub New()
        sqlGuardarMeta = Consultas.AREA_DE_SERVICIO_METAS_CREAR
        sqlGuardarConfigDiaFactura = Consultas.CONFIGURACION_DIA_FACTURACION_CREAR

        dtFacturacion = New DataTable
        navegador = New BindingSource
        dtFacturacion.Columns.Add("No", Type.GetType("System.Int32"))
        dtFacturacion.Columns.Add("Modulo", Type.GetType("System.String"))
        dtFacturacion.Columns.Add("N_Registro", Type.GetType("System.String"))
        dtFacturacion.Columns.Add("Paciente", Type.GetType("System.String"))
        dtFacturacion.Columns.Add("Identificación", Type.GetType("System.String"))
        dtFacturacion.Columns.Add("Tipo_Documento", Type.GetType("System.String"))
        dtFacturacion.Columns.Add("Area", Type.GetType("System.String"))
        dtFacturacion.Columns.Add("EPS", Type.GetType("System.String"))
        dtFacturacion.Columns.Add("DiasEstancia", Type.GetType("System.Int32"))
        dtFacturacion.Columns.Add("Factura", Type.GetType("System.String"))
        dtFacturacion.Columns.Add("E", Type.GetType("System.Double"))
        dtFacturacion.Columns.Add("T", Type.GetType("System.Double"))
        dtFacturacion.Columns.Add("O", Type.GetType("System.Double"))
        dtFacturacion.Columns.Add("P", Type.GetType("System.Double"))
        dtFacturacion.Columns.Add("H", Type.GetType("System.Double"))
        dtFacturacion.Columns.Add("PR", Type.GetType("System.Double"))
        dtFacturacion.Columns.Add("M", Type.GetType("System.Double"))
        dtFacturacion.Columns.Add("I", Type.GetType("System.Double"))
        dtFacturacion.Columns.Add("Valor", Type.GetType("System.Double"))
        dtFacturacion.Columns.Add("PromedioDia", Type.GetType("System.Double"))
        dtFacturacion.Columns.Add("PromedioMeta", Type.GetType("System.Double"))
        dtFacturacion.Columns.Add("Diferencia", Type.GetType("System.Double"))
        dtFacturacion.Columns.Add("Color", Type.GetType("System.Int32"))
        navegador.DataSource = dtFacturacion

        dtFacturacionDetalle = New DataTable
        navegadorDetalle = New BindingSource
        dtFacturacionDetalle.Columns.Add("dia", Type.GetType("System.String"))
        dtFacturacionDetalle.Columns.Add("Modulo", Type.GetType("System.String"))
        dtFacturacionDetalle.Columns.Add("N_Registro", Type.GetType("System.String"))
        dtFacturacionDetalle.Columns.Add("Paciente", Type.GetType("System.String"))
        dtFacturacionDetalle.Columns.Add("Identificación", Type.GetType("System.String"))
        dtFacturacionDetalle.Columns.Add("Tipo_Documento", Type.GetType("System.String"))
        dtFacturacionDetalle.Columns.Add("Area", Type.GetType("System.String"))
        dtFacturacionDetalle.Columns.Add("EPS", Type.GetType("System.String"))
        dtFacturacionDetalle.Columns.Add("DiasEstancia", Type.GetType("System.Int32"))
        dtFacturacionDetalle.Columns.Add("Factura", Type.GetType("System.String"))
        dtFacturacionDetalle.Columns.Add("E", Type.GetType("System.Double"))
        dtFacturacionDetalle.Columns.Add("T", Type.GetType("System.Double"))
        dtFacturacionDetalle.Columns.Add("O", Type.GetType("System.Double"))
        dtFacturacionDetalle.Columns.Add("P", Type.GetType("System.Double"))
        dtFacturacionDetalle.Columns.Add("H", Type.GetType("System.Double"))
        dtFacturacionDetalle.Columns.Add("PR", Type.GetType("System.Double"))
        dtFacturacionDetalle.Columns.Add("M", Type.GetType("System.Double"))
        dtFacturacionDetalle.Columns.Add("I", Type.GetType("System.Double"))
        dtFacturacionDetalle.Columns.Add("Valor", Type.GetType("System.Double"))
        dtFacturacionDetalle.Columns.Add("PromedioDia", Type.GetType("System.Double"))
        navegadorDetalle.DataSource = dtFacturacionDetalle

        dtTotales = New DataTable
        dtTotales.Columns.Add("Modulo", Type.GetType("System.String"))
        dtTotales.Columns.Add("Total", Type.GetType("System.Double"))
        dtTotales.Columns.Add("Diferencia", Type.GetType("System.Double"))
        dtTotales.Columns.Add("Porcentaje", Type.GetType("System.Double"))

        dtTotalesG = New DataTable
        navegadorTotalesG = New BindingSource
        dtTotalesG.Columns.Add("No", Type.GetType("System.Int32"))
        dtTotalesG.Columns.Add("Modulo", Type.GetType("System.String"))
        dtTotalesG.Columns.Add("N_Registro", Type.GetType("System.String"))
        dtTotalesG.Columns.Add("Paciente", Type.GetType("System.String"))
        dtTotalesG.Columns.Add("Identificación", Type.GetType("System.String"))
        dtTotalesG.Columns.Add("Tipo_Documento", Type.GetType("System.String"))
        dtTotalesG.Columns.Add("Area", Type.GetType("System.String"))
        dtTotalesG.Columns.Add("EPS", Type.GetType("System.String"))
        dtTotalesG.Columns.Add("DiasEstancia", Type.GetType("System.Int32"))
        dtTotalesG.Columns.Add("Factura", Type.GetType("System.String"))
        dtTotalesG.Columns.Add("E", Type.GetType("System.Double"))
        dtTotalesG.Columns.Add("T", Type.GetType("System.Double"))
        dtTotalesG.Columns.Add("O", Type.GetType("System.Double"))
        dtTotalesG.Columns.Add("P", Type.GetType("System.Double"))
        dtTotalesG.Columns.Add("H", Type.GetType("System.Double"))
        dtTotalesG.Columns.Add("PR", Type.GetType("System.Double"))
        dtTotalesG.Columns.Add("M", Type.GetType("System.Double"))
        dtTotalesG.Columns.Add("I", Type.GetType("System.Double"))
        dtTotalesG.Columns.Add("Valor", Type.GetType("System.Double"))
        dtTotalesG.Columns.Add("PromedioDia", Type.GetType("System.Double"))
        dtTotalesG.Columns.Add("PromedioMeta", Type.GetType("System.Double"))
        dtTotalesG.Columns.Add("Diferencia", Type.GetType("System.Double"))
        navegadorTotalesG.DataSource = dtTotalesG

        dtAsignarMetas = New DataTable
        dtAsignarMetas.Columns.Add("Codigo_Area_Servicio", Type.GetType("System.String"))
        dtAsignarMetas.Columns.Add("Descripcion_Area_Servicio", Type.GetType("System.String"))
        dtAsignarMetas.Columns.Add("HC", Type.GetType("System.Double"))
        dtAsignarMetas.Columns.Add("AM", Type.GetType("System.Double"))
        dtAsignarMetas.Columns.Add("AF", Type.GetType("System.Double"))

        dtPromedio = New DataTable
        navegadorPromedio = New BindingSource
        dtPromedio.Columns.Add("Area", Type.GetType("System.String"))
        dtPromedio.Columns.Add("Modulo", Type.GetType("System.String"))
        dtPromedio.Columns.Add("N_Pacientes", Type.GetType("System.String"))
        dtPromedio.Columns.Add("N_Registro", Type.GetType("System.String"))
        dtPromedio.Columns.Add("Fecha", Type.GetType("System.String"))
        dtPromedio.Columns.Add("E", Type.GetType("System.Double"))
        dtPromedio.Columns.Add("T", Type.GetType("System.Double"))
        dtPromedio.Columns.Add("O", Type.GetType("System.Double"))
        dtPromedio.Columns.Add("P", Type.GetType("System.Double"))
        dtPromedio.Columns.Add("H", Type.GetType("System.Double"))
        dtPromedio.Columns.Add("PR", Type.GetType("System.Double"))
        dtPromedio.Columns.Add("M", Type.GetType("System.Double"))
        dtPromedio.Columns.Add("I", Type.GetType("System.Double"))
        dtPromedio.Columns.Add("Valor", Type.GetType("System.Double"))
        dtPromedio.Columns.Add("Meta", Type.GetType("System.Double"))
        navegadorPromedio.DataSource = dtPromedio

        dtPromedioTG = New DataTable
        navegadorPromedioTG = New BindingSource
        dtPromedioTG.Columns.Add("Area", Type.GetType("System.String"))
        dtPromedioTG.Columns.Add("Modulo", Type.GetType("System.String"))
        dtPromedioTG.Columns.Add("N_Pacientes", Type.GetType("System.String"))
        dtPromedioTG.Columns.Add("N_Registro", Type.GetType("System.String"))
        dtPromedioTG.Columns.Add("Fecha", Type.GetType("System.String"))
        dtPromedioTG.Columns.Add("E", Type.GetType("System.Double"))
        dtPromedioTG.Columns.Add("T", Type.GetType("System.Double"))
        dtPromedioTG.Columns.Add("O", Type.GetType("System.Double"))
        dtPromedioTG.Columns.Add("P", Type.GetType("System.Double"))
        dtPromedioTG.Columns.Add("H", Type.GetType("System.Double"))
        dtPromedioTG.Columns.Add("PR", Type.GetType("System.Double"))
        dtPromedioTG.Columns.Add("M", Type.GetType("System.Double"))
        dtPromedioTG.Columns.Add("I", Type.GetType("System.Double"))
        dtPromedioTG.Columns.Add("Valor", Type.GetType("System.Double"))
        dtPromedioTG.Columns.Add("Meta", Type.GetType("System.Double"))
        navegadorPromedioTG.DataSource = dtPromedioTG

        dtConfigDiaFactura = New DataTable

    End Sub
End Class
