Public Class Hemodialisis
    Inherits Procedimientos
    Public Property dtSigno As DataTable
    Public Property dtMedicamento As DataTable
    Public Property nota As String
    Public Property resultado As String
    Public Property banderaVerificarSolic As Boolean
    Public Property moduloreporte As Integer
    Public Property nombreReporte As String
    Public Property registro As Integer
    Public Sub New()
        dtSigno = New DataTable
        dtMedicamento = New DataTable
        titulo = "HEMODIALISIS: HISTORIA CLÍNICA"
        area = ConstantesHC.NOMBRE_PDF_HEMODIALISIS
        sqlBuscarPaciente = Consultas.BUSQUEDA_BUSCAR_PACIENTE_HEMO
        sqlCargarPaciente = Consultas.CARGAR_PACIENTE_HEMO
        sqlCargarRegistro = Consultas.CARGAR_HEMO
        sqlBuscarRegistro = Consultas.BUSQUEDA_BUSCAR_HEMODIALISI
        sqlGuardarRegistro = Consultas.CREAR_HEMO
        sqlAnularRegistro = Consultas.HEMOD_ANULAR
        sqlCargarInsumoConfig = Consultas.INSUMO_CONF_AUDITORIA_CARGAR
        sqlCargarParaclinicoConfig = Consultas.PARACLINICO_CONF_AUDITORIA_CARGAR
        moduloreporte = Constantes.REPORTE_HC
        nombreReporte = ConstantesHC.NOMBRE_PDF_TRANSFUSION

        dtSigno.Columns.Add("HORA", Type.GetType("System.String"))
        dtSigno.Columns.Add("TA", Type.GetType("System.String"))
        dtSigno.Columns.Add("FC", Type.GetType("System.String"))
        dtSigno.Columns.Add("FR", Type.GetType("System.String"))
        dtSigno.Columns.Add("TAM", Type.GetType("System.String"))
        dtSigno.Columns.Add("SPO3%", Type.GetType("System.String"))
        dtSigno.Columns.Add("TEMP", Type.GetType("System.String"))

        dtMedicamento.Columns.Add("Codigo", Type.GetType("System.String"))
        dtMedicamento.Columns.Add("Medicamento", Type.GetType("System.String"))
        dtMedicamento.Columns.Add("Cantidad", Type.GetType("System.String"))

    End Sub
    Public Overrides Sub guardarRegistro()
        If String.IsNullOrEmpty(codigo) Then
            codigo = -1
        End If
        HemodialisiBLL.guardarHemodialisi(Me)
    End Sub

    Public Overridable Sub imprimirReporte()
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre As String

            codigoNombre = codigo
            nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = Path.GetTempPath() & nombreArchivo
            formula = "{VISTA_HEMODIALISIS.Codigo_hemo} = " & codigoNombre & " AND {VISTA_HEMODIALISIS.MODULO}=" & moduloreporte & ""
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(nombreReporte, nRegistro, New rptHemodialisis,
                                  codigoNombre, formula,
                                  nombreReporte, IO.Path.GetTempPath)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub anularHemodialisis(activoAM As Boolean, activoAF As Boolean)
        Dim params As New List(Of String)
        params.Add(SesionActual.idUsuario)
        params.Add(codigo)
        params.Add(SesionActual.codigoEP)
        params.Add(editado)

        General.ejecutarSQL(sqlAnularRegistro, params)
    End Sub
End Class
