Public Class InformeQx
    Inherits Procedimientos
    Public Property sqlcargarDiagUltimaEvo As String
    Public Property fechaInicio As DateTime
    Public Property fechaFin As DateTime
    Public Property codigoAyudante As Integer
    Private codigoAnestesia As String
    Private codigoAnestesiologo As String
    Public Property codigoVia As Integer
    Public Property hallazgo As String
    Public Property justificacionProcedimiento As String
    Public Property descripcionProcedimiento As String
    Public Property conductaSeguir As String
    Public Property dtDiagPre As DataTable
    Public Property dtDiagpos As DataTable
    Public Property dtMedicamento As DataTable
    Public Property dtProcedimiento As DataTable
    Public Property banderaVerificarSolic As Boolean
    Public Property nombreReporte As String
    Public Property moduloreporte As Integer
    Public Sub New()
        dtDiagpos = New DataTable
        dtDiagPre = New DataTable
        dtMedicamento = New DataTable
        dtProcedimiento = New DataTable

        sqlCargarRegistro = Consultas.CARGAR_QX
        sqlBuscarRegistro = Consultas.BUSQUEDA_BUSCAR_QX
        sqlBuscarPaciente = Consultas.BUSQUEDA_BUSCAR_PACIENTE_QX
        sqlCargarPaciente = Consultas.CARGAR_PACIENTE_QX
        sqlcargarDiagUltimaEvo = Consultas.EVO_DIAG_ULTIMO_REG
        sqlGuardarRegistro = Consultas.QX_CREAR
        sqlAnularRegistro = Consultas.QX_ANULAR
        titulo = TitulosForm.TITULO_INFORME_QX
        sqlCargarInsumoConfig = Consultas.INSUMO_CONF_AUDITORIA_CARGAR
        sqlCargarParaclinicoConfig = Consultas.PARACLINICO_CONF_AUDITORIA_CARGAR

        dtDiagpos.Columns.Add("Codigo", Type.GetType("System.String"))
        dtDiagpos.Columns.Add("Descripcion", Type.GetType("System.String"))

        dtDiagPre.Columns.Add("Codigo", Type.GetType("System.String"))
        dtDiagPre.Columns.Add("Descripcion", Type.GetType("System.String"))

        dtProcedimiento.Columns.Add("Codigo", Type.GetType("System.String"))
        dtProcedimiento.Columns.Add("Descripcion", Type.GetType("System.String"))

        dtMedicamento.Columns.Add("Codigo", Type.GetType("System.String"))
        dtMedicamento.Columns.Add("Medicamento", Type.GetType("System.String"))
        dtMedicamento.Columns.Add("Cantidad", Type.GetType("System.Int32")).DefaultValue = 0
        editado = 0

        area = ConstantesHC.NOMBRE_PDF_INFORMEQX
        nombreReporte = area
        moduloreporte = modulo
    End Sub

    Public Property setGetCodigoMedicamento As String
        Get
            Return codigoAnestesia
        End Get
        Set(value As String)
            codigoAnestesia = If(value = -1, Nothing, value)
        End Set
    End Property
    Public Property setGetCodigoMedico As String
        Get
            Return codigoAnestesiologo
        End Get
        Set(value As String)
            codigoAnestesiologo = If(value = -1, Nothing, value)
        End Set
    End Property

    Public Sub consultaReporte(codigo As Integer, idEmpresa As Integer, modulo As String)
        formula = "{VISTA_INFORME_QX.Codigo_QX}=" & codigo & " and" &
                  "{VISTA_EMPRESAS.Id_empresa}= " & idEmpresa & " and" & "
                  {VISTA_INFORME_QX.MODULO}= " & modulo & ""
    End Sub
    Public Overrides Sub guardarRegistro()
        If String.IsNullOrEmpty(codigo) Then
            codigo = -1
        End If
        InformeQuirurgicoBLL.guardarInformeQx(Me)
    End Sub
    Public Sub anularInformeQX(activoAM As Boolean, activoAF As Boolean)
        Dim params As New List(Of String)
        params.Add(SesionActual.idUsuario)
        params.Add(codigo)
        params.Add(SesionActual.codigoEP)
        params.Add(editado)

        General.ejecutarSQL(sqlAnularRegistro, params)

    End Sub
    Public Overridable Sub imprimirReporte()
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre As String

            codigoNombre = codigo
            nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = Path.GetTempPath() & nombreArchivo
            formula = "{VISTA_INFORME_QX.Codigo_QX} = " & codigoNombre & " AND {VISTA_INFORME_QX.MODULO}=" & moduloreporte & ""
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(nombreReporte, nRegistro, New rptInformeQx,
                                  codigoNombre, formula,
                                  nombreReporte, IO.Path.GetTempPath)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
