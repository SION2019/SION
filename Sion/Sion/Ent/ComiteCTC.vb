Public Class ComiteCTC
    Inherits Procedimientos
    Public Property Codigo_Solic As String
    Public Property Codigo_PEM As String
    Public Property Encabezado As String
    Public Property Decision As String
    Public Property dtdiag As DataTable
    Public Property dtAsistentes As DataTable
    Public Property dtMedicamento As DataTable
    Public Property sqlCargarDiagnostico As String
    Public Property moduloReporte As Integer

    Sub New()
        sqlBuscarPaciente = Consultas.BUSQUEDA_BUSCAR_PACIENTE_CTC
        sqlCargarPaciente = Consultas.CARGAR_PACIENTE_CTC
        sqlCargarRegistro = Consultas.CARGAR_CTC
        sqlBuscarRegistro = Consultas.BUSQUEDA_BUSCAR_CTC
        sqlCargarDiagnostico = Consultas.EVO_DIAG_ULTIMO_REG
        sqlGuardarRegistro = Consultas.CREAR_CTC
        sqlAnularRegistro = Consultas.ANULAR_CTC
        area = ConstantesHC.NOMBRE_PDF_CTC
        moduloReporte = Constantes.REPORTE_HC
        titulo = "COMITE TECNICO CIENTIFICO (C.T.C): HISTORIA CLÍNICA"
        '---- Diagnostico
        dtdiag = New DataTable
        '---- Asistentes
        dtAsistentes = New DataTable
        '---- Medicamento
        dtMedicamento = New DataTable

        dtdiag.Columns.Add("Codigo", Type.GetType("System.String"))
        dtdiag.Columns.Add("Descripcion", Type.GetType("System.String"))

        dtAsistentes.Columns.Add("Indice", Type.GetType("System.Int32"))
        dtAsistentes.Columns.Add("Cargo", Type.GetType("System.String"))
        dtAsistentes.Columns.Add("ID_Empleado", Type.GetType("System.Int32"))
        dtAsistentes.Columns.Add("Cedula", Type.GetType("System.String"))
        dtAsistentes.Columns.Add("Nombre", Type.GetType("System.String"))

        dtMedicamento.Columns.Add("Codigo", Type.GetType("System.Int32"))
        dtMedicamento.Columns.Add("Medicamento", Type.GetType("System.String"))
        dtMedicamento.Columns.Add("dosis", Type.GetType("System.String"))
        dtMedicamento.Columns.Add("cant", Type.GetType("System.String"))
        dtMedicamento.Columns.Add("dias", Type.GetType("System.String"))
    End Sub
    Public Overrides Sub guardarRegistro()
        ComiteTecnicoCientificoBLL.guardarCTC(Me)
    End Sub
    Public Sub anularCTC(activoAM As Boolean, activoAF As Boolean)
        Dim params As New List(Of String)
        params.Add(SesionActual.idUsuario)
        params.Add(codigo)
        params.Add(SesionActual.codigoEP)
        params.Add(editado)
        General.ejecutarSQL(sqlAnularRegistro, params)
    End Sub
    Public Sub imprimir()
        Dim ruta, nombreArchivo As String

        nombreArchivo = area & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigo & ConstantesHC.EXTENSION_ARCHIVO_PDF
        ruta = IO.Path.GetTempPath() & nombreArchivo
        Dim reporte As New ftp_reportes
        reporte.crearReportePDF(area, codigo, New rptCTC, codigo,
                                      "{VISTA_CTC.Codigo_CTC}= " & codigo & " and {VISTA_CTC.MODULO}= " & moduloReporte,
                                      area, IO.Path.GetTempPath)
    End Sub
End Class
