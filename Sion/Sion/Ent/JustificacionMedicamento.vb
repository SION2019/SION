Public Class JustificacionMedicamento
    Public Property codigo As String
    Public Property identificacion As String
    Public Property nRegistro As Integer
    Public Property codigoOrden As Integer
    Public Property nombreCompleto As String
    Public Property empresa As String
    Public Property Ingreso As String
    Public Property areaServicio As String
    Public Property medEspecial As String
    Public Property timoMuestra As Boolean

    Public Property fechaMuestra As DateTime
    Public Property fecha As Date
    Public Property tipoMuestra As String
    Public Property aislamiento As String
    Public Property codigoInterno As Integer
    Public Property antibiotico As String
    Public Property tiempoUso As String
    Public Property justificacion As String
    Public Property usuario As Integer
    Public Property nombreUsuario As String
    Public Property dtDiagnosticos As DataTable
    Public Property dtAntibiotico As DataTable
    Public Property codigoEP As Integer
    Public Property diagnosticoSolicitado As String
    Public Property antibioticoTipo As String
    Public Property anular As String
    Public Property medicamento As String
    Public Property dtMedicamentos As New DataTable
    Public Property EPS As String
    Public Property Antibioticos As String
    Public Property diagnosticoEvo As String
    Public Property diagnosticoEvo2 As String
    Public Property dgvMedicamentos As DataGridView
    Public Property dgvAntibioticos As DataGridView
    Public Property guardarAntibiotico As String
    Public Property actualizarAntibiotico As String
    Public Property nombrePDF As String
    Public Property dtDiagnosticosCheck As New DataTable
    Public Property consultaBuscar As String
    Public Property cargarBusqueda As String
    Public Property titulo As String
    Public Property nombreReporte As String
    Public Property moduloreporte As Integer
    Sub New()
        diagnosticoSolicitado = Consultas.JUSTIFICACION_CARGAR_DIAG
        antibioticoTipo = Consultas.JUSTIFICACION_ANT_DETALLE_CARGAR
        anular = Consultas.ANULAR_JUSTIFICACION_ANTIBIOTICO
        medicamento = Consultas.JUSTIFICACION_ANT_MEDICAMENTO_CARGAR
        Antibioticos = Consultas.JUSTIFICACION_ANTIBIOTICO_CARGAR
        diagnosticoEvo = Consultas.DIAGNOSTICO_EVOLUCION_NOPOS_CARGAR
        diagnosticoEvo2 = Consultas.DIAGNOSTICO_EVOLUCION_NOPOS_CARGAR2
        guardarAntibiotico = Consultas.JUSTIFICACION_GUARDAR
        actualizarAntibiotico = Consultas.JUSTIFICACION_ACTUALIZAR
        consultaBuscar = Consultas.JUSTIFICACION_ANT_REGISTRO_BUSCAR
        cargarBusqueda = Consultas.JUSTIFICACION_ANT_BUSCAR
        nombrePDF = ConstantesHC.NOMBRE_PDF_JUSTIFICACION
        titulo = "JUSTIFICACIÓN DE ANTIBIOTICOS: HISTORIA CLÍNICA"
        iniciarTbDiagnosticos()
        iniciarTbAntibiotico()
        nombreReporte = ConstantesHC.NOMBRE_PDF_TRANSFUSION
        moduloreporte = Constantes.REPORTE_HC
    End Sub


    Public Overridable Sub anulaJustificacion()
        Dim lista As New List(Of String)
        lista.Add(codigo)
        lista.Add(SesionActual.idUsuario)
        lista.Add(SesionActual.codigoEP)
        General.ejecutarSQL(anular, lista)
    End Sub
    Public Overridable Sub guardarJustificacion()

        AntibioticoJustificacionBLL.guardarAntibiotico(Me)
    End Sub

    Public Sub cargarDiagnosticosSol()
        Dim params As New List(Of String)
        params.Add(codigo)
        General.llenarTabla(diagnosticoSolicitado, params, dtDiagnosticos)

    End Sub

    Public Sub cargarAntibioticoTipo(pCodigo As Integer)
        Dim param As New List(Of String)
        param.Add(pCodigo)
        General.llenarTabla(antibioticoTipo, param, dtAntibiotico)
        dtAntibiotico.Rows(0).Item(2).Value = dtAntibiotico.Rows(0).Item("Tiempo_uso")
        dtAntibiotico.Rows(0).Item(3).Value = dtAntibiotico.Rows(0).Item("Antibiotico anterior")

    End Sub
    Public Sub cargarDiagnosticosEvo()
        Dim params As New List(Of String)
        params.Add(nRegistro)
        General.llenarTabla(diagnosticoEvo, params, dtDiagnosticos)
        If dtDiagnosticos.Rows.Count = 0 Then
            General.llenarTabla(diagnosticoEvo2, params, dtDiagnosticos)
        End If
    End Sub
    Public Sub cargarDatosPacientes()
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(nRegistro)
        General.llenarTabla(Consultas.JUSTIFICACION_ANT_DATOS_PACIENTES_CARGAR, params, dt)
        If dt.Rows.Count > 0 Then
            nombreCompleto = dt.Rows(0).Item("Paciente")
            identificacion = dt.Rows(0).Item("Documento_paciente")
            EPS = dt.Rows(0).Item("Nombre_EPS")
            areaServicio = dt.Rows(0).Item("Descripcion_Area_Servicio")
        End If
    End Sub

    Public Sub cargarAntibioticos()
        Dim params As New List(Of String)
        params.Add(codigo)
        General.llenarTabla(Antibioticos, params, dtAntibiotico)
    End Sub
    Public Sub cargarMedicamentos()
        Dim param As New List(Of String)
        param.Add(nRegistro)
        dtMedicamentos = New DataTable
        General.llenarTabla(medicamento, param, dtMedicamentos)
    End Sub
    Public Sub iniciarTbDiagnosticos()
        dtDiagnosticos = New DataTable
        dtDiagnosticos.Columns.Add("Codigo", Type.GetType("System.String"))
        dtDiagnosticos.Columns.Add("Codigo Cie", Type.GetType("System.String"))
        dtDiagnosticos.Columns.Add("Descripcion", Type.GetType("System.String"))
        dtDiagnosticos.Columns.Add("Codigo_evo", Type.GetType("System.String"))
        dtDiagnosticos.Columns.Add("Seleccion", Type.GetType("System.Boolean"))
    End Sub

    Public Sub iniciarTbAntibiotico()
        dtAntibiotico = New DataTable
        dtAntibiotico.Columns.Add("Antibiotico anterior", Type.GetType("System.String"))
        dtAntibiotico.Columns.Add("Tiempo uso", Type.GetType("System.String"))
    End Sub


    Public Overridable Sub imprimirReporte()
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre As String


            codigoNombre = codigo
            nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = Path.GetTempPath() & nombreArchivo
            formula = "{VISTA_JUSTIFICACION_ANTIBIOTICO.Codigo} = " & codigoNombre & " AND {VISTA_JUSTIFICACION_ANTIBIOTICO.Modulo}=" & moduloreporte & ""
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(nombreReporte, codigo, New rptJustificacion_Antibioticos,
                                  codigoNombre, formula,
                                  nombreReporte, IO.Path.GetTempPath)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
