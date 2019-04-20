

Public Class Admision
    Public Property nRegistro As String
    Public Property identiPaciente As Integer
    Public Property fechaAdmision As DateTime
    Public Property CodigoEp As Integer
    Public Property codigoEspecialidad As Integer
    Public Property codigoTriage As Integer
    Public Property codigoContrato As Integer
    Public Property codigoArea As Integer
    Public Property idContacto As Integer
    Public Property codigoIdAcompanate As Integer
    Public Property idAcompanate As String
    Public Property acompanante As String
    Public Property Direccion As String
    Public Property telAcompanate As String
    Public Property codigoPais As String
    Public Property codigoDepartamento As String
    Public Property codigoMunicipio As String
    Public Property codigoIdResponsable As Integer
    Public Property idResponsable As String
    Public Property responsable As String
    Public Property direccionResponsable As String
    Public Property telResponsable As String
    Public Property codigoPaisResponsable As String
    Public Property codigoDepartamentoResponsable As String
    Public Property codigoMunicipioResponsable As String
    Public Property usuario As String
    Public Property numAutorizacion As String
    Public Property viaIngreso As Integer
    Public Property institucion As Integer
    Public Property causaExterna As Integer
    Public Property codigoEstado As Integer
    Public Property codigoCama As Integer
    Public Property dtDiagnostico As DataTable
    Public Property dtTraslado As DataTable
    Public Property tipoDocumento As Integer
    Public Property opcionAcomp As Boolean
    Public Property opcionRespon As Boolean
    Public Property imagen As Byte()
    Public Property extensionDoc As String
    Public Property observacion As String
    Public Property solicitud As String
    Public Property nombreReporte As String
    Public Property documentoPaciente As String
    Public Property valorSOAT As Integer
    Public Property tipo As String

    Sub New()
        dtDiagnostico = New DataTable
        dtDiagnostico.Columns.Add("N_Registro", Type.GetType("System.String"))
        dtDiagnostico.Columns.Add("Codigo_CIE", Type.GetType("System.String"))
        dtDiagnostico.Columns.Add("Observacion", Type.GetType("System.String"))
        nombreReporte = ConstantesHC.NOMBRE_PDF_ADMISION
    End Sub
    Public Overridable Sub imprimirReporte()
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre As String

            codigoNombre = nRegistro
            nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = Path.GetTempPath() & nombreArchivo
            formula = "{VISTA_ADMISION.N_Registro} = " & nRegistro
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(nombreReporte, nRegistro, New rpt_Admisiones,
                                  codigoNombre, formula,
                                  nombreReporte, IO.Path.GetTempPath)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub imprimirCertificado()
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre As String

            codigoNombre = nRegistro
            nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = Path.GetTempPath() & nombreArchivo
            formula = "{VISTA_CERTIFICADO_EPS_VALOR.N_Registro} = " & nRegistro
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(nombreReporte, nRegistro, New CertificadoSoat,
                                  codigoNombre, formula,
                                  nombreReporte, IO.Path.GetTempPath)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
