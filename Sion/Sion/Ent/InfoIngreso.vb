Public MustInherit Class InfoIngreso

    Public Property nRegistro As Integer
    Public Property estadoAtencion As String
    Public Property autorizacion As String
    Public Property institucion As String
    Public Property viaIngreso As String
    Public Property causaExterna As String
    Public Property cama As String
    Public Property pesoUltimo As String
    Public Property peso As String
    Public Property motivo As String
    Public Property signosVitales As String
    Public Property cabezaYCuello As String
    Public Property torax As String
    Public Property cardio As String
    Public Property abdomen As String
    Public Property gentoUrinario As String
    Public Property extremidades As String
    Public Property sistemaNervioso As String
    Public Property analisis As String
    Public Property pronostico As String
    Public Property usuario As String
    Public Property usuarioReal As String
    Public Property usuarioNombre As String
    Public Property usuarioCreacion As String
    Public Property dtDiagImpresion As New DataTable
    Public Property dtDiagRemision As New DataTable
    Public Property codigoEP As Integer
    Public Property sqlDetalleCarga As String
    Public Property sqlDetalleImpresionCarga As String
    Public Property sqlDetalleRemisionCarga As String
    Public Property nombreReporte As String
    Public Property objReporte As Object
    Public Property codEps As String
    Public Property elementoAEliminar As New List(Of String)
    Public Property moduloReporte As Integer
    Public Property vistaReporte As String
    Public Sub New()

    End Sub
    Public MustOverride Sub cargarDetalle()
    Public MustOverride Sub guardarDetalle()
    Public Sub imprimirHoja()
        Try
            Dim nombreArchivo, ruta, codigoNombre, formula As String
            Dim reporte As New ftp_reportes

            'codigoNombre = nRegistro
            'nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            'ruta = Path.GetTempPath() & nombreArchivo
            'ftp_reportes.llamarArchivo(ruta, nombreArchivo, nRegistro, nombreReporte)

            Dim params As New List(Of String)
            params.Add(nRegistro)
            params.Add(moduloReporte)
            codigoNombre = nRegistro
            nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = Path.GetTempPath() & nombreArchivo
            formula = "{" & vistaReporte & ".N_Registro} = " & nRegistro & " AND {" & vistaReporte & ".Modulo}=" & moduloReporte & ""
            reporte.crearReportePDF(nombreReporte, nRegistro, objReporte,
                                  codigoNombre, formula,
                                  nombreReporte, IO.Path.GetTempPath,,, params)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
