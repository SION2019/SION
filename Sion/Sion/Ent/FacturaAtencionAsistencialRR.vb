Public Class FacturaAtencionAsistencialRR
    Inherits FacturaAtencionAsistencial

    Public Overrides Sub consultasCargarDatos()
        sqlPrecargarEstancias = ConsultasAsis.FACTURA_ESTANCIAS_CARGARRR
        sqlPrecargarOxigenos = ConsultasAsis.FACTURA_OXIGENO_CARGARRR
        sqlPrecargarParaclinicos = ConsultasAsis.FACTURA_PARACLINICOS_CARGARRR
        sqlPrecargarHemoderivados = ConsultasAsis.FACTURA_HEMODERIVADOS_CARGARRR
        sqlPrecargarProcedimientos = ConsultasAsis.FACTURA_PROCEDIMIENTOS_CARGARRR
        sqlPrecargarMedicamentos = ConsultasAsis.FACTURA_MEDICAMENTOS_CARGARRR
        sqlPrecargarMedicamentosCTC = ConsultasAsis.FACTURA_CTC_MEDICAMENTOS_CARGARRR
        sqlPrecargarInsumos = ConsultasAsis.FACTURA_INSUMOS_CARGARRR
        sqlBuscarContrato = ConsultasAsis.FACTURA_ATENCION_CLIENTE_BUSCARRR
        sqlbuscarRegistroAFacturar = ConsultasAsis.FACTURA_ATENCION_REGISTRO_A_FACTURAR_BUSCARRR
        sqlCargarRegistroAFacturar = ConsultasAsis.FACTURA_ATENCION_REGISTRO_A_FACTURAR_CARGARRR
    End Sub
    Public Overrides Sub nombreReportePDF()
        If CTC = True Then
            nombrePDF = ConstantesHC.NOMBRE_PDF_FACTURA_CTC_RR
            tipo = 1
        Else
            nombrePDF = ConstantesHC.NOMBRE_PDF_FACTURA_RR
            tipo = 0
        End If
        nombreCertificadoPDF = ConstantesHC.NOMBRE_PDF_FACTURA_CERTIFICADO_RR
        moduloReporte = Constantes.REPORTE_AF
    End Sub
End Class
