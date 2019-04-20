Public Class PrefacturaRR
    Inherits Prefactura
    Public Sub New()
        codigoMenu = Constantes.CODIGO_MENU_AUDF
        consultaCrear = ConsultasHC.PREFACTURA_CREAR_RR
        nombreReporte = ConstantesHC.NOMBRE_PDF_PREFACTURA_RR
        vistaReporte = "VISTA_PREFACTURA"
        objReporte = New rptPrefactura
        moduloReporte = Constantes.REPORTE_AF
        consultaSOAT = ConsultasHC.PREFACTURA_SOAT_RR
    End Sub
End Class
