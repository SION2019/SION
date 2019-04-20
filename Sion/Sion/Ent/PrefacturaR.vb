Public Class PrefacturaR
    Inherits Prefactura
    Public Sub New()
        codigoMenu = Constantes.CODIGO_MENU_AUDM
        consultaCrear = ConsultasHC.PREFACTURA_CREAR_R
        nombreReporte = ConstantesHC.NOMBRE_PDF_PREFACTURA_R
        vistaReporte = "VISTA_PREFACTURA"
        objReporte = New rptPrefactura
        moduloReporte = Constantes.REPORTE_AM
        consultaSOAT = ConsultasHC.PREFACTURA_SOAT_R
    End Sub
End Class
