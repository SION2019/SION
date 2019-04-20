Public Class VistaPreviaRR
    Inherits VistaPrevia
    Public Sub New()
        docCarga = ConsultasHC.VISTA_PREVIA_DOC_BD_LISTAR_RR
        codigoMenu = Constantes.CODIGO_MENU_AUDF
        listDocumentos = New DataTable
        listDocumentos.Columns.Add("Bandera", Type.GetType("System.String"))
        listDocumentos.Columns.Add("Direccion", Type.GetType("System.String"))
        codigoReporte = Constantes.REPORTE_AF
    End Sub
End Class
