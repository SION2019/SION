Public Class VistaPreviaR
    Inherits VistaPrevia
    Public Sub New()
        docCarga = ConsultasHC.VISTA_PREVIA_DOC_BD_LISTAR_R
        codigoMenu = Constantes.CODIGO_MENU_AUDM
        listDocumentos = New DataTable
        listDocumentos.Columns.Add("Bandera", Type.GetType("System.String"))
        listDocumentos.Columns.Add("Direccion", Type.GetType("System.String"))
        codigoReporte = Constantes.REPORTE_AM
    End Sub
End Class
