Public Class Documentacion
    Inherits ExamenResultado
    Public Sub New()
        titulo = TitulosForm.TITULO_ARCHIVO
        sqlCargarRegistro = ConsultasHC.RESULTADO_DOCUMEN_CARGAR
        sqlAnularRegistro = ConsultasHC.RESULTADO_DOCUMEN_ANULAR
        sqlGuardarRegistro = ConsultasHC.RESULTADO_DOCUMEN_GUARDAR
        sqlCargarRegistroDetalle = ConsultasHC.RESULTADO_DOCUMEN_CARGAR_D
        area = ConstantesHC.NOMBRE_PDF_IMAGENOLOGIA
    End Sub

End Class
