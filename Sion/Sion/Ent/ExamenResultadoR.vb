Public Class ExamenResultadoR
    Inherits ExamenResultado
    Sub New()
        titulo = TitulosForm.TITULO_RESULT_R
        area = ConstantesHC.NOMBRE_PDF_IMAGENOLOGIA_R
        sqlCargarRegistro = ConsultasHC.RESULTADO_EXAMENES_CARGAR_R
        sqlCargarRegistroDetalle = ConsultasHC.RESULTADO_EXAMENES_CARGAR_D_R
        sqlAnularRegistro = ConsultasHC.RESULTADO_EXAMEN_ANULAR_R
        sqlGuardarRegistro = ConsultasHC.RESULTADO_EXAMEN_GUARDAR_R
        sqlGuardarRegistroHosp = "SP_RESULTADOS_HOSPITAL_EXAMENES_CREAR_R"
        editado = 1
        moduloReporte = Constantes.REPORTE_AM
        nombrePDF = ConstantesHC.NOMBRE_PDF_IMAGENOLOGIA_R
    End Sub
    Public Overrides Sub guardarRegistro()
        Dim objExamen As New ExamenResultadoBLL
        objExamen.guardarResultado(Me)
    End Sub
End Class
