Public Class EcocardiogramaRR
    Inherits Ecocardiograma
    Public Sub New()
        titulo = TitulosForm.TITULO_ECO_RR
        editado = Constantes.EDITADO
        sqlGuardarRegistro = Consultas.ECOCARDIOGRAMA_CREAR_RR
        sqlAnularRegistro = Consultas.ANULAR_ECOCARDIOGRAMARR
        sqlCargarDetalle = Consultas.CARGAR_ECOCARDIOGRAMARR_D
        editado = Constantes.EDITADO
        formula = "{VISTA_REPORTE_ECO.Codigo_Eco} = $" &
               " AND {VISTA_PACIENTES.Id_empresa}=" & idEmpresa &
               " AND {VISTA_REPORTE_ECO.Modulo}= 2"
        area = ConstantesHC.NOMBRE_PDF_ECO_RR
    End Sub
End Class
