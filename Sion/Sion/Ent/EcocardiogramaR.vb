Public Class EcocardiogramaR
    Inherits Ecocardiograma
    Public Sub New()
        titulo = TitulosForm.TITULO_ECO_R
        editado = Constantes.EDITADO
        sqlGuardarRegistro = Consultas.ECOCARDIOGRAMA_CREAR_R
        sqlAnularRegistro = Consultas.ANULAR_ECOCARDIOGRAMAR
        sqlCargarDetalle = Consultas.CARGAR_ECOCARDIOGRAMAR_D
        editado = Constantes.EDITADO
        formula = "{VISTA_REPORTE_ECO.Codigo_Eco} = $" &
               " AND {VISTA_PACIENTES.Id_empresa}=" & idEmpresa &
               " AND {VISTA_REPORTE_ECO.Modulo}= 1"
        area = ConstantesHC.NOMBRE_PDF_ECO_R
    End Sub
End Class
