Public Class EcoR
    Inherits Eco
    Public Sub New()
        sqlBuscarPaciente = Consultas.BUSQUEDA_BUSCAR_PACIENTE_ECOCARDIOGRAMAR
        sqlCargarPaciente = Consultas.CARGAR_PACIENTE_ECOCARDIOGRAMAR
        sqlBuscarRegistro = Consultas.BUSQUEDA_BUSCAR_ECOCARDIOGRAMAR
        sqlCargarRegistro = Consultas.CARGAR_ECOCARDIOGRAMAR
        moduloReporte = Constantes.REPORTE_AM
    End Sub
    Public Overrides Sub guardarRegistro()
        If String.IsNullOrEmpty(codigo) Then
            codigo = -1
        Else
            editado = 1
        End If
        EcocardiogramaBLL.guardarEcocardiograma(Me)
    End Sub
    Public Overrides Sub seleccionarDatos()
        Select Case codigoTipoEco
            Case Constantes.CODIGO_ECOCARDIOGRAMA
                nombrePDF = ConstantesHC.NOMBRE_PDF_ECO_R
                vista = "VISTA_REPORTE_ECO"
                rpte = New rptEcocardiograma
            Case Constantes.CODIGO_ECO_ESTRES
                nombrePDF = ConstantesHC.NOMBRE_PDF_ECOSTRES_R
                vista = "VISTA_REPORTE_ECOSTRES"
                rpte = New rptEcoEstres
        End Select
    End Sub


End Class
