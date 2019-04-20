Public Class EcoRR
    Inherits Eco
    Public Sub New()
        sqlBuscarPaciente = Consultas.BUSQUEDA_BUSCAR_PACIENTE_ECOCARDIOGRAMARR
        sqlCargarPaciente = Consultas.CARGAR_PACIENTE_ECOCARDIOGRAMARR
        sqlBuscarRegistro = Consultas.BUSQUEDA_BUSCAR_ECOCARDIOGRAMARR
        sqlCargarRegistro = Consultas.CARGAR_ECOCARDIOGRAMARR
        moduloReporte = Constantes.REPORTE_AF
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
                nombrePDF = ConstantesHC.NOMBRE_PDF_ECO_RR
                vista = "VISTA_REPORTE_ECO"
                rpte = New rptEcocardiograma
            Case Constantes.CODIGO_ECO_ESTRES
                nombrePDF = ConstantesHC.NOMBRE_PDF_ECOSTRES_RR
                vista = "VISTA_REPORTE_ECOSTRES"
                rpte = New rptEcoEstres
        End Select
    End Sub
End Class
