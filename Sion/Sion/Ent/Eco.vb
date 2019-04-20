Public Class Eco
    Inherits Procedimientos
    Public Property codigoMenu As String
    Public Property codigoTipoEco As String
    Public Property objEcocardiograma As Ecocardiograma
    Public Property objEcocardiogramaEstres As EcocardiogramaEstres
    Public Property descripcionHallazgo As String
    Public Property rpte As ReportClass
    Public Property moduloReporte As Integer
    Public Property nombrePDF As String
    Public Property vista As String
    Public Sub New()
        '----- Consultas 
        sqlBuscarPaciente = Consultas.BUSQUEDA_BUSCAR_PACIENTE_ECOCARDIOGRAMA
        sqlCargarPaciente = Consultas.CARGAR_PACIENTE_ECOCARDIOGRAMA
        sqlBuscarRegistro = Consultas.BUSQUEDA_BUSCAR_ECOCARDIOGRAMA
        sqlCargarRegistro = Consultas.CARGAR_ECOCARDIOGRAMA
        moduloReporte = Constantes.REPORTE_HC

    End Sub
    Public Sub instanciaObjeto(modulo As String)
        '----- iniciar las clase segun eñ modulo 
        codigotipo()
        objEcocardiograma = GeneralHC.fabricaHC.crear(Constantes.CODIGO_ECOCARDIOGRAMA & Constantes.ECO & modulo)
        objEcocardiogramaEstres = GeneralHC.fabricaHC.crear(Constantes.CODIGO_ECO_ESTRES & Constantes.ECOSTRES & modulo)
    End Sub
    Public Function validarFormulario() As String
        '------valida el tirulo del formulario segun el modulo y el tipo de examen 
        Select Case codigoTipoEco
            Case Constantes.CODIGO_ECOCARDIOGRAMA
                titulo = objEcocardiograma.titulo
            Case Constantes.CODIGO_ECO_ESTRES
                titulo = objEcocardiogramaEstres.titulo
        End Select
        Return titulo
    End Function
    Public Overrides Sub guardarRegistro()
        If String.IsNullOrEmpty(codigo) Then
            codigo = Constantes.VALOR_PREDETERMINADO '--------validamos que el codigo este sin registro para poder generar uno nuevo 
        End If
        EcocardiogramaBLL.guardarEcocardiograma(Me)
    End Sub

    Public Sub MostrarReporte()
        '----------validamos el area segun el tipo de examen de ecocardiograma 
        Select Case codigoTipoEco
            Case Constantes.CODIGO_ECOCARDIOGRAMA
                area = objEcocardiograma.area
            Case Constantes.CODIGO_ECO_ESTRES
                area = objEcocardiogramaEstres.area
        End Select
    End Sub
    Public Sub codigotipo()
        '---- consultamos el codigo del examen eco
        Try
            If codigoMenu = Constantes.CODIGO_MENU_ECOSTRES_HC _
                                Or codigoMenu = Constantes.CODIGO_MENU_ECOSTRES_AM _
                                Or codigoMenu = Constantes.CODIGO_MENU_ECOSTRES_AF Then
                codigoTipoEco = Constantes.CODIGO_ECO_ESTRES
            Else
                codigoTipoEco = Constantes.CODIGO_ECOCARDIOGRAMA
            End If
            seleccionarDatos()
        Catch ex As Exception
            general.manejoErrores(ex)
        End Try
    End Sub
    Public Overridable Sub seleccionarDatos()
        Select Case codigoTipoEco
            Case Constantes.CODIGO_ECOCARDIOGRAMA
                nombrePDF = ConstantesHC.NOMBRE_PDF_ECO
                vista = "VISTA_REPORTE_ECO"
                rpte = New rptEcocardiograma
            Case Constantes.CODIGO_ECO_ESTRES
                nombrePDF = ConstantesHC.NOMBRE_PDF_ECOSTRES
                vista = "VISTA_REPORTE_ECOSTRES"
                rpte = New rptEcoEstres
        End Select
    End Sub

    Public Sub anularEco(activoAM As Boolean, activoAF As Boolean)
        '---------------- codigo para anular
        Dim nombrePdf2, nombrePdfR, nombrePdfRR, consulta As String
        Dim params As New List(Of String)

        params.Add(usuario)
        params.Add(codigo)
        params.Add(idEmpresa)
        params.Add(editado)

        Select Case codigoTipoEco
            Case Constantes.CODIGO_ECOCARDIOGRAMA
                consulta = objEcocardiograma.sqlAnularRegistro
                nombrePdf2 = ConstantesHC.NOMBRE_PDF_ECO
                nombrePdfR = ConstantesHC.NOMBRE_PDF_ECO_R
                nombrePdfRR = ConstantesHC.NOMBRE_PDF_ECO_RR
            Case Constantes.CODIGO_ECO_ESTRES
                consulta = objEcocardiogramaEstres.sqlAnularRegistro
                nombrePdf2 = ConstantesHC.NOMBRE_PDF_ECOSTRES
                nombrePdfR = ConstantesHC.NOMBRE_PDF_ECOSTRES_R
                nombrePdfRR = ConstantesHC.NOMBRE_PDF_ECOSTRES_RR
            Case Else
                consulta = ""
        End Select
        General.ejecutarSQL(consulta, params)
    End Sub

    Public Sub imprimir()
        Try
            Dim ruta, nombreArchivo As String

            nombreArchivo = nombrePDF & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigo & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = IO.Path.GetTempPath() & nombreArchivo
            'ftp_reportes.llamarArchivo(ruta, nombreArchivo, registro, nombrePDF)
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(nombrePDF, codigo, rpte,
                                      codigo, "{" & vista & ".Codigo_Eco} = " & codigo &
                                         " AND {" & vista & ".Modulo}= " & moduloReporte,
                                      nombrePDF, IO.Path.GetTempPath)
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try

    End Sub
End Class
