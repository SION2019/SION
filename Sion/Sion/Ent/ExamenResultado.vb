Public Class ExamenResultado
    Inherits Procedimientos
    Public Property codigoTipo As Integer
    Public Property observacion As String
    Public Property dtExamen As DataTable
    Public Property dtExamenEliminados As DataTable
    Public Property verificarRegistroGuardado As Boolean
    Public Property contador As Integer = 1
    Public Property tipoExam As String
    Public Property sqlCargarRegistroDetalle As String
    Public Property nombreArchivo As String
    Public Property archivoAudio As Byte()
    Public Property moduloReporte As String
    Public Property nombrePDF As String
    Property sqlGuardarRegistroHosp As String
    Property procedencia As Integer
    Property existeAudio As Boolean
    Private muestraPDF As Boolean
    Private mustraImage As Boolean
    Sub New()

        titulo = TitulosForm.TITULO_RESULT
        sqlCargarRegistro = ConsultasHC.RESULTADO_EXAMENES_CARGAR
        sqlAnularRegistro = ConsultasHC.RESULTADO_EXAMEN_ANULAR
        sqlGuardarRegistro = ConsultasHC.RESULTADO_EXAMEN_GUARDAR
        sqlGuardarRegistroHosp = "SP_RESULTADOS_HOSPITAL_EXAMENES_CREAR"
        sqlCargarRegistroDetalle = ConsultasHC.RESULTADO_EXAMENES_CARGAR_D
        area = ConstantesHC.NOMBRE_PDF_IMAGENOLOGIA

        dtExamen = New DataTable
        dtExamen.Columns.Add("nombreArchivo", Type.GetType("System.String"))
        dtExamen.Columns.Add("Archivo", Type.GetType("System.Byte[]"))
        dtExamen.Columns.Add("ruta", Type.GetType("System.String"))
        dtExamen.Columns.Add("idArchivo", Type.GetType("System.Int32"))

        moduloReporte = Constantes.REPORTE_HC
        nombrePDF = ConstantesHC.NOMBRE_PDF_IMAGENOLOGIA
    End Sub
    Public Function cargarConsultaGuardar() As String
        Return If(procedencia = 0, sqlGuardarRegistro, sqlGuardarRegistroHosp)
    End Function
    Public Overrides Sub guardarRegistro()
        Dim objExamenBLL As New ExamenResultadoBLL
        If tipoExam = Constantes.TipoEXAM Then
            objExamenBLL.guardarResultado(Me)
        ElseIf tipoExam = Constantes.TipoEXAMAtencion Then
            objExamenBLL.guardarResultadoLab(Me)
        Else
            objExamenBLL.guardarResultadoExamen(Me)
        End If
    End Sub

    Public Sub cargarResultado()
        Dim params As New List(Of String)
        Dim drResultado As DataRow
        Dim consulta As String

        If tipoExam = Constantes.TipoEXAM Then
            consulta = sqlCargarRegistro
        ElseIf tipoExam = Constantes.TipoEXAMAtencion Then
            consulta = Consultas.RESULTADO_LABORATORIO_CARGAR
        Else
            consulta = Consultas.RESULTADO_EXAMEN_LABORATORIO_CARGAR
        End If

        Try
            params.Add(CodigoOrden)
            params.Add(CodigoProcedimiento)
            params.Add(procedencia)
            drResultado = General.cargarItem(consulta, params)
            If drResultado IsNot Nothing Then
                codigoTipo = drResultado.Item(0).ToString
                observacion = drResultado.Item(1).ToString
                existeAudio = drResultado.Item(2)
                archivoAudio = If(IsDBNull(drResultado.Item(3)), Nothing, drResultado.Item(3))
                usuarioReal = If(IsDBNull(drResultado.Item(4)), SesionActual.idUsuario, drResultado.Item(4))
                verificarRegistroGuardado = True '-------------------bandera que me identifica que los datos estan en la base de datos 
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub cargarImagenDicom()
        Dim params As New List(Of String)
        Dim contador As Integer
        Dim ruta As String
        Dim consulta As String

        If tipoExam = Constantes.TipoEXAM Then
            consulta = sqlCargarRegistroDetalle
        ElseIf tipoExam = Constantes.TipoEXAMAtencion Then
            consulta = Consultas.RESULTADO_LABORATORIO_CARGAR_DETALLE
        Else
            consulta = Consultas.RESULTADO_EXAMEN_LAB_IMG_CARGAR
        End If

        Try
            params.Add(CodigoOrden)
            params.Add(CodigoProcedimiento)
            params.Add(procedencia)
            General.llenarTabla(consulta, params, dtExamen)
            While (dtExamen.Rows.Count - 1) >= contador
                ruta = IO.Path.GetTempPath & Format(DateTime.Now, "ddMMyyyyHHmmss") &
                        ConstantesHC.NOMBRE_PDF_SEPARADOR & contador &
                        ConstantesHC.NOMBRE_PDF_SEPARADOR & Replace(dtExamen.Rows(contador).Item("nombreArchivo"), ":", "_")
                IO.File.WriteAllBytes(ruta, dtExamen.Rows(contador).Item("Archivo"))
                dtExamen.Rows(contador).Item("ruta") = ruta
                contador = contador + 1
            End While
        Catch ex As Exception
            Throw
        End Try

    End Sub
    Public Function muestraVisorPDF() As Boolean
        Return If(codigoTipo = ConstantesHC.CODIGO_PDF, True, False)
    End Function
    Public Function muestraVisorDicom() As Boolean
        Return If(codigoTipo = ConstantesHC.CODIGO_IMAGEN_DCM, True, False)
    End Function
    Public Function cargarImagen(abrirArchivo As OpenFileDialog) As DataTable
        Return General.subirimagenDiagnostica(abrirArchivo)
    End Function

    Public Sub imprimir()
        Try
            Dim ruta, nombreArchivo As String
            Dim codigoCompuesto As String
            Dim cadena As String
            Dim repor As ReportClass
            Dim reporte As New ftp_reportes
            codigoCompuesto = CodigoOrden & CodigoProcedimiento
            nombreArchivo = nombrePDF & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoCompuesto & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = IO.Path.GetTempPath() & nombreArchivo
            If tipoExam = Constantes.TipoEXAM Then
                repor = New rptResultado
                cadena = "{VISTA_RESULTADOS_PARACLINICOS.Codigo_Orden} = " & CInt(CodigoOrden) &
                                                       " And {VISTA_RESULTADOS_PARACLINICOS.Codigo_Procedimiento}='" & CodigoProcedimiento &
                                                       "' AND {VISTA_RESULTADOS_PARACLINICOS.Modulo}= " & moduloReporte
            Else
                repor = New rptResultadoAtencion
                cadena = "{VISTA_RESULTADO_PARACLINICO_ATENCION.Codigo_Orden} = " & CInt(CodigoOrden) &
                                                       " And {VISTA_RESULTADO_PARACLINICO_ATENCION.Codigo_Procedimiento}='" & CodigoProcedimiento &
                                                       "' AND {VISTA_RESULTADO_PARACLINICO_ATENCION.Modulo}= 4"
            End If
            reporte.crearReportePDF(nombrePDF, codigoCompuesto, repor,
                                      codigoCompuesto, cadena,
                                      nombrePDF, IO.Path.GetTempPath, CodigoProcedimiento)
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try

    End Sub


End Class
