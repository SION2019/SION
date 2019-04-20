Public Class Remision
    Public Property modalidad As String
    Public Property datosMedicos As String
    Public Property prioridad As String
    Public Property complejidad As String
    Public Property otras As String
    Public Property antecedentes As String
    Public Property glasgow As String
    Public Property descripglasgow As String
    Public Property presionsis As String
    Public Property presiondias As String
    Public Property freccar As String
    Public Property frecresp As String
    Public Property oxigeno As String
    Public Property remisionListar As String
    Public Property remisionCargar As String
    Public Property moduloReporte As Integer
    Public Property codigoRemision As String
    Public Property codigoOrden As String
    Public Property traslados As String
    Public Property especialidad As String
    Public Property fechaRemision As DateTime
    Public Property ambulancia As Boolean
    Public Property editado As Boolean
    Public Property usuario As String
    Public Property usuarioReal As String
    Public Property usuarioNombre As String
    Public Property nRegistro As Integer
    Public Property codigoEP As Integer
    Public Property nombreReporte As String
    Public Property consultaVerificar As String
    Public Property consultaSolicitudVerificar As String
    Public Property opcionCancelar As Boolean

    Sub New()
        nombreReporte = ConstantesHC.NOMBRE_PDF_REMISION
        moduloReporte = Constantes.REPORTE_HC
        remisionListar = Consultas.REMISIONES_LISTAR
        remisionCargar = Consultas.REMISION_CARGAR
        consultaVerificar = Consultas.REMISION_VERIFICAR
        consultaSolicitudVerificar = Consultas.REMISION_SOLICITUD_VERIFICAR
    End Sub

    Public Sub cargarListaRemisiones(ByRef listRemision As ListBox)
        Dim params As New List(Of String)
        params.Add(nRegistro)
        General.cargarLista(remisionListar, params, "Nombre", "Codigo_remision", listRemision)
    End Sub

    Public Sub cargarRemision()
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(codigoRemision)
        General.llenarTabla(remisionCargar, params, dt)
        If dt.Rows.Count > 0 Then
            modalidad = dt.Rows(0).Item("Codigo_Motivo_observacion").ToString
            datosMedicos = dt.Rows(0).Item("Datos_medicos").ToString
            fechaRemision = dt.Rows(0).Item("Fecha_Remision").ToString
            prioridad = dt.Rows(0).Item("Codigo_Prioridad").ToString
            complejidad = dt.Rows(0).Item("Codigo_complejidad").ToString
            otras = dt.Rows(0).Item("Otras").ToString
            antecedentes = dt.Rows(0).Item("Antecedentes").ToString
            glasgow = dt.Rows(0).Item("Glasgow").ToString
            descripglasgow = dt.Rows(0).Item("Descripcion_Glasgow").ToString
            presionsis = dt.Rows(0).Item("Presion_Sistolica").ToString
            presiondias = dt.Rows(0).Item("Presion_Diastolica").ToString
            freccar = dt.Rows(0).Item("Frecuencia_Cardiaca").ToString
            frecresp = dt.Rows(0).Item("Frecuencia_Respiratoria").ToString
            ambulancia = dt.Rows(0).Item("Ambulancia").ToString
            traslados = dt.Rows(0).Item("Codigo_traslado").ToString
            oxigeno = dt.Rows(0).Item("Oxigeno").ToString
            especialidad = dt.Rows(0).Item("Codigo_especialidad").ToString
            usuarioNombre = dt.Rows(0).Item("Usuario").ToString
            fechaRemision = dt.Rows(0).Item("Fecha_remision").ToString
        End If

    End Sub

    Public Overridable Sub guardarRemision()
        Try
            If String.IsNullOrEmpty(codigoOrden) Then
                codigoOrden = Constantes.VALOR_PREDETERMINADO
            End If
            HistoriaClinicaBLL.guardarRemision(Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub imprimirRemision(indicelista As Integer)
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre, campo As String

            If indicelista = 0 Then
                codigoNombre = nRegistro
                campo = "N_Registro"
                'nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                '                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
                'ruta = Path.GetTempPath() & nombreArchivo
                'ftp_reportes.llamarArchivo(ruta, nombreArchivo, nRegistro, nombreReporte)
            Else
                codigoNombre = codigoRemision
                campo = "Codigo_remision"

            End If
            nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF

            formula = "{VISTA_REMISION." & campo & "} = " & codigoNombre & " 
                                            AND {VISTA_REMISION.Modulo}=" & moduloReporte & ""
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(nombreReporte, nRegistro, New rptRemision,
                                  codigoNombre, formula,
                                  nombreReporte, IO.Path.GetTempPath)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Overridable Sub anularRemision()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoRemision)
        params.Add(codigoEP)
        General.ejecutarSQL(Consultas.ANULAR_REMISION, params)
    End Sub
End Class
