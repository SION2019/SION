Public Class Enfermeria
    Public Property registro As Integer
    Public Property codigoOrden As String
    Public Property listNota As DataTable
    Public Property listInsumo As DataTable
    Public Property listGlucometria As DataTable
    Public Property listExamenes As DataTable
    Public Property listExamenesHemo As DataTable
    Public Property codigoEP As Integer
    Public Property objInsumosEnfer As New InsumoEnfermeria
    Public Property objNotaEnfer As New NotaEnfermeria
    Public Property objGlucomEnfer As New GlucometriaEnfermeria
    Public Property objParaclinico As New ParaclinicoEnfermeria
    Public Property objHemoderivado As New HemoderivadoEnfermeria
    Public Property insumoCarga As String
    Public Property notaCarga As String
    Public Property glucomCargar As String
    Public Property examenCargar As String
    Public Property examenHemoCargar As String
    Public Property moduloReporte As Integer
    Public Property nombreReporte As String
    Public Property opcionCancelar As Boolean

    Sub New()
        moduloReporte = Constantes.REPORTE_HC
        nombreReporte = ConstantesHC.NOMBRE_PDF_GLUCOMETRIA_ENFER
        insumoCarga = ConsultasHC.ENFERMERIA_INSUMO_LISTAR
        listInsumo = New DataTable
        listInsumo.Columns.Add("Codigo_Orden", Type.GetType("System.String"))
        listInsumo.Columns.Add("Nombre", Type.GetType("System.String"))
        listInsumo.Columns.Add("Fecha_Orden", Type.GetType("System.DateTime"))

        notaCarga = ConsultasHC.ENFERMERIA_NOTA_LISTAR
        listNota = New DataTable
        listNota.Columns.Add("Codigo_Orden", Type.GetType("System.String"))
        listNota.Columns.Add("Nombre", Type.GetType("System.String"))
        listNota.Columns.Add("Nota", Type.GetType("System.String"))
        listNota.Columns.Add("Fecha_Nota", Type.GetType("System.DateTime"))

        glucomCargar = ConsultasHC.ENFERMERIA_GLUCOMETRIA_LISTA
        listGlucometria = New DataTable
        listGlucometria.Columns.Add("Codigo_Orden", Type.GetType("System.String"))
        listGlucometria.Columns.Add("Nombre", Type.GetType("System.String"))
        listGlucometria.Columns.Add("Fecha", Type.GetType("System.DateTime"))

        examenCargar = ConsultasHC.ENFERMERIA_EXAMENES_LISTA
        listExamenes = New DataTable
        listExamenes.Columns.Add("Codigo_Orden", Type.GetType("System.String"))
        listExamenes.Columns.Add("Nombre", Type.GetType("System.String"))
        listExamenes.Columns.Add("Fecha_Orden", Type.GetType("System.DateTime"))

        examenHemoCargar = ConsultasHC.ENFERMERIA_EXAMENES_HEMODERIVADOS_LISTAR
        listExamenesHemo = New DataTable
        listExamenesHemo.Columns.Add("Codigo_Orden", Type.GetType("System.String"))
        listExamenesHemo.Columns.Add("Nombre", Type.GetType("System.String"))
        listExamenesHemo.Columns.Add("Fecha_Orden", Type.GetType("System.DateTime"))

        objInsumosEnfer = New InsumoEnfermeria
        objNotaEnfer = New NotaEnfermeria
        objGlucomEnfer = New GlucometriaEnfermeria
        objParaclinico = New ParaclinicoEnfermeria
        objHemoderivado = New HemoderivadoEnfermeria
    End Sub


    Public Sub cargarListaOrdenesInsumoEnfer()
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(insumoCarga, params, listInsumo)
    End Sub

    Public Sub cargarOrdenInsumoEnfer()
        If String.IsNullOrEmpty(codigoOrden) Then
            codigoOrden = -1
        End If

        objInsumosEnfer.codigoOrden = codigoOrden
        objInsumosEnfer.cargarInsumoEnfer()
    End Sub
    Public Sub cargarInsumoEnferAutomatico()
        objInsumosEnfer.cargarInsumoFisioConfigAud()
        objInsumosEnfer.cargarInsumoEnferPeriodicidad()
    End Sub
    Public Overridable Sub guardarOrdenInsumoEnfer()
        Try
            If String.IsNullOrEmpty(objInsumosEnfer.codigoOrden) Then
                HistoriaClinicaDAL.guardarDetalleInsumosEnfer(objInsumosEnfer)
            Else
                HistoriaClinicaDAL.actualizarDetalleInsumosEnfer(objInsumosEnfer)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overridable Sub guardarGlucometriaEnfe()
        Try
            HistoriaClinicaDAL.guardarGlucometriaHC(Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub anularOrdenInsumoEnfer()
        objInsumosEnfer.anularOrden()
    End Sub
    Public Sub cargarListaNotaEnfer()
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(notaCarga, params, listNota)
    End Sub

    Public Overridable Sub guardarNotaEnfer()
        Try
            If String.IsNullOrEmpty(objNotaEnfer.codigoNota) Then
                HistoriaClinicaDAL.guardarNotaEnfer(objNotaEnfer)
            Else
                HistoriaClinicaDAL.actualizarNotaEnfer(objNotaEnfer)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub anularNotaEnfer()
        objNotaEnfer.anularNota()
    End Sub

    Public Sub imprimirOrdenInsumoEnfer(indicelista As Integer)
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre As String

            If indicelista = 0 Then
                codigoNombre = registro
                nombreArchivo = ConstantesHC.NOMBRE_PDF_ORDEN_INSUMO_ENFER & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
                formula = "{VISTA_ORDEN_INSUMO_ENFER.N_Registro} = " & registro & " 
                                            AND {VISTA_ORDEN_INSUMO_ENFER.Modulo}=" & moduloReporte & ""
            Else
                codigoNombre = codigoOrden
                nombreArchivo = ConstantesHC.NOMBRE_PDF_ORDEN_INSUMO_ENFER & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
                formula = "{VISTA_ORDEN_INSUMO_ENFER.Codigo_Orden_Insumo_E} = " & codigoOrden & " 
                                            AND {VISTA_ORDEN_INSUMO_ENFER.Modulo}=" & moduloReporte & ""
            End If
            ruta = Path.GetTempPath() & nombreArchivo
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(ConstantesHC.NOMBRE_PDF_ORDEN_INSUMO_ENFER, registro, New rptOrdenInsumoEnfer,
                                  codigoNombre, formula,
                                  ConstantesHC.NOMBRE_PDF_ORDEN_INSUMO_ENFER, IO.Path.GetTempPath,,)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub imprimirNotaEnfer(indicelista As Integer)
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre As String

            If indicelista = 0 Then
                codigoNombre = registro
                nombreArchivo = ConstantesHC.NOMBRE_PDF_ORDEN_INSUMO_ENFER & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
                formula = "{VISTA_NOTA_ENFERMERIA.N_Registro} = " & registro & " 
                                            AND {VISTA_NOTA_ENFERMERIA.Modulo}=" & moduloReporte & ""
            Else
                codigoNombre = codigoOrden
                nombreArchivo = ConstantesHC.NOMBRE_PDF_NOTA_ENFER & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
                formula = "{VISTA_NOTA_ENFERMERIA.ID_NOTA} = " & codigoOrden & " 
                                            AND {VISTA_NOTA_ENFERMERIA.Modulo}=" & moduloReporte & ""
            End If
            ruta = Path.GetTempPath() & nombreArchivo
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(ConstantesHC.NOMBRE_PDF_NOTA_ENFER, registro, New rptNotaEnfermeria,
            codigoNombre, formula,
            ConstantesHC.NOMBRE_PDF_NOTA_ENFER, IO.Path.GetTempPath)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#Region "Glucometria Enfermeria"
    Public Sub cargarListaOrdenesGlucomEnfer()
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(glucomCargar, params, listGlucometria)
    End Sub
    Public Sub cargarOrdenGlucomEnfer()
        If String.IsNullOrEmpty(codigoOrden) Then
            codigoOrden = -1
        End If

        objGlucomEnfer.codigoOrden = codigoOrden
        objGlucomEnfer.cargarGlucomEnfer()
    End Sub

    Public Sub imprimirGlucometria(indicelista As Integer)
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre, campo As String
            If indicelista = 0 Then
                codigoNombre = registro
                campo = "N_Registro"
            Else
                codigoNombre = codigoOrden
                campo = "Codigo_orden"

            End If
            nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            formula = "{VISTA_GLUCOMETRIA_ENFE." & campo & "} = " & codigoNombre & " 
                                            AND {VISTA_GLUCOMETRIA_ENFE.Modulo}=" & moduloReporte & ""
            ruta = Path.GetTempPath() & nombreArchivo
                Dim reporte As New ftp_reportes
            reporte.crearReportePDF(nombreReporte, registro, New rptGlucometria,
                codigoNombre, formula, nombreReporte, IO.Path.GetTempPath)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region
#Region "paraclinicos"
    Public Sub cargarListaOrdenesParaclinicos()
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(examenCargar, params, listExamenes)
    End Sub
#End Region
#Region "Hemoderivados"
    Public Sub cargarListaOrdenesHemoderivados()
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(examenHemoCargar, params, listExamenesHemo)
    End Sub
#End Region

End Class
