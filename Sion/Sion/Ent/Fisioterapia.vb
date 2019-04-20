Public Class Fisioterapia
    Public Property registro As Integer
    Public Property codigoOrden As String
    Public Property listOxigeno As DataTable
    Public Property listNebulizacion As DataTable
    Public Property listTerapiaFyR As DataTable
    Public Property listNota As DataTable
    Public Property listInsumo As DataTable
    Public Property codigoEP As Integer
    Public Property objTerapiaFyR As New TerapiaFisicaYRespiratoria
    Public Property objInsumosFisio As New InsumoFisioterapia
    Public Property objNotaFisio As New NotaFisioterapia
    Public Property objOxigeno As New Oxigeno
    Public Property objnebulizacion As New Nebulizacion
    Public Property terapiaCarga As String
    Public Property insumoCarga As String
    Public Property notaCarga As String
    Public Property oxigenoCarga As String
    Public Property nebulizacionCarga As String
    Public Property notaReporte As String
    Public Property nebulizacionReporte As String
    Public Property moduloReporte As Integer
    Public Property opcionCancelar As Boolean

    Sub New()
        notaReporte = ConstantesHC.NOMBRE_PDF_NOTA_FISIO
        moduloReporte = Constantes.REPORTE_HC
        terapiaCarga = ConsultasHC.FISIOTERAPIA_FISICA_Y_RESPIRATORIA_LISTAR
        listTerapiaFyR = New DataTable
        listTerapiaFyR.Columns.Add("Codigo_Orden", Type.GetType("System.String"))
        listTerapiaFyR.Columns.Add("Codigo_Procedimiento", Type.GetType("System.String"))
        listTerapiaFyR.Columns.Add("Nombre", Type.GetType("System.String"))

        insumoCarga = ConsultasHC.FISIOTERAPIA_INSUMO_LISTAR
        listInsumo = New DataTable
        listInsumo.Columns.Add("Codigo_Orden", Type.GetType("System.String"))
        listInsumo.Columns.Add("Nombre", Type.GetType("System.String"))
        listInsumo.Columns.Add("Fecha_Orden", Type.GetType("System.DateTime"))

        notaCarga = ConsultasHC.FISIOTERAPIA_NOTA_LISTAR
        listNota = New DataTable
        listNota.Columns.Add("Codigo_Orden", Type.GetType("System.String"))
        listNota.Columns.Add("Nombre", Type.GetType("System.String"))
        listNota.Columns.Add("Nota", Type.GetType("System.String"))
        listNota.Columns.Add("Fecha_Nota", Type.GetType("System.DateTime"))

        oxigenoCarga = ConsultasHC.OXIGENO_LISTAR
        listOxigeno = New DataTable
        listOxigeno.Columns.Add("Codigo_Orden", Type.GetType("System.String"))
        listOxigeno.Columns.Add("Nombre", Type.GetType("System.String"))

        nebulizacionCarga = ConsultasHC.NEBULIZACION_LISTAR
        listNebulizacion = New DataTable
        listNebulizacion.Columns.Add("Codigo_Orden", Type.GetType("System.String"))
        listNebulizacion.Columns.Add("Nombre", Type.GetType("System.String"))

        objTerapiaFyR = New TerapiaFisicaYRespiratoria
        objInsumosFisio = New InsumoFisioterapia
        objNotaFisio = New NotaFisioterapia
        objOxigeno = New Oxigeno
        objnebulizacion = New Nebulizacion
    End Sub
    Public Sub cargarListaOxigenosOrdenados()
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(oxigenoCarga, params, listOxigeno)
        objOxigeno.nRegistro = registro
    End Sub
    Public Sub cargarListaNebulizacion()
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(nebulizacionCarga, params, listNebulizacion)
        objnebulizacion.nRegistro = registro
    End Sub
    Public Sub cargarListaTerapia()
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(terapiaCarga, params, listTerapiaFyR)
    End Sub

    Public Sub cargarDetalleTerapia(codigoProcedimiento As String)

        objTerapiaFyR.codigoOrden = codigoOrden
        objTerapiaFyR.codigoProcedimiento = codigoProcedimiento
        objTerapiaFyR.cargarTerapiaFyR()
    End Sub
    Public Overridable Sub guardarDetalleTerapia()
        Try
            HistoriaClinicaDAL.guardarDetalleTerapia(objTerapiaFyR)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub cargarListaOrdenesInsumoFisio()
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(insumoCarga, params, listInsumo)
    End Sub

    Public Sub cargarOrdenInsumoFisio()
        If String.IsNullOrEmpty(codigoOrden) Then
            codigoOrden = -1
        End If

        objInsumosFisio.codigoOrden = codigoOrden
        objInsumosFisio.cargarInsumoFisio()
    End Sub
    Public Sub cargarInsumoFisioAutomatico()
        objInsumosFisio.cargarInsumoFisioConfigAud()
        objInsumosFisio.cargarInsumoFisioPeriodicidad()
    End Sub
    Public Overridable Sub guardarOrdenInsumoFisio()
        Try
            If String.IsNullOrEmpty(objInsumosFisio.codigoOrden) Then
                HistoriaClinicaDAL.guardarDetalleInsumosFisio(objInsumosFisio)
            Else
                HistoriaClinicaDAL.actualizarDetalleInsumosFisio(objInsumosFisio)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub anularOrdenInsumoFisio()
        objInsumosFisio.anularOrden()
    End Sub
    Public Sub cargarListaNotaFisio()
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(notaCarga, params, listNota)
    End Sub
    Public Overridable Sub guardarNotaFisio()
        Try
            If String.IsNullOrEmpty(objNotaFisio.codigoNota) Then
                HistoriaClinicaDAL.guardarNotaFisio(objNotaFisio)
            Else
                HistoriaClinicaDAL.actualizarNotaFisio(objNotaFisio)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub anularNotaFisio()
        objNotaFisio.anularNota()

    End Sub
    Public Overridable Sub guardarOxigeno()
        Try
            HistoriaClinicaDAL.guardarFisioterapiaOxigeno(objOxigeno)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overridable Sub guardarNebulizacion()
        Try
            HistoriaClinicaDAL.guardarFisioterapiaNebulizacion(objnebulizacion)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub imprimirOrdenInsumoFisio(indicelista As Integer)
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre As String

            If indicelista = 0 Then
                codigoNombre = registro
                nombreArchivo = ConstantesHC.NOMBRE_PDF_ORDEN_INSUMO_FISIO & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
                formula = "{VISTA_ORDEN_INSUMO_FISIO.N_Registro} = " & registro & " 
                                            AND {VISTA_ORDEN_INSUMO_FISIO.Modulo}=" & moduloReporte & ""
            Else
                codigoNombre = codigoOrden
                nombreArchivo = ConstantesHC.NOMBRE_PDF_ORDEN_INSUMO_FISIO & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
                formula = "{VISTA_ORDEN_INSUMO_FISIO.Codigo_Orden_Insumo_F} = " & codigoOrden & " 
                                            AND {VISTA_ORDEN_INSUMO_FISIO.Modulo}=" & moduloReporte & ""
            End If
            ruta = Path.GetTempPath() & nombreArchivo
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(ConstantesHC.NOMBRE_PDF_ORDEN_INSUMO_FISIO, registro, New rptOrdenInsumoFisio,
                              codigoNombre, formula,
                              ConstantesHC.NOMBRE_PDF_ORDEN_INSUMO_FISIO, IO.Path.GetTempPath)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Sub imprimirNotaFisio(indicelista As Integer)
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre, campo As String

            If indicelista = 0 Then
                'codigoNombre = registro
                'nombreArchivo = notaReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                '                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
                'ruta = Path.GetTempPath() & nombreArchivo
                'ftp_reportes.llamarArchivo(ruta, nombreArchivo, registro, notaReporte)
                campo = "N_Registro"
                codigoNombre = registro
            Else
                codigoNombre = codigoOrden
                campo = "ID_NOTA"
            End If
            nombreArchivo = notaReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = Path.GetTempPath() & nombreArchivo
            formula = "{VISTA_NOTA_FISIOTERAPIA." & campo & "} = " & codigoNombre & " 
                                            AND {VISTA_NOTA_FISIOTERAPIA.Modulo}=" & moduloReporte & ""
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(notaReporte, registro, New rptNotaFisioterapia,
                                  codigoNombre, formula,
                                  notaReporte, IO.Path.GetTempPath)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Sub imprimirNebulizacion(indicelista As Integer)
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre As String

            If indicelista = 0 Then
                codigoNombre = registro
                nombreArchivo = ConstantesHC.NOMBRE_PDF_NEBULIZACION_FISIO & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
                ruta = Path.GetTempPath() & nombreArchivo
            Else
                codigoNombre = codigoOrden
                nombreArchivo = notaReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
                ruta = Path.GetTempPath() & nombreArchivo
                formula = "{VISTA_FISIOTERAPIA_NEBULIZACION.Codigo_orden} = " & codigoOrden & " AND {VISTA_FISIOTERAPIA_NEBULIZACION.Modulo}=" & moduloReporte & ""
                Dim reporte As New ftp_reportes
                reporte.crearReportePDF(notaReporte, registro, New rptFisioterapiaNebulizacion,
                                  codigoNombre, formula,
                                  notaReporte, IO.Path.GetTempPath)
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

End Class
