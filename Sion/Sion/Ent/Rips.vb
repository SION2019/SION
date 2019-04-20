Public Class Rips
    Public Property dtFactura As DataTable
    Public Property dtAm As New DataTable
    Public Property dtAc As New DataTable
    Public Property dtUs As New DataTable
    Public Property dtAt As New DataTable
    Public Property dtAf As New DataTable
    Public Property dtAp As New DataTable
    Public Property dtAh As New DataTable
    Public Property fecha As Date
    Public Property fecha2 As Date
    Public Property codigoEPS As String
    Public Property fechaExp As Date
    Public Property ruta As String
    Public Property consecutivo As String
    Public Property dtArchivos As DataTable
    Public Property codigoPrestador As String
    Public consultaCargaFactura As String
    Public Property consultacargarFacturaAm As String
    Public Property rac As Boolean
    Public Property ram As Boolean
    Public Property rat As Boolean
    Public Property rap As Boolean
    Public Property raf As Boolean
    Public Property rus As Boolean
    Public Property rah As Boolean
    Public Property rct As Boolean
    Public Property codigoTipoRips As Integer
    Public Property dtDinamico As DataTable
    Public Property dtLista As New DataTable
    Public Property dtListaConfigurada As New DataTable
    Public Property codigoGrupo As Integer
    Public Property usuario As Integer
    Public Property descripcion As String
    Public Property dtArchivosDinamicos As New DataTable
    Public Property Modulo As String
    Public Property dtNew As New DataTable
    Public Property idEPS As Integer


    Public Sub cargarFacturas()
        If Modulo = Constantes.CODIGO_MENU_ASIS Then
            Dim params As New List(Of String)
            params.Add(fecha)
            params.Add(fecha2)
            params.Add(codigoEPS)
            params.Add(SesionActual.idEmpresa)
            General.llenarTabla(consultaCargaFactura, params, dtFactura)
        Else
            Dim params As New List(Of String)
            params.Add(fecha)
            params.Add(fecha2)
            params.Add(codigoEPS)
            params.Add(SesionActual.idEmpresa)
            General.llenarTabla(consultacargarFacturaAm, params, dtFactura)
        End If
    End Sub

    Public Sub guardarAfAmbulancia()
        dtDinamico = New DataTable
        Dim filas As DataRow()
        dtNew = dtFactura.Clone
        filas = dtFactura.Select("Agregar=1")
        For Each row As DataRow In filas
            dtNew.ImportRow(row)
        Next
        RipsDAL.ripsAmbulanciaAf(Me)
    End Sub

    Public Sub guardarAtAmbulancia()
        dtDinamico = New DataTable
        Dim filas As DataRow()
        dtNew = dtFactura.Clone
        filas = dtFactura.Select("Agregar=1")
        For Each row As DataRow In filas
            dtNew.ImportRow(row)
        Next
        RipsDAL.ripsAmbulanciaAt(Me)
    End Sub

    Public Sub guardarUsAmbulancia()
        dtDinamico = New DataTable
        Dim filas As DataRow()
        dtNew = dtFactura.Clone
        filas = dtFactura.Select("Agregar=1")
        For Each row As DataRow In filas
            dtNew.ImportRow(row)
        Next
        RipsDAL.ripsAmbulanciaUs(Me)
    End Sub

    Public Sub generarArchivosDinamicos()
        If Modulo = Constantes.CODIGO_MENU_ASIS Then
            dtArchivosDinamicos = New DataTable
            Select Case dtDinamico.Rows.Count > 0
                Case codigoTipoRips = 0
                    generarArchivo(dtArchivosDinamicos, "AM")
                Case codigoTipoRips = 1
                    generarArchivo(dtArchivosDinamicos, "AF")
                Case codigoTipoRips = 2
                    generarArchivo(dtArchivosDinamicos, "US")
                Case codigoTipoRips = 3
                    generarArchivo(dtArchivosDinamicos, "AC")
                Case codigoTipoRips = 4
                    generarArchivo(dtArchivosDinamicos, "AP")
                Case codigoTipoRips = 5
                    generarArchivo(dtArchivosDinamicos, "AH")
                Case codigoTipoRips = 6
                    generarArchivo(dtArchivosDinamicos, "AT")
                Case codigoTipoRips = 10
                    generarArchivo(dtArchivosDinamicos, "TT")
            End Select
        Else
            dtArchivosDinamicos = New DataTable
            Select Case dtDinamico.Rows.Count > 0
                Case codigoTipoRips = 7
                    generarArchivo(dtArchivosDinamicos, "AF")
                Case codigoTipoRips = 8
                    generarArchivo(dtArchivosDinamicos, "US")
                Case codigoTipoRips = 9
                    generarArchivo(dtArchivosDinamicos, "AT")
            End Select
        End If
    End Sub

    Private Sub generarArchivo(ByRef dt As DataTable, pSigla As String)
        dt = dtDinamico.Copy
        dtArchivos.Rows.Add(codigoPrestador, Format(fechaExp.Date, Constantes.FORMATO_FECHA2), pSigla + consecutivo, dt.Rows.Count.ToString())
        Dim consultaAM As String = ""
        General.concatenacionComas(dt, consultaAM)
        File.WriteAllText(ruta + "\" + pSigla + "" + consecutivo + ".TXT", consultaAM)
    End Sub

    Public Sub ripsDinamico()
        dtDinamico = New DataTable
        If dtNew.Rows.Count = 0 Then
            Dim filas As DataRow()
            dtNew = dtFactura.Clone
            filas = dtFactura.Select("Agregar=1")
            For Each row As DataRow In filas
                dtNew.ImportRow(row)
            Next
        End If
        RipsDAL.configuracionDinamica(Me)
    End Sub

    Public Sub guardarConfiguracion(ByRef form As FormRips)
        For i = 0 To dtListaConfigurada.Rows.Count - 1
            dtListaConfigurada.Rows(i).Item("orden") = i + 1
        Next

        RipsDAL.configuracionGuardar(Me, form)
    End Sub

    Public Sub cargarListaParametro()
        Dim params As New List(Of String)
        params.Add(codigoTipoRips)
        params.Add(codigoGrupo)
        General.llenarTabla(Consultas.RIPS_LISTA_PARAMETRO, params, dtLista)
    End Sub

    Public Sub cargarListaConfigurada()
        Dim params As New List(Of String)
        params.Add(codigoGrupo)
        params.Add(codigoTipoRips)
        General.llenarTabla(Consultas.RIPS_LISTA_CONFIGURADA, params, dtListaConfigurada)
    End Sub
    Public Sub codigoHabilitacion()
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(SesionActual.idEmpresa)
        General.llenarTabla(Consultas.CODIGO_HABILITACION, params, dt)
        If dt.Rows.Count > 0 Then
            codigoPrestador = dt.Rows(0).Item("Cod_Habilitacion").ToString
        End If
    End Sub

    Public Sub generarArchivoCT()
        If rct = True Then
            Dim dt As New DataTable
            dt = dtArchivos.Copy
            If dt.Rows.Count > 0 Then
                Dim consultaCT As String = ""
                General.concatenacionComas(dt, consultaCT)
                File.WriteAllText(ruta + "\CT" + consecutivo + ".TXT", consultaCT)
            End If
        End If
    End Sub

    Public Function verificarFormulario(form As Windows.Forms.Form) As Boolean
        Dim valor As Boolean = False
        For Each f As Windows.Forms.Form In Application.OpenForms
            If f.Name = form.Name Then
                valor = True
            End If
        Next
        Return valor
    End Function
    Public Sub postAM()
        If dtAm.Rows.Count > 0 Then
            dtArchivos.Rows.Add(codigoPrestador, fechaExp.ToShortDateString, "AM" + consecutivo, dtAm.Rows.Count.ToString())
            Dim consultaAM As String = ""
            General.concatenacionComas(dtAm, consultaAM)
            File.WriteAllText(ruta + "\AM" + consecutivo + ".TXT", consultaAM)
        End If
    End Sub

    Sub New()
        dtFactura = New DataTable
        dtFactura.Columns.Add("Codigo factura", Type.GetType("System.String"))
        dtFactura.Columns.Add("Paciente", Type.GetType("System.String"))
        dtFactura.Columns.Add("Fecha factura", Type.GetType("System.DateTime"))
        dtFactura.Columns.Add("Valor factura", Type.GetType("System.Double"))
        dtFactura.Columns.Add("Agregar", Type.GetType("System.Boolean"))
        dtFactura.Columns.Add("Registro", Type.GetType("System.String"))
        dtFactura.Columns.Add("idempresa", Type.GetType("System.String"))
        dtFactura.Columns.Add("Autorizacion", Type.GetType("System.String"))

        dtArchivos = New DataTable
        dtArchivos.Columns.Add("Codigo prestador")
        dtArchivos.Columns.Add("Fecha expedicion")
        dtArchivos.Columns.Add("Nombre archivo")
        dtArchivos.Columns.Add("Cantidad registro")
        consultaCargaFactura = Consultas.RIPS_EPS_FACTURA
        consultacargarFacturaAm = Consultas.RIP_EPS_FACTURA_AM

        dtListaConfigurada.Columns.Add("Codigo_parametro", Type.GetType("System.String"))
        dtListaConfigurada.Columns.Add("descripcion", Type.GetType("System.String"))
        dtListaConfigurada.Columns.Add("orden", Type.GetType("System.String"))

    End Sub
End Class
