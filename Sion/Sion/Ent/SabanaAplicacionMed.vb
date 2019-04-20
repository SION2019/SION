Public Class SabanaAplicacionMed
    Inherits Sabana
    Public Property pMostrarSignos As Integer
    Public Property editado As Boolean
    Public Property servicioNeonatal As Boolean
    Public Property pesoAnterior As Double
    Public Property pesoActual As Double
    Public Property pesoGanado As Double
    Public Property ingresoTotal As Double
    Public Property egresoTotal As Double
    Public Property balanceActual As Double
    Public Property balanceAnterior As Double
    Public Property gastoUrinario As Double
    Public Property balanceAcumulado As Double
    Public Property dtHoraSabanaSilverman As DataTable
    Public Property dtHoraSabanaValoracion As DataTable
    Public Property dtHoraSabanaIngreso As DataTable
    Public Property dtHoraSabanaEgreso As DataTable
    Public Property dtHoraSabanaBalance As DataTable
    Public Property dtHoraSabanaAplicacion As DataTable
    Public Property dtHoraSabanaMedicamento As DataTable
    Public Property dtHoraSabanaNPT As DataTable
    Public Property dtHorasabanaGoteo As DataTable
    Public Property sqlDetalleCargaSilverman As String
    Public Property sqlDetalleCargaValoracion As String
    Public Property sqlDetalleCargaIngreso As String
    Public Property sqlDetalleCargaEgreso As String
    Public Property sqlDetalleCargaBalance As String
    Public Property sqlDetalleCargaAplicacion As String
    Public Property sqlDetalleCargaGoteo As String
    Public Property sqlPesoBalanceCarga As String
    Public Property identificadorPredeterminado As String

    Public Sub New()
        pMostrarSignos = ConstantesHC.SABANA_CON_TODO
        identificadorPredeterminado = ConstantesHC.IDENTIFICADOR_SABANA_NO_APLICADO
        sqlDatoPacienteCarga = ConsultasHC.SABANA_PACIENTE_CARGAR
        sqlCodigoSabanaCarga = ConsultasHC.SABANA_CODIGO_CARGAR
        sqlPesoBalanceCarga = ConsultasHC.SABANA_PESO_CARGAR
        sqlDetalleCarga = ConsultasHC.SABANA_SIGNO_VITAL_CARGAR
        sqlDetalleCargaSilverman = ConsultasHC.SABANA_SILVERMAN_CARGAR
        sqlDetalleCargaValoracion = ConsultasHC.SABANA_VALORACION_CARGAR
        sqlDetalleCargaIngreso = ConsultasHC.SABANA_INGRESO_CARGAR
        sqlDetalleCargaEgreso = ConsultasHC.SABANA_EGRESO_CARGAR
        sqlDetalleCargaBalance = ConsultasHC.SABANA_BALANCE_CARGAR
        sqlDetalleCargaAplicacion = ConsultasHC.SABANA_APLICACION_MED_CARGAR
        sqlDetalleCargaGoteo = ConsultasHC.SABANA_GOTEO_CARGAR
        GeneralHC.crearTablaSabana(dtHoraSabanaSilverman)
        GeneralHC.crearTablaSabana(dtHoraSabanaValoracion)
        GeneralHC.crearTablaSabana(dtHoraSabanaIngreso)
        GeneralHC.crearTablaSabana(dtHoraSabanaEgreso)
        GeneralHC.crearTablaSabana(dtHoraSabanaAplicacion)
        GeneralHC.crearTablaSabana(dtHoraSabanaMedicamento)
        GeneralHC.crearTablaSabana(dtHoraSabanaNPT)
        GeneralHC.crearTablaSabanaGoteo(dtHorasabanaGoteo)
        dtHoraSabanaBalance = New DataTable
        dtHoraSabanaBalance.Columns.Add("codigo", Type.GetType("System.Int32"))
        dtHoraSabanaBalance.Columns.Add("0 - 6", Type.GetType("System.String"))
        dtHoraSabanaBalance.Columns.Add("6 - 12", Type.GetType("System.String"))
        dtHoraSabanaBalance.Columns.Add("12 - 18", Type.GetType("System.String"))
        dtHoraSabanaBalance.Columns.Add("18 - 0", Type.GetType("System.String"))
        dtHoraSabanaAplicacion.Columns.Add("CodigoOrden", Type.GetType("System.String"))
        dtHoraSabanaAplicacion.Columns.Add("Nombre", Type.GetType("System.String"))
        dtHoraSabanaAplicacion.Columns.Add("Tipo", Type.GetType("System.String"))
        dtHoraSabanaMedicamento.Columns.Add("CodigoOrden", Type.GetType("System.String"))
        dtHoraSabanaMedicamento.Columns.Add("Nombre", Type.GetType("System.String"))
        dtHoraSabanaMedicamento.Columns.Add("Tipo", Type.GetType("System.String"))
        dtHoraSabanaNPT.Columns.Add("CodigoOrden", Type.GetType("System.String"))
        dtHoraSabanaNPT.Columns.Add("Nombre", Type.GetType("System.String"))
        dtHoraSabanaNPT.Columns.Add("Tipo", Type.GetType("System.String"))
        dtHorasabanaGoteo.Columns.Add("ConsecutivoInfusion", Type.GetType("System.String"))
        dtHorasabanaGoteo.Columns.Add("Nombre", Type.GetType("System.String"))

        nombreReporte = ConstantesHC.NOMBRE_PDF_SABANA_APLICACION
        vistaReporte = "VISTA_SABANA_DIAS"
        objReporte = New rptSabanaAplicacion
        moduloReporte = Constantes.REPORTE_HC
    End Sub
    Public Overrides Sub cargarCodigoSabana()
        Try
            Dim params As New List(Of String)
            Dim dr As DataRow
            params.Add(nRegistro)
            params.Add(Format(fechaReal.Date, Constantes.FORMATO_FECHA_GEN))
            dr = General.cargarItem(sqlCodigoSabanaCarga, params)
            If Not IsNothing(dr) Then
                codigoSabana = dr.Item(0)
            Else
                codigoSabana = ""
            End If
            usuario = SesionActual.idUsuario
            usuarioReal = usuario
            cargarIniciales()
        Catch ex As Exception
            MsgBox(ex.Message & "Error en cargarCodigoSabana")
        End Try

    End Sub
    Public Sub cargarIniciales()
        Try
            Dim params As New List(Of String)
            Dim dfila As DataRow
            params.Add(usuarioReal)
            params.Add(SesionActual.idEmpresa)
            dfila = General.cargarItem(ConsultasHC.EMPLEADO_INICIALES, params)
            inicialesUsuario = dfila.Item(0).ToString
        Catch ex As Exception
            MsgBox(ex.Message & "Error en cargarIniciales")
        End Try
    End Sub
    Public Sub cambiarUsuarioAutomatico()
        Try
            Dim params As New List(Of String)
            Dim dfila As DataRow
            params.Add(Constantes.LISTA_CARGO_NOTA_ENFERMERIA)
            params.Add(SesionActual.idEmpresa)
            dfila = General.cargarItem(Consultas.BUSQUEDA_EMPLEADO_HC_AUTOMATICO, params)
            usuarioReal = dfila.Item(0).ToString
        Catch ex As Exception
            MsgBox(ex.Message & "Error en cambiarUsuarioAutomatico")
        End Try

    End Sub
    Public Sub cargarPesoYBalanceSabana()

        Try
            Dim params As New List(Of String)
            Dim dr As DataRow
            params.Add(nRegistro)
            params.Add(fechaReal.Date)
            dr = General.cargarItem(sqlPesoBalanceCarga, params)
            pesoAnterior = IIf(String.IsNullOrEmpty(dr.Item(0).ToString), 0, dr.Item(0))
            pesoActual = IIf(String.IsNullOrEmpty(dr.Item(1).ToString), 0, dr.Item(1))
            balanceAnterior = IIf(String.IsNullOrEmpty(dr.Item(2).ToString), 0, dr.Item(2))
        Catch ex As Exception
            MsgBox(ex.Message & "Error en cargarPesoYBalanceSabana")
        End Try

    End Sub
    Public Sub calcularPesoGanado()
        pesoGanado = pesoActual - pesoAnterior
    End Sub
    Public Sub calcularBalanceHora()
        calcularInfusionesIV()
        sumaIngreso()
        sumaEgreso()
        If servicioNeonatal = True Then
            dtHoraSabanaBalance.Rows(2).Item(1) = String.Format("{0:N3}", IIf(pesoActual = 0, 0, ((CDbl(dtHoraSabanaBalance.Rows(1).Item(1)) / (pesoActual / 1000))) / If(numeroHorastranscurridas >= 6, 6, numeroHorastranscurridas)))
            dtHoraSabanaBalance.Rows(2).Item(2) = String.Format("{0:N3}", IIf(pesoActual = 0, 0, (CDbl(dtHoraSabanaBalance.Rows(1).Item(2)) / (pesoActual / 1000))) / If(numeroHorastranscurridas >= 12, 12, numeroHorastranscurridas))
            dtHoraSabanaBalance.Rows(2).Item(3) = String.Format("{0:N3}", IIf(pesoActual = 0, 0, (CDbl(dtHoraSabanaBalance.Rows(1).Item(3)) / (pesoActual / 1000))) / If(numeroHorastranscurridas >= 18, 18, numeroHorastranscurridas))
            dtHoraSabanaBalance.Rows(2).Item(4) = String.Format("{0:N3}", IIf(pesoActual = 0, 0, (CDbl(dtHoraSabanaBalance.Rows(1).Item(4)) / (pesoActual / 1000))) / numeroHorastranscurridas)
        Else
            dtHoraSabanaBalance.Rows(2).Item(1) = String.Format("{0:N3}", IIf(pesoActual = 0, 0, (CDbl(dtHoraSabanaBalance.Rows(1).Item(1)) / pesoActual)) / If(numeroHorastranscurridas >= 6, 6, numeroHorastranscurridas))
            dtHoraSabanaBalance.Rows(2).Item(2) = String.Format("{0:N3}", IIf(pesoActual = 0, 0, (CDbl(dtHoraSabanaBalance.Rows(1).Item(2)) / pesoActual)) / If(numeroHorastranscurridas >= 12, 12, numeroHorastranscurridas))
            dtHoraSabanaBalance.Rows(2).Item(3) = String.Format("{0:N3}", IIf(pesoActual = 0, 0, (CDbl(dtHoraSabanaBalance.Rows(1).Item(3)) / pesoActual)) / If(numeroHorastranscurridas >= 18, 18, numeroHorastranscurridas))
            dtHoraSabanaBalance.Rows(2).Item(4) = String.Format("{0:N3}", IIf(pesoActual = 0, 0, (CDbl(dtHoraSabanaBalance.Rows(1).Item(4)) / pesoActual)) / numeroHorastranscurridas)
        End If

        dtHoraSabanaBalance.Rows(3).Item(1) = String.Format("{0:N2}", CDbl(dtHoraSabanaBalance.Rows(0).Item(1)) - CDbl(dtHoraSabanaBalance.Rows(1).Item(1)))
        dtHoraSabanaBalance.Rows(3).Item(2) = String.Format("{0:N2}", CDbl(dtHoraSabanaBalance.Rows(0).Item(2)) - CDbl(dtHoraSabanaBalance.Rows(1).Item(2)))
        dtHoraSabanaBalance.Rows(3).Item(3) = String.Format("{0:N2}", CDbl(dtHoraSabanaBalance.Rows(0).Item(3)) - CDbl(dtHoraSabanaBalance.Rows(1).Item(3)))
        dtHoraSabanaBalance.Rows(3).Item(4) = String.Format("{0:N2}", CDbl(dtHoraSabanaBalance.Rows(0).Item(4)) - CDbl(dtHoraSabanaBalance.Rows(1).Item(4)))
        calcularBalanceDia()
    End Sub

    Friend Sub limpiarDts()
        dtHoraSabanaSilverman.Clear()
        dtHoraSabanaValoracion.Clear()
        dtHoraSabanaIngreso.Clear()
        dtHoraSabanaEgreso.Clear()
        dtHoraSabanaBalance.Clear()
        dtHoraSabanaAplicacion.Clear()
        dtHoraSabanaMedicamento.Clear()
        dtHoraSabanaNPT.Clear()
        dtHorasabanaGoteo.Clear()
    End Sub

    Private Sub calcularBalanceDia()
        ingresoTotal = CDbl(dtHoraSabanaBalance.Rows(0).Item(1)) +
                       CDbl(dtHoraSabanaBalance.Rows(0).Item(2)) +
                       CDbl(dtHoraSabanaBalance.Rows(0).Item(3)) +
                       CDbl(dtHoraSabanaBalance.Rows(0).Item(4))
        egresoTotal = CDbl(dtHoraSabanaBalance.Rows(1).Item(1)) +
                       CDbl(dtHoraSabanaBalance.Rows(1).Item(2)) +
                       CDbl(dtHoraSabanaBalance.Rows(1).Item(3)) +
                       CDbl(dtHoraSabanaBalance.Rows(1).Item(4))
        balanceActual = CDbl(dtHoraSabanaBalance.Rows(3).Item(1)) +
                       CDbl(dtHoraSabanaBalance.Rows(3).Item(2)) +
                       CDbl(dtHoraSabanaBalance.Rows(3).Item(3)) +
                       CDbl(dtHoraSabanaBalance.Rows(3).Item(4))

        balanceAcumulado = balanceAnterior + balanceActual
        If pesoActual = 0 Then
            gastoUrinario = 0
            Exit Sub
        End If

        If servicioNeonatal = True Then
            gastoUrinario = (egresoTotal / (pesoActual / 1000)) / numeroHorastranscurridas
        Else
            gastoUrinario = (egresoTotal / pesoActual) / numeroHorastranscurridas
        End If

    End Sub
    Private Sub calcularInfusionesIV()
        If dtHorasabanaGoteo.Rows.Count > 0 Then
            For I = 0 To 23
                dtHoraSabanaIngreso.Rows(ConstantesHC.POSICION_FILA_INFUSIONES_IV).Item("" & I & "") = dtHorasabanaGoteo.Compute("SUM([" & I & "])", "[" & I & "]>0")
            Next
        End If


    End Sub

    Private Sub sumaIngreso()
        For i = 0 To 23
            dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("" & i & "") = CDbl(sumaColumna(dtHoraSabanaIngreso, "" & i & ""))
        Next
        dtHoraSabanaBalance.Rows(0).Item(1) = CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("0")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("1")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("2")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("3")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("4")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("5"))
        dtHoraSabanaBalance.Rows(0).Item(2) = CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("6")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("7")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("8")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("9")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("10")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("11"))
        dtHoraSabanaBalance.Rows(0).Item(3) = CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("12")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("13")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("14")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("15")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("16")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("17"))
        dtHoraSabanaBalance.Rows(0).Item(4) = CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("18")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("19")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("20")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("21")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("22")) +
                                              CDbl(dtHoraSabanaIngreso.Rows(dtHoraSabanaIngreso.Rows.Count - 1).Item("23"))
    End Sub
    Private Sub sumaEgreso()
        For i = 0 To 23
            dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("" & i & "") = CDbl(sumaColumna(dtHoraSabanaEgreso, "" & i & ""))
        Next
        dtHoraSabanaBalance.Rows(1).Item(1) = CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("0")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("1")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("2")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("3")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("4")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("5"))
        dtHoraSabanaBalance.Rows(1).Item(2) = CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("6")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("7")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("8")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("9")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("10")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("11"))
        dtHoraSabanaBalance.Rows(1).Item(3) = CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("12")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("13")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("14")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("15")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("16")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("17"))
        dtHoraSabanaBalance.Rows(1).Item(4) = CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("18")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("19")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("20")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("21")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("22")) +
                                              CDbl(dtHoraSabanaEgreso.Rows(dtHoraSabanaEgreso.Rows.Count - 1).Item("23"))
    End Sub

    Private Function sumaColumna(dt As DataTable, columna As String)
        Dim total = (From a In dt.AsEnumerable()
                     Where Not String.IsNullOrEmpty(a.Field(Of String)("" & columna & "")) AndAlso a.Field(Of Integer)("codigo") <> ConstantesHC.CODIGO_TOTAL_INGRESO AndAlso a.Field(Of Integer)("codigo") <> ConstantesHC.CODIGO_TOTAL_EGRESO
                     Select Convert.ToDouble(a.Field(Of String)("" & columna & ""))).Sum()
        Return total
    End Function
    Public Overrides Sub cargarDetalle()
        cargarDetalleSignoVital()
        cargarDetalleSilverman()
        cargarDetalleValoracion()
        cargarDetalleIngreso()
        cargarDetalleEgreso()
        cargarDetalleBalance()
        cargarDetalleAplicacion()
        cargarDetalleGoteo()
    End Sub
    Public Sub cargarDetalleSignoVital()
        Dim params As New List(Of String)
        params.Add(ConstantesHC.CODIGO_DESCRIPCION_SIGNOS_VITALES)
        params.Add(codigoSabana)
        General.llenarTabla(sqlDetalleCarga, params, dtHorasabana)
    End Sub
    Public Sub cargarDetalleSilverman()
        Dim params As New List(Of String)
        params.Add(ConstantesHC.CODIGO_DESCRIPCION_SILVERMAN)
        params.Add(codigoSabana)
        General.llenarTabla(sqlDetalleCargaSilverman, params, dtHoraSabanaSilverman)
    End Sub
    Public Sub cargarDetalleValoracion()
        Dim params As New List(Of String)
        params.Add(ConstantesHC.CODIGO_DESCRIPCION_NEUROLOGIA)
        params.Add(codigoSabana)
        General.llenarTabla(sqlDetalleCargaValoracion, params, dtHoraSabanaValoracion)
    End Sub
    Public Sub cargarDetalleIngreso()
        Dim params As New List(Of String)
        params.Add(ConstantesHC.CODIGO_DESCRIPCION_INGRESO)
        params.Add(codigoSabana)
        General.llenarTabla(sqlDetalleCargaIngreso, params, dtHoraSabanaIngreso)
    End Sub
    Public Sub cargarDetalleEgreso()
        Dim params As New List(Of String)
        params.Add(ConstantesHC.CODIGO_DESCRIPCION_EGRESO)
        params.Add(codigoSabana)
        General.llenarTabla(sqlDetalleCargaEgreso, params, dtHoraSabanaEgreso)
    End Sub

    Public Sub cargarDetalleBalance()
        Dim params As New List(Of String)
        params.Add(ConstantesHC.CODIGO_DESCRIPCION_BALANCE)
        params.Add(codigoSabana)
        General.llenarTabla(sqlDetalleCargaBalance, params, dtHoraSabanaBalance)
    End Sub
    Public Sub cargarDetalleAplicacion()
        Dim params As New List(Of String)
        params.Add(nRegistro)
        params.Add(fechaReal.Date)
        params.Add(identificadorPredeterminado)
        General.llenarTabla(sqlDetalleCargaAplicacion, params, dtHoraSabanaAplicacion)
        cargarDetalleMedicamento()
        cargarDetalleNPT()
    End Sub

    Private Sub cargarDetalleNPT()
        dtHoraSabanaNPT.Clear()
        Dim filas As DataRow()
        filas = dtHoraSabanaAplicacion.Select("Tipo = '" & ConstantesHC.CODIGO_IDENTIFICADOR_NUTRICION & "' ")
        For Each row As DataRow In filas
            dtHoraSabanaNPT.ImportRow(row)
        Next
    End Sub

    Private Sub cargarDetalleMedicamento()
        dtHoraSabanaMedicamento.Clear()
        Dim filas As DataRow()
        filas = dtHoraSabanaAplicacion.Select("Tipo = '" & ConstantesHC.CODIGO_IDENTIFICADOR_MEDICAMENTO & "'  ")
        For Each row As DataRow In filas
            dtHoraSabanaMedicamento.ImportRow(row)
        Next
        filas = Nothing
        filas = dtHoraSabanaAplicacion.Select(" Tipo = '" & ConstantesHC.CODIGO_IDENTIFICADOR_IMPREGNACION & "' ")
        For Each row As DataRow In filas
            dtHoraSabanaMedicamento.ImportRow(row)
        Next
    End Sub

    Public Sub cargarDetalleGoteo()
        Dim params As New List(Of String)
        params.Add(nRegistro)
        params.Add(fechaReal.Date)
        General.llenarTabla(sqlDetalleCargaGoteo, params, dtHorasabanaGoteo)
    End Sub

    Public Overrides Sub guardarDetalle()
        cargarCodigoSabana()
        If String.IsNullOrEmpty(codigoSabana) Then
            SabanaAplicacionMedDAL.guardarSabanaHC(Me)
        Else
            SabanaAplicacionMedDAL.actualizarSabanaHC(Me)
        End If
    End Sub


    Public Sub guardarReporte2doPlano()
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre As String
            Dim objReporte2doPlano As New rptSabanaAplicacion
            Dim reporte As New ftp_reportes
            codigoNombre = nRegistro
            nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            formula = "{" & vistaReporte & ".N_Registro} = " & nRegistro & " 
                                            AND {" & vistaReporte & ".Modulo}=" & moduloReporte & ""
            ruta = Path.GetTempPath() & nombreArchivo & Now.Hour & Now.Minute & Now.Second & Now.Millisecond

            reporte.usuarioActual = usuario
            Dim params As New List(Of String)
            params.Add(pMostrarSignos)
            params.Add(nRegistro)
            reporte.crearReportePDF(nombreReporte, nRegistro, objReporte2doPlano,
                                    codigoNombre, formula,
                                    nombreReporte, IO.Path.GetTempPath,,, params)
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
        estadoGuardando = False
    End Sub

End Class
