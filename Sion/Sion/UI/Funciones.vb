'Option Explicit On
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Class Funciones
    Private cadena, Consulta As String
    Private CodigoBodega, CodigoProducto, NumeroLote, stoc_total As String
    Public Shared Property empresaActual As New Empresa

    Public Structure cnsctvos
        Public cdgo As String
        Public sma, stoc_total As Integer
        Public CodigoBodega, CodigoProducto, NumeroLote, Consulta As String
    End Structure

    Friend Shared Function Fecha(ByVal frmto As Integer) As String
        Dim fcha As String
        If Validar(frmto) Then
            fcha = solicita("SELECT CONVERT(varchar,GETDATE()," & frmto.ToString & ") AS Fecha", "Fecha")
            Return fcha
        Else
            Return "0"
        End If
    End Function


    Public Shared Sub validarCantidadCaracteresDGV(dgv As DataGridView, nombreColumna As String, cantidad As Integer)
        With dgv
            CType(.Columns(nombreColumna), DataGridViewTextBoxColumn).MaxInputLength = cantidad
        End With
    End Sub

    Public Shared Function solicita(ByVal cdna As String, ByVal itm As String) As String
        Dim vlor As String = ""
        Try

            Using cnslta = New SqlCommand(cdna, FormPrincipal.cnxion)
                Using rsltdo = cnslta.ExecuteReader()
                    If rsltdo.HasRows Then
                        rsltdo.Read()
                        vlor = rsltdo.Item(itm).ToString
                    End If
                End Using
            End Using
        Catch ex As Exception
            General.manejoErrores(ex)
            'Finally
            '    rsltdo.Close()
            '    FormPrincipal.cnxion.desconectarbd()
        End Try
        Return vlor
    End Function
    Public Shared Function getReporteNoFTP(pReporte As Object, pFormula As String, pArea As String, Optional ext As String = Constantes.FORMATO_PDF, Optional tbl As Hashtable = Nothing, Optional rutaAlterna As String = Nothing, Optional ejecutarReporte As Boolean = True) As Boolean
        Try
            FormPrincipal.Cursor = Cursors.WaitCursor
            General.getConnReporte(pReporte.Database.Tables)
            If Not IsNothing(tbl) Then
                For Each item As DictionaryEntry In tbl
                    pReporte.SetParameterValue(item.Key, item.Value)
                Next
            End If
            If pFormula IsNot Nothing Then pReporte.RecordSelectionFormula = pFormula
            Dim ruta As String
            If Not IsNothing(rutaAlterna) Then
                ruta = rutaAlterna
            Else
                ruta = IO.Path.GetTempPath & pArea & ConstantesHC.NOMBRE_PDF_SEPARADOR & Now.Ticks & ext
            End If

            Dim crformat As ExportFormatType
            Select Case ext
                Case Constantes.FORMATO_PDF : crformat = ExportFormatType.PortableDocFormat
                Case Constantes.FORMATO_WORD : crformat = ExportFormatType.WordForWindows
                Case Constantes.FORMATO_EXCEL : crformat = ExportFormatType.Excel
            End Select
            pReporte.ExportToDisk(crformat, ruta)
            pReporte.close()
            If ejecutarReporte Then
                Process.Start(ruta)
            End If

        Catch ex As Exception
            General.manejoErrores(ex)

            Return False
        Finally
            FormPrincipal.Cursor = Cursors.Default
        End Try
        Return True
    End Function

    Public Shared Sub getReporte(pReporte As ReportClass,
                                 pNombre As String,
                                 pFormula As String,
                                 pFormato As ExportFormatType)

        Try
            FormPrincipal.Cursor = Cursors.WaitCursor
            General.getConnReporte(pReporte.Database.Tables)

            If pFormula IsNot Nothing Then
                pReporte.RecordSelectionFormula = pFormula
            End If

            Dim ruta = IO.Path.GetTempPath &
                       pNombre &
                       ConstantesHC.NOMBRE_PDF_SEPARADOR &
                       Now.Ticks &
                       getReporteExtension(pFormato)

            pReporte.ExportToDisk(pFormato, ruta)
            pReporte.Load(ruta)
            pReporte.Close()
            Process.Start(ruta)

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            FormPrincipal.Cursor = Cursors.Default
        End Try
    End Sub

    Public Shared Function validarComillaSimple(ByVal busqueda As String) As String
        Return Replace(busqueda, "'", "")
    End Function

    Private Shared Function Validar(ByVal frmto As Integer) As Boolean
        If (0 <= frmto < 15) Or (99 < frmto < 115) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function getSexoDesc(ByVal sexoCod As Char) As String

        Return If(sexoCod = "M", "MASCULINO", "FEMENINO")
    End Function

    Private Function niveltriaje(ByVal codigo As Integer) As String
        Dim descripcion As String = ""
        Select Case codigo
            Case 1
                descripcion = "NIVEL I"
            Case 2
                descripcion = "NIVEL II"
            Case 3
                descripcion = "NIVEL III"
            Case 4
                descripcion = "NIVEL IV"
            Case 5
                descripcion = "NIVEL V"
        End Select
        Return descripcion
    End Function

    Public Shared Function getReporteExtension(pTipoArchivo As String) As String

        Select Case pTipoArchivo
            Case ExportFormatType.PortableDocFormat
                Return ConstantesHC.EXTENSION_ARCHIVO_PDF
            Case ExportFormatType.Excel
                Return ConstantesHC.EXTENSION_ARCHIVO_EXCEL
        End Select

        Return ConstantesHC.EXTENSION_ARCHIVO_PDF
    End Function


    Public Shared Function getParametros(lista As List(Of String)) As String

        If lista Is Nothing OrElse lista.Count = 0 Then 'OrElse lista.First Is Nothing Then
            Return String.Empty
        End If

        Dim comilla As String = Chr(39) 'comilla simple
        Dim separador As String = Chr(44) '(,)
        Dim listaParams As String

        listaParams = comilla & lista.First & comilla
        For i = 0 To lista.Count - 1

            If i > 0 Then
                listaParams += separador & comilla & lista(i) & comilla
            End If
        Next

        Return listaParams

    End Function

    'A diferencia de la funcion getParametros, esta concatena la palabra NULL cuando encuentra un parametro 
    'con valor Nothing
    Public Shared Function getParametrosWithNull(lista As List(Of String)) As String
        Const NULL_PARAM = "NULL"

        If lista Is Nothing OrElse lista.Count = 0 Then 'OrElse lista.First Is Nothing Then
            Return String.Empty
        End If

        Dim comilla As String = Chr(39) 'comilla simple
        Dim separador As String = Chr(44) '(,)
        Dim listaParams As String

        'Se determina el primer parametro
        If lista.First Is Nothing Then
            listaParams = NULL_PARAM
        Else
            listaParams = comilla & lista.First & comilla
        End If

        'Se recorren los demas parametros y se van concatenando a la lista
        For Each param As String In lista
            If param IsNot lista.First Then
                If param Is Nothing Then
                    listaParams += separador & NULL_PARAM
                Else
                    listaParams += separador & comilla & param & comilla
                End If
            End If
        Next

        Return listaParams

    End Function

    Public Shared Function getParametrosHistorial(lista As List(Of String)) As String

        If lista Is Nothing OrElse lista.Count = 0 OrElse lista.First Is Nothing Then
            Return String.Empty
        End If

        Dim comilla As String = Chr(39) 'comilla simple
        Dim separador As String = Chr(44) '(,)
        Dim listaParams As String

        listaParams = comilla & lista.First & comilla
        For Each param As String In lista

            If param IsNot lista.First Then
                If param <> Nothing Then
                    listaParams += separador & comilla & param & comilla
                End If
            End If
        Next

        Return listaParams

    End Function

    Public Shared Function textToNum(texto As String) As Integer
        If String.IsNullOrEmpty(texto) Then
            Return Nothing
        Else
            Return CInt(texto)
        End If

    End Function
    Public Shared Sub agregarFilas(ByRef dt As DataTable, ByVal nFilas As Integer)
        Try
            For indiceFilas As Int16 = 0 To nFilas - 1
                dt.Rows.Add()
            Next
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try

    End Sub
    Public Shared Function getNombreComida(ByVal codigoComida As Integer) As String
        Dim nombreComidaRetornado As String = ""
        Select Case codigoComida
            Case Constantes.NO_PUEDE
                nombreComidaRetornado = ""
            Case Constantes.DESAYUNO
                nombreComidaRetornado = "Desayuno"
            Case Constantes.ALMUERZO
                nombreComidaRetornado = "Almuerzo"
            Case Constantes.CENA
                nombreComidaRetornado = "Cena"
        End Select
        Return nombreComidaRetornado
    End Function
    Public Shared Function getCodigoComida(ByVal nombreComida As String) As Integer
        Dim codigoComidaRetornado As String = ""
        Select Case nombreComida
            Case ""
                codigoComidaRetornado = Constantes.NO_PUEDE
            Case "- - - Seleccione - - -"
                codigoComidaRetornado = Constantes.NO_PUEDE
            Case "Desayuno"
                codigoComidaRetornado = Constantes.DESAYUNO
            Case "Almuerzo"
                codigoComidaRetornado = Constantes.ALMUERZO
            Case "Cena"
                codigoComidaRetornado = Constantes.CENA
        End Select
        Return codigoComidaRetornado
    End Function
    Public Shared Function llenarMesesAnio() As SortedList
        Dim tblMes As New SortedList
        tblMes.Add(0, "-- Seleccione el mes --")
        tblMes.Add(1, "Enero")
        tblMes.Add(2, "Febrero")
        tblMes.Add(3, "Marzo")
        tblMes.Add(4, "Abril")
        tblMes.Add(5, "Mayo")
        tblMes.Add(6, "Junio")
        tblMes.Add(7, "Julio")
        tblMes.Add(8, "Agosto")
        tblMes.Add(9, "Septiempre")
        tblMes.Add(10, "Octubre")
        tblMes.Add(11, "Noviembre")
        tblMes.Add(12, "Diciembre")
        Return tblMes
    End Function
    Public Shared Function consultarEmpleadoNotaAudit(idEmpleado As String) As Integer
        Dim params As New List(Of String)
        Dim verificar As Integer
        Dim dFila As DataRow
        params.Add(idEmpleado)
        Try
            dFila = General.cargarItem("SP_NOTA_AUDITORIA_NOTIFICACION", params)
            verificar = dFila.Item(0)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return verificar
    End Function
    Public Shared Function consultarEmpleadoHojaRuta(idEmpleado As String, fecha As Date, registro As Integer) As Boolean
        Dim params As New List(Of String)
        Dim verificar As Integer
        Dim dFila As DataRow
        params.Add(idEmpleado)
        params.Add(fecha.Date)
        params.Add(registro)
        Try
            dFila = General.cargarItem("SP_HOJA_RUTA_NOTIFICACION", params)
            verificar = dFila.Item(0)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return verificar
    End Function
    Public Shared Function consultarPeriodoContratoActivo(idEmpleado As String, fechaInicio As Date, fechaFin As Date, empresa As Integer) As Boolean
        Dim params As New List(Of String)
        Dim verificar As Integer
        Dim dFila As DataRow
        params.Add(idEmpleado)
        params.Add(fechaInicio.Date)
        params.Add(fechaFin)
        params.Add(empresa)
        Try
            dFila = General.cargarItem("[SP_VALIDAR_PERIODO_CONTRATO]", params)
            verificar = dFila.Item(0)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return verificar
    End Function
    Public Shared Function consultarCodigoTipoExamen(CodigoProcedimiento As String) As Integer
        Dim params As New List(Of String)
        Dim codigo As Integer
        Dim dFila As DataRow
        params.Add(CodigoProcedimiento)
        Try
            dFila = General.cargarItem(ConsultasHC.CODIGO_TIPO_EXAMEN_CONSULTAR, params)
            codigo = dFila.Item(0)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return codigo
    End Function
    Public Shared Function consultarCodigoAreaServicio(registro As Integer) As Integer
        Dim params As New List(Of String)
        Dim codigo As Integer
        Dim dFila As DataRow
        params.Add(registro)
        Try
            dFila = General.cargarItem(ConsultasHC.CODIGO_AREA_SERVICIO_CONSULTAR, params)
            codigo = dFila.Item(0)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return codigo
    End Function
    Public Shared Function consultarAplicaParaclinico(codigoTipoExamen As Integer) As Integer
        Dim params As New List(Of String)
        Dim codigo As Integer
        Dim dFila As DataRow
        params.Add(codigoTipoExamen)
        Try
            dFila = General.cargarItem(ConsultasHC.CODIGO_APLICA_PARACLINICO, params)
            codigo = dFila.Item(0)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return codigo
    End Function

    Public Shared Function consultarFormulario(formulario As String) As Boolean
        Dim consultar As Boolean
        Dim dFila As DataRow
        Try
            dFila = General.cargarItem(Replace(ConsultasHC.CONSULTAR_FORMULARIO, "$", formulario), Nothing)
            consultar = dFila.Item(0)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return consultar
    End Function


    Public Shared Function consultaUltimoCodImage() As Integer
        Dim vlor As Integer
        Try
            Using cnslta = New SqlCommand(Consultas.CONSULTA_ULTIMO_COD_IMAGEN, FormPrincipal.cnxion)
                Using rsltdo = cnslta.ExecuteReader()
                    If rsltdo.HasRows Then
                        rsltdo.Read()
                        vlor = rsltdo.Item("Consecutivo").ToString
                    End If
                End Using
            End Using
        Catch ex As Exception
            General.manejoErrores(ex)
        End Try
        Return vlor
    End Function

    Public Shared Function convertirImagenToByte(imagen As Image) As Byte()
        Using memoryStream = New MemoryStream()
            imagen.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png)
            Return memoryStream.ToArray()
        End Using

    End Function
    Public Shared Function getCodigoTipoPedido(ByVal tipo As String) As Integer
        If tipo = "Pedido Farmacia" Then
            Return Constantes.TIPO_PEDIDO_FARMACIA
        Else
            Return Constantes.TIPO_PEDIDO_LABORATORIO
        End If
    End Function
    Public Function getTipoPedido(ByVal tipo As Integer) As String
        If tipo = Constantes.TIPO_PEDIDO_FARMACIA Then
            Return "Pedido Farmacia"
        Else
            Return "Solicitud Laboratorio"
        End If
    End Function
    Public Shared Function obtenerValorPredeterminado(ByVal codigoPredeterminado As String) As String
        Dim params As New List(Of String)
        params.Add(codigoPredeterminado)
        params.Add(SesionActual.codigoEP)
        Dim fila As DataRow = General.cargarItem("[SP_OBTENER_VALOR_PREDETERMINADO]", params)
        If Not IsNothing(fila) Then
            Return fila(0)
        Else
            Return ""
        End If
    End Function
#Region "Inventario"
    Public Shared Function validarCeldaActiva(ByRef dgv As DataGridView,
                                              ByRef e As DataGridViewCellEventArgs,
                                              ByVal nombreColumna As String,
                                              ByRef btn As ToolStripButton) As Boolean
        dgv.ReadOnly = False
        If btn.Enabled = True AndAlso verificacionPosicionActual(dgv, e, nombreColumna) Then
            Return True
        End If
        Return False
    End Function
    Public Shared Function verificacionPosicionActual(ByRef dgv As DataGridView,
                                                      ByRef e As DataGridViewCellEventArgs,
                                                      ByVal nombreColumna As String)
        If filaValida(e.RowIndex) AndAlso dgv.Columns(nombreColumna).Index = e.ColumnIndex Then
            Return True
        End If
        Return False
    End Function
    Public Shared Function validarColumnaActual(ByVal indiceColumna As Integer,
                                                ByVal fila As Integer,
                                                ByVal nombreColumna As String,
                                                ByRef dgv As DataGridView)
        If filaValida(fila) AndAlso indiceColumna = dgv.Columns(nombreColumna).Index Then
            Return True
        End If
        Return False
    End Function
    Public Shared Function filaValida(ByVal fila As Integer) As Boolean
        If fila < 0 Then
            Return False
        End If
        Return True
    End Function
    Public Shared Function isFilaActiva(ByVal contenido As String) As Boolean
        If contenido = "" Then
            Return False
        End If
        Return True
    End Function
    Public Shared Function nombrarTabla(ByVal indiceFila As Int16,
                                        ByVal codigo As String) As String

        Return "Tabla-" & indiceFila & "-" & codigo
    End Function
    Public Shared Function contarColumnasTablas(ByVal tbl As DataTable) As Integer
        Return tbl.Columns.Count
    End Function
    Public Shared Function validarEntrada(ByVal numeroComparado As Integer,
                                          ByVal numeroComparador As Integer,
                                          ByVal msm As String) As Integer

        If numeroComparador > numeroComparado Then
            MsgBox(msm, MsgBoxStyle.Exclamation)
            Return 0
        End If
        Return numeroComparador
    End Function

    Public Shared Function indiceValidoListas(ByVal indice As Integer) As Boolean
        If indice <= 0 Then
            Return False
        End If
        Return True
    End Function

    Public Shared Function quitarFilas(ByRef tablaReferencia As DataTable, ByVal indiceFila As Integer) As Boolean
        If MsgBox(Mensajes.QUITAR, MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Return False
        Else
            If tablaReferencia.Rows.Count > 1 Then
                tablaReferencia.Rows.RemoveAt(indiceFila)
            Else
                MsgBox("¡El detalle debe tener por lo menos 1 registro!", MsgBoxStyle.Exclamation)
            End If
        End If
        Return True
    End Function
    Public Shared Function quitarFilasSinRestriccion(ByRef tablaReferencia As DataTable, ByVal indiceFila As Integer) As Boolean
        If MsgBox(Mensajes.QUITAR, MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Return False
        Else
            tablaReferencia.Rows.RemoveAt(indiceFila)
        End If
        Return True
    End Function
    Public Shared Sub obtenerDatosEmpresaActual(idEmpresa As String)

        Dim drEmpresa As DataRow
        Dim params As New List(Of String)
        params.Add(idEmpresa)

        Try
            drEmpresa = General.cargarItem("SP_EMPRESA_CONSULTAR", params)

            empresaActual.dv = drEmpresa.Item(1).ToString
            empresaActual.nit = drEmpresa.Item(2).ToString
            empresaActual.codHabilitacion = drEmpresa.Item(3).ToString
            empresaActual.razonSocial = drEmpresa.Item(4).ToString
            empresaActual.direccion = drEmpresa.Item(5).ToString
            empresaActual.telefono = drEmpresa.Item(6).ToString
            empresaActual.celular = drEmpresa.Item(7).ToString
            empresaActual.codPais = drEmpresa.Item(8).ToString
            empresaActual.codDepartamento = drEmpresa.Item(9).ToString
            empresaActual.codMunicipio = drEmpresa.Item(10).ToString
            empresaActual.email = drEmpresa.Item(11).ToString
            empresaActual.portalWeb = drEmpresa.Item(12).ToString
            empresaActual.sigla = drEmpresa.Item(13).ToString
            empresaActual.encabezadoFactura = drEmpresa.Item(14).ToString
            empresaActual.pieFactura = drEmpresa.Item(15).ToString
            empresaActual.logoFactura = drEmpresa.Item(16)
            empresaActual.logoPrincipal = drEmpresa.Item(17)
            empresaActual.idRepresentante = drEmpresa.Item(18).ToString

        Catch ex As Exception
            General.manejoErrores(ex)
        End Try


    End Sub
    Public Shared Function totalizarColumna(ByRef tabla As DataTable,
                                            ByVal nombreColumna As String) As Double

        Dim suma As String = "SUM('" & nombreColumna & "')"
        Dim filtro As String = nombreColumna & " > 0"
        Return If(IsDBNull(tabla.Compute(suma, filtro)), 0, tabla.Compute(suma, filtro))
    End Function

    Public Shared Function castFromDbItemVacio(ByVal DbItem As Object) As Object
        If IsDBNull(DbItem) Then
            Return ""
        Else
            Return DbItem
        End If
    End Function


    Public Shared Function castFromDbItem(ByVal DbItem As Object) As Object
        If IsDBNull(DbItem) Then
            Return Nothing
        Else
            Return DbItem
        End If
    End Function

    Public Shared Function castToSqlInt32(ByVal item As Integer?) As SqlInt32
        If item Is Nothing Then
            Return SqlTypes.SqlInt32.Null
        Else
            Return item
        End If
    End Function
    Public Shared Function formatComboItem(ByVal item As Object) As Object
        If item Is Nothing Then
            Return Constantes.COMBO_VALOR_PREDETERMINADO
        Else
            Return item
        End If
    End Function

    Public Shared Function formatMoneda(valor As String) As String

        If valor <> Nothing AndAlso valor <> String.Empty AndAlso valor > 0 AndAlso IsNumeric(valor) Then
            Return CDbl(valor).ToString("C2")
        End If

        Return valor
    End Function

#End Region
    Public Shared Function verificaExistenciaRecaudo(comprobante As String) As String
        Dim params As New List(Of String)
        Dim codigo As String = Nothing
        Dim dFila As DataRow
        params.Add(comprobante)
        Try
            dFila = General.cargarItem(ConsultasContabilidad.CONSULTAR_EXISTENCIA_RECAUDO, params)
            If Not IsNothing(dFila) Then
                codigo = dFila.Item(0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return codigo
    End Function
    Public Shared Function consultaContratoFecha(idEmpleado As Integer,
                                                 idEmpresa As Integer,
                                                 fechaInicio As Date,
                                                 fechaFin As Date) As Boolean
        Dim params As New List(Of String)
        Dim codigo As Boolean
        Dim dFila As DataRow
        params.Add(idEmpresa)
        params.Add(idEmpleado)
        params.Add(fechaInicio)
        params.Add(fechaFin)
        Try
            dFila = General.cargarItem("[SP_CONTRATO_CONSULTAR_FECHA_VIGENTE]", params)
            If Not IsNothing(dFila) Then
                codigo = dFila.Item(0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return codigo
    End Function
    Public Shared Function consultaExistenciaAntibiograma(codigoOrden As Integer,
                                                          codigoProcedimiento As String,
                                                          modulo As String) As Boolean
        Dim params As New List(Of String)
        Dim codigo As Boolean
        Dim dFila As DataRow
        params.Add(codigoOrden)
        params.Add(codigoProcedimiento)
        params.Add(modulo)
        Try
            dFila = General.cargarItem(Consultas.CONSULTAR_EXISTENCIA_ANTIBIOGRAMA, params)
            If Not IsNothing(dFila) Then
                codigo = dFila.Item(0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return codigo
    End Function
    Public Shared Function esCorreoValido(ByVal email As String) As Boolean
        If email = String.Empty Then Return False
        ' Compruebo si el formato de la dirección es correcto.
        Dim re As New Regex("^[\w._%-]+@[\w.-]+\.[a-zA-Z]{2,4}$")
        Dim m As Match = re.Match(email)
        Return (m.Captures.Count <> 0)
    End Function
    Public Shared Function enviarCorreo(Optional ByVal correo As Correo = Nothing) As Boolean
        Try

            Using Smtp_Server As New SmtpClient
                Using e_mail As New MailMessage()
                    Smtp_Server.UseDefaultCredentials = False
                    Smtp_Server.Credentials = New Net.NetworkCredential(correo.correo, correo.pass)
                    Smtp_Server.Port = correo.puerto
                    Smtp_Server.EnableSsl = True
                    Smtp_Server.Host = correo.configuracionSMTP

                    Dim archivos() As String = Directory.GetFiles(correo.rutaArchivo & "\")
                    For Each archivo In archivos
                        Dim archivoAdjunto As New System.Net.Mail.Attachment(archivo)
                        e_mail.Attachments.Add(archivoAdjunto)
                    Next

                    e_mail.From = New MailAddress(correo.correo)
                    'e_mail.To.Add(correoDestino)
                    e_mail.To.Add(correo.correoDestinatario)
                    e_mail.Subject = correo.asunto
                    e_mail.IsBodyHtml = False
                    e_mail.Body = correo.cuerpo
                    Smtp_Server.Send(e_mail)
                End Using
            End Using
        Catch error_t As Exception
            Return False
        End Try

        Return True
    End Function
    Public Shared Function enviarCorreo(ByVal correo As Correo, rutaArchivo As String) As Boolean
        Try

            Using Smtp_Server As New SmtpClient
                Using e_mail As New MailMessage()
                    Smtp_Server.UseDefaultCredentials = False
                    Smtp_Server.Credentials = New Net.NetworkCredential(correo.correo, correo.pass)
                    Smtp_Server.Port = correo.puerto
                    Smtp_Server.EnableSsl = False
                    Smtp_Server.Host = correo.configuracionSMTP

                    Dim archivoAdjunto As New System.Net.Mail.Attachment(rutaArchivo)
                    e_mail.Attachments.Add(archivoAdjunto)

                    e_mail.From = New MailAddress(correo.correo)
                    'e_mail.To.Add(correoDestino)
                    e_mail.To.Add(correo.correoDestinatario)
                    e_mail.Subject = correo.asunto
                    e_mail.IsBodyHtml = False
                    e_mail.Body = correo.cuerpo
                    Smtp_Server.Send(e_mail)
                End Using
            End Using
        Catch error_t As Exception
            Return False
        End Try

        Return True
    End Function
    Public Shared Function crearCarpeta(ByVal codigo As String) As String
        Try
            Dim ruta As String = Path.GetTempPath & codigo
            If Directory.Exists(ruta) Then
                System.IO.Directory.Delete(ruta, True)
            End If
            System.IO.Directory.CreateDirectory(ruta)
            Return ruta
        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Public Shared Function verificarServidor50Only() As Boolean
        Dim MyPing As New System.Net.NetworkInformation.Ping
        Dim Myreply As System.Net.NetworkInformation.PingReply
        Myreply = MyPing.Send("192.168.3.5")
        If Myreply.Buffer.Length > 0 Then
            Return True
        End If
        Return False
    End Function
    Public Shared Function consultaConvencion(convencion As String) As String
        Dim params As New List(Of String)
        Dim descripcion As String = Nothing
        Dim dFila As DataRow
        params.Add(convencion)
        Try
            dFila = General.cargarItem("[SP_VERIFICAR_CONVENCION]", params)
            descripcion = dFila.Item(0)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return descripcion
    End Function
End Class
