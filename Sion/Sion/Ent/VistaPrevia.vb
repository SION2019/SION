Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.Threading
Public Class VistaPrevia
    Inherits PdfPageEventHelper
    Public Property registro As Integer
    Public Property codigoMenu As String
    Public Property usuario As Integer
    Public Property listDocumentos As DataTable
    Public Property listDocumentosBD As DataTable
    Public Property rutaArchivos As New DataTable
    Public Property listaCarga As String
    Public Property docCarga As String
    Public Property codigoReporte As Integer
    Public Property nombrePaciente As String
    Private direccion As String
    Private rutaUnion As String
    Public Property rutaDefinitiva As String
    Public Property objConsolidado As Consolidado
    Public Property formPlataforma As FormVisorPlataformaDoc
    Public Property reporteExame As ReportClass
    Public Property erroresEncontrados As String
    Private dtFamiliaRegistro As New DataTable
    Public Sub New()
        listaCarga = ConsultasHC.VISTA_PREVIA_DOC_LISTAR
        docCarga = ConsultasHC.VISTA_PREVIA_DOC_BD_LISTAR
        codigoMenu = Constantes.CODIGO_MENU_HISTC
        listDocumentos = New DataTable
        listDocumentos.Columns.Add("Bandera", Type.GetType("System.String"))
        listDocumentos.Columns.Add("Direccion", Type.GetType("System.String"))
        listDocumentosBD = New DataTable
        listDocumentosBD.Columns.Add("Bandera", Type.GetType("System.String"))
        listDocumentosBD.Columns.Add("Archivo", Type.GetType("System.Byte[]"))
        codigoReporte = Constantes.REPORTE_HC
        reporteExame = If(SesionActual.codigoEnlace = 2, New rptExamenParaclinico_ReyFalse, New rptExamenParaclinico)

    End Sub
    Public Sub cargarDocumentos(ByRef vConsolidado As Consolidado, ByRef pformPlataforma As FormVisorPlataformaDoc)
        Try
            erroresEncontrados = ""
            formPlataforma = pformPlataforma
            objConsolidado = vConsolidado
            Dim params As New List(Of String)
            params.Add(registro)
            General.llenarTabla(ConsultasHC.FAMILIA_REGISTRO_LISTAR, params, dtFamiliaRegistro)
            crearCarpetaDetalle()
            generarDocumentos()
            Dim contador As Integer
            Do
                contador = 1
            Loop While objConsolidado.Generando2doPlanoParte1 = True OrElse objConsolidado.Generando2doPlanoParte2 = True OrElse
                       objConsolidado.Generando2doPlanoParte3 = True OrElse objConsolidado.Generando2doPlanoParte4 = True

            unirDocumentos()
            eliminarCarpetaDetalle()
            objConsolidado.dtProcesos.Rows(7).Item(1) = 1
            objConsolidado.Generando = False
        Catch ex As Exception
            objConsolidado.Generando = False
            objConsolidado.errorGeneracion = True
            erroresEncontrados += "Error en cargarDocumentos " & ex.Message
            MsgBox("Intente nuevamente o comuníquese con sistemas. " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
    Public Overridable Sub crearCarpetaDetalle()
        direccion = Path.GetTempPath & Replace(nombrePaciente, " ", "")
        eliminarCarpetaDetalle()
        Do
            Directory.CreateDirectory(direccion)
            Threading.Thread.Sleep(Constantes.TIEMPO_ESPERA)
        Loop While Not Directory.Exists(direccion)
        objConsolidado.dtProcesos.Rows(0).Item(1) = 1
    End Sub
    Public Overridable Sub unirDocumentos()
        Try
            Dim listaRutas As New List(Of String)
            organizarLista(listaRutas)
            crearCarpetaConsolidado()
            Dim Doc As New Document()
            Dim fs As New FileStream(rutaUnion, FileMode.Create, FileAccess.Write, FileShare.None)
            Dim copy As New PdfCopy(Doc, fs)
            Doc.Open()
            Dim archivoPDF As PdfReader
            Dim archivoImg As PdfImage
            Dim numeroPagina As Integer 'Número de páginas de cada pdf
            Dim numeroPaginaActual, totalNumeroPaginas As Integer
            For Each pdf2 In listaRutas
                If Not File.Exists(pdf2) Then
                Else
                    Try
                        archivoPDF = New PdfReader(pdf2)
                        totalNumeroPaginas = totalNumeroPaginas + archivoPDF.NumberOfPages
                        archivoPDF.Close()
                    Catch ex As Exception
                    End Try
                End If
            Next

            For Each pdf2 In listaRutas
                If Not File.Exists(pdf2) Then
                    If MsgBox("Hubo un problema al momento de cargar algun soporte, aceptar: ignorar,  cancelar: cargar de nuevo", MsgBoxStyle.OkCancel, "Informaciòn") = MsgBoxResult.Ok Then
                    Else
                        Exit Sub
                    End If
                Else
                    Try
                        'If Not pdf2.ToString.Contains(".pdf") Then
                        '    Doc.Add(iTextSharp.text.Image.GetInstance(pdf2))
                        'Else
                        archivoPDF = New PdfReader(enumerarPaginasPDF(pdf2, numeroPaginaActual, totalNumeroPaginas))
                        numeroPagina = archivoPDF.NumberOfPages
                        Dim pagina As Integer = 0
                        Do While pagina < numeroPagina
                            pagina += 1
                            copy.AddPage(copy.GetImportedPage(archivoPDF, pagina))
                        Loop
                        copy.FreeReader(archivoPDF)
                        archivoPDF.Close()
                        'End If
                    Catch ex As Exception
                    End Try
                End If
            Next
            Doc.Close()
            Doc.Dispose()
            copy.Close()
            copy.Dispose()
            formPlataforma.obtenerRuta(rutaUnion)
        Catch ex As Exception
            Throw ex
        End Try
        objConsolidado.dtProcesos.Rows(5).Item(1) = 1
        objConsolidado.dtProcesos.Rows(6).Item(1) = 1
    End Sub
    Private Function enumerarPaginasPDF(ruta As String, ByRef numeroPaginaActual As Integer, totalNumeroPaginas As Integer) As String
        ' Creamos una lista de archivos para concatenar
        Dim oldFile As String = ruta
        Dim newFile As String = oldFile
        'Dim newFile As String = Replace(ruta, ".pdf", "Paginado.pdf")

        '' Create reader
        'Dim reader As New PdfReader(oldFile)
        'Dim size As Rectangle = reader.GetPageSizeWithRotation(1)
        'Dim document As New Document(size)

        '' Create the writer
        'Dim fs As New FileStream(newFile, FileMode.Create, FileAccess.Write)
        'Dim writer As PdfWriter = PdfWriter.GetInstance(document, fs)
        'document.Open()
        'Dim cb As PdfContentByte = writer.DirectContent
        ''document.SetPageSize(iTextSharp.text.PageSize.A4.Rotate())
        '' Set the font, color and size properties for writing text to the PDF
        'Dim bf As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)

        '' Write text in the PDF
        'Dim n As Integer = reader.NumberOfPages
        'Dim pagina As Integer
        'Do While pagina < n
        '    cb.SetColorFill(BaseColor.BLACK)
        '    cb.SetFontAndSize(bf, 8)
        '    pagina += 1
        '    numeroPaginaActual += 1

        '    ' Put the text on a new page in the PDF 
        '    Dim page As PdfImportedPage = writer.GetImportedPage(reader, pagina)
        '    cb.AddTemplate(page, 0, 0)
        '    Dim text As String = "Folio " & numeroPaginaActual & " de " & totalNumeroPaginas
        '    cb.BeginText()
        '    ' Set the alignment and coordinates here
        '    cb.ShowTextAligned(1, text, 50, 20, 0)
        '    cb.EndText()

        '    document.NewPage()
        'Loop
        '' Close the objects
        'document.Close()
        'document.Dispose()
        'fs.Dispose()
        'writer.Dispose()
        'reader.Dispose()
        Return newFile
    End Function

    Public Overridable Sub eliminarCarpetaDetalle()

        While Directory.Exists(direccion)
            My.Computer.FileSystem.DeleteDirectory(direccion, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            Threading.Thread.Sleep(Constantes.TIEMPO_ESPERA + 2000)
        End While

    End Sub

    Public Overridable Sub crearCarpetaConsolidado()
        rutaUnion = direccion & ConstantesHC.CARPETA_VISTA_PREVIA

        While Directory.Exists(rutaUnion)
            My.Computer.FileSystem.DeleteDirectory(rutaUnion, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            Threading.Thread.Sleep(Constantes.TIEMPO_ESPERA + 1000)
        End While
        Do
            Directory.CreateDirectory(direccion & ConstantesHC.CARPETA_VISTA_PREVIA)
            Threading.Thread.Sleep(Constantes.TIEMPO_ESPERA)
        Loop While Not Directory.Exists(direccion & ConstantesHC.CARPETA_VISTA_PREVIA)
        rutaUnion = direccion & ConstantesHC.CARPETA_VISTA_PREVIA & ConstantesHC.CARPETA_VISTA_PREVIA_ARCHIVO & registro & ConstantesHC.NOMBRE_PDF_SEPARADOR & Format(Now, "HH'h'mm'm'ss's'") & ConstantesHC.EXTENSION_ARCHIVO_PDF
    End Sub

    Public Overridable Sub organizarLista(listaRutas As List(Of String))
        Dim dtLista As New DataTable
        General.llenarTabla(listaCarga, Nothing, dtLista)
        For j = 0 To dtLista.Rows.Count - 1
            For i = 0 To listDocumentos.Rows.Count - 1
                If dtLista.Rows(j).Item("Bandera").ToString.ToUpper = listDocumentos.Rows(i).Item("Bandera").ToString.ToUpper Then
                    listaRutas.Add(listDocumentos.Rows(i).Item("Direccion").ToString.Trim)
                End If
            Next
        Next

    End Sub
    Private Sub generarDocumentos()

        Dim guardarEn2doPlano, guardarEn2doPlano2, guardarEn2doPlano3, guardarEn2doPlano4 As Thread
        listDocumentos.Clear()

        guardarEn2doPlano = New Thread(AddressOf crearParte1)
        guardarEn2doPlano.Name = "Vista Previa parte1"
        guardarEn2doPlano.SetApartmentState(ApartmentState.STA)
        guardarEn2doPlano.Start()
        guardarEn2doPlano2 = New Thread(AddressOf crearParte2)
        guardarEn2doPlano2.Name = "Vista Previa parte2"
        guardarEn2doPlano2.SetApartmentState(ApartmentState.STA)
        guardarEn2doPlano2.Start()
        guardarEn2doPlano3 = New Thread(AddressOf crearParte3)
        guardarEn2doPlano3.Name = "Vista Previa parte3"
        guardarEn2doPlano3.SetApartmentState(ApartmentState.STA)
        guardarEn2doPlano3.Start()
        guardarEn2doPlano4 = New Thread(AddressOf crearParte4)
        guardarEn2doPlano4.Name = "Vista Previa parte4"
        guardarEn2doPlano4.SetApartmentState(ApartmentState.STA)
        guardarEn2doPlano4.Start()
    End Sub
    Public Overridable Sub crearParte1()
        Try
            Dim listReportes As New List(Of DocumentoVistaPrevia)

            objConsolidado.Generando2doPlanoParte1 = True
            Dim reporte As New ftp_reportes
            Dim paso As String
            Dim errores As New List(Of String)
            Dim raiz As String
            crearCateterismo(listReportes)
            crearOrdenMedica(listReportes)
            crearImagenologia(listReportes)
            crearFactura(listReportes)
            crearFacturaCTC(listReportes)

            For i = 0 To listReportes.Count - 1
                paso = listReportes.Item(i).nombrePDF & " - " & listReportes.Item(i).formula
                Try
                    Dim params As New List(Of String)
                    params.Add(listReportes.Item(i).codigoAdicional)
                    params.Add(listReportes.Item(i).codigoAdicional2)
                    raiz = direccion & "\" & listReportes.Item(i).nombrePDF & ConstantesHC.NOMBRE_PDF_SEPARADOR & listReportes.Item(i).codigoAdicional & ".pdf"
                    reporte.crearReportePDF(listReportes.Item(i).nombrePDF, registro, listReportes.Item(i).reporte,
                                           listReportes.Item(i).codigoAdicional, listReportes.Item(i).formula,
                                           listReportes.Item(i).area, direccion,, True, params, raiz)
                    listDocumentos.Rows.Add(listReportes.Item(i).nombrePDF, raiz)
                Catch ex As Exception
                    errores.Add(paso)
                End Try
            Next
            If errores.Count > 0 Then
                paso = ""
                For i = 0 To errores.Count - 1
                    paso = paso & errores.Item(i) & vbCrLf
                Next
                erroresEncontrados += vbCrLf & "Error en Consolidado: parte 1. " & paso
            End If
            objConsolidado.Generando2doPlanoParte1 = False
        Catch ex As Exception

            objConsolidado.Generando2doPlanoParte1 = False
            erroresEncontrados += vbCrLf & "Error en Consolidado: parte 1. " & ex.Message
        End Try
        objConsolidado.dtProcesos.Rows(1).Item(1) = 1
    End Sub

    Public Overridable Sub crearParte2()
        Try

            Dim listReportes As New List(Of DocumentoVistaPrevia)

            objConsolidado.Generando2doPlanoParte2 = True
            Dim inicioExamenes As Integer
            crearElectrolitos(listReportes)
            Dim reporte As New ftp_reportes
            Dim params As New List(Of String)
            Dim paso As String
            Dim errores As New List(Of String)
            Dim raiz As String
            For i = inicioExamenes To listReportes.Count - 1
                paso = listReportes.Item(i).nombrePDF & " - " & listReportes.Item(i).formula
                Try
                    params.Clear()
                    params.Add(listReportes.Item(i).codigoAdicional)
                    params.Add(listReportes.Item(i).codigoAdicional2)
                    params.Add(listReportes.Item(i).codigoAdicional3)
                    raiz = direccion & "\" & listReportes.Item(i).nombrePDF & ConstantesHC.NOMBRE_PDF_SEPARADOR & listReportes.Item(i).codigoAdicional & ".pdf"
                    reporte.crearReportePDF(listReportes.Item(i).nombrePDF, registro, listReportes.Item(i).reporte,
                                                listReportes.Item(i).codigoAdicional, listReportes.Item(i).formula,
                                                listReportes.Item(i).area, direccion, listReportes.Item(i).codigoAdicional2, True, params, raiz)
                    listDocumentos.Rows.Add(listReportes.Item(i).nombrePDF, raiz)
                Catch ex As Exception
                    errores.Add(paso)
                End Try
            Next
            If errores.Count > 0 Then
                paso = ""
                For i = 0 To errores.Count - 1
                    paso = paso & errores.Item(i) & vbCrLf
                Next
                erroresEncontrados += vbCrLf & "Error en Consolidado: parte 3. " & paso
            End If
            objConsolidado.Generando2doPlanoParte2 = False
        Catch ex As Exception
            objConsolidado.Generando2doPlanoParte2 = False
            erroresEncontrados += vbCrLf & "Error en Consolidado: parte 2. " & ex.Message
        End Try
        objConsolidado.dtProcesos.Rows(2).Item(1) = 1
    End Sub
    Public Overridable Sub crearParte3()
        objConsolidado.Generando2doPlanoParte3 = True
        Try
            Dim listReportes As New List(Of DocumentoVistaPrevia)
            Dim paso As String
            Dim errores As New List(Of String)
            crearAdmision(listReportes)
            crearCertificados(listReportes)
            crearAnexo1(listReportes)
            crearAnexo2(listReportes)
            crearAnexo3(listReportes)
            crearConsentimiento(listReportes)
            crearRemision(listReportes)
            crearEpicrisis(listReportes)
            crearSalidaVoluntaria(listReportes)
            crearIngresoMedico(listReportes)
            crearEvolucionMedica(listReportes)
            crearHojaFisio(listReportes)
            crearHojaNeb(listReportes)
            crearNotaFisio(listReportes)
            crearInformeQX(listReportes)
            crearSolicitudNoPos(listReportes)
            crearIIH(listReportes)
            crearInterconsulta(listReportes)
            crearHemodialisis(listReportes)
            crearTransfusionSanguinea(listReportes)
            crearEco(listReportes)
            crearEcostress(listReportes)
            crearHojaGlucometria(listReportes)
            Dim reporte As New ftp_reportes
            Dim raiz As String
            For i = 0 To listReportes.Count - 1
                paso = listReportes.Item(i).nombrePDF & " - " & listReportes.Item(i).formula
                Try
                    raiz = direccion & "\" & listReportes.Item(i).nombrePDF & ConstantesHC.NOMBRE_PDF_SEPARADOR & listReportes.Item(i).codigoAdicional & ".pdf"
                    reporte.crearReportePDF(listReportes.Item(i).nombrePDF, registro, listReportes.Item(i).reporte,
                                            listReportes.Item(i).codigoAdicional, listReportes.Item(i).formula,
                                           listReportes.Item(i).area, direccion,, True,, raiz)
                    listDocumentos.Rows.Add(listReportes.Item(i).nombrePDF, raiz)
                Catch ex As Exception
                    errores.Add(paso)
                End Try
            Next
            descargarDocumentos(paso, errores)
            If errores.Count > 0 Then
                paso = ""
                For i = 0 To errores.Count - 1
                    paso = paso & errores.Item(i) & vbCrLf
                Next
                erroresEncontrados += vbCrLf & "Error en Consolidado: parte 3. " & paso
            End If

            objConsolidado.Generando2doPlanoParte3 = False
        Catch ex As Exception
            objConsolidado.Generando2doPlanoParte3 = False
            erroresEncontrados += vbCrLf & "Error en Consolidado: parte 3. " & ex.Message
        End Try
        objConsolidado.dtProcesos.Rows(3).Item(1) = 1
    End Sub

    Public Overridable Sub descargarDocumentos(ByRef paso As String, errores As List(Of String))
        Dim params As New List(Of String)
        Dim raiz As String
        params.Add(registro)
        General.llenarTabla(docCarga, params, listDocumentosBD)
        For i = 0 To listDocumentosBD.Rows.Count - 1

            raiz = direccion & "\" & listDocumentosBD.Rows(i).Item("Bandera") & ConstantesHC.NOMBRE_PDF_SEPARADOR & i & listDocumentosBD.Rows(i).Item("Extension")
            paso = raiz
            Try
                File.WriteAllBytes(raiz, listDocumentosBD.Rows(i).Item("Archivo"))
                listDocumentos.Rows.Add(listDocumentosBD.Rows(i).Item("Bandera"), raiz)
                FileClose()
            Catch ex As Exception
                errores.Add(paso)
            End Try
        Next
    End Sub

    Public Overridable Sub crearParte4()
        Try

            Dim listReportes As New List(Of DocumentoVistaPrevia)
            objConsolidado.Generando2doPlanoParte4 = True
            Dim inicioExamenes As Integer

            Dim reporte As New ftp_reportes
            Dim params As New List(Of String)
            Dim paso As String
            Dim errores As New List(Of String)
            Dim raiz As String

            crearSabanaAplicacion(listReportes)
            For i = inicioExamenes To listReportes.Count - 1
                paso = listReportes.Item(i).nombrePDF & " - " & listReportes.Item(i).formula
                Try
                    params.Clear()
                    params.Add(listReportes.Item(i).codigoAdicional)
                    params.Add(listReportes.Item(i).codigoAdicional2)
                    raiz = direccion & "\" & listReportes.Item(i).nombrePDF & ConstantesHC.NOMBRE_PDF_SEPARADOR & listReportes.Item(i).codigoAdicional & ".pdf"
                    reporte.crearReportePDF(listReportes.Item(i).nombrePDF, registro, listReportes.Item(i).reporte,
                                           listReportes.Item(i).codigoAdicional, listReportes.Item(i).formula,
                                           listReportes.Item(i).area, direccion,, True, params, raiz)
                    listDocumentos.Rows.Add(listReportes.Item(i).nombrePDF, raiz)
                Catch ex As Exception
                    errores.Add(paso)
                End Try
            Next
            inicioExamenes = listReportes.Count
            For i = inicioExamenes To listReportes.Count - 1
                paso = listReportes.Item(i).nombrePDF & " - " & listReportes.Item(i).formula
                Try
                    params.Clear()
                    params.Add(listReportes.Item(i).codigoAdicional)
                    params.Add(listReportes.Item(i).codigoAdicional2)
                    Dim parametro As String
                    If listReportes.Item(i).nombrePDF = ConstantesHC.NOMBRE_PDF_IMAGENOLOGIA OrElse
                       listReportes.Item(i).nombrePDF = ConstantesHC.NOMBRE_PDF_IMAGENOLOGIA_R OrElse
                       listReportes.Item(i).nombrePDF = ConstantesHC.NOMBRE_PDF_IMAGENOLOGIA_RR Then
                        parametro = listReportes.Item(i).codigoAdicional + listReportes.Item(i).codigoAdicional2
                    Else
                        parametro = listReportes.Item(i).codigoAdicional
                    End If
                    raiz = direccion & "\" & listReportes.Item(i).nombrePDF & ConstantesHC.NOMBRE_PDF_SEPARADOR & parametro & ".pdf"
                    reporte.crearReportePDF(listReportes.Item(i).nombrePDF, registro, listReportes.Item(i).reporte,
                                                parametro, listReportes.Item(i).formula,
                                                listReportes.Item(i).area, direccion, listReportes.Item(i).codigoAdicional2, True, params, raiz)
                    listDocumentos.Rows.Add(listReportes.Item(i).nombrePDF, raiz)
                Catch ex As Exception
                    errores.Add(paso)
                End Try
            Next
            If errores.Count > 0 Then
                paso = ""
                For i = 0 To errores.Count - 1
                    paso = paso & errores.Item(i) & vbCrLf
                Next
                erroresEncontrados += vbCrLf & "Error en Consolidado: parte 4. " & paso
            End If
            objConsolidado.Generando2doPlanoParte4 = False
        Catch ex As Exception
            objConsolidado.Generando2doPlanoParte4 = False
            erroresEncontrados += vbCrLf & "Error en Consolidado: parte 4. " & ex.Message
        End Try
        objConsolidado.dtProcesos.Rows(4).Item(1) = 1
    End Sub
    Public Overridable Sub crearAdmision(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim params As New List(Of String)
            Dim valor As New DataTable
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                General.llenarTabla("SP_ADMISION_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_ADMISION
                    documento.reporte = New rpt_Admisiones
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_ADMISION.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & ""
                    documento.area = ConstantesHC.NOMBRE_PDF_ADMISION
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en ADMISION." & ex.Message
        End Try

    End Sub
    Private Sub crearCertificados(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim documento As New DocumentoVistaPrevia
            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_FACTURA_CER_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    documento = New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_FACTURA_CERTIFICADO
                    documento.reporte = New rptFacturaCertificado
                    documento.codigoAdicional = valor.Rows(I).Item("Codigo_Factura")
                    documento.formula = "{VISTA_FACTURACION_CERTIFICADOS.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("n_registro") & " AND {VISTA_FACTURACION_CERTIFICADOS.Modulo}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next

        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en CERTIFICADOS. " & ex.Message
        End Try
    End Sub
    Private Sub crearConsentimiento(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim params As New List(Of String)
            Dim valor As New DataTable
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                General.llenarTabla("SP_CONSENTIMIENTO_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_CONSETIMIENTO
                    documento.reporte = New rptConsetimiento
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_PACIENTES.N_Registro}=" & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") &
                                            " And {VISTA_EMPRESAS.Id_empresa}=" & valor.Rows(0).Item("ID_EMPRESA")
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en CONSENTIMIENTO. " & ex.Message
        End Try
    End Sub
    Private Sub crearAnexo1(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For J = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(J).Item("N_REGISTRO"))
                General.llenarTabla("SP_ANEXO1_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As DocumentoVistaPrevia
                    For I = 0 To valor.Rows.Count - 1
                        documento = New DocumentoVistaPrevia
                        documento.nombrePDF = ConstantesHC.NOMBRE_PDF_ANEXO_1
                        documento.reporte = New rptAnexo1
                        documento.codigoAdicional = valor.Rows(I).Item("N_Orden")
                        documento.formula = "{VISTA_ANEXO1.N_Orden} = " & valor.Rows(I).Item("N_Orden") & ""
                        documento.area = ConstantesHC.NOMBRE_PDF_ANEXO_1
                        listReportes.Add(documento)
                    Next
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en ANEXO1. " & ex.Message
        End Try
    End Sub
    Private Sub crearAnexo2(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For J = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(J).Item("N_REGISTRO"))
                General.llenarTabla("SP_ANEXO2_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As DocumentoVistaPrevia
                    For I = 0 To valor.Rows.Count - 1
                        documento = New DocumentoVistaPrevia
                        documento.nombrePDF = ConstantesHC.NOMBRE_PDF_ANEXO_2
                        documento.reporte = New rptAnexo2
                        documento.codigoAdicional = valor.Rows(I).Item("N_Solicitud")
                        documento.formula = "{VISTA_ANEXO2.N_Solicitud} = " & valor.Rows(I).Item("N_Solicitud") & ""
                        documento.area = ConstantesHC.NOMBRE_PDF_ANEXO_2
                        listReportes.Add(documento)
                    Next
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en ANEXO2. " & ex.Message
        End Try
    End Sub
    Private Sub crearAnexo3(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For J = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(J).Item("N_REGISTRO"))
                General.llenarTabla("SP_ANEXO3_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As DocumentoVistaPrevia
                    For I = 0 To valor.Rows.Count - 1
                        documento = New DocumentoVistaPrevia
                        documento.nombrePDF = ConstantesHC.NOMBRE_PDF_ANEXO_3
                        documento.reporte = New rptAnexo3
                        documento.codigoAdicional = valor.Rows(I).Item("N_Orden")
                        documento.formula = "{VISTA_ANEXO3.N_Orden} = " & valor.Rows(I).Item("N_Orden") & ""
                        documento.area = ConstantesHC.NOMBRE_PDF_ANEXO_3
                        listReportes.Add(documento)
                    Next
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en ANEXO3. " & ex.Message
        End Try
    End Sub
    Private Sub crearRemision(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_REMISION_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_REMISION
                    documento.reporte = New rptRemision
                    documento.codigoAdicional = registro
                    documento.formula = "{VISTA_REMISION.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & " AND {VISTA_REMISION.Modulo}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en REMISION. " & ex.Message
        End Try
    End Sub
    Private Sub crearEpicrisis(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_EPICRISIS_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_EPICRISIS
                    documento.reporte = New rptEpicrisis
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_EPICRISIS.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & " AND {VISTA_EPICRISIS.Modulo}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en EPICRISIS. " & ex.Message
        End Try
    End Sub
    Private Sub crearIngresoMedico(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For J = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(J).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_INGRESO_REGISTRO_LISTAR", params, valor)
                Dim servicioNeonatal As Boolean
                servicioNeonatal = General.getEstadoVF(Consultas.AREA_SERVICIO_REGISTRO_VERIFICAR & dtFamiliaRegistro.Rows(J).Item("N_REGISTRO"))

                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_INGRESO_MEDICO
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(J).Item("N_REGISTRO")
                    If servicioNeonatal = True Then
                            documento.reporte = New rptInfoIngresoNeonato
                        documento.formula = "{VISTA_H_INGRESO_NEONATAL.N_Registro} = " & dtFamiliaRegistro.Rows(J).Item("N_REGISTRO") & " AND {VISTA_H_INGRESO_NEONATAL.Modulo}=" & codigoReporte & ""
                    Else
                            documento.reporte = New rptInfoIngresoAdulto
                        documento.formula = "{VISTA_H_INGRESO.N_Registro} = " & dtFamiliaRegistro.Rows(J).Item("N_REGISTRO") & " AND {VISTA_H_INGRESO.Modulo}=" & codigoReporte & ""
                    End If
                        documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en INGRESO MEDICO. " & ex.Message
        End Try
    End Sub
    Private Sub crearSalidaVoluntaria(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                General.llenarTabla("SP_SALIDA_VOL_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_RETIRO
                    documento.reporte = New rptRetiroVoluntario
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_PACIENTES.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & " "
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next

        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en SALIDA VOLUNTARIA. " & ex.Message
        End Try
    End Sub
    Private Sub crearOrdenMedica(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_ORDEN_MEDICA_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_ORDEN_MEDICA
                    documento.reporte = New rptOrdenMedica
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.codigoAdicional2 = codigoReporte
                    documento.formula = "{VISTA_ORDEN_MEDICA.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & " AND {VISTA_ORDEN_MEDICA.Modulo}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next

        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en ORDEN MEDICA. " & ex.Message
        End Try
    End Sub
    Private Sub crearEvolucionMedica(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("[SP_EVOLUCION_MEDICA_REGISTRO_LISTAR]", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_EVOLUCION_MEDICA
                    documento.reporte = New rptEvolucionMedica
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_EVOLUCION.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & " AND {VISTA_EVOLUCION.Modulo}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next

        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en EVOLUCION MEDICA. " & ex.Message
        End Try
    End Sub
    Private Sub crearNotaFisio(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_NOTA_FISIO_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_NOTA_FISIO
                    documento.reporte = New rptNotaFisioterapia
                    documento.codigoAdicional = registro
                    documento.formula = "{VISTA_NOTA_FISIOTERAPIA.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & "  AND {VISTA_NOTA_FISIOTERAPIA.Modulo}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next

        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en NOTA FISIO. " & ex.Message
        End Try
    End Sub

    Private Sub crearHojaFisio(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_HOJA_FISIO_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_OXIGENO_FISIO
                    documento.reporte = New rptFisioterapiaOxigeno
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_PACIENTES.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & "  AND {VISTA_FISIOTERAPIA_OXIGENO.Modulo}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en OXIGENO FISIO. " & ex.Message
        End Try
    End Sub

    Private Sub crearHojaNeb(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_HOJA_NEB_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_NEBULIZACION_FISIO
                    documento.reporte = New rptFisioterapiaNebulizacion
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_PACIENTES.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & "  AND {VISTA_FISIOTERAPIA_NEBULIZACION.Modulo}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en OXIGENO FISIO. " & ex.Message
        End Try
    End Sub
    Private Sub crearInformeQX(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_INFORME_QX_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_INFORMEQX
                    documento.reporte = New rptInformeQx
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_INFORME_QX.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & " and  {VISTA_INFORME_QX.MODULO}= " & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en INFORME QX. " & ex.Message
        End Try
    End Sub
    Private Sub crearSolicitudNoPos(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For J = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(J).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_SOLICITUD_NOPOS_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As DocumentoVistaPrevia
                    For I = 0 To valor.Rows.Count - 1
                        documento = New DocumentoVistaPrevia
                        documento.nombrePDF = ConstantesHC.NOMBRE_PDF_SOLICITUD_MED_NOSPOS
                        Dim reporte As Object = Nothing
                        Dim tipo As Integer
                        Dim dt As New DataTable
                        General.llenarTablaYdgv(Consultas.SOLICITUD_MED_FORMATO_EPS & registro, dt)
                        tipo = dt.Rows(0).Item("Codigo_Reporte").ToString
                        Select Case tipo
                            Case Constantes.COD_REPORTE_NOPOS_GENERAL
                                reporte = New rptSolicitud_Medicamento
                            Case Constantes.COD_REPORTE_NOPOS_AMBUQ
                                reporte = New rptAMBUQ
                            Case Constantes.COD_REPORTE_NOPOS_CAJACOPI
                                reporte = New rptCaja_Copi
                            Case Constantes.COD_REPORTE_NOPOS_COMPARTA
                                reporte = New rptComparta
                            Case Constantes.COD_REPORTE_NOPOS_COOMEVA
                                reporte = New rpt_Coomeva
                        End Select
                        documento.reporte = reporte
                        documento.codigoAdicional = valor.Rows(I).Item("Codigo")
                        documento.formula = "{VISTA_SOLICITUD_MED_NOPOS.codigo} = " & valor.Rows(I).Item("Codigo") & " AND {VISTA_SOLICITUD_MED_NOPOS.Modulo}=" & codigoReporte & ""
                        documento.area = documento.nombrePDF
                        listReportes.Add(documento)
                    Next
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en SOLICITUD NO POS. " & ex.Message
        End Try
    End Sub

    Private Sub crearCTC(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim valor As New DataTable
            Dim params As New List(Of String)
            For J = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(J).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_CTC_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As DocumentoVistaPrevia
                    For I = 0 To valor.Rows.Count - 1
                        documento = New DocumentoVistaPrevia
                        documento.nombrePDF = ConstantesHC.NOMBRE_PDF_CTC
                        documento.reporte = New rptCTC
                        documento.codigoAdicional = valor.Rows(I).Item("Codigo_CTC")
                        documento.formula = "{VISTA_CTC.Codigo_CTC}= " & valor.Rows(I).Item("Codigo_CTC") & " and {VISTA_EMPRESAS.Id_empresa}= " & valor.Rows(I).Item("ID_EMPRESA") & " and  {VISTA_CTC.MODULO}= " & codigoReporte & ""
                        documento.area = documento.nombrePDF
                        listReportes.Add(documento)
                    Next
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en CTC. " & ex.Message
        End Try
    End Sub
    Private Sub crearIIH(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim params As New List(Of String)
            Dim valor As New DataTable
            For J = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(J).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_IIH_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As DocumentoVistaPrevia
                    For I = 0 To valor.Rows.Count - 1
                        documento = New DocumentoVistaPrevia
                        documento.nombrePDF = ConstantesHC.NOMBRE_PDF_INFECCION_IH
                        documento.reporte = New rptInfeccionIH
                        documento.codigoAdicional = valor.Rows(I).Item("Codigo_Solicitud")
                        documento.formula = "{VISTA_INFECCION_IH.Codigo_Solicitud} = " & valor.Rows(I).Item("Codigo_Solicitud") & " AND {VISTA_INFECCION_IH.N_registro} = " & dtFamiliaRegistro.Rows(J).Item("N_REGISTRO") & ""
                        documento.area = documento.nombrePDF
                        listReportes.Add(documento)
                    Next
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en IIH. " & ex.Message
        End Try
    End Sub
    Private Sub crearInterconsulta(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_INTERCONSULTA_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_INTERCONSULTA
                    documento.reporte = New rptInterconsulta
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_INTERCONSULTA.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & "  AND {VISTA_INTERCONSULTA.Modulo}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en INTERCONSULTA. " & ex.Message
        End Try
    End Sub
    Private Sub crearHemodialisis(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_HEMODIALISIS_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_HEMODIALISIS
                    documento.reporte = New rptHemodialisis
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_HEMODIALISIS.n_registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & " AND {VISTA_HEMODIALISIS.MODULO}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en HEMODIALISIS. " & ex.Message
        End Try
    End Sub
    Private Sub crearCateterismo(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim valor As New DataTable
            Dim params As New List(Of String)
            For J = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(J).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_CATETERISMO_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As DocumentoVistaPrevia
                    For I = 0 To valor.Rows.Count - 1
                        documento = New DocumentoVistaPrevia
                        documento.nombrePDF = ConstantesHC.NOMBRE_PDF_CATETERISMO
                        documento.reporte = New rptInformeCateterismo
                        documento.codigoAdicional = valor.Rows(I).Item("Codigo_Informe")
                        If valor.Rows(I).Item("Codigo_Procedimiento") = Constantes.CODIGO_EXAMEN_CATETERISMO Then
                            documento.codigoAdicional2 = Constantes.NOMBRE_CATETERISMO
                        Else
                            documento.codigoAdicional2 = Constantes.NOMBRE_ANGIOPLASTIA_STENT
                        End If
                        documento.formula = "{VISTA_INFORME_CATETERISMO.Codigo_Informe}=" & valor.Rows(I).Item("Codigo_Informe") & " AND {VISTA_INFORME_CATETERISMO.MODULO}=" & codigoReporte & ""
                        documento.area = documento.nombrePDF
                        listReportes.Add(documento)
                    Next
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en HEMODIALISIS. " & ex.Message
        End Try
    End Sub
    Public Overridable Sub crearElectrolitos(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim params As New List(Of String)
            Dim valor As New DataTable
            For i = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(i).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_ELECTROLITOS_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then

                    Dim documento As New DocumentoVistaPrevia

                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_EXAMEN_PARA
                    documento.reporte = If(SesionActual.codigoEnlace = 2, New rptExamenParaclinico_ReyFalse, New rptExamenParaclinico)
                    documento.codigoAdicional = valor.Rows(0).Item("Codigo_Orden")
                    documento.codigoAdicional2 = valor.Rows(0).Item("Codigo_Procedimiento")
                    documento.codigoAdicional3 = Constantes.TipoEXAM
                    documento.formula = " {VISTA_PARACLINICOS_RESULTADO.n_registro}=" & dtFamiliaRegistro.Rows(i).Item("N_REGISTRO") &
                        " AND {VISTA_PARACLINICOS_RESULTADO.codigoBandera}<> ''" &
                        " AND {VISTA_PARACLINICOS_RESULTADO.Tipo_Examen}='" & Constantes.TipoEXAM & "'" &
                       " AND {VISTA_PARACLINICOS_RESULTADO.MODULO}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next

        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en ELECTROLITO. " & ex.Message
        End Try
    End Sub

    Public Overridable Sub crearImagenologia(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim params As New List(Of String)
            Dim valor As New DataTable
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_IMAGENOLOGIA_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_IMAGENOLOGIA
                    documento.reporte = New rptResultado
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = " {VISTA_RESULTADOS_PARACLINICOS.n_registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") &
                           " AND {VISTA_RESULTADOS_PARACLINICOS.Modulo}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en IMAGENOLOGIA. " & ex.Message
        End Try
    End Sub

    Private Sub crearTransfusionSanguinea(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim valor As New DataTable
            Dim params As New List(Of String)
            For J = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(J).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_TRANSFUSION_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As DocumentoVistaPrevia
                    For I = 0 To valor.Rows.Count - 1
                        documento = New DocumentoVistaPrevia
                        documento.nombrePDF = ConstantesHC.NOMBRE_PDF_TRANSFUSION
                        documento.reporte = New rptTransfusion_Sanguinea
                        documento.codigoAdicional = valor.Rows(I).Item("Codigo_TS")
                        documento.formula = "{VISTA_TRANSFUSION_SANGUINEA_MED.Codigo_TS} = " & valor.Rows(I).Item("Codigo_TS") & " AND {VISTA_TRANSFUSION_SANGUINEA_MED.Modulo}=" & codigoReporte & ""
                        documento.area = documento.nombrePDF
                        listReportes.Add(documento)
                    Next
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en TRANSFUSION SANGUINEA. " & ex.Message
        End Try
    End Sub
    Private Sub crearEco(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_ECO_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_ECO
                    documento.reporte = New rptEcocardiograma
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_REPORTE_ECO.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") &
                                                                          " AND {VISTA_REPORTE_ECO.Modulo}= " & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en ECO. " & ex.Message
        End Try
    End Sub
    Private Sub crearEcostress(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_ECOSTRESS_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento As New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_ECOSTRES
                    documento.reporte = New rptEcoEstres
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_REPORTE_ECOSTRES.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") &
                                                                          " AND {VISTA_REPORTE_ECOSTRES.Modulo}= " & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en ECOSTRESS. " & ex.Message
        End Try
    End Sub
    Private Sub crearHojaGlucometria(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_GLUCOMETRIA_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento = New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_GLUCOMETRIA_ENFER
                    documento.reporte = New rptGlucometria
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    documento.formula = "{VISTA_GLUCOMETRIA_ENFE.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & "  AND {VISTA_GLUCOMETRIA_ENFE.Modulo}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en HOJA GLUCOMETRIA. " & ex.Message
        End Try
    End Sub
    Private Sub crearSabanaAplicacion(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_SABANA_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    Dim documento = New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_SABANA_APLICACION
                    documento.reporte = New rptSabanaAplicacion
                    documento.codigoAdicional = dtFamiliaRegistro.Rows(I).Item("N_REGISTRO")
                    If codigoReporte = 0 Then
                        documento.codigoAdicional2 = ConstantesHC.SABANA_CON_TODO
                    Else
                        documento.codigoAdicional2 = ConstantesHC.SABANA_SIN_SIGNOS
                    End If
                    documento.formula = "{VISTA_SABANA_DIAS.N_Registro} = " & dtFamiliaRegistro.Rows(I).Item("N_REGISTRO") & " AND {VISTA_SABANA_DIAS.Modulo}=" & codigoReporte & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next


        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en SABANA. " & ex.Message
        End Try
    End Sub
    Public Overridable Sub crearFactura(listReportes As List(Of DocumentoVistaPrevia))
        Try

            Dim valor As New DataTable
            Dim params As New List(Of String)
            For J = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(J).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_FACTURA_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then
                    For i = 0 To valor.Rows.Count - 1
                        Dim documento = New DocumentoVistaPrevia
                        documento.nombrePDF = ConstantesHC.NOMBRE_PDF_FACTURA
                        documento.reporte = New rptFacturaAsistencial
                        documento.codigoAdicional = valor.Rows(0).Item("Codigo_Factura")
                        documento.codigoAdicional2 = FuncionesContables.Convertir_Numero(valor.Rows(0).Item("Total_Factura"))
                        documento.codigoAdicional3 = Format(valor.Rows(0).Item("Fecha_Egreso"), Constantes.FORMATO_FECHA_GEN)
                        documento.formula = "{VISTA_FACTURACION.codigo_factura} = '" & valor.Rows(0).Item("Codigo_Factura") & "' 
                                  AND {VISTA_FACTURACION.TipoFactura}=" & ConstantesAsis.TIPO_FACTURA_ASISTENCIAL & "
                                  AND {VISTA_FACTURACION.Id_Empresa}=" & valor.Rows(0).Item("ID_EMPRESA") & ""
                        documento.area = documento.nombrePDF
                        listReportes.Add(documento)
                        documento = New DocumentoVistaPrevia
                        documento.nombrePDF = ConstantesHC.NOMBRE_PDF_FACTURA_CERTIFICADO
                        documento.reporte = New rptFacturaCertificado
                        documento.codigoAdicional = dtFamiliaRegistro.Rows(J).Item("N_REGISTRO")
                        documento.formula = "{VISTA_FACTURACION_CERTIFICADOS.N_Registro} = " & dtFamiliaRegistro.Rows(J).Item("N_REGISTRO") & "
                                  AND {VISTA_FACTURACION_CERTIFICADOS.Modulo}=" & codigoReporte & ""
                        documento.area = documento.nombrePDF
                        listReportes.Add(documento)
                    Next

                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en FACTURA. " & ex.Message
        End Try
    End Sub
    Private Sub crearFacturaCTC(listReportes As List(Of DocumentoVistaPrevia))
        Try
            Dim documento As New DocumentoVistaPrevia
            Dim valor As New DataTable
            Dim params As New List(Of String)
            For I = 0 To dtFamiliaRegistro.Rows.Count - 1
                params.Clear()
                params.Add(dtFamiliaRegistro.Rows(I).Item("N_REGISTRO"))
                params.Add(codigoReporte)
                General.llenarTabla("SP_FACTURA_CTC_REGISTRO_LISTAR", params, valor)
                If valor.Rows.Count > 0 Then

                    documento = New DocumentoVistaPrevia
                    documento.nombrePDF = ConstantesHC.NOMBRE_PDF_FACTURA_CTC
                    documento.reporte = New rptFacturaAsistencial
                    documento.codigoAdicional = FuncionesContables.Convertir_Numero(valor.Rows(0).Item("Total_Factura"))
                    documento.codigoAdicional2 = Format(valor.Rows(0).Item("Fecha_Egreso"), Constantes.FORMATO_FECHA_GEN)
                    documento.formula = "{VISTA_FACTURACION.codigo_factura} = '" & valor.Rows(0).Item("Codigo_Factura") & "' 
                                  AND {VISTA_FACTURACION.TipoFactura}=" & ConstantesAsis.TIPO_FACTURA_ASISTENCIAL_CTC & "
                                 AND {VISTA_FACTURACION.Id_Empresa}=" & valor.Rows(0).Item("ID_EMPRESA") & ""
                    documento.area = documento.nombrePDF
                    listReportes.Add(documento)
                End If
            Next
        Catch ex As Exception
            erroresEncontrados += vbCrLf & "Error en FACTURA CTC. " & ex.Message
        End Try
    End Sub


    Public Sub piePagina2()
        Dim oldFile As String = rutaUnion

        Dim newFile As String = direccion & ConstantesHC.CARPETA_VISTA_PREVIA & ConstantesHC.CARPETA_VISTA_PREVIA_ARCHIVO & "COPIA_" & registro & ConstantesHC.EXTENSION_ARCHIVO_PDF

        Dim reader As PdfReader = Nothing

        Dim stamper As PdfStamper = Nothing

        Dim cb As PdfContentByte = Nothing

        Dim rect As Rectangle = Nothing

        Dim pageCount As Integer = 0

        Try
            reader = New PdfReader(oldFile)

            rect = reader.GetPageSizeWithRotation(1)

            stamper = New PdfStamper(reader, New FileStream(newFile, FileMode.Create))

            cb = stamper.GetOverContent(1)
            Dim ct = New ColumnText(cb)
            ct.Alignment = Element.ALIGN_LEFT
            ct.SetSimpleColumn(70, 36, PageSize.A4.Width - 36, PageSize.A4.Height - 300)
            Dim nTbl As PdfPTable = New PdfPTable(2)

            Dim rows As Single() = {135.0F, 145.0F}

            nTbl.SetTotalWidth(rows)

            nTbl.AddCell(New Paragraph("Application Ref No:", FontFactory.GetFont("Arial", 15)))
            nTbl.AddCell(New Paragraph("HPOBM00017", FontFactory.GetFont("Arial", 15)))
            nTbl.WriteSelectedRows(0, 20, 300, 835, stamper.GetOverContent(1))

            stamper.Close()
            reader.Close()
            ct.Go()
            rutaUnion = newFile
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub piePagina()
        Dim copyDoc As New Document()
        Dim rutaCopia As String
        Dim reader As PdfReader
        Dim totalPages As Integer
        rutaCopia = direccion & ConstantesHC.CARPETA_VISTA_PREVIA & ConstantesHC.CARPETA_VISTA_PREVIA_ARCHIVO & "COPIA_" & registro & ConstantesHC.EXTENSION_ARCHIVO_PDF
        Dim copyPdf As New PdfCopy(copyDoc, New FileStream(rutaCopia, FileMode.Create))
        copyPdf.SetPageSize(PageSize.A4.Rotate())
        copyDoc.Open()
        reader = New PdfReader(rutaUnion)
        totalPages = reader.NumberOfPages
        Dim copiedPage As PdfImportedPage
        Dim stamper As PdfCopy.PageStamp
        Dim georgia As Font = FontFactory.GetFont("georgia", 10.0F)
        georgia.Color = BaseColor.GRAY
        Dim c1 As Chunk
        Dim pagina As Integer = 0
        Try

            Do While pagina < totalPages
                pagina += 1
                copiedPage = copyPdf.GetImportedPage(reader, (pagina))
                stamper = copyPdf.CreatePageStamp(copiedPage)
                c1 = New Chunk("You can of course force a newline using \", georgia)
                ColumnText.ShowTextAligned(stamper.GetUnderContent(), Element.ALIGN_RIGHT, New Phrase(c1), 820.0F, 15, 0)
                stamper.AlterContents()

                copyPdf.AddPage(copiedPage)
            Loop
            copyPdf.Flush()
            reader.Close()
        Catch ex As Exception
            Throw ex
        End Try
        rutaUnion = rutaCopia
        copyPdf.Close()

    End Sub

End Class
