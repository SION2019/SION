Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Data.SqlClient
Imports System.IO

Public Class ftp_reportes
    Dim Convertir As New ConvertirNumeros
    Dim parametros As Boolean
    Public usuarioActual As String
    Dim consulta As SqlCommand
    Dim resultado As SqlDataReader

    Public Sub reportePDF(ruta As String, nRegistro As String, formula As String, area As String, reporte As ReportClass,
                          Optional ByVal texto As String = "", Optional params As List(Of String) = Nothing)
        Try
            Dim tblas As Tables = reporte.Database.Tables
            General.getConnReporte(tblas)
            reporte.Refresh()
            enviarParametros(area, reporte, texto, params)
            If ruta.Contains("\/") Then
                ruta = ruta.Replace("\/", "\")
            End If
            reporte.RecordSelectionFormula = formula
            reporte.ExportToDisk(ExportFormatType.PortableDocFormat, ruta)
            reporte.Close()
            reporte.Dispose()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Shared Sub enviarParametros(area As String, ByRef reporte As Object, ByRef texto As String, ByRef params As List(Of String))

        Select Case area
            Case ConstantesHC.NOMBRE_PDF_BRADEN
                reporte.SetParameterValue("Observacion", texto)
            Case ConstantesHC.NOMBRE_PDF_COMPROBANTE
                reporte.SetParameterValue("suma", texto)
            Case Constantes.NOMBRE_HISTORIAL_ATENCION
                reporte.SetParameterValue("@FECHA_INICIO", params(0))
                reporte.SetParameterValue("@FECHA_FIN", params(1))
                reporte.SetParameterValue("@COMBOCONTRATO", params(2))
                reporte.SetParameterValue("@COMBOCONTACTO", params(3))
                reporte.SetParameterValue("@CADENA", params(4))
            Case Constantes.NOMBRE_PDF_LISTA_EMPLEADO
                reporte.SetParameterValue("@FECHA_INICIO", params(0))
                reporte.SetParameterValue("@FECHA_FIN", params(1))
                reporte.SetParameterValue("@ESTADO", params(2))
                reporte.SetParameterValue("@CADENA", params(3))
            Case ConstantesHC.NOMBRE_PDF_FACTURA, ConstantesHC.NOMBRE_PDF_FACTURA_R, ConstantesHC.NOMBRE_PDF_FACTURA_RR _
                , ConstantesHC.NOMBRE_PDF_FACTURA_CTC, ConstantesHC.NOMBRE_PDF_FACTURA_CTC_R, ConstantesHC.NOMBRE_PDF_FACTURA_CTC_RR
                reporte.SetParameterValue("valorLetras", params(0))
                reporte.SetParameterValue("fechaEgreso", params(1))
            Case ConstantesHC.NOMBRE_PDF_FACTURA_EXTERNA
                reporte.SetParameterValue("valorLetras", params(0))
                reporte.SetParameterValue("observacion", params(1))
            Case ConstantesHC.NOMBRE_PDF_FACTURA_ATENCION
                reporte.SetParameterValue("valorLetras", params(0))
                reporte.SetParameterValue("observacion", params(1))
            Case ConstantesHC.NOMBRE_PDF_FACTURA_AMBULANCIA
                reporte.SetParameterValue("valorLetras", params(0))
            Case ConstantesHC.NOMBRE_PDF_FACTURA_REMISION
                reporte.SetParameterValue("valorLetras", params(0))
                reporte.SetParameterValue("valorIVA", params(1))
                reporte.SetParameterValue("observacion", params(2))
                reporte.SetParameterValue("remisiones", params(3))
            Case ConstantesHC.NOMBRE_PDF_SABANA_APLICACION,
                 ConstantesHC.NOMBRE_PDF_SABANA_APLICACION_R,
                ConstantesHC.NOMBRE_PDF_SABANA_APLICACION_RR
                reporte.SetParameterValue("TIPO", params(1))
                reporte.SetParameterValue("@NREGISTRO", params(0))
            Case ConstantesHC.NOMBRE_PDF_ELECTROLITO,
                ConstantesHC.NOMBRE_PDF_HEMOGRAMA,
                ConstantesHC.NOMBRE_PDF_QUIMICASANGUINEA,
                ConstantesHC.NOMBRE_PDF_UROCULTIVO,
                ConstantesHC.NOMBRE_PDF_UROANALISIS,
                ConstantesHC.NOMBRE_PDF_HEMOCULTIVO,
                ConstantesHC.NOMBRE_PDF_COPROLOGICO,
                ConstantesHC.NOMBRE_PDF_SEROLOGIA,
                ConstantesHC.NOMBRE_PDF_TINTACHINA,
                ConstantesHC.NOMBRE_PDF_GASESARTERIALES,
                ConstantesHC.NOMBRE_PDF_ELECTROLITO_R,
                ConstantesHC.NOMBRE_PDF_HEMOGRAMA_R,
                ConstantesHC.NOMBRE_PDF_QUIMICASANGUINEA_R,
                ConstantesHC.NOMBRE_PDF_UROCULTIVO_R,
                ConstantesHC.NOMBRE_PDF_UROANALISIS_R,
                ConstantesHC.NOMBRE_PDF_HEMOCULTIVO_R,
                ConstantesHC.NOMBRE_PDF_COPROLOGICO_R,
                ConstantesHC.NOMBRE_PDF_SEROLOGIA_R,
                ConstantesHC.NOMBRE_PDF_TINTACHINA_R,
                ConstantesHC.NOMBRE_PDF_GASESARTERIALES_R,
                ConstantesHC.NOMBRE_PDF_ELECTROLITO_RR,
                ConstantesHC.NOMBRE_PDF_HEMOGRAMA_RR,
                ConstantesHC.NOMBRE_PDF_QUIMICASANGUINEA_RR,
                ConstantesHC.NOMBRE_PDF_UROCULTIVO_RR,
                ConstantesHC.NOMBRE_PDF_UROANALISIS_RR,
                ConstantesHC.NOMBRE_PDF_HEMOCULTIVO_RR,
                ConstantesHC.NOMBRE_PDF_COPROLOGICO_RR,
                ConstantesHC.NOMBRE_PDF_SEROLOGIA_RR,
                ConstantesHC.NOMBRE_PDF_TINTACHINA_RR,
                ConstantesHC.NOMBRE_PDF_GASESARTERIALES_RR,
                ConstantesHC.NOMBRE_PDF_TIEMPOP,
                ConstantesHC.NOMBRE_PDF_TIEMPOP_R,
                ConstantesHC.NOMBRE_PDF_TIEMPOP_RR,
                ConstantesHC.NOMBRE_PDF_TIEMPOPT,
                ConstantesHC.NOMBRE_PDF_TIEMPOPT_R,
                ConstantesHC.NOMBRE_PDF_TIEMPOPT_RR,
                ConstantesHC.NOMBRE_PDF_EXAMEN_PARA,
                ConstantesHC.NOMBRE_PDF_EXAMEN_PARA_R,
                ConstantesHC.NOMBRE_PDF_EXAMEN_PARA_RR
                reporte.SetParameterValue("@Codigo_Orden", params(0))
                reporte.SetParameterValue("@Codigo_Procedimiento", params(1))
                reporte.SetParameterValue("@tipo_examen", params(2))
            Case ConstantesHC.NOMBRE_PDF_LAB
                reporte.SetParameterValue("@Codigo_Orden", params(0))
                reporte.SetParameterValue("@Codigo_Procedimiento", params(1))
                reporte.SetParameterValue("@tipo_examen", params(2))
            Case ConstantesHC.NOMBRE_PDF_CATETERISMO,
                ConstantesHC.NOMBRE_PDF_CATETERISMO_R,
                ConstantesHC.NOMBRE_PDF_CATETERISMO_RR
                reporte.SetParameterValue("@titulo", params(1))
            Case ConstantesHC.NOMBRE_PDF_ORDEN_MEDICA,
                 ConstantesHC.NOMBRE_PDF_ORDEN_MEDICAR,
                 ConstantesHC.NOMBRE_PDF_ORDEN_MEDICARR
                reporte.SetParameterValue("@pCodigoOrden", params(0))
                reporte.SetParameterValue("@pModulo", params(1))
            Case ConstantesHC.CONSOLIDADO_EXAMEN
                reporte.SetParameterValue("@Cadena", params(0))
                reporte.SetParameterValue("@Fecha_Inicio", params(1))
                reporte.SetParameterValue("@Fecha_Fin", params(2))
                reporte.SetParameterValue("@Codigo_Laboratorio", params(3))
        End Select
    End Sub

    Public Sub reporteExcel(ruta As String, nRegistro As String, formula As String, area As String, reporte As Object,
                       Optional ByVal texto As String = "", Optional params As List(Of String) = Nothing)
        Try
            Dim tblas As Tables = reporte.Database.Tables
            General.getConnReporte(tblas)
            enviarParametros(area, reporte, texto, params)
            reporte.refresh()
            reporte.RecordSelectionFormula = formula
            reporte.ExportToDisk(ExportFormatType.Excel, ruta)
            reporte.close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub crearReportePDF(carpeta As String, nRegistro As String, reporte As Object, codigo As String, formula As String, area As String,
                                    ruta As String, Optional ByVal texto As String = "", Optional ByVal noAbrirArchivo As Boolean = False,
                                    Optional params As List(Of String) = Nothing, Optional rutaDiferente As String = "")
        Dim cnd As String
        reporte.Close()
        If String.IsNullOrWhiteSpace(rutaDiferente) Then
            cnd = area & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigo & Now.Hour & Now.Minute & Now.Second & Now.Millisecond & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = ruta & "/" & cnd
        Else
            ruta = rutaDiferente
        End If

        Me.reportePDF(ruta, codigo, formula, area, reporte, texto, params)
        If noAbrirArchivo Then Exit Sub
        abrirReporte(ruta)
    End Sub
    Private Sub abrirReporte(rutaPDF As String)
        Process.Start(rutaPDF)
    End Sub
    Public Sub crearReporteExcel(carpeta As String, nRegistro As String, reporte As Object, codigo As String, formula As String, area As String,
                                    ruta As String, Optional ByVal texto As String = "", Optional ByVal noSubirArchivo As Boolean = False,
                                    Optional params As List(Of String) = Nothing)
        Dim cnd As String
        reporte.Close()
        cnd = area & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigo & ConstantesHC.EXTENSION_ARCHIVO_EXCEL

        ruta = ruta & "/" & cnd
        Me.reporteExcel(ruta, codigo, formula, area, reporte, texto, params)
    End Sub

    'Public Shared Sub eliminarArchivo(nombreArchivo As String)
    '    Try

    '        Using dbCommand = New SqlCommand()
    '            dbCommand.Connection = FormPrincipal.cnxion2doPlano
    '            dbCommand.CommandText = "EXEC [SP_ARCHIVO_PDF_BORRAR] '" & nombreArchivo & "'"
    '            dbCommand.ExecuteNonQuery()
    '        End Using
    '        FormPrincipal.cnxion2doPlano.Close()
    '    Catch ex As Exception
    '        general.manejoErrores(ex)
    '    End Try
    'End Sub
End Class
