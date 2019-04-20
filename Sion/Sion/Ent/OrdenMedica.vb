Public Class OrdenMedica
    Public Property codigoOrden As String
    Public Property registro As Integer
    Public Property indicacion As String
    Public Property fechaOrden As DateTime
    Public Property objetoNutricion As OrdenMedicaNutricion
    Public Property usuario As Integer
    Public Property usuarioNombre As String
    Public Property usuarioCreacion As String
    Public Property usuarioReal As Integer
    Public Property codigoEP As Integer
    Public Property nombreReporte As String
    Public Property moduloReporte As Integer
    Public Property ordenListar As String
    Public Property ordenCargar As String
    Public Property estanciaCargar As String
    Public Property indicacionCargar As String
    Public Property comidaCargar As String
    Public Property oxigenoCargar As String
    Public Property oxigenoAutoCargar As String
    Public Property medicamentoAutoCargar As String
    Public Property medicamentoAuditoriaCargar As String
    Public Property procedimientoAutoCargar As String
    Public Property paraclinicoCargar As String
    Public Property hemoderivadoCargar As String
    Public Property glucometriaCargar As String
    Public Property procedimientoCargar As String
    Public Property medicamentoCargar As String
    Public Property impregnacionCargar As String
    Public Property infusionCargar As String
    Public Property mezclaCargar As String
    Public Property dtEstancias As New DataTable
    Public Property dtComidas As New DataTable
    Public Property dtOxigenos As New DataTable
    Public Property dtParaclinicos As New DataTable
    Public Property dtHemoderivados As New DataTable
    Public Property dtGlucometrias As New DataTable
    Public Property dtProcedimientos As New DataTable
    Public Property dtProcedimientosAuto As New DataTable
    Public Property dtInfusiones As New DataTable
    Public Property dtMedicamentos As New DataTable
    Public Property dtMedicamentosAuto As New DataTable
    Public Property dtImpregnaciones As New DataTable
    Public Property dsMezcla As New DataSet
    Public Property elementoAEliminar As New List(Of String)
    Public Property modulo As String
    Public Property opcionCancelar As Boolean

    Sub New()
        objetoNutricion = New OrdenMedicaNutricion
        nombreReporte = ConstantesHC.NOMBRE_PDF_ORDEN_MEDICA
        moduloReporte = Constantes.REPORTE_HC
        ordenListar = Consultas.ORDEN_MEDICA_LISTAR
        ordenCargar = Consultas.ORDEN_MEDICA_CARGAR
        estanciaCargar = Consultas.ORDEN_MEDICA_ESTANCIA_CARGAR
        indicacionCargar = Consultas.ORDEN_MEDICA_INDICACION_CARGAR
        comidaCargar = Consultas.ORDEN_MEDICA_COMIDA_CARGAR
        oxigenoCargar = Consultas.ORDEN_MEDICA_OXIGENO_CARGAR
        oxigenoAutoCargar = Consultas.ORDEN_MEDICA_OXIGENO_AUTO_CARGAR
        medicamentoAutoCargar = Consultas.ORDEN_MEDICA_MEDICAMENTO_AUTO_CARGAR
        medicamentoAuditoriaCargar = Consultas.ORDEN_MEDICA_MEDICAMENTO_AUDITORIA_CARGAR
        procedimientoAutoCargar = Consultas.ORDEN_MEDICA_PROCEDIMIENTO_AUTO_CARGAR
        paraclinicoCargar = Consultas.ORDEN_MEDICA_PARACLINICO_CARGAR
        hemoderivadoCargar = Consultas.ORDEN_MEDICA_HEMODERIVADO_CARGAR
        glucometriaCargar = Consultas.ORDEN_MEDICA_GLUCOMETRIA_CARGAR
        procedimientoCargar = Consultas.ORDEN_MEDICA_PROCEDIMIENTO_CARGAR
        medicamentoCargar = Consultas.ORDEN_MEDICA_MEDICAMENTO_CARGAR
        impregnacionCargar = Consultas.ORDEN_MEDICA_IMPREGNACION_CARGAR
        infusionCargar = Consultas.ORDEN_MEDICA_INFUSION_CARGAR
        mezclaCargar = Consultas.ORDEN_MEDICA_MEZCLA_CARGAR
        asignarColDtEstancia()
        asignarColDtComida()
        asignarColDtOxigeno()
        asignarColDtParaclinico()
        asignarColDtHemoderivado()
        asignarColDtGlucometria()
        asignarColDtProcedimiento()
        asignarColDtMedicamentos()
        asignarColDtImpregnacion()
        asignarColDtInfusion()
    End Sub

    Public Sub asignarColDtEstancia()
        Dim col1, col2, col3, col4, col5 As New DataColumn
        col1.ColumnName = "Código"
        col1.DataType = Type.GetType("System.String")
        dtEstancias.Columns.Add(col1)
        col2.ColumnName = "Descripción"
        col2.DataType = Type.GetType("System.String")
        dtEstancias.Columns.Add(col2)
        col3.ColumnName = "Dia"
        col3.DataType = Type.GetType("System.DateTime")
        dtEstancias.Columns.Add(col3)
        col4.ColumnName = "Justificación"
        col4.DataType = Type.GetType("System.String")
        dtEstancias.Columns.Add(col4)
        col5.ColumnName = "Estado"
        col5.DataType = Type.GetType("System.String")
        dtEstancias.Columns.Add(col5)
    End Sub
    Public Sub asignarColDtComida()
        Dim col1, col2, col3, col4, col5, col6, col7, col8 As New DataColumn
        col1.ColumnName = "Código"
        col1.DataType = Type.GetType("System.String")
        dtComidas.Columns.Add(col1)
        col2.ColumnName = "Descripción"
        col2.DataType = Type.GetType("System.String")
        dtComidas.Columns.Add(col2)
        col3.ColumnName = "Desayuno"
        col3.DataType = Type.GetType("System.Boolean")
        dtComidas.Columns.Add(col3)
        col4.ColumnName = "Almuerzo"
        col4.DataType = Type.GetType("System.Boolean")
        dtComidas.Columns.Add(col4)
        col5.ColumnName = "Cena"
        col5.DataType = Type.GetType("System.Boolean")
        dtComidas.Columns.Add(col5)
        col6.ColumnName = "Justificación"
        col6.DataType = Type.GetType("System.String")
        dtComidas.Columns.Add(col6)
        col7.ColumnName = "Virtual"
        col7.DataType = Type.GetType("System.Boolean")
        dtComidas.Columns.Add(col7)
        col8.ColumnName = "Estado"
        col8.DataType = Type.GetType("System.String")
        dtComidas.Columns.Add(col8)
    End Sub
    Public Sub asignarColDtOxigeno()
        Dim col1, col2, col3, col4, col5, col6 As New DataColumn
        col1.ColumnName = "Código"
        col1.DataType = Type.GetType("System.String")
        dtOxigenos.Columns.Add(col1)
        col2.ColumnName = "Descripción"
        col2.DataType = Type.GetType("System.String")
        dtOxigenos.Columns.Add(col2)
        col3.ColumnName = "Porcentaje"
        col3.DataType = Type.GetType("System.Double")
        dtOxigenos.Columns.Add(col3)
        col4.ColumnName = "Justificación"
        col4.DataType = Type.GetType("System.String")
        dtOxigenos.Columns.Add(col4)
        col5.ColumnName = "Estado"
        col5.DataType = Type.GetType("System.String")
        dtOxigenos.Columns.Add(col5)
        col6.ColumnName = "Suspender"
        col6.DataType = Type.GetType("System.Boolean")
        col6.DefaultValue = False
        dtOxigenos.Columns.Add(col6)
    End Sub

    Public Sub asignarColDtParaclinico()
        Dim col1, col2, col3, col4, col5, col6, col7 As New DataColumn
        col1.ColumnName = "Código"
        col1.DataType = Type.GetType("System.String")
        dtParaclinicos.Columns.Add(col1)
        col2.ColumnName = "Descripción"
        col2.DataType = Type.GetType("System.String")
        dtParaclinicos.Columns.Add(col2)
        col3.ColumnName = "Cantidad"
        col3.DataType = Type.GetType("System.Int32")
        dtParaclinicos.Columns.Add(col3)
        col4.ColumnName = "Indicación"
        col4.DataType = Type.GetType("System.String")
        dtParaclinicos.Columns.Add(col4)
        col5.ColumnName = "Justificación"
        col5.DataType = Type.GetType("System.String")
        dtParaclinicos.Columns.Add(col5)
        col6.ColumnName = "Resultado"
        col6.DataType = Type.GetType("System.String")
        dtParaclinicos.Columns.Add(col6)
        col7.ColumnName = "Estado"
        col7.DataType = Type.GetType("System.String")
        dtParaclinicos.Columns.Add(col7)
    End Sub
    Public Sub asignarColDtHemoderivado()
        Dim col1, col2, col3, col4, col5, col6, col7, col8 As New DataColumn
        col1.ColumnName = "Código"
        col1.DataType = Type.GetType("System.String")
        dtHemoderivados.Columns.Add(col1)
        col2.ColumnName = "Descripción"
        col2.DataType = Type.GetType("System.String")
        dtHemoderivados.Columns.Add(col2)
        col3.ColumnName = "Cantidad"
        col3.DataType = Type.GetType("System.Int32")
        dtHemoderivados.Columns.Add(col3)
        col4.ColumnName = "Indicación"
        col4.DataType = Type.GetType("System.String")
        dtHemoderivados.Columns.Add(col4)
        col5.ColumnName = "Justificación"
        col5.DataType = Type.GetType("System.String")
        dtHemoderivados.Columns.Add(col5)
        col6.ColumnName = "Resultado"
        col6.DataType = Type.GetType("System.String")
        dtHemoderivados.Columns.Add(col6)
        col7.ColumnName = "Formato"
        col7.DataType = Type.GetType("System.String")
        dtHemoderivados.Columns.Add(col7)
        col8.ColumnName = "Estado"
        col8.DataType = Type.GetType("System.String")
        dtHemoderivados.Columns.Add(col8)
    End Sub
    Public Sub asignarColDtGlucometria()
        Dim col1, col2, col3, col4, col5, col6 As New DataColumn
        col1.ColumnName = "Código"
        col1.DataType = Type.GetType("System.String")
        dtGlucometrias.Columns.Add(col1)
        col2.ColumnName = "Descripción"
        col2.DataType = Type.GetType("System.String")
        dtGlucometrias.Columns.Add(col2)
        col3.ColumnName = "Cantidad"
        col3.DataType = Type.GetType("System.Int32")
        dtGlucometrias.Columns.Add(col3)
        col4.ColumnName = "Inicio"
        col4.DataType = Type.GetType("System.String")
        dtGlucometrias.Columns.Add(col4)
        col5.ColumnName = "Frecuencia"
        col5.DataType = Type.GetType("System.Int32")
        dtGlucometrias.Columns.Add(col5)
        col6.ColumnName = "Estado"
        col6.DataType = Type.GetType("System.String")
        dtGlucometrias.Columns.Add(col6)
    End Sub
    Public Sub asignarColDtProcedimiento()
        Dim col1, col2, col3, col4, col5, col6, col7, col8 As New DataColumn
        col1.ColumnName = "Código"
        col1.DataType = Type.GetType("System.String")
        dtProcedimientos.Columns.Add(col1)
        col2.ColumnName = "Descripción"
        col2.DataType = Type.GetType("System.String")
        dtProcedimientos.Columns.Add(col2)
        col3.ColumnName = "Cantidad"
        col3.DataType = Type.GetType("System.Int32")
        dtProcedimientos.Columns.Add(col3)
        col4.ColumnName = "Indicación"
        col4.DataType = Type.GetType("System.String")
        dtProcedimientos.Columns.Add(col4)
        col5.ColumnName = "Justificación"
        col5.DataType = Type.GetType("System.String")
        dtProcedimientos.Columns.Add(col5)
        col6.ColumnName = "Resultado"
        col6.DataType = Type.GetType("System.String")
        dtProcedimientos.Columns.Add(col6)
        col7.ColumnName = "Formato"
        col7.DataType = Type.GetType("System.String")
        dtProcedimientos.Columns.Add(col7)
        col8.ColumnName = "Estado"
        col8.DataType = Type.GetType("System.String")
        dtProcedimientos.Columns.Add(col8)
    End Sub
    Public Sub asignarColDtMedicamentos()
        Dim col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12 As New DataColumn
        col1.ColumnName = "Código"
        col1.DataType = Type.GetType("System.String")
        dtMedicamentos.Columns.Add(col1)
        col2.ColumnName = "Descripción"
        col2.DataType = Type.GetType("System.String")
        dtMedicamentos.Columns.Add(col2)
        col3.ColumnName = "Unidad"
        col3.DataType = Type.GetType("System.String")
        dtMedicamentos.Columns.Add(col3)
        col4.ColumnName = "Dosis"
        col4.DataType = Type.GetType("System.Double")
        dtMedicamentos.Columns.Add(col4)
        col5.ColumnName = "CódigoVia"
        col5.DataType = Type.GetType("System.String")
        dtMedicamentos.Columns.Add(col5)
        col6.ColumnName = "DescripciónVia"
        col6.DataType = Type.GetType("System.String")
        dtMedicamentos.Columns.Add(col6)
        col7.ColumnName = "Horario"
        col7.DataType = Type.GetType("System.String")
        dtMedicamentos.Columns.Add(col7)
        col8.ColumnName = "Inicio"
        col8.DataType = Type.GetType("System.String")
        dtMedicamentos.Columns.Add(col8)
        col9.ColumnName = "Frecuencia"
        col9.DataType = Type.GetType("System.Int32")
        dtMedicamentos.Columns.Add(col9)
        col10.ColumnName = "Dias"
        col10.DataType = Type.GetType("System.String")
        dtMedicamentos.Columns.Add(col10)
        col11.ColumnName = "Suspender"
        col11.DataType = Type.GetType("System.Boolean")
        dtMedicamentos.Columns.Add(col11)
        col12.ColumnName = "Estado"
        col12.DataType = Type.GetType("System.String")
        dtMedicamentos.Columns.Add(col12)
    End Sub
    Public Sub asignarColDtMedicamentosAuto()
        Dim col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12 As New DataColumn
        col1.ColumnName = "Código"
        col1.DataType = Type.GetType("System.String")
        dtMedicamentosAuto.Columns.Add(col1)
        col2.ColumnName = "Descripción"
        col2.DataType = Type.GetType("System.String")
        dtMedicamentosAuto.Columns.Add(col2)
        col3.ColumnName = "Unidad"
        col3.DataType = Type.GetType("System.String")
        dtMedicamentosAuto.Columns.Add(col3)
        col4.ColumnName = "Dosis"
        col4.DataType = Type.GetType("System.Double")
        dtMedicamentosAuto.Columns.Add(col4)
        col5.ColumnName = "CódigoVia"
        col5.DataType = Type.GetType("System.String")
        dtMedicamentosAuto.Columns.Add(col5)
        col6.ColumnName = "DescripciónVia"
        col6.DataType = Type.GetType("System.String")
        dtMedicamentosAuto.Columns.Add(col6)
        col7.ColumnName = "Horario"
        col7.DataType = Type.GetType("System.String")
        dtMedicamentosAuto.Columns.Add(col7)
        col8.ColumnName = "Inicio"
        col8.DataType = Type.GetType("System.String")
        dtMedicamentosAuto.Columns.Add(col8)
        col9.ColumnName = "Frecuencia"
        col9.DataType = Type.GetType("System.Int32")
        dtMedicamentosAuto.Columns.Add(col9)
        col10.ColumnName = "Dias"
        col10.DataType = Type.GetType("System.String")
        dtMedicamentosAuto.Columns.Add(col10)
        col11.ColumnName = "Suspender"
        col11.DataType = Type.GetType("System.Boolean")
        dtMedicamentosAuto.Columns.Add(col11)
        col12.ColumnName = "Estado"
        col12.DataType = Type.GetType("System.String")
        dtMedicamentosAuto.Columns.Add(col12)
    End Sub
    Public Sub asignarColDtImpregnacion()
        Dim col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12, col13, col14, col15 As New DataColumn
        col1.ColumnName = "Código"
        col1.DataType = Type.GetType("System.String")
        dtImpregnaciones.Columns.Add(col1)
        col2.ColumnName = "Descripción"
        col2.DataType = Type.GetType("System.String")
        dtImpregnaciones.Columns.Add(col2)
        col3.ColumnName = "Unidad"
        col3.DataType = Type.GetType("System.String")
        dtImpregnaciones.Columns.Add(col3)
        col4.ColumnName = "Dosis"
        col4.DataType = Type.GetType("System.Double")
        dtImpregnaciones.Columns.Add(col4)
        col5.ColumnName = "Velocidad"
        col5.DataType = Type.GetType("System.Double")
        dtImpregnaciones.Columns.Add(col5)
        col6.ColumnName = "CódigoDisolvente"
        col6.DataType = Type.GetType("System.Int32")
        dtImpregnaciones.Columns.Add(col6)
        col7.ColumnName = "DescripciónDisolvente"
        col7.DataType = Type.GetType("System.String")
        dtImpregnaciones.Columns.Add(col7)
        col8.ColumnName = "Cantidad"
        col8.DataType = Type.GetType("System.Int32")
        dtImpregnaciones.Columns.Add(col8)
        col9.ColumnName = "TotalDisolvente"
        col9.DataType = Type.GetType("System.String")
        dtImpregnaciones.Columns.Add(col9)
        col10.ColumnName = "Inicio"
        col10.DataType = Type.GetType("System.String")
        dtImpregnaciones.Columns.Add(col10)
        col11.ColumnName = "Dias"
        col11.DataType = Type.GetType("System.String")
        dtImpregnaciones.Columns.Add(col11)
        col12.ColumnName = "Suspender"
        col12.DataType = Type.GetType("System.Boolean")
        dtImpregnaciones.Columns.Add(col12)
        col13.ColumnName = "Estado"
        col13.DataType = Type.GetType("System.String")
        dtImpregnaciones.Columns.Add(col13)
        col14.ColumnName = "ConcentracionDisolvente"
        col14.DataType = Type.GetType("System.String")
        dtImpregnaciones.Columns.Add(col14)
        col15.ColumnName = "UnidadDisolvente"
        col15.DataType = Type.GetType("System.String")
        dtImpregnaciones.Columns.Add(col15)
    End Sub
    Public Sub asignarColDtInfusion()
        Dim col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12, col13, col14 As New DataColumn
        col1.ColumnName = "Código"
        col1.DataType = Type.GetType("System.String")
        dtInfusiones.Columns.Add(col1)
        col2.ColumnName = "Descripción"
        col2.DataType = Type.GetType("System.String")
        dtInfusiones.Columns.Add(col2)
        col3.ColumnName = "Unidad"
        col3.DataType = Type.GetType("System.String")
        dtInfusiones.Columns.Add(col3)
        col4.ColumnName = "Dosis"
        col4.DataType = Type.GetType("System.Double")
        dtInfusiones.Columns.Add(col4)
        col5.ColumnName = "Velocidad"
        col5.DataType = Type.GetType("System.Double")
        dtInfusiones.Columns.Add(col5)
        col6.ColumnName = "CódigoDisolvente"
        col6.DataType = Type.GetType("System.String")
        dtInfusiones.Columns.Add(col6)
        col7.ColumnName = "DescripciónDisolvente"
        col7.DataType = Type.GetType("System.String")
        dtInfusiones.Columns.Add(col7)
        col8.ColumnName = "TotalDosis"
        col8.DataType = Type.GetType("System.String")
        dtInfusiones.Columns.Add(col8)
        col9.ColumnName = "Inicio"
        col9.DataType = Type.GetType("System.String")
        dtInfusiones.Columns.Add(col9)
        col10.ColumnName = "Dias"
        col10.DataType = Type.GetType("System.String")
        dtInfusiones.Columns.Add(col10)
        col11.ColumnName = "Suspender"
        col11.DataType = Type.GetType("System.Boolean")
        dtInfusiones.Columns.Add(col11)
        col12.ColumnName = "Estado"
        col12.DataType = Type.GetType("System.String")
        dtInfusiones.Columns.Add(col12)
        col13.ColumnName = "Consecutivo"
        col13.DataType = Type.GetType("System.String")
        dtInfusiones.Columns.Add(col13)

    End Sub

    Public Sub cargarOrdenNutricion()
        objetoNutricion.codigoOrden = codigoOrden
        objetoNutricion.cargarNutricion()
    End Sub
    Public Overridable Sub guardarSabanaAutomatica()
        Dim objSabanaAplicacion As New SabanaAplicacionMed
        objSabanaAplicacion = GeneralHC.fabricaHC.crear(ConstantesHC.IDENTIFICADOR_SABANA_APLICACION & Constantes.CODIGO_MENU_HISTC)
        objSabanaAplicacion.usuario = usuario
        objSabanaAplicacion.nRegistro = registro
        objSabanaAplicacion.guardarReporte2doPlano()
    End Sub
    Public Sub cargarListaOrdenes(ByRef plistOrden As ListBox)
        Dim params As New List(Of String)
        params.Add(registro)
        General.cargarLista(ordenListar, params, "Nombre", "Codigo_Orden", plistOrden)
    End Sub
    Friend Sub cargarOrdenMedica()
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(codigoOrden)
        General.llenarTabla(ordenCargar, params, dt)
        If dt.Rows.Count > 0 Then
            usuarioNombre = dt.Rows(0).Item("Usuario")
            usuarioReal = dt.Rows(0).Item("usuarioInforme")
            usuarioCreacion = dt.Rows(0).Item("usuarioCreacion")
            fechaOrden = dt.Rows(0).Item("Fecha_Orden")
            indicacion = dt.Rows(0).Item("Indicacion")
        End If

    End Sub
    Public Sub imprimirOrden(indicelista As Integer)
        Try
            Dim nombreArchivo, ruta, formula, codigoNombre, campo As String
            If indicelista = 0 Then
                codigoNombre = registro
                campo = "N_Registro"
            Else
                campo = "Codigo_Orden"
                codigoNombre = codigoOrden
            End If
            Dim params As New List(Of String)
            params.Add(registro)
            params.Add(moduloReporte)

            nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = Path.GetTempPath() & nombreArchivo
            formula = "{VISTA_ORDEN_MEDICA." & campo & "} = " & codigoNombre & " AND {VISTA_ORDEN_MEDICA.Modulo}=" & moduloReporte & ""
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(nombreReporte, registro, New rptOrdenMedica,
                                  codigoNombre, formula,
                                  nombreReporte, IO.Path.GetTempPath,,, params)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Overridable Function verificarHoraAplicacion(dt As DataTable) As Boolean
        Return False
    End Function
End Class
