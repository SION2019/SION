Public Class Nebulizacion
    Public Property empresa As Integer
    Public Property codigoOrden As Integer
    Public Property fechaOrden As DateTime
    Public Property horaInicioNebulizacion As Integer
    Public Property frecuenciaNebulizacion As Integer
    Public Property usuario As Integer
    Public Property usuarioReal As Integer
    Public Property nRegistro As Integer
    Public Property codigoEP As Integer
    Public Property totalLitrosNebulizacion As Integer
    Public Property totalPrecioNebulizacion As Double
    Public Property oxigenoCargarGuardados As String
    Public Property fechaEgreso As DateTime
    Public Property tablaNebulacion As DataTable
    Public Property modulo As Integer
    Public Property nebulizacionReporte As String
    Public Property consultaCargar As String
    Public Property consultaFechaEgreso As String
    Sub New()
        consultaCargar = ConsultasHC.NEBULIZACION_CARGAR
        consultaFechaEgreso = ConsultasHC.OXIGENO_FECHA_EGRESO
        modulo = Constantes.REPORTE_HC
        oxigenoCargarGuardados = ConsultasHC.OXIGENO_CARGAR_PRIMERA_VEZ
        tablaNebulacion = New DataTable
        tablaNebulacion.Columns.Add("Id_Neb", Type.GetType("System.Int32"))
        tablaNebulacion.Columns.Add("Codigo_Orden", Type.GetType("System.Int32"))
        tablaNebulacion.Columns.Add("Codigo_interno", Type.GetType("System.Int32"))
        tablaNebulacion.Columns.Add("nombreMedicamento", Type.GetType("System.String"))
        tablaNebulacion.Columns.Add("Fecha_inicio_neb", Type.GetType("System.String"))
        tablaNebulacion.Columns.Add("Fecha_fin_neb", Type.GetType("System.String"))
        tablaNebulacion.Columns.Add("Codigo", Type.GetType("System.String"))
        tablaNebulacion.Columns.Add("Descripcion", Type.GetType("System.String"))
        tablaNebulacion.Columns.Add("Porcentaje", Type.GetType("System.Double"))
        tablaNebulacion.Columns.Add("Factor", Type.GetType("System.Double"))
        tablaNebulacion.Columns.Add("Litro", Type.GetType("System.Double"))
        tablaNebulacion.Columns.Add("Hora", Type.GetType("System.Double"))
        tablaNebulacion.Columns.Add("Valor", Type.GetType("System.Double"))
        tablaNebulacion.Columns.Add("Total", Type.GetType("System.Double"), "Valor*Litro")
        tablaNebulacion.Columns.Add("ident_fila", Type.GetType("System.Int32"))
        tablaNebulacion.Columns.Add("Usuario", Type.GetType("System.Int32"))
        tablaNebulacion.Columns.Add("Estado", Type.GetType("System.Int32"))
    End Sub
    Public Sub colocarNebulizaciones()
        For indiceAplicacionNebulizacion = horaInicioNebulizacion To ConstantesHC.LIMITE_NEBULIZACION
            If (indiceAplicacionNebulizacion - horaInicioNebulizacion) Mod frecuenciaNebulizacion = 0 Then

            End If
        Next
    End Sub
    Public Sub calculosNebulizacion()
        For indiceFila = 0 To tablaNebulacion.Rows.Count - 1
            If tablaNebulacion.Rows(indiceFila).Item("Descripcion").ToString <> "" Then
                calcularMinutos(indiceFila)
            End If
        Next
    End Sub
    Public Sub calcularMinutos(ByVal fila As Integer)
        tablaNebulacion.Rows(fila).Item("Hora") = DateDiff(DateInterval.Minute, CDate(tablaNebulacion.Rows(fila).Item("Fecha_inicio_neb")), CDate(tablaNebulacion.Rows(fila).Item("Fecha_fin_neb")))
        calcularLitros(fila)
    End Sub
    Public Sub calcularLitros(ByVal fila As Integer)
        tablaNebulacion.Rows(fila).Item("Litro") = (tablaNebulacion.Rows(fila).Item("Factor") * (DateDiff(DateInterval.Minute, tablaNebulacion.Rows(fila).Item("Fecha_inicio_neb"), tablaNebulacion.Rows(fila).Item("Fecha_fin_neb")) / 60)) * 1000 'SE DIVIDE ENTRE 60 PARA SABER CUANTO EN HORA Y 100O PARA LLEVARLO A LA UNIDAD DE LITROS
    End Sub
    Public Overridable Sub getFechaEgreso()
        Dim params As New List(Of String)
        params.Add(nRegistro)
        Dim fila As DataRow
        fila = General.cargarItem(consultaFechaEgreso, params)
        If Not IsDBNull(fila(0)) Then
            fechaEgreso = fila(0)
        Else
            fechaEgreso = Funciones.Fecha(Constantes.FORMATO_FECHA_HORA_NUMERICA)
        End If
    End Sub
    Public Sub cargarListaNebulizacion()
        Dim params As New List(Of String)
        params.Add(nRegistro)
        General.llenarTabla(oxigenoCargarGuardados, params, tablaNebulacion)
    End Sub
    Public Overridable Sub cargarNebulizacion()
        Dim params As New List(Of String)
        params.Add(codigoOrden)
        General.llenarTabla(consultaCargar, params, tablaNebulacion)
    End Sub
    Public Sub totaliarPrecioNebulizacion()
        totalPrecioNebulizacion = If(IsDBNull(tablaNebulacion.Compute("SUM(Total)", "Litro > 0")), 0, tablaNebulacion.Compute("SUM(Total)", "Litro > 0"))
    End Sub
    Public Sub totalizarLitrosOxigeno()
        totalLitrosNebulizacion = If(IsDBNull(tablaNebulacion.Compute("SUM(Litro)", "Litro > 0")), 0, tablaNebulacion.Compute("SUM(Litro)", "Litro > 0"))
    End Sub
    Public Function agregarNuevaNebulizacion(ByVal indiceFila As Integer) As Boolean
        getFechaEgreso()
        If compararFechaEgreso(tablaNebulacion.Rows(indiceFila).Item("Fecha_fin_neb"), Format(fechaEgreso, Constantes.FORMATO_FECHA_HORA_GEN)) = False Then
            Dim filaNueva As DataRow = tablaNebulacion.NewRow
            filaNueva("Id_Neb") = DBNull.Value
            filaNueva("Codigo_Orden") = tablaNebulacion.Rows(indiceFila).Item("Codigo_Orden")
            filaNueva("Codigo_interno") = tablaNebulacion.Rows(indiceFila).Item("Codigo_interno")
            filaNueva("nombreMedicamento") = tablaNebulacion.Rows(indiceFila).Item("nombreMedicamento")
            filaNueva("Fecha_inicio_neb") = tablaNebulacion.Rows(indiceFila).Item("Fecha_fin_neb")
            filaNueva("Fecha_fin_neb") = tablaNebulacion.Rows(indiceFila).Item("Fecha_fin_neb")
            filaNueva("Codigo") = tablaNebulacion.Rows(indiceFila).Item("Codigo")
            filaNueva("Descripcion") = tablaNebulacion.Rows(indiceFila).Item("Descripcion")
            filaNueva("Porcentaje") = tablaNebulacion.Rows(indiceFila).Item("Porcentaje")
            filaNueva("Factor") = tablaNebulacion.Rows(indiceFila).Item("Factor")
            filaNueva("Litro") = tablaNebulacion.Rows(indiceFila).Item("Litro")
            filaNueva("Hora") = tablaNebulacion.Rows(indiceFila).Item("Hora")
            filaNueva("Valor") = tablaNebulacion.Rows(indiceFila).Item("Valor")
            filaNueva("Total") = tablaNebulacion.Rows(indiceFila).Item("Total")
            filaNueva("Usuario") = tablaNebulacion.Rows(indiceFila).Item("Usuario")
            filaNueva("Estado") = tablaNebulacion.Rows(indiceFila).Item("Estado")
            tablaNebulacion.Rows.InsertAt(filaNueva, indiceFila + 1)
        Else
            Return False
        End If
        Return True
    End Function
    Public Function compararFechaEgreso(ByVal fechaComprada As DateTime, ByVal fechaEgreso As DateTime) As Boolean
        Return HistoriaClinicaBLL.isFechaMayor(fechaComprada, fechaEgreso)
    End Function
    Public Overridable Sub imprimirNebulizacion(ByVal lista As ListBox)
        Dim nombreArchivo, ruta, formula, codigoNombre, campo As String

        If lista.SelectedIndex = 0 Then
            'nombreArchivo = ConstantesHC.NOMBRE_PDF_NEBULIZACION_FISIO & ConstantesHC.NOMBRE_PDF_SEPARADOR & nRegistro & ConstantesHC.EXTENSION_ARCHIVO_PDF
            'ruta = Path.GetTempPath() & nombreArchivo
            'ftp_reportes.llamarArchivo(ruta, nombreArchivo, nRegistro, ConstantesHC.NOMBRE_PDF_NEBULIZACION_FISIO)
            campo = "N_Registro"
            codigoNombre = nRegistro
        Else
            codigoNombre = codigoOrden
            campo = "Codigo_Orden"
        End If
        nombreArchivo = ConstantesHC.NOMBRE_PDF_NEBULIZACION_FISIO & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
        ruta = Path.GetTempPath() & nombreArchivo
        formula = "{VISTA_FISIOTERAPIA_NEBULIZACION." & campo & "} = " & codigoOrden & " AND {VISTA_FISIOTERAPIA_NEBULIZACION.Modulo}=" & modulo
        Dim reporte As New ftp_reportes
        reporte.crearReportePDF(ConstantesHC.NOMBRE_PDF_NEBULIZACION_FISIO,
                                         nRegistro, New rptFisioterapiaNebulizacion,
                                         codigoNombre,
                                         formula,
                                         ConstantesHC.NOMBRE_PDF_NEBULIZACION_FISIO,
                                         IO.Path.GetTempPath)
    End Sub
End Class

