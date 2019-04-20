Public Class Oxigeno
    Public Property empresa As Integer
    Public Property codigoOrden As Integer
    Public Property fechaOrden As DateTime
    Public Property editado As Boolean
    Public Property usuario As Integer
    Public Property usuarioReal As Integer
    Public Property nRegistro As Integer
    Public Property codigoEP As Integer
    Public Property totalLitrosOxigeno As Integer
    Public Property totalPrecioOxigeno As Double
    Public Property oxigenoCargarGuardados As String
    Public Property fechaEgreso As DateTime
    Public Property tablaOxigeno As DataTable
    Public Property modulo As Integer
    Public Property OxiegenReporte As String
    Public Property ObjOxigenoModo As OxigenoCPAP
    Public Property nombreReporte As String
    Public Property moduloReporte As Integer
    Public Property consultaFechaEgreso As String
    Sub New()
        consultaFechaEgreso = ConsultasHC.OXIGENO_FECHA_EGRESO
        nombreReporte = ConstantesHC.NOMBRE_PDF_OXIGENO_FISIO
        moduloReporte = Constantes.REPORTE_HC
        ObjOxigenoModo = New OxigenoCPAP
        modulo = Constantes.REPORTE_HC
        oxigenoCargarGuardados = ConsultasHC.OXIGENO_CARGAR_PRIMERA_VEZ
        tablaOxigeno = New DataTable
        tablaOxigeno.Columns.Add("Id_oxigeno", Type.GetType("System.Int32"))
        tablaOxigeno.Columns.Add("Codigo_Orden", Type.GetType("System.Int32"))
        tablaOxigeno.Columns.Add("Fecha_Inicio", Type.GetType("System.String"))
        tablaOxigeno.Columns.Add("Fecha_Fin", Type.GetType("System.String"))
        tablaOxigeno.Columns.Add("Codigo", Type.GetType("System.String"))
        tablaOxigeno.Columns.Add("Descripcion", Type.GetType("System.String"))
        tablaOxigeno.Columns.Add("Porcentaje", Type.GetType("System.Double"))
        tablaOxigeno.Columns.Add("Factor", Type.GetType("System.Double"))
        tablaOxigeno.Columns.Add("Litro", Type.GetType("System.Double"))
        tablaOxigeno.Columns.Add("Hora", Type.GetType("System.Double"))
        tablaOxigeno.Columns.Add("Valor", Type.GetType("System.Double"))
        tablaOxigeno.Columns.Add("Total", Type.GetType("System.Double"), "Valor*Litro")
        tablaOxigeno.Columns.Add("ident_fila", Type.GetType("System.Int32"))
        tablaOxigeno.Columns.Add("Usuario", Type.GetType("System.Int32"))
        tablaOxigeno.Columns.Add("Estado", Type.GetType("System.Int32"))
    End Sub
    Public Sub cargarListaOxigenoOrdenados()
        Dim params As New List(Of String)
        params.Add(nRegistro)
        General.llenarTabla2doPlano(oxigenoCargarGuardados, params, tablaOxigeno)
    End Sub
    Public Sub configurarOxigeno()
        Dim tieneFechaEgreso As Boolean = getFechaEgreso()
        Dim diferenciaOxigeno As Integer
        For indiceFilaOxigeno = 0 To (tablaOxigeno.Rows.Count - 1)
            diferenciaOxigeno = calcularMinutosDiferencia(indiceFilaOxigeno)
            If diferenciaOxigeno > 1440 Then
                agregarOxigeno(indiceFilaOxigeno)

                ' configurarOxigeno()
                ' Exit For
                'Else
                'If tablaOxigeno.Rows.Count = 1 Then
                '    tablaOxigeno.Rows(indiceFilaOxigeno).Item("Fecha_Fin") = Format(Convert.ToDateTime(tablaOxigeno.Rows(indiceFilaOxigeno).Item("Fecha_inicio")).AddDays(1), "dd-MM-yyyy 00:00")
                'Else
                '    If indiceFilaOxigeno < tablaOxigeno.Rows.Count - 1 Then
                '        tablaOxigeno.Rows(indiceFilaOxigeno).Item("Fecha_Fin") = tablaOxigeno.Rows(indiceFilaOxigeno + 1).Item("Fecha_inicio")
                '    Else
                '        tablaOxigeno.Rows(indiceFilaOxigeno).Item("Fecha_Fin") = Format(Convert.ToDateTime(tablaOxigeno.Rows(indiceFilaOxigeno).Item("Fecha_inicio")).AddDays(1), "dd-MM-yyyy 00:00")
                '    End If
                'End If
            ElseIf diferenciaOxigeno = 0 Then
                tablaOxigeno.Rows.RemoveAt(indiceFilaOxigeno)
                configurarOxigeno()
                Exit For
            End If
            totalizarMinutosLitrosPorFila(indiceFilaOxigeno)
        Next
    End Sub
    Public Sub totalizarLitrosOxigeno()
        totalLitrosOxigeno = If(IsDBNull(tablaOxigeno.Compute("SUM(Litro)", "Litro > 0")), 0, tablaOxigeno.Compute("SUM(Litro)", "Litro > 0"))
    End Sub
    Public Sub totaliarPrecioOxigeno()
        totalPrecioOxigeno = If(IsDBNull(tablaOxigeno.Compute("SUM(Total)", "Litro > 0")), 0, tablaOxigeno.Compute("SUM(Total)", "Litro > 0"))
    End Sub
    Public Sub colocarFechaFinal()
        If getFechaEgreso() Then
            tablaOxigeno.Rows(tablaOxigeno.Rows.Count - 1).Item("Fecha_Fin") = Format(fechaEgreso, Constantes.FORMATO_FECHA_OXIGENO_MASCARA)
        Else
            tablaOxigeno.Rows(tablaOxigeno.Rows.Count - 1).Item("Fecha_Fin") = Format(Convert.ToDateTime(tablaOxigeno.Rows(tablaOxigeno.Rows.Count - 1).Item("Fecha_inicio")).AddDays(1), Constantes.FORMATO_FECHA_OXIGENO_MASCARA)
        End If

        totalizarMinutosLitrosPorFila(tablaOxigeno.Rows.Count - 1)

    End Sub
    Public Overridable Function getFechaEgreso() As Boolean
        Dim result As Boolean = False
        Dim params As New List(Of String)
        params.Add(nRegistro)
        Dim fila As DataRow
        fila = General.cargarItem(consultaFechaEgreso, params)

        If Not IsNothing(fila) Then
            If Not IsDBNull(fila(0)) Then
                fechaEgreso = fila(0)
                result = True
            End If
        Else
            fechaEgreso = Nothing
        End If
        Return result
    End Function
    Public Function calcularMinutosDiferencia(ByRef filaActual As Integer) As Double
        Return DateDiff(DateInterval.Minute, tablaOxigeno.Rows(filaActual).Item("Fecha_Inicio"), tablaOxigeno.Rows(filaActual).Item("Fecha_Fin"))
    End Function
    Sub totalizarMinutosLitrosPorFila(ByVal fila As Integer)
        calcularMinutos(fila)
        calcularLitros(fila)
    End Sub
    Public Sub calcularMinutos(ByVal fila As Integer)
        If Not IsNothing(tablaOxigeno.Rows(fila)) Then
            tablaOxigeno.Rows(fila).Item("Hora") = calcularMinutosDiferencia(fila)
        End If
    End Sub
    Public Sub calcularLitros(ByVal fila As Integer)
        If Not IsNothing(tablaOxigeno.Rows(fila)) Then
            tablaOxigeno.Rows(fila).Item("Litro") = Math.Ceiling((tablaOxigeno.Rows(fila).Item("Factor") * (calcularMinutosDiferencia(fila) / 60)) * 1200)        'SE DIVIDE ENTRE 60 PARA SABER CUANTO EN HORA Y 1000 PARA LLEVARLO A LA UNIDAD DE LITROS
        End If
    End Sub
    Public Function agregarOxigeno(ByVal indiceFila As Integer) As Boolean
        Dim respuesta As Boolean = False
        ' If getFechaEgreso() = False Then
        respuesta = True
            colocarNuevaFila(indiceFila)
        '  End If
        Return respuesta
    End Function
    Sub colocarNuevaFila(ByVal indiceFila As Integer)
        Dim fechaActual As DateTime = obtenerFechaActual()
        Dim diasDiferencia As Integer = DateDiff(DateInterval.Day, tablaOxigeno.Rows(indiceFila).Item("Fecha_Fin"), CDate(Format(fechaActual, "dd-MM-yyyy HH:mm:ss")))
        'If diasDiferencia > 1 Then
        tablaOxigeno.Rows(indiceFila).Item("Fecha_Fin") = Format(DateAdd(DateInterval.Day, 1, tablaOxigeno.Rows(indiceFila).Item("Fecha_Inicio")), Constantes.FORMATO_FECHA_GEN & " 00:00")
            Dim filaNueva As DataRow = tablaOxigeno.NewRow
            filaNueva("Codigo_Orden") = tablaOxigeno.Rows(indiceFila).Item("Codigo_Orden")
            filaNueva("Fecha_Inicio") = tablaOxigeno.Rows(indiceFila).Item("Fecha_Fin")
            filaNueva("Fecha_Fin") = tablaOxigeno.Rows(indiceFila).Item("Fecha_Fin")
            filaNueva("Codigo") = tablaOxigeno.Rows(indiceFila).Item("Codigo")
            filaNueva("Descripcion") = tablaOxigeno.Rows(indiceFila).Item("Descripcion")
            filaNueva("Porcentaje") = tablaOxigeno.Rows(indiceFila).Item("Porcentaje")
            filaNueva("Factor") = tablaOxigeno.Rows(indiceFila).Item("Factor")
            filaNueva("Litro") = tablaOxigeno.Rows(indiceFila).Item("Litro")
            filaNueva("Hora") = tablaOxigeno.Rows(indiceFila).Item("Hora")
            filaNueva("Valor") = tablaOxigeno.Rows(indiceFila).Item("Valor")
            filaNueva("ident_fila") = tablaOxigeno.Rows(indiceFila).Item("ident_fila")
            filaNueva("Usuario") = tablaOxigeno.Rows(indiceFila).Item("Usuario")
            filaNueva("Estado") = tablaOxigeno.Rows(indiceFila).Item("Estado")
            tablaOxigeno.Rows.InsertAt(filaNueva, indiceFila + 1)
        ' End If
    End Sub
    Sub colocarCPAP(ByVal indiceActual As Integer)
        Dim codigoOxigenoMecanica As Integer = Constantes.VENTLACION_MECANICA
        Dim codigoOxigenoVenturi As Integer = Constantes.HUMIFICADOR_VENTURY
        Dim horaInicio, horaFinal As DateTime
        Dim numeroFilas As Integer = tablaOxigeno.Rows.Count - 1

        ObjOxigenoModo.cargarMetodo(codigoOxigenoMecanica)

        horaInicio = tablaOxigeno.Rows(indiceActual).Item("Fecha_Inicio")
        tablaOxigeno.Rows(indiceActual).Item("Fecha_Fin") = Format(If(ObjOxigenoModo.codigoOxigeno = Constantes.VENTLACION_MECANICA, horaInicio.AddHours(4), horaInicio.AddHours(8)), Constantes.FORMATO_FECHA_OXIGENO_MASCARA)
        tablaOxigeno.Rows(indiceActual).Item("Codigo") = ObjOxigenoModo.codigoOxigeno
        tablaOxigeno.Rows(indiceActual).Item("Descripcion") = ObjOxigenoModo.descripcion
        tablaOxigeno.Rows(indiceActual).Item("Porcentaje") = ObjOxigenoModo.porcentaje
        tablaOxigeno.Rows(indiceActual).Item("Factor") = ObjOxigenoModo.factor


        For indiceFila = 0 To 2
            If indiceFila Mod 2 = 0 Then
                ObjOxigenoModo.cargarMetodo(codigoOxigenoVenturi)
            Else
                ObjOxigenoModo.cargarMetodo(codigoOxigenoMecanica)
            End If
            Dim filaNueva As DataRow = tablaOxigeno.NewRow
            horaFinal = tablaOxigeno.Rows(indiceActual + indiceFila).Item("Fecha_Fin")
            filaNueva("Codigo_Orden") = tablaOxigeno.Rows(indiceActual + indiceFila).Item("Codigo_Orden")
            filaNueva("Fecha_Inicio") = Format(horaFinal, Constantes.FORMATO_FECHA_OXIGENO_MASCARA)
            horaInicio = Format(If(ObjOxigenoModo.codigoOxigeno = Constantes.VENTLACION_MECANICA, horaFinal.AddHours(4), horaFinal.AddHours(8)), Constantes.FORMATO_FECHA_OXIGENO_MASCARA)
            filaNueva("Fecha_Fin") = Format(horaInicio, Constantes.FORMATO_FECHA_OXIGENO_MASCARA)
            filaNueva("Codigo") = ObjOxigenoModo.codigoOxigeno
            filaNueva("Descripcion") = ObjOxigenoModo.descripcion
            filaNueva("Porcentaje") = ObjOxigenoModo.porcentaje
            filaNueva("Factor") = ObjOxigenoModo.factor
            filaNueva("Litro") = 0
            filaNueva("Hora") = 0
            filaNueva("Valor") = ObjOxigenoModo.Valor
            filaNueva("ident_fila") = tablaOxigeno.Rows(indiceActual + indiceFila).Item("ident_fila")
            filaNueva("Usuario") = tablaOxigeno.Rows(indiceActual + indiceFila).Item("Usuario")
            filaNueva("Estado") = tablaOxigeno.Rows(indiceActual + indiceFila).Item("Estado")
            tablaOxigeno.Rows.InsertAt(filaNueva, indiceActual + indiceFila + 1)
        Next
        If numeroFilas > indiceActual Then
            tablaOxigeno.Rows(indiceActual + 4).Item("Fecha_Inicio") = Format(horaInicio, Constantes.FORMATO_FECHA_OXIGENO_MASCARA)
        Else
            tablaOxigeno.Rows(indiceActual + 3).Item("Fecha_Inicio") = Format(horaFinal, Constantes.FORMATO_FECHA_OXIGENO_MASCARA)
        End If

        configurarOxigeno()
        'For i = 0 To tablaOxigeno.Rows.Count - 1
        '    totalizarMinutosLitrosPorFila(i)
        'Next
    End Sub
    Function obtenerFechaActual() As DateTime
        Return Funciones.Fecha(Constantes.FORMATO_FECHA_HORA_NUMERICA)
    End Function
    Public Overridable Sub imprimirOxigeno()
        Try
            Dim nombreArchivo, ruta, codigoNombre, formula As String
            codigoNombre = nRegistro
            'nombreArchivo = ConstantesHC.NOMBRE_PDF_OXIGENO_FISIO & ConstantesHC.NOMBRE_PDF_SEPARADOR & codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            'ruta = Path.GetTempPath() & nombreArchivo
            'ftp_reportes.llamarArchivo(ruta, nombreArchivo, nRegistro, ConstantesHC.NOMBRE_PDF_OXIGENO_FISIO)
            nombreArchivo = nombreReporte & ConstantesHC.NOMBRE_PDF_SEPARADOR &
                                codigoNombre & ConstantesHC.EXTENSION_ARCHIVO_PDF
            ruta = Path.GetTempPath() & nombreArchivo
            formula = "{VISTA_PACIENTES.N_Registro} = " & codigoNombre & " 
                                            AND {VISTA_FISIOTERAPIA_OXIGENO.Modulo}=" & moduloReporte & ""
            Dim reporte As New ftp_reportes
            reporte.crearReportePDF(nombreReporte, nRegistro, New rptFisioterapiaOxigeno,
                                  codigoNombre, formula,
                                  nombreReporte, IO.Path.GetTempPath)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
End Class
