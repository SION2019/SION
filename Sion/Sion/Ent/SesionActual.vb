Public Class SesionActual

    Public Shared Property idEmpresa As Integer
    Public Shared Property codigoEP As Integer
    Public Shared Property idUsuario As Integer
    Public Shared Property usuario As String
    Public Shared Property codigoPerfil As Integer
    Public Shared Property nombreCompleto As String

    Public Shared Property nombreEmpresa As String
    Public Shared Property nombreSede As String
    Public Shared Property codigoEnlace As Integer

    Public Shared Property dtPermisos As New DataTable

    Public Shared Property htError As New Hashtable
    Public Shared Sub cargarError()
        htError.Add(ConstantesError.CODIGO_ERROR_CONEXION, ConstantesError.ERROR_SIN_CONEXION & ConstantesError.COMUNICAR_SISTEMA)
        htError.Add(ConstantesError.CODIGO_ERROR_CONEXION2, ConstantesError.ERROR_SIN_CONEXION & ConstantesError.COMUNICAR_SISTEMA)
        htError.Add(ConstantesError.CODIGO_ERROR_SIN_CONEXION, ConstantesError.ERROR_SIN_CONEXION & ConstantesError.COMUNICAR_SISTEMA)
        htError.Add(ConstantesError.CODIGO_ERROR_SIN_CONEXION2, ConstantesError.ERROR_SIN_CONEXION & ConstantesError.COMUNICAR_SISTEMA)
        htError.Add(ConstantesError.CODIGO_ERROR_TIEMPO_ESPERA, ConstantesError.ERROR_TIEMPO_ESPERA & ConstantesError.COMUNICAR_SISTEMA)
        htError.Add(ConstantesError.CODIGO_ERROR_SIN_CONEXION3, ConstantesError.ERROR_SIN_CONEXION & ConstantesError.COMUNICAR_SISTEMA)
        htError.Add(ConstantesError.CODIGO_ERROR_SIN_CONEXION4, ConstantesError.ERROR_SIN_CONEXION & ConstantesError.COMUNICAR_SISTEMA)
        htError.Add(ConstantesError.CODIGO_ERROR_SIN_CONEXION5, ConstantesError.ERROR_SIN_CONEXION & ConstantesError.COMUNICAR_SISTEMA)
        'htError.Add(ConstantesError.CODIGO_ERROR_SINTAXIS, ConstantesError.ERROR_SINTAXIS & ConstantesError.COMUNICAR_SISTEMA)
        'htError.Add(ConstantesError.CODIGO_ERROR_PREFIJO_INEXISTENTE, ConstantesError.ERROR_PREFIJO_INEXISTENTE & ConstantesError.COMUNICAR_SISTEMA)
        ' htError.Add(ConstantesError.CODIGO_ERROR_ANIDAMIENTO_CASE, ConstantesError.ERROR_ANIDAMIENTO_CASE & ConstantesError.COMUNICAR_SISTEMA)
        'htError.Add(ConstantesError.CODIGO_ERROR_COLUMNA_INVALIDA, ConstantesError.ERROR_COLUMNA_INVALIDA & ConstantesError.COMUNICAR_SISTEMA)
        'htError.Add(ConstantesError.CODIGO_ERROR_ENTRADA_SALIDA, ConstantesError.ERROR_ENTRADA_SALIDA & ConstantesError.COMUNICAR_SISTEMA)
        'htError.Add(ConstantesError.CODIGO_ERROR_PERMISOS_INSUFICIENTES, ConstantesError.ERROR_PERMISOS_INSUFICIENTES & ConstantesError.COMUNICAR_SISTEMA)
        'htError.Add(ConstantesError.CODIGO_ERROR_INTERBLOQUEO, ConstantesError.ERROR_INTERBLOQUEO & ConstantesError.COMUNICAR_SISTEMA)
        'htError.Add(ConstantesError.CODIGO_ERROR_INICIO_SESION, ConstantesError.ERROR_INICIO_SESION & ConstantesError.COMUNICAR_SISTEMA)
        ' htError.Add(ConstantesError.CODIGO_ERROR_DATOS_TRUNCAR, ConstantesError.ERROR_DATOS_TRUNCAR & ConstantesError.COMUNICAR_SISTEMA)
    End Sub
    Public Shared Function mostrarError(ByVal er As Integer) As String
        Return htError(er)
    End Function
    Public Sub New()
    End Sub

    Public Shared Function tienePermiso(pCodigoPermiso As String) As Boolean
        If dtPermisos.Select("Codigo_Menu='" & pCodigoPermiso & "'", "").Count > 0 Then
            Return True
        End If
        Return False
    End Function

End Class
