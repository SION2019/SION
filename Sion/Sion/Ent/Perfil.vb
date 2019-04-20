Public Class TerceroPerfil

    Public Property idTercero As Integer
    Public Property idEmpresa As Integer
    Public Property codigoPerfil As String
    Public Property nombre As String
    Public Property observacion As String

    Sub New()

    End Sub

    Sub New(tercero As Tercero, empresa As Integer)
        idTercero = tercero.idTercero
        idEmpresa = idEmpresa
    End Sub
    Sub New(drPerfilTercero As DataRow)
        codigoPerfil = Funciones.castFromDbItem(drPerfilTercero.Item("codigo_perfil"))
        nombre = Funciones.castFromDbItem(drPerfilTercero.Item("nombre_perfil"))

    End Sub

    Public Sub guardarPerfil()
        If String.IsNullOrEmpty(codigoPerfil) Then
            perfilDAL.crearPerfil(Me)
        Else
            perfilDAL.actualizarPerfil(Me)
        End If
    End Sub

    Friend Sub anularPerfil()
        perfilDAL.anularPerfil(Me)
    End Sub

    Friend Sub cargarPerfil()
        Dim dt As New DataTable
        Dim param As New List(Of String)
        param.Add(codigoPerfil)
        param.Add(SesionActual.idEmpresa)
        General.llenarTabla(Consultas.BUSQUEDA_PERFIL_CARGAR, param, dt)
        nombre = dt.Rows(0).Item("Nombre_Perfil").ToString
        observacion = dt.Rows(0).Item("Observacion").ToString
    End Sub
End Class
