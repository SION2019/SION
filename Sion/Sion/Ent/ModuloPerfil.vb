Public Class ModuloPerfil

    'Atributos
    Private codigoMenu As String
    Private descripcion As String
    Private formulario As String
    Private menuPadre As String

    'Constructor
    Public Sub New()

        End Sub

    Public Sub New(drModuloPerfil As DataRow)
        codigoMenu = drModuloPerfil.Item("Codigo_Menu")
        descripcion = drModuloPerfil.Item("Descripcion_Menu")
        formulario = drModuloPerfil.Item("Formulario")
        menuPadre = drModuloPerfil.Item("Padre_Menu")
    End Sub
    Public Sub setCodigoMenu(pCodigoMenu As String)
        codigoMenu = pCodigoMenu
    End Sub

    Public Function getCodigoMenu() As String
        Return codigoMenu
    End Function

    Public Sub setDescripcion(pDescripcion As String)
        descripcion = pDescripcion
    End Sub
    Public Function getDescripcion() As String
        Return descripcion
    End Function
    Public Sub setFormulario(pFormulario As String)
        formulario = pFormulario
    End Sub
    Public Function getFormulario() As String
        Return formulario
    End Function


    Public Function getMenuPadre() As Integer
        Return menuPadre
    End Function

    Public Sub setMenuPadre(pMenuPadre As String)
        menuPadre = pMenuPadre
    End Sub



End Class
