Public Class PerfilParaclinico
    Public Property codigo As String
    Public Property Descripcion As String
    Public Property Usuario As Integer
    Public Property sqlGuardarPerfilParaclinico As String
    Public Property sqlBuscarPerfilParaclinico As String
    Public Property sqlBuscarPerfilParaclinicoG As String
    Public Property sqlCargarPerfilParaclinico As String
    Public Property sqlAnularPerfilParaclinico As String
    Public Sub New()
        sqlGuardarPerfilParaclinico = Consultas.PERFIL_PARACLINICO_CREAR
        sqlBuscarPerfilParaclinico = Consultas.PERFIL_PARACLINICO_BUSCAR
        sqlBuscarPerfilParaclinicoG = Consultas.PERFIL_PARACLINICO_G_BUSCAR
        sqlCargarPerfilParaclinico = Consultas.PERFIL_PARACLINICO_CARGAR
        sqlAnularPerfilParaclinico = Consultas.PERFIL_PARACLINICO_ANULAR
    End Sub
    Public Sub AnularPerfilParaclinico()
        General.anularRegistro(sqlAnularPerfilParaclinico & codigo &
             ConstantesHC.NOMBRE_PDF_SEPARADOR3 & Usuario)
    End Sub
End Class
