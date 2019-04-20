Imports System.Data.SqlClient
Public Class FuRips
    Public Property codigo As Integer
    Public Property registro As Integer
    Public Property dtParametro As DataTable

    Public Property nombre As String
    Public Property segundoNombre As String
    Public Property apellido As String
    Public Property segundoApellido As String
    Public Property telefono As Integer
    Public Property genero As Integer
    Public Property tipoDocumento As Integer
    Public Property documento As Integer
    Public Property codigoDepartamento As String
    Public Property codigoMunicipio As String
    Public Property departamento As String
    Public Property municipio As String
    Public Property direccion As String
    Public Property fechaNacimiento As String
    Public Property descripcion As String
    Public Property usuario As Integer
    Public Sub guardarFuRips()
        FuripsDAL.guardarRegistro(Me)
    End Sub

    Public Sub cargarDatoPaciente()
        Dim dt As New DataTable
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(Consultas.FURIPS_DATOS_PACIENTE_CARGAR, params, dt)
        If dt.Rows.Count > 0 Then
            nombre = dt.Rows(0).Item("Primer_Nombre")
            segundoNombre = Funciones.castFromDbItem(dt.Rows(0).Item("Segundo_Nombre"))
            apellido = dt.Rows(0).Item("Primer_Apellido")
            segundoApellido = dt.Rows(0).Item("Segundo_Apellido")
            tipoDocumento = dt.Rows(0).Item("Codigo_Identificacion")
            documento = dt.Rows(0).Item("Documento_paciente")
            departamento = dt.Rows(0).Item("Codigo_Departamento_Res")
            codigoDepartamento = dt.Rows(0).Item("Codigo_Departamento_Res")
            municipio = dt.Rows(0).Item("Codigo_Municipio_Res")
            codigoMunicipio = dt.Rows(0).Item("Codigo_Municipio_Res")
            direccion = dt.Rows(0).Item("Direccion")
            genero = dt.Rows(0).Item("Codigo_Genero")
            fechaNacimiento = dt.Rows(0).Item("Fecha_Nacimiento")
        End If
    End Sub

    Sub New()
        dtParametro = New DataTable
        dtParametro.Columns.Add("codigoParametro", Type.GetType("System.String"))
        dtParametro.Columns.Add("Valor", Type.GetType("System.String"))
    End Sub
End Class
