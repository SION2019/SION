
Public Class Fabrica(Of claveTipo, objetoGeneral)
    Private Delegate Function creadorFuncion() As objetoGeneral
    Private hashtableHC As New Hashtable

    Public Sub registrar(Of objetoEspecifico As {objetoGeneral, New})(clave As claveTipo)
        Dim CreadorFuncion As creadorFuncion = AddressOf creadorObj(Of objetoEspecifico)
        hashtableHC.Add(clave, CreadorFuncion)
    End Sub

    Public Function crear(clave As claveTipo) As objetoGeneral
        Dim CreadorFuncion As creadorFuncion = hashtableHC(clave)
        Return CreadorFuncion()
    End Function

    Private Function creadorObj(Of objetoEspecifico As {objetoGeneral, New})() As objetoGeneral
        Return New objetoEspecifico()
    End Function

End Class

