Public Class NotaEnfermeriaRR
    Inherits NotaEnfermeria

    Sub New()
        notaEnferAnular = ConsultasHC.FISIOTERAPIA_NOTA_ANULAR_RR

    End Sub
    Public Overrides Sub anularNota()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoNota)
        General.ejecutarSQL(notaEnferAnular, params)
    End Sub
End Class
