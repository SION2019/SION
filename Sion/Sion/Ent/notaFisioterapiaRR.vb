Public Class NotaFisioterapiaRR
    Inherits NotaFisioterapia

    Sub New()
        notaFisioAnular = ConsultasHC.FISIOTERAPIA_NOTA_ANULAR_RR
    End Sub
    Public Overrides Sub anularNota()
        Dim params As New List(Of String)
        params.Add(usuario)
        params.Add(codigoNota)
        General.ejecutarSQL(notaFisioAnular, params)
    End Sub
End Class
