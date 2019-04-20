Public Class EscalaMorse

    Public Property registro As String
    Public Property l1 As String
    Public Property l2 As String
    Public Property l3 As String
    Public Property l4 As String
    Public Property l5 As String
    Public Property l6 As String
    Public Property dateFecha As Date
    Public Property r2 As Boolean
    Public Property usuario As Integer
    Public Property r4 As Boolean
    Public Property r6 As Boolean
    Public Property r7 As Boolean
    Public Property r5 As Boolean
    Public Property r9 As Boolean
    Public Property r10 As Boolean
    Public Property r11 As Boolean
    Public Property r12 As Boolean
    Public Property r1 As Boolean
    Public Property r3 As Boolean
    Public Property r8 As Boolean
    Public Property r13 As Boolean
    Public Property r14 As Boolean
    Public Property dt As DataTable
    Public Property total As String
    Public Sub cargarDatos()
        dt = New DataTable
        Dim params As New List(Of String)
        params.Add(registro)
        General.llenarTabla(Consultas.MORSE_CARGAR, params, dt)
        If dt.Rows.Count > 0 Then
            l1 = dt.Rows(0).Item("hc")
            l2 = dt.Rows(0).Item("ds")
            l3 = dt.Rows(0).Item("apd")
            l4 = dt.Rows(0).Item("v")
            l5 = dt.Rows(0).Item("m")
            l6 = dt.Rows(0).Item("em")
            dateFecha = dt.Rows(0).Item("Fecha")
            If dt.Rows(0).Item("hc") > 0 Then
                r2 = True
            End If
            If dt.Rows(0).Item("ds") > 0 Then
                r4 = True
            End If
            If dt.Rows(0).Item("apd") = 15 Then
                r6 = True
            ElseIf dt.Rows(0).Item("apd") = 30 Then
                r7 = True
            Else
                r5 = True
            End If
            If dt.Rows(0).Item("v") > 0 Then
                r9 = True
            End If
            If dt.Rows(0).Item("m") = 10 Then
                r11 = True
            ElseIf dt.Rows(0).Item("m") = 20 Then
                R12 = True
            Else
                r10 = True
            End If
            If dt.Rows(0).Item("em") > 0 Then
                R14 = True
            End If
        End If
    End Sub
    Public Sub guardarEscala()
        EscalaMorseDAL.guardarMorse(Me)
    End Sub

End Class

