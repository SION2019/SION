Public Class ListaCheckDAL

    Public Shared Sub cargarMenuPadres(objLista As ListaCheck)
        Try
            Using dbCommand As New SqlCommand(objLista.consultaPendiente & objLista.registro & "", FormPrincipal.cnxion2doPlano)
                Using daCuentaPadre As New SqlDataAdapter(dbCommand)
                    daCuentaPadre.Fill(objLista.dsLista)
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    'Public Shared Sub cargarMenuHijos(objLista As ListaCheck)
    '    Try
    '        Using dbCommand As New SqlCommand(objLista.consultaPendiente & objLista.registro & "," & objLista.accion & "", FormPrincipal.cnxion)
    '            Using daCuentaPadre As New SqlDataAdapter(dbCommand)
    '                daCuentaPadre.Fill(objLista.dsLista)
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub


    '------Cargar listado Admision-------

    Public Shared Sub cargarDocumentos(objLista As ListaCheck)
        Try
            Using dbCommand As New SqlCommand("SP_LISTA_DOCUMENTOS_PENDIENTES " & objLista.registro & "," & objLista.accion & "", FormPrincipal.cnxion)
                Using daCuentaPadre As New SqlDataAdapter(dbCommand)
                    daCuentaPadre.Fill(objLista.dsLista, "padre")
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub cargarDocumentosH(objLista As ListaCheck)
        Try
            Using dbCommand As New SqlCommand("SP_LISTA_DOCUMENTOS_PENDIENTES " & objLista.registro & "," & objLista.accion & "", FormPrincipal.cnxion)
                Using daCuentaPadre As New SqlDataAdapter(dbCommand)
                    daCuentaPadre.Fill(objLista.dsLista, "Hijos")
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
