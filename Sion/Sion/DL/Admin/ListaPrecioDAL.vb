Imports System.Data.SqlClient
Public Class ListaPrecioDAL
    Public Function Guardarproveedor(ByVal codigo As String, ByVal idempresa As Integer, ByVal descripcion As String,
                                       ByVal fechalista As DateTime, ByVal usuario As String) As String

        Dim codigor As String = ""
        Try

            Using consulta As New SqlCommand
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction()
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    If codigo <> "" Then
                        consulta.Parameters.Clear()
                        consulta.CommandType = CommandType.StoredProcedure
                        consulta.CommandText = "SP_PROVEEDOR_LISTA_PRECIO_ACTUALIZAR"
                        consulta.Parameters.Add(New SqlParameter("@IDEMPRESA", SqlDbType.Int))
                        consulta.Parameters("@IDEMPRESA").Value = idempresa
                        consulta.Parameters.Add(New SqlParameter("@DESCRIPCION", SqlDbType.NVarChar))
                        consulta.Parameters("@DESCRIPCION").Value = descripcion
                        consulta.Parameters.Add(New SqlParameter("@FECHALISTA", SqlDbType.NVarChar))
                        consulta.Parameters("@FECHALISTA").Value = fechalista
                        consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                        consulta.Parameters("@USUARIO").Value = usuario
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_LISTA", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_LISTA").Value = codigo
                        consulta.ExecuteNonQuery()
                        codigor = codigo
                    Else
                        consulta.Parameters.Clear()
                        consulta.CommandType = CommandType.StoredProcedure
                        consulta.CommandText = "SP_PROVEEDOR_LISTA_PRECIO_CREAR"
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_LISTA", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_LISTA").Value = codigo
                        consulta.Parameters.Add(New SqlParameter("@IDEMPRESA", SqlDbType.Int))
                        consulta.Parameters("@IDEMPRESA").Value = idempresa
                        consulta.Parameters.Add(New SqlParameter("@DESCRIPCION", SqlDbType.NVarChar))
                        consulta.Parameters("@DESCRIPCION").Value = descripcion
                        consulta.Parameters.Add(New SqlParameter("@FECHALISTA", SqlDbType.NVarChar))
                        consulta.Parameters("@FECHALISTA").Value = fechalista
                        consulta.Parameters.Add(New SqlParameter("@USUARIO", SqlDbType.Int))
                        consulta.Parameters("@USUARIO").Value = usuario
                        codigor = CType(consulta.ExecuteScalar, String)

                    End If

                    Try
                        trnsccion.Commit()
                    Catch ex As Exception
                        codigor = ""
                        trnsccion.Rollback()
                        general.manejoErrores(ex) 
                    End Try
                End Using
            End Using
        Catch ex As Exception
            codigor = ""
            general.manejoErrores(ex) 
        End Try
        Return codigor
    End Function
    Public Function guardar_lista(ByVal tabla As DataTable, ByVal codigo_lista As Integer, ByVal costo As Double, ByVal precio As Double) As Boolean
        Try
            Using consulta = New SqlCommand()
                Using trnsccion = FormPrincipal.cnxion.BeginTransaction
                    consulta.Connection = FormPrincipal.cnxion
                    consulta.Transaction = trnsccion
                    consulta.CommandType = CommandType.StoredProcedure

                    For i = 0 To tabla.Rows.Count - 1
                        consulta.Parameters.Clear()
                        consulta.CommandText = "[SP_PROVEEDOR_LISTA_PRECIO_D_CREAR]"
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_LISTA", SqlDbType.Int))
                        consulta.Parameters("@CODIGO_LISTA").Value = codigo_lista
                        consulta.Parameters.Add(New SqlParameter("@CODIGO_PRODUCTO", SqlDbType.NVarChar))
                        consulta.Parameters("@CODIGO_PRODUCTO").Value = tabla.Rows(i).Item("Codigo")
                        consulta.Parameters.Add(New SqlParameter("@COSTO", SqlDbType.Decimal))
                        consulta.Parameters("@COSTO").Value = tabla.Rows(i).Item("Costo")
                        consulta.Parameters.Add(New SqlParameter("@PRECIO", SqlDbType.Decimal))
                        consulta.Parameters("@PRECIO").Value = tabla.Rows(i).Item("precio")
                        consulta.ExecuteNonQuery()
                    Next

                    Try
                        trnsccion.Commit()
                        Return True
                    Catch ex As Exception
                        trnsccion.Rollback()
                        general.manejoErrores(ex) 
                        Return False

                    End Try
                End Using
            End Using

        Catch ex As Exception
            general.manejoErrores(ex) 
            Return False
        End Try
    End Function
    Function buscar_sin_asignar(ByVal Codigo As Integer) As DataTable
        Using dt As New DataTable
            Using consultor As New SqlDataAdapter("[SP_PROVEEDOR_LISTA_NO_ASIGNADOS] '" & Codigo & "'", FormPrincipal.cnxion)
                consultor.Fill(dt)
            End Using

            Return dt
        End Using
    End Function
    Function buscar_asignados(ByVal Codigo As Integer) As DataTable
        Using dt As New DataTable
            Using consultor As New SqlDataAdapter("[SP_PROVEEDOR_LISTA_ASIGNADOS] '" & Codigo & "'", FormPrincipal.cnxion)
                consultor.Fill(dt)
            End Using
            Return dt
        End Using
    End Function

End Class
