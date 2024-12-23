Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient

Public Class clsCombo

    Dim ClasseConexao As New ConexaoSQLServer
    Public Property Codigo As Integer
    Public Property Descricao As String
    Public Function PreencherComboBox(query As String, campoId As String, campoNome As String) As List(Of clsCombo)
        Dim lista = New List(Of clsCombo)
        Try
            Using cn = New SqlConnection(ClasseConexao.strConexao)
                cn.Open()
                Using cmd = New SqlCommand(query, cn)
                    Using RDR As SqlDataReader = cmd.ExecuteReader()
                        While RDR.Read()
                            Dim item As New clsCombo With {
                            .Codigo = RDR.GetInt32(RDR.GetOrdinal(campoId)),
                            .Descricao = RDR.GetString(RDR.GetOrdinal(campoNome))
                        }
                            lista.Add(item)
                        End While
                        RDR.Close()
                    End Using
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show("Não foi possível consultar os dados!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Throw
        End Try
        Return lista
    End Function

End Class
