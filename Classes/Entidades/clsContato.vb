Imports System.Data.SqlClient
Imports iText.Layout.Element
Public Class clsContato

    Public Property CodContato As Integer
    Public Property CodTipoContato As Integer
    Public Property TipoContato As String
    Public Property NomeContato As String
    Public Property Email As String
    Public Property Telefone As String
#Region "CONSTRUTORES"
    Public Sub New()

    End Sub
    Public Sub New(_codConta As Integer, _nomecontato As String, _email As String, _telefone As String)
        CodContato = _codConta
        NomeContato = _nomecontato
        Email = _email
        Telefone = _telefone
    End Sub
#End Region
#Region "CADASTRO TIPO CONTATO"
    Public Sub CarregaTipoContato(ByRef DadosContato As clsContato)

        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "SELECT IDENT_CURRENT('Tbl_TipoContato')+1 AS Codigo"
                Using CMD = New SqlCommand(sql, cn)
                    Using RDR As SqlDataReader = CMD.ExecuteReader()
                        If RDR.HasRows() Then
                            While RDR.Read()
                                DadosContato.CodTipoContato = RDR.Item("Codigo")
                            End While
                        Else
                            Exit Sub
                        End If
                        RDR.Close()
                    End Using
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Falha ao realizar a consulta no banco de dados" & vbCrLf & ex.Message)
            Throw
        End Try
    End Sub
    Public Function ComboTipoContato() As List(Of clsContato)
        Dim ListaDepartamento = New List(Of clsContato)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "SELECT * FROM Tbl_TipoContato ORDER BY TipoContato"
                Using CMD = New SqlCommand(sql, cn)
                    Using RDR As SqlDataReader = CMD.ExecuteReader()
                        While RDR.Read()
                            Dim Cust As New clsContato
                            Cust.CodContato = RDR.GetInt32(RDR.GetOrdinal("Codigo"))
                            Cust.TipoContato = If(Not RDR.IsDBNull(RDR.GetOrdinal("TipoContato")), RDR.GetString(RDR.GetOrdinal("TipoContato")), String.Empty)
                            ListaDepartamento.Add(Cust)
                        End While
                        RDR.Close()
                    End Using
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível consultar o departamento!" & vbCrLf & ex.Message)
            Throw
        End Try
        Return ListaDepartamento
    End Function
    Public Sub ConsultaTipoContato(Grid As DataGridView)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "SELECT * FROM Tbl_TipoContato"
                Using CMD = New SqlCommand(sql, cn)
                    Dim RDR As SqlClient.SqlDataReader
                    RDR = CMD.ExecuteReader()
                    While RDR.Read()
                        Grid.Rows.Add(RDR.Item("Codigo").ToString,
                                       If(Not IsDBNull(RDR("TipoContato")), RDR("TipoContato"), 0))
                    End While
                    RDR.Close()
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível realizar a consulta!" & vbCrLf & ex.Message)
            Throw
        End Try
    End Sub
    Public Sub SalvaTipoContato(TipoContato As String)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "INSERT INTO Tbl_TipoContato (TipoContato) VALUES (@TipoContato)"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@TipoContato", TipoContato)
                    CMD.ExecuteNonQuery()
                    MsgBox("Cadastro efetuado com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não é possível inserir os dados!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
        End Try
    End Sub
    Public Sub AtualizaTipoContato(Codigo As Integer, TipoContato As String)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "UPDATE Tbl_TipoContato SET TipoContato = @TipoContato WHERE Codigo = @Codigo"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@TipoContato", TipoContato)
                    CMD.Parameters.AddWithValue("@Codigo", Codigo)
                    CMD.ExecuteNonQuery()
                    MsgBox("Cadastro atualizado com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não é possível alterar os dados!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
        End Try
    End Sub
    Public Sub ExcluiTipoContato(Codigo As Integer)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "DELETE FROM Tbl_TipoContato WHERE Codigo = @Codigo"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@Codigo", Codigo)
                    CMD.ExecuteNonQuery()
                    MsgBox("Cadastro excluído com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não é possível alterar os dados!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
        End Try
    End Sub
#End Region
#Region "CLIENTE"
    Public Sub ConsultaContatoCliente(Grid As DataGridView, CodCliente As Integer)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "Select * FROM Cs_ContatoClientes WHERE CodCli = @CodCliente"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@CodCliente", CodCliente)
                    Using RDR As SqlDataReader = CMD.ExecuteReader()
                        If RDR.HasRows() Then
                            While RDR.Read()
                                Grid.Rows.Add(RDR.Item("Codigo").ToString,
                                            If(Not IsDBNull(RDR("Contato")), RDR("Contato"), ""),
                                            If(Not IsDBNull(RDR("TipoContato")), RDR("TipoContato"), ""),
                                            If(Not IsDBNull(RDR("Email")), RDR("Email"), ""),
                                            If(Not IsDBNull(RDR("Telefone")), RDR("Telefone"), ""))
                            End While
                        Else
                            Exit Sub
                        End If
                        RDR.Close()
                    End Using
                End Using
                cn.Close()
            End Using
        Catch ex As SqlException
            MsgBox("Não foi possível realizar a consulta!" & ex.Message, MsgBoxStyle.Critical, "Erro de Banco de Dados")
            Throw
        Catch ex As Exception
            MsgBox("Não foi possível realizar a consulta!" & vbCrLf & ex.Message)
            Throw
        End Try
    End Sub
    Public Sub SalvarContatoCliente(CodCliente As Integer, TipoContato As Integer, Contato As String, Email As String, Telefone As String)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "INSERT INTO Tbl_Contatos (CodCli,Contato,Email,Telefone,Depto) VALUES (@CodCliente, @Contato,@Email,@Telefone,@TipoContato)"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@CodCliente", CodCliente)
                    CMD.Parameters.AddWithValue("@Contato", Contato)
                    CMD.Parameters.AddWithValue("@Email", Email)
                    CMD.Parameters.AddWithValue("@Telefone", Telefone)
                    CMD.Parameters.AddWithValue("@TipoContato", TipoContato)
                    CMD.ExecuteNonQuery()
                    MsgBox("Contato inserido com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível inserir os Contato!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
            Throw
        End Try
    End Sub
    Public Sub AtualizaContatoCliente(Codigo As Integer, CodCliente As Integer, TipoContato As Integer, Contato As String, Email As String, Telefone As String)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "UPDATE Tbl_Contatos Set  Contato = @Contato, 
                                                    Email = @Email,
                                                    Telefone = @Telefone,
                                                    Depto = @TipoContato
                                            WHERE   CodCli = @CodCliente 
                                            AND     Codigo = @Codigo"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@CodCliente", CodCliente)
                    CMD.Parameters.AddWithValue("@Codigo", Codigo)
                    CMD.Parameters.AddWithValue("@Contato", Contato)
                    CMD.Parameters.AddWithValue("@Email", Email)
                    CMD.Parameters.AddWithValue("@Telefone", Telefone)
                    CMD.Parameters.AddWithValue("@TipoContato", TipoContato)
                    CMD.ExecuteNonQuery()
                    MsgBox("Dados de contato atualizado com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível atualizar os contatos!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
            Throw
        End Try
    End Sub
    Public Sub ExcluirContatoCliente(Codigo As Integer, CodCliente As Integer)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "DELETE FROM Tbl_Contatos WHERE CodCli = @CodCliente AND Codigo = @Codigo"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@CodCliente", CodCliente)
                    CMD.Parameters.AddWithValue("@Codigo", Codigo)
                    CMD.ExecuteNonQuery()
                    MsgBox("Dados de contato excluidos com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível atualizar os contatos!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
            Throw
        End Try
    End Sub
    Public Sub ExcluirContatosCliente(CodCliente As Integer)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "DELETE FROM Tbl_Contatos WHERE CodCli = @CodCliente "
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@CodCliente", CodCliente)
                    CMD.ExecuteNonQuery()
                    MsgBox("Dados de contato excluidos com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível atualizar os contatos!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
            Throw
        End Try
    End Sub
    Public Function ContatoCliente(CodCliente As Integer)
        Dim Listagem = New List(Of clsContato)
        Try
            Dim sql As String = "SELECT * FROM CS_ContatoClientes WHERE CodCli = @CodCliente"
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Using CMD As New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@CodCliente", CodCliente)
                    Using RDR As SqlDataReader = CMD.ExecuteReader()
                        While RDR.Read()
                            Dim Contato As New clsContato()
                            Contato.CodContato = Convert.ToInt32(RDR("Codigo"))
                            Contato.NomeContato = RDR("Contato").ToString()
                            Contato.Email = RDR("Email").ToString
                            Contato.Telefone = RDR("Telefone").ToString()
                            Listagem.Add(Contato)
                        End While
                    End Using
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível consultar o pedido de venda!" & vbCrLf & ex.Message)
            Throw
        End Try
        Return Listagem
    End Function
#End Region
#Region "FORNECEDOR"
    Public Sub IncluiContatoForn(CodForn As Integer, Contato As String, Email As String, Telefone As String, Depto As Integer)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "INSERT INTO Tbl_ContatosForn (CodForn,Contato, Email,Telefone,Depto) VALUES (@CodForn,@Contato,@Email,@Telefone,@Depto)"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@CodForn", CodForn)
                    CMD.Parameters.AddWithValue("@Contato", Contato)
                    CMD.Parameters.AddWithValue("@Email", Email)
                    CMD.Parameters.AddWithValue("@Telefone", Telefone)
                    CMD.Parameters.AddWithValue("@Depto", Depto)
                    CMD.ExecuteNonQuery()
                    MsgBox("Contato inserido com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível inserir o contato!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
            Throw
        End Try
    End Sub
    Public Sub AtualizaContatoForn(Codigo As Integer, CodForn As Integer, TipoContato As Integer, Contato As String, Email As String, Telefone As String)

        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "UPDATE Tbl_ContatosForn SET  Contato = @Contato, 
                                                        Email = @Email,
                                                        Telefone = @Telefone,
                                                        Depto = @TipoContato
                                                WHERE   CodForn = @CodForn AND Codigo = @Codigo"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@Codigo", Codigo)
                    CMD.Parameters.AddWithValue("@CodForn", CodForn)
                    CMD.Parameters.AddWithValue("@Contato", Contato)
                    CMD.Parameters.AddWithValue("@Email", Email)
                    CMD.Parameters.AddWithValue("@Telefone", Telefone)
                    CMD.Parameters.AddWithValue("@TipoContato", TipoContato)
                    CMD.ExecuteNonQuery()
                    MsgBox("Dados de contato atualizado com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível atualizar os contatos!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
            Throw
        End Try
    End Sub
    Public Sub ExcluirContatoForn(Codigo As Integer, CodForn As Integer)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "DELETE FROM Tbl_ContatosForn WHERE CodForn = @CodForn AND Codigo = @Codigo"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@CodForn", CodForn)
                    CMD.Parameters.AddWithValue("@Codigo", Codigo)
                    CMD.ExecuteNonQuery()
                    MsgBox("Dados de contato excluidos com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível atualizar os contatos!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
            Throw
        End Try
    End Sub
    Public Sub ExcluirContatosForn(CodForn As Integer)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "DELETE FROM Tbl_ContatosForn WHERE CodForn = @CodForn"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@CodForn", CodForn)
                    CMD.ExecuteNonQuery()
                    MsgBox("Dados de contato excluidos com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível atualizar os contatos!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
            Throw
        End Try
    End Sub
#End Region
#Region "TRANSPORTADORA"
    Public Sub IncluiContatoTransp(CodTransp As Integer, Contato As String, Email As String, Telefone As String, Depto As Integer)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "INSERT INTO Tbl_ContatosTransp (CodTransp,Contato,Email,Telefone,Depto) VALUES (@CodTransp,@Contato,@Email,@Telefone,@Depto)"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@CodTransp", CodTransp)
                    CMD.Parameters.AddWithValue("@Contato", Contato)
                    CMD.Parameters.AddWithValue("@Email", Email)
                    CMD.Parameters.AddWithValue("@Telefone", Telefone)
                    CMD.Parameters.AddWithValue("@Depto", Depto)
                    CMD.ExecuteNonQuery()
                    MsgBox("Contato inserido com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível inserir o contato!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
            Throw
        End Try
    End Sub
    Public Sub AtualizaContatoTransp(Codigo As Integer, CodTransp As Integer, TipoContato As Integer, Contato As String, Email As String, Telefone As String)

        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "UPDATE Tbl_ContatosTransp SET  Contato = @Contato, 
                                                        Email = @Email,
                                                        Telefone = @Telefone,
                                                        Depto = @TipoContato
                                                WHERE   CodTransp = @CodTransp AND Codigo = @Codigo"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@Codigo", Codigo)
                    CMD.Parameters.AddWithValue("@CodTransp", CodTransp)
                    CMD.Parameters.AddWithValue("@Contato", Contato)
                    CMD.Parameters.AddWithValue("@Email", Email)
                    CMD.Parameters.AddWithValue("@Telefone", Telefone)
                    CMD.Parameters.AddWithValue("@TipoContato", TipoContato)
                    CMD.ExecuteNonQuery()
                    MsgBox("Dados de contato atualizado com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível atualizar os contatos!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
            Throw
        End Try
    End Sub
    Public Sub ExcluirContatoTransp(Codigo As Integer, CodTransp As Integer)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "DELETE FROM Tbl_ContatosTransp WHERE CodTransp = @CodTransp AND Codigo = @Codigo"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@CodTransp", CodTransp)
                    CMD.Parameters.AddWithValue("@Codigo", Codigo)
                    CMD.ExecuteNonQuery()
                    MsgBox("Dados de contato excluidos com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível atualizar os contatos!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
            Throw
        End Try
    End Sub
    Public Sub ExcluirContatosTransp(CodTransp As Integer)
        Try
            Using cn = New SqlConnection(strConexao)
                cn.Open()
                Dim sql = "DELETE FROM Tbl_ContatosTransp WHERE CodTransp = @CodTransp"
                Using CMD = New SqlCommand(sql, cn)
                    CMD.Parameters.AddWithValue("@CodTransp", CodTransp)
                    CMD.ExecuteNonQuery()
                    MsgBox("Dados de contato excluidos com sucesso!", MsgBoxStyle.Information, "Sucesso")
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox("Não foi possível atualizar os contatos!" & vbCrLf & vbCrLf & ex.Message, vbCritical)
            Throw
        End Try
    End Sub
#End Region
End Class
