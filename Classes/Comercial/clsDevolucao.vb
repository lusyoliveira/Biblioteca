Imports System.Data
Imports System.Text
Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports Xceed.Wpf.Toolkit
Namespace Classes.Comercial

    ''' <summary>
    ''' Esta classe representa todas as rotinas que envolve uma devolução.
    ''' </summary>
    Public Class clsDevolucao
        Inherits clsComercial
        Dim ClasseConexao As New ConexaoSQLServer

#Region "METODOS"
        Public Sub SalvarDevolucao(DataDevolucao As Date, CodVendas As Integer, Loja As Integer, Cliente As Integer, Motivo As Integer, Total As Decimal, Observacao As String, Optional CodOrcamento As Integer? = Nothing)
            Dim sql As String = "INSERT Tbl_Devolucao (DataDevolucao, 
                                                CodVendas,
                                                CodOrcamento,
                                                Loja,
                                                Cliente,
                                                Motivo,
                                                Status,
                                                Total, 
                                                DataCriacao, 
                                                Observacao)
                                  VALUES		(@DataDevolucao, 
                                                @CodVendas,
                                                @CodOrcamento,
                                                @Loja,
                                                @Cliente,
                                                @Motivo,
                                                0,
                                                @Total,
                                                GETDATE(),
                                                @Observacao)"
            Dim parameters As SqlParameter() = {
                        New SqlParameter("@DataDevolucao", DataDevolucao),
                        New SqlParameter("@CodVendas", CodVendas),
                        New SqlParameter("@Loja", Loja),
                        New SqlParameter("@Cliente", Cliente),
                        New SqlParameter("@Motivo", Motivo),
                        New SqlParameter("@Total", Total),
                        New SqlParameter("@Observacao", Observacao)
        }
            ' Adicionar parâmetros opcionais, verificando se são Nothing para definir como DBNull
            If CodOrcamento <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodOrcamento", CodOrcamento.Value)
            Else
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodOrcamento", DBNull.Value)
            End If

            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub SalvarDevolucaoDet(CodDevolucao As Integer, ValorUnit As Decimal, Quantidade As Integer, CodSimples As Integer, ValorTotal As Decimal)
            Dim sql As String = "INSERT INTO Tbl_DevolucaoDet (CodDevolucao,ValorUnit,Quantidade,Cod_Simples,ValorTotal) VALUES  (@CodDevolucao,@VUNITARIO,@QUANTIDADE,@CODIGO,@TOTAL)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodDevolucao", CodDevolucao),
                         New SqlParameter("@QUANTIDADE", Quantidade),
                         New SqlParameter("@VUNITARIO", ValorUnit),
                         New SqlParameter("@CODIGO", CodSimples),
                         New SqlParameter("@TOTAL", ValorTotal)
        }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub AtualizaDevolucao(CodDevolucao As Integer, DataDevolucao As Date, CodVendas As Integer, CodOrcamento As Integer, Loja As Integer, Cliente As Integer, Motivo As Integer, Total As Decimal, Observacao As String)
            Dim sql As String = "UPDATE Tbl_Devolucao SET DataDevolucao = @DataDevolucao, 
                                                    CodVendas = @CodVendas,
                                                    CodOrcamento = @CodOrcamento,
                                                    Loja = @Loja,
                                                    Cliente = @Cliente,
                                                    Motivo = @Motivo,
                                                    Status = @Status,
                                                    Total = @Total,
                                                    DataAlteracao = GETDATE(),
                                                    Observacao = @Observacao
                                            WHERE   CodDevolucao = @CodDevolucao"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodDevolucao", CodDevolucao),
                         New SqlParameter("@DataDevolucao", DataDevolucao),
                         New SqlParameter("@CodVendas", CodVendas),
                         New SqlParameter("@CodOrcamento", CodOrcamento),
                         New SqlParameter("@Loja", Loja),
                         New SqlParameter("@Cliente", Cliente),
                         New SqlParameter("@Motivo", Motivo),
                         New SqlParameter("@Total", Total),
                         New SqlParameter("@Observacao", Observacao)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo consulta o banco de dados para verificar se o item já existe para o pedido informado. Se existe o item no pedido informado retorna 1 e atualiza os dados do item.Caso o item não exista no pedido informado, ele é inserido no pedido.
        ''' </summary>
        ''' <param name="CodDevolucao">Representa o código da devolução do tipo integer.</param>
        ''' <param name="ValorUnit">Representa o valor unitário do item do tipo decimal.</param>
        ''' <param name="Quantidade">Representa a quantidade do item do tipo integer.</param>
        ''' <param name="CodSimples">Representa o código do item do tipo integer.</param>
        ''' <param name="ValorTotal">Representa o valor total do item do tipo decimal.</param>
        Public Sub AtualizaDetDevolucao(CodDevolucao As Integer, ValorUnit As Decimal, Quantidade As Integer, CodSimples As Integer, ValorTotal As Decimal)
            Dim parameters As SqlParameter() = {
                 New SqlParameter("@CodDevolucao", CodDevolucao),
                 New SqlParameter("@CodSimples", CodSimples)
    }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT COUNT(Cod_Simples) FROM Cs_DevolucaoDetalhes WHERE CodDevolucao = @CodDevolucao AND Cod_Simples = @CodSimples", parameters)
            If Tabela.Rows.Count > 0 Then
                Dim parametersUp As SqlParameter() = {
                 New SqlParameter("@CodDevolucao", CodDevolucao),
                 New SqlParameter("@CodSimples", CodSimples),
                 New SqlParameter("@Quantidade", Quantidade),
                 New SqlParameter("@ValorUnit", ValorUnit),
                 New SqlParameter("@ValorTotal", ValorTotal)
    }

                Dim sql As String = "UPDATE    Tbl_DevolucaoDet SET      
                                                            Quantidade = @Quantidade, 
                                                            ValorUnit = @ValorUnit, 
                                                            ValorTotal = @ValorTotal,
                                                    WHERE   CodDevolucao = @CodDevolucao
                                                    AND     CodSimples = @CodSimples"

                ClasseConexao.Operar(sql, parametersUp)
            Else
                SalvarDevolucaoDet(CodDevolucao, ValorUnit, Quantidade, CodSimples, ValorTotal)
            End If
        End Sub
        Public Sub ExcluirDevolucao(CodDevolucao As Integer)
            Dim sql As String = "DELETE FROM Tbl_Devolucao WHERE CodDevolucao = @CodDevolucao"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodDevolucao", CodDevolucao)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluirDetDevolucao(CodDevolucao As Integer)
            Dim sql As String = "DELETE FROM Tbl_DevolucaoDet WHERE CodDevolucao = @CodDevolucao"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodDevolucao", CodDevolucao)
        }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub ExcluirItemDevolucao(CodDevolucao As Integer, CodSimples As Integer)
            Dim parameters As SqlParameter() = {
                New SqlParameter("@CodDevolucao", CodDevolucao),
                New SqlParameter("@CodSimples", CodSimples)
    }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_CotacaoDetalhes WHERE CodCotacao = @CodCotacao AND CodSimples = @CodSimples", parameters)

            If Tabela.Rows.Count > 0 Then
                Dim parametersDel As SqlParameter() = {
                        New SqlParameter("@CodDevolucao", CodDevolucao),
                        New SqlParameter("@CodSimples", CodSimples)
            }
                Dim sql As String = "DELETE FROM Tbl_ItemCotacao WHERE CodCotacao = @CodCotacao AND CodSimples = @CodSimples"
                ClasseConexao.Operar(sql, parametersDel)
            Else
                Exit Sub
            End If
        End Sub
#End Region
#Region "FUNCOES"
        ''' <summary>
        ''' Esta função realiza a requisa de devoluções conforme parâmetros inforamdos.
        ''' </summary>
        ''' <param name="Status">Representa o status da devolução.</param>
        ''' <param name="CodDevolucao">Representa o cóidgo identificador da devolução.</param>
        ''' <param name="CodPedido">Representa o código identificador do pedido.</param>
        ''' <param name="Codorcamento">Representa o código identificador do orcamento.</param>
        ''' <param name="NomeCliente">Representa o nome do cliente.</param>
        ''' <param name="Motivo">Representa o motivo da devolução.</param>
        ''' <returns></returns>
        Public Function PesquisaDevolucao(Status As String, CodDevolucao As Integer, CodPedido As Integer, Codorcamento As Integer, NomeCliente As String, Motivo As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_Devolucao WHERE 1=1")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Em Aberto"
                        sql.AppendLine("AND Status LIKE 'EM ABERTO'")
                    Case "Devolver"
                        sql.AppendLine("AND Status LIKE 'DEVOLVER'")
                    Case "Devolvido"
                        sql.AppendLine("AND Status LIKE 'DEVOLVIDO'")
                    Case "Todos"
                        sql.AppendLine("AND Status IS NOT NULL")
                End Select

                If CodDevolucao <> 0 Then
                    sql.AppendLine("AND CodDevolucao = @CodDevolucao")
                    parameters.Add(New SqlParameter("@CodDevolucao", CodDevolucao))
                End If

                If CodPedido <> 0 Then
                    sql.AppendLine("AND CodPedido = @CodPedido")
                    parameters.Add(New SqlParameter("@CodPedido", CodPedido))
                End If

                If Codorcamento <> 0 Then
                    sql.AppendLine("AND CodOrcamento = @Codorcamento")
                    parameters.Add(New SqlParameter("@Codorcamento", Codorcamento))
                End If

                If Not String.IsNullOrEmpty(NomeCliente) Then
                    sql.AppendLine("AND Entidade LIKE @NomeCliente")
                    parameters.Add(New SqlParameter("@NomeCliente", NomeCliente))
                End If

                If Not String.IsNullOrEmpty(Motivo) Then
                    sql.AppendLine("AND Motivo LIKE @Motivo")
                    parameters.Add(New SqlParameter("@Motivo", Motivo))
                End If

                sql.AppendLine("ORDER BY DataDevolucao DESC")

                ' Chama a função Consultar com a query e os parâmetros
                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Esta função consulta as cotação no banco de dados.
        ''' </summary>
        ''' <param name="CodDevolucao">Representa o código identificado da devolucação.</param>
        ''' <param name="sql">Representa a query sql necessária para realizar a consulta.</param>
        ''' <returns>Retorna os itens de uma cotação.</returns>
        Public Function ConsultaDevolucao(sql As String, Optional CodDevolucao As Integer = 0)
            If CodDevolucao <> 0 Then
                Dim parameters As SqlParameter() = {
                New SqlParameter("@CodDevolucao", CodDevolucao)
            }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
#End Region
    End Class
End Namespace


