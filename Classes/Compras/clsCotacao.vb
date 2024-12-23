Imports System.Data
Imports System.Text
Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports Xceed.Wpf.Toolkit
Namespace Classes.Compras
    ''' <summary>
    ''' Esta classe representa todas as rotinas acerca das transações de cotação de produtos.
    ''' </summary>
    Public Class clsCotacao
        Inherits clsCompras
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property DataInicio As Date
        Public Property DataFim As Date
#End Region
#Region "CONSTRUTORES"
        Public Sub New()

        End Sub
        Public Sub New(_codigosimples As Integer, _quantidade As Integer, _valorunit As Decimal, _valortotal As Decimal)
            CodItem = _codigosimples
            Quantidade = _quantidade
            ValorUnitario = _valorunit
            ValorTotal = _valortotal
        End Sub
#End Region
#Region "METODOS"
        ''' <summary>
        ''' Este metódo registra os dados de uma cotação no banco de dados.
        ''' </summary>
        ''' <param name="Descricao">Representa a descrição de uma cotação do tipo string.</param>
        ''' <param name="DataInicio">Representa a data de inicio da cotação do tipo string</param>
        ''' <param name="DataFim">Representa a data final da cotação do tipo string</param>
        Public Sub SalvaCotacao(Descricao As String, DataInicio As DateTime, DataFim As DateTime)
            Dim sql As String = "INSERT   Tbl_Cotacao         (Descricao, 
                                                        DataInicial, 
                                                        DataFinal,
                                                        Status,
                                                        DataCriacao) 
                                             VALUES     (@Descricao, 
                                                        @DataInicio, 
                                                        @DataFim, 
                                                        0,
                                                        GETDATE())"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@Descricao", Descricao),
                         New SqlParameter("@DataInicio", DataInicio),
                         New SqlParameter("@DataFim", DataFim)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo grava os dados do itens que pertencem a uma cotação.
        ''' </summary>
        ''' <param name="CodCotacao">Representa o código de uma cotação do tipo integer.</param>
        ''' <param name="Quantidade">Representa a quantidade de itens que deve ser cotada do tipo integer.</param>
        ''' <param name="CodSimples">Represente o código do item que está na cotação do tipo integer.</param>
        Public Sub SalvaItemCotacao(CodCotacao As Integer, Quantidade As Integer, CodSimples As Integer)
            Dim sql As String = "INSERT INTO Tbl_ItemCotacao (CodCotacao,
                                                            CodSimples,
                                                            Quantidade) 
                                                    VALUES(@CodCotacao,
                                                            @CodSimples,
                                                            @Quantidade)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodCotacao", CodCotacao),
                         New SqlParameter("@Quantidade", Quantidade),
                         New SqlParameter("@CodSimples", CodSimples)
        }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        ''' <summary>
        ''' Este metódo realiza a exclusão de uma cotação do banco de dados.
        ''' </summary>
        ''' <param name="CodCotacao">Representa o código da cotação.</param>
        Public Sub ExcluirCotacao(CodCotacao As Integer)
            Dim sqlDet As String = "DELETE FROM Tbl_ItemCotacao WHERE CodCotacao = @CodCotacao"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodCotacao", CodCotacao)
        }
            ClasseConexao.Operar(sqlDet, parameters)

            Dim sql As String = "DELETE FROM Tbl_Cotacao WHERE CodCotacao = @CodCotacao"

            ClasseConexao.Operar(sql, parameters)

            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo consulta se um item está em um cotação, caso positivo o item é removido da cotação.
        ''' </summary>
        ''' <param name="CodCotacao">Representa o código da cotação.</param>
        ''' <param name="CodSimples">Represente o código do item que está na cotação do tipo integer.</param>
        Public Sub ExcluirItemCotacao(CodCotacao As Integer, CodSimples As Integer)
            Dim parameters As SqlParameter() = {
                        New SqlParameter("@CodCotacao", CodCotacao),
                        New SqlParameter("@CodSimples", CodSimples)
            }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_CotacaoDetalhes WHERE CodCotacao = @CodCotacao AND CodSimples = @CodSimples", parameters)

            If Tabela.Rows.Count > 0 Then
                Dim parametersdel As SqlParameter() = {
                        New SqlParameter("@CodCotacao", CodCotacao),
                        New SqlParameter("@CodSimples", CodSimples)
            }
                Dim sql As String = "DELETE FROM Tbl_ItemCotacao WHERE CodCotacao = @CodCotacao AND CodSimples = @CodSimples"
                ClasseConexao.Operar(sql, parametersdel)
            Else
                Exit Sub
            End If
        End Sub
        ''' <summary>
        ''' Este metódo atualiza os dados de uma cotação no banco de dados.
        ''' </summary>
        ''' <param name="CodCotacao">Representa o código de uma cotação do tipo integer.</param>
        ''' <param name="Descricao">Representa a descrição de uma cotação do tipo string.</param>
        ''' <param name="DataInicio">Representa a data de inicio da cotação do tipo string</param>
        ''' <param name="DataFim">Representa a data final da cotação do tipo string</param>
        Public Sub AtualizaCotacao(CodCotacao As Integer, Descricao As String, DataInicio As Date, DataFim As Date)
            Dim sql As String = "UPDATE Tbl_Cotacao  SET      Descricao = @Descricao,
                                                        DataInicial = @DataInicio,
                                                        DataFinal = @DataFim
                                                WHERE   CodCotacao = @CodCotacao"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodCotacao", CodCotacao),
                         New SqlParameter("@Descricao", Descricao),
                         New SqlParameter("@DataInicio", DataInicio),
                         New SqlParameter("@DataFim", DataFim)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo atualiza os itens de uma cotação.
        ''' </summary>
        ''' <param name="CodCotacao">Representa o código de uma cotação do tipo integer.</param>
        ''' <param name="Quantidade">Representa a quantidade do item que foi cotada do tipo integer.</param>
        ''' <param name="CodSimples">Represente o código do item que está na cotação do tipo integer.</param>
        Public Sub AtualizaItemCotacao(CodCotacao As Integer, Quantidade As Integer, CodSimples As Integer)
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodCotacao", CodCotacao),
                         New SqlParameter("@CodSimples", CodSimples)
        }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_CotacaoDetalhes WHERE CodCotacao = @CodCotacao AND CodSimples = @CodSimples", parameters)
            If Tabela.Rows.Count > 0 Then
                Dim parametersUp As SqlParameter() = {
                         New SqlParameter("@CodCotacao", CodCotacao),
                         New SqlParameter("@CodSimples", CodSimples),
                         New SqlParameter("@Quantidade", Quantidade)
        }
                Dim sql As String = "UPDATE     Tbl_ItemCotacao SET      
                                                            Quantidade = @Quantidade
                                                    WHERE   CodCotacao = @CodCotacao
                                                    AND     CodSimples = @CodSimples"

                ClasseConexao.Operar(sql, parametersUp)
            Else
                SalvaItemCotacao(CodCotacao, Quantidade, CodSimples)
            End If
        End Sub
        ''' <summary>
        ''' Este metódo atualiza o estado de uma cotação no banco de dados.
        ''' </summary>
        ''' <param name="Status">Representa o estado desejado da uma cotação do tipo integer.</param>
        ''' <param name="CodCotacao">Represente o código da cotação do tipo integer.</param>
        Public Sub AtualizaStatusCotacao(Status As Integer, CodCotacao As Integer)
            'STATUS DA COTAÇÃO
            '0 - EM ABERTO
            '1 - EM COTAÇÃO
            '2 - ENCERRADO
            '3 - APURADO
            '4 - CONCLUÍDO
            Dim sql As String = "UPDATE Tbl_Cotacao SET Status = @Status WHERE CodCotacao = @CodCotacao"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodCotacao", CodCotacao),
                         New SqlParameter("@Status", Status)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cotação atualizada com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo gravar a preceficação do fornecedor no banco de dados.
        ''' </summary>
        ''' <param name="CodForn">Representa o código do fornecedot do tipo integer.</param>
        ''' <param name="CodSimples">Representa o código do item do tipo integer.</param>
        ''' <param name="CodCotacao">Represente o código da cotação do tipo integer.</param>
        ''' <param name="Prazo">Representa o valor do prazo do tipo string.</param>
        ''' <param name="Desconto">Represente o valor do desconto do tipo decimal.</param>
        ''' <param name="Frete">Represente o valor do desconto do tipo decimal.</param>
        ''' <param name="Valor">Represente o valor do frete do tipo decimal.</param>
        ''' <param name="Total">Represente o valor total do item do tipo decimal.</param>
        ''' <param name="Quantidade">Representa o quantidade do item do tipo integer.</param>
        Public Sub SalvaItemFornecedor(CodForn As Integer, CodSimples As Integer, CodCotacao As String, Prazo As String, Desconto As Decimal, Frete As Decimal, Valor As Decimal, Total As Decimal, Quantidade As Integer)
            Dim sql As String = "INSERT   Tbl_ItemCotacaoFornecedor  (CodForn,
                                                        CodSimples, 
                                                        CodCotacao, 
                                                        Prazo,
                                                        Desconto,
                                                        Frete,
                                                        Valor,
                                                        Total,
                                                        DataCadastro,
                                                        Quantidade) 
                                             VALUES     (@CodForn,
                                                        @CodSimples, 
                                                        @CodCotacao, 
                                                        @Prazo,
                                                        @Desconto,
                                                        @Frete,
                                                        @Valor,
                                                        @Total,
                                                        GETDATE(),
                                                        @Quantidade)"
            Dim parameters As SqlParameter() = {
                New SqlParameter("@CodForn", CodForn),
                New SqlParameter("@CodSimples", CodSimples),
                New SqlParameter("@CodCotacao", CodCotacao),
                New SqlParameter("@Prazo", Prazo),
                New SqlParameter("@Desconto", Desconto),
                New SqlParameter("@Frete", Frete),
                New SqlParameter("@Valor", Valor),
                New SqlParameter("@Total", Total),
                New SqlParameter("@Quantidade", Quantidade)
        }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        ''' <summary>
        ''' Este metódo executa uma procedure que criar o pedido de compra para cada fornecedor de acordo com o vencedor.
        ''' </summary>
        ''' <param name="CODCOTACAO">Representa o código da cotação do tipo integer.</param>
        ''' <param name="CODFORN">Represente o código do fornecedor do tipo integer.</param>
        Public Sub CriarPedidoCotacao(CODCOTACAO As Integer, CODFORN As Integer)
            Dim parameters As SqlParameter() = {
                New SqlParameter("@CODCOTACAO", CODCOTACAO),
                New SqlParameter("@CODFORN", CODFORN)
        }
            ClasseConexao.ExecutarProcedure("spCriaPedidoCompraCotacao", parameters)
        End Sub
#End Region
#Region "FUNCOES"
        ''' <summary>
        ''' Esta função consulta os dados da cotação conforme filtro informados e retorna num componente DataGridView.
        ''' </summary>
        ''' <param name="Status">Representa o status de uma cotação do tipo string.</param>
        ''' <param name="CodCotacao">Representa o código de uma cotação do tipo integer.</param>
        ''' <param name="Descricao">Representa a descrição de uma cotação do tipo string.</param>
        ''' <param name="DataInicio">Representa a data de inicio da cotação do tipo string</param>
        ''' <param name="DataFim">Representa a data final da cotação do tipo string</param>
        Public Function PesquisaCotacao(Status As String, CodCotacao As String, Descricao As String, DataInicio As String, DataFim As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_Cotacao WHERE 1=1")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Em Aberto"
                        sql.AppendLine("AND Status LIKE 'Em Aberto'")
                    Case "Concluído"
                        sql.AppendLine("AND Status LIKE 'Concluído'")
                    Case "Todos"
                        sql.AppendLine("AND Status IS NOT NULL")
                End Select

                If Not String.IsNullOrEmpty(CodCotacao) Then
                    sql.AppendLine("AND CodCotacao = @CodCotacao")
                    parameters.Add(New SqlParameter("@CodCotacao", CodCotacao))
                End If

                If Not String.IsNullOrEmpty(Descricao) Then
                    sql.AppendLine("AND Descricao LIKE @Descricao")
                    parameters.Add(New SqlParameter("@Descricao", Descricao))
                End If

                'Pesquisa pela data de inicio
                If Not String.IsNullOrEmpty(DataInicio) And Not String.IsNullOrEmpty(DataFim) Then
                    sql.AppendLine("AND DataInicial BETWEEN @DataInicio AND @DataFim")
                    parameters.Add(New SqlParameter("@DataInicio", DataInicio))
                    parameters.Add(New SqlParameter("@DataFim", DataFim))
                ElseIf Not String.IsNullOrEmpty(DataInicio) And String.IsNullOrEmpty(DataFim) Then
                    sql.AppendLine("AND DataInicial >= @DataInicio")
                    parameters.Add(New SqlParameter("@DataInicio", DataInicio))
                ElseIf String.IsNullOrEmpty(DataInicio) And Not String.IsNullOrEmpty(DataFim) Then
                    sql.AppendLine("AND DataFinal <= @DataFinal")
                    parameters.Add(New SqlParameter("@DataFim", DataFim))
                Else
                    sql.AppendLine("AND DataInicial IS NOT NULL")
                End If

                sql.AppendLine("ORDER BY DataInicial DESC")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function
        ''' <summary>
        ''' Este metódo consulta todos o preços digitados pelo fornecedor para um cotação e retorna no componente DataGridView.
        ''' </summary>
        ''' <param name="CodCotacao">Representa o código da cotação do tipo integer.</param>
        ''' <param name="Total">Represente o valor total da precificação do fornecedor do tipo decimal.</param>
        Public Function ConsultaItensCotacao(sql As String, CodCotacao As Integer, Total As Decimal)

            Dim parameters As SqlParameter() = {
                New SqlParameter("@CODCOTACAO", CodCotacao),
                New SqlParameter("@Total", Total)
        }
            Return ClasseConexao.Consultar(sql, parameters)
        End Function
        ''' <summary>
        ''' Esta função consulta as cotação no banco de dados.
        ''' </summary>
        ''' <param name="CodCotacao">Representa o número da cotação.</param>
        ''' <param name="sql">Representa a query sql necessária para realizar a consulta.</param>
        ''' <returns>Retorna os itens de uma cotação.</returns>
        Public Function ConsultaCotacao(sql As String, Optional CodCotacao As Integer = 0)
            If CodCotacao <> 0 Then
                Dim parameters As SqlParameter() = {
                New SqlParameter("@CodCotacao", CodCotacao)
            }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
        ''' <summary>
        ''' Esta função consulta todas as cotações conforme o código do status da cotação.
        ''' </summary>
        ''' <param name="CodStatus">Representa o código do status da cotação.</param>
        ''' <param name="CodCotacao">Representa o código da cotação do tipo integer.</param>
        ''' <param name="Descricao">Representa a descrição da cotação</param>
        ''' <returns>Retorna os dados da cotação conforme status.</returns>
        Public Function ConsultaStatusCotacao(CodStatus As Integer, Optional CodCotacao As Integer = 0, Optional Descricao As String = Nothing)

            If CodCotacao <> 0 And Descricao Is Nothing Then
                Dim sql As String = "SELECT * FROM Cs_CotacaoFornecedorDetalhes WHERE CodStatus = @CodStatus AND CodCotacao = @CodCotacao"
                Dim parameters As SqlParameter() = {
            New SqlParameter("@CodStatus", CodStatus),
            New SqlParameter("@CodCotacao", CodCotacao)
         }
                Return ClasseConexao.Consultar(sql, parameters)
            ElseIf CodCotacao = 0 And Descricao IsNot Nothing Then
                Dim sql As String = "SELECT * FROM Cs_CotacaoFornecedorDetalhes WHERE CodStatus = @CodStatus AND Descricao LIKE @Descricao"
                Dim parameters As SqlParameter() = {
            New SqlParameter("@CodStatus", CodStatus),
            New SqlParameter("@Descricao", Descricao)
         }
                Return ClasseConexao.Consultar(sql, parameters)
            ElseIf CodCotacao <> 0 And Descricao IsNot Nothing Then
                Dim sql As String = "SELECT * FROM Cs_CotacaoFornecedorDetalhes WHERE CodStatus = @CodStatus AND CodCotacao = @CodCotacao AND Descricao LIKE @Descricao"
                Dim parameters As SqlParameter() = {
            New SqlParameter("@CodStatus", CodStatus),
            New SqlParameter("@CodCotacao", CodCotacao),
            New SqlParameter("@Descricao", Descricao)
         }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Dim sql As String = "SELECT * FROM Cs_CotacaoFornecedorDetalhes WHERE CodStatus = @CodStatus"
                Dim parameters As SqlParameter() = {
            New SqlParameter("@CodStatus", CodStatus)
         }
                Return ClasseConexao.Consultar(sql, parameters)
            End If
        End Function
        ''' <summary>
        ''' Este metódo executa uma procedure que faz a apuração do melhor fornecedor e retorna o resultado num componente DataGridView.
        ''' </summary>
        ''' <param name="CODCOTACAO">Representa o código da cotação do tipo integer.</param>
        Public Function ApuraCotacao(CodCotacao As Integer)
            Dim parameters As SqlParameter() = {
                New SqlParameter("@CODCOTACAO", CodCotacao)
        }
            Return ClasseConexao.ExecProcedureRetorno("spApuraCotacao", parameters)

        End Function

#End Region
    End Class
End Namespace
