Imports System.Data
Imports System.Text
Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports Xceed.Wpf.Toolkit
Namespace Classes.Comercial
    Public Class ClsOrcamento
        Inherits clsComercial
        Dim ClasseConexao As New ConexaoSQLServer

#Region "METODOS"
        Public Sub SalvarOrcamento(DataOrcamento As Date, NomeCliente As Integer, CustoTotal As Decimal, CodItem As Integer, Bordas As String, Quantidade As Integer, Miolo As Integer, Estampa As Integer, Loja As Integer)
            Dim sql As String = "INSERT tbl_Orcamento (DataOrcamento, 
                                        Cliente, 
                                        TotalPago, 
                                        Cod_Simples,
                                        Bordas,
                                        Status,
                                        DataCriacao,
                                        Quantidade,
                                        Miolo,
                                        Estampa,
                                        Loja)
                          VALUES		(@DataOrcamento, 
        	                            @NomeCliente, 
        	                            @Total,
        	                            @Cod_Simples,
        	                            @Bordas,
        	                            0,
        	                            GETDATE(),
        	                            @Quantidade,
        	                            @Miolo,
                                        @Estampa, 
                                        @Loja)"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@DataOrcamento", DataOrcamento),
                     New SqlParameter("@NomeCliente", NomeCliente),
                     New SqlParameter("@Total", CustoTotal),
                     New SqlParameter("@Cod_Simples", CodItem),
                     New SqlParameter("@Bordas", Bordas),
                     New SqlParameter("@Quantidade", Quantidade),
                     New SqlParameter("@Miolo", Miolo),
                     New SqlParameter("@Estampa", Estampa),
                     New SqlParameter("@Loja", Loja)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub SalvarDetOrcamento(CodOrcamento As Integer, Quantidade As Integer, ValorUnit As Decimal, CodItem As Integer, ValorTotal As Decimal)
            Dim sql As String = "INSERT INTO tbl_OrcamentoDet (CodigoOrcamento,ValorUnit,Quantidade,Cod_Simples,ValorTotal) VALUES  (@CODIGOORCAMENTO,@VUNITARIO,@QUANTIDADE,@CODIGO,@TOTAL)"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CODIGOORCAMENTO", CodOrcamento),
                     New SqlParameter("@QUANTIDADE", Quantidade),
                     New SqlParameter("@VUNITARIO", ValorUnit),
                     New SqlParameter("@CODIGO", CodItem),
                     New SqlParameter("@TOTAL", ValorTotal)
    }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub AtualizaOrcamento(CodOrcamento As Integer, DataOrcamento As Date, NomeCliente As String, Bordas As String, Miolo As Integer, Estampa As Integer, Loja As Integer)
            Dim sql As String = "UPDATE tbl_Orcamento SET  DataOrcamento = @DATAORCAMENTO,Cliente = @CLIENTE,Bordas = @BORDAS,Miolo = @MIOLO, Estampa = @Estampa Loja = @Loja WHERE CodOrcamento = @CodOrcamento"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodOrcamento", CodOrcamento),
                     New SqlParameter("@DATAORCAMENTO", DataOrcamento),
                     New SqlParameter("@CLIENTE", NomeCliente),
                     New SqlParameter("@BORDAS", Bordas),
                     New SqlParameter("@MIOLO", Miolo),
                     New SqlParameter("@Estampa", Estampa),
                     New SqlParameter("@Loja", Loja)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaDetOrcamento(CodOrcamento As Integer, ValorUnit As Decimal, Quantidade As Integer, CodSimples As Integer, ValorTotal As Decimal)
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodOrcamento", CodOrcamento),
                     New SqlParameter("@CodSimples", CodSimples)
    }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_OrcamentoDetalhes WHERE CodOrcamento = @CodOrcamento AND CodItem = @CodSimples", parameters)
            If Tabela.Rows.Count > 0 Then
                Dim parametersUp As SqlParameter() = {
                     New SqlParameter("@CodOrcamento", CodOrcamento),
                     New SqlParameter("@CodSimples", CodSimples),
                     New SqlParameter("@Quantidade", Quantidade),
                     New SqlParameter("@ValorUnit", ValorUnit),
                     New SqlParameter("@ValorTotal", ValorTotal)
    }
                Dim sql As String = "UPDATE tbl_OrcamentoDet SET    Quantidade = @Quantidade,
                                                                ValorUnit = @ValorUnit,
                                                                ValorTotal = @ValorTotal
                                                        WHERE   CodigoOrcamento = @CodOrcamento
                                                        AND     Cod_Simples = @CodSimples"
                ClasseConexao.Operar(sql, parametersUp)
            Else
                SalvarDetOrcamento(CodOrcamento, Quantidade, ValorUnit, CodSimples, ValorTotal)
            End If

        End Sub
        Public Sub ExcluirOrcamento(CodOrcamento As Integer)
            Dim sql As String = "DELETE FROM tbl_Orcamento WHERE CodOrcamento = @CODIGOORCAMENTO"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CODIGOORCAMENTO", CodOrcamento)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Orçamento excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluirDetOrcamento(CodOrcamento As Integer)
            Dim sql As String = "DELETE FROM tbl_OrcamentoDet WHERE CodigoOrcamento = @CODIGOORCAMENTO"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CODIGOORCAMENTO", CodOrcamento)
    }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub ExcluirItemOrcamento(CodOrcamento As Integer, CodSimples As Integer)
            Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodOrcamento", CodOrcamento),
                    New SqlParameter("@CodSimples", CodSimples)
        }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT *FROM Cs_OrcamentoDetalhes WHERE CodOrcamento = @CodOrcamento AND Cod_Simples = @CodSimples", parameters)

            If Tabela.Rows.Count > 0 Then
                Dim parametersDel As SqlParameter() = {
                    New SqlParameter("@CodOrcamento", CodOrcamento),
                    New SqlParameter("@CodSimples", CodSimples)
        }

                Dim sql As String = "DELETE FROM tbl_OrcamentoDet WHERE CodigoOrcamento = @CodOrcamento AND Cod_Simples = @CodSimples"
                ClasseConexao.Operar(sql, parametersDel)
            Else
                Exit Sub
            End If
        End Sub
        Public Sub AtualizaStatusOrcamento(CodOrcamento As Integer, Status As Integer)
            Dim sql As String = "UPDATE Tbl_Orcamento SET Status = @Status WHERE CodOrcamento = @CODORCARMENTO"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@Status", Status),
                     New SqlParameter("@CODORCARMENTO", CodOrcamento)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Orçamento atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo executa uma procedure que criar o pedido de compra para cada fornecedor de acordo com o vencedor.
        ''' </summary>
        ''' <param name="CodOrcamento">Representa o código do orçamento do tipo integer.</param>
        ''' <param name="CodCli">Represente o código do cliente do tipo integer.</param>
        Public Sub CriarPedidoOrcamento(CodOrcamento As Integer, CodCli As Integer)
            Dim parameters As SqlParameter() = {
            New SqlParameter("@CodOrcamento", CodOrcamento),
            New SqlParameter("@CodCli", CodCli)
    }
            ClasseConexao.ExecutarProcedure("spCriaPedidoVendaOrc", parameters)
        End Sub

        ''' <summary>
        ''' Este metódo registra uma estampa no banco de dados.
        ''' </summary>
        ''' <param name="Descricao">Representa a descrição da estampa.</param>
        ''' <param name="Impressa">Representa se a estampa é impresa ou não.</param>
        Public Sub SalvarEstampa(Descricao As String, Impressa As Boolean)
            Dim sql As String = "INSERT INTO tbl_Estampas (Descricao,Impressa, Inativo) VALUES (@DESCRICAO,@IMPRESSA,0)"
            Dim parameters As SqlParameter() = {
           New SqlParameter("@DESCRICAO", Descricao),
           New SqlParameter("@IMPRESSA", Impressa)
   }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo atualiza os dados da estampa no banco de dados.
        ''' </summary>
        ''' <param name="Codigo">Representa o código da estampa no banco de dados.</param>
        ''' <param name="Descricao">Representa a descrição da estampa.</param>
        ''' <param name="Impressa">Representa se a estampa é impresa ou não.</param>
        ''' <param name="Inativo">Representa a a situação da estampa.</param>
        Public Sub AtualizarEstampa(Codigo As Integer, Descricao As String, Impressa As Boolean, Inativo As Boolean)
            Dim sql As String = "UPDATE tbl_Estampas SET Descricao = @DESCRICAO, Impressa = @IMPRESSA, Inativo = @INATIVO WHERE Codigo = @CODIGO"
            Dim parameters As SqlParameter() = {
         New SqlParameter("@CODIGO", Codigo),
         New SqlParameter("@DESCRICAO", Descricao),
         New SqlParameter("@IMPRESSA", Impressa),
         New SqlParameter("@INATIVO", Inativo)
   }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo realiza a exclusão os dados da estampa no banco de dados.
        ''' </summary>
        ''' <param name="Codigo">Representa o código da estampa no banco de dados.</param>
        Public Sub ExcluirEstampa(Codigo As Integer)
            Dim sql As String = "DELETE FROM tbl_Estampas WHERE Codigo = @CODIGO"
            Dim parameters As SqlParameter() = {
         New SqlParameter("@CODIGO", Codigo)
   }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
#End Region
#Region "FUNCOES"
        Public Function PesquisaOrcarmento(Status As String, CodOrcamento As String, NomeCliente As String, Produto As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_Orcamentos WHERE 1=1 ")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Em Aberto"
                        sql.AppendLine("AND Status LIKE 'EM ABERTO'")
                    Case "Produzido"
                        sql.AppendLine("AND Status LIKE 'PRODUZIDO'")
                    Case "Concluído"
                        sql.AppendLine("AND Status LIKE 'CONCLUÍDO'")
                    Case "Finalizado"
                        sql.AppendLine("AND Status LIKE 'FINALIZADO'")
                    Case "Todos"
                        sql.AppendLine("AND Status IS NOT NULL")
                End Select

                If Not String.IsNullOrEmpty(CodOrcamento) Then
                    sql.AppendLine("AND CodOrcamento = @CodOrcamento")
                    parameters.Add(New SqlParameter("@CodOrcamento", CodOrcamento))
                End If

                If Not String.IsNullOrEmpty(NomeCliente) Then
                    sql.AppendLine("AND Entidade LIKE @NomeCliente")
                    parameters.Add(New SqlParameter("@NomeCliente", NomeCliente))
                End If

                If Not String.IsNullOrEmpty(Produto) Then
                    sql.AppendLine("AND Produto LIKE @Produto")
                    parameters.Add(New SqlParameter("@Produto", Produto))
                End If

                sql.AppendLine("ORDER BY DataOrcamento DESC")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function
        ''' <summary>
        ''' Esta função consulta um orçamento no banco de dados.
        ''' </summary>
        ''' <param name="CodOrcamento">Representa o número do orçamento.</param>
        ''' <param name="sql">Representa a query sql necessária para realizar a consulta.</param>
        ''' <returns>Retorna os itens de uma cotação.</returns>
        Public Function ConsultaOrcamento(sql As String, Optional CodOrcamento As Integer = 0)

            If CodOrcamento <> 0 Then
                Dim parameters As SqlParameter() = {
                 New SqlParameter("@CodOrcamento", CodOrcamento)
            }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
        ''' <summary>
        ''' Esta função consulta os dados de um estampas no banco de dados.
        ''' </summary>
        ''' <param name="sql">Query sql necessária para a consulta.</param>
        ''' <returns>Retorna os dados da estampa.</returns>
        Public Function ConsultaEstampa(sql As String)
            Return ClasseConexao.Consultar(sql, Nothing)
        End Function
#End Region
    End Class

End Namespace


