Imports Microsoft.Data.SqlClient
Imports Biblioteca.Classes.Conexao
Imports System.Data
Imports Xceed.Wpf.Toolkit
Namespace Classes.Configuracao

    Public Class clsConfiguracoes
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property CodDatabase As Integer
        Public Property NomeDatabase As String
        Public Property TipoOperacao As String
        Public Property Tipo As String
        Public Property CodOperacao As Integer
        Public Property Operacao As String
        Public Property CaminhoExportacao As String
        Public Property CaminhoImportacao As String
        Public Property EmailPedVenda As String
        Public Property EmailPedCompra As String
#End Region
#Region "METODOS"
        Public Sub InativaRotinas()
            ClasseConexao.ExecutarProcedure("spInativaRotinasVencida", Nothing)
        End Sub
        ''' <summary>
        ''' Este metódo registra um operação no banco de dados.
        ''' </summary>
        ''' <param name="Descricao">Representa a descrição de operação.</param>
        ''' <param name="TipoOperacao">Representa o tipode operação.</param>
        ''' <param name="Tipo">Representa o tipo.</param>
        ''' <param name="AtualizaCusto">Representa se a operação atualiza o custo.</param>
        ''' <param name="AtualizaEstoque">Representa se a operação atualiza o estoque.</param>
        ''' <param name="GeraPagamento">Representa se a operação gera um pagamento.</param>
        ''' <param name="GeraRecebimento">Representa se a operação gera um recebimento.</param>
        ''' <param name="Orcamento">Representa se a operação obriga informar um orçamento</param>
        ''' <param name="PedidoCompra">Representa se a operação obriga informar um pedido de compra.</param>
        Public Sub SalvarOperacoes(Descricao As String, TipoOperacao As String, Tipo As String, AtualizaCusto As Boolean, AtualizaEstoque As Boolean, GeraPagamento As Boolean, GeraRecebimento As Boolean, Orcamento As Boolean, PedidoCompra As Boolean)
            Dim sql As String = "INSERT INTO tbl_ConfigOperacao (Descricao,TipoOperacao, Tipo,AtualizaCusto,AtualizaEstoque) VALUES (@Descricao,@TipoOperacao,@Tipo,@AtualizaCusto,@AtualizaEstoque)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@Descricao", Descricao),
                         New SqlParameter("@TipoOperacao", TipoOperacao),
                         New SqlParameter("@Tipo", Tipo),
                         New SqlParameter("@AtualizaCusto", AtualizaCusto),
                         New SqlParameter("@AtualizaEstoque", AtualizaEstoque),
                         New SqlParameter("@GeraPagamento", GeraPagamento),
                         New SqlParameter("@GeraRecebimento", GeraRecebimento),
                         New SqlParameter("@Orcamento", Orcamento),
                         New SqlParameter("@PedidoCompra", PedidoCompra)
            }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo atualiza os dados de uma operação no banco de dados.
        ''' </summary>
        ''' <param name="Codigo">Representa o código da operação.</param>
        ''' <param name="Descricao">Representa a descrição de operação.</param>
        ''' <param name="TipoOperacao">Representa o tipode operação.</param>
        ''' <param name="Tipo">Representa o tipo.</param>
        ''' <param name="AtualizaCusto">Representa se a operação atualiza o custo.</param>
        ''' <param name="AtualizaEstoque">Representa se a operação atualiza o estoque.</param>
        ''' <param name="GeraPagamento">Representa se a operação gera um pagamento.</param>
        ''' <param name="GeraRecebimento">Representa se a operação gera um recebimento.</param>
        ''' <param name="Orcamento">Representa se a operação obriga informar um orçamento</param>
        ''' <param name="PedidoCompra">Representa se a operação obriga informar um pedido de compra.</param>
        Public Sub AtualizaOperacoes(Codigo As Integer, Descricao As String, TipoOperacao As String, Tipo As String, AtualizaCusto As Boolean, AtualizaEstoque As Boolean, GeraPagamento As Boolean, GeraRecebimento As Boolean, Orcamento As Boolean, PedidoCompra As Boolean)
            Dim sql As String = "UPDATE tbl_ConfigOperacao  SET Descricao = @Descricao,TipoOperacao = @TipoOperacao, Tipo = @Tipo,AtualizaCusto = @AtualizaCusto,AtualizaEstoque = @AtualizaEstoque WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@Codigo", Codigo),
                         New SqlParameter("@Descricao", Descricao),
                         New SqlParameter("@TipoOperacao", TipoOperacao),
                         New SqlParameter("@Tipo", Tipo),
                         New SqlParameter("@AtualizaCusto", AtualizaCusto),
                         New SqlParameter("@AtualizaEstoque", AtualizaEstoque),
                         New SqlParameter("@GeraPagamento", GeraPagamento),
                         New SqlParameter("@GeraRecebimento", GeraRecebimento),
                         New SqlParameter("@Orcamento", Orcamento),
                         New SqlParameter("@PedidoCompra", PedidoCompra)
            }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo registra as configurações do sistema.
        ''' </summary>
        ''' <param name="CaminhoExportacao">Representa um caminho padrão onde um arquivo deve ser exportado.</param>
        ''' <param name="CaminhoImportacao">Representa um caminho padrão onde um arquivo deve ser importado.</param>
        ''' <param name="EmailPedVenda">Representa o email configurado para envio do pedido de venda.</param>
        ''' <param name="EmailPedCompra">Representa o email configurado para envio do pedido de compra.</param>
        Public Sub SalvarConfigSistema(CaminhoExportacao As String, CaminhoImportacao As String, EmailPedVenda As Integer, EmailPedCompra As Integer)
            Dim sql As String = "INSERT INTO tbl_ConfigSistema (CaminhoExportacao,CaminhoImportacao, EmailPedVenda,EmailPedCompra) VALUES (@CaminhoExportacao,@CaminhoImportacao,@EmailPedVenda,@EmailPedCompra)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CaminhoExportacao", CaminhoExportacao),
                         New SqlParameter("@CaminhoImportacao", CaminhoImportacao),
                         New SqlParameter("@EmailPedVenda", EmailPedVenda),
                         New SqlParameter("@EmailPedCompra", EmailPedCompra)
            }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaConfigurcoes(CaminhoExportacao As String, CaminhoImportacao As String, EmailPedVenda As Integer, EmailPedCompra As Integer, ExibeDashboard As Integer, ValidaCPFCNPJ As Integer, PermiteMesmoDocumento As Integer)
            Dim sql As String = "UPDATE tbl_ConfigSistema SET   CaminhoExportacao = @CaminhoExportacao,
                                                            CaminhoImportacao = @CaminhoImportacao, 
                                                            EmailPedVenda = @EmailPedVenda,
                                                            EmailPedCompra = @EmailPedCompra, 
                                                            ExibeDashboard = @ExibeDashboard,
                                                            ValidaCPFCNPJ = @ValidaCPFCNPJ,
                                                            PermiteMesmoCNPJ = @PermiteMesmoDocumento
                                                            WHERE Codigo = 1"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CaminhoExportacao", CaminhoExportacao),
                         New SqlParameter("@CaminhoImportacao", CaminhoImportacao),
                         New SqlParameter("@EmailPedVenda", EmailPedVenda),
                         New SqlParameter("@EmailPedCompra", EmailPedCompra),
                         New SqlParameter("@ExibeDashboard", ExibeDashboard),
                         New SqlParameter("@ValidaCPFCNPJ", ValidaCPFCNPJ),
                         New SqlParameter("@PermiteMesmoDocumento", PermiteMesmoDocumento)
            }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
#End Region
#Region "FUNCOES"
        ''' <summary>
        ''' Esta função consulta os dados de operações de notas fiscais.
        ''' </summary>
        ''' <param name="sql">Query sql necessária para a consulta</param>
        ''' <returns>Retorna dados solicitados na query sql.</returns>
        Public Function ConsultaOperacoes(sql As String)
            Return ClasseConexao.Consultar(sql, Nothing)
        End Function
        ''' <summary>
        ''' Esta função consulta os dados de configurações do sistema.
        ''' </summary>
        ''' <param name="sql">Query sql necessária para a consulta</param>
        ''' <returns>Retorna dados solicitados na query sql.</returns>
        Public Function ConsultaConfiguracao(sql As String)
            Return ClasseConexao.Consultar(sql, Nothing)
        End Function
        ''' <summary>
        ''' Este função valida se a operação informada requer que o sistema realizar certas operações.
        ''' As validações disponíveis são: Custo, Estoque, Pagamento, Recebimento, Orcamento e Compra.
        ''' </summary>
        ''' <param name="Operacao">Nome da operação que deseja validar.</param>
        ''' <param name="CodOperacao">Código identificador da operação a ser validada.</param>
        ''' <returns>Retorna verdadeiro ou falso</returns>
        Public Function ValidaOpercaoes(Operacao As String, CodOperacao As Integer) As Boolean
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CODOPERACAO", CodOperacao)
            }
            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT * FROM tbl_ConfigOperacao WHERE Codigo = @CODOPERACAO", parameters)

            Select Case Operacao
                Case "CUSTO"
                    If Tabela.Rows(0)("AtualizaCusto").ToString() = True Then
                        Return True
                    Else
                        Return False
                    End If
                Case "ESTOQUE"
                    If Tabela.Rows(0)("AtualizaEstoque").ToString() = True Then
                        Return True
                    Else
                        Return False
                    End If
                Case "PAGAMENTO"
                    If Tabela.Rows(0)("GeraRecebimento").ToString() = True Then
                        Return True
                    Else
                        Return False
                    End If
                Case "RECEBIMENTO"
                    If Tabela.Rows(0)("GeraRecebimento").ToString() = True Then
                        Return True
                    Else
                        Return False
                    End If
                Case "ORCAMENTO"
                    If Tabela.Rows(0)("Orcamento").ToString() = True Then
                        Return True
                    Else
                        Return False
                    End If
                Case "COMPRA"
                    If Tabela.Rows(0)("PedidoCompra").ToString() = True Then
                        Return True
                    Else
                        Return False
                    End If
            End Select
        End Function
#End Region
    End Class
End Namespace

