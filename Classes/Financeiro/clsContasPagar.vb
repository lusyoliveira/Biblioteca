Imports System.Data
Imports System.Text
Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports Xceed.Wpf.Toolkit
Namespace Classes.Financeiro
    ''' <summary>
    ''' Esta classe representa todas as rotina de contas a pagar do sistema.Esta é uma classe derivada de clsFinanceiro.
    ''' </summary>
    Public Class clsContasPagar
        Inherits clsFinanceiro
        Dim ClasseConexao As New ConexaoSQLServer
#Region "METODOS"
        ''' <summary>
        ''' Este metódo realiza a gravação do pagamento no banco de dados.
        ''' </summary>
        ''' <param name="CodCompra">Representa o código do lançamento da nota fiscal relacionada ao pagamento do tipo integer.</param>
        ''' <param name="CodDevolucao">Representa o código da devolução relacionada ao pagamento do tipo integer.</param>
        ''' <param name="DtVencimento">Representa a data em que o pagamento irá vencer do tipo date.</param>
        ''' <param name="Parcelas">Represente o número do parcelas do tipo integer.</param>
        ''' <param name="ValorParcela">Representa o valor da parcelas a ser pago do tipo decimal.</param>
        ''' <param name="ValorPago">Representa o valor pago do pagamento do tipo decimal.</param>
        ''' <param name="Complemento">Representa a descrição do pagamento do tipo string.</param>
        ''' <param name="Desconto">Representa o valor de desconto aplicado no pagamento do tipo decimal.</param>
        ''' <param name="Frete">Representa o valor de frete aplicado no pagamento do tipo decimal.</param>
        ''' <param name="Taxa">Representa o valor de taxa aplicado no pagamento do tipo decimal.</param>
        ''' <param name="Acrescimo">Representa o valor de acréscimo aplicado no pagamento do tipo decimal.</param>
        ''' <param name="Entidade">Representa o código da entidade para qual o pagamento foi realizado do tipo integer.</param>
        ''' <param name="Cobranca">Representa o código da  cobrança do pagamento do tipo integer.</param>
        ''' <param name="FormaPagto">Representa o código do da forma de pagamento do tipo integer.</param>
        Public Sub GerarContasaPagar(DtVencimento As Date, Parcelas As Integer, ValorParcela As Decimal, ValorPago As Decimal, Complemento As String, Desconto As Integer, Frete As Integer, Taxa As Decimal, Acrescimo As Decimal, Entidade As Integer, Cobranca As Integer, FormaPagto As Integer, Optional CodCompra As Integer? = 0, Optional CodDevolucao As Integer? = 0, Optional CodRecebimento As Integer? = 0)

            Dim sql As String = "INSERT INTO Tbl_ContasAPagar (CodCompra,
                                                        Dt_Vencimento,
                                                        Parcelas,
                                                        Valor_Parcela,
                                                        Valor_Pago,
                                                        Quitar,
                                                        Complemento,
                                                        CodDevolucao,
                                                        Desconto,
                                                        Frete,
                                                        Entidade,
                                                        FormaPagto,
                                                        Cobranca,
                                                        Taxa,
                                                        Acrescimo) 
                          VALUES (@CODCOMPRA, 
                                  @DATAVENCTO, 
                                  @PARCELAS,
                                  @VALORPARC,
                                  @VALORPAGO,
                                  0,
                                  @COMPLEMENTO,
                                  @CodDevolucao,
                                  @DESCONTO,
                                  @FRETE,
                                  @ENTIDADE,
                                  @FORMAPAGTO,
                                  @COBRANCA,
                                  @TAXA,
                                  @ACRESCIMO)"

            ' Preparar os parâmetros, configurando valores nulos para CodCompra e CodDevolucao, se necessário
            Dim parameters As SqlParameter() = {
            New SqlParameter("@DATAVENCTO", DtVencimento),
            New SqlParameter("@PARCELAS", Parcelas),
            New SqlParameter("@VALORPARC", ValorParcela),
            New SqlParameter("@VALORPAGO", ValorPago),
            New SqlParameter("@COMPLEMENTO", Complemento),
            New SqlParameter("@DESCONTO", Desconto),
            New SqlParameter("@FRETE", Frete),
            New SqlParameter("@ENTIDADE", Entidade),
            New SqlParameter("@COBRANCA", Cobranca),
            New SqlParameter("@FORMAPAGTO", FormaPagto),
            New SqlParameter("@TAXA", Taxa),
            New SqlParameter("@ACRESCIMO", Acrescimo)
        }

            ' Adicionar parâmetros opcionais, verificando se são Nothing para definir como DBNull
            If CodCompra <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CODCOMPRA", CodCompra.Value)
            Else
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CODCOMPRA", DBNull.Value)
            End If

            If CodDevolucao <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodDevolucao", CodDevolucao.Value)
            Else
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodDevolucao", DBNull.Value)
            End If

            If CodRecebimento <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodRecebimento", CodRecebimento.Value)
            Else
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodRecebimento", DBNull.Value)
            End If
            ' Executar a operação
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Pagamento registrado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo atualiza o estado do pagamento para quitado.
        ''' </summary>
        ''' <param name="CodPagar">Representa o código identificador do pagamento.</param>
        ''' <param name="ValorPago">Representa o valor pago</param>
        Public Sub BaixarTituloAPagar(CodPagar As Integer, ValorPago As Decimal)
            Dim sql As String = "UPDATE Tbl_ContasAPagar SET Valor_Pago = @VALORPAGO, Dt_Pgto = GETDATE(), Quitar = '1' WHERE  Cod_ContasPagar = @CODPAGAR"
            Dim parameters As SqlParameter() = {
             New SqlParameter("@CODPAGAR", CodPagar),
            New SqlParameter("@VALORPAGO", ValorPago)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Baixa efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo atualiza as informações do pagamento.
        ''' </summary>
        ''' <param name="CodPagar">Representa o código do pagamento do tipo integer.</param>
        ''' <param name="DataVencto">Representa a data em que o pagamento irá vencer do tipo date.</param>
        ''' <param name="ValorPago">Representa o valor pago do pagamento do tipo decimal.</param>
        ''' <param name="DataPagto">Representa a data na qual o pagamento foi realizado do tipo date.</param>
        ''' <param name="Cobranca">Representa o código da  cobrança do pagamento do tipo integer.</param>
        ''' <param name="FormaPagto">Representa o código do da forma de pagamento do tipo integer.</param>
        ''' <param name="Complemento">Representa a descrição do pagamento do tipo string.</param>
        ''' <param name="Desconto">Representa o valor de desconto aplicado no pagamento do tipo decimal.</param>
        ''' <param name="Frete">Representa o valor de frete aplicado no pagamento do tipo decimal.</param>
        ''' <param name="Entidade">Representa o código da entidade para qual o pagamento foi realizado do tipo integer.</param>
        Public Sub AtualizarTituloAPagar(CodPagar As Integer, DataVencto As Date, ValorPago As Decimal, DataPagto As Date, Cobranca As Integer, FormaPagto As Integer, Complemento As String, Desconto As Integer, Frete As Integer, Entidade As Integer, Parcelas As Integer, ValorParcela As Decimal)
            Dim sql As String = "UPDATE   Tbl_ContasAPagar    SET     Dt_Vencimento = @DTVENCTO, 
                                                                Valor_Pago = @VALORPAGO,  
                                                                Dt_Pgto = @DTPAGTO, 
                                                                Complemento = @COMPLEMENTO, 
                                                                Cobranca = @COBRANCA, 
                                                                FormaPgto = @FORMAPAGTO, 
                                                                Desconto = @DESCONTO, 
                                                                Frete = @FRETE, 
                                                                Entidade = @ENTIDADE,
                                                                Parcelas = @PARCELAS,
                                                                Valor_Parcela = @VALORPPARCELA,
                                                        WHERE   Cod_ContasPagar = @CODPAGAR"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CODPAGAR", CodPagar),
                         New SqlParameter("@Entidade", Entidade),
                         New SqlParameter("@DTVENCTO", DataVencto),
                         New SqlParameter("@VALORPAGO", ValorPago),
                         New SqlParameter("@DTPAGTO", DataPagto),
                         New SqlParameter("@COBRANCA", Cobranca),
                         New SqlParameter("@COMPLEMENTO", Complemento),
                         New SqlParameter("@FORMAPAGTO", FormaPagto),
                         New SqlParameter("@DESCONTO", Desconto),
                         New SqlParameter("@FRETE", Frete),
                         New SqlParameter("@PARCELAS", Parcelas),
                         New SqlParameter("@VALORPPARCELA", ValorParcela)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Pagamento atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        ''' <summary>
        ''' Este metódo realiza a exclusão de um pagamento estando ou não vinculado a uma nota fiscal de entrada.
        ''' </summary>
        ''' <param name="CodCompra">Representa o código da nota fiscal de entrada do tipo integer.</param>
        ''' <param name="CodContPagar">Representa o código do pagamento do tipo integer.</param>
        Public Sub ExcluirPagamento(Optional CodContPagar As Integer? = 0, Optional CodCompra As Integer? = 0, Optional CodDevolucao As Integer? = 0)
            Dim sql As String
            If CodCompra <> 0 Then
                sql = "DELETE FROM Tbl_ContasAPagar WHERE CodCompra = @CODCOMPRA"
            ElseIf CodDevolucao <> 0 Then
                sql = "DELETE FROM Tbl_ContasAPagar WHERE CodDevolucao = @CodDevolucao"
            Else
                sql = "DELETE FROM Tbl_ContasApagar WHERE Cod_ContasPagar = @CodContasPagar"
            End If

            Dim parameters As SqlParameter() = {}

            If CodDevolucao <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodDevolucao", CodDevolucao.Value)
            End If

            If CodCompra <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CODCOMPRA", CodCompra.Value)
            End If

            If CodContPagar <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CODPAGAR", CodContPagar.Value)
            End If
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Pagamento excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
#End Region
#Region "FUNCOES"
        ''' <summary>
        ''' Este metódo consulta os dados relacionados a um pagamento conforme filtro informados e retorna num componente DataGridView.
        ''' </summary>
        ''' <param name="Status">Representa o estado do pagamento do tipo string.</param>
        ''' <param name="CodTitulo">Representa o código do pagamento do tipo integer.</param>
        ''' <param name="Fornecedor">Representa o nome da entidade para qual o pagamento deve ser realizado do tipo string.</param>
        ''' <param name="FormaPagto">Representa a forma de pagamento realizada do tipo string.</param>
        ''' <param name="Cobranca">Representa o form de cobranças do pagamento do tipo string.</param>
        Public Function PesquisaPagamentos(Status As String, CodTitulo As Integer, Fornecedor As String, FormaPagto As String, Cobranca As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_ContasPagar WHERE 1=1")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Em Aberto"
                        sql.AppendLine("AND Status LIKE 'EM ABERTO'")
                    Case "Baixado"
                        sql.AppendLine("AND Status LIKE 'BAIXADO'")
                    Case "Todos"
                        sql.AppendLine("AND Status IS NOT NULL")
                End Select

                'Pesquisa por codigo da venda
                If CodTitulo <> 0 Then
                    sql.AppendLine("AND CodTitulo = @CodTitulo")
                    parameters.Add(New SqlParameter("@CodTitulo", CodTitulo))
                End If

                'Pesquisa pelo Fornecedor
                If Not String.IsNullOrEmpty(Fornecedor) Then
                    sql.AppendLine("AND Entidade = @Fornecedor")
                    parameters.Add(New SqlParameter("@Fornecedor", Fornecedor))
                End If

                'Pesquisa pelo tipo do forma de pagamento
                If Not String.IsNullOrEmpty(FormaPagto) Then
                    sql.AppendLine("AND FormaPagto = @FormaPagto")
                    parameters.Add(New SqlParameter("@FormaPagto", FormaPagto))
                End If

                'Pesquisa pelo tipo do cobrança
                If Not String.IsNullOrEmpty(Cobranca) Then
                    sql.AppendLine("AND Cobranca = @Cobranca")
                    parameters.Add(New SqlParameter("@Cobranca", Cobranca))
                End If

                sql.AppendLine("ORDER BY DataCompra DESC")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function
        ''' <summary>
        ''' Este metódo consulta o último código de pagamento cadastrado e retorna o próximo número disponível.
        ''' </summary>
        ''' <param name="sql">query sql necessária para a consulta.</param>
        Public Function ConsultaPagamento(sql As String, Optional CodTitulo As Integer = 0)

            If CodTitulo <> 0 Then
                Dim parameters As SqlParameter() = {
             New SqlParameter("@CodTitulo", CodTitulo)
                }
                Return ClasseConexao.Consultar(sql, parameters)
            End If
            Return ClasseConexao.Consultar(sql, Nothing)
        End Function
        ''' <summary>
        ''' Este função verifica se existe um pagamento referente a uma compra.
        ''' </summary>
        ''' <param name="CodCompra">Codigo identificado da compra relacionada ao pagamento.</param>
        ''' <returns></returns>
        Public Function VerificaVinculoPagamento(CodCompra As Integer)
            Dim sql As String = "SELECT * FROM Tbl_ContasApagar WHERE CodCompra = @CodCompra"
            Dim parameters As SqlParameter() = {
             New SqlParameter("@CodCompra", CodCompra)
            }

            Return ClasseConexao.Consultar(sql, parameters)
        End Function
#End Region
    End Class
End Namespace

