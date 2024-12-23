Imports System.Data
Imports System.Text
Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports Xceed.Wpf.Toolkit
Namespace Classes.Financeiro
    ''' <summary>
    ''' Esta classe representa todas as rotina de contas a receber do sistema.Esta é uma classe derivada de clsFinanceiro.
    ''' </summary>
    Public Class clsContasReceber
        Inherits clsFinanceiro
        Dim ClasseConexao As New ConexaoSQLServer
#Region "METODOS"
        ''' <summary>
        ''' Este metódo realiza a gravação do recebimento no banco de dados.
        ''' </summary>
        ''' <param name="CodVenda">Representa o código do pedido de venda relacionada ao recebimento do tipo integer.</param>
        ''' <param name="DtVencimento">Representa a data em que o recebimento irá vencer do tipo date.</param>
        ''' <param name="Parcelas">Represente o número do parcelas do tipo integer.</param>
        ''' <param name="ValorParcela">Representa o valor da parcelas a ser pago do tipo decimal.</param>
        ''' <param name="ValorPago">Representa o valor pago do recebimento do tipo decimal.</param>
        ''' <param name="Complemento">Representa a descrição do recebimento do tipo string.</param>
        ''' <param name="Desconto">Representa o valor de desconto aplicado no recebimento do tipo decimal.</param>
        ''' <param name="Frete">Representa o valor de frete aplicado no recebimento do tipo decimal.</param>
        ''' <param name="Taxa">Representa o valor de taxa aplicado no recebimento do tipo decimal.</param>
        ''' <param name="Acrescimo">Representa o valor de acréscimo aplicado no recebimento do tipo decimal.</param>
        ''' <param name="Entidade">Representa o código da entidade para qual o recebimento foi realizado do tipo integer.</param>
        ''' <param name="Cobranca">Representa o código da  cobrança do recebimento do tipo integer.</param>
        ''' <param name="FormaPagto">Representa o código do da forma de pagamento do recebimento do tipo integer.</param>
        Public Sub GerarContasaReceber(DtVencimento As Date, Parcelas As Integer, ValorParcela As Decimal, ValorPago As Decimal, Complemento As String, Desconto As Decimal, Frete As Decimal, Taxa As Decimal, Acrescimo As Decimal, Entidade As Integer, Cobranca As Integer, FormaPagto As Integer, Optional CodVenda As Integer? = Nothing)
            Dim sql As String = "INSERT INTO Tbl_ContasAreceber       (Cod_TabVenda,
                                                                Dt_Vencimento,
                                                                Parcelas,
                                                                Valor_Parcela,
                                                                Valor_Pago,
                                                                Quitar,
                                                                Complemento,
                                                                Desconto,
                                                                Frete,
                                                                Entidade,
                                                                FormaPagto,
                                                                Cobranca,
                                                                Taxa,
                                                                Acrescimo) 
                                        VALUES                  (@CODVENDA, 
                                                                @DATAVENCTO, 
                                                                @PARCELAS,
                                                                @VALORPARC,
                                                                @VALORPAGO,
                                                                0,
                                                                @COMPLEMENTO,
                                                                @DESCONTO,
                                                                @FRETE,
                                                                @ENTIDADE,
                                                                @FORMAPAGTO,
                                                                @COBRANCA,
                                                                @TAXA,
                                                                @ACRESCIMO)"
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

            If CodVenda <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodVenda", CodVenda.Value)
            Else
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodVenda", DBNull.Value)
            End If

            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub

        ''' <summary>
        ''' Este metódo realiza a baixa de um recebimento.
        ''' </summary>
        ''' <param name="CodReceber">Representa o código identificador do recebimento.</param>
        ''' <param name="ValorPago">Representa o valor pago</param>
        Public Sub BaixarTituloAReceber(CodReceber As Integer, Optional ValorPago As Decimal = 0)
            If ValorPago <> 0 Then
                Dim sql As String = "UPDATE Tbl_ContasAreceber SET Valor_Pago = @VALORPAGO,  Dt_Pgto = GETDATE(), Quitar = '1' WHERE Cod_AutContaAreceber = @CODRECEBER"
                Dim parameters As SqlParameter() = {
            New SqlParameter("@CODRECEBER", CodReceber),
            New SqlParameter("@VALORPAGO", ValorPago)
        }
                ClasseConexao.Operar(sql, parameters)
            Else
                Dim sql As String = "UPDATE Tbl_ContasAreceber Dt_Pgto = GETDATE(), Quitar = '1' WHERE Cod_AutContaAreceber = @CODRECEBER"
                Dim parameters As SqlParameter() = {
            New SqlParameter("@CODRECEBER", CodReceber)
        }
                ClasseConexao.Operar(sql, parameters)
            End If
            MessageBox.Show("Recebimento baixado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo atualiza as informações do recebimento.
        ''' </summary>
        ''' <param name="CodReceber">Representa o código do recebimento do tipo integer.</param>
        ''' <param name="DataVencto">Representa a data em que o recebimento irá vencer do tipo date.</param>
        ''' <param name="ValorPago">Representa o valor pago do recebimento do tipo decimal.</param>
        ''' <param name="DataPagto">Representa a data na qual o recebimento foi realizado do tipo date.</param>
        ''' <param name="Cobranca">Representa o código da  cobrança do recebimento do tipo integer.</param>
        ''' <param name="FormaPagto">Representa o código do da forma de pagamento do recebimento do tipo integer.</param>
        ''' <param name="CodVenda">Representa o código do pedido de venda relacionada ao recebimento do tipo integer.</param>
        ''' <param name="Complemento">Representa a descrição do recebimento do tipo string.</param>
        ''' <param name="Desconto">Representa o valor de desconto aplicado no recebimento do tipo decimal.</param>
        ''' <param name="Frete">Representa o valor de frete aplicado no recebimento do tipo decimal.</param>
        ''' <param name="Entidade">Representa o código da entidade para qual o recebimento foi realizado do tipo integer.</param>
        Public Sub AtualizarTituloAReceber(CodReceber As Integer, DataVencto As Date, ValorPago As Decimal, DataPagto As Date, Cobranca As Integer, FormaPagto As Integer, CodVenda As Integer, Complemento As String, Desconto As Decimal, Frete As Decimal, Taxa As Decimal, Acrescimo As Decimal, Entidade As Integer)
            Dim sql As String = "UPDATE Tbl_ContasAreceber  SET Dt_Vencimento = @DTVENCTO, Valor_Pago = @VALORPAGO, Dt_Pgto = @DTPAGTO, Complemento = @COMPLEMENTO WHERE Cod_AutContaAreceber = @CODRECEBER"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CODRECEBER", CodReceber),
                         New SqlParameter("@DTVENCTO", DataVencto),
                         New SqlParameter("@VALORPAGO", ValorPago),
                         New SqlParameter("@DTPAGTO", DataPagto),
                         New SqlParameter("@COMPLEMENTO", Complemento),
                         New SqlParameter("@FORMAPAGTO", FormaPagto),
                         New SqlParameter("@DESCONTO", Desconto),
                         New SqlParameter("@FRETE", Frete),
                         New SqlParameter("@ENTIDADE", Entidade),
                         New SqlParameter("@TAXA", Taxa),
                         New SqlParameter("@ACRESCIMO", Acrescimo)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Recebimento atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo consultar se o título possui uma nota fiscal de entrada, e caso não existe realiza a exclusão do pagamento no banco de dados.
        ''' </summary>
        ''' <param name="CodVenda">Representa o código do pedido de venda do tipo integer </param>
        ''' <param name="CodContReceber">Representa o código do recebimento do tipo integer</param>
        Public Sub ExcluirRecebimento(Optional CodVenda As Integer? = 0, Optional CodContReceber As Integer? = 0)
            Dim sql As String
            If CodVenda <> 0 Then
                sql = "DELETE FROM Tbl_ContasAreceber WHERE Cod_TabVenda = @CODVENDA"
            Else
                sql = "DELETE FROM Tbl_ContasAreceber WHERE Cod_AutContaAreceber = @CodContReceber"
            End If

            Dim parameters As SqlParameter() = {}

            If CodVenda <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CODVENDA", CodVenda.Value)
            End If

            If CodContReceber <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodContReceber", CodContReceber.Value)
            End If
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Recebimento excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub
#End Region
#Region "FUNCOES"
        ''' <summary>
        ''' Esta função consulta os dados relacionados a um recebimento conforme filtro informados e retorna num componente DataGridView.
        ''' </summary>
        ''' <param name="Status">Representa o estado do recebimento do tipo string.</param>
        ''' <param name="CodTitulo">Representa o código do recebimento do tipo integer.</param>
        ''' <param name="Cliente">Representa o nome da entidade para qual o recebimento deve ser realizado do tipo string.</param>
        ''' <param name="FormaPagto">Representa a forma de pagamento do recebimento realizada do tipo string.</param>
        ''' <param name="Cobranca">Representa o form de cobranças do recebimento do tipo string.</param>
        Public Function PesquisaRecebimentos(Status As String, CodTitulo As String, Cliente As String, FormaPagto As String, Cobranca As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_ContasReceber WHERE 1=1")
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
                If Not String.IsNullOrEmpty(Cliente) Then
                    sql.AppendLine("AND Cliente = @Cliente")
                    parameters.Add(New SqlParameter("@Cliente", Cliente))
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

                sql.AppendLine("ORDER BY DataVenda DESC")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())
            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Este metódo consulta o último código de recebimento cadastrado e retorna o próximo número disponível.
        ''' </summary>
        ''' <param name="sql">query sql necessária para a consulta.</param>
        Public Function ConsultaRecebimento(sql As String, Optional CodTitulo As Integer = 0)

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
        ''' <param name="CodVenda">Codigo identificado da venda relacionada ao pagamento.</param>
        ''' <returns></returns>
        Public Function VerificaVinculoRecebimento(CodVenda As Integer)
            Dim sql As String = "SELECT * FROM Tbl_ContasAreceber WHERE Cod_TabVenda = @CODVENDA"
            Dim parameters As SqlParameter() = {
             New SqlParameter("@CodVenda", CodVenda)
            }
            Return ClasseConexao.Consultar(sql, parameters)
        End Function
#End Region
    End Class

End Namespace
