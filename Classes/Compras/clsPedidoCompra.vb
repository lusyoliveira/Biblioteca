Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports System.Data
Imports System.Text
Imports Xceed.Wpf.Toolkit

Namespace Classes.Compras
    Public Class clsPedidoCompra
        Inherits clsCompras
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property CodRastreamento As String
        Public Property EstoqueAnterior As Integer
        Public Property EstoqueAtual As Integer
        Public Property Obs As String
        Public Property Motivo As String
        Public Property DataCompra As Date
        Public Property DataMovimento As Date
        Public Property Transportadora As String
        Public Property CodComprador As Integer
        Public Property Comprador As String
        Public Property TipoInventario As String
        Public Property CodTipoInventario As Integer
        Public Property DataEmissao As Date
        Public Property DataEntrada As Date
        Public Property Serie As Integer
        Public Property ChaveAcesso As String
        Public Property CustoAtual As Decimal
        Public Property CNPJ As String
        Public Property Telefone As String
        Public Property Endereco As String
        Public Property Email As String
#End Region
#Region "METODOS"
        Public Sub SalvaPedidoCompra(DataPedCompra As Date, Fornecedor As Integer, TotalPago As Decimal, FormaPgto As Integer, Transporte As Decimal, Desconto As Decimal, Obs As String, PrazoEntrega As Integer, CodigoRastreio As String, Transportadora As Integer, Comprador As Integer, Optional CodCotacao As Integer? = Nothing)
            Dim sql As String = "INSERT  INTO Tbl_PedCompra   (DataPedCompra, 
                                                        Fornecedor, 
                                                        TotalPago, 
                                                        FormaPgto, 
                                                        Transporte, 
                                                        Desconto, 
                                                        Obs, 
                                                        PrazoEntrega, 
                                                        Status,
                                                        CodigoRastreio,
                                                        Transportadora,
                                                        Comprador,
                                                        DataCadastro,
                                                        CodCotacao) 
                                             VALUES     (@DataPedCompra, 
                                                        @Fornecedor, 
                                                        @TotalPago, 
                                                        @FormaPgto, 
                                                        @Transporte, 
                                                        @Desconto, 
                                                        @Obs, 
                                                        @PrazoEntrega, 
                                                        0,
                                                        @CodigoRastreio,
                                                        @Transportadora,
                                                        @Comprador,
                                                        GETDATE(),
                                                        @CodCotacao)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@DataPedCompra", DataPedCompra),
                         New SqlParameter("@Fornecedor", Fornecedor),
                         New SqlParameter("@TotalPago", TotalPago),
                         New SqlParameter("@FormaPgto", FormaPgto),
                         New SqlParameter("@Transporte", Transporte),
                         New SqlParameter("@Desconto", Desconto),
                         New SqlParameter("@Obs", Obs),
                         New SqlParameter("@PrazoEntrega", PrazoEntrega),
                         New SqlParameter("@CodigoRastreio", CodigoRastreio),
                         New SqlParameter("@Transportadora", Transportadora),
                         New SqlParameter("@Comprador", Comprador)
                         }
            If CodCotacao <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodCotacao", CodCotacao.Value)
            Else
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodCotacao", DBNull.Value)
            End If
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub SalvaDetPedCompra(CodigoPedCompra As Integer, ValorUnit As Decimal, Quantidade As Integer, CodSimples As Integer, ValorTotal As Decimal)
            Dim sql As String = "INSERT INTO Tbl_PedCompraDet (CodigoPedCompra,
                                                            ValorUnit,
                                                            Quantidade,
                                                            Cod_Simples,
                                                            ValorTotal) 
                                                    VALUES(@CodigoPedCompra,
                                                            @ValorUnit,
                                                            @Quantidade,
                                                            @Cod_Simples,
                                                            @ValorTotal)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodigoPedCompra", CodigoPedCompra),
                         New SqlParameter("@ValorUnit", ValorUnit),
                         New SqlParameter("@Quantidade", Quantidade),
                         New SqlParameter("@Cod_Simples", CodSimples),
                         New SqlParameter("@ValorTotal", ValorTotal)
                         }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub ExcluirDetPedidoCompra(CodigoPedCompra As Integer)
            Dim sql As String = "DELETE FROM Tbl_PedCompraDet WHERE CodigoPedCompra = @CodigoPedCompra"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodigoPedCompra", CodigoPedCompra)
                         }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub ExcluirPedidoCompra(CodPedCompra As Integer)
            Dim sql As String = "DELETE FROM Tbl_PedCompra WHERE CodPedCompra = @CodPedCompra"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodPedCompra", CodPedCompra)
                         }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluirItemPedidoCompra(CodigoPedCompra As Integer, CodSimples As Integer)
            Dim parameters As SqlParameter() = {
                        New SqlParameter("@CodigoPedCompra", CodigoPedCompra),
                        New SqlParameter("@CodSimples", CodSimples)
            }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_PedidosCompraDetalhes WHERE CodPedCompra = @CodigoPedCompra AND Cod_Simples = @CodSimples", parameters)

            If Tabela.Rows.Count > 0 Then
                Dim parametersDel As SqlParameter() = {
                        New SqlParameter("@CodigoPedCompra", CodigoPedCompra),
                        New SqlParameter("@CodSimples", CodSimples)
            }
                Dim sql As String = "DELETE FROM Tbl_PedCompraDet WHERE CodigoPedCompra = @CodigoPedCompra AND Cod_Simples = @CodSimples"
                ClasseConexao.Operar(sql, parametersDel)
            Else
                Exit Sub
            End If
        End Sub
        Public Sub AtualizaPedidoCompra(CodPedCompra As Integer, DataPedCompra As Date, Fornecedor As Integer, TotalPago As Decimal, FormaPgto As Integer, Transporte As Decimal, Desconto As Decimal, Obs As String, PrazoEntrega As Integer, CodigoRastreio As String, Transportadora As Integer, Comprador As Integer)
            Dim sql As String = "UPDATE Tbl_PedCompra  SET    DataPedCompra = @DataPedCompra,
                                                        Fornecedor = @Fornecedor,
                                                        TotalPago = @TotalPago,
                                                        FormaPgto = @FormaPgto,
                                                        Transporte = @Transporte,
                                                        Desconto = @Desconto,
                                                        Obs = @Obs,
                                                        PrazoEntrega =  @PrazoEntrega,
                                                        CodigoRastreio = @CodigoRastreio,
                                                        Transportadora = @Transportadora,
                                                        Comprador = @Comprador
                                                WHERE   CodPedCompra = @CodPedCompra"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodPedCompra", CodPedCompra),
                         New SqlParameter("@DataPedCompra", DataPedCompra),
                         New SqlParameter("@Fornecedor", Fornecedor),
                         New SqlParameter("@TotalPago", TotalPago),
                         New SqlParameter("@FormaPgto", FormaPgto),
                         New SqlParameter("@Transporte", Transporte),
                         New SqlParameter("@Desconto", Desconto),
                         New SqlParameter("@Obs", Obs),
                         New SqlParameter("@PrazoEntrega", PrazoEntrega),
                         New SqlParameter("@CodigoRastreio", CodigoRastreio),
                         New SqlParameter("@Transportadora", Transportadora),
                         New SqlParameter("@Comprador", Comprador)
                         }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaDetPedidoCompra(CodigoPedCompra As Integer, ValorUnit As Decimal, Quantidade As Integer, CodSimples As Integer, ValorTotal As Decimal)
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodigoPedCompra", CodigoPedCompra),
                         New SqlParameter("@CodSimples", CodSimples)
        }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_PedidosCompraDetalhes WHERE CodPedCompra = @CodigoPedCompra AND Cod_Simples = @CodSimples", parameters)
            If Tabela.Rows.Count > 0 Then
                Dim parametersDel As SqlParameter() = {
                         New SqlParameter("@CodigoPedCompra", CodigoPedCompra),
                         New SqlParameter("@CodSimples", CodSimples),
                         New SqlParameter("@Quantidade", Quantidade),
                         New SqlParameter("@ValorUnit", ValorUnit),
                         New SqlParameter("@ValorTotal", ValorTotal)
        }

                Dim sql As String = "UPDATE Tbl_PedCompraDet SET      
                                                            ValorUnit = @ValorUnit,
                                                            Quantidade = @Quantidade,
                                                            ValorTotal = @ValorTotal
                                                    WHERE CodigoPedCompra = @CodigoPedCompra
                                                    AND Cod_Simples = @CodSimples"

                ClasseConexao.Operar(sql, parametersDel)
            Else
                SalvaDetPedCompra(CodigoPedCompra, ValorUnit, Quantidade, CodSimples, ValorTotal)
            End If

        End Sub
        Public Sub AtualizaStatusPedCompra(Status As Integer, CodPedCompra As Integer)
            Dim sql As String = "UPDATE Tbl_PedCompra SET Status = @Status WHERE CodPedCompra = @CodPedCompra"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodPedCompra", CodPedCompra),
                         New SqlParameter("@Status", Status)
                         }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Pedido atualizado com sucesso!", "Sucesso")
        End Sub
#End Region
#Region "FUNCOES"
        Public Function PesquisaPedidoCompra(StatusPed As String, CodPedido As String, Fornecedor As String, FormaPagto As String, Comprador As String, Transportadora As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_PedidosCompra WHERE 1=1")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case StatusPed
                    Case "Em Aberto"
                        sql.AppendLine("AND StatusPed = 0")
                    Case "Concluído"
                        sql.AppendLine("AND StatusPed = 1")
                    Case "Finalizado"
                        sql.AppendLine("AND StatusPed = 2")
                    Case "Todos"
                        sql.AppendLine("AND StatusPed IS NOT NULL")
                End Select

                If Not String.IsNullOrEmpty(CodPedido) Then
                    sql.AppendLine("AND CodPedido = @CodPedido")
                    parameters.Add(New SqlParameter("@CodPedido", CodPedido))
                End If

                If Not String.IsNullOrEmpty(Fornecedor) Then
                    sql.AppendLine("AND Fornecedor LIKE @Fornecedor")
                    parameters.Add(New SqlParameter("@Fornecedor", Fornecedor))
                End If

                If Not String.IsNullOrEmpty(FormaPagto) Then
                    sql.AppendLine("AND FormaPagto LIKE @FormaPagto")
                    parameters.Add(New SqlParameter("@FormaPgto", FormaPagto))
                End If

                If Not String.IsNullOrEmpty(Comprador) Then
                    sql.AppendLine("AND Comprador LIKE @Comprador")
                    parameters.Add(New SqlParameter("@Comprador", Comprador))
                End If

                If Not String.IsNullOrEmpty(Transportadora) Then
                    sql.AppendLine("AND Transportadora LIKE @Transportadora")
                    parameters.Add(New SqlParameter("@Transportadora", Transportadora))
                End If

                sql.AppendLine("ORDER BY DataCompra desc")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Esta função consulta um pedido de compra no banco de dados.
        ''' </summary>
        ''' <param name="CodPedCompra">Representa o número do pedido.</param>
        ''' <param name="sql">Representa a query sql necessária para realizar a consulta.</param>
        ''' <returns>Retorna dados do pedido de compra.</returns>
        Public Function ConsultaPedidoCompra(sql As String, Optional CodPedCompra As Integer = 0)
            If CodPedCompra <> 0 Then
                Dim parameters As SqlParameter() = {
                New SqlParameter("@CodPedCompra", CodPedCompra)
            }

                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function

#End Region
    End Class
End Namespace

