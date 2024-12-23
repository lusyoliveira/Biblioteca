Imports System.Data
Imports System.Text
Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Namespace Classes.Comercial
    Public Class clsPedidoVenda

        Inherits clsComercial
        Dim ClasseConexao As New ConexaoSQLServer

#Region "METODOS"
        Public Sub SalvarPedidoVenda(DataVenda As Date, Cliente As Integer, FormaPgto As Integer, Transporte As Decimal, Desconto As Decimal, Transportadora As Integer, ObsVenda As String, Vendedor As Integer, Cobranca As Integer, NotaFiscal As Integer, Loja As Integer, taxas As Decimal, acrescimo As Decimal, TotalProdutos As Decimal, PesoTotal As Decimal, Lucro As Decimal, TotalPedido As Decimal, TotalARceber As Decimal, Operacao As Integer, Optional CodOrcamento As Integer? = Nothing)
            Dim sql = "INSERT INTO tbl_Vendas   (DataVenda, 
                                                    Cliente, 
                                                    FormaPgto, 
                                                    Transporte,  
                                                    Desconto, 
                                                    Transportadora, 
                                                    Obs_Venda, 
                                                    Vendedor, 
                                                    Cobranca, 
                                                    NotaFiscal,
                                                    CodOrcamento, 
                                                    Loja, 
                                                    taxas, 
                                                    acrescimo,
                                                    Status,
                                                    TotalProdutos,
                                                    PesoTotal,
                                                    Lucro,
                                                    TotalPedido,
                                                    TotalARceber,
                                                    DataCriacao,
                                                    Operacao) 
                                            VALUES  (@DataVenda, 
                                                    @Cliente, 
                                                    @FormaPgto, 
                                                    @Transporte,  
                                                    @Desconto, 
                                                    @Transportadora, 
                                                    @ObsVenda, 
                                                    @Vendedor, 
                                                    @Cobranca, 
                                                    @NotaFiscal,
                                                    @CodOrcamento, 
                                                    @Loja, 
                                                    @taxas, 
                                                    @acrescimo,
                                                    0,
                                                    @TotalProdutos,
                                                    @PesoTotal,
                                                    @Lucro,
                                                    @Total,
                                                    @TotalARceber,
                                                    GETDATE(),
                                                    @Operacao)"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@DataVenda", DataVenda),
                     New SqlParameter("@Cliente", Cliente),
                     New SqlParameter("@FormaPgto", FormaPgto),
                     New SqlParameter("@Transporte", Transporte),
                     New SqlParameter("@Desconto", Desconto),
                     New SqlParameter("@Transportadora", Transportadora),
                     New SqlParameter("@ObsVenda", ObsVenda),
                     New SqlParameter("@Vendedor", Vendedor),
                     New SqlParameter("@Cobranca", Cobranca),
                     New SqlParameter("@NotaFiscal", NotaFiscal),
                     New SqlParameter("@Loja", Loja),
                     New SqlParameter("@taxas", taxas),
                     New SqlParameter("@acrescimo", acrescimo),
                     New SqlParameter("@TotalProdutos", TotalProdutos),
                     New SqlParameter("@PesoTotal", PesoTotal),
                     New SqlParameter("@Lucro", Lucro),
                     New SqlParameter("@Total", TotalPedido),
                     New SqlParameter("@TotalARceber", TotalARceber),
                     New SqlParameter("@Operacao", Operacao)
                     }
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

        Public Sub SalvarDetPedidoVenda(CodPedido As Integer, Quantidade As Integer, ValorUnit As Decimal, CodItem As Integer, ValorTotal As Decimal)
            Dim sql As String = "INSERT INTO Tbl_VendasDet (CodigoVendas,ValorUnit,Quantidade,Cod_Simples,ValorTotal) 
                                                VALUES    (@CODVENDA,@VUNITARIO,@QUANTIDADE,@CODIGO,@TOTAL)"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CODVENDA", CodPedido),
                     New SqlParameter("@QUANTIDADE", Quantidade),
                     New SqlParameter("@VUNITARIO", ValorUnit),
                     New SqlParameter("@CODIGO", CodItem),
                     New SqlParameter("@TOTAL", ValorTotal)
        }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub AtualizaStatusPedidoVenda(CodVenda As Integer, Status As Integer)
            Dim sql As String = "UPDATE Tbl_Vendas SET Status = @Status WHERE CodVenda = @CODVENDA"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@Status", Status),
                     New SqlParameter("@CODVENDA", CodVenda)
                     }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Pedido atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaPedidoVenda(CodVenda As Integer, Transporte As Decimal, Desconto As Decimal, CodigoRastreamento As String, ObsVenda As String, NotaFiscal As Integer, taxas As Decimal, acrescimo As Decimal, FormaPagto As Integer, Cobranca As Integer, TotalProdutos As Decimal, PesoTotal As Decimal, Lucro As Decimal, TotalPedido As Decimal, TotalARceber As Decimal, Cliente As Integer, Operacao As Integer)
            Dim sql As String = "UPDATE Tbl_Vendas SET        Desconto = @Desconto,
		                        	                    Transporte = @Transporte,
	                        		                    CodigoRastreamento = @CodigoRastreamento,
	                           		                    NotaFiscal = @NotaFiscal,
	                        		                    Obs_Venda = @ObsVenda,
	                        		                    taxas = @taxas,
	                           		                    acrescimo = @acrescimo,
                                                        Cobranca = @Cobranca,
                                                        FormaPgto = @FormaPagto, 
                                                        TotalProdutos = @TotalProdutos,
                                                        PesoTotal = @PesoTotal,
                                                        Lucro = @Lucro,
                                                        TotalPedido = @Total,
                                                        TotalARceber = @TotalARceber,
                                                        Cliente = @Cliente,
                                                        Operacao = @Operacao
                                            WHERE       CodVenda = @CodPedido"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodPedido", CodVenda),
                     New SqlParameter("@Transporte", Transporte),
                     New SqlParameter("@Desconto", Desconto),
                     New SqlParameter("@ObsVenda", ObsVenda),
                     New SqlParameter("@NotaFiscal", NotaFiscal),
                     New SqlParameter("@CodigoRastreamento", CodigoRastreamento),
                     New SqlParameter("@taxas", taxas),
                     New SqlParameter("@acrescimo", acrescimo),
                     New SqlParameter("@TotalProdutos", TotalProdutos),
                     New SqlParameter("@PesoTotal", PesoTotal),
                     New SqlParameter("@Lucro", Lucro),
                     New SqlParameter("@FormaPagto", FormaPagto),
                     New SqlParameter("@Cobranca", Cobranca),
                     New SqlParameter("@Total", TotalPedido),
                     New SqlParameter("@TotalARceber", TotalARceber),
                     New SqlParameter("@Cliente", Cliente),
                     New SqlParameter("@Operacao", Operacao)
                     }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Pedido atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaDetPedidoVenda(CodPedido As Integer, Quantidade As Integer, ValorUnit As Decimal, CodItem As Integer, ValorTotal As Decimal)
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodigoVendas", CodPedido),
                     New SqlParameter("@CodSimples", CodItem)
    }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_VendasDetalhes WHERE CodPedido = @CodigoVendas AND CodSimples = @CodSimples", parameters)
            If Tabela.Rows.Count > 0 Then
                Dim parametersUp As SqlParameter() = {
                     New SqlParameter("@CodigoVendas", CodPedido),
                     New SqlParameter("@CodSimples", CodItem),
                     New SqlParameter("@Quantidade", Quantidade),
                     New SqlParameter("@ValorUnit", ValorUnit),
                     New SqlParameter("@ValorTotal", ValorTotal)
    }


                Dim sql As String = "UPDATE Tbl_VendasDet SET      
                                                    ValorUnit = @ValorUnit,
                                                    Quantidade = @Quantidade,
                                                    ValorTotal = @ValorTotal
                                            WHERE CodigoVendas = @CodigoVendas
                                            AND Cod_Simples = @CodSimples"

                ClasseConexao.Operar(sql, parametersUp)
            Else
                SalvarDetPedidoVenda(CodPedido, Quantidade, ValorUnit, CodItem, ValorTotal)
            End If
        End Sub
        Public Sub ExcluirPedidoVenda(CodVenda As Integer)
            Dim sql As String = "DELETE FROM Tbl_Vendas WHERE CodVenda = @CodPedido"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodPedido", CodVenda)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluirItemPedidoVenda(CodVenda As Integer, Optional CodSimples As Integer = 0)
            If CodSimples <> 0 Then
                Dim sql As String = "DELETE FROM Tbl_VendasDet WHERE CodigoVendas = @CodPedido AND Cod_Simples = @CodSimples"
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodPedido", CodVenda),
                         New SqlParameter("@CodSimples", CodSimples)
            }
                ClasseConexao.Operar(sql, parameters)
            Else
                Dim sql As String = "DELETE FROM Tbl_VendasDet WHERE CodigoVendas = @CodPedido"
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodPedido", CodVenda)
            }
                ClasseConexao.Operar(sql, parameters)
            End If
            MessageBox.Show("Item excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaEntrega(CodPedido As Integer)
            Dim sql As String = "UPDATE Tbl_Vendas SET DtEntrega = GETDATE() WHERE CodVenda = @CODVENDA"
            Dim parameters As SqlParameter() = {
                    New SqlParameter("@CODVENDA", CodPedido)
                                         }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Pedido atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaRecVenda(Cobranca As Integer, FormaPagto As Integer, CodVenda As Integer, Desconto As Decimal, Frete As Decimal, Taxa As Decimal, Acrescimo As Decimal, Entidade As Integer)
            Dim sql As String = "UPDATE Tbl_Vendas  SET Cobranca = @COBRANCA, FormaPgto = @FORMAPAGTO WHERE CodPedido = @CODVENDA"
            Dim parameters As SqlParameter() = {
          New SqlParameter("@CODVENDA", CodVenda),
          New SqlParameter("@COBRANCA", Cobranca),
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
#End Region
#Region "FUNCOES"
        Public Function PesquisaPedidoVenda(Status As String, CodPedido As String, Cliente As String, FormPagto As String, Cobranca As String, Transportadora As String, Loja As String, Vendedor As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_Vendas WHERE 1=1 ")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Em Aberto"
                        sql.AppendLine("AND Status LIKE 'EM ABERTO'")
                    Case "Concluído"
                        sql.AppendLine("AND Status LIKE 'CONCLUÍDO'")
                    Case "Devolver"
                        sql.AppendLine("AND Status LIKE 'DEVOLVER'")
                    Case "Devolvido"
                        sql.AppendLine("AND Status LIKE 'DEVOLVIDO'")
                    Case "Cancelado"
                        sql.AppendLine("AND Status LIKE 'CANCELADO'")
                    Case "Finalizado"
                        sql.AppendLine("AND Status LIKE 'FINALIZADO'")
                    Case "Todos"
                        sql.AppendLine("AND Status IS NOT NULL")
                End Select

                If Not String.IsNullOrEmpty(CodPedido) Then
                    sql.AppendLine("AND CodPedido = @CodPedido")
                    parameters.Add(New SqlParameter("@CodPedido", CodPedido))
                End If

                If Not String.IsNullOrEmpty(Cliente) Then
                    sql.AppendLine("AND Entidade LIKE @Cliente")
                    parameters.Add(New SqlParameter("@Cliente", Cliente))
                End If

                If Not String.IsNullOrEmpty(FormPagto) Then
                    sql.AppendLine("AND Forma_Pgto LIKE @FormaPagto")
                    parameters.Add(New SqlParameter("@FormaPgto", FormPagto))
                End If

                If Not String.IsNullOrEmpty(Cobranca) Then
                    sql.AppendLine("AND Cobranca LIKE @Cobranca")
                    parameters.Add(New SqlParameter("@Cobranca", Cobranca))
                End If

                If Not String.IsNullOrEmpty(Transportadora) Then
                    sql.AppendLine("AND Transportadora LIKE @Transportadora")
                    parameters.Add(New SqlParameter("@Transportadora", Transportadora))
                End If

                If Not String.IsNullOrEmpty(Loja) Then
                    sql.AppendLine("AND Loja LIKE @Origem")
                    parameters.Add(New SqlParameter("@Origem", Loja))
                End If

                If Not String.IsNullOrEmpty(Vendedor) Then
                    sql.AppendLine("AND Vendedor LIKE @Vendedor")
                    parameters.Add(New SqlParameter("@Vendedor", Vendedor))
                End If

                sql.AppendLine("ORDER BY CodPedido DESC")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function

#End Region
    End Class
End Namespace
