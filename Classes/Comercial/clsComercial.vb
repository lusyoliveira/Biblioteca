Imports System.Data
Imports System.Text
Imports Biblioteca.Classes.Conexao
Imports Biblioteca.Classes.Produtos
Imports Microsoft.Data.SqlClient
Imports Xceed.Wpf.Toolkit
Namespace Classes.Comercial
    ''' <summary>
    ''' Este classe representa todas as rotinas em comum a transações comercial do sistema. Esta é uma classe Base.
    ''' </summary>
    Public Class clsComercial
        Inherits clsItem
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property CodDevolucao As Integer
        Public Property CodOrcamento As Integer
        Public Property CodPedido As Integer
        Public Property CodCliente As Integer
        Public Property Entidade As String
        Public Property CPF As String
        Public Property Telefone As String
        Public Property Endereco As String
        Public Property Email As String
        Public Property Loja As String
        Public Property ValorUnit As Decimal
        Public Property ValorTotal As Decimal
        Public Property Total As Decimal
        Public Property Status As String
        Public Property FormaPagto As String
        Public Property Cobranca As String
        Public Property Frete As Decimal
        Public Property Desconto As Decimal
        Public Property Acrescimo As Decimal
        Public Property Taxa As Decimal
        Public Property CodEstampa As Integer
        Public Property Estampa As String
        Public Property CodElastico As Integer
        Public Property Elastico As String
        Public Property CodMiolo As Integer
        Public Property Miolo As String
        Public Property DataPrimeiraVenda As DateTime
        Public Property DataUltimaVenda As DateTime
        Public Property QtdCompra As Integer
        Public Property CAMINHOARQUIVO As String
        Public Property Dias As Integer
        Public Property Frequencia As Integer
        Public Property Periodo As Double
        Public Property DataDespache As DateTime
        Public Property DataEntrega As DateTime
        Public Property DataVenda As DateTime
        Public Property CustoTotal As Decimal
        Public Property Margem As Decimal
        Public Property VendaTotal As Double
        Public Property ValorUnitMedio As Double
        Public Property VendasMes As Decimal
        Public Property Lucro As Double
        Public Property LucroMes As Decimal
        Public Property LucroAno As Decimal
        Public Property Ano As Integer
        Public Property Mes As Integer
        Public Property Transportadora As String
        Public Property Vendedor As String
        Public Property CodRastreamento As String
        Public Property CodVendedor As Integer
        Public Property CodMotivo As Integer
        Public Property DataDevolucao As DateTime
        Public Property Motivo As String
        Public Property TotalARceber As Decimal
        Public Property TotalPedido As Decimal
        Private _DataOrcamento As DateTime
        Public Property DataOrcamento As DateTime
            Get
                Return _DataOrcamento
            End Get
            Set(value As DateTime)
                _DataOrcamento = value
            End Set
        End Property
        Private Property _Produto As String
        Public Property Produto As String
            Get
                Return _Produto
            End Get
            Set(value As String)
                _Produto = value
            End Set
        End Property
#End Region
#Region "CONSTRUTORES"
        Public Sub New()

        End Sub
        Public Sub New(_codigo As Integer, _item As String, _quantidade As Integer, _valorunit As Decimal)
            CodItem = _codigo
            Item = _item
            Quantidade = _quantidade
            ValorUnit = _valorunit
        End Sub
#End Region
#Region "METODOS"
        Public Sub SalvaVendedor(Vendedor As String)

            Dim sql As String = "INSERT INTO Tbl_Vendedores (vendedor) VALUES (@vendedor)"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@vendedor", Vendedor)
        }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub AtualizaVendedor(Codigo As Integer, Vendedor As String)
            Dim sql As String = "UPDATE Tbl_Vendedores SET vendedor = @Vendedor WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@vendedor", Vendedor),
            New SqlParameter("@Codigo", Codigo)
        }
            ClasseConexao.Operar(sql, parameters)

        End Sub
        Public Sub ExcluiVendedor(Codigo As Integer)
            Dim sql As String = "DELETE FROM Tbl_Vendedores Where Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@Codigo", Codigo)
        }
            ClasseConexao.Operar(sql, parameters)
        End Sub

#End Region
#Region "FUNCOES"
        Public Function ConsultaPedidoVenda(sql As String, Optional CodVenda As Integer = 0)
            If CodVenda <> 0 Then
                Dim parameters As SqlParameter() = {
            New SqlParameter("@CodVenda", CodVenda)
                }

                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
        Public Function ConsultaVendedor(sql As String)

            Return ClasseConexao.Consultar(sql, Nothing)
        End Function
        Public Function ConsultaFaturamento(sql As String, Optional Ano As Integer = 0, Optional Mes As Integer = 0)

            If Ano <> 0 And Mes <> 0 Then
                Dim parameters As SqlParameter() = {
            New SqlParameter("@Ano", Ano),
            New SqlParameter("@Mes", Mes)
                }

                Return ClasseConexao.Consultar(sql, parameters)
            ElseIf Ano <> 0 And Mes = 0 Then
                Dim parameters As SqlParameter() = {
            New SqlParameter("@Ano", Ano)
                }

                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
        Public Function MelhoresClientes(Grafico As Chart) As List(Of clsComercial)
            Dim resultado As New List(Of clsComercial)()
            Dim Clientes As New List(Of String)()
            Dim Quantidades As New List(Of Integer)()
            Dim dt As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_MelhoresClientes WHERE Compras > 1 ORDER BY Compras DESC", Nothing)

            Try
                For Each row As DataRow In dt.Rows
                    Dim Cliente As String = If(Not row.IsNull("NomeCliente"), row("NomeCliente").ToString(), String.Empty)
                    Dim Quantidade As Integer = If(Not row.IsNull("Compras"), Convert.ToInt32(row("Compras")), 0)

                    Clientes.Add(Cliente)
                    Quantidades.Add(Quantidade)

                    ' Adiciona ao resultado
                    resultado.Add(New clsComercial() With {.Entidade = Cliente, .Quantidade = Quantidade})
                Next

                ' Preencher o gráfico com os dados coletados
                Grafico.Series(0).Points.DataBindXY(Clientes, Quantidades)

            Catch ex As Exception
                MessageBox.Show("Não foi possível gerar o gráfico!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try

            Return resultado

        End Function
        Public Function RankingProdutos(Grafico As Chart, Ano As Integer) As List(Of clsComercial)
            Dim resultado As New List(Of clsComercial)()
            Dim ListaItens As New List(Of String)()
            Dim ListaQuantidades As New List(Of Integer)()
            Dim dt As DataTable = ConsultaFaturamento("SELECT * FROM Cs_RankingProdutos WHERE ANO = @ano ORDER BY Quantidade DESC", Ano)

            Try
                For Each row As DataRow In dt.Rows
                    Dim Item As String = If(Not row.IsNull("Descricao"), row("Descricao").ToString, String.Empty)
                    Dim Quantidade As Integer = If(Not row.IsNull("Quantidade"), Convert.ToInt32(row("Quantidade")), 0)

                    ListaItens.Add(Item)
                    ListaQuantidades.Add(Quantidade)


                    ' Também pode adicionar ao resultado se necessário:
                    resultado.Add(New clsComercial() With {.Item = Item, .Quantidade = Quantidade})
                Next
                ' Preencher o gráfico com os dados coletados
                Grafico.Series(0).Points.DataBindXY(ListaItens, ListaQuantidades)

            Catch ex As Exception
                MessageBox.Show("Não foi possível gerar o gráfico!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
            Return resultado
        End Function

        Public Function FaturamentoAnual(Grafico As Chart) As List(Of clsComercial)
            Dim resultado As New List(Of clsComercial)()
            Dim anos As New List(Of Integer)()
            Dim valores As New List(Of Decimal)()
            Dim dt As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_FaturamentoAnual ORDER BY Ano DESC", Nothing)

            Try
                For Each row As DataRow In dt.Rows
                    Dim ano As Integer = If(Not row.IsNull("Ano"), Convert.ToInt32(row("Ano")), 0)
                    Dim valorTotal As Decimal = If(Not row.IsNull("TOTAL"), Convert.ToInt32(row("TOTAL")), 0)

                    anos.Add(ano)
                    valores.Add(valorTotal)

                    ' Também pode adicionar ao resultado se necessário:
                    resultado.Add(New clsComercial() With {.Ano = ano, .ValorTotal = valorTotal})
                Next

                ' Preencher o gráfico com os dados coletados
                Grafico.Series(0).Points.DataBindXY(anos, valores)

            Catch ex As Exception
                MessageBox.Show("Não foi possível gerar o gráfico!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
            Return resultado
        End Function
        Public Function DespesasAnual(Grafico As Chart, Ano As Integer) As List(Of clsComercial)
            Dim resultado As New List(Of clsComercial)()
            Dim Meses As New List(Of Integer)()
            Dim valores As New List(Of Decimal)()
            Dim dt As DataTable = ConsultaFaturamento("SELECT * FROM Cs_DespesasAnual WHERE ANO = @ano ORDER BY Ano, Mes", Ano)

            Try
                For Each row As DataRow In dt.Rows
                    Dim Mes As Integer = If(Not row.IsNull("Mes"), Convert.ToInt32(row("Mes")), 0)
                    Dim ValorTotal As Decimal = If(Not row.IsNull("Pagamentos"), Convert.ToDecimal(row("Pagamentos")), 0)

                    Meses.Add(Mes)
                    valores.Add(ValorTotal)

                    ' Também pode adicionar ao resultado se necessário:
                    resultado.Add(New clsComercial() With {.Mes = Mes, .ValorTotal = ValorTotal})
                Next
                Grafico.Series(0).Points.DataBindXY(Meses, valores)

            Catch ex As Exception
                MessageBox.Show("Não foi possível gerar o gráfico!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
            Return resultado
        End Function
        Public Function ReceitasAnual(Grafico As Chart, Ano As Integer)
            Dim resultado As New List(Of clsComercial)()
            Dim Meses As New List(Of Integer)()
            Dim valores As New List(Of Decimal)()
            Dim dt As DataTable = ConsultaFaturamento("SELECT * FROM Cs_ReceitasAnual WHERE ANO = @ano ORDER BY Ano, Mes", Ano)

            Try
                For Each row As DataRow In dt.Rows
                    Dim Mes As Integer = If(Not row.IsNull("Mes"), Convert.ToInt32(row("Mes")), 0)
                    Dim ValorTotal As Decimal = If(Not row.IsNull("Recebimento"), Convert.ToDecimal(row("Recebimento")), 0)


                    Meses.Add(Mes)
                    valores.Add(ValorTotal)

                    ' Também pode adicionar ao resultado se necessário:
                    resultado.Add(New clsComercial() With {.Mes = Mes, .ValorTotal = ValorTotal})
                Next
                Grafico.Series(0).Points.DataBindXY(Meses, valores)

            Catch ex As Exception
                MessageBox.Show("Não foi possível gerar o gráfico!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
            Return resultado
        End Function
        Public Function FaturamentoMensal(Grafico As Chart, Ano As Integer) As List(Of clsComercial)
            Dim resultado As New List(Of clsComercial)()
            Dim Meses As New List(Of Integer)()
            Dim valores As New List(Of Decimal)()
            Dim dt As DataTable = ConsultaFaturamento("SELECT * FROM Cs_FaturamentoMensal WHERE ANO = @ano ORDER BY Ano, Mes", Ano)

            Try
                For Each row As DataRow In dt.Rows
                    Dim Mes As Integer = If(Not row.IsNull("Mes"), Convert.ToInt32(row("Mes")), 0)
                    Dim ValorTotal As Decimal = If(Not row.IsNull("TOTAL"), Convert.ToDecimal(row("TOTAL")), 0)

                    Meses.Add(Mes)
                    valores.Add(ValorTotal)

                    ' Também pode adicionar ao resultado se necessário:
                    resultado.Add(New clsComercial() With {.Mes = Mes, .ValorTotal = ValorTotal})
                Next
                Grafico.Series(0).Points.DataBindXY(Meses, valores)

            Catch ex As Exception
                MessageBox.Show("Não foi possível gerar o gráfico!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
            Return resultado
        End Function
        Public Function FaturamentoPorLoja(Grafico As Chart, Ano As Integer)
            Dim resultado As New List(Of clsComercial)()
            Dim Valores As New List(Of Decimal)()
            Dim Lojas As New List(Of String)()
            Dim Quantidades As New List(Of Integer)()
            Dim dt As DataTable = ConsultaFaturamento("SELECT NomeFantasia, Valor, Quantidade FROM Cs_FaturamentoPorOrigem WHERE ANO = @Ano", Ano)

            Try
                For Each row As DataRow In dt.Rows
                    Dim ValorTotal As Decimal = If(Not row.IsNull("Valor"), Convert.ToDecimal(row("Valor")), 0)
                    Dim Loja As String = If(Not row.IsNull("NomeFantasia"), row("NomeFantasia").ToString, String.Empty)
                    Dim QuantidadeTotal As Integer = If(Not row.IsNull("Quantidade"), Convert.ToInt32(row("Quantidade")), 0)

                    Lojas.Add(Loja)
                    Valores.Add(ValorTotal)
                    Quantidades.Add(QuantidadeTotal) ' Adiciona a quantidade à lista

                    ' Adiciona ao resultado se necessário
                    resultado.Add(New clsComercial() With {.Loja = Loja, .ValorTotal = ValorTotal, .Quantidade = QuantidadeTotal})
                Next

                ' Vincula os dados ao gráfico
                Grafico.Series(0).Points.DataBindXY(Lojas, Valores)

                ' Adiciona os dados de quantidade a uma segunda série no gráfico, se necessário
                If Grafico.Series.Count > 1 Then
                    Grafico.Series(1).Points.DataBindXY(Lojas, Quantidades) ' Assumindo que a segunda série existe
                Else
                    ' Caso não exista uma segunda série, você pode criar uma
                    Dim serieQuantidade As New Series("Quantidade")
                    serieQuantidade.ChartType = SeriesChartType.Column ' ou outro tipo que você preferir
                    Grafico.Series.Add(serieQuantidade)
                    serieQuantidade.Points.DataBindXY(Lojas, Quantidades)
                End If

            Catch ex As Exception
                MessageBox.Show("Não foi possível gerar o gráfico!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
            Return resultado
        End Function
#End Region
    End Class
End Namespace

