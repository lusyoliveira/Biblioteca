Imports System.Data.SqlClient
Imports System.Text
Imports Biblioteca.Classes.Produtos
Imports Biblioteca.Classes.Conexao
''' <summary>
''' Esta classe representa todos os metódos e funções que manipulam o preço do item.
''' </summary>
Public Class clsPreco
    Inherits clsItem
    Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADE"
    Public Property CodPromocao As Integer
    Public Property TotalItens As Integer
    Public Property ValorUnit As Decimal
    Public Property Tipo As String
    Public Property TipoDesconto As String
    Public Property CustoFixo As Decimal
    Public Property Motivo As String
    Public Property DataMovimento As Date
    Public Property DataInicial As Date
    Public Property DataFinal As Date
    Public Property CodMotivo As Integer
    Public Property Loja As String
    Public Property CodLoja As Integer
    Public Property Status As String
#End Region
#Region "CONSTRUTORES"
    Public Sub New()

    End Sub
    Public Sub New(_codigo As Integer, _item As String, _quantidade As Integer, _valorunit As Decimal, _compra As Decimal, _venda As Decimal, _atacado As Decimal, _fab As Decimal)
        CodItem = _codigo
        Item = _item
        Quantidade = _quantidade
        ValorUnit = _valorunit
        PrecoCompra = _compra
        PrecoVenda = _venda
        PrecoAtacado = _atacado
        PrecoFab = _fab
    End Sub
#End Region
#Region "METODOS"
    Public Sub AtualizaHistoricoPreco(CodItem As Integer, Motivo As Integer, Compra As Decimal, Venda As Decimal, Fab As Decimal, Atacado As Decimal, CodLoja As Integer)
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@COD_SIMPLES", CodItem),
                    New SqlParameter("@MOTIVO", Motivo),
                    New SqlParameter("@PRECOCOMPRA", Compra),
                    New SqlParameter("@PRECOVENDA", Venda),
                    New SqlParameter("@PRECOATACADO", Atacado),
                    New SqlParameter("@PRECOFAB", Fab),
                    New SqlParameter("@CODLOJA", CodLoja)
        }
        ClasseConexao.ExecutarProcedure("spAlteraHistoricoPreco", parameters)
    End Sub
    ''' <summary>
    ''' Este metódo executa uma procedure no banco de dados para voltar o preço de um itens.
    ''' </summary>
    ''' <param name="CodItem">Representa o código do item.</param>
    Public Sub VoltaPrecoItem(CodItem As Integer)

        Dim parameters As SqlParameter() = {
            New SqlParameter("@COD_SIMPLES", CodItem)
    }
        ClasseConexao.ExecutarProcedure("spVoltarHistoricoPreco", parameters)
    End Sub

    Public Sub CalculaPreco(Grid As DataGridView, CustoFixo As Decimal, TaxaCartao As Decimal, Comissao As Decimal, TaxaFixa As Decimal)
        TaxaCartao = TaxaCartao / 100
        Comissao = Comissao / 100
        CustoFixo = CustoFixo / 220

        For Each col As DataGridViewRow In Grid.Rows
            col.Cells("CustoFixo").Value = Math.Round((CustoFixo / 60) * col.Cells("Prazo").Value, 2) 'Custo Fixo Hora do Atelie
            col.Cells("PrecoFab").Value = col.Cells("CustoProd").Value + col.Cells("CustoFixo").Value 'Preço de Fabricação
            col.Cells("Taxa").Value = Math.Round(col.Cells("PrecoFab").Value * TaxaCartao, 2) 'Valor da Taxa de Cartão
            col.Cells("vComissao").Value = Math.Round((col.Cells("PrecoFab").Value + col.Cells("Taxa").Value) * Comissao, 2) 'Valor da Comissão
            col.Cells("CustoReal").Value = Math.Round((col.Cells("PrecoFab").Value + col.Cells("Taxa").Value + col.Cells("vComissao").Value + TaxaFixa), 2) 'Custo Real
            col.Cells("vLucro").Value = Math.Round((col.Cells("CustoReal").Value) * col.Cells("Lucro").Value, 2) 'Valor do Lucro
            col.Cells("PrecoVenda").Value = ArredondarCima((col.Cells("CustoReal").Value + col.Cells("vLucro").Value), 0.5D) 'Preço de Venda
            col.Cells("PrecoCompra").Value = Math.Round(col.Cells("CustoReal").Value, 2) 'Preço de compra
            col.Cells("PrecoAtacado").Value = ArredondarCima(col.Cells("CustoReal").Value, 0.5D) 'Preço de Atacado
        Next
    End Sub
    Public Function ConsultaPromocao(sql As String, Optional CodProd As Integer = 0)
        If CodProd <> 0 Then
            Dim parameters As SqlParameter() = {
                New SqlParameter("@CodProd", CodProd)
                }
            Return ClasseConexao.Consultar(sql, parameters)
        Else
            Return ClasseConexao.Consultar(sql, Nothing)
        End If
    End Function
    Public Sub ExcluiHistoricoPreco(CodSimples As Integer)
        Dim sql As String = "DELETE FROM Tbl_HistoricoPreco WHERE Cod_Simples = @CodSimples"
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodSimples", CodSimples)
        }
        ClasseConexao.Operar(sql, parameters)
    End Sub
    Public Sub SalvaPromocao(Descricao As String, DataVigInicio As Date, DataVigFim As Date, Inativo As Boolean, Loja As Integer, TipoPromocao As String)
        Dim sql As String = "INSERT INTO Tbl_Promocao (Descricao,DataVigInicio,DataVigFim,Inativo,DataCadastro,Loja,TipoPromocao) VALUES (@Descricao,@DataVigInicio,@DataVigFim,@Inativo,GETDate(),@Loja,@TipoPromocao)"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@Descricao", Descricao),
                    New SqlParameter("@DataVigInicio", DataVigInicio),
                    New SqlParameter("@DataVigFim", DataVigFim),
                    New SqlParameter("@Loja", Loja),
                    New SqlParameter("@TipoPromocao", TipoPromocao),
                    New SqlParameter("@Inativo", Inativo)
        }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub AtualizaPromocao(CodPromocao As Integer, Descricao As String, DataVigInicio As Date, DataVigFim As Date, Inativo As Boolean)
        Dim sql As String = "UPDATE Tbl_Promocao SET   Descricao = @Descricao,
                                                    DataVigInicio = @DataVigInicio,
                                                    DataVigFim = @DataVigFim,
                                                    Inativo = @Inativo 
                                                    WHERE CodPromocao = @CodPromocao"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@Descricao", Descricao),
                    New SqlParameter("@DataVigInicio", DataVigInicio),
                    New SqlParameter("@DataVigFim", DataVigFim),
                    New SqlParameter("@Loja", Loja),
                    New SqlParameter("@CodPromocao", CodPromocao),
                    New SqlParameter("@Inativo", Inativo)
        }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub ExcluirPromocao(CodPromocao As Integer)
        Dim sql As String = "DELETE FROM Tbl_Promocao WHERE CodPromocao = @CodPromocao"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodPromocao", CodPromocao)
        }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub SalvaPromocaoCampanha(Descricao As String, DataVigInicio As Date, DataVigFim As Date, Inativo As Boolean, Loja As Integer, TipoPromocao As String, Valor As Integer)
        Dim sql As String = "INSERT INTO Tbl_Promocao (Descricao,DataVigInicio,DataVigFim,Inativo,DataCadastro,Loja,TipoPromocao) VALUES (@Descricao,@DataVigInicio,@DataVigFim,@Inativo,GETDate(),@Loja,@TipoPromocao)"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@Descricao", Descricao),
                    New SqlParameter("@DataVigInicio", DataVigInicio),
                    New SqlParameter("@DataVigFim", DataVigFim),
                    New SqlParameter("@Loja", Loja),
                    New SqlParameter("@TipoPromocao", TipoPromocao),
                    New SqlParameter("@Inativo", Inativo),
                    New SqlParameter("@Valor", Valor)
        }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub AtualizaPromocaoCampanha(CodPromocao As Integer, Descricao As String, DataVigInicio As Date, DataVigFim As Date, Inativo As Boolean, TipoPromocao As String, Valor As Integer)
        Dim sql As String = "UPDATE Tbl_Promocao SET  Descricao = @Descricao,
                                                    DataVigInicio = @DataVigInicio,
                                                    DataVigFim = @DataVigFim,
                                                    Inativo = @Inativo,
                                                    TipoPromocao = @TipoPromocao,
                                                    Valor = @Valor
                                                    WHERE CodPromocao = @CodPromocao"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodPromocao", CodPromocao),
                    New SqlParameter("@Descricao", Descricao),
                    New SqlParameter("@DataVigInicio", DataVigInicio),
                    New SqlParameter("@DataVigFim", DataVigFim),
                    New SqlParameter("@Loja", Loja),
                    New SqlParameter("@TipoPromocao", TipoPromocao),
                    New SqlParameter("@Inativo", Inativo),
                    New SqlParameter("@Valor", Valor)
        }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub AtualizaItemPromocaoCampanha(CodSimples As Integer, Preco As Decimal?, CodPromo As Integer)
        Dim parameters As SqlParameter() = {
        New SqlParameter("@CodPromocao", CodPromo),
        New SqlParameter("@CodSimples", CodSimples)
    }

        Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT COUNT(Cod_Simples) FROM Cs_PromocaoDetalhes WHERE Cod_Simples = @CodSimples AND CodPromocao = @CodPromo", parameters)
        If Tabela.Rows.Count > 0 Then
            Dim sql As String = "UPDATE     Tbl_ItemPromocaoCampanha SET      
                                                            Preco = @Preco
                                                    WHERE   CodPromocao = @CodPromocao
                                                    AND     Cod_Simples = @CodSimples"

            Array.Resize(parameters, parameters.Length + 1)
            parameters(parameters.Length - 1) = New SqlParameter("@Preco", Preco.Value)

            ClasseConexao.Operar(sql, parameters)
        ElseIf Tabela.Rows.Count = 0 Then
            SalvaItemPromocaoCampanha(CodPromo, CodSimples, Preco)
        End If
    End Sub
    Public Sub SalvaItemPromocaoCampanha(CodPromocao As String, CodSimples As String, Preco As Decimal)
        Dim sql As String = "INSERT INTO Tbl_ItemPromocaoCampanha (CodPromocao,Cod_Simples,Preco) VALUES (@CodPromocao,@CodSimples,@Preco)"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodPromocao", CodPromocao),
                    New SqlParameter("@CodSimples", CodSimples),
                    New SqlParameter("@Preco", Preco)
                    }
        ClasseConexao.Operar(sql, parameters)
    End Sub
    Public Sub ExcluiDetPromocaoCampanha(CodPromocao As Integer)
        Dim sql As String = "DELETE FROM Tbl_ItemPromocaoCampanha WHERE CodPromocao = @CodPromocao"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodPromocao", CodPromocao)
                    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub ExcluiItemPromocaoCampanha(CodPromocao As Integer, CodSimples As Integer)
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodPromocao", CodPromocao),
                    New SqlParameter("@CodSimples", CodSimples)
        }

        Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT COUNT(CodPromocao) FROM Cs_PromocaoDetalhes WHERE CodPromocao = @CodPromocao AND Cod_Simples = @CodSimples", parameters)

        If Tabela.Rows.Count > 0 Then
            Dim sql As String = "DELETE FROM Tbl_ItemPromocaoCampanha WHERE CodPromocao = @CodPromocao AND Cod_Simples = @CodSimples"
            ClasseConexao.Operar(Sql, parameters)
        Else
            Exit Sub
        End If
    End Sub
    Public Sub SalvaItemPromocaoFaixaPreco(CodPromocao As Integer, CodSimples As Integer, De As Integer, Ate As Integer, Preco As Decimal)
        Dim sql As String = "INSERT INTO Tbl_ItemPromocaoFaixaPreco (CodPromocao,Cod_Simples,De,Preco,Ate) VALUES (@CodPromocao,@CodSimples,@De,@Preco,@Ate)"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodPromocao", CodPromocao),
                    New SqlParameter("@CodSimples", CodSimples),
                    New SqlParameter("@De", De),
                    New SqlParameter("@Ate", Ate),
                    New SqlParameter("@Preco", Preco)
                    }
        ClasseConexao.Operar(sql, parameters)
    End Sub
    Public Sub ExcluiDetPromocaoFaixaPreco(CodPromocao As Integer)
        Dim sql As String = "DELETE FROM Tbl_ItemPromocaoFaixaPreco WHERE CodPromocao = @CodPromocao"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodPromocao", CodPromocao)
                    }
        ClasseConexao.Operar(sql, parameters)
    End Sub
    Public Sub ExcluiItemPromocaoFaixaPreco(CodPromocao As Integer, CodSimples As Integer)
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodPromocao", CodPromocao),
                    New SqlParameter("@CodSimples", CodSimples)
        }

        Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT COUNT(CodPromocao) FROM Cs_PromocaoDetalhes WHERE CodPromocao = @CodPromocao AND Cod_Simples = @CodSimples", parameters)

        If Tabela.Rows.Count > 0 Then
            Dim sql As String = "DELETE FROM Tbl_ItemPromocaoFaixaPreco WHERE CodPromocao = @CodPromocao AND Cod_Simples = @CodSimples"
            ClasseConexao.Operar(Sql, parameters)
        Else
            Exit Sub
        End If
    End Sub

    Public Sub AtualizaItemPromocaoFaixaPreco(CodSimples As Integer, De As Integer?, Ate As Integer?, Preco As Decimal?, CodPromo As Integer)
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodPromo", CodPromo),
                     New SqlParameter("@CodSimples", CodSimples)
    }

        Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT COUNT(Cod_Simples) FROM Cs_PromocaoDetalhes WHERE Cod_Simples = @CodSimples AND CodPromocao = @CodPromo", parameters)
        If Tabela.Rows.Count > 0 Then
            Dim sql As String = "UPDATE     Tbl_ItemPromocaoFaixaPreco SET      
                                                            De = @De,
                                                            Preco = @Preco,
                                                            Ate = @Ate
                                                    WHERE   CodPromocao = @CodPromocao
                                                    AND     CodSimples = @CodSimples"

            Array.Resize(parameters, parameters.Length + 1)
            parameters(parameters.Length - 1) = New SqlParameter("@De", De.Value)

            Array.Resize(parameters, parameters.Length + 1)
            parameters(parameters.Length - 1) = New SqlParameter("@Ate", Ate.Value)

            Array.Resize(parameters, parameters.Length + 1)
            parameters(parameters.Length - 1) = New SqlParameter("@Preco", Preco.Value)

            ClasseConexao.Operar(sql, parameters)

        ElseIf Tabela.Rows.Count = 0 Then
            SalvaItemPromocaoFaixaPreco(CodPromo, CodSimples, De, Ate, Preco)
        End If
    End Sub
    Public Function ConsultaMotivoPreco(sql As String, Optional CodMotivo As Integer? = 0)
        If CodMotivo <> 0 Then
            Return ClasseConexao.Consultar(sql, Nothing)
        Else
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodMotivo", CodMotivo)
                        }
            Return ClasseConexao.Consultar(sql, parameters)
        End If
    End Function
    Public Sub SalvaMotivoPreco(Motivo As String)
        Dim sql As String = "INSERT INTO Tbl_MotivoAlteracaoPreco (Motivo) VALUES (@Motivo)"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@Motivo", Motivo)
                    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub AtualizaMotivoPreco(Codigo As Integer, Motivo As String)
        Dim sql As String = "UPDATE Tbl_MotivoAlteracaoPreco SET Motivo = @Motivo WHERE Codigo = @Codigo"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@Codigo", Codigo),
                    New SqlParameter("@Motivo", Motivo)
                    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub ExcluiMotivoPreco(Codigo As Integer)
        Dim sql As String = "DELETE FROM Tbl_MotivoAlteracaoPreco WHERE Codigo = @Codigo"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@Codigo", Codigo)
                    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
#End Region
#Region "FUNCOES"
    Public Function PesquisaItens(Status As String, CodItem As Integer, Produto As String, Departamento As String, TipoItem As String) As DataTable
        Dim sql As New StringBuilder("SELECT * FROM CS_CadProd WHERE 1=1 ")
        Dim parameters As New List(Of SqlParameter)()

        Try
            Select Case Status
                Case "Inativo"
                    sql.AppendLine("AND Descontinuado = 1")
                Case "Ativo"
                    sql.AppendLine("AND Descontinuado = 0")
                Case "Todos"
                    sql.AppendLine("and Descontinuado IS NOT NULL")
            End Select

            'Pesquisa pela código do item

            If CodItem <> 0 Then
                sql.AppendLine("AND CodItem = @CodSimples")
                parameters.Add(New SqlParameter("@CodSimples", CodItem))
            End If

            'Pesquisa pela item
            If Not String.IsNullOrEmpty(Produto) Then
                sql.AppendLine("AND Item LIKE @Produto")

                parameters.Add(New SqlParameter("@Produto", "%" & Produto & "%"))
            End If

            'Pesquisa pela tipo do item
            If Not String.IsNullOrEmpty(TipoItem) Then
                sql.AppendLine("AND TipoItem LIKE @TipoItem")
                parameters.Add(New SqlParameter("@TipoItem", "%" & TipoItem & "%"))
            End If

            'Pesquisa pelo departamento
            If Not String.IsNullOrEmpty(Departamento) Then

                sql.AppendLine("AND Departamento LIKE @Departamento")
                parameters.Add(New SqlParameter("@Departamento", "%" & Departamento & "%"))
            End If

            sql.AppendLine("AND CodTipoItem IN (4, 6)")

            sql.AppendLine("ORDER BY Item")

            Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

        Catch ex As SqlException
            MessageBox.Show("Não foi possível realizar a consulta!" & ex.Message, "Erro de Banco de Dados", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Throw
        Catch ex As Exception
            MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Throw
        End Try
    End Function
    Public Function PesquisaPromocao(Status As String, CodPromo As String, Promocao As String, DataCadIni As String, DataCadFim As String, Loja As String, TipoPromo As String) As DataTable
        Dim sql As New StringBuilder("SELECT * FROM Cs_Promocao WHERE 1=1 ")
        Dim parameters As New List(Of SqlParameter)()

        Try
            Select Case Status
                Case "Ativo"
                    sql.AppendLine("AND Inativo = 0")
                Case "Inativo"
                    sql.AppendLine("AND Inativo = 1")
                Case "Todos"
                    sql.AppendLine("AND Inativo IS NOT NULL")
            End Select

            'Pesquisa por código
            If Not String.IsNullOrEmpty(CodPromo) Then
                sql.AppendLine("AND CodPromocao = @CodPromocao")
                parameters.Add(New SqlParameter("@CodPromocao", CodPromo))
            End If

                'Pesquisa por descrição
                If Not String.IsNullOrEmpty(Promocao) Then
                sql.AppendLine("AND Descricao LIKE @Promocao")
                parameters.Add(New SqlParameter("@Promocao", Promocao))
            End If

                'Pesquisa por loja
                If Not String.IsNullOrEmpty(Loja) Then
                sql.AppendLine("AND Loja LIKE @Loja")
                parameters.Add(New SqlParameter("@Loja", Loja))
            End If

            'Tipo de Promoção
            If Not String.IsNullOrEmpty(TipoPromo) Then
                sql.AppendLine("AND Tipo LIKE @Tipo")
                parameters.Add(New SqlParameter("@Tipo", TipoPromo))
            End If

            'Pesquisa pela data de vigência
            If Not String.IsNullOrEmpty(DataCadIni) And Not String.IsNullOrEmpty(DataCadFim) Then
                sql.AppendLine("AND DataVigInicio >= @DataCadIni AND DataVigFim <= @DataCadFim")
                parameters.Add(New SqlParameter("@DataCadIni", DataCadIni))
                parameters.Add(New SqlParameter("@DataCadFim", DataCadFim))
            ElseIf Not String.IsNullOrEmpty(DataCadIni) And String.IsNullOrEmpty(DataCadFim) Then
                sql.AppendLine("AND DataVigInicio >= @DataCadIni")
                parameters.Add(New SqlParameter("@DataCadIni", DataCadIni))
            ElseIf String.IsNullOrEmpty(DataCadIni) And Not String.IsNullOrEmpty(DataCadFim) Then
                sql.AppendLine("AND DataVigFim <= @DataCadFim")
                parameters.Add(New SqlParameter("@DataCadFim", DataCadFim))
            Else
                    sql.AppendLine("AND DataVigInicio IS NOT NULL OR DataVigFim IS NOT NULL")
                End If

            sql.AppendLine("ORDER BY DataVigInicio")

            Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())
        Catch ex As Exception
            MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Throw
        End Try
    End Function
    Public Function ValidarItemNaPromocao(CodItem As Integer) As Boolean
        Dim sql As String = "SELECT * FROM Cs_PromocaoDetalhes WHERE Cod_Simples = @CodSimples AND Inativo = 0"

        Dim parameters As SqlParameter() = {
    New SqlParameter("@CodSimples", CodItem)
        }

        Dim Tabela As DataTable = ClasseConexao.Consultar(sql, parameters)
        If Tabela IsNot Nothing AndAlso Tabela.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function PesquisalHistoricoPreco(DataInicio As String, DataFim As String, Status As String, CodItem As Integer, Produto As String, Motivo As String, Loja As String, Departamento As String, TipoItem As String)
        Dim sql As New StringBuilder("SELECT * FROM Cs_HistoricoPreco WHERE 1=1")

        Dim parameters As New List(Of SqlParameter)()

        Try
            Select Case Status
                Case "Inativo"
                    sql.AppendLine("AND Descontinuado = 1")
                Case "Ativo"
                    sql.AppendLine("AND Descontinuado = 0")
                Case "Todos"
                    sql.AppendLine("and Descontinuado IS NOT NULL")
            End Select

            'Pesquisa pela código do item
            If CodItem <> 0 Then
                sql.AppendLine("AND CodItem = @CodSimples")
                parameters.Add(New SqlParameter("@CodSimples", CodItem))
            End If

            'Pesquisa pela item
            If Not String.IsNullOrEmpty(Produto) Then
                sql.AppendLine("AND Item LIKE @Produto")
                parameters.Add(New SqlParameter("@Produto", "%" & Produto & "%"))
            End If

            'Pesquisa por loja
            If Not String.IsNullOrEmpty(Loja) Then
                sql.AppendLine("AND Loja LIKE @Loja")
                parameters.Add(New SqlParameter("@Loja", "%" & Loja & "%"))
            End If

            'Pesquisa pela tipo do item
            If Not String.IsNullOrEmpty(TipoItem) Then
                sql.AppendLine("AND Tipo_Prod LIKE @TipoItem")
                parameters.Add(New SqlParameter("@TipoItem", "%" & TipoItem & "%"))
            End If

            'Pesquisa pelo departamento
            If Not String.IsNullOrEmpty(Departamento) Then
                sql.AppendLine("AND Depto LIKE @Departamento")
                parameters.Add(New SqlParameter("@Departamento", "%" & Departamento & "%"))
            End If

            'Pesquisa por loja
            If Not String.IsNullOrEmpty(Motivo) Then
                sql.AppendLine("AND Motivo LIKE @Motivo")
                parameters.Add(New SqlParameter("@Motivo", "%" & Motivo & "%"))
            End If

            ''Pesquisa pela data de cadastro
            If Not String.IsNullOrEmpty(DataInicio) And Not String.IsNullOrEmpty(DataFim) Then
                sql.AppendLine("AND DataMovimento BETWEEN @DataIni AND @DataFim")
                parameters.Add(New SqlParameter("@DataIni", DataInicio))
                parameters.Add(New SqlParameter("@DataFim", DataFim))
            ElseIf Not String.IsNullOrEmpty(DataInicio) And String.IsNullOrEmpty(DataFim) Then
                sql.AppendLine("AND DataMovimento >= @DataIni")
                parameters.Add(New SqlParameter("@DataIni", DataInicio))
            ElseIf String.IsNullOrEmpty(DataInicio) And Not String.IsNullOrEmpty(DataFim) Then
                sql.AppendLine("AND DataMovimento <= @DataFim")
                parameters.Add(New SqlParameter("@DataFim", DataFim))
            Else
                sql.AppendLine("AND DataMovimento IS NOT NULL")
            End If

            sql.AppendLine("ORDER BY Item")

            Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())


        Catch ex As SqlException
            MessageBox.Show("Não foi possível realizar a consulta!" & ex.Message, "Erro de Banco de Dados", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Throw
        Catch ex As Exception
            MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Throw
        End Try
    End Function
    Public Function ConsultaPreco(sql As String, CodLoja As Integer, Optional CodItem As Integer = 0, Optional Item As String = Nothing)
        If CodItem <> 0 Then
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodItem", CodItem),
                     New SqlParameter("@CodLoja", CodLoja)
                        }
            Return ClasseConexao.Consultar(sql, parameters)
        ElseIf Item <> "" Or Item IsNot Nothing Then
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@Item", Item),
                     New SqlParameter("@CodLoja", CodLoja)
                        }
            Return ClasseConexao.Consultar(sql, parameters)
        Else
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodLoja", CodLoja)
                        }
            Return ClasseConexao.Consultar(sql, parameters)
        End If
    End Function
#End Region
End Class
