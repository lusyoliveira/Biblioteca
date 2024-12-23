Imports Microsoft.Data.SqlClient
Imports System.Text
Imports Biblioteca.Classes.Conexao
Imports System.Data

Namespace Classes.Produtos
    Public Class clsItem
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property CodItem As Integer
        Public Property CodProduto As Integer
        Public Property CodDepartamento As Integer
        Public Property CodTipoItem As Integer
        Public Property CodForn As Integer
        Public Property TipoItem As String
        Public Property Unidade As String
        Public Property CodUnidade As Integer
        Public Property ProdutoAcabado As String
        Public Property Departamento As String
        Public Property Item As String
        Public Property NCM As String
        Public Property Investimento As String
        Public Property Peso As String
        Public Property PrecoVenda As Decimal
        Public Property PrecoCompra As Decimal
        Public Property PrecoFab As Decimal
        Public Property PrecoAtacado As Decimal
        Public Property PrecoPromocao As Decimal
        Public Property Estoque As Integer
        Public Property Quantidade As Integer
        Public Property Pai As Boolean
        Public Property Inativo As Integer
        Public Property DepartamentoPai As String
        Public Property NivelDepto As Integer
        Public Property EspecificacaoProd As String
        Public Property EspecificacaoVenda As String
        Public Property LeadTime As Integer
        Public Property Fornecedor As String
        Private Property _CodProd As Integer
        Public Property CodProd As Integer
            Get
                Return _CodProd
            End Get
            Set(value As Integer)
                _CodProd = value
            End Set
        End Property
        Public Property CodigoBarras As String
        Public Property Descricao As String
        Public Property Fator As Integer
        Public Property UnidadeCompra As String
#End Region
#Region "CONSTRUTORES"
        Public Sub New()

        End Sub
        Public Sub New(_coditem As Integer, _item As String, _preco As Decimal)
            CodItem = _coditem
            Item = _item
            PrecoVenda = _preco
        End Sub
#End Region
#Region "METODOS"
        Public Sub ValidaItem(Operacao As String, CodItem As Integer, Inativo As Integer, CodForn As Integer)
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@OPERACAO", Operacao),
                     New SqlParameter("@CODITEM", CodItem),
                     New SqlParameter("@SITUACAOSITEMA", Inativo),
                     New SqlParameter("@CODFORN", CodForn)
        }
            Dim Tabela As DataTable = ClasseConexao.ExecProcedureRetorno("spValidatem", parameters)
            If Tabela IsNot Nothing AndAlso Tabela.Rows.Count > 0 Then
                If Tabela.Rows(0)("Mensagem").ToString() <> "OK" Then
                    Dim Mensagem As String = Tabela.Rows(0)("Mensagem").ToString()
                    MessageBox.Show(Mensagem, "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        End Sub
        Public Sub SalvarItem(Descricao As String, Und As Integer, NCM As Integer, Peso As Integer, Descontinuado As Integer, TipoProd As Integer, LeadTime As Integer, Notas As String, Investimento As Decimal, IDProdutoShopee As Integer, Departamento As Integer)
            Dim sql As String = "INSERT INTO Tbl_CadProd      (Descricao,
                                                        Und,
                                                        Departamento, 
                                                        NCM,
                                                        Peso,
                                                        Descontinuado,
                                                        Tipo_Prod,
                                                        Lead_Time,
                                                        Notas,
                                                        Investimento,
                                                        DataCadastro,
                                                        IDProdutoShopee)
                                 VALUES                 (@Descricao,
                                                        @Und,
                                                        @CodDepto,
                                                        @NCM,
                                                        @Peso,
                                                        @Descontinuado,
                                                        @TipoProd,
                                                        @LeadTime,
                                                        @Notas,
                                                        @Investimento,
                                                        GETDATE(),
                                                        @IDProdutoShopee)"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@Descricao", Descricao),
                     New SqlParameter("@Und", Und),
                     New SqlParameter("@NCM", NCM),
                     New SqlParameter("@Peso", Peso),
                     New SqlParameter("@Descontinuado", Descontinuado),
                     New SqlParameter("@TipoProd", TipoProd),
                     New SqlParameter("@LeadTime", LeadTime),
                     New SqlParameter("@Notas", Notas),
                     New SqlParameter("@Investimento", Investimento),
                     New SqlParameter("@IDProdutoShopee", IDProdutoShopee),
                     New SqlParameter("@CodDepto", Departamento)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizarItem(CodSimples As Integer, Descricao As String, Und As Integer, NCM As Integer, Peso As String, Descontinuado As Integer, TipoProd As Integer, LeadTime As Integer, Notas As String, Investimento As Decimal, IDProdutoShopee As Integer, Departamento As Integer)
            Dim sql As String = "UPDATE    Tbl_CadProd SET    Descricao = @Descricao, 
                                                        Und = @Und, 
                                                        NCM = @NCM, 
                                                        Peso = @Peso, 
                                                        Descontinuado = @Descontinuado, 
                                                        Tipo_Prod = @TipoProd,  
                                                        Lead_Time = @LeadTime,
                                                        Investimento = @Investimento, 
                                                        DataAlteracao = GETDATE(),
                                                        Notas = @Notas,
                                                        IDProdutoShopee = @IDProdutoShopee, 
                                                        Departamento = @CodDepto
                                                WHERE   Cod_Simples = @CodSimples"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodSimples", CodSimples),
                     New SqlParameter("@Descricao", Descricao),
                     New SqlParameter("@Und", Und),
                     New SqlParameter("@NCM", NCM),
                     New SqlParameter("@Peso", Peso),
                     New SqlParameter("@Descontinuado", Descontinuado),
                     New SqlParameter("@TipoProd", TipoProd),
                     New SqlParameter("@LeadTime", LeadTime),
                     New SqlParameter("@Notas", Notas),
                     New SqlParameter("@Investimento", Investimento),
                     New SqlParameter("@IDProdutoShopee", IDProdutoShopee),
                     New SqlParameter("@CodDepto", Departamento)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluiItem(CodSimples As Integer)
            Dim sql As String = "DELETE FROM Tbl_CadProd WHERE Cod_Simples = @CodSimples"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodSimples", CodSimples)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub IncluirItemFornecedor(CodForn As Integer, CodProd As String, CodSimples As Integer, CodigoBarras As String, UnidadeCompra As Integer, Fator As Integer)

            Dim sql = "INSERT INTO Tbl_itemFornecedor
                                            (CodForn,
                                            CodProd,
                                            CodSimples,
                                            CodigoBarras,
                                            UnidadeCompra,
                                            Fator,
                                            DataCadastro)
                                    VALUES (@CodForn,
                                            @CodProd,
                                            @CodSimples,
                                            @CodigoBarras,
                                            @UnidadeCompra,
                                            @Fator,
                                            GETDATE())"
            Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodForn", CodForn),
                    New SqlParameter("@CodProd", CodProd),
                    New SqlParameter("@CodSimples", CodSimples),
                    New SqlParameter("@CodigoBarras", CodigoBarras),
                    New SqlParameter("@UnidadeCompra", UnidadeCompra),
                    New SqlParameter("@Fator", Fator)
}
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Fornecedor inserido com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizarItemFornecedor(CodForn As Integer, CodProd As String, CodSimples As Integer, CodigoBarras As String, UnidadeCompra As Integer, Fator As Integer)
            Dim sql As String = "UPDATE Tbl_itemFornecedor SET        CodProd = @CodProd, 
                                                                CodigoBarras = @CodigoBarras,
                                                                UnidadeCompra = @UnidadeCompra,
                                                                Fator = @Fator,
                                                                DataAlteracao =  GETDATE()
                                                      WHERE     CodSimples = @CodSimples
                                                      AND       CodForn = @CodForn"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodForn", CodForn),
                     New SqlParameter("@CodProd", CodProd),
                     New SqlParameter("@CodSimples", CodSimples),
                     New SqlParameter("@CodigoBarras", CodigoBarras),
                     New SqlParameter("@UnidadeCompra", UnidadeCompra),
                     New SqlParameter("@Fator", Fator)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluirItemFornecedor(Codigo As Integer, CodSimples As Integer)
            Dim sql As String = "DELETE FROM Tbl_itemFornecedor WHERE CodSimples = @CodSimples AND Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@Codigo", Codigo),
                     New SqlParameter("@CodSimples", CodSimples)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluirItensFornecedor(CodSimples As Integer)
            Dim sql As String = "DELETE FROM Tbl_itemFornecedor WHERE CodSimples = @CodSimples"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodSimples", CodSimples)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo registro um departamento no banco de dados.
        ''' </summary>
        ''' <param name="Depto">Representa o nome do departamento.</param>
        ''' <param name="Pai">Representa o departamento pai.</param>
        ''' <param name="Inativo">Representa a situação do departamento.</param>
        ''' <param name="CodDeptoPai">Representa o código do departamento pai.</param>
        ''' <param name="Nivel">Representa o nível do departamento</param>
        Public Sub SalvaDepartamento(Depto As String, Pai As Integer, Inativo As Integer, CodDeptoPai As Integer, Nivel As Integer, Lucro As Decimal)
            Dim sql As String = "INSERT INTO Tbl_Depto (Depto,Pai,Inativo,CodDeptoPai,Nivel,Lucro) VALUES (@Depto,@Pai,@Inativo,@CodDeptoPai,@Nivel,@Lucro)"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@Depto", Depto),
        New SqlParameter("@Pai", Pai),
        New SqlParameter("@Inativo", Inativo),
        New SqlParameter("@CodDeptoPai", CodDeptoPai),
        New SqlParameter("@Nivel", Nivel),
        New SqlParameter("@Lucro", Lucro)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo atualiza as infomrações do departamento no banco de dados.
        ''' </summary>
        ''' <param name="CodDepto">Representa o código do departamento.</param>
        ''' <param name="Depto">Representa o nome do departamento.</param>
        ''' <param name="Pai">Representa o departamento pai.</param>
        ''' <param name="Inativo">Representa a situação do departamento.</param>
        ''' <param name="CodDeptoPai">Representa o código do departamento pai.</param>
        ''' <param name="Nivel">Representa o nível do departamento</param>
        Public Sub AtualizaDepartamento(CodDepto As String, Depto As String, Pai As Integer, Inativo As Integer, CodDeptoPai As Integer, Nivel As Integer, Lucro As Decimal)
            Dim sql As String = "UPDATE Tbl_Depto SET Depto = @Depto,Pai = @Pai,Inativo = @Inativo,CodDeptoPai = @CodDeptoPai,Nivel = @Nivel, Lucro = @Lucro WHERE Codigo = @CodDepto"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@CodDepto", CodDepto),
        New SqlParameter("@Depto", Depto),
        New SqlParameter("@Pai", Pai),
        New SqlParameter("@Inativo", Inativo),
        New SqlParameter("@CodDeptoPai", CodDeptoPai),
        New SqlParameter("@Nivel", Nivel),
        New SqlParameter("@Lucro", Lucro)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo realiza a exclusão de um departamento no banco de dados.
        ''' </summary>
        ''' <param name="CodDepto">Representa o código do departamento.</param>
        Public Sub ExcluirDepartamento(CodDepto As Integer)
            Dim sql As String = "DELETE FROM Tbl_Depto WHERE Codigo = @CodDepto"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@CodDepto", CodDepto)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo registra um tipo de item no banco de dados.
        ''' </summary>
        ''' <param name="TipoItem">Representa a descrição do tipo.</param>
        Public Sub SalvaTipoItem(TipoItem As String)
            Dim sql As String = "INSERT INTO Tbl_TipoItem (TipoItem) VALUES (@TipoItem)"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@TipoItem", TipoItem)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo atualiza os dados do tipo de item.
        ''' </summary>
        ''' <param name="Codigo">Representa o código do tipo de item.</param>
        ''' <param name="TipoItem">Representa a descrição do tipo.</param>
        Public Sub AtualizaTipoItem(Codigo As Integer, TipoItem As String)
            Dim sql As String = "UPDATE Tbl_TipoItem SET TipoItem = @TipoItem WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@Codigo", Codigo),
        New SqlParameter("@TipoItem", TipoItem)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo realiza a exclusão de um tipo de item.
        ''' </summary>
        ''' <param name="Codigo">Representa o código do tipo de item.</param>
        Public Sub ExcluiTipoItem(Codigo As Integer)
            Dim sql As String = "DELETE Tbl_TipoItem WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@Codigo", Codigo)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        ''' <summary>
        ''' Este metódo registra um unidade de medida do banco de dados.
        ''' </summary>
        ''' <param name="Unidade">Representa o nome da unidade de medida.</param>
        ''' <param name="Simbolo">Representa a sigla da unidade de medida.</param>
        Public Sub SalvaUnidadeMedida(Unidade As String, Simbolo As String)
            Dim sql As String = "INSERT INTO Tbl_UnidadeMedida (Descricao,Simbolo) VALUES (@Unidade,@Simbolo)"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@Unidade", Unidade),
        New SqlParameter("@Simbolo", Simbolo)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub
        ''' <summary>
        ''' Este metódo atualiza os dados da unidade de medida no banco de dados.
        ''' </summary>
        ''' <param name="Codigo">Representa o código da unidade de medida</param>
        ''' <param name="Unidade">Representa o nome da unidade de medida.</param>
        ''' <param name="Simbolo">Representa a sigla da unidade de medida.</param>
        Public Sub AtualizaUnidadeMedida(Codigo As Integer, Unidade As String, Simbolo As String)
            Dim sql As String = "UPDATE Tbl_UnidadeMedida SET Descricao = @Unidade, Simbolo = @Simbolo WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@Codigo", Codigo),
        New SqlParameter("@Unidade", Unidade),
        New SqlParameter("@Simbolo", Simbolo)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub
        ''' <summary>
        ''' Este metódo realiza a exclusão de unidade de medida do banco de dados.
        ''' </summary>
        ''' <param name="Codigo">Representa o código da unidade de medida</param>
        Public Sub ExcluirUnidadeMedida(Codigo As Integer)
            Dim sql As String = "DELETE FROM Tbl_UnidadeMedida WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@Codigo", Codigo)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
#End Region
#Region "FUNCOES"
        Public Function PesquisaItem(Status As String, CodItem As String, Produto As String, CadastroIni As String, CadastroFim As String, AlteracaoIni As String, AlteracaoFim As String, InativacaoIni As String, InativacaoFim As String, Departamento As String, TipoItem As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_CadProd WHERE 1=1 ")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Inativo"
                        sql.AppendLine("AND Descontinuado = 1")
                    Case "Ativo"
                        sql.AppendLine("AND Descontinuado = 0")
                    Case "Todos"
                        sql.AppendLine("AND Descontinuado IS NOT NULL")
                End Select

                'Pesquisa pela código do item
                If Not String.IsNullOrEmpty(CodItem) Then
                    sql.AppendLine("AND CodItem = @CodSimples")
                    parameters.Add(New SqlParameter("@CodSimples", CodItem))
                End If

                'Pesquisa pela item
                If Not String.IsNullOrEmpty(Produto) Then
                    sql.AppendLine("AND Item LIKE @Produto")
                    parameters.Add(New SqlParameter("@Produto", "%" & Produto & "%"))
                End If

                'Pesquisa pela data de cadastro
                If Not String.IsNullOrEmpty(CadastroIni) And Not String.IsNullOrEmpty(CadastroFim) Then
                    sql.AppendLine("AND DataCadastro BETWEEN @DataCadIni AND @DataCadFim")
                    parameters.Add(New SqlParameter("@DataCadIni", CadastroIni))
                    parameters.Add(New SqlParameter("@DataCadFim", CadastroFim))
                ElseIf Not String.IsNullOrEmpty(CadastroIni) And String.IsNullOrEmpty(CadastroFim) Then
                    sql.AppendLine("AND DataCadastro >= @DataCadIni")
                    parameters.Add(New SqlParameter("@DataCadIni", CadastroIni))
                ElseIf String.IsNullOrEmpty(CadastroIni) And Not String.IsNullOrEmpty(CadastroFim) Then
                    sql.AppendLine("AND DataCadastro <= @DataCadFim")
                    parameters.Add(New SqlParameter("@DataCadFim", CadastroFim))
                End If

                'Pesquisa pela data de alteração
                If Not String.IsNullOrEmpty(AlteracaoIni) And Not String.IsNullOrEmpty(AlteracaoFim) Then
                    sql.AppendLine("AND DataAlteracao BETWEEN @AlteracaoIni AND @AlteracaoFim")
                    parameters.Add(New SqlParameter("@AlteracaoIni", AlteracaoIni))
                    parameters.Add(New SqlParameter("@AlteracaoFim", AlteracaoFim))
                ElseIf Not String.IsNullOrEmpty(AlteracaoIni) And Not String.IsNullOrEmpty(AlteracaoFim) Then
                    sql.AppendLine("AND DataAlteracao >= @AlteracaoIni")
                    parameters.Add(New SqlParameter("@AlteracaoIni", AlteracaoIni))
                ElseIf Not String.IsNullOrEmpty(AlteracaoIni) And Not String.IsNullOrEmpty(AlteracaoFim) Then
                    sql.AppendLine("AND DataAlteracao <= @AlteracaoFim")
                    parameters.Add(New SqlParameter("@AlteracaoFim", AlteracaoFim))
                End If

                'Pesquisa pela data de inativação
                If Not String.IsNullOrEmpty(InativacaoIni) And Not String.IsNullOrEmpty(InativacaoFim) Then
                    sql.AppendLine("AND DataInativacao BETWEEN @InativacaoIni AND @InativacaoFim")
                    parameters.Add(New SqlParameter("@InativacaoIni", InativacaoIni))
                    parameters.Add(New SqlParameter("@InativacaoFim", InativacaoFim))
                ElseIf Not String.IsNullOrEmpty(InativacaoIni) And Not String.IsNullOrEmpty(InativacaoFim) Then
                    sql.AppendLine("AND DataInativacao >= @InativacaoIni")
                    parameters.Add(New SqlParameter("@InativacaoIni", InativacaoIni))
                ElseIf Not String.IsNullOrEmpty(InativacaoIni) And Not String.IsNullOrEmpty(InativacaoFim) Then
                    sql.AppendLine("AND DataInativacao <= @InativacaoFim")
                    parameters.Add(New SqlParameter("@InativacaoFim", InativacaoFim))
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

                sql.AppendLine("ORDER BY Item")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As SqlException
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function
        Public Function ConsultaItem(sql As String, Optional CodItem As Integer = 0, Optional Item As String = Nothing)
            If CodItem <> 0 Then
                Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodItem", CodItem)
                        }
                Return ClasseConexao.Consultar(sql, parameters)
            ElseIf Item <> "" Or Item IsNot Nothing Then
                Dim parameters As SqlParameter() = {
                     New SqlParameter("@Item", Item)
                        }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function

        ''' <summary>
        ''' Esta função realizar a consulta do dados do tipo de item.
        ''' </summary>
        ''' <param name="sql">Representa uma query sql.</param>
        ''' <returns>Retorna os dados da consulta.</returns>
        Public Function ConsultaTipoItem(sql As String)
            Return ClasseConexao.Consultar(sql, Nothing)
        End Function
        ''' <summary>
        ''' Esta função realizar a consulta do dados do unidade de medida.
        ''' </summary>
        ''' <param name="sql">Representa uma query sql.</param>
        ''' <returns>Retorna os dados da consulta.</returns>
        Public Function ConsultaUnidadeMedida(sql As String)
            Return ClasseConexao.Consultar(sql, Nothing)
        End Function
        ''' <summary>
        ''' Esta função realizar a consulta do dados do departamento.
        ''' </summary>
        ''' <param name="sql">Representa uma query sql.</param>
        ''' <returns>Retorna os dados da consulta.</returns>
        Public Function ConsultaDepartamento(sql As String, Optional Departamento As String = Nothing)
            If Departamento IsNot Nothing Then
                Dim parameters As SqlParameter() = {
                     New SqlParameter("@Departamento", Departamento)
                        }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
        ''' <summary>
        ''' Esta função preenche uma TreeView com os departamentos existente.
        ''' </summary>
        ''' <param name="Arvore">Representa o nome do componente TreeView.</param>
        ''' <returns>Retorna a lista de departamentos.</returns>
        Public Function CriarArvore(Arvore As TreeView)
            Dim nodeMap As New Dictionary(Of Integer, TreeNode)()
            Dim tbArvore As DataTable = ConsultaDepartamento("SELECT * FROM Cs_Departamentos")

            ' Lê os dados do DataTable e constrói os nós
            For Each row As DataRow In tbArvore.Rows
                Dim id As Integer = Convert.ToInt32(row("Codigo"))
                Dim nome As String = Convert.ToString(row("Depto"))
                Dim Departamento As String = Convert.ToString(row("Departamento"))
                Dim nivel As Integer = Convert.ToInt32(row("CodNivel"))
                Dim paiId As Integer? = If(IsDBNull(row("CodDeptoPai")), Nothing, Convert.ToInt32(row("CodDeptoPai")))

                ' Cria o nó correspondente
                Dim newNode As TreeNode = New TreeNode(Departamento)
                nodeMap(id) = newNode

                ' Adiciona o nó ao TreeView ou ao seu pai
                If paiId = 0 Then
                    ' Se não tem pai, adiciona ao nível raiz
                    Arvore.Nodes.Add(newNode)
                ElseIf nodeMap.ContainsKey(paiId.Value) Then
                    ' Se tem pai, adiciona ao nó pai
                    nodeMap(paiId.Value).Nodes.Add(newNode)
                End If
            Next
        End Function
#End Region

    End Class

End Namespace
