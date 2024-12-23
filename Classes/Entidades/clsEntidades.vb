Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports System.Data
Imports System.Text
Imports Xceed.Wpf.Toolkit

Namespace Classes.Entidades.Vendas

    Public Class clsEntidades
        Inherits clsLocalidades
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property SaldoCredito As Decimal
        Public Property TipoPessoa As String
        Public Property Status As String
        Public Property CodContato As Integer
        Public Property CodTipoContato As Integer
        Public Property TipoContato As String
        Public Property NomeContato As String
        Public Property Email As String
        Public Property Telefone As String
        Public Property CodConta As Integer
        Public Property NomeConta As String
        Public Property Conta As String
        Public Property Agencia As String
        Public Property CodEntidade As Integer
        Public Property Entidade As String
        Public Property TipoEntidade As String
        Public Property EnderecoInativo As Boolean
        Public Property CodEndereco As Integer
        Public Property TipoEndereco As String
#End Region
#Region "METODOS"
        ''' <summary>
        ''' Este metódo grava os dados do cliente no banco de dados.
        ''' </summary>
        ''' <param name="NomeFantasia">Representa o nome do cliente do tipo string.</param>
        ''' <param name="Endereco">Representa o endereço do cliente do tipo string</param>
        ''' <param name="Bairro">Representa o bairro do cliente do tipo string.</param>
        ''' <param name="Cep">representa o cep do cliente do tipo string</param>
        ''' <param name="UF">Representa o estado do cliente do tipo string.</param>
        ''' <param name="Cidade">Representa o cidade do cliente do tipo string.</param>
        ''' <param name="Cpfcnpj">Representa o CPF ou CNPJ do cliente do tipo string.</param>
        ''' <param name="RazaoSocial">Representa a razão social do cliente do tipo string</param>
        ''' <param name="TipoPessoa">representa o tipo de pessoa do cliente do tipo string</param>
        Public Sub SalvarEntidade(NomeFantasia As String, RazaoSocial As String, Endereco As String, Bairro As String, Cidade As Integer, UF As Integer, Cep As String, INSC As String, Cpfcnpj As String, RG As String, RamoAtividade As String, TipoEstabelecimento As String, Site As String, Referencia As String, TipoPessoa As String, TipoEntidade As String)
            Dim sql As String = "INSERT INTO Tbl_Entidades (NomeFantasia,
                                                   RazaoSocial,
                                                   Endereco,
                                                   Bairro,
                                                   Cidade,
                                                   UF,
                                                   CEP,
                                                   INSC,
                                                   CPFCNPJ,
                                                   RG,
                                                   RamoAtividade,
                                                   TipoEstabelecimento,
                                                   Site,
                                                   Referencia,
                                                   Inativo,
                                                   DataCadastro,
                                                   TipoPessoa,
                                                   TipoEntidade) 
                                                    VALUES (@NomeFantasia,
                                                    @RazaoSocial,
                                                    @Endereco,
                                                    @Bairro,
                                                    @Cidade,
                                                    @UF,
                                                    @CEP,
                                                    @INSC,
                                                    @CPFCNPJ,
                                                    @RG,
                                                    @RamoAtividade,
                                                    @TipoEstabelecimento,
                                                    @Site,
                                                    @Referencia,
                                                    0,
                                                    GETDATE(),
                                                    @TipoPessoa,
                                                     @TipoEntidade)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@NomeFantasia", NomeFantasia),
                         New SqlParameter("@RazaoSocial", RazaoSocial),
                         New SqlParameter("@Endereco", Endereco),
                         New SqlParameter("@Bairro", Bairro),
                         New SqlParameter("@Cidade", Cidade),
                         New SqlParameter("@UF", UF),
                         New SqlParameter("@Cep", Cep),
                         New SqlParameter("@INSC", INSC),
                         New SqlParameter("@CPFCNPJ", Cpfcnpj),
                         New SqlParameter("@RG", RG),
                         New SqlParameter("@RamoAtividade", RamoAtividade),
                         New SqlParameter("@TipoEstabelecimento", TipoEstabelecimento),
                         New SqlParameter("@Site", Site),
                         New SqlParameter("@Referencia", Referencia),
                         New SqlParameter("@TipoPessoa", TipoPessoa),
                         New SqlParameter("@TipoEntidade", TipoEntidade)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo atualiza os dados do cliente no banco de dados.
        ''' </summary>
        ''' <param name="Codigo">Representa o código do cliente no banco de dados.</param>
        ''' <param name="NomeFantasia">Representa o nome do cliente do tipo string.</param>
        ''' <param name="Endereco">Representa o endereço do cliente do tipo string</param>
        ''' <param name="Bairro">Representa o bairro do cliente do tipo string.</param>
        ''' <param name="Cep">representa o cep do cliente do tipo string</param>
        ''' <param name="UF">Representa o estado do cliente do tipo string.</param>
        ''' <param name="Cidade">Representa o cidade do cliente do tipo string.</param>
        ''' <param name="Cpfcnpj">Representa o CPF ou CNPJ do cliente do tipo string.</param>
        ''' <param name="RazaoSocial">Representa a razão social do cliente do tipo string</param>
        ''' <param name="TipoPessoa">representa o tipo de pessoa do cliente do tipo string</param>
        ''' <param name="Inativo">Representa o estado do cliente no sistema e banco de dados</param>
        Public Sub AtualizarEntidade(Codigo As Integer, NomeFantasia As String, RazaoSocial As String, Endereco As String, Bairro As String, Cidade As Integer, UF As Integer, Cep As String, INSC As String, Cpfcnpj As String, RG As String, RamoAtividade As String, TipoEstabelecimento As String, Site As String, Referencia As String, Inativo As Integer, TipoPessoa As String, TipoEntidade As String)
            Dim sql As String = "UPDATE Tbl_Entidades    SET  NomeFantasia=@NomeFantasia,
                                                    RazaoSocial = @RazaoSocial, 
                                                    Endereco = @Endereco, 
                                                    Bairro = @Bairro, 
                                                    Cidade = @Cidade, 
                                                    UF = @UF, 
                                                    CEP = @Cep, 
                                                    INSC = @INSC, 
                                                    CPFCNPJ = @CPFCNPJ, 
                                                    RG = @RG,
                                                    RamoAtividade = @RamoAtividade,
                                                    TipoEstabelecimento = @TipoEstabelecimento,
                                                    Site = @Site,
                                                    Referencia = @Referencia,
                                                    Inativo = @Inativo, 
                                                    DataAlteracao = GETDATE(),
                                                    TipoPessoa = @TipoPessoa,
                                                    TipoEntidade = @TipoEntidade
                                              WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@Codigo", Codigo),
                         New SqlParameter("@NomeFantasia", NomeFantasia),
                         New SqlParameter("@RazaoSocial", RazaoSocial),
                         New SqlParameter("@Endereco", Endereco),
                         New SqlParameter("@Bairro", Bairro),
                         New SqlParameter("@Cidade", Cidade),
                         New SqlParameter("@UF", UF),
                         New SqlParameter("@Cep", Cep),
                         New SqlParameter("@INSC", INSC),
                         New SqlParameter("@CPFCNPJ", Cpfcnpj),
                         New SqlParameter("@RG", RG),
                         New SqlParameter("@RamoAtividade", RamoAtividade),
                         New SqlParameter("@TipoEstabelecimento", TipoEstabelecimento),
                         New SqlParameter("@Site", Site),
                         New SqlParameter("@Referencia", Referencia),
                         New SqlParameter("@Inativo", Inativo),
                         New SqlParameter("@TipoPessoa", TipoPessoa),
                         New SqlParameter("@TipoEntidade", TipoEntidade)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo realiza a exclusão do cadastro do cliente do banco de dados.
        ''' </summary>
        ''' <param name="CodEntidade">Código do cliente do tipo integer</param>
        Public Sub ExcluirEntidade(CodEntidade As Integer)
            ExcluirConta(CodEntidade)

            ExcluirContato(CodEntidade)

            ExcluirEndereco(CodEntidade)

            Dim sql As String = "DELETE FROM Tbl_Entidades WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@Codigo", CodEntidade)
        }
            ClasseConexao.Operar(sql, parameters)

            MessageBox.Show("Cadastro excluído com sucesso!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub IncluiContato(CodEntidade As Integer, Contato As String, Email As String, Telefone As String, Tipo As Integer)
            Dim sql As String = "INSERT INTO Tbl_EntidadesContatos (CodEntidade,Contato,Email,Telefone,TipoContato) VALUES (@CodEntidade,@Contato,@Email,@Telefone,@Tipo)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodEntidade", CodEntidade),
                         New SqlParameter("@Contato", Contato),
                         New SqlParameter("@Email", Email),
                         New SqlParameter("@Telefone", Telefone),
                         New SqlParameter("@Tipo", Tipo)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Contato inserido com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaContato(Codigo As Integer, CodEntidade As Integer, Tipo As Integer, Contato As String, Email As String, Telefone As String)
            Dim sql As String = "UPDATE Tbl_EntidadesContatos SET  Contato = @Contato, 
                                                        Email = @Email,
                                                        Telefone = @Telefone,
                                                        TipoContato = @Tipo
                                                WHERE   CodEntidade = @CodEntidade 
                                                AND Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@Codigo", Codigo),
                         New SqlParameter("@CodEntidade", CodEntidade),
                         New SqlParameter("@Contato", Contato),
                         New SqlParameter("@Email", Email),
                         New SqlParameter("@Telefone", Telefone),
                         New SqlParameter("@Tipo", Tipo)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Contato atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluirContato(CodEntidade As Integer, Optional CodContato As Integer? = 0)
            If CodContato <> 0 Then
                Dim sql As String = "DELETE FROM Tbl_EntidadesContatos WHERE CodEntidade = @CodEntidade AND Codigo = @Codigo"
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@Codigo", CodContato),
                         New SqlParameter("@CodEntidade", CodEntidade)
            }
                ClasseConexao.Operar(sql, parameters)
            Else
                Dim sql As String = "DELETE FROM Tbl_EntidadesContatos WHERE CodEntidade = @CodEntidade"
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodEntidade", CodEntidade)
            }
                ClasseConexao.Operar(sql, parameters)
            End If

            MessageBox.Show("Contato excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub IncluiConta(CodEntidade As Integer, Banco As Integer, Agencia As String, Conta As String)
            Dim sql As String = "INSERT INTO Tbl_EntidadesContas (CodEntidade,Banco,Agencia,Conta) VALUES (@CodEntidade,@Banco,@Agencia,@Conta)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodEntidade", CodEntidade),
                         New SqlParameter("@Banco", Banco),
                         New SqlParameter("@Agencia", Agencia),
                         New SqlParameter("@Conta", Conta)
            }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Conta inserida com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaConta(Codigo As Integer, CodEntidade As Integer, ctaBanco As Integer, Agencia As String, Conta As String)
            Dim sql As String = "UPDATE Tbl_EntidadesContas SET Banco = @Banco,Agencia = @Agencia,Conta = @Conta WHERE CodEntidade = @CodEntidade AND Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodEntidade", CodEntidade),
                         New SqlParameter("@Codigo", Codigo),
                         New SqlParameter("@Conta", Conta),
                         New SqlParameter("@Agencia", Agencia),
                         New SqlParameter("@Banco", ctaBanco)
            }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Dados bancários atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluirConta(CodEntidade As Integer, Optional CodConta As Integer? = 0)
            If CodConta <> 0 Then
                Dim sql As String = "DELETE FROM Tbl_EntidadesContas WHERE CodEntidade = @CodEntidade AND Codigo = @Codigo"
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@Codigo", CodConta),
                         New SqlParameter("@CodEntidade", CodEntidade)
            }
                ClasseConexao.Operar(sql, parameters)
            Else
                Dim sql As String = "DELETE FROM Tbl_EntidadesContas WHERE CodEntidade = @CodEntidade"
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodEntidade", CodEntidade)
            }
                ClasseConexao.Operar(sql, parameters)
            End If

            MessageBox.Show("Dados bancários excluidos com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Public Sub IncluirEndereco(CodEntidade As Integer, CEP As String, Endereco As String, Bairro As String, Cidade As Integer, Estado As Integer, Principal As Boolean, TipoEndereco As String)
            Dim sql As String = "INSERT INTO Tbl_EntidadesEndereco (CodEntidade,Cep,Endereco,Bairro,Cidade,Estado,Principal,TipoEndereco,DataCadastro) VALUES (@CodEntidade,@cep,@Endereco,@Bairro,@Cidade,@Estado,@Principal,@TipoEndereco,GETDATE())"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodEntidade", CodEntidade),
                         New SqlParameter("@Cep", CEP),
                         New SqlParameter("@Endereco", Endereco),
                         New SqlParameter("@Bairro", Bairro),
                         New SqlParameter("@Cidade", Cidade),
                         New SqlParameter("@Estado", Estado),
                         New SqlParameter("@Principal", Principal),
                         New SqlParameter("@TipoEndereco", TipoEndereco)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Endereço inserido com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizarEndereco(Codigo As Integer, CodEntidade As Integer, CEP As String, Endereco As String, Bairro As String, Cidade As Integer, Estado As Integer, Principal As Boolean, TipoEndereco As String)
            Dim sql As String = "UPDATE Tbl_EntidadesEndereco SET   Endereco = @Endereco,
                                                                CEP = @Cep,
                                                                Bairro = @Bairro,
                                                                Cidade = @Cidade,
                                                                Estado = @Estado,
                                                                Principal = @Principal,
                                                                TipoEndereco = @TipoEndereco
                                                        WHERE   CodEntidade = @CodEntidade 
                                                        AND     Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@Codigo", Codigo),
                         New SqlParameter("@CodEntidade", CodEntidade),
                         New SqlParameter("@Cep", CEP),
                         New SqlParameter("@Endereco", Endereco),
                         New SqlParameter("@Bairro", Bairro),
                         New SqlParameter("@Cidade", Cidade),
                         New SqlParameter("@Estado", Estado),
                         New SqlParameter("@Principal", Principal),
                         New SqlParameter("@TipoEndereco", TipoEndereco)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Endereço atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluirEndereco(CodEntidade As Integer, Optional CodContato As Integer? = 0)
            If CodContato <> 0 Then
                Dim sql As String = "DELETE FROM Tbl_EntidadesEndereco WHERE CodEntidade = @CodEntidade AND Codigo = @Codigo"
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@Codigo", CodContato),
                         New SqlParameter("@CodEntidade", CodEntidade)
            }
                ClasseConexao.Operar(sql, parameters)
            Else
                Dim sql As String = "DELETE FROM Tbl_EntidadesEndereco WHERE CodEntidade = @CodEntidade"
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodEntidade", CodEntidade)
            }
                ClasseConexao.Operar(sql, parameters)
            End If

            MessageBox.Show("Endereco excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub SalvaTipoContato(TipoContato As String)
            Dim sql As String = "INSERT INTO Tbl_TipoContato (TipoContato) VALUES (@TipoContato)"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@TipoContato", TipoContato)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaTipoContato(Codigo As Integer, TipoContato As String)
            Dim sql As String = "UPDATE Tbl_TipoContato SET TipoContato = @TipoContato WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@Codigo", Codigo),
            New SqlParameter("@TipoContato", TipoContato)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluiTipoContato(Codigo As Integer)
            Dim sql As String = "DELETE FROM Tbl_TipoContato WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@Codigo", Codigo)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo atualiza o saldo de crédito do cliente no banco de dados.
        ''' </summary>
        ''' <param name="CodDevolucao">Representa o código da devolução que gerou o saldo do tipo integer.</param>
        ''' <param name="Cliente">Representa o código do cliente que gerou o saldo do tipo integer</param>
        ''' <param name="Valor">Representa o valor de crédito que foi gerado do tipo decimal.</param>
        ''' <param name="Operacao">Representa o tipo de operação que será realizada, sendo C para crédito e D para débito.</param>
        ''' <param name="CodPedido">Representa o código do pedido de venda que gerou o saldo do tipo integer</param>
        Public Sub AtualizaSaldoCliente(Cliente As Integer, Valor As Decimal, Operacao As String, Optional CodDevolucao As Integer? = Nothing, Optional CodPedido As Integer? = Nothing)
            Dim sql As String = "INSERT INTO tbl_SaldoCliente (CodDevolucao, 
                                                       Cliente,
                                                       Valor, 
                                                       Operacao, 
                                                       DataMovimentacao,
                                                       CodPedido)  
                                                VALUES (@CodDevolucao, 
                                                       @Cliente,
                                                       @Valor, 
                                                       @Operacao, 
                                                        GETDATE(),
                                                        @CodPedido)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@Cliente", Cliente),
                         New SqlParameter("@Valor", Valor),
                         New SqlParameter("@Operacao", Operacao)
        }
            If CodPedido <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodPedido", CodPedido.Value)
            Else
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodPedido", DBNull.Value)
            End If

            If CodDevolucao <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodDevolucao", CodDevolucao.Value)
            Else
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@CodDevolucao", DBNull.Value)
            End If

            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Saldo atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
#End Region
#Region "FUNÇÕES"
        Public Function PesquisaEntidade(Codigo As Integer, NomeFantasia As String, Bairro As String, ctaEstado As String, ctaCidade As String, Cpfcnpj As String, CadastroIni As String, CadastroFim As String, AlteracaoIni As String, AlteracaoFim As String, InativacaoIni As String, InativacaoFim As String, Status As String, TipoEntidade As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_Entidades WHERE 1=1 ")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Ativos"
                        sql.AppendLine("AND Inativo = 0")
                    Case "Inativos"
                        sql.AppendLine("AND Inativo = 1")
                    Case "Todos"
                        sql.AppendLine(" AND Inativo IS NOT NULL")
                End Select

                'Pesquisa por codigo do cliente
                If Codigo <> 0 Then
                    sql.AppendLine("AND CodEntidade = @Codigo")
                    parameters.Add(New SqlParameter("@Codigo", Codigo))
                End If

                'Pesquisa pelo nome
                If Not String.IsNullOrEmpty(NomeFantasia) Then
                    sql.AppendLine("AND Entidade LIKE @NomeFantasia")
                    parameters.Add(New SqlParameter("@NomeFantasia", "%" & NomeFantasia & "%"))
                End If

                'Pesquisa pelo CPF
                If Not String.IsNullOrEmpty(Cpfcnpj) Then
                    sql.AppendLine("AND CPFCNPJ LIKE @Cpfcnpj")
                    parameters.Add(New SqlParameter("@Cpfcnpj", "%" & Cpfcnpj & "%"))
                End If

                'Pesquisa pelo Cidade
                If Not String.IsNullOrEmpty(ctaCidade) Then
                    sql.AppendLine("AND Localidade LIKE @Cidade")
                    parameters.Add(New SqlParameter("@Cidade", "%" & ctaCidade & "%"))
                End If

                'Pesquisa pelo Estado
                If Not String.IsNullOrEmpty(ctaEstado) Then
                    sql.AppendLine("AND UF LIKE @Estado")
                    parameters.Add(New SqlParameter("@Estado", "%" & ctaEstado & "%"))
                End If

                'Pesquisa pelo Bairro
                If Not String.IsNullOrEmpty(Bairro) Then
                    sql.AppendLine("AND BAIRRO LIKE @Bairro")
                    parameters.Add(New SqlParameter("@Bairro", "%" & Bairro & "%"))
                End If

                'Pesquisa pela data de cadastro
                If Not String.IsNullOrEmpty(CadastroIni) And Not String.IsNullOrEmpty(CadastroFim) Then
                    sql.AppendLine("AND DataCadastro BETWEEN @CadastroIni AND @CadastroFim")
                    parameters.Add(New SqlParameter("@CadastroIni", CadastroIni))
                    parameters.Add(New SqlParameter("@CadastroFim", CadastroFim))
                ElseIf Not String.IsNullOrEmpty(CadastroIni) And String.IsNullOrEmpty(CadastroFim) Then
                    sql.AppendLine("AND DataCadastro >= @CadastroIni")
                    parameters.Add(New SqlParameter("@CadastroIni", CadastroIni))
                ElseIf String.IsNullOrEmpty(CadastroIni) And Not String.IsNullOrEmpty(CadastroFim) Then
                    sql.AppendLine("AND DataCadastro <= @CadastroFim")
                    parameters.Add(New SqlParameter("@CadastroFim", CadastroFim))
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

                ' Pesquisa pelo Tipo de Entidade
                If Not String.IsNullOrEmpty(TipoEntidade) Then
                    ' Verifica se o TipoEntidade contém mais de um valor, separado por vírgula
                    Dim tipos = TipoEntidade.Split(","c).Select(Function(tipo) tipo.Trim()).ToArray()

                    ' Adiciona uma condição IN para múltiplos valores
                    If tipos.Length > 1 Then
                        sql.AppendLine("AND Tipo IN (" & String.Join(", ", tipos.Select(Function(t, i) "@TipoEntidade" & i)) & ")")

                        ' Adiciona cada tipo como um parâmetro
                        For i = 0 To tipos.Length - 1
                            parameters.Add(New SqlParameter("@TipoEntidade" & i, tipos(i)))
                        Next
                    Else
                        ' Caso contrário, usa a condição de igualdade padrão
                        sql.AppendLine("AND Tipo = @TipoEntidade")
                        parameters.Add(New SqlParameter("@TipoEntidade", TipoEntidade))
                    End If
                End If

                sql.AppendLine("ORDER BY Entidade")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function
        Public Function ConsultaTipoContato(Sql As String)
            Return ClasseConexao.Consultar(Sql, Nothing)
        End Function
        Public Function ConsultaEntidade(sql As String, Optional CodEntidade As Integer = 0, Optional Entidade As String = Nothing)
            If CodEntidade <> 0 Then
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodEntidade", CodEntidade)
                            }
                Return ClasseConexao.Consultar(sql, parameters)
            ElseIf Entidade IsNot Nothing Then
                Dim parameters As SqlParameter() = {
                       New SqlParameter("@Entidade", "%" & Entidade & "%")
                          }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function

        ''' <summary>
        ''' Esta função valida CPF/CPNJ, Situação e vinculos da entidades no sistema, para permitir exclusão de ativação/inativação da entidades.
        ''' </summary>
        ''' <param name="Operacao">Representa a operação que deve ser validade, sendo as possíveis (Situação, CPF/CNPJ, Exclusão.</param>
        ''' <param name="CodEntidade">Represente o código da entidade do tipo Integer.</param>
        ''' <param name="Situacao">Representa a situação atual da entidade no sistema</param>
        ''' <param name="CPFCNPJ">Representa o número do CPF/CNPJ da entidade.</param>
        ''' <returns></returns>
        Public Function ValidaEntidade(Operacao As String, CodEntidade As Integer, Situacao As Integer, Optional ByVal CPFCNPJ As String = "") As Boolean
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodEntidade", CodEntidade)
                            }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_EntidadeContato WHERE CodEntidade = @CodEntidade", parameters)
            Select Case Operacao
                Case "SITUACAO"
                    If Situacao = 1 And Tabela.Rows(0)("Inativo").ToString() = 0 Then
                        Dim sql As String = "UPDATE Tbl_Entidades SETInativo = 1 WHERE Codigo = @CODIGO"
                        Dim parametersinativa As SqlParameter() = {
                        New SqlParameter("@Codigo", CodEntidade)
                    }
                        ClasseConexao.Operar(sql, parametersinativa)
                        MessageBox.Show("Cliente inativado com sucesso!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return True
                    ElseIf Situacao = 0 And Tabela.Rows(0)("Inativo").ToString() = 1 Then
                        Dim sql As String = "UPDATE Tbl_Entidades SETInativo = 0 WHERE Codigo = @CODIGO"
                        Dim parametersinativa As SqlParameter() = {
                        New SqlParameter("@Codigo", CodEntidade)
                    }
                        ClasseConexao.Operar(sql, parametersinativa)
                        MessageBox.Show("Cliente ativado com sucesso!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return True
                    ElseIf Situacao = Tabela.Rows(0)("Inativo").ToString() Then
                        Return False
                    End If
                Case "CPFCNPJ"
                    Dim tbConfig As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_ConfigSistema", Nothing)
                    If tbConfig IsNot Nothing AndAlso tbConfig.Rows.Count > 0 Then
                        Dim ValidarDocumento As Boolean = tbConfig.Rows(0)("ValidaCPFCNPJ").ToString()
                        Dim DocumentoExiste As Boolean = tbConfig.Rows(0)("PermiteMesmoCNPJ").ToString()

                        If ValidarDocumento = True Then
                            If CPFCNPJ.Length = 11 Then
                                If ValidarCPF(CPFCNPJ) Then
                                    If DocumentoExiste = True Then
                                        Return True
                                        Exit Function
                                    Else
                                        Dim parametercpfcnpj As SqlParameter() = {
                                           New SqlParameter("@CPFCNPJ", CPFCNPJ)
                                           }
                                        Dim tbsituacao As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_Entidades WHERE CPFCNPJ LIKE @CPFCNPJ", parametercpfcnpj)
                                        If tbsituacao.Rows.Count() = 1 Then
                                            MessageBox.Show("CPF informado já cadastrado no banco de dados!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                            Return False
                                        Else
                                            Return True
                                        End If
                                    End If
                                Else
                                    MessageBox.Show("CPF inválido!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                    Return False
                                End If
                            ElseIf CPFCNPJ.Length = 14 Then
                                If ValidarCNPJ(CPFCNPJ) Then
                                    If DocumentoExiste = True Then
                                        Return True
                                        Exit Function
                                    Else
                                        Dim parametercpfcnpj As SqlParameter() = {
                                           New SqlParameter("@CPFCNPJ", CPFCNPJ)
                                           }
                                        Dim tbsituacao As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_Entidades WHERE CPFCNPJ LIKE @CPFCNPJ", parametercpfcnpj)
                                        If tbsituacao.Rows.Count() = 1 Then
                                            MessageBox.Show("CNPJ informado já cadastrado no banco de dados!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                            Return False
                                        Else
                                            Return True
                                        End If
                                    End If
                                Else
                                    MessageBox.Show("CPNJ inválido!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                    Return False
                                End If
                            Else
                                MessageBox.Show("Tipo de pessoa inválido!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Return False
                            End If
                        Else
                            If DocumentoExiste = True Then
                                Return True
                                Exit Function
                            Else
                                Dim parametercpfcnpj As SqlParameter() = {
                                           New SqlParameter("@CPFCNPJ", CPFCNPJ)
                                           }
                                Dim tbsituacao As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_Entidades WHERE CPFCNPJ LIKE @CPFCNPJ", parametercpfcnpj)
                                If tbsituacao.Rows.Count() = 1 Then
                                    MessageBox.Show("CNPJ informado já cadastrado no banco de dados!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                    Return False
                                Else
                                    Return True
                                End If
                            End If
                        End If
                    Else
                        MessageBox.Show("Não é possível validar o documento da etidade!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Function
                    End If
            End Select
        End Function

#End Region
    End Class
End Namespace

Namespace Classes.Entidades.Locadora
    Public Class clsEntidades
        Dim ClasseConexao As New ConexaoSQLServer
#Region "CONSTRUTORES"

#End Region
#Region "PROPRIEDADES"
        Private Property _CodEntidade As Integer
        Public Property CodEntidade As Integer
            Get
                Return _CodEntidade
            End Get
            Set(value As Integer)
                _CodEntidade = value
            End Set
        End Property
        Private Property _CodContato As Integer
        Public Property CodContato As Integer
            Get
                Return _CodContato
            End Get
            Set(value As Integer)
                _CodContato = value
            End Set
        End Property
        Private Property _CodCargo As Integer
        Public Property CodCargo As Integer
            Get
                Return _CodCargo
            End Get
            Set(value As Integer)
                _CodCargo = value
            End Set
        End Property
        Private Property _NumeroFicha As Integer
        Public Property NumeroFicha As Integer
            Get
                Return _NumeroFicha
            End Get
            Set(value As Integer)
                _NumeroFicha = value
            End Set
        End Property
        Private Property _NomeFantasia As String
        Public Property NomeFantasia As String
            Get
                Return _NomeFantasia
            End Get
            Set(value As String)
                _NomeFantasia = value
            End Set
        End Property
        Private Property _Telefone1 As String
        Public Property Telefone1 As String
            Get
                Return _Telefone1
            End Get
            Set(value As String)
                _Telefone1 = value
            End Set
        End Property
        Private Property _Telefone2 As String
        Public Property Telefone2 As String
            Get
                Return _Telefone2
            End Get
            Set(value As String)
                _Telefone2 = value
            End Set
        End Property
        Private Property _Celular As String
        Public Property Celular As String
            Get
                Return _Celular
            End Get
            Set(value As String)
                _Celular = value
            End Set
        End Property
        Private Property _Email As String
        Public Property Email As String
            Get
                Return _Email
            End Get
            Set(value As String)
                _Email = value
            End Set
        End Property
        Private Property _Matricula As Integer
        Public Property Matricula As Integer
            Get
                Return _Matricula
            End Get
            Set(value As Integer)
                _Matricula = value
            End Set
        End Property
        Private Property _CarteiraProfissional As String
        Public Property CarteiraProfissional As String
            Get
                Return _CarteiraProfissional
            End Get
            Set(value As String)
                _CarteiraProfissional = value
            End Set
        End Property
        Private Property _Cargo As Integer
        Public Property Cargo As Integer
            Get
                Return _Cargo
            End Get
            Set(value As Integer)
                _Cargo = value
            End Set
        End Property
        Private Property _Salario As Decimal
        Public Property Salario As Decimal
            Get
                Return _Salario
            End Get
            Set(value As Decimal)
                _Salario = value
            End Set
        End Property
        Private Property _Expediente As String
        Public Property Expediente As String
            Get
                Return _Expediente
            End Get
            Set(value As String)
                _Expediente = value
            End Set
        End Property

#End Region
#Region "METODOS"
        Public Sub SalvarEntidade(NomeFantasia As String, RazaoSocial As String, Datanasc As String, Estadocivil As String, Endereco As String, Complemento As String, Bairro As String, Cidade As String, Uf As String, Cep As String, sexo As String, Rg As String, Documento As String, Obs As String, Tipo As String)

            Dim sql As String = "INSERT INTO tbEntidades   (NomeFantasia,
                                                           RazaoSocial,
                                                           Datanasc,
                                                           Estadocivil,
                                                           Endereco,
                                                           Complemento,
                                                           Bairro,
                                                           Cidade,
                                                           Uf,
                                                           Cep,
                                                           sexo,
                                                           Rg,
                                                           Documento,
                                                           Obs,
                                                           DataCadastro,
                                                           Tipo)
                                                VALUES     (@NomeFantasia,
                                                           @RazaoSocial,
                                                           @Datanasc,
                                                           @Estadocivil,
                                                           @Endereco,
                                                           @Complemento,
                                                           @Bairro,
                                                           @Cidade,
                                                           @Uf,
                                                           @Cep,
                                                           @sexo,
                                                           @Rg,
                                                           @Documento,
                                                           @Obs,
                                                           GETDATE(),
                                                           @Tipo)"

            Dim parameters As SqlParameter() = {
                    New SqlParameter("@NomeFantasia", NomeFantasia),
                    New SqlParameter("@RazaoSocial", RazaoSocial),
                    New SqlParameter("@Datanasc", Datanasc),
                    New SqlParameter("@Estadocivil", Estadocivil),
                    New SqlParameter("@Endereco", Endereco),
                    New SqlParameter("@Complemento", Complemento),
                    New SqlParameter("@Bairro", Bairro),
                    New SqlParameter("@Cidade", Cidade),
                    New SqlParameter("@Uf", Uf),
                    New SqlParameter("@Cep", Cep),
                    New SqlParameter("@sexo", sexo),
                    New SqlParameter("@Rg", Rg),
                    New SqlParameter("@Documento", Documento),
                    New SqlParameter("@Obs", Obs),
                    New SqlParameter("@Tipo", Tipo)
                  }
            ClasseConexao.Operar(sql, parameters)

            MessageBox.Show("Entidade salva com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AlterarEntidade(codigo As Integer, NomeFantasia As String, RazaoSocial As String, Datanasc As String, Estadocivil As String, Endereco As String, Complemento As String, Bairro As String, Cidade As String, Uf As String, Cep As String, sexo As String, Rg As String, Documento As String, Obs As String)

            Dim sql As String = "UPDATE	tbEntidades
                               SET	NomeFantasia = @NomeFantasia,
		                            RazaoSocial = @RazaoSocial,
		                            Datanasc = @Datanasc,
		                            Estadocivil = @Estadocivil,
		                            Endereco = @Endereco,
		                            Complemento = @Complemento,
		                            Bairro = @Bairro,
		                            Cidade = @Cidade,
		                            Uf = @Uf,
		                            CEP = @Cep,
		                            sexo = @sexo,
		                            Rg = @Rg,
		                            Documento = @Documento,
		                            Obs = @Obs,
		                            DataAlteracao = GETDATE()
                             WHERE	Codigo = @Codigo"

            Dim parameters As SqlParameter() = {
                    New SqlParameter("@Codigo", codigo),
                    New SqlParameter("@NomeFantasia", NomeFantasia),
                    New SqlParameter("@RazaoSocial", RazaoSocial),
                    New SqlParameter("@Datanasc", Datanasc),
                    New SqlParameter("@Estadocivil", Estadocivil),
                    New SqlParameter("@Endereco", Endereco),
                    New SqlParameter("@Complemento", Complemento),
                    New SqlParameter("@Bairro", Bairro),
                    New SqlParameter("@Cidade", Cidade),
                    New SqlParameter("@Uf", Uf),
                    New SqlParameter("@Cep", Cep),
                    New SqlParameter("@sexo", sexo),
                    New SqlParameter("@Rg", Rg),
                    New SqlParameter("@Documento", Documento),
                    New SqlParameter("@Obs", Obs)
                  }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Entidade alterada com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub
        Public Sub ExcluirEntidade(codigo As Integer)

            Dim sql As String = "DELETE FROM tbEntidades WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
                      New SqlParameter("@Codigo", codigo)
                      }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Entidade excluída com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Public Sub SalvarContato(CodEntidade As Integer, Telefone1 As String, Telefone2 As String, Celular As String, Email As String)

            Dim sql As String = "INSERT INTO    tbEntidadeContatos
                                                (CodEntidade,
                                                Telefone1,
                                                Telefone2,
                                                Celular,
                                                Email)
                                        VALUES
                                                (@CodEntidade,
                                                @Telefone1,
                                                @Telefone2,
                                                @Celular,
                                                @Email)"
            Dim parameters As SqlParameter() = {
                        New SqlParameter("@CodEntidade", CodEntidade),
                        New SqlParameter("@Telefone1", Telefone1),
                        New SqlParameter("@Telefone2", Telefone2),
                        New SqlParameter("@Celular", Celular),
                        New SqlParameter("@Email", Email)
                        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Contato salva com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub
        Public Sub AlterarContato(CodEntidade As Integer, Codigo As Integer, Telefone1 As String, Telefone2 As String, Celular As String, Email As String)

            Dim sql As String = "UPDATE	tbEntidadeContatos
                                SET     CodEntidade = @CodEntidade, 
                                        Telefone1 = @Telefone1,
                                        Telefone2 = @Telefone2,
                                        Celular = @Celular,
                                        Email = @Email
                                WHERE   Codigo = @Codigo AND CodEntidade = @CodEntidade"

            Dim parameters As SqlParameter() = {
                        New SqlParameter("@CodEntidade", CodEntidade),
                        New SqlParameter("@Codigo", Codigo),
                        New SqlParameter("@Telefone1", Telefone1),
                        New SqlParameter("@Telefone2", Telefone2),
                        New SqlParameter("@Celular", Celular),
                        New SqlParameter("@Email", Email)
                        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Contato alterado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub
        Public Sub ExcluirContato(CodEntidade As Integer, Optional Codcontato As Integer = 0)

            If Codcontato <> 0 Then
                Dim sql As String = "DELETE FROM tbEntidadeContatos WHERE CodEntidade = @CodEntidade AND Codigo = @Codigo"
                Dim parameters As SqlParameter() = {
                        New SqlParameter("@CodEntidade", CodEntidade),
                        New SqlParameter("@Codigo", Codcontato)
                        }
                ClasseConexao.Operar(sql, parameters)
                MessageBox.Show("Contato excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                Dim sql As String = "DELETE FROM tbEntidadeContatos WHERE CodEntidade = @CodEntidade"
                Dim parameters As SqlParameter() = {
                        New SqlParameter("@CodEntidade", CodEntidade),
                        New SqlParameter("@Codigo", Codcontato)
                        }
                ClasseConexao.Operar(sql, parameters)
                MessageBox.Show("Contato excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Public Sub SalvarCargo(CodEntidade As Integer, Matricula As Integer, CarteiraProfissional As String, Cargo As String, Salario As String, Expediente As String)

            Dim sql As String = "INSERT INTO tbEntidadeCargo   (CodEntidade,
                                                                   Matricula,
                                                                   CarteiraProfissional,
                                                                   Cargo,
                                                                   Salario,
                                                                   Expediente,
                                                                   GETDATE())
                                        VALUES                     (@CodEntidade,
                                                                   @Matricula,
                                                                   @CarteiraProfissional,
                                                                   @Cargo,
                                                                   @Salario,
                                                                   @Expediente,
                                                                   @DataCadastro)"
            Dim parameters As SqlParameter() = {
                        New SqlParameter("@CodEntidade", CodEntidade),
                        New SqlParameter("@Matricula", Matricula),
                        New SqlParameter("@CarteiraProfissional", CarteiraProfissional),
                        New SqlParameter("@Cargo", Cargo),
                        New SqlParameter("@Salario", Salario),
                        New SqlParameter("@Expediente", Expediente)
                        }
            ClasseConexao.Operar(sql, parameters)

            MessageBox.Show("Cargo salva com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AlterarCargo(CodEntidade As Integer, Codigo As Integer, Matricula As Integer, CarteiraProfissional As String, Cargo As String, Salario As String, Expediente As String)

            Dim sql As String = "UPDATE	tbEntidadeCargo
                                    SET     CodEntidade = @CodEntidade, 
                                            Matricula = @Matricula,
                                            CarteiraProfissional = @CarteiraProfissional,
                                            Cargo = @Cargo,
                                            Salario = @Salario,
                                            Expediente = @Expediente,
                                            DataAlteracao = GETDATE()
                                WHERE   Codigo = @Codigo AND CodEntidade = @CodEntidade"
            Dim parameters As SqlParameter() = {
                      New SqlParameter("@CodEntidade", CodEntidade),
                      New SqlParameter("@Matricula", Matricula),
                      New SqlParameter("@CarteiraProfissional", CarteiraProfissional),
                      New SqlParameter("@Cargo", Cargo),
                      New SqlParameter("@Salario", Salario),
                      New SqlParameter("@Expediente", Expediente)
                      }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cargo alterado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Sub
        Public Sub ExcluirCargo(CodEntidade As Integer, Optional CodCargo As Integer = 0)
            If CodCargo <> 0 Then
                Dim sql As String = "DELETE FROM tbEntidadeCargo WHERE CodEntidade = @CodEntidade AND Codigo = @CodCargo"

                Dim parameters As SqlParameter() = {
                          New SqlParameter("@CodEntidade", CodEntidade),
                          New SqlParameter("@CodCargo", CodCargo)
                          }
                ClasseConexao.Operar(sql, parameters)
                MessageBox.Show("Cargo excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                Dim sql As String = "DELETE FROM tbEntidadeCargo WHERE CodEntidade = @CodEntidade"

                Dim parameters As SqlParameter() = {
                      New SqlParameter("@CodEntidade", CodEntidade)
                      }
                ClasseConexao.Operar(sql, parameters)
                MessageBox.Show("Cargo excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub
#End Region
#Region "FUNCOES"
        Public Function PesquisaEntidade(Codigo As Integer, NomeFantasia As String, Tipo As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM tbEntidades WHERE 1=1")
            Dim parameters As New List(Of SqlParameter)()

            Try
                If Codigo <> 0 Then
                    sql.AppendLine("AND Codigo = @Codigo")
                    parameters.Add(New SqlParameter("@Codigo", Codigo))
                End If

                If Not String.IsNullOrEmpty(NomeFantasia) Then
                    sql.AppendLine("AND NomeFantasia LIKE @NomeFantasia")
                    parameters.Add(New SqlParameter("@Nome", "%" & NomeFantasia & "%"))
                End If

                If Not String.IsNullOrEmpty(Tipo) Then
                    sql.AppendLine("AND Tipo = @Tipo")
                    parameters.Add(New SqlParameter("@Tipo", Tipo))
                End If

                sql.AppendLine("ORDER BY NomeFantasia")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Erro ao consultar a entidades: " & ex.Message)
            End Try
        End Function
        Public Function ConsultaEntidade(sql As String, CodEntidade As Integer)
            Dim parameters As SqlParameter() = {
                       New SqlParameter("@CodEntidade", CodEntidade)
                          }
            Return ClasseConexao.Consultar(sql, parameters)
        End Function
#End Region
    End Class
End Namespace