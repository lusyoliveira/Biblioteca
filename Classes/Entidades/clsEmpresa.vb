Imports System.Data.SqlClient
Imports Biblioteca.Classes.Conexao
Imports Biblioteca.Classes.Entidades.Vendas
Public Class clsEmpresa
    Inherits clsEntidades
  Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
    Public Property CodLoja As Integer
    Public Property CodEstabelecimento As Integer
    Public Property CodRamo As Integer
    Public Property TipoEstabelecimento As String
    Public Property RamoAtividade As String
    Public Property Loja As String
    Public Property CodEmpresa As Integer
    Public Property DataInauguracao As Date
    Public Property DataCadastro As Date
    Public Property DataAlteracao As Date
    Public Property DataAtivacao As String
    Public Property Titular As String
    Public Property Debito As Decimal
    Public Property Credito As Decimal
    Public Property Boleto As Decimal
    Public Property Taxas As Decimal
    Public Property Comissao As Integer
    Public Property Lucro As Integer
    Public Property Pix As Decimal
    Public Property Inativo As Boolean
    Public Property CustoRevenda As Decimal
    Public Property CustoProducao As Decimal
    Public Property CustoFixo As Decimal
    Public Property CodEmail As Integer
    Public Property Senha As String
    Public Property FaturamentoTotal As Decimal
    Public Property FaturamentoMedio As Decimal
    Public Property LucroTotal As Decimal
    Public Property VendasTotal As Integer
    Public Property Orcamentos As Integer
    Public Property Vendas As Integer
    Public Property Devolucoes As Integer
    Public Property Cotacoes As Integer
    Public Property PedidosCompra As Integer
    Public Property NotasEntrada As Integer
    Public Property ContasPagar As Integer
    Public Property ContasReceber As Integer

#End Region
#Region "METODOS"
    ''' <summary>
    ''' Este metódo registra um nova empresa no banco de dados.
    ''' </summary>
    ''' <param name="NomeFantasia">Representa o nome fantasia da empresa.</param>
    ''' <param name="RazaoSocial">Representa a razão social da empresa.</param>
    ''' <param name="DataAtivacao">Representa da data de ativação da empresa.</param>
    ''' <param name="CNPJ">Representa o CNPJ da empresa.</param>
    ''' <param name="Titular">Representa o nome do titular de empresa.</param>
    ''' <param name="RamoAtividade">Representa o ramo de atividade da empresa.</param>
    ''' <param name="TipoEstabelecimento">Representa o tipo de estabelecimento da empresa.</param>
    ''' <param name="Telefone">Representa o telefone da empresa.</param>
    ''' <param name="Site">Representa o site da empresa.</param>
    ''' <param name="CEP">Representa o CEP da empresa.</param>
    ''' <param name="Endereco">Representa o endereço da empresa.</param>
    ''' <param name="Bairro">Representa o bairro da empresa</param>
    ''' <param name="Municipio">Representa o município da empresa</param>
    ''' <param name="Estado">Representa o estado da empresa.</param>
    Public Sub SalvarEmpresa(NomeFantasia As String, RazaoSocial As String, DataAtivacao As Date, CNPJ As String, Titular As String, RamoAtividade As Integer, TipoEstabelecimento As Integer, Telefone As String, Site As String, CEP As String, Endereco As String, Bairro As String, Municipio As Integer, Estado As Integer)
        Dim sql As String = "INSERT INTO  Tbl_Empresa             (NomeFantasia,
                                                                RazaoSocial,
                                                                DataAtivacao,
                                                                CNPJ,
                                                                Titular,
                                                                RamoAtividade,
                                                                TipoEstabelecimento,
                                                                Telefone,
                                                                Site,
                                                                CEP,
                                                                Endereco,
                                                                Bairro,
                                                                Cidade,
                                                                Estado,
                                                                Inativo,
                                                                DataCadastro) 
                                          VALUES                (@NomeFantasia,
                                                                @RazaoSocial,
                                                                @DataAtivacao, 
                                                                @CPFCNPJ, 
                                                                @Titular, 
                                                                @RamoAtividade, 
                                                                @TipoEstabelecimento,
                                                                @Telefone, 
                                                                @Site, 
                                                                @CEP, 
                                                                @Endereco, 
                                                                @Bairro,
                                                                @Municipio, 
                                                                @Estado, 
                                                                0,
                                                                GETDATE())"
        Dim parameters As SqlParameter() = {
        New SqlParameter("@NomeFantasia", NomeFantasia),
        New SqlParameter("@RazaoSocial", RazaoSocial),
        New SqlParameter("@DataAtivacao", DataAtivacao),
        New SqlParameter("@CPFCNPJ", CNPJ),
        New SqlParameter("@Titular", Titular),
        New SqlParameter("@RamoAtividade", RamoAtividade),
        New SqlParameter("@TipoEstabelecimento", TipoEstabelecimento),
        New SqlParameter("@Telefone", Telefone),
        New SqlParameter("@Site", Site),
        New SqlParameter("@CEP", CEP),
        New SqlParameter("@Endereco", Endereco),
        New SqlParameter("@Bairro", Bairro),
        New SqlParameter("@Municipio", Municipio),
        New SqlParameter("@Estado", Estado)
   }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    ''' <summary>
    ''' Este metódo atualizado os dados de uma empresa no banco de dados.
    ''' </summary>
    ''' <param name="CodEmpresa">Representa o código da empresa no banco de dados.</param>
    ''' <param name="NomeFantasia">Representa o nome fantasia da empresa.</param>
    ''' <param name="RazaoSocial">Representa a razão social da empresa.</param>
    ''' <param name="DataAtivacao">Representa da data de ativação da empresa.</param>
    ''' <param name="CNPJ">Representa o CNPJ da empresa.</param>
    ''' <param name="Titular">Representa o nome do titular de empresa.</param>
    ''' <param name="RamoAtividade">Representa o ramo de atividade da empresa.</param>
    ''' <param name="TipoEstabelecimento">Representa o tipo de estabelecimento da empresa.</param>
    ''' <param name="Telefone">Representa o telefone da empresa.</param>
    ''' <param name="Site">Representa o site da empresa.</param>
    ''' <param name="CEP">Representa o CEP da empresa.</param>
    ''' <param name="Endereco">Representa o endereço da empresa.</param>
    ''' <param name="Bairro">Representa o bairro da empresa</param>
    ''' <param name="Municipio">Representa o município da empresa</param>
    ''' <param name="Estado">Representa o estado da empresa.</param>
    Public Sub AtualizaEmpresa(CodEmpresa As Integer, NomeFantasia As String, RazaoSocial As String, DataAtivacao As Date, CNPJ As String, Titular As String, RamoAtividade As Integer, TipoEstabelecimento As Integer, Telefone As String, Site As String, CEP As String, Endereco As String, Bairro As String, Municipio As Integer, Estado As Integer, Inativo As Boolean)
        Dim sql As String = "UPDATE Tbl_Empresa SET       NomeFantasia = @NomeFantasia,
                                                        RazaoSocial = @RazaoSocial,
                                                        DataAtivacao  = @DataAtivacao, 
                                                        CNPJ = @CPFCNPJ, 
                                                        Titular = @Titular, 
                                                        RamoAtividade = @RamoAtividade, 
                                                        TipoEstabelecimento = @TipoEstabelecimento, 
                                                        Telefone = @Telefone, 
                                                        Site =  @Site, 
                                                        CEP = @CEP, 
                                                        Endereco = @Endereco, 
                                                        Bairro = @Bairro,
                                                        Cidade = @Municipio, 
                                                        Estado = @Estado, 
                                                        Inativo = @Inativo,
                                                        DataAlteracao = GETDATE()
                                               WHERE    Codigo = @CodEmpresa"
        Dim parameters As SqlParameter() = {
        New SqlParameter("@CodEmpresa", CodEmpresa),
        New SqlParameter("@NomeFantasia", NomeFantasia),
        New SqlParameter("@RazaoSocial", RazaoSocial),
        New SqlParameter("@DataAtivacao", DataAtivacao),
        New SqlParameter("@CPFCNPJ", CNPJ),
        New SqlParameter("@Titular", Titular),
        New SqlParameter("@RamoAtividade", RamoAtividade),
        New SqlParameter("@TipoEstabelecimento", TipoEstabelecimento),
        New SqlParameter("@Telefone", Telefone),
        New SqlParameter("@Site", Site),
        New SqlParameter("@Inativo", Inativo),
        New SqlParameter("@CEP", CEP),
        New SqlParameter("@Endereco", Endereco),
        New SqlParameter("@Bairro", Bairro),
        New SqlParameter("@Municipio", Municipio),
        New SqlParameter("@Estado", Estado)
   }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    ''' <summary>
    ''' Este metódo realiza a exclusão de uma empresa no banco de dados.
    ''' </summary>
    ''' <param name="CodLoja">Representa o código da empresa no banco de dados.</param>
    Public Sub ExcluirEmpresa(CodLoja As Integer)
        ExcluirTaxa(CodLoja)
        ExcluirEmail(CodLoja)
        Dim sql As String = "DELETE FROM Tbl_Empresa WHERE Codigo = @CodLoja"
        Dim parameters As SqlParameter() = {
        New SqlParameter("@CodLoja", CodLoja)
   }
        ClasseConexao.Operar(sql, parameters)

        MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    ''' <summary>
    ''' Este método registra as taxa de uma empresa no banco de dados.
    ''' </summary>
    ''' <param name="Debito">Representa o valor da taxa de cartão de débito da empresa.</param>
    ''' <param name="Credito">Representa o valor da taxa de cartão de crédito da empresa.</param>
    ''' <param name="Boleto">Representa o valor da taxa de boleto da empresa.</param>
    ''' <param name="Pix">Representa o valor da taxa de pix da empresa.</param>
    ''' <param name="Comissao">Representa o valor da taxa de comissão da empresa.</param>
    ''' <param name="CodEmpresa">Represente o código da empresa no banco de dados.</param>
    Public Sub SalvarTaxa(Debito As Integer, Credito As Integer, Boleto As Decimal, Pix As Integer, Comissao As Integer, CodEmpresa As Integer)
        Dim sql As String = "INSERT INTO Tbl_EmpresasTaxas (Debito,Credito,Boleto,Pix,Comissao,Taxas,CodEmpresa) VALUES (@Debito,@Credito,@Boleto,@Pix,@Comissao,@CodEmpresa)"
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@Debito", Debito),
                     New SqlParameter("@Credito", Credito),
                     New SqlParameter("@Boleto", Boleto),
                     New SqlParameter("@Pix", Pix),
                     New SqlParameter("@Comissao", Comissao),
                     New SqlParameter("@CodEmpresa", CodEmpresa)
                     }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    ''' <summary>
    ''' Este método atualiza os dados das taxa de uma empresa no banco de dados.
    ''' </summary>
    ''' <param name="Debito">Representa o valor da taxa de cartão de débito da empresa.</param>
    ''' <param name="Credito">Representa o valor da taxa de cartão de crédito da empresa.</param>
    ''' <param name="Boleto">Representa o valor da taxa de boleto da empresa.</param>
    ''' <param name="Pix">Representa o valor da taxa de pix da empresa.</param>
    ''' <param name="Comissao">Representa o valor da taxa de comissão da empresa.</param>
    ''' <param name="CodEmpresa">Represente o código da empresa no banco de dados.</param>
    ''' <param name="CodLoja">Represente o código da loja no banco de dados.</param>
    Public Sub AtualizaTaxa(CodLoja As Integer, Debito As Integer, Credito As Integer, Taxas As Decimal, Boleto As Decimal, Pix As Integer, Comissao As Integer, CodEmpresa As Integer)
        Dim sql As String = "UPDATE Tbl_EmpresasTaxas SET Debito = @Debito,
                                                        Credito = @Credito,
                                                        Boleto = @Boleto,
                                                        Pix = @Pix,
                                                        Comissao = @Comissao,
                                                        CodEmpresa = @CodEmpresa    
                                                        WHERE Codigo = @CodLoja"
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodLoja", CodLoja),
                     New SqlParameter("@Debito", Debito),
                     New SqlParameter("@Credito", Credito),
                     New SqlParameter("@Boleto", Boleto),
                     New SqlParameter("@Pix", Pix),
                     New SqlParameter("@Comissao", Comissao),
                     New SqlParameter("@Taxas", Taxas),
                     New SqlParameter("@CodEmpresa", CodEmpresa)
                     }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    ''' <summary>
    ''' Este método realiza a exclusão das taxas de uma empresa no banco de dados.
    ''' </summary>
    ''' <param name="CodLoja">Represente o código da loja no banco de dados.</param>
    Public Sub ExcluirTaxa(CodLoja As Integer)
        Dim sql As String = "DELETE FROM Tbl_EmpresasTaxas WHERE Codigo = @CodLoja"
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodLoja", CodLoja)
                     }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>
    ''' Este metódo registra um conta de email da empresa no banco de dados.
    ''' </summary>
    ''' <param name="SMTP">Represente o servidor SMTP do endereço de email.</param>
    ''' <param name="POP">Represente o servidor POP do endereço de email.</param>
    ''' <param name="IMAP">Represente o servidor IMAP do endereço de email.</param>
    ''' <param name="Email">Represente o endereço de e-mail.</param>
    ''' <param name="Senha">Representa sa senha do endereço de email.</param>
    Public Sub SalvarEmail(SMTP As String, POP As String, IMAP As String, Email As String, Senha As String)
        Dim sql As String = "INSERT INTO tbl_EmpresaEmail (SMTP,POP,IMAP,Email,Senha,CodEmpresa) VALUES (@SMTP,@POP,@IMAP,@Email,@Senha,2)"
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@SMTP", SMTP),
                     New SqlParameter("@POP", POP),
                     New SqlParameter("@IMAP", IMAP),
                     New SqlParameter("@Email", Email),
                     New SqlParameter("@Senha", Senha)
                     }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    ''' <summary>
    ''' Este metódo atualiza os dados de uma conta de email da empresa no banco de dados.
    ''' </summary>
    ''' <param name="Codigo">Representa o código identificador do e-mail no banco de dados</param>
    ''' <param name="SMTP">Represente o servidor SMTP do endereço de email.</param>
    ''' <param name="POP">Represente o servidor POP do endereço de email.</param>
    ''' <param name="IMAP">Represente o servidor IMAP do endereço de email.</param>
    ''' <param name="Email">Represente o endereço de e-mail.</param>
    ''' <param name="Senha">Representa sa senha do endereço de email.</param>
    Public Sub AtualizarEmail(Codigo As Integer, SMTP As String, POP As String, IMAP As String, Email As String, Senha As String)
        Dim sql As String = "UPDATE tbl_EmpresaEmail SET SMTP = @SMTP,POP = @POP, IMAP =@IMAP, Email = @Email, Senha = @Senha,TipoEmail = @TipoEmail, CodEmpresa = 2 WHERE Codigo = @Codigo"
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@Codigo", Codigo),
                     New SqlParameter("@SMTP", SMTP),
                     New SqlParameter("@POP", POP),
                     New SqlParameter("@IMAP", IMAP),
                     New SqlParameter("@Email", Email),
                     New SqlParameter("@Senha", Senha)
                     }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    ''' <summary>
    ''' Este metódo realiza a exclusão de uma conta de email da empresa no banco de dados.
    ''' </summary>
    ''' <param name="CodEmpresa">Representa o código da empresa que o e-mail pertence.</param>
    Public Sub ExcluirEmail(CodEmpresa As Integer)

        Dim sql As String = "DELETE FROM tbl_EmpresaEmail WHERE CodEmpresa = @CodEmpresa"
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodEmpresa", CodEmpresa)
                     }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    ''' <summary>
    ''' Este metódo registra a conta da empresa no banco de dados.
    ''' </summary>
    ''' <param name="NomeConta">Representa o nome da conta.</param>
    ''' <param name="Conta">Representa o número da conta.</param>
    ''' <param name="Agencia">Representa o número da agência</param>
    ''' <param name="Banco">Representa o nome do banco</param>
    ''' <param name="CodEmpresa">Representa o código da empresa que a conta pertence.</param>
    Public Sub SalvarContaBancaria(NomeConta As String, Conta As String, Agencia As Integer, Banco As Integer, CodEmpresa As Integer)
        Dim sql As String = "INSERT INTO Tbl_EmpresaContas (NomeConta,Conta,Agencia,Banco) VALUES (@NOMECONTA,@CONTA,@AGENCIA,@BANCO)"
        Dim parameters As SqlParameter() = {
          New SqlParameter("@NOMECONTA", NomeConta),
          New SqlParameter("@CONTA", Conta),
          New SqlParameter("@AGENCIA", Agencia),
          New SqlParameter("@BANCO", Banco),
          New SqlParameter("@CodEmpresa", CodEmpresa)
    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro realizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    ''' <summary>
    ''' Este metódo atualiza os dados da conta da empresa no banco de dados.
    ''' </summary>
    ''' <param name="NomeConta">Representa o nome da conta.</param>
    ''' <param name="Conta">Representa o número da conta.</param>
    ''' <param name="Agencia">Representa o número da agência</param>
    ''' <param name="Banco">Representa o nome do banco</param>
    ''' <param name="CodEmpresa">Representa o código da empresa que a conta pertence.</param>
    Public Sub AtualizarContaBancaria(CodConta As Integer, NomeConta As String, Conta As String, Agencia As Integer, Banco As Integer, CodEmpresa As Integer)
        Dim sql As String = "UPDATE Tbl_EmpresaContas SET NomeConta = @NOMECONTA,  Conta = @CONTA, Agencia = @AGENCIA, Banco = @BANCO WHERE CodigoConta = @CODIGO"
        Dim parameters As SqlParameter() = {
          New SqlParameter("@CODIGO", CodConta),
          New SqlParameter("@NOMECONTA", NomeConta),
          New SqlParameter("@CONTA", Conta),
          New SqlParameter("@AGENCIA", Agencia),
          New SqlParameter("@BANCO", Banco),
          New SqlParameter("@CodEmpresa", CodEmpresa)
    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    ''' <summary>
    ''' Este metódo exclui a conta da empresa do banco de dados.
    ''' </summary>
    ''' <param name="CodConta">Código identificador da conta no banco de dados</param>
    Public Sub ExcluirContaBancaria(CodConta As Integer)
        Dim sql As String = "DELETE FROM Tbl_EmpresaContas WHERE CodigoConta = @CODIGO"
        Dim parameters As SqlParameter() = {
          New SqlParameter("@CODIGO", CodConta)
    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    ''' <summary>
    ''' Este metódo executa uma procedure no banco de dados que atualiza o custo de produção dos itens com base na composição.
    ''' </summary>
    Public Sub AtualizaoCustoProducao(CodLoja As Integer)
        Dim parameters As SqlParameter() = {
            New SqlParameter("@CodLoja", CodLoja)
        }
        ClasseConexao.ExecutarProcedure("spAtualizaCustoProducao", parameters)
    End Sub
    ''' <summary>
    ''' Este metódo executa uma procedure no banco de dados que atualiza o custo da lojas.
    ''' </summary>
    Public Sub CalculaCustoLojas()
        ClasseConexao.ExecutarProcedure("spCalculaCustos", Nothing)
    End Sub

#End Region
#Region "FUNCOES"
    ''' <summary>
    ''' Esta função consulta o dados da empresa.
    ''' </summary>
    ''' <param name="sql">Query sql necessária para a consulta.</param>
    ''' <returns>Retorna os dados solicitado na query.</returns>
    Public Function ConsultaEmpresa(sql As String, Optional CodEmpresa As Integer = 0)
        If CodEmpresa <> 0 Then
            Dim parameters As SqlParameter() = {
                New SqlParameter("@CodEmpresa", CodEmpresa)
            }
            Return ClasseConexao.Consultar(sql, parameters)
        Else
            Return ClasseConexao.Consultar(sql, Nothing)
        End If
    End Function
    ''' <summary>
    ''' Esta função realiza a consulta dos dados de conta bancária da empresa.
    ''' </summary>
    ''' <returns>Retorna os dados da conta bancária da empresa.</returns>
    Public Function ConsultaContaBancaria(sql As String, Optional CodConta As Integer = 0)
        If CodConta <> 0 Then
            Dim parameters As SqlParameter() = {
                New SqlParameter("@CodConta", CodConta)
            }
            Return ClasseConexao.Consultar(sql, parameters)
        Else
            Return ClasseConexao.Consultar(sql, Nothing)
        End If
    End Function

#End Region
End Class
