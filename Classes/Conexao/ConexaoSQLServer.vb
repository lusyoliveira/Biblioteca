Imports System.Data
Imports Microsoft.Data.SqlClient
Imports System.IO
Imports System.ServiceProcess
Namespace Classes.Conexao
    Public Class ConexaoSQLServer
#Region "PROPRIEDADES"
        Private pServidor = GetNomeSQLServer()
        Public Property Servidor As String
            Get
                Return pServidor
            End Get
            Set(value As String)
                pServidor = value
            End Set
        End Property
        Private pDataBase = "SysVendasTeste"
        Public Property DataBase As String
            Get
                Return pDataBase
            End Get
            Set(value As String)
                pDataBase = value
            End Set
        End Property
        Private pUser = "sa"
        Public Property user As String
            Get
                Return pUser
            End Get
            Set(value As String)
                pUser = value
            End Set
        End Property

        Private pPassword = "123456"
        Public Property password As String
            Get
                Return pPassword
            End Get
            Set(value As String)
                pPassword = value
            End Set
        End Property
        Private pstrConexao As String = $"Data Source={Servidor};Initial Catalog={DataBase};User ID={user}; Password={password};Integrated Security=True"
        Public Property strConexao As String
            Get
                Return pstrConexao
            End Get
            Set(value As String)
                pPassword = value
            End Set
        End Property
        Public Property NomeArquivo As String = "Banco_Dados.ini"
#End Region
#Region "CONSTRUTORES"
        Public Sub New()
            Dim Arquivo As String = ObterArquivo(NomeArquivo)

            If File.Exists(Arquivo) Then
                Servidor = LeArquivoINI(Arquivo, "Geral", "Servidor", "INSIRA O SERVIDOR")
                DataBase = LeArquivoINI(Arquivo, "Geral", "Banco", "INSIRA O BANCO DE DADOS")
                pstrConexao = $"Data Source={Servidor};Initial Catalog={DataBase};User ID=sa; Password=123456;Integrated Security=True"
            Else
                MessageBox.Show("Arquivo de configuração não encontrado! Será carrregado as configurações padrão do sistema.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub
#End Region
#Region "METODOS"
        ''' <summary>
        ''' Este método realiza operações de inserção, alteração e exclusão numa tabela do banco de dados.
        ''' </summary>
        ''' <param name="sql">Representa a string sql para realizar uma inserção, alteração ou exclusão numa tabela do banco de dados.</param>
        Public Sub Operar(ByVal sql As String, ByVal parameters As SqlParameter())
            Try
                Using cn As New SqlConnection(strConexao)
                    cn.Open()
                    Using cmd As New SqlCommand(sql, cn)
                        ' Adiciona parâmetros ao comando, se houver
                        If parameters IsNot Nothing Then
                            cmd.Parameters.AddRange(parameters)
                        End If
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As Exception '
                MessageBox.Show("Não foi possível realizar a operação" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Sub
        ''' <summary>
        ''' Este metódo executar uma procedure no banco de dados.
        ''' </summary>
        ''' <param name="procedureName">Representa o nome da procedure no banco de dados.</param>
        ''' <param name="parameters">Representa os parâmetro necessário para a procedure.</param>
        Public Sub ExecutarProcedure(ByVal procedureName As String, ByVal parameters As SqlParameter())
            Try
                Using cn As New SqlConnection(strConexao)
                    Using cmd As New SqlCommand(procedureName, cn)
                        ' Define o tipo de comando como stored procedure
                        cmd.CommandType = CommandType.StoredProcedure

                        ' Adiciona parâmetros ao comando, se houver
                        If parameters IsNot Nothing Then
                            cmd.Parameters.AddRange(parameters)
                        End If

                        ' Abre a conexão e executa a stored procedure
                        cn.Open()
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Não foi possível executar a procedure!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
        Private Sub AlterarStringDeConexao()
            Dim Config = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.None)
            Dim ConnectionStrings = Config.ConnectionStrings
            For Each ConnectionString As System.Configuration.ConnectionStringSettings In ConnectionStrings.ConnectionStrings
                ConnectionString.ConnectionString = String.Format($"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={Servidor}\{DataBase}.mdb", Environment.CurrentDirectory)
            Next
            Config.Save(System.Configuration.ConfigurationSaveMode.Modified)
            System.Configuration.ConfigurationManager.RefreshSection("cnStrings")
        End Sub
        ''' <summary>
        ''' Esta função obtem o nome do servidor SQL Server.
        ''' </summary>
        ''' <returns>Retorna o nome do servidor.</returns>
        Public Function GetNomeSQLServer() As String
            'Nome do PC local
            Dim strPCname As String = Environment.MachineName
            ' nome do serviço do SQL Server Express
            Dim strInstancia As String = "MSSQL$SQLEXPRESS"
            Dim strNomeSQLServer As String = String.Empty

            ' Inclua uma referência a : System.ServiceProcess;
            Dim servicos As ServiceController() = ServiceController.GetServices()
            ' percorre os serviços 
            For Each servico As ServiceController In servicos
                If servico Is Nothing Then
                    Continue For
                End If
                Dim strNomeDoServico As String = servico.ServiceName
                If strNomeDoServico.Contains(strInstancia) Then
                    strNomeSQLServer = strNomeDoServico
                End If
            Next
            Dim IndiceInicio As Integer = strNomeSQLServer.IndexOf("$")
            If IndiceInicio > -1 Then
                'strSqlServerName=NomeDoSeuPC\SQLEXPRESS;
                strNomeSQLServer = strPCname + "\" + strNomeSQLServer.Substring(IndiceInicio + 1)
            End If
            Return strNomeSQLServer
        End Function

        'Private Function verificaSeBancoDadosExiste(ByVal strNomeBD As String) As Boolean
        'Inclua referências a todas As .dll's que estão na pasta  "Scripts" 
        'estes arquivos dll's estão na pasta C:\Program Files\Microsoft SQL Server\100\SDK\Assemblies
        'Dim dbServer As New Server(GetNomeSQLServer())
        'Dim dbServer As New Server{GetNomeSQLServer()}
        'If dbServer.Databases(strNomeBD) IsNot Nothing Then
        '    Return True
        'End If
        'Return False
        'Select Case* From sys.databases 

        'End Function

#End Region
#Region "FUNCOES"
        ''' <summary>
        ''' Esta função realiza uma consulta no banco de dados e retorna a lista dos dados.
        ''' </summary>
        ''' <param name="sql">Representa a string sql para realizar a consulta.</param>
        ''' <returns>Retorna a tabela consultada.</returns>
        Public Function Consultar(ByVal sql As String, ByVal parameters As SqlParameter()) As DataTable
            Dim dt As New DataTable()

            Try
                Using cn As New SqlConnection(strConexao)
                    Using cmd As New SqlCommand(sql, cn)
                        ' Adiciona parâmetros ao comando, se houver
                        If parameters IsNot Nothing Then
                            cmd.Parameters.AddRange(parameters)
                        End If

                        Using adpt As New SqlDataAdapter(cmd)
                            adpt.Fill(dt)
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Não foi possível consultar os dados!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ' Opcional: log do erro ou tratamento adicional
            End Try

            Return dt
        End Function
        Public Function ExecProcedureRetorno(ByVal procedureName As String, ByVal parameters As SqlParameter()) As DataTable
            Dim dt As New DataTable()

            Try
                Using cn As New SqlConnection(strConexao)
                    Using cmd As New SqlCommand(procedureName, cn)
                        ' Define o tipo de comando como stored procedure
                        cmd.CommandType = CommandType.StoredProcedure

                        ' Adiciona parâmetros ao comando, se houver
                        If parameters IsNot Nothing Then
                            cmd.Parameters.AddRange(parameters)
                        End If

                        Using adpt As New SqlDataAdapter(cmd)
                            adpt.Fill(dt)
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Não foi possível consultar os dados!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ' Opcional: log do erro ou tratamento adicional
            End Try

            Return dt
        End Function

#End Region
    End Class
End Namespace

