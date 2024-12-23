Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports System.Data
Imports Xceed.Wpf.Toolkit
Imports System.Windows.Forms
Namespace Classes.Acessos
    ''' <summary>
    ''' Esta classe representa as rotinas de controle de acesso do sistema.
    ''' </summary>
    Public Class clsAcessos
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property Codigo As Integer
        Public Property NomeMenu As String
        Public Property DescricaoMenu As String
        Public Property NivelMenu As Integer
        Public Property Liberado As Byte
        Public Property CodigoNivel As Integer
        Public Property CodigoMenu As Integer
        Public Property NivelAcesso As String
#End Region
#Region "CONSTRUTORES"
        Public Sub New()

        End Sub
        Public Sub New(_nome As String, _descricao As String, _nivel As Integer)
            NomeMenu = _nome
            DescricaoMenu = _descricao
            NivelMenu = _nivel
        End Sub
        Public Sub New(_Codigo As Integer, _NomeMenu As String, _DescricaoMenu As String, _NivelMenu As Integer, _Liberado As Byte, _CodigoNivel As Integer, _CodigoMenu As Integer)
            Codigo = _Codigo
            NomeMenu = _NomeMenu
            DescricaoMenu = _DescricaoMenu
            NivelMenu = _NivelMenu
            Liberado = _Liberado
            CodigoNivel = _CodigoNivel
            CodigoMenu = _CodigoMenu
        End Sub

        Public Sub New(_Codigo As Integer, _Liberado As Byte, _CodigoNivel As Integer, _CodigoMenu As Integer)
            Codigo = _Codigo
            Liberado = _Liberado
            CodigoNivel = _CodigoNivel
            CodigoMenu = _CodigoMenu
        End Sub
#End Region
#Region "FUNÇÕES"
        Public Function ConsultaNivel(sql As String, Optional CodNivel As Integer = 0)
            If CodNivel <> 0 Then
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodNivel", CodNivel)
                            }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
        Public Function ConsultaAcessos(Descricao As String) As Boolean

            Dim sql As String = "SELECT * FROM Tbl_MenuAcessos WHERE Descricao LIKE @Descricao"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@Descricao", Descricao)
                        }
            Dim Tabela As DataTable = ClasseConexao.Consultar(sql, parameters)
            If Tabela IsNot Nothing AndAlso Tabela.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        End Function
        ''' <summary>
        ''' Esta função realiza a coleta do menus existentes e guarda numa lista.
        ''' </summary>
        ''' <param name="Menu">Representa o nome do componente menu</param>
        ''' <returns>Retorna a lista de menus</returns>
        Public Function MostrarOpcoes(Menu As MenuStrip) As List(Of clsAcessos)
            Dim listaopcoes = New List(Of clsAcessos)()
            'Nivel 1
            For Each item In Menu.Items
                Dim descricao1 = item.Text
                If item.HasDropDownItems Then
                    'Nivel 2
                    For Each opcao In item.DropDownItems
                        Dim descricao2 = descricao1 & "/" & opcao.Text
                        If opcao.HasDropDownItems Then
                            'Nivel 3
                            For Each subOpcao In opcao.DropDownItems
                                Dim descricao3 = descricao2 & "/" & subOpcao.Text
                                If subOpcao.HasDropDownItems Then
                                    'Nivel 4
                                    For Each Nivel4 In subOpcao.DropDownItems
                                        Dim descricao4 = descricao3 & "/" & Nivel4.Text
                                        listaopcoes.Add(New clsAcessos(Nivel4.Name, descricao4, 4))
                                    Next
                                Else
                                    listaopcoes.Add(New clsAcessos(subOpcao.Name, descricao3, 3))
                                End If
                            Next
                        Else
                            listaopcoes.Add(New clsAcessos(opcao.Name, descricao2, 2))
                        End If
                    Next
                Else
                    listaopcoes.Add(New clsAcessos(item.Name, descricao1, 1))
                End If
            Next
            Return listaopcoes
        End Function
        ''' <summary>
        ''' Esta função realiza a coleta dos menus do sistema e salva numa lista, verificando se este menu já existe no banco de dados. Caso ele não existe é inserido na lista.
        ''' </summary>
        ''' <param name="Menu">Representa o nome do componente do menu</param>
        ''' <returns>Retorna uma lista de menus do sistema.</returns>
        Public Function CriarMenus(Menu As MenuStrip) As List(Of clsAcessos)
            Dim listaopcoes = New List(Of clsAcessos)()

            'Nivel 1
            For Each item In Menu.Items
                Dim descricao1 = item.Text
                If item.HasDropDownItems Then
                    'Nivel 2
                    For Each opcao In item.DropDownItems
                        Dim descricao2 = descricao1 & "/" & opcao.Text
                        If opcao.HasDropDownItems Then
                            'Nivel 3
                            For Each subOpcao In opcao.DropDownItems
                                Dim descricao3 = descricao2 & "/" & subOpcao.Text
                                If subOpcao.HasDropDownItems Then
                                    'Nivel 4
                                    For Each Nivel4 In subOpcao.DropDownItems
                                        Dim descricao4 = descricao3 & "/" & Nivel4.Text
                                        If ConsultaAcessos(descricao4) = False Then
                                            listaopcoes.Add(New clsAcessos(Nivel4.Name, descricao4, 4))
                                        End If
                                    Next
                                Else
                                    If ConsultaAcessos(descricao3) = False Then
                                        listaopcoes.Add(New clsAcessos(subOpcao.Name, descricao3, 3))
                                    End If
                                End If
                            Next
                        Else
                            If ConsultaAcessos(descricao2) = False Then
                                listaopcoes.Add(New clsAcessos(opcao.Name, descricao2, 2))
                            End If
                        End If
                    Next
                Else
                    If ConsultaAcessos(descricao1) = False Then
                        listaopcoes.Add(New clsAcessos(item.Name, descricao1, 1))
                    End If
                End If
            Next
            Return listaopcoes
        End Function
#End Region
#Region "METODOS"

        ''' <summary>
        ''' Este metódo realizar a inserção dos menus de uma lista no banco de dados.
        ''' </summary>
        Public Sub SalvarListaMenus(listaopcoes As List(Of clsAcessos)) '(DescricaoMenu As String, NomeMenu As String, NivelMenu As Integer)
            For Each item In listaopcoes
                Dim sql = "INSERT INTO Tbl_MenuAcessos (nome, NivelMenu, descricao,DataCriacao)  VALUES (@nome,@nivel,@descricao,GETDATE())"

                Dim parameters As SqlParameter() = {
                         New SqlParameter("@nome", item.NomeMenu),
                         New SqlParameter("@nivel", item.NivelMenu),
                         New SqlParameter("@Descricao", item.DescricaoMenu)
        }
                ClasseConexao.Operar(sql, parameters)
            Next
        End Sub
        Public Sub ValidaAcessos(Menu As MenuStrip, CodigoNivel As Integer)
            Dim sql = "SELECT * FROM Cs_NivelAcesso WHERE CodigoNivel = @CodigoNivel"

            Dim parameters As SqlParameter() = {
                                 New SqlParameter("@CodigoNivel", CodigoNivel)
                }
            Dim Tabela As DataTable = ClasseConexao.Consultar(sql, parameters)
            If Tabela IsNot Nothing AndAlso Tabela.Rows.Count > 0 Then
                For Each item In Menu.Items
                    LiberarMenu(item, Tabela)
                    For Each opcao In item.DropDownItems
                        LiberarMenu(opcao, Tabela)
                        For Each subOpcao In opcao.DropDownItems
                            LiberarMenu(subOpcao, Tabela)
                            For Each Nivel4 In subOpcao.DropDownItems
                                LiberarMenu(Nivel4, Tabela)
                            Next
                        Next
                    Next
                Next
            End If
        End Sub
        ''' <summary>
        ''' Este função receber uma lista de acessos do nivel e verificar se o nível está liberado, sendo True para o menu liberado e False para o não liberador.
        ''' </summary>
        ''' <param name="opcao">Representa a opção do nível de acesso.</param>
        ''' <param name="usuarioAcessos">Representa a lista de acesso do nível</param>
        ''' <returns>Retorna verdadeiro ou falso para o nivel.</returns>
        Public Function LiberarMenu(opcao As ToolStripMenuItem, usuarioAcessos As DataTable)
            For Each row As DataRow In usuarioAcessos.Rows
                If row("nomeOpcao") = opcao.Name Then
                    If row("Liberado") = 0 Then
                        opcao.Visible = False
                        Exit For
                    End If
                End If
            Next
        End Function
        ''' <summary>
        ''' Este metódo recebe uma lista de menus para que seja verificado se este menu existe no sistema, caso ele não existe ele será excluido das tabelas do banco de dados.
        ''' </summary>
        ''' <param name="listaopcoes">Representa uma lista de menus.</param>
        Public Sub ExcluirMenus(listaopcoes As List(Of clsAcessos))
            ' Exclui menu não existentes no sistema
            Dim sql As String = "spExcluiMenusObsoletos"
            Try
                Using cn As New SqlConnection(ClasseConexao.strConexao)
                    cn.Open()

                    ' Criar tabela temporária para armazenar os menus da segunda tabela
                    Using cmdCreateTable As New SqlCommand("CREATE TABLE #Tbl_MenuAcessosTemp (Nome NVARCHAR(255), Descricao NVARCHAR(255));", cn)
                        cmdCreateTable.ExecuteNonQuery()
                    End Using

                    ' Inserir os menus na tabela temporária
                    Using cmdInsert As New SqlCommand("INSERT INTO #Tbl_MenuAcessosTemp (Nome, Descricao) VALUES (@NOME, @DESCRICAO);", cn)
                        cmdInsert.Parameters.Add("@NOME", SqlDbType.VarChar)
                        cmdInsert.Parameters.Add("@DESCRICAO", SqlDbType.VarChar)
                        For Each item In listaopcoes
                            cmdInsert.Parameters("@NOME").Value = item.NomeMenu
                            cmdInsert.Parameters("@DESCRICAO").Value = item.DescricaoMenu
                            cmdInsert.ExecuteNonQuery()
                        Next
                    End Using

                    ' Executar a stored procedure para excluir menus obsoletos
                    Using cmd As New SqlCommand(sql, cn)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.ExecuteNonQuery()
                    End Using

                    ' Limpar a tabela temporária
                    Using cmdDropTable As New SqlCommand("DROP TABLE #Tbl_MenuAcessosTemp;", cn)
                        cmdDropTable.ExecuteNonQuery()
                    End Using

                    cn.Close()
                End Using
            Catch ex As Exception
                MessageBox.Show("Não foi possível excluir os acessos!" & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' Esse metódo atualiza o nível de acesso se uma usuário.
        ''' </summary>
        ''' <param name="CodNivel">Representa o código do nível de acesso do tipo integer.</param>
        ''' <param name="Liberado">Representa se o nível está liberado, sendo 0 para não liberar e 1 para liberado.</param>
        ''' <param name="CodigoMenu">Representa o código do menu de acesso que será atualizado.</param>
        Public Sub AtualizaPermissoes(CodNivel As Integer, Liberado As Integer, CodigoMenu As Integer)
            Dim sql = "UPDATE Tbl_Permissoes SET Liberado = @Liberado WHERE CodigoNivel = @CodNivel AND CodigoMenu = @CodigoMenu"

            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodigoMenu", CodigoMenu),
                         New SqlParameter("@Liberado", Liberado),
                         New SqlParameter("@CodNivel", CodNivel)
        }
            ClasseConexao.Operar(sql, parameters)
        End Sub

        ''' <summary>
        ''' Este metódo salva no banco de dados um novo nível de acesso.
        ''' </summary>
        ''' <param name="Nivel">Representa o nome do nível do tipo string.</param>
        Public Sub SalvarNivel(Nivel As String)
            Dim sql = "INSERT INTO Tbl_NivelUsuario (Nivel) VALUES (@NIVEL)"

            Dim parameters As SqlParameter() = {
                         New SqlParameter("@NIVEL", Nivel)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro realizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
#End Region
    End Class
End Namespace

