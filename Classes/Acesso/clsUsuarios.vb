Imports System.Data
Imports System.Net.Mime.MediaTypeNames
Imports System.Text
Imports Microsoft.Data.SqlClient
Imports Xceed.Wpf.Toolkit
Imports Biblioteca.Classes.Entidades.Vendas
Imports Biblioteca.Classes.Conexao
Namespace Bibliotecas.Classes.Acessos
    Public Class clsUsuarios
        Inherits clsEntidades
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property ID As Integer
        Public Property Usuario As String
        Public Property Nivel As Integer
        Public Property Nome As String
        Public Property NomeNivel As String
        Public Property Inativo As Boolean
        Public Property Tentativas As Integer = 3
        Public Property Autenticado As Boolean = False
        Public Property MenuAcesso As List(Of clsAcessos)
#End Region
#Region "CONSTRUTORES"
        Public Sub New()

        End Sub

        Public Sub New(_id As Integer, _usuario As String, _nivel As Integer, _nome As String, _nivelusuario As String, _inativo As Boolean, _email As String)
            ID = _id
            Usuario = _usuario
            Nivel = _nivel
            Nome = _nome
            NomeNivel = _nivelusuario
            Inativo = _inativo
            Email = _email
        End Sub
#End Region
#Region "METODOS"
        Public Sub SalvarUsuario(Usuario As String, Nivel As Integer, Nome As String, Senha As String, Ativo As Boolean, Email As String)
            Dim sql As String = "INSERT INTO Tbl_Usuarios (Login,Senha,Nivel,Nome,Email,Inativo) VALUES (@Usuario,@Senha,@Permissao,@Nome,@Email,@Ativo)"
            Dim parameters As SqlParameter() = {
                        New SqlParameter("@Usuario", Usuario),
                        New SqlParameter("@Senha", geraHash(Senha)),
                        New SqlParameter("@Permissao", Nivel),
                        New SqlParameter("@Nome", Nome),
                        New SqlParameter("@Email", Email),
                        New SqlParameter("@Ativo", Ativo)
            }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizarUsuario(Usuario As String, Nivel As Integer, Nome As String, Senha As String, Inativo As Boolean, Email As String)
            Dim Sql As String = "UPDATE Tbl_Usuarios SET Nivel=@Nivel, Nome=@Nome, email=@Email, Inativo=@Inativo "
            If Senha <> "" Then
                Sql &= ", senha = '" & geraHash(Senha) & "'"
            End If
            Sql &= "  WHERE Login = '" & Usuario & "'"

            Dim parameters As SqlParameter() = {
                        New SqlParameter("@Nivel", Nivel),
                        New SqlParameter("@Nome", Nome),
                        New SqlParameter("@Email", Email),
                        New SqlParameter("@Inativo", Inativo)
            }
            ClasseConexao.Operar(Sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluirUsuario(Usuario As String)
            Dim sql As String = "DELETE FROM Tbl_Usuarios WHERE Login = @Usuario"
            Dim parameters As SqlParameter() = {
                        New SqlParameter("@Usuario", Usuario)
            }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AlteraSenha(Usuario As String, Senha As String)
            Dim sql As String = "UPDATE Tbl_Usuarios SET Senha=@Senha WHERE Login = @Usuario"
            Dim parameters As SqlParameter() = {
                        New SqlParameter("@Usuario", Usuario),
                        New SqlParameter("@Senha", geraHash(Senha))
            }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Senha atualizada com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
#End Region
#Region "FUNCOES"
        Public Function ConsultaUsuario(sql As String, Optional Usuario As String = Nothing)
            If Usuario IsNot Nothing Then
                Dim parameters As SqlParameter() = {
                New SqlParameter("@Usuario", Usuario)
            }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
        Public Function PesquisaUsuarios(Status As String, Usuario As String, Nivel As String, Nome As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_Usuarios WHERE 1=1 ")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Ativos"
                        sql.AppendLine("AND Inativo = 0")
                    Case "Inativos"
                        sql.AppendLine("AND Inativo = 1")
                    Case "Todos"
                        sql.AppendLine("AND Inativo IS NOT NULL")
                End Select

                If Not String.IsNullOrEmpty(Usuario) Then
                    sql.AppendLine("AND Usuario LIKE @Usuario")
                    parameters.Add(New SqlParameter("@Usuario", "%" & Usuario & "%"))
                End If

                If Not String.IsNullOrEmpty(Nivel) Then
                    sql.AppendLine("AND Nivel LIKE @Nivel")
                    parameters.Add(New SqlParameter("@Nivel", "%" & Nivel & "%"))
                End If

                If Not String.IsNullOrEmpty(Nome) Then
                    sql.AppendLine("AND Nome LIKE @Nome")
                    parameters.Add(New SqlParameter("@Nome", "%" & Nome & "%"))
                End If

                sql.AppendLine("ORDER BY Nome")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As SqlException
                MessageBox.Show("Não foi possível realizar a consulta: " & ex.Message, "Erro de Banco de Dados", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function

        Public Function Autenticar(Usuario As String, senha As String, ByRef DadosUsuario As clsUsuarios) As Boolean

            Dim Resposta As Boolean = ValidarUsuario("SELECT * FROM Cs_Usuarios WHERE Usuario LIKE @Usuario AND INATIVO = 0", Usuario, senha)

            Select Case DadosUsuario.Tentativas
                Case Is <> 0
                    If Resposta Then
                        DadosUsuario.Tentativas = DadosUsuario.Tentativas - 1
                        Autenticado = True
                        Return True
                    Else
                        DadosUsuario.Tentativas = DadosUsuario.Tentativas - 1
                        MessageBox.Show("Usuário ou senha incorreto!" & vbCrLf & "Restam: " & DadosUsuario.Tentativas & " tentativas.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Autenticado = False
                        Return False
                    End If
                Case Is <= 0
                    DadosUsuario.Tentativas = DadosUsuario.Tentativas - 1
                    MessageBox.Show("Suas tentativas expiraram!" & vbCrLf & "O aplicativo será fechado!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Application.Exit()
                    Autenticado = False
                    Return False
            End Select
        End Function

        Public Sub ValidaInativacaoUsuario(CodUsuario As Integer, Inativo As Integer)
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodUsuario", CodUsuario),
                         New SqlParameter("@SITUACAOSITEMA", Inativo)
            }
            Dim Tabela As DataTable = ClasseConexao.ExecProcedureRetorno("spAtivaInativaUsuario", parameters)
            If Tabela IsNot Nothing AndAlso Tabela.Rows.Count > 0 Then

                If Tabela.Rows(0)("Mensagem").ToString() = "Ativo" Then
                    MessageBox.Show("Usuário já está ativo!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ElseIf Tabela.Rows(0)("Mensagem").ToString() = "Inativo" Then
                    MessageBox.Show("Usuário já está inativo!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Usuário alterado com Sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

            End If
        End Sub
        ''' <summary>
        ''' Esta função validar se o usuário informado existe no sistema.
        ''' </summary>
        ''' <param name="senha">Representa a senha informada pelo usuário</param>
        ''' <param name="Usuario">Representa o usuario do tipo string.</param>
        ''' <returns></returns>
        Public Function ValidarUsuario(sql As String, Optional Usuario As String = Nothing, Optional senha As String = Nothing) As Boolean

            If Usuario IsNot Nothing Then
                Dim parameters As SqlParameter() = {
              New SqlParameter("@Usuario", Usuario)
      }

                Dim Tabela As DataTable = ClasseConexao.Consultar(sql, parameters)
                If Tabela IsNot Nothing AndAlso Tabela.Rows.Count > 0 Then
                    If senha IsNot Nothing Then
                        senha = geraHash(senha)
                        If senha = Tabela.Rows(0)("senha").ToString() Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                    Return True
                Else
                    Return False
                End If
            End If

        End Function
#End Region
    End Class
End Namespace

