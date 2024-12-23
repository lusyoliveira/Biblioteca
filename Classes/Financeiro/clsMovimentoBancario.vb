Imports System.Data
Imports System.Text
Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports Xceed.Wpf.Toolkit
Namespace Classes.Financeiro
    Public Class clsMovimentoBancario
        Inherits clsFinanceiro
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property CodPedido As Integer
        Public Property ValorPago As Decimal
        Public Property ValorParcela As Decimal
        Public Property CodEnt As Integer
        Public Property Entidade As String
        Public Property TipoEntidade As String
        Public Property DataVencto As Date
        Public Property DataPagto As Date
        Public Property Dias As Integer
        Public Property NParcelas As Integer
        Public Property NumeroConta As String
        Public Property CodTitulo As Integer
        Public Property DataMovimento As DateTime
        Public Property Credito As Decimal
        Public Property Debito As Decimal
        Public Property CredDeb As String
        Public Property Saldo As Decimal
        Public Property TipoDocumento As String
        Public Property IDBanco As Integer
        Public Property Banco As String
        Public Property CodTipoDocumento As Integer
        Public Property CodMov As Integer
        Public Property DataInicial As Date
        Public Property DataFinal As Date
#End Region
#Region "METODOS"

        Public Sub IncluirMovBancario(DataMovimento As Date, CodConta As Integer, Valor As Decimal, Movimento As String, TipoDocumento As Integer, CodTitulo As Integer, Complemento As String)
            Dim sql As String = "INSERT INTO tbl_MovimentoBancario (DataMovimento,CodConta,Valor,Movimento,TipoDocumento,CodTitulo,Complemento)
                                                      VALUES (@DTMOV,@CODCONTA,@VALOR,@MOV,@TIPODOC,@CODTITULO,@COMPLEMENTO)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CODTITULO", CodTitulo),
                         New SqlParameter("@CODCONTA", CodConta),
                         New SqlParameter("@DTMOV", DataMovimento),
                         New SqlParameter("@MOV", Movimento),
                         New SqlParameter("@VALOR", Valor),
                         New SqlParameter("@TIPODOC", TipoDocumento),
                         New SqlParameter("@COMPLEMENTO", Complemento)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaMovBancario(CodConta As Integer, Valor As Decimal, Movimento As String, TipoDocumento As Integer, CodTitulo As Integer, Complemento As String)
            Dim sql As String = "UPDATE Tbl_MovimentoBancario  SET    CodConta = @CODCONTA,
                                                                Valor = @VALOR,
                                                                Movimento = @MOV,
                                                                TipoDocumento = @TIPODOC
                                                                Complemento = @COMPLEMENTO
                                                         WHERE  CodMovimento = @CODTITULO"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CODTITULO", CodTitulo),
                         New SqlParameter("@CODCONTA", CodConta),
                         New SqlParameter("@MOV", Movimento),
                         New SqlParameter("@VALOR", Valor),
                         New SqlParameter("@TIPODOC", TipoDocumento),
                         New SqlParameter("@COMPLEMENTO", Complemento)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluiMovBancario(CodTitulo As Integer)
            Dim sql As String = "UPDATE Tbl_MovimentoBancario  SET    CodConta = @CODCONTA,
                                                                Valor = @VALOR,
                                                                Movimento = @MOV,
                                                                TipoDocumento = @TIPODOC
                                                                Complemento = @COMPLEMENTO
                                                         WHERE  CodMovimento = @CODTITULO"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CODTITULO", CodTitulo)
                         }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
#End Region
#Region "FUNÇÕES"
        Public Function PesquisaMovimentoBancario(Status As String, Codigo As String, TipoDoc As String, Conta As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_MovimentoBancario WHERE 1=1 ")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Crédito"
                        sql.AppendLine("AND Movimento LIKE 'C'")
                    Case "Débito"
                        sql.AppendLine("AND Movimento LIKE 'D'")
                    Case "Todos"
                        sql.AppendLine("AND Movimento IS NOT NULL")
                End Select

                'Pesquisa por codigo do pagamento/recebimento
                If Not String.IsNullOrEmpty(Codigo) Then
                    sql.AppendLine("AND CodTitulo = @CodTitulo")
                    parameters.Add(New SqlParameter("@CodTitulo", Codigo))
                End If

                'Pesquisa pelo tipo de documento
                If Not String.IsNullOrEmpty(TipoDoc) Then
                    sql.AppendLine("AND TipoDocumento = @TipoDocumento")
                    parameters.Add(New SqlParameter("@TipoDocumento", TipoDoc))
                End If

                'Pesquisa pela conta
                If Not String.IsNullOrEmpty(Conta) Then
                    sql.AppendLine("AND NomeConta = @Conta")
                    parameters.Add(New SqlParameter("@Conta", Conta))
                End If

                sql.AppendLine("ORDER BY DataMovimento")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function
        Public Function ConsultaMovimento(sql As String, Optional DATAINI As String = Nothing, Optional DATAFI As String = Nothing)

            If DATAINI IsNot Nothing And DATAFI IsNot Nothing Then
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@dataini", DATAINI),
                         New SqlParameter("@datafim", DATAFI)
        }
                Return ClasseConexao.Consultar(sql, parameters)

            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If

        End Function
#End Region
    End Class
End Namespace

