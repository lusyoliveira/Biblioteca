Imports Biblioteca.Classes.Conexao
Imports Biblioteca.Classes.Produtos
Imports Microsoft.Data.SqlClient
Imports Xceed.Wpf.Toolkit

Namespace Classes.Compras
    ''' <summary>
    ''' Este classe representa todas as rotinas em comum a transações de compra do sistema. Esta é uma classe Base.
    ''' </summary>
    Public Class clsCompras
        Inherits clsItem
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property CodCotacao As Integer
        Public Property CodPedido As Integer
        Public Property CodNotaFiscal As Integer
        Public Property NotaFiscal As Integer
        Public Property FormaPagto As String
        Public Property Cobranca As String
        Public Property Status As String
        Public Property Prazo As Integer
        Public Property Desconto As Decimal
        Public Property Frete As Decimal
        Public Property ValorUnit As Double
        Public Property ValorUnitario As Decimal
        Public Property ValorTotal As Decimal
        Public Property Total As Decimal

#End Region
#Region "METODOS"
        Public Sub SalvaComprador(Comprador As String)
            Dim sql As String = "INSERT INTO Tbl_Comprador (Comprador) VALUES (@Comprador)"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@Comprador", Comprador)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaComprador(CodComprador As Integer, Comprador As String)
            Dim sql As String = "UPDATE Tbl_Comprador SET Comprador = @Comprador WHERE Codigo = @CodComprador"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@CodComprador", CodComprador),
            New SqlParameter("@Comprador", Comprador)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluiComprador(CodComprador As Integer)
            Dim sql As String = "DELETE FROM Tbl_Comprador WHERE Codigo = @CodComprador"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@CodComprador", CodComprador)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
#End Region
#Region "FUNCOES"
        Public Function ConsultaComprador(sql As String)
            Return ClasseConexao.Consultar(sql, Nothing)
        End Function
#End Region
    End Class
End Namespace

