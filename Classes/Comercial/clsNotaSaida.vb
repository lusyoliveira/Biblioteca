Imports System.Data
Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports Xceed.Wpf.Toolkit
Namespace Classes.Comercial
    Public Class clsNotaSaida
        Inherits clsComercial
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property Destinatario As String
        Public Property TotalReceber As Decimal
        Public Property DataPedido As Date
        Public Property CodEntrega As Integer
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
        Public Sub AtualizaDespache(CodPedido As Integer, DataDespache As Date, CodigoRastreamento As String, CodEntrega As Integer)
            Dim sql As String = "UPDATE Tbl_Vendas SET  DtDespache = @DataDespache,
                                                CodigoRastreamento = @CodigoRastreamento, 
                                                Status = 1,
                                                CodEntrega = @CodEntrega 
                                            WHERE CodVenda = @CODVENDA"
            Dim parameters As SqlParameter() = {
                    New SqlParameter("@CODVENDA", CodPedido),
                    New SqlParameter("@DataDespache", DataDespache),
                    New SqlParameter("@CodigoRastreamento", CodigoRastreamento),
                    New SqlParameter("@CodEntrega", CodEntrega)
                                         }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Pedido atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
#End Region
#Region "FUNCOES"
        Public Function ConsultaNotaSaida(sql As String, Optional CodCli As Integer = 0)
            If CodCli <> 0 Then
                Dim parameters As SqlParameter() = {
            New SqlParameter("@CodCli", CodCli)
    }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function

        Public Function ItensPedido(CodPedido As Integer) As List(Of clsNotaSaida)
            Dim listaitens = New List(Of clsNotaSaida)()
            Dim sql As String = "SELECT * FROM Cs_VendasDetalhes WHERE CodPedido = @CodPedido"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@CodPedido", CodPedido)
    }

            Dim dt As DataTable = ClasseConexao.Consultar(sql, parameters)

            ' Percorre cada linha do DataTable
            For Each row As DataRow In dt.Rows
                ' Cria uma nova instância de clsNotaSaida
                Dim nota As New clsNotaSaida()

                ' Preenche as propriedades com os valores da linha
                nota.CodItem = Convert.ToInt32(row("CodSimples"))
                nota.Item = row("Item").ToString()
                nota.Quantidade = row("Quantidade").ToString()
                nota.ValorUnit = row("ValorUnit").ToString()

                ' Adiciona o objeto à lista
                listaitens.Add(nota)
            Next
            Return listaitens
        End Function
#End Region
    End Class

End Namespace
