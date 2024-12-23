Imports System.Text
Imports Biblioteca.Classes.Conexao
Imports Biblioteca.Classes.Comercial
Public Class clsProducao
    Inherits clsComercial
    Dim ClasseConexao As New ConexaoSQLServer
#Region "CONSTRUTORES"
    Public Sub New()

    End Sub
#End Region
#Region "PROPRIEDADE"
    Public Property CodComposicao As Integer
    Public Property NovoCusto As Decimal
    Public Property Rendimento As Integer
    Public Property Configuracao As String
    Public Property DataCadastro As Date
    Public Property PrazoProducao As Integer
    Public Property CodVariacao As Integer
#End Region
#Region "METODOS"
    Public Sub SalvaProducao(Cod_Simples As Integer, Rendimento As Integer, CustoTotal As Decimal, Configuracao As String, TamanhoFinal As String, TamanhoMiolo As String, TamanhoCapa As String, TamanhoRevestimento As String, TamanhoLombada As String, Gap As String, Seixa As String, Elastico As String, FitaCetim As String)
        Dim sql = "INSERT Tbl_Producao (Cod_Simples, 
                                        Rendimento, 
                                        CustoTotal,
                                        Configuracao,
                                        Inativo,
                                        TamanhoFinal,
                                        TamanhoMiolo,
                                        TamanhoCapa,
                                        TamanhoRevestimento,
                                        TamanhoLombada,
                                        Gap,
                                        Seixa,
                                        Elastico,
                                        FitaCetim,
                                        DataCadastro) 
                                VALUES (@Cod_Simples, 
                                        @Rendimento, 
                                        @CustoTotal,
                                        @Configuracao,
                                        0,
                                        @TamanhoFinal,
                                        @TamanhoMiolo,
                                        @TamanhoCapa,
                                        @TamanhoRevestimento,
                                        @TamanhoLombada,
                                        @Gap,
                                        @Seixa,
                                        @Elastico,
                                        @FitaCetim,
                                        GETDATE())"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@Cod_Simples", Cod_Simples),
                    New SqlParameter("@Rendimento", Rendimento),
                    New SqlParameter("@CustoTotal", CustoTotal),
                    New SqlParameter("@Configuracao", Configuracao),
                    New SqlParameter("@TamanhoFinal", TamanhoFinal),
                    New SqlParameter("@TamanhoMiolo", TamanhoMiolo),
                    New SqlParameter("@TamanhoCapa", TamanhoCapa),
                    New SqlParameter("@TamanhoRevestimento", TamanhoRevestimento),
                    New SqlParameter("@TamanhoLombada", TamanhoLombada),
                    New SqlParameter("@Gap", Gap),
                    New SqlParameter("@Seixa", Seixa),
                    New SqlParameter("@Elastico", Elastico),
                    New SqlParameter("@FitaCetim", FitaCetim)
        }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub SalvaDetProducao(CodProd As Integer, CodProducao As Integer, ValorCusto As Decimal, Quantidade As Integer, ValorTotal As Decimal)
        Dim sql = "INSERT INTO              Tbl_CustosProd
                                            (CodProd,
                                            ValorCusto,
                                            Quantidade,
                                            ValorTotal,
                                            CodProducao)
                                    VALUES (@CodProd,
                                            @ValorCusto,
                                            @Quantidade,
                                            @ValorTotal,
                                            @CodProducao)"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodProd", CodProd),
                    New SqlParameter("@ValorCusto", ValorCusto),
                    New SqlParameter("@Quantidade", Quantidade),
                    New SqlParameter("@ValorTotal", ValorTotal),
                    New SqlParameter("@CodProducao", CodProducao)
        }
        ClasseConexao.Operar(sql, parameters)
    End Sub
    Public Sub AtualizaProducao(CodProducao As Integer, CodSimples As Integer, Rendimento As String, CustoTotal As Decimal, Configuracao As String, TamanhoFinal As String, TamanhoMiolo As String, TamanhoCapa As String, TamanhoRevestimento As String, TamanhoLombada As String, Gap As String, Seixa As String, Elastico As String, FitaCetim As String)
        Dim sql = "UPDATE Tbl_Producao  SET Cod_Simples = @CodSimples,
                                            Rendimento = @Rendimento, 
                                            CustoTotal = @CustoTotal, 
                                            Configuracao = @Configuracao,
                                            TamanhoFinal = @TamanhoFinal,
                                            TamanhoMiolo = @TamanhoMiolo,
                                            TamanhoCapa = @TamanhoCapa,
                                            TamanhoRevestimento = @TamanhoRevestimento,
                                            TamanhoLombada = @TamanhoLombada,
                                            Gap = @Gap,
                                            Seixa = @Seixa,
                                            Elastico = @Elastico,
                                            FitaCetim = @FitaCetim
                                    WHERE   CodProducao = @CodProducao"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodProducao", CodProducao),
                    New SqlParameter("@CodSimples", CodSimples),
                    New SqlParameter("@Rendimento", Rendimento),
                    New SqlParameter("@CustoTotal", CustoTotal),
                    New SqlParameter("@Configuracao", Configuracao),
                    New SqlParameter("@TamanhoFinal", TamanhoFinal),
                    New SqlParameter("@TamanhoMiolo", TamanhoMiolo),
                    New SqlParameter("@TamanhoCapa", TamanhoCapa),
                    New SqlParameter("@TamanhoRevestimento", TamanhoRevestimento),
                    New SqlParameter("@TamanhoLombada", TamanhoLombada),
                    New SqlParameter("@Gap", Gap),
                    New SqlParameter("@Seixa", Seixa),
                    New SqlParameter("@Elastico", Elastico),
                    New SqlParameter("@FitaCetim", FitaCetim)
                    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub AtualizaDetProducao(Quantidade As Integer, ValorCusto As Decimal, CodProd As Integer, ValorTotal As Decimal, CodProducao As Integer)
        Dim parameters As SqlParameter() = {
        New SqlParameter("@CodProducao", CodProducao),
        New SqlParameter("@CodProd", CodProd)
    }

        Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_ProducaoDetalhes WHERE CodProducao = @CodProducao AND CodItem = @CodProd", parameters)

        If Tabela.Rows.Count > 0 Then
            Dim parametersUp As SqlParameter() = {
        New SqlParameter("@CodProd", CodProd),
        New SqlParameter("@Quantidade", Quantidade),
        New SqlParameter("@ValorCusto", ValorCusto),
        New SqlParameter("@ValorTotal", ValorTotal),
        New SqlParameter("@CodProducao", CodProducao)
    }
            Dim sql As String = "UPDATE     Tbl_CustosProd SET      
                                                            ValorCusto = @ValorCusto,
                                                            Quantidade = @Quantidade,
                                                            ValorTotal = @ValorTotal
                                                    WHERE   CodProd = @CodProd
                                                    AND     CodProducao = @CodProducao"

            ClasseConexao.Operar(sql, parametersUp)
        Else
            SalvaDetProducao(CodProd, CodProducao, ValorCusto, Quantidade, ValorTotal)
        End If
    End Sub
    Public Sub ExcluiProducao(CodProducao As Integer)
        Dim sql As String = "DELETE FROM Tbl_Producao WHERE CodProducao = @CodProducao"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodProducao", CodProducao)
                    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub ExcluirItemProducao(CodProducao As Integer, Optional CodProd As Integer = 0)
        If CodProd <> 0 Then
            Dim parameters As SqlParameter() = {
            New SqlParameter("@CodProd", CodProd),
            New SqlParameter("@CodProducao", CodProducao)
}
            Dim sql As String = "DELETE FROM Tbl_CustosProd WHERE CodProducao = @CodProducao AND CodProd = @CodProd"
            ClasseConexao.Operar(sql, parameters)
        Else
            Dim sql As String = "DELETE FROM Tbl_CustosProd WHERE CodProducao = @CodProducao"
            Dim parameters As SqlParameter() = {
                        New SqlParameter("@CodProducao", CodProducao)
                        }
            ClasseConexao.Operar(sql, parameters)

        End If
    End Sub
    Public Sub SalvarVariacao(CodSimples As Integer, CodProducao As Integer, CodVariacao As Integer, SKU As String, SKUPrincipal As String, Variante As String, IDProduto As String)
        Dim sql As String = "INSERT INTO tbl_Variacao_Producao (CodSimples,CodProducao,CodVariacao,SKU,SKUPrincipal,Variante,IDProduto) VALUES (@CodSimples,@CodProducao,@CodVariacao,@SKU,@SKUPrincipal,@Variante,@IDProduto)"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodSimples", CodSimples),
                    New SqlParameter("@CodProducao", CodProducao),
                    New SqlParameter("@CodVariacao", CodVariacao),
                    New SqlParameter("@SKU", SKU),
                    New SqlParameter("@SKUPrincipal", SKUPrincipal),
                    New SqlParameter("@Variante", Variante),
                    New SqlParameter("@IDProduto", IDProduto)
                    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Variação incluída com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub ExcluirVariacao(CodVariacao As Integer)
        Dim sql As String = "DELETE FROM tbl_Variacao_Producao WHERE CodVariacao = @CodVariacao"
        Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodVariacao", CodVariacao)
                    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Variação incluída com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
#End Region
#Region "FUNCOES"
    Public Function PesquisaProducao(Status As String, CodProd As Integer, CodSimples As Integer, Produto As String) As DataTable
        Dim sql As New StringBuilder("SELECT * FROM Cs_Producao WHERE 1=1 ")
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

            If CodProd <> 0 Then
                sql.AppendLine("AND CodProducao LIKE @CodProducao")
                parameters.Add(New SqlParameter("@CodProducao", CodProd))
            End If

            If CodSimples <> 0 Then
                sql.AppendLine("AND Cod_Simples LIKE @CodSimples")
                parameters.Add(New SqlParameter("@CodSimples", CodSimples))
            End If

            If Not String.IsNullOrEmpty(Produto) Then
                sql.AppendLine("AND Descricao LIKE @Produto")
                parameters.Add(New SqlParameter("@Produto", "%" & Produto & "%"))
            End If

            sql.AppendLine("ORDER BY Descricao")

            Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

        Catch ex As Exception
            MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Throw
        End Try
    End Function
    ''' <summary>
    ''' Este metodo realiza a consulta de produção no banco de dados.
    ''' </summary>
    ''' <param name="sql">Representa a query sql necessária para a consulta.</param>
    ''' <param name="CodComposicao">Representa o código identificador da composição no banco de dados.</param>
    ''' <returns>Retorna dados solicitado na query sql.</returns>
    Public Function ConsultaProducao(sql As String, Optional CodComposicao As Integer = 0)
        If CodComposicao <> 0 Then
            Dim parameters As SqlParameter() = {
            New SqlParameter("@CodComposicao", CodComposicao)
        }

            Return ClasseConexao.Consultar(sql, parameters)
        Else
            Return ClasseConexao.Consultar(sql, Nothing)
        End If
    End Function
    Public Function ConsultaVariacoes(sql As String, CodProducao As Integer)
        If CodProducao <> 0 Then
            Dim parameters As SqlParameter() = {
            New SqlParameter("@CodProducao", CodProducao)
        }
            Return ClasseConexao.Consultar(sql, parameters)
        Else
            Return ClasseConexao.Consultar(sql, Nothing)
        End If
    End Function
    Public Function ValidarComposicao(CodComposicao As Integer) As Boolean
        Dim sql = "SELECT * FROM Cs_ComposicaoProducao WHERE CodComposicao = @CodComposicao"
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodComposicao", CodComposicao)
        }
        Dim Tabela As DataTable = ClasseConexao.Consultar(sql, parameters)
        If Tabela.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If

    End Function
    Public Function GerarSKU(nomeProduto As String, variacao As String, CodItem As String) As (sku As String, skuPrincipal As String)

        Dim partes As New List(Of String)
        Dim removerProduto As String = " MIOLO"
        Dim novoProduto As String = nomeProduto.Replace(removerProduto, "").Trim()

        Dim removerVaricao As String = " PÁGINAS"
        Dim novaVariacao As String = variacao.Replace(removerVaricao, "").Trim()

        Dim descricao As String = String.Concat(novoProduto, " ", novaVariacao)

        Dim indiceInicio As Integer = -1
        Dim parteARemover As String = ""

        If descricao.IndexOf("BRANCO") > 0 Then
            indiceInicio = descricao.IndexOf("BRANCO")
            parteARemover = "BRANCO"
        ElseIf descricao.IndexOf("KRAFT") > 0 Then
            indiceInicio = descricao.IndexOf("KRAFT")
            parteARemover = "KRAFT"
        ElseIf descricao.IndexOf("MESCLADO") > 0 Then
            indiceInicio = descricao.IndexOf("MESCLADO")
            parteARemover = "MESCLADO"
        End If

        If indiceInicio <> -1 Then
            ' Divide a frase em duas partes: antes e depois da parte removida
            Dim antes As String = descricao.Substring(0, indiceInicio).Trim()
            Dim depois As String = descricao.Substring(indiceInicio + parteARemover.Length).Trim()

            ' Constrói a nova frase com a parte removida no final
            Dim fraseReorganizada As String = (antes & " " & depois & " " & parteARemover).Trim()

            ' Remove espaços extras
            fraseReorganizada = System.Text.RegularExpressions.Regex.Replace(fraseReorganizada, "\s{2,}", " ")

            Dim palavras As String() = fraseReorganizada.Split(" "c) ' Divide a string em palavras

            For Each palavra In palavras
                If palavra = "A5" Or palavra = "A6" Then
                    partes.Add(palavra)
                ElseIf palavra = "060" Or palavra = "100" Or palavra = "160" Or palavra = "200" Then
                    partes.Add(String.Concat("-", palavra))
                ElseIf palavra = "CADERNO" Or palavra = "CAPA" Or palavra = "DURA" Or palavra = "BROCHURA" Or palavra = "QUADRADO" Then
                    partes.Add(palavra.Substring(0, 1))
                ElseIf palavra = "LISO" Or palavra = "PONTILHADO" Or palavra = "QUADRICULADO" Then
                    partes.Add(palavra.Substring(0, 1))
                ElseIf palavra = "BRANCO" Or palavra = "KRAFT" Or palavra = "MESCLADO" Then
                    partes.Add(palavra.Substring(0, 1))
                Else
                    partes.Add(palavra.Substring(0, 2))
                End If
            Next

            Dim parteJuntas As String = String.Join("", partes) ' Junta novamente as palavras

            Dim divisor As String = "-"
            Dim indicePartes As Integer = parteJuntas.IndexOf(divisor)

            If indicePartes <> -1 Then
                ' Divide a frase em duas partes: antes e depois da parte removida
                Dim skuPrincipal As String = parteJuntas.Substring(0, indicePartes).Trim()
                Dim sku As String = parteJuntas.Substring(indicePartes + divisor.Length).Trim()

                ' Constrói a nova frase com a parte removida no final
                Dim skuReorganizada As String = String.Concat(skuPrincipal, sku, CodItem).Trim()

                ' Remove espaços extras
                skuReorganizada = System.Text.RegularExpressions.Regex.Replace(skuReorganizada, "\s{2,}", " ")

                Return (skuReorganizada, skuPrincipal) ' Retorna os dois valores como tuple
            Else
                MessageBox.Show("Divisor a ser removida não encontrada na frase" & vbCrLf, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            MessageBox.Show("Parte a ser removida não encontrada na frase" & vbCrLf, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Function
        End If
    End Function

#End Region
End Class
