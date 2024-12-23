Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports System.Data
Imports System.Text
Imports System.Xml
Imports Xceed.Wpf.Toolkit

Namespace Classes.Compras
    Public Class clsNotaEntrada
        Inherits clsCompras
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property NaturezaOperacao As String
        Public Property DataEmissao As String
        Public Property NumeroNota As Integer
        Public Property Serie As Integer
        Public Property ChaveAcesso As String
#End Region
#Region "CONSTRUTORES"
        Public Sub New()

        End Sub
        Public Sub New(_codigo As String, _item As String, _fator As Integer, _custoatual As String)
            CodItem = _codigo
            Item = _item
            Fator = _fator
            PrecoCompra = _custoatual
        End Sub
#End Region
#Region "METODOS"
        Public Sub SalvarNotaFiscalEntrada(DataCompra As Date, DataEmissao As Date, Fornecedor As Integer, Cobranca As Integer, FormaPgto As Integer, Frete As Decimal, Desconto As Decimal, Total As Decimal, DataEntrada As Date, Serie As String, ChaveAcesso As String, Operacao As Integer, Optional NrPedido As Integer? = Nothing, Optional NotaFiscal As Integer? = Nothing)
            Dim sql As String = "INSERT INTO Tbl_Compras (NrPedido,
                                            DataCompra, 
        	                                DataEmissao, 
        	                                Fornecedor,
        	                                Cobranca,
        	                                FormaPgto,
        	                                NotaFiscal,
        	                                Frete,
        	                                Desconto,
        	                                DataCadastro,
        	                                Total,
        	                                DataEntrada,
        	                                Serie,
        	                                ChaveAcesso,
                                            Status,
                                            Operacao)
                              VALUES		(@NrPedido,
                                             @DataCompra, 
        	                                @DataEmissao, 
        	                                @Fornecedor,
        	                                @Cobranca,
        	                                @FormaPgto,
        	                                @NotaFiscal,
        	                                @Frete,
        	                                @Desconto,
        	                                GETDATE(),
        	                                @Total,
        	                                @DataEntrada,
        	                                @Serie,
        	                                @ChaveAcesso,
                                            0,
                                            @Operacao)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@DataCompra", DataCompra),
                         New SqlParameter("@DataEmissao", DataEmissao),
                         New SqlParameter("@Fornecedor", Fornecedor),
                         New SqlParameter("@Cobranca", Cobranca),
                         New SqlParameter("@FormaPgto", FormaPgto),
                         New SqlParameter("@Frete", Frete),
                         New SqlParameter("@Desconto", Desconto),
                         New SqlParameter("@Total", Total),
                         New SqlParameter("@DataEntrada", DataEntrada),
                         New SqlParameter("@Serie", Serie),
                         New SqlParameter("@ChaveAcesso", ChaveAcesso),
                         New SqlParameter("@Operacao", Operacao)
        }

            If NrPedido <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@NrPedido", NrPedido.Value)
            Else
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@NrPedido", DBNull.Value)
            End If

            If NotaFiscal <> 0 Then
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@NotaFiscal", NotaFiscal.Value)
            Else
                Array.Resize(parameters, parameters.Length + 1)
                parameters(parameters.Length - 1) = New SqlParameter("@NotaFiscal", DBNull.Value)
            End If

            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub SalvaDetNotaEntrada(NrPedidoDet As String, CodSimples As Integer, ValorUnit As Decimal, Quantidade As Integer, ValorTotal As Decimal, CodCompra As Integer)
            Dim sql As String = "INSERT INTO Tbl_ComprasDet (NrPedidoDet,
                                                            Cod_Simples,
                                                            Qtde,
                                                            ValorUnit,
                                                            Total,
                                                            CodCompra) 
                                                   VALUES(@NrPedidoDet,
                                                            @CodSimples,
                                                            @Quantidade,
                                                            @ValorUnit,
                                                            @ValorTotal,
                                                            @CodCompra)"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@NrPedidoDet", NrPedidoDet),
                         New SqlParameter("@CodSimples", CodSimples),
                         New SqlParameter("@ValorUnit", ValorUnit),
                         New SqlParameter("@Quantidade", Quantidade),
                         New SqlParameter("@ValorTotal", ValorTotal),
                         New SqlParameter("@CodCompra", CodCompra)
            }

            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub ExcluirDetNotaEntrada(CodCompra As Integer)
            Dim sql As String = "DELETE FROM Tbl_ComprasDet WHERE CodCompra = @CodCompra"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodCompra", CodCompra)
            }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub ExcluirNotaEntrada(Codigo As Integer)
            Dim sql As String = "DELETE FROM Tbl_Compras WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@Codigo", Codigo)
            }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluirItemNota(CodCompra As Integer, CodItem As Integer)
            Dim parameters As SqlParameter() = {
                        New SqlParameter("@CodCompra", CodCompra),
                        New SqlParameter("@CodSimples", CodItem)
            }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_ComprasDetalhes WHERE Codigo = @CodCompra AND Cod_Simples = @CodSimples", parameters)

            If Tabela.Rows.Count > 0 Then
                Dim parametersDel As SqlParameter() = {
                        New SqlParameter("@CodCompra", CodCompra),
                        New SqlParameter("@CodSimples", CodItem)
            }
                Dim sql As String = "DELETE FROM Tbl_ComprasDet WHERE CodCompra = @CodCompra AND Cod_Simples = @CodSimples"
                ClasseConexao.Operar(sql, parametersDel)
            Else
                Exit Sub
            End If
        End Sub
        Public Sub AtualizaNotaEntrada(CodCompra As Integer, DataCompra As Date, DataEmissao As Date, Fornecedor As Integer, Cobranca As Integer, FormaPgto As Integer, NotaFiscal As String, Frete As Decimal, Desconto As Decimal, Total As Decimal, DataEntrada As Date, Serie As Integer, ChaveAcesso As String)
            Dim sql As String = "UPDATE Tbl_Compras 
                               SET DataCompra = @DataCompra,
						            DataEmissao = @DataEmissao,
                                    DataEntrada = @DataEntrada,
                                    Fornecedor = @Fornecedor,
                                    NotaFiscal = @NotaFiscal,
                                    Serie = @Serie,
                                    Frete = @Frete,
							        Total = @Total,
							        Cobranca = @Cobranca,
							        FormaPgto = @FormaPgto,
                                    ChaveAcesso = @ChaveAcesso
                               WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@Codigo", CodCompra),
                         New SqlParameter("@DataCompra", DataCompra),
                         New SqlParameter("@DataEmissao", DataEmissao),
                         New SqlParameter("@Fornecedor", Fornecedor),
                         New SqlParameter("@Cobranca", Cobranca),
                         New SqlParameter("@FormaPgto", FormaPgto),
                         New SqlParameter("@NotaFiscal", NotaFiscal),
                         New SqlParameter("@Frete", Frete),
                         New SqlParameter("@Desconto", Desconto),
                         New SqlParameter("@Total", Total),
                         New SqlParameter("@DataEntrada", DataEntrada),
                         New SqlParameter("@Serie", Serie),
                         New SqlParameter("@ChaveAcesso", ChaveAcesso)
            }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaDetNota(CodCompra As Integer?, NrPedidoDet As Integer, ValorUnit As Decimal?, Quantidade As Integer?, CodSimples As Integer, ValorTotal As Decimal?)
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodCompra", CodCompra),
                         New SqlParameter("@CodSimples", CodSimples)
        }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT * FROM Cs_ComprasDetalhes WHERE Codigo = @CodCompra AND Cod_Simples = @CodSimples", parameters)
            If Tabela.Rows.Count > 0 Then
                Dim parametersUp As SqlParameter() = {
                         New SqlParameter("@CodCompra", CodCompra),
                         New SqlParameter("@CodSimples", CodSimples),
                         New SqlParameter("@Quantidade", Quantidade),
                         New SqlParameter("@ValorUnit", ValorUnit),
                         New SqlParameter("@ValorTotal", ValorTotal)
        }

                Dim sql As String = "UPDATE Tbl_ComprasDet  SET     Qtde = @Quantidade,
                                                                ValorUnit = @ValorUnit,
                                                                Total = @ValorTotal
                                                        WHERE   CodCompra = @CodCompra
                                                        AND     Cod_Simples = @CodSimples"

                ClasseConexao.Operar(sql, parametersUp)
            ElseIf Tabela.Rows.Count = 0 Then
                SalvaDetNotaEntrada(NrPedidoDet, CodSimples, ValorUnit, Quantidade, ValorTotal, CodCompra)
            End If
        End Sub
        Public Sub AtualizaPagNota(ValorPago As Decimal, Cobranca As Integer, FormaPagto As Integer, CodCompra As Integer, Complemento As String, Desconto As Integer, Frete As Integer, Entidade As Integer)
            Dim sql As String = "UPDATE Tbl_Compras     SET     Cobranca = @COBRANCA, 
                                                                FormaPgto = @FORMAPAGTO, 
                                                                Total = @VALORPAGO, 
                                                                Desconto = @DESCONTO, 
                                                                Frete = @FRETE, 
                                                                Entidade = @ENTIDADE 
                                                        WHERE   Codigo = @CODCOMPRA"
            Dim parameters As SqlParameter() = {
                            New SqlParameter("@CODCOMPRA", CodCompra),
         New SqlParameter("@COBRANCA", Cobranca),
         New SqlParameter("@FORMAPAGTO", FormaPagto),
         New SqlParameter("@VALORPAGO", ValorPago),
         New SqlParameter("@COMPLEMENTO", Complemento),
         New SqlParameter("@DESCONTO", Desconto),
         New SqlParameter("@FRETE", Frete),
         New SqlParameter("@ENTIDADE", Entidade)
        }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub AtualizaStatusNotaEntrada(CodNotaEntrada As Integer, Status As Integer)
            Dim sql As String = "UPDATE Tbl_Compras SET Status = @Status WHERE Codigo = @CodNotaEntrada"
            Dim parameters As SqlParameter() = {
                New SqlParameter("@CodNotaEntrada", CodNotaEntrada),
                New SqlParameter("@Status", Status)
        }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Nota Fiscal baixada com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
#End Region
#Region "FUNCOES"
        Public Function PesquisaNotaEntrada(Status As String, CodPedido As String, Fornecedor As String, FormaPagto As String, Cobranca As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_Compras WHERE 1=1")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Em Aberto"
                        sql.AppendLine("AND Status = 0")
                    Case "Concluído"
                        sql.AppendLine("AND Status = 1")
                    Case "Finalizado"
                        sql.AppendLine("AND Status = 2")
                    Case "Todos"
                        sql.AppendLine("AND Status IS NOT NULL")
                End Select

                If Not String.IsNullOrEmpty(CodPedido) Then
                    sql.AppendLine("AND CodPedido = @CodPedido")
                    parameters.Add(New SqlParameter("@CodPedido", CodPedido))
                End If

                If Not String.IsNullOrEmpty(Fornecedor) Then
                    sql.AppendLine("AND Fornecedor LIKE @Fornecedor")
                    parameters.Add(New SqlParameter("@Fornecedor", Fornecedor))
                End If

                If Not String.IsNullOrEmpty(FormaPagto) Then
                    sql.AppendLine("AND Forma_Pgto LIKE @FormaPagto")
                    parameters.Add(New SqlParameter("@FormaPgto", FormaPagto))
                End If

                If Not String.IsNullOrEmpty(Cobranca) Then
                    sql.AppendLine("AND Cobranca LIKE @Cobranca")
                    parameters.Add(New SqlParameter("@Cobranca", Cobranca))
                End If

                sql.AppendLine("ORDER BY DataCompra DESC")
                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function

        Public Function ConsultaNotaEntrada(sql As String, Optional CodNota As Integer = 0, Optional CodPedido As Integer = 0)
            If CodNota <> 0 Then
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodNotaEntrada", CodNota)
                            }
                Return ClasseConexao.Consultar(sql, parameters)
            ElseIf CodPedido <> 0 Then
                Dim parameters As SqlParameter() = {
                         New SqlParameter("@NRPedido", CodPedido)
                            }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
        Public Function ConsultaItensFornecedor(Sql As String, CodItem As String, CodFornecedor As Integer)
            Dim parameters As SqlParameter() = {
                         New SqlParameter("@CodSimples", CodItem),
                         New SqlParameter("@CodFornecedor", CodFornecedor)
                            }
            Return ClasseConexao.Consultar(Sql, parameters)
        End Function

        Public Function ImportarXML(xmlDoc As XmlDocument, Grid As DataGridView, CodForn As Integer, Documento As String, ByRef DadosNotaEntrada As clsNotaEntrada)
            Dim nsmgr As New XmlNamespaceManager(xmlDoc.NameTable)
            Dim numNFe, nNatop, nNota, nSerie, nDataEmi, cProd, xProd, qCom, vUnCom, vProd, vFrete, vDesc, vNF, CustoAtual As String

            Try
                nsmgr.AddNamespace("nfe", "http://www.portalfiscal.inf.br/nfe")

                Dim infNFeNode As XmlNode = xmlDoc.SelectSingleNode("//nfe:infNFe", nsmgr)
                If infNFeNode IsNot Nothing Then
                    numNFe = infNFeNode.Attributes("Id").Value.Substring(3)

                    Dim ideNode As XmlNode = infNFeNode.SelectSingleNode("nfe:ide", nsmgr)
                    If ideNode IsNot Nothing Then
                        nNatop = ideNode.SelectSingleNode("nfe:natOp", nsmgr).InnerText
                        nNota = ideNode.SelectSingleNode("nfe:nNF", nsmgr).InnerText
                        nSerie = ideNode.SelectSingleNode("nfe:serie", nsmgr).InnerText
                        nDataEmi = ideNode.SelectSingleNode("nfe:dhEmi", nsmgr).InnerText
                    End If

                    'Cabeçalho
                    Dim emitenteNode As XmlNode = infNFeNode.SelectSingleNode("nfe:emit", nsmgr)
                    If emitenteNode IsNot Nothing Then
                        Dim cnpjEmitente As String = emitenteNode.SelectSingleNode("nfe:CNPJ", nsmgr).InnerText
                        If Documento <> cnpjEmitente Then
                            MessageBox.Show("CPFCNPJ do emitente difere do cadastrado!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Function
                        ElseIf Documento = "" Then
                            MessageBox.Show("Selecione o emitente para importar a nota fiscal", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Function
                        ElseIf Documento = cnpjEmitente Then
                            DadosNotaEntrada.NaturezaOperacao = nNatop
                            DadosNotaEntrada.DataEmissao = nDataEmi
                            DadosNotaEntrada.NumeroNota = nNota
                            DadosNotaEntrada.Serie = nSerie
                            DadosNotaEntrada.ChaveAcesso = numNFe
                        End If
                    Else
                        MessageBox.Show("Emitente não encontrado no XML!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Function
                    End If

                    'Itens
                    Dim xmlNodeList As XmlNodeList = xmlDoc.GetElementsByTagName("det")
                    Dim nodeCount As Integer = xmlNodeList.Count 'Conta quantidade de itens no nó

                    For Each node As XmlNode In xmlNodeList
                        Dim value = node.Attributes.GetNamedItem("nItem").Value

                        cProd = Convert.ToString(node("prod")("cProd").InnerText)
                        xProd = Convert.ToString(node("prod")("xProd").InnerText)
                        qCom = Convert.ToString(node("prod")("qCom").InnerText)
                        vUnCom = Convert.ToString(node("prod")("vUnCom").InnerText)
                        vProd = Convert.ToString(node("prod")("vProd").InnerText)
                        CustoAtual = 0

                        vUnCom = vUnCom.Replace(".", ",")
                        qCom = qCom.Replace(".", ",")
                        vProd = vProd.Replace(".", ",")

                        Dim Tabela As DataTable = ConsultaItensFornecedor("SELECT * FROM Cs_ItemFornecedor WHERE CodProd LIKE @CodSimples AND CodForn = @CodFornecedor", cProd, CodForn)
                        For Each linha As DataRow In Tabela.Rows
                            cProd = Tabela.Rows(0)("CodItem").ToString()
                            xProd = Tabela.Rows(0)("Descricao").ToString()
                            qCom = Tabela.Rows(0)("Fator").ToString()
                            CustoAtual = Tabela.Rows(0)("PrecoCompra").ToString()
                            vUnCom = (vUnCom / qCom)
                            vProd = (vUnCom * qCom)
                        Next

                        ' Verifica se o item já existe no DataGridView
                        Dim itemExists As Boolean = False
                        For Each row As DataGridViewRow In Grid.Rows
                            If row.Cells("CodSimples").Value IsNot Nothing AndAlso row.Cells("CodSimples").Value.ToString() = cProd Then
                                itemExists = True
                                Exit For
                            End If
                        Next

                        ' Adiciona o item ao DataGridView se ele não existir
                        If Not itemExists Then
                            Grid.Rows.Add(cProd, Nothing, xProd, Convert.ToDecimal(vUnCom), CInt(qCom), Convert.ToDecimal(vProd), CustoAtual.Replace(".", ","))
                        End If
                    Next

                    'Total, Frete e Desconto
                    Dim totalNode As XmlNode = infNFeNode.SelectSingleNode("nfe:total", nsmgr)
                    If totalNode IsNot Nothing Then
                        Dim ICMSNode As XmlNode = totalNode.SelectSingleNode("nfe:ICMSTot", nsmgr)
                        If ICMSNode IsNot Nothing Then
                            vFrete = ICMSNode.SelectSingleNode("nfe:vFrete", nsmgr).InnerText
                            vDesc = ICMSNode.SelectSingleNode("nfe:vDesc", nsmgr).InnerText
                            vNF = ICMSNode.SelectSingleNode("nfe:vNF", nsmgr).InnerText
                            DadosNotaEntrada.Frete = vFrete.Replace(".", ",")
                            DadosNotaEntrada.Desconto = vDesc.Replace(".", ",")
                            DadosNotaEntrada.Total = vNF.Replace(".", ",")
                        End If
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show("Não foi possível importar o XML!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function
#End Region
    End Class
End Namespace


