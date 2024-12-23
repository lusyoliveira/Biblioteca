Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports Xceed.Wpf.Toolkit

Namespace Classes.Financeiro
    Public Class clsFinanceiro
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property CodReceber As Integer
        Public Property CodPagar As Integer
        Public Property CodFormaPagto As Integer
        Public Property CodCobranca As Integer
        Public Property CodConta As Integer
        Public Property FormaPagto As String
        Public Property Cobranca As String
        Public Property NomeConta As String
        Public Property NomeBanco As String
#End Region
#Region "METODOS"
        ''' <summary>
        ''' Este método realiza o calculo das parcelas de pagamentos e recebimentos.
        ''' </summary>
        ''' <param name="ValorTotal"></param>
        ''' <param name="Data"></param>
        ''' <param name="NumeroParcelas"></param>
        ''' <param name="DadosParcela"></param>
        Public Sub CalculaParcelas(ValorTotal As Decimal, Data As DateTime, NumeroParcelas As Integer, ByRef DadosParcela As clsMovimentoBancario)
            Dim valorParcela As Decimal = FormatCurrency(ValorTotal / NumeroParcelas)
            Dim ListaParcela = New List(Of clsMovimentoBancario)

            Try
                For i = 0 To Val(NumeroParcelas) - 1
                    Dim novaParcela As DateTime
                    novaParcela = Data.AddDays(i * 30)
                    If novaParcela.DayOfWeek = DayOfWeek.Sunday Then
                        novaParcela = novaParcela.AddDays(1)
                    ElseIf novaParcela.DayOfWeek = DayOfWeek.Saturday Then
                        novaParcela = novaParcela.AddDays(2)
                    End If
                    DadosParcela.NParcelas = i + 1
                    DadosParcela.DataVencto = Mid(novaParcela.ToString, 1, 10)
                    DadosParcela.ValorParcela = (Decimal.Parse(valorParcela).ToString("N2"))
                Next
            Catch ex As Exception
                MessageBox.Show("Erro ao calcular valor", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
        ''' <summary>
        ''' Este metódo registra o banco no banco de dados.
        ''' </summary>
        ''' <param name="CodBanco">Represente o código do banco</param>
        ''' <param name="Banco">Representa o nome do banco.</param>
        ''' <param name="ISPB">Representa código ISPB do banco</param>
        Public Sub SalvarBanco(CodBanco As Integer, Banco As String, ISPB As Integer)

            Dim sql As String = "INSERT INTO Tbl_Bancos (CodigoBanco,Banco,ISPB) VALUES (@CODBANCO, @BANCO, @ISPB)"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@CODBANCO", CodBanco),
        New SqlParameter("@BANCO", Banco),
        New SqlParameter("@ISPB", ISPB)
    }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        ''' <summary>
        ''' Este metódo atualiza dados do banco no banco de dados.
        ''' </summary>
        ''' <param name="Codigo">Código identificador do banco no banco de dados</param>
        ''' <param name="CodBanco">Represente o código do banco</param>
        ''' <param name="Banco">Representa o nome do banco.</param>
        ''' <param name="ISPB">Representa código ISPB do banco</param>
        Public Sub AtualizarBanco(Codigo As Integer, CodBanco As Integer, Banco As String, ISPB As Integer)
            Dim sql As String = "UPDATE Tbl_Bancos SET CodigoBanco= @CODBANCO, Banco = @Banco, ISPB = @ISPB WHERE Codigo = @CODIGO"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@CODIGO", Codigo),
        New SqlParameter("@CODBANCO", CodBanco),
        New SqlParameter("@BANCO", Banco),
        New SqlParameter("@ISPB", ISPB)
    }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        ''' <summary>
        ''' Este metódo exclui um banco do banco de dados.
        ''' </summary>
        ''' <param name="Codigo">Código identificador do banco no banco de dados</param>
        Public Sub ExcluirBanco(Codigo As Integer)
            Dim sql As String = "DELETE FROM Tbl_Bancos WHERE Codigo = @CODIGO"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@CODIGO", Codigo)
    }
            ClasseConexao.Operar(sql, parameters)

        End Sub
        ''' <summary>
        ''' Este metódo registra a cobrança no banco de dados.
        ''' </summary>
        ''' <param name="Cobranca">Represente o nome da cobrança.</param>
        Public Sub SalvarCobranca(Cobranca As String)
            Dim sql As String = "INSERT INTO Tbl_Cobranca (Cobranca) VALUES (@COBRANCA)"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@COBRANCA", Cobranca)
    }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        ''' <summary>
        ''' Este metódo registra a cobrança no banco de dados.
        ''' </summary>
        ''' <param name="Codigo">Código identificador da cobrança no banco de dados.</param>
        ''' <param name="Cobranca">Represente o nome da cobrança.</param>
        Public Sub AtualizarCobranca(Codigo As Integer, Cobranca As String)
            Dim sql As String = "UPDATE Tbl_Cobranca SET Cobranca = @COBRANCA WHERE Codigo = @CODIGO"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@CODIGO", Codigo),
        New SqlParameter("@COBRANCA", Cobranca)
    }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        ''' <summary>
        ''' Este metódo exclui uma cobrança do banco de dados.
        ''' </summary>
        ''' <param name="Codigo">Código identificador da cobrança no banco de dados.</param>
        Public Sub ExcluirCobranca(Codigo As Integer)
            Dim sql As String = "DELETE FROM Tbl_Cobranca WHERE Codigo = @CODIGO"
            Dim parameters As SqlParameter() = {
            New SqlParameter("@CODIGO", Codigo)
    }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        ''' <summary>
        ''' Este metódo registra a forma de pagamento no banco de dados.
        ''' </summary>
        ''' <param name="Forma_Pgto">Representa o nome da forma de pagamento.</param>
        ''' <param name="DiasParc">Representa os dias de parcelamentos.</param>
        ''' <param name="Dias">Representa q quantidade de dias da parcela.</param>
        ''' <param name="NParcelas">Representa o número de parcelas.</param>
        Public Sub SalvarFormaPagto(Forma_Pgto As String, DiasParc As Integer, Dias As Integer, NParcelas As Integer)
            Dim sql As String = "INSERT INTO tbl_FormaPgto (Forma_Pgto,Forma_Pgto_Dias1parc,Forma_Pgto_Dias,NParcelas) VALUES (@FORMPAGTO,@DIASPARC,@DIAS,@NPARCELAS)"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@FORMPAGTO", Forma_Pgto),
        New SqlParameter("@DIASPARC", DiasParc),
        New SqlParameter("@DIAS", Dias),
        New SqlParameter("@NPARCELAS", NParcelas)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo registra a forma de pagamento no banco de dados.
        ''' </summary>
        ''' <param name="CodForma">Representa o código identificador da forma de pagamento</param>
        ''' <param name="Forma_Pgto">Representa o nome da forma de pagamento.</param>
        ''' <param name="DiasParc">Representa os dias de parcelamentos.</param>
        ''' <param name="Dias">Representa q quantidade de dias da parcela.</param>
        ''' <param name="NParcelas">Representa o número de parcelas.</param>
        Public Sub AtualizarFormaPagto(CodForma As Integer, Forma_Pgto As String, DiasParc As Integer, Dias As Integer, NParcelas As Integer)
            Dim sql As String = "UPDATE tbl_FormaPgto SET Forma_Pgto = @FORMPAGTO, Forma_Pgto_Dias1parc = @DIASPARC, Forma_Pgto_Dias = @DIAS, NParcelas = @NPARCELAS WHERE Cod_Pgto = @CODFORMA"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@CODFORMA", CodForma),
        New SqlParameter("@FORMPAGTO", Forma_Pgto),
        New SqlParameter("@DIASPARC", DiasParc),
        New SqlParameter("@DIAS", Dias),
        New SqlParameter("@NPARCELAS", NParcelas)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo exclui uma forma de pagamento do banco de dados.
        ''' </summary>
        ''' <param name="CodForma">Código identificador do forma de pagamento no banco de dados</param>
        Public Sub ExcluirFormaPagto(CodForma As Integer)
            Dim sql As String = "DELETE FROM tbl_FormaPgto WHERE Cod_Pgto = @CODFORMA"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@CODFORMA", CodForma)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo registra um tipo de documento no banco de dados.
        ''' </summary>
        ''' <param name="TipoDocumento">Represente a descrição do tipo de documento.</param>
        ''' <param name="CreDeb">Represente se o tipo de documento é do tipo débito ou crédito.</param>
        Public Sub SalvarTipoDoc(TipoDocumento As String, CreDeb As String)
            Dim sql As String = "INSERT INTO tbl_TipoDocumento (TipoDocumento,CreDeb) VALUES (@TIPODOC,@CREDDEB)"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@TIPODOC", TipoDocumento),
        New SqlParameter("@CREDDEB", CreDeb)
   }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo atualiza um tipo de documento no banco de dados.
        ''' </summary>
        ''' <param name="CodTipoDoc">Representa o código do tipo de documento.</param>
        ''' <param name="TipoDocumento">Represente a descrição do tipo de documento.</param>
        ''' <param name="CreDeb">Represente se o tipo de documento é do tipo débito ou crédito.</param>
        Public Sub AtualizarTipoDoc(CodTipoDoc As Integer, TipoDocumento As String, CreDeb As String)
            Dim sql As String = "UPDATE tbl_TipoDocumento SET TipoDocumento = @TIPODOC, CreDeb = @CREDDEB WHERE Codigo = @CODIGO"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@CODIGO", CodTipoDoc),
        New SqlParameter("@TIPODOC", TipoDocumento),
        New SqlParameter("@CREDDEB", CreDeb)
   }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        ''' <summary>
        ''' Este metódo realiza a exclusão de um tipo de documento no banco de dados.
        ''' </summary>
        ''' <param name="CodTipoDoc">Representa o código do tipo de documento.</param>
        Public Sub ExcluirTipoDoc(CodTipoDoc As Integer)
            Dim sql As String = "DELETE FROM tbl_TipoDocumento WHERE Codigo = @CODIGO"
            Dim parameters As SqlParameter() = {
        New SqlParameter("@CODIGO", CodTipoDoc)
   }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub TransferirSaldo(CodConta As Integer, Valor As Decimal, Movimento As String, TipoDocumento As Integer, Complemento As String)
            Dim sql As String = "INSERT INTO tbl_MovimentoBancario    (CodTitulo,
                                                                DataMovimento,
                                                                CodConta,
                                                                Valor,
                                                                Movimento,
                                                                TipoDocumento,
                                                                Complemento)
                                                        VALUES (CONVERT(VARCHAR, GETDATE(), 12),                                            
                                                                GETDATE(),
                                                                @CodConta,
                                                                @Valor,
                                                                @Movimento,
                                                                @TipoDocumento,
                                                                @Complemento)"
            Dim parameters As SqlParameter() = {
              New SqlParameter("@CodConta", CodConta),
              New SqlParameter("@Valor", Valor),
              New SqlParameter("@Movimento", Movimento),
              New SqlParameter("@TipoDocumento", TipoDocumento),
              New SqlParameter("@Complemento", Complemento)
    }
            ClasseConexao.Operar(sql, parameters)
        End Sub
#End Region
#Region "FUNÇÕES"
        ''' <summary>
        ''' Esta função realiza a consulta dos dados  do banco.
        ''' </summary>
        ''' <returns>Retorna os dados do banco</returns>
        Public Function ConsultaBanco(sql As String)
            Return ClasseConexao.Consultar(sql, Nothing)
        End Function
        ''' <summary>
        ''' Esta função realiza a consulta dos dados de cobrança.
        ''' </summary>
        ''' <returns>Retorna os dados da cobrança</returns>
        Public Function ConsultaCobranca(sql As String)
            Return ClasseConexao.Consultar(sql, Nothing)
        End Function
        ''' <summary>
        ''' Esta função realiza a consulta dos dados de forma de pagamento.
        ''' </summary>
        ''' <param name="sql">Query sql com os dados de consulta.</param>
        ''' <returns>Retorna os dados da forma de pagamento</returns>
        Public Function ConsultaFormaPagto(sql As String, Optional CodFormaPagto As Integer = 0)
            If CodFormaPagto <> 0 Then
                Dim parameters As SqlParameter() = {
                New SqlParameter("@CodFormaPagto", CodFormaPagto)
            }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
        ''' <summary>
        ''' Esta função consulta os dados do tipo de documento no banco de dados.
        ''' </summary>
        ''' <param name="sql">Query sql necessária para a consulta</param>
        ''' <returns>Retorna os dados do tipo de documento.</returns>
        Public Function ConsultaTipoDoc(sql As String, Optional CodTipoDoc As Integer = 0)
            If CodTipoDoc <> 0 Then
                Dim parameters As SqlParameter() = {
                New SqlParameter("@CodTipoDoc", CodTipoDoc)
            }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
#End Region
    End Class

End Namespace
