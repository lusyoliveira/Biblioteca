Imports System.Data
Imports System.Text
Imports Biblioteca.Classes.Conexao
Imports Biblioteca.Classes.Produtos
Imports Microsoft.Data.SqlClient
Imports Xceed.Wpf.Toolkit
Namespace Classes.Compras
    ''' <summary>
    ''' Esta classe representa todos o mtódos e funções que manipulam o estoque de um item.
    ''' </summary>
    Public Class clsEstoque
        Inherits clsItem
        Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
        Public Property CodInventario As Integer
        Public Property Lote As Integer
        Public Property CodMotivo As Integer
        Public Property Motivo As String
        Public Property Operacao As String
        Public Property DataInventario As Date
        Public Property Tipo As String
        Public Property Status As String
        Public Property EstMinimo As Integer
        Public Property EstMaximo As Integer
        Public Property MediaConsumo As Integer
        Public Property TipoInventario As String
        Public Property CodTipoInventario As Integer
        Public Property EstoqueAnterior As Integer
        Public Property EstoqueAtual As Integer
        Public Property DataMovimento As Date
        Public Property DataInicial As Date
        Public Property DataFinal As Date
#End Region
#Region "METODOS"
        ''' <summary>
        ''' Este metódo executa uma procedure no banco de dados que atualiza o estoque de um item.
        ''' </summary>
        ''' <param name="CodItem">Representa o código do item</param>
        ''' <param name="Motivo">Representa o código do motivo de alteração de estoque.</param>
        ''' <param name="Quantidade">Representa a quantidade movimentada do item.</param>
        Public Sub AlteraEstoqueItem(CodItem As Integer, Motivo As Integer, Quantidade As Integer)
            Dim parameters As SqlParameter() = {
                    New SqlParameter("@COD_SIMPLES", CodItem),
                    New SqlParameter("@MOTIVO", Motivo),
                    New SqlParameter("@QTDE_MOVIMENTADA", Quantidade)
                    }
            ClasseConexao.ExecutarProcedure("spAlteraHistoricoEstoque", parameters)
        End Sub
        ''' <summary>
        ''' Este metódo executa uma procedure que volta o estoque do item para a quantidade anterior.
        ''' </summary>
        ''' <param name="CodItem"></param>
        Public Sub VoltaEstoqueItem(CodItem As Integer)
            Dim parameters As SqlParameter() = {
                    New SqlParameter("@COD_SIMPLES", CodItem)
                    }
            ClasseConexao.ExecutarProcedure("spVoltaHistoricoEstoque", parameters)
        End Sub
        Public Sub ExcluiHistoricoEstoque(CodSimples As Integer)
            Dim sql As String = "DELETE FROM Tbl_HistoricoEstoque WHERE Cod_Simples = @CodSimples"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodSimples", CodSimples)
    }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub AtualizaStatusInventario(CodInventario As Integer, Status As Integer)
            Dim sql As String = "UPDATE Tbl_Inventario SET Status = @Status WHERE Codigo = @CodInventario"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodInventario", CodInventario),
                     New SqlParameter("@Status", Status)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Inventário " & CodInventario & " finalizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub SalvarInventario(Descricao As String, DataInventario As Date, Status As Integer, TipoInventario As Integer)
            Dim sql As String = "INSERT INTO  tbl_Inventario (Descricao, 
                                                            DataInventario, 
                                                            Status,
                                                            Tipo_Inventario,
                                                            DataCriacao) 
                                                VALUES      (@Descricao, 
                                                            @DataInventario, 
                                                            @Status,
                                                            @TipoInventario,
                                                            GETDATE())"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@Descricao", Descricao),
                     New SqlParameter("@DataInventario", DataInventario),
                     New SqlParameter("@Status", Status),
                     New SqlParameter("@TipoInventario", TipoInventario)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub SalvarItemInventario(CodSimples As Integer, Quantidade As Integer, CodInventario As Integer)
            Dim sql As String = "INSERT INTO  Tbl_Inventario_Item (Cod_Simples,
                                                            Quantidade,
                                                            Cod_Inventario) 
                                                    VALUES  (@CodSimples,
                                                            @Quantidade,
                                                            @CodInventario)"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodSimples", CodSimples),
                     New SqlParameter("@Quantidade", Quantidade),
                     New SqlParameter("@CodInventario", CodInventario)
    }
            ClasseConexao.Operar(sql, parameters)
        End Sub
        Public Sub AtualizaInventario(CodInventario As Integer, Descricao As String, Data As Date, Status As Boolean, TipoInventario As Integer)
            Dim sql As String = "UPDATE Tbl_Inventario SET  Descricao = @Descricao,
                                                    DataInventario = @Data,
                                                    Status = @Status,
                                                    Tipo_Inventario = @TipoInventario,
                                                    DataAlteracao = GETDATE()
                                                    WHERE Codigo = @CodInventario"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodInventario", CodInventario),
                     New SqlParameter("@Descricao", Descricao),
                     New SqlParameter("@Data", Data),
                     New SqlParameter("@Status", Status),
                     New SqlParameter("@TipoInventario", TipoInventario)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaItemInventario(CodSimples As Integer, Quantidade As Integer, CodInventario As Integer)
            Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodSimples", CodSimples),
                    New SqlParameter("@CodInventario", CodInventario)
        }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT COUNT(Cod_Simples) FROM CS_InventarioDetalhes WHERE Cod_Inventario = @CodInventario AND Cod_Simples = @CodSimples", parameters)
            If Tabela.Rows.Count > 0 Then
                Dim sql As String = "UPDATE Tbl_Inventario_Item SET Quantidade = @Quantidade WHERE Cod_Inventario = @CodInventario AND Cod_Simples = @CodSimples"
                Dim UpParameters As SqlParameter() = {
                    New SqlParameter("@CodSimples", CodSimples),
                    New SqlParameter("@CodInventario", CodInventario),
                    New SqlParameter("@Quantidade", Quantidade)
    }
                ClasseConexao.Operar(sql, UpParameters)
            Else
                SalvarItemInventario(CodSimples, Quantidade, CodInventario)
            End If
        End Sub
        Public Sub ExcluirInventario(CodInventario As Integer)
            Dim sqlIem As String = "DELETE FROM Tbl_Inventario_Item WHERE Cod_Inventario = @CodInventario"
            Dim sql As String = "DELETE FROM tbl_Inventario WHERE Codigo = @CodInventario"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodInventario", CodInventario)
    }
            ClasseConexao.Operar(sqlIem, parameters)
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluiItemInventario(CodInventario As Integer, CodSimples As Integer)
            Dim parameters As SqlParameter() = {
                    New SqlParameter("@CodSimples", CodSimples),
                    New SqlParameter("@CodInventario", CodInventario)
        }

            Dim Tabela As DataTable = ClasseConexao.Consultar("SELECT COUNT(Cod_Simples) FROM CS_InventarioDetalhes WHERE Cod_Inventario = @CodInventario AND Cod_Simples = @CodSimples", parameters)
            If Tabela.Rows.Count > 0 Then
                Dim sql As String = "DELETE FROM Tbl_Inventario_Item WHERE Cod_Inventario = @CodInventario AND Cod_Simples = @CodSimples"
                ClasseConexao.Operar(sql, parameters)
            Else
                Exit Sub
            End If
        End Sub
        Public Sub SalvaMotivoEstoque(Motivo As String, Operacao As String)
            Dim sql As String = "INSERT INTO Tbl_MotivoAlteracaoEstoque (Motivo,Operacao) VALUES (@Motivo, @Operacao)"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@Motivo", Motivo),
                     New SqlParameter("@Operacao", Operacao)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub AtualizaMotivoEstoque(Codigo As Integer, Motivo As String, Operacao As String)
            Dim sql As String = "UPDATE Tbl_MotivoAlteracaoEstoque SET Motivo = @Motivo, Operacao = @Operacao WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@Codigo", Codigo),
                     New SqlParameter("@Motivo", Motivo),
                     New SqlParameter("@Operacao", Operacao)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
        Public Sub ExcluiMotivoEstoque(Codigo As Integer)
            Dim sql As String = "DELETE FROM Tbl_MotivoAlteracaoEstoque WHERE Codigo = @Codigo"
            Dim parameters As SqlParameter() = {
                     New SqlParameter("@Codigo", Codigo)
    }
            ClasseConexao.Operar(sql, parameters)
            MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
#End Region
#Region "FUNCOES"
        Public Function PesquisaInventario(Status As String, CodInventario As Integer, Descricao As String, TipoInventario As String, DataInicio As String, DataFim As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM CS_Inventario WHERE 1=1 ")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Em Aberto"
                        sql.AppendLine("AND CodStatus = 0")
                    Case "Concluído"
                        sql.AppendLine("AND CodStatus = 1")
                    Case "Todos"
                        sql.AppendLine("AND CodStatus IS NOT NULL")
                End Select

                If CodInventario <> 0 Then
                    sql.AppendLine("AND Codigo = @CodInventario")
                    parameters.Add(New SqlParameter("@CodInventario", CodInventario))
                End If

                If Not String.IsNullOrEmpty(Descricao) Then
                    sql.AppendLine("AND Descricao LIKE @Descricao")
                    parameters.Add(New SqlParameter("@Descricao", "%" & Descricao & "%"))
                End If

                If Not String.IsNullOrEmpty(TipoInventario) Then
                    sql.AppendLine("AND TipoInventario = @TipoInventario")
                    parameters.Add(New SqlParameter("@TipoInventario", TipoInventario))
                End If

                'Pesquisa pela data
                If Not String.IsNullOrEmpty(DataInicio) And Not String.IsNullOrEmpty(DataFim) Then
                    sql.AppendLine("AND DataInventario BETWEEN @DataIni AND @DataFim")
                    parameters.Add(New SqlParameter("@DataIni", DataInicio))
                    parameters.Add(New SqlParameter("@DataFim", DataFim))
                ElseIf Not String.IsNullOrEmpty(DataInicio) And String.IsNullOrEmpty(DataFim) Then
                    sql.AppendLine("AND DataInventario >= @DataIni")
                    parameters.Add(New SqlParameter("@DataIni", DataInicio))
                ElseIf String.IsNullOrEmpty(DataInicio) And Not String.IsNullOrEmpty(DataFim) Then
                    sql.AppendLine("AND DataInventario <= @DataFim")
                    parameters.Add(New SqlParameter("@DataFim", DataFim))
                End If

                sql.AppendLine("ORDER BY Codigo DESC")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function
        Public Function GerarInventario(Positivo As String, Negativo As String, Zerado As String, Status As String, TipoItem As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM CS_Estoque WHERE 1=1 ")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Ativo"
                        sql.AppendLine("AND Descontinuado = 0")
                    Case "Inativo"
                        sql.AppendLine("AND Descontinuado = 1")
                    Case "Todos"
                        sql.AppendLine("AND Descontinuado IS NOT NULL")
                End Select

                If Not String.IsNullOrEmpty(Positivo) And String.IsNullOrEmpty(Negativo) And String.IsNullOrEmpty(Zerado) Then
                    sql.AppendLine("AND Estoque > 0")
                ElseIf Not String.IsNullOrEmpty(Positivo) And Not String.IsNullOrEmpty(Negativo) Then
                    sql.AppendLine("AND Estoque > 0 OR Estoque < 0")
                ElseIf Not String.IsNullOrEmpty(Positivo) And Not String.IsNullOrEmpty(Zerado) Then
                    sql.AppendLine("AND Estoque > 0 OR Estoque = 0")
                ElseIf Not String.IsNullOrEmpty(Negativo) And String.IsNullOrEmpty(Positivo) And String.IsNullOrEmpty(Zerado) Then
                    sql.AppendLine("AND Estoque < 0")
                ElseIf Not String.IsNullOrEmpty(Negativo) And Not String.IsNullOrEmpty(Zerado) Then
                    sql.AppendLine("AND Estoque < 0 OR Estoque = 0")
                ElseIf Not String.IsNullOrEmpty(Zerado) And String.IsNullOrEmpty(Positivo) And String.IsNullOrEmpty(Negativo) Then
                    sql.AppendLine("AND Estoque = 0")
                Else
                    sql.AppendLine("AND Estoque IS NOT NULL")
                End If

                If Not String.IsNullOrEmpty(TipoItem) Then
                    sql.AppendLine("AND Tipo_Prod = @TipoItem")
                    parameters.Add(New SqlParameter("@TipoItem", TipoItem))
                End If

                sql.AppendLine("ORDER BY Item")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function
        Public Function ConsultaInventario(sql As String, Optional CodInventario As Integer = 0)
            If CodInventario <> 0 Then
                Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodInventario", CodInventario)
                        }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
        Public Function ConsultaMotivoEstoque(sql As String, Optional CodMotivo As Integer? = 0)
            If CodMotivo <> 0 Then
                Dim parameters As SqlParameter() = {
                     New SqlParameter("@CodMotivo", CodMotivo)
                        }
                Return ClasseConexao.Consultar(sql, parameters)
            Else
                Return ClasseConexao.Consultar(sql, Nothing)
            End If
        End Function
        Public Function CalcularEstoqueMinMax()
            Return ClasseConexao.ExecProcedureRetorno("spCalculaEstoqueMaxMin", Nothing)
        End Function
        Public Function PesquisaEstoque(Positivo As String, Negativo As String, Zerado As String, Status As String, CodItem As Integer, Item As String, TipoItem As String, Departamento As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM CS_Estoque WHERE 1=1 ")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Ativo"
                        sql.AppendLine("AND Descontinuado = 0")
                    Case "Inativo"
                        sql.AppendLine("AND Descontinuado = 1")
                    Case "Todos"
                        sql.AppendLine("AND Descontinuado IS NOT NULL")
                End Select


                If Not String.IsNullOrEmpty(Positivo) And String.IsNullOrEmpty(Negativo) And String.IsNullOrEmpty(Zerado) Then
                    sql.AppendLine("AND Estoque > 0")
                ElseIf Not String.IsNullOrEmpty(Negativo) And String.IsNullOrEmpty(Positivo) And String.IsNullOrEmpty(Zerado) Then
                    sql.AppendLine("AND Estoque < 0")
                ElseIf Not String.IsNullOrEmpty(Zerado) And String.IsNullOrEmpty(Positivo) And String.IsNullOrEmpty(Negativo) Then
                    sql.AppendLine("AND Estoque = 0")
                ElseIf Not String.IsNullOrEmpty(Negativo) And Not String.IsNullOrEmpty(Zerado) Then
                    sql.AppendLine("AND Estoque < 0 OR Estoque = 0")
                ElseIf Not String.IsNullOrEmpty(Positivo) And Not String.IsNullOrEmpty(Negativo) Then
                    sql.AppendLine("AND Estoque > 0 OR Estoque < 0")
                ElseIf Not String.IsNullOrEmpty(Positivo) And Not String.IsNullOrEmpty(Zerado) Then
                    sql.AppendLine("AND Estoque > 0 OR Estoque = 0")
                Else
                    sql.AppendLine("AND Estoque IS NOT NULL")
                End If

                'Pesquisa pela código do item
                If CodItem <> 0 Then
                    sql.AppendLine("AND CodSimples = @CodSimples")
                    parameters.Add(New SqlParameter("@CodSimples", CodItem))
                End If

                'Pesquisa pela item
                If Not String.IsNullOrEmpty(Item) Then
                    sql.AppendLine("AND Item LIKE @Descricao")
                    parameters.Add(New SqlParameter("@Descricao", "%" & Item & "%"))
                End If

                'Pesquisa pela tipo do item
                If Not String.IsNullOrEmpty(TipoItem) Then
                    sql.AppendLine("AND Tipo_Prod LIKE @TipoItem")
                    parameters.Add(New SqlParameter("@TipoItem", "%" & TipoItem & "%"))
                End If

                'Pesquisa pelo departamento
                If Not String.IsNullOrEmpty(Departamento) Then
                    sql.AppendLine("AND Depto LIKE @Departamento")
                    parameters.Add(New SqlParameter("@Departamento", "%" & Departamento & "%"))
                End If

                sql.AppendLine("ORDER BY Item")
                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())
            Catch ex As SqlException
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function
        Public Function PesquisaHistEstoque(Status As String, CodItem As Integer, Item As String, TipoItem As String, Departamento As String, DATAINI As String, DATAFI As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_HistoricoEstoque WHERE 1=1")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Ativo"
                        sql.AppendLine("AND Descontinuado = 0")
                    Case "Inativo"
                        sql.AppendLine("AND Descontinuado = 1")
                    Case "Todos"
                        sql.AppendLine("AND Descontinuado IS NOT NULL")
                End Select


                'Pesquisa pela código do item
                If CodItem <> 0 Then
                    sql.AppendLine("AND CodSimples = @CodSimples")
                    parameters.Add(New SqlParameter("@CodSimples", CodItem))
                End If

                'Pesquisa pela item
                If Not String.IsNullOrEmpty(Item) Then
                    sql.AppendLine("AND Item LIKE @Descricao")
                    parameters.Add(New SqlParameter("@Descricao", "%" & Item & "%"))
                End If

                'Pesquisa pela tipo do item
                If Not String.IsNullOrEmpty(TipoItem) Then
                    sql.AppendLine("AND Tipo_Prod LIKE @TipoItem")
                    parameters.Add(New SqlParameter("@TipoItem", "%" & TipoItem & "%"))
                End If

                'Pesquisa pelo departamento
                If Not String.IsNullOrEmpty(Departamento) Then
                    sql.AppendLine("AND Depto LIKE @Departamento")
                    parameters.Add(New SqlParameter("@Departamento", "%" & Departamento & "%"))
                End If

                'Pesquisa pelo DATA
                If Not String.IsNullOrEmpty(DATAINI) And String.IsNullOrEmpty(DATAFI) Then
                    sql.AppendLine("AND DataMovimento > @DATAINI")
                    parameters.Add(New SqlParameter("@DATAINI", DATAINI))
                ElseIf String.IsNullOrEmpty(DATAINI) And Not String.IsNullOrEmpty(DATAFI) Then
                    sql.AppendLine("AND DataMovimento < @DATAFI")
                    parameters.Add(New SqlParameter("@DATAFI", DATAFI))
                ElseIf Not String.IsNullOrEmpty(DATAINI) And Not String.IsNullOrEmpty(DATAFI) Then
                    sql.AppendLine("AND DataMovimento BETWEEN @DATAINI AND @DATAFI")
                    parameters.Add(New SqlParameter("@DATAINI", DATAINI))
                    parameters.Add(New SqlParameter("@DATAFI", DATAFI))
                ElseIf Not String.IsNullOrEmpty(DATAINI) And Not String.IsNullOrEmpty(DATAFI) Then
                    sql.AppendLine("AND DataMovimento IS NOT NULL")
                End If

                sql.AppendLine("ORDER BY Item")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())
            Catch ex As SqlException
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function

#End Region
    End Class
End Namespace
