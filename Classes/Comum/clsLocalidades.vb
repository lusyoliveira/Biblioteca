Imports Microsoft.Data.SqlClient
Imports System.Text
Imports Biblioteca.Classes.Conexao

Public Class clsLocalidades
    Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
    Public Property Microrregiao As New Microrregiao()
    Public Property Regiao As ClsRegiao()
    Public Property CodRegiao As String
    Public Property NomeRegiao As String
    Public Property SiglaRegiao As String
    Private pIspb As String
    Public Property CodISPB() As String
        Get
            Return pIspb
        End Get
        Set(value As String)
            pIspb = value
        End Set
    End Property
    Private pBanco As String
    Public Property Banco() As String
        Get
            Return pBanco
        End Get
        Set(value As String)
            pBanco = value
        End Set
    End Property
    Private pCodBanco As String
    Public Property CodBanco() As String
        Get
            Return pCodBanco
        End Get
        Set(value As String)
            pCodBanco = value
        End Set
    End Property
    Private pNomeBanco As String
    Public Property NomeBanco() As String
        Get
            Return pNomeBanco
        End Get
        Set(value As String)
            pNomeBanco = value
        End Set
    End Property
    Private Logradouro As String
    Public Property Endereco As String
        Get
            Return Logradouro
        End Get
        Set(value As String)
            Logradouro = value
        End Set
    End Property
    Private pBairro As String
    Public Property Bairro As String
        Get
            Return pBairro
        End Get
        Set(value As String)
            pBairro = value
        End Set
    End Property
    Private idCidade As Integer
    Public Property CodCidade As String
        Get
            Return idCidade
        End Get
        Set(value As String)
            idCidade = value
        End Set
    End Property
    Private Cidade As String
    Public Property Localidade As String
        Get
            Return Cidade
        End Get
        Set(value As String)
            Cidade = value
        End Set
    End Property
    Private Estado As String
    <JsonProperty("nome")>
    Public Property UF As String
        Get
            Return Estado
        End Get
        Set(value As String)
            Estado = value
        End Set
    End Property
    Private pSigla As String
    <JsonProperty("sigla")>
    Public Property Sigla As String
        Get
            Return pSigla
        End Get
        Set(value As String)
            pSigla = value
        End Set
    End Property
    Private pCapital As String
    Public Property Capital As String
        Get
            Return pCapital
        End Get
        Set(value As String)
            pCapital = value
        End Set
    End Property
    Private pCEP As String
    Public Property CEP As String
        Get
            Return pCEP
        End Get
        Set(value As String)
            pCEP = value
        End Set
    End Property
    Private pCodEstado As Integer
    <JsonProperty("id")>
    Public Property CodEstado As Integer
        Get
            Return pCodEstado
        End Get
        Set(value As Integer)
            pCodEstado = value
        End Set
    End Property
    Private IBGECidade As String
    Public Property CodIBGECidade As String
        Get
            Return IBGECidade
        End Get
        Set(value As String)
            IBGECidade = value
        End Set
    End Property
    Private IBGEEstado As String
    Public Property CodIBGEEstado As String
        Get
            Return IBGEEstado
        End Get
        Set(value As String)
            IBGEEstado = value
        End Set
    End Property
    Private Gia As String
    Public Property CodGia As String
        Get
            Return Gia
        End Get
        Set(value As String)
            Gia = value
        End Set
    End Property
    Private CodDDD As String
    Public Property DDD As String
        Get
            Return CodDDD
        End Get
        Set(value As String)
            CodDDD = value
        End Set
    End Property
    Private psiafi As String
    Public Property Siafi As String
        Get
            Return psiafi
        End Get
        Set(value As String)
            psiafi = value
        End Set
    End Property
    Private _CPFCNPJ As String
    Public Property CPFCNPJ As String
        Get
            Return _CPFCNPJ
        End Get
        Set(value As String)
            _CPFCNPJ = value
        End Set
    End Property
    Private _razao_social As String
    Public Property RazaoSocial As String
        Get
            Return _razao_social
        End Get
        Set(value As String)
            _razao_social = value
        End Set
    End Property
    Private _nome_fantasia As String
    Public Property NomeFantasia As String
        Get
            Return _nome_fantasia
        End Get
        Set(value As String)
            _nome_fantasia = value
        End Set
    End Property
    Private _situacao_cadastral As Integer
    Public Property SituacaoCadastral As Integer
        Get
            Return _situacao_cadastral
        End Get
        Set(value As Integer)
            _situacao_cadastral = value
        End Set
    End Property
    Private _descricao_situacao_cadastral As String
    Public Property DescricaoSituacaoCadastral As String
        Get
            Return _descricao_situacao_cadastral
        End Get
        Set(value As String)
            _descricao_situacao_cadastral = value
        End Set
    End Property
    Private _data_situacao_cadastral As Date
    Public Property DataSituacaoCadastral As Date
        Get
            Return _data_situacao_cadastral
        End Get
        Set(value As Date)
            _data_situacao_cadastral = value
        End Set
    End Property
    Private _motivo_situacao_cadastral As Integer
    Public Property MotivoSituacaoCadastral As Integer
        Get
            Return _motivo_situacao_cadastral
        End Get
        Set(value As Integer)
            _motivo_situacao_cadastral = value
        End Set
    End Property
    Private _nome_cidade_exterior As String
    Public Property NomeCidadeExterior As String
        Get
            Return _nome_cidade_exterior
        End Get
        Set(value As String)
            _nome_cidade_exterior = value
        End Set
    End Property
    Private _descricao_tipo_de_logradouro As String
    Public Property DescricaoTipoDeLogradouro As String
        Get
            Return _descricao_tipo_de_logradouro
        End Get
        Set(value As String)
            _descricao_tipo_de_logradouro = value
        End Set
    End Property

    Private _numero As String
    Public Property Numero As String
        Get
            Return _numero
        End Get
        Set(value As String)
            _numero = value
        End Set
    End Property
    Private _complemento As String
    Public Property Complemento As String
        Get
            Return _complemento
        End Get
        Set(value As String)
            _complemento = value
        End Set
    End Property
    Private _codigo_municipio As Integer
    Public Property CodigoMunicipio As Integer
        Get
            Return _codigo_municipio
        End Get
        Set(value As Integer)
            _codigo_municipio = value
        End Set
    End Property
    Private _ddd_telefone_1 As String
    Public Property DDDTelefone1 As String
        Get
            Return _ddd_telefone_1
        End Get
        Set(value As String)
            _ddd_telefone_1 = value
        End Set
    End Property
    Private _ddd_telefone_2 As String
    Public Property DDDTelefone2 As String
        Get
            Return _ddd_telefone_2
        End Get
        Set(value As String)
            _ddd_telefone_2 = value
        End Set
    End Property
#End Region
#Region "METODOS"
    Public Sub SalvarEstado(Sigla As String, Estado As String, Capital As String, Regiao As String)
        Dim sql As String = "INSERT INTO Tbl_Estados (Sigla,Estado,Capital,Regiao) VALUES (@SIGLA,@ESTADO,@CAPITAL,@REGIAO)"
        Dim parameters As SqlParameter() = {
                New SqlParameter("@SIGLA", Sigla),
                New SqlParameter("@ESTADO", Estado),
                New SqlParameter("@CAPITAL", Capital),
                New SqlParameter("@REGIAO", Regiao)
    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub AtualizaEstado(CodEstado As Integer, Sigla As String, Estado As String, Capital As String, Regiao As String)
        Dim sql As String = "UPDATE Tbl_Estados SET Sigla = @SIGLA, Estado = @ESTADO, Capital = @CAPITAL, Regiao = @REGIAO   WHERE Codigo = @CODIGO"
        Dim parameters As SqlParameter() = {
                New SqlParameter("@CODIGO", CodEstado),
                New SqlParameter("@SIGLA", Sigla),
                New SqlParameter("@ESTADO", Estado),
                New SqlParameter("@CAPITAL", Capital),
                New SqlParameter("@REGIAO", Regiao)
    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub ExcluirEstado(CodEstado As Integer)
        Dim sql As String = "DELETE FROM Tbl_Estados WHERE Codigo = @CODIGO"
        Dim parameters As SqlParameter() = {
                New SqlParameter("@CODIGO", CodEstado)
    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub AtualizaEstadoIBGE(CodEstado As Integer, CodEstadoIBGE As Integer, Sigla As String, Estado As String, Capital As String, Regiao As String)
        Dim sql As String = "UPDATE Tbl_Estados SET Sigla = @SIGLA,Estado = @ESTADO, Capital = @CAPITAL,Regiao = @REGIAO, CodigoIBGE = @CODIGOIBGE  WHERE Codigo = @CODIGO"
        Dim parameters As SqlParameter() = {
                New SqlParameter("@CODIGO", CodEstado),
                New SqlParameter("@SIGLA", Sigla),
                New SqlParameter("@ESTADO", Estado),
                New SqlParameter("@CAPITAL", Capital),
                New SqlParameter("@REGIAO", Regiao),
                New SqlParameter("@CODIGOIBGE", CodEstadoIBGE)
    }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub SalvarCidade(Cidade As String, Estado As Integer, CodigoIBGE As Integer)
        Dim sql As String = "INSERT INTO Tbl_Municipios (Municipio,Estado,CodigoIBGE) VALUES (@CIDADE, @ESTADO, @CodigoIBGE)"
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@CIDADE", Cidade),
                     New SqlParameter("@ESTADO", Estado),
                     New SqlParameter("@CodigoIBGE", CodigoIBGE)
                     }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro efetuado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub AtualizarCidade(CodCidade As Integer, Cidade As String, Estado As Integer)
        Dim sql As String = "UPDATE Tbl_Municipios SET Municipio = @MUNICIPIO, Estado = @ESTADO  WHERE Codigo = @CODIGO"
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@CODIGO", CodCidade),
                     New SqlParameter("@MUNICIPIO", Cidade),
                     New SqlParameter("@ESTADO", Estado)
        }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub ExcluirCidade(CodCidade As Integer)
        Dim sql As String = "DELETE FROM Tbl_Municipios WHERE Codigo = @CODIGO"
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@CODIGO", CodCidade)
        }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro excluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Public Sub AtualizarCidadeIBGE(CodCidade As Integer, Cidade As String, Estado As Integer, CodIBGECidade As Integer)
        Dim sql As String = "UPDATE Tbl_Municipios SET Municipio = @MUNICIPIO, Estado = @ESTADO, CodigoIBGE = @CODIGOIBGE  WHERE Codigo = @CODIGO"
        Dim parameters As SqlParameter() = {
                     New SqlParameter("@CODIGO", CodCidade),
                     New SqlParameter("@MUNICIPIO", Cidade),
                     New SqlParameter("@ESTADO", Estado),
                     New SqlParameter("@CODIGOIBGE", CodIBGECidade)
        }
        ClasseConexao.Operar(sql, parameters)
        MessageBox.Show("Cadastro atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
#End Region
#Region "FUNCOES"
    Public Function PesquisaEstado(CodEstado As Integer, Sigla As String, Estado As String) As DataTable
        Dim sql As New StringBuilder("SELECT * FROM Tbl_Estados WHERE 1=1 ")
        Dim parameters As New List(Of SqlParameter)()

        Try
            If CodEstado <> 0 Then
                sql.AppendLine("AND Codigo = @CodEstado")
                parameters.Add(New SqlParameter("@CodEstado", CodEstado))
            End If

            If Not String.IsNullOrEmpty(Estado) Then
                sql.AppendLine("AND Estado LIKE @Estado")
                parameters.Add(New SqlParameter("@Estado", "%" & Estado & "%"))
            End If

            If Not String.IsNullOrEmpty(Sigla) Then
                sql.AppendLine("AND Sigla LIKE @Sigla")
                parameters.Add(New SqlParameter("@Sigla", "%" & Sigla & "%"))
            End If

            sql.AppendLine("ORDER BY Estado")

            Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

        Catch ex As Exception
            MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Throw
        End Try
    End Function
    Public Function PesquisaCidade(CodCidade As String, Cidade As String, Estado As String) As DataTable
        Dim sql As New StringBuilder("SELECT * FROM Cs_Municipios WHERE 1=1 ")
        Dim parameters As New List(Of SqlParameter)()

        Try

            If CodCidade <> 0 Then
                sql.AppendLine("AND Codigo = @CodCidade")
                parameters.Add(New SqlParameter("@CodCidade", CodCidade))

            End If

            If Not String.IsNullOrEmpty(Estado) Then
                sql.AppendLine("AND Estado LIKE @Estado")
                parameters.Add(New SqlParameter("@Estado", "%" & Estado & "%"))

            End If

            If Not String.IsNullOrEmpty(Cidade) Then
                sql.AppendLine("AND Municipio LIKE @Municipio")
                parameters.Add(New SqlParameter("@Municipio", "%" & Cidade & "%"))

            End If
            sql.AppendLine("ORDER BY Municipio")

            Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

        Catch ex As Exception
            MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Throw
        End Try
    End Function
    Public Function ConsultaEstado(Sql As String, Optional CodEstado As Integer = 0)
        If CodEstado <> 0 Then
            Dim parameters As SqlParameter() = {
        New SqlParameter("@CodEstado", CodEstado)
            }

            Return ClasseConexao.Consultar(Sql, parameters)
        Else
            Return ClasseConexao.Consultar(Sql, Nothing)
        End If
    End Function
    Public Function ConsultaCidade(Sql As String)
        Return ClasseConexao.Consultar(Sql, Nothing)

    End Function

#End Region
End Class
