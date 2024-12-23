Imports System.Text
Imports Newtonsoft.Json
Imports System.Net
Imports System.Net.Http
Imports Newtonsoft.Json.Linq
Imports SDKBrasilAP
Imports Biblioteca.Classes.Conexao
Imports Microsoft.Data.SqlClient
Imports System.Data
Imports Xceed.Wpf.Toolkit
Namespace Bibliotecas.Classes.Integracao

    Public Class clsIntegracoes
        Inherits clsLocalidades

        Dim ClasseConexao As New ConexaoSQLServer
#Region "CONSTRUTORES"
        Public Sub New()
            httpClient = New HttpClient()
            ' Configurar httpClient, se necessário
        End Sub

        Public Sub New(_codcidade As Integer, _codestado As Integer, _cidade As String, _estado As Integer, _sigla As String, _capital As String, _regiao As String, _codibgeCidade As Integer, _codibgeestado As Integer)
            CodIBGECidade = _codibgeCidade
            CodCidade = _codcidade
            Localidade = _cidade
            CodIBGEEstado = _codibgeestado
            CodEstado = _codestado
            UF = _estado
            Sigla = _sigla
            Capital = _capital
        End Sub
#End Region
#Region "PROPRIEDADES"
        Private ReadOnly httpClient As HttpClient

        Public Property ID As Integer
        Public Property Descricao As String
        Public Property Notas As String
        Public Property Availability As String
        Public Property Condition As String
        Public Property Price As Decimal
        Public Property Link As String
        Public Property ImageLink As String
        Public Property Brand As String
        Public Property google_product_category As String
        Public Property fb_product_category As String
        Public Property quantity_to_sell_on_facebook As String
        Public Property sale_price As String
        Public Property sale_price_effective_date As String
        Public Property item_group_id As String
        Public Property Gender As String
        Public Property Color As String
        Public Property Size As String
        Public Property age_group As String
        Public Property material As String
        Public Property pattern As String
        Public Property shipping As String
        Public Property shipping_weight As String
        Public Property style As String
        Public Property CodNCM As String
        Public Property descricaoNCM As String
        Public Property datainicio As String
        Public Property datafim As String
        Public Property tipoato As String
        Public Property numeroato As String
        Public Property anoato As String
#End Region
#Region "METODOS"
        Public Function ExportacaoProdutos(Status As String, CodItem As Integer, Produto As String, Departamento As String, TipoItem As String) As DataTable
            Dim sql As New StringBuilder("SELECT * FROM Cs_ExportacaoProdutos WHERE 1=1")
            Dim parameters As New List(Of SqlParameter)()

            Try
                Select Case Status
                    Case "Inativo"
                        sql.AppendLine("AND Descontinuado = 1")
                    Case "Ativo"
                        sql.AppendLine("AND Descontinuado = 0")
                    Case "Todos"
                        sql.AppendLine("and Descontinuado IS NOT NULL")
                End Select

                'Pesquisa pela código do item
                If CodItem <> 0 Then
                    sql.AppendLine("AND Cod_Simples = @CodSimples")
                    parameters.Add(New SqlParameter("@CodSimples", CodItem))
                End If

                'Pesquisa pela item
                If Not String.IsNullOrEmpty(Produto) Then
                    sql.AppendLine("AND NomeProduto LIKE @Produto")
                    parameters.Add(New SqlParameter("@Produto", "%" & Produto & "%"))
                End If

                'Pesquisa pela tipo do item
                If Not String.IsNullOrEmpty(TipoItem) Then
                    sql.AppendLine("AND Tipo_Prod LIKE @TipoItem")
                    parameters.Add(New SqlParameter("@TipoItem", "%" & TipoItem & "%"))
                End If

                'Pesquisa pelo departamento
                If Not String.IsNullOrEmpty(Departamento) Then
                    sql.AppendLine("AND Departamento LIKE @Departamento")
                    parameters.Add(New SqlParameter("@Departamento", "%" & Departamento & "%"))
                End If

                sql.AppendLine("ORDER BY NomeProduto")

                Return ClasseConexao.Consultar(sql.ToString(), parameters.ToArray())

            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta da shopee!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function
#End Region
#Region "FUNCOES"
        ''' <summary>
        ''' Esta função realiza a consulta de um CEP na Brasil API e retorna as informações de endereço do CEP.
        ''' </summary>
        ''' <param name="CEP">Representa o CEP e deve possui 8 caracteres</param>
        ''' <returns>Retorna o Logradouro, Complemento,Bairro,Localidade,UF</returns>
        Public Async Function ObterEndereco(CEP As String) As Task
            Using brasilAPI As New BrasilAPI()
                Dim response = Await brasilAPI.CEP_V2(CEP)
                Endereco = response.Street.ToString
                Bairro = response.Neighborhood.ToString
                Localidade = response.City.ToString
                UF = response.UF.ToString
            End Using
        End Function
        ''' <summary>
        ''' Esta função realiza a consulta de um DDD na Brasil API e retorna os nomes das cidades
        ''' </summary>
        ''' <returns></returns>
        Public Async Function ListarDDD() As Task(Of List(Of String))
            Dim citiesList As New List(Of String)()

            Try
                ' Chamada assíncrona para buscar os dados da API
                Using brasilAPI = New BrasilAPI()
                    Dim response = Await brasilAPI.DDD(17)

                    ' Adicionar os nomes das cidades à lista
                    For Each city In response.Cities
                        citiesList.Add(city)
                    Next
                End Using
            Catch ex As Exception
                ' Pode lançar a exceção ou logar aqui, dependendo do seu requisito
                Throw New Exception("Erro ao obter os nomes das cidades: " & ex.Message)
            End Try

            ' Retorna a lista de nomes das cidades
            Return citiesList
        End Function
        ''' <summary>
        ''' Esta função realiza a consulta de um CNPJ na Brasil API e retorna as informações da empresa
        ''' </summary>
        ''' <param name="CNPJ">Representa o CNPJ.</param>
        ''' <returns>Retorna o nome da empresa, razão social, endereço, bairro, localidade e UF.</returns>
        Public Async Function ObterEmpresa(CNPJ As String) As Task
            Using brasilAPI As New BrasilAPI()
                Dim response = Await brasilAPI.CNPJ(CNPJ)
                NomeFantasia = response.NomeFantasia
                RazaoSocial = response.RazaoSocial
                Endereco = response.Logradouro
                Numero = response.Numero
                Complemento = response.Complemento
                Localidade = response.Municipio
                UF = response.UF
                Bairro = response.Bairro
                CEP = response.CEP
                DDDTelefone1 = response.DDD_Telefone1
            End Using
        End Function
        ''' <summary>
        ''' Esta função realiza a consulta de um código do banco na Brasil API e retorna as informações do banco.
        ''' </summary>
        ''' <param name="Codigo">Representa o Código do banco.</param>
        ''' <returns>Retorna o nome do banco,Código ISPB e Código do banco</returns>
        Public Function ObterBancos(Codigo As Integer) As clsIntegracoes
            Dim url = $"https://brasilapi.com.br/api/banks/v1/{Codigo}"

            Try
                Using client As New WebClient()
                    Dim json = client.DownloadString(url)

                    ' Desserializar o JSON manualmente para um objeto da classe clsEstado
                    Dim Bancos As New clsIntegracoes()
                    Dim jsonData As JObject = JObject.Parse(json)

                    Bancos.CodISPB = jsonData("ispb").ToObject(Of Integer)()
                    Bancos.Banco = jsonData("name").ToString()
                    Bancos.CodBanco = jsonData("code").ToString()
                    Bancos.NomeBanco = jsonData("fullName").ToString()

                    Return Bancos
                End Using
            Catch ex As Exception
                MessageBox.Show("Erro ao obter banco!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End Try
        End Function
        ''' <summary>
        ''' Esta função realiza a consulta de um Código do estado na Brasil API e retorna as informações do estado.
        ''' </summary>
        ''' <param name="Codigo">Representa o código IBGE do estado</param>
        ''' <returns>Retorna o nome do estado, Código do estado e Código da região</returns>
        Public Function ObterEstado(Codigo As Integer) As clsIntegracoes
            Dim url = $"https://brasilapi.com.br/api/ibge/uf/v1/{Codigo}"

            Try
                Using client As New WebClient()
                    Dim json = client.DownloadString(url)
                    Dim estado As New clsIntegracoes()
                    Dim jsonData As JObject = JObject.Parse(json)

                    estado.CodEstado = jsonData("id").ToObject(Of Integer)()
                    estado.Sigla = jsonData("sigla").ToString()
                    estado.UF = jsonData("nome").ToString()


                    Dim regiao As JObject = jsonData("regiao")
                    estado.CodRegiao = regiao("id").ToObject(Of Integer)()
                    estado.SiglaRegiao = regiao("sigla").ToString()
                    estado.NomeRegiao = regiao("nome").ToString()

                    Return estado
                End Using
            Catch ex As Exception
                MessageBox.Show("Erro ao obter estado!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End Try
        End Function
        ''' <summary>
        ''' Esta função realiza a consulta de um Código do município na Brasil API e retorna as informações do estado.
        ''' </summary>
        ''' <param name="SiglaUF">Representa a sigla do estado.</param>
        ''' <returns>Retorna o nome do município e o código IBGE.</returns>
        Public Function ObterMunicipios(SiglaUF As String) As List(Of clsIntegracoes)
            Dim url = $"https://brasilapi.com.br/api/ibge/municipios/v1/{SiglaUF}?providers=gov"
            Dim listaMunicipios As New List(Of clsIntegracoes)

            Try
                Using client As New WebClient()
                    Dim json = client.DownloadString(url)
                    Dim MunicipioList As JArray = JArray.Parse(json)

                    For Each municipio In MunicipioList
                        Dim MunicipioObj As New clsIntegracoes()
                        MunicipioObj.Localidade = municipio("nome").ToString()
                        MunicipioObj.CodIBGECidade = municipio("codigo_ibge").ToString()
                        listaMunicipios.Add(MunicipioObj)
                    Next
                End Using
            Catch ex As Exception
                MessageBox.Show("Erro ao obter municípios!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            Return listaMunicipios
            'Try
            '    Using client As New WebClient()
            '        Dim json = client.DownloadString(url)
            '        Dim Municipio As New clsIntegracoes()
            '        Dim jsonData As JObject = JObject.Parse(json)

            '        Municipio.Localidade = jsonData("nome").ToObject(Of Integer)()
            '        Municipio.CodIBGECidade = jsonData("codigo_ibge").ToString()

            '        Return Municipio
            '    End Using
            'Catch ex As Exception
            '    MessageBox.Show("Erro ao obter estado!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    Return Nothing
            '    Throw
            'End Try
        End Function
        ''' <summary>
        ''' Esta função realiza a consulta de um Código do NCM na Brasil API e retorna as informações do NCM.
        ''' </summary>
        ''' <param name="Codigo">Representa o código do NCM.</param>
        ''' <returns>Retorna o nome do NCM</returns>
        Public Function ObterNCM(Codigo As Integer) As clsIntegracoes
            Dim url = $"https://brasilapi.com.br/api/ncm/v1/{Codigo}"

            Try
                Using client As New WebClient()
                    Dim json = client.DownloadString(url)
                    Dim ncm As New clsIntegracoes()
                    Dim jsonData As JObject = JObject.Parse(json)

                    ncm.CodEstado = jsonData("codigo").ToObject(Of Integer)()
                    ncm.Sigla = jsonData("descricao").ToString()
                    ncm.UF = jsonData("data_inicio").ToString()
                    ncm.CodRegiao = Regiao("data_fim").ToString()
                    ncm.SiglaRegiao = Regiao("tipo_ato").ToString()
                    ncm.NomeRegiao = Regiao("numero_ato").ToString()

                    Return ncm
                End Using
            Catch ex As Exception
                MessageBox.Show("Erro ao obter estado!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End Try
        End Function
        ''' <summary>
        ''' Esta funcão realiza a consulta de todos os municípios na API do IBGE
        ''' </summary>
        ''' <returns>Retorna a lista de municípios.</returns>
        Public Function GetMunicipios() As List(Of clsIntegracoes)
            Dim url = "https://servicodados.ibge.gov.br/api/v1/localidades/municipios?orderBy=nome"
            Dim json = New Net.WebClient().DownloadString(url)

            Return JsonConvert.DeserializeObject(Of List(Of clsIntegracoes))(json)
        End Function
        ''' <summary>
        ''' Esta função realiza a consulta de todos os municípios por estado na API do IBGE
        ''' </summary>
        ''' <param name="UF">Representa a sigla do estado.</param>
        ''' <returns>Retorna a lista de municípios.</returns>
        Public Function GetMunicipiosPorEstados(UF As String) As List(Of clsIntegracoes)
            Dim url = $"https://servicodados.ibge.gov.br/api/v1/localidades/estados/{UF}/municipios"
            Dim json = New Net.WebClient().DownloadString(url)

            Return JsonConvert.DeserializeObject(Of List(Of clsIntegracoes))(json)
        End Function
        ''' <summary>
        ''' Esta função realiza a consulta de todos os estados na API do IBGE
        ''' </summary>
        ''' <returns>Retorna a lista de estados.</returns>
        Public Function GetEstados() As List(Of clsIntegracoes)
            Dim url = "https://servicodados.ibge.gov.br/api/v1/localidades/estados?orderBy=nome"
            Dim json = New Net.WebClient().DownloadString(url)

            Return JsonConvert.DeserializeObject(Of List(Of clsIntegracoes))(json)
        End Function

        ''' <summary>
        ''' Esta função realiza a consulta de um CEP na API Via CEP e retorna as informações de endereço do CEP.
        ''' </summary>
        ''' <param name="CEP">Representa o número do CEP e deve possui 8 caracteres sem máscara.</param>
        ''' <returns>Retorna o Logradouro, Complemento,Bairro,Localidade,UF,Código do IBGE, Gia, DDD e Siafi.</returns>
        Public Function GetCEP(CEP As String) As clsIntegracoes
            Dim url = $"https://viacep.com.br/ws/{CEP}/json/"
            Dim json = New Net.WebClient().DownloadString(url)

            ' Desserializar o JSON para um objeto da classe clsLocalidades
            Return JsonConvert.DeserializeObject(Of clsIntegracoes)(json)
        End Function
#End Region
    End Class
    Public Class ClsRegiao
        <JsonProperty("id")>
        Public Property Id As Integer

        <JsonProperty("sigla")>
        Public Property SiglaRegiao As String

        <JsonProperty("nome")>
        Public Property NomeRegiao As String

    End Class
    Public Class Mesorregiao
        <JsonProperty("id")>
        Public Property Id As Integer

        <JsonProperty("nome")>
        Public Property Nome As String

        <JsonProperty("UF")>
        Public Property Estado As clsLocalidades
    End Class
    Public Class Microrregiao
        <JsonProperty("id")>
        Public Property Id As Integer

        <JsonProperty("nome")>
        Public Property Nome As String

        <JsonProperty("mesorregiao")>
        Public Property Mesorregiao As Mesorregiao
    End Class
End Namespace
