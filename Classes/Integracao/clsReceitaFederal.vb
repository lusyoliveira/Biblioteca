Imports System.Net
Imports Newtonsoft.Json
Namespace Classes.Integracao

    Public Class AtividadePrincipal
        Property code As String
        Property text As String
    End Class
    Public Class AtividadeSecundaria
        Property code As String
        Property text As String
    End Class
    Public Class Billing
        Property free As Boolean
        Property database As Boolean
    End Class
    Public Class Extra

    End Class
    Public Class Qsa
        Property nome As String
        Property qual As String
    End Class
    Public Class Empresa
        Property abertura As String
        Property situacao As String
        Property tipo As String
        Property nome As String
        Property porte As String
        Property natureza_juridica As String
        Property qsa As List(Of Qsa)
        Property logradouro As String
        Property numero As String
        Property municipio As String
        Property bairro As String
        Property uf As String
        Property cep As String
        Property telefone As String
        Property data_situacao As String
        Property motivo_situacao As String
        Property cnpj As String
        Property ultima_situacao As DateTime
        Property status As String
        Property fantasia As String
        Property complemento As String
        Property email As String
        Property efr As String
        Property situcao_especial As String
        Property data_situacao_especial As List(Of AtividadePrincipal)
        Property atividade_principal As List(Of AtividadeSecundaria)
        Property atividade_secundaria As String
        Property capital_social As String
        Property extra As Extra
        Property billing As Billing

        Public Shared Function ObterDadosCNPJ(cnpj As String) As Empresa
            Dim url = "https://www.receitaws.com.br/v1/cnpj/" + cnpj
            Dim json = New WebClient().DownloadString(url)

            Dim Empresa = JsonConvert.DeserializeObject(Of Empresa)(json)

            Return Empresa
        End Function
    End Class
End Namespace
