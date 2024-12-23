Imports Newtonsoft.Json
Imports System.Net
Imports System.Text
Imports Xceed.Wpf.Toolkit
Namespace Classes.Integracao
    ''' <summary>
    ''' Esta classe respresenta todas as rotina pertinentes a integração com os correios para retorna informações de um determinado rastreio. Esta integração está descontinuada.
    ''' </summary>
    Public Class Endereco
        Public Property cidade As String
        Public Property uf As String
        Public Property bairro As String
        Public Property cep As String
        Public Property logradouro As String
        Public Property numero As String
    End Class

    Public Class Unidade
        Public Property codMcu As String
        Public Property codSro As String
        Public Property endereco As Endereco
        Public Property nome As String
        Public Property tipo As String
    End Class

    Public Class UnidadeDestino
        Public Property codMcu As String
        Public Property codSro As String
        Public Property endereco As Endereco
        Public Property nome As String
        Public Property tipo As String
    End Class

    Public Class Destinatario
        Public Property cep As String
    End Class

    Public Class Evento
        Public Property codigo As String
        Public Property descricao As String
        Public Property dtHrCriado As DateTime
        Public Property tipo As String
        Public Property unidade As Unidade
        Public Property urlIcone As String
        Public Property unidadeDestino As UnidadeDestino
        Public Property detalhe As String
        Public Property destinatario As Destinatario
    End Class

    Public Class TipoPostal
        Public Property categoria As String
        Public Property descricao As String
        Public Property sigla As String
    End Class

    Public Class Objeto
        Public Property codObjeto As String
        Public Property dtPrevista As DateTime
        Public Property eventos As List(Of Evento)
        Public Property modalidade As String
        Public Property tipoPostal As TipoPostal
        Public Property habilitaAutoDeclaracao As Boolean
        Public Property permiteEncargoImportacao As Boolean
        Public Property habilitaPercorridaCarteiro As Boolean
        Public Property bloqueioObjeto As Boolean
        Public Property possuiLocker As Boolean
        Public Property habilitaLocker As Boolean
        Public Property habilitaCrowdshipping As Boolean
    End Class

    Public Class Rastreio
        Public Property objetos As List(Of Objeto)
        Public Property quantidade As Integer
        Public Property resultado As String
        Public Property versao As String
    End Class

    Public Class Pacote
        ''' <summary>
        ''' Esta função consulta o código de rastreio no Correios e retorna as informações sobre o pacote.
        ''' </summary>
        ''' <param name="codigo">Representa o código de rastreio de um pacote do tipo string.</param>
        ''' <returns></returns>
        Public Function ObterPacote(codigo As String) As Rastreio
            Dim strEnd As String = String.Format("https://proxyapp.correios.com.br/v1/sro-rastro/{0}", codigo)
            Dim result As Rastreio

            Try
                Dim wc As New WebClient
                wc.Encoding = Encoding.UTF8
                Dim strJson As String = wc.DownloadString(strEnd)

                result = JsonConvert.DeserializeObject(Of Rastreio)(strJson)

                Return result
            Catch ex As Exception
                MessageBox.Show("Não foi possível realizar a consulta!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try
        End Function
    End Class
End Namespace
