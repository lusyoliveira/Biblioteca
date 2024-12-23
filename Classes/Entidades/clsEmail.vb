Imports System.Data.SqlClient
Imports System.Net.Mail
Imports Biblioteca.Classes.Conexao

''' <summary>
''' Esta classe representa todas as rotinas do sistema que envolve o envio de e-mail
''' </summary>
Public Class clsEmail
    Dim ClasseConexao As New ConexaoSQLServer
#Region "PROPRIEDADES"
    Public Property CodPedido As Integer
    Public Property ServidorSMTP As String
    Public Property CodEmail As Integer
    Public Property Email As String
    Public Property Senha As String
    Public Property Destinatario As String
    Public Property Assunto As String
    Public Property Mensagem As String
    Public Property CAMINHOARQUIVO As String
    Public Property TipoMensagem As String
    Public Property Remetente As String
    Public Property TipoContato As String

#End Region
#Region "CONSTRUTORES"
    Public Sub New()

    End Sub

    Public Sub New(_codigo As Integer, _servidorSMTP As String, _email As String, _senha As String, _destinatario As String, _assunto As String, _mensagem As String)
        CodEmail = _codigo
        ServidorSMTP = _servidorSMTP
        Email = _email
        Senha = _senha
        Destinatario = _destinatario
        Assunto = _assunto
        Mensagem = _mensagem
    End Sub
#End Region
#Region "METODOS"
    Public Sub EnviaEmail(CodPedido As Integer, TipoMensagem As String, Anexo As String)
        Dim Tabela As DataTable = ConsultaEmail(CodPedido, TipoMensagem, Anexo)
        If Tabela IsNot Nothing AndAlso Tabela.Rows.Count > 0 Then
            Dim ServidorSMTP As String = Tabela.Rows(0)("ServidorSMTP").ToString()
            Dim Senha As String = Tabela.Rows(0)("Senha").ToString()
            Dim Destinatario As String = Tabela.Rows(0)("Destinatario").ToString()
            Dim Assunto As String = Tabela.Rows(0)("Assunto").ToString()
            Dim Mensagem As String = Tabela.Rows(0)("Mensagem").ToString()
            Dim Remetente As String = Tabela.Rows(0)("Remetente").ToString()

            Try
                Using smtp As New SmtpClient
                    Using email As New MailMessage()
                        'Servidor SMTP
                        smtp.Host = ServidorSMTP
                        smtp.UseDefaultCredentials = False
                        smtp.Credentials = New Net.NetworkCredential(Remetente, Senha)
                        smtp.Port = 587
                        smtp.EnableSsl = True
                        'OAuth2/Modern Auth
                        'Define a configurações para o envio da mensagem
                        email.From = New MailAddress(Remetente)
                        email.To.Add(Destinatario)
                        email.Subject = Assunto
                        email.IsBodyHtml = True
                        email.Body = Mensagem

                        'Se houver anexo
                        If Anexo <> "" Or Anexo Is Nothing Then
                            email.Attachments.Add(New Attachment(Anexo))
                        End If
                        'Envio o email
                        smtp.Send(email)
                    End Using
                End Using
                MessageBox.Show("Email enviado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As ArgumentException
                MessageBox.Show("Não foi possível enviar a mensagem, pois o destinatário naõ foi informado!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            Catch ex As Exception
                MessageBox.Show("Não foi possível enviar a mensagem!" & vbCrLf & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw
            End Try

        End If

    End Sub
#End Region
#Region "FUNCOES"
    Public Function ConsultaEmail(CodPedido As Integer, TipoMensagem As String, Anexo As String)
        Dim parameters As SqlParameter() = {
                                    New SqlParameter("@CODPEDIDO", CodPedido),
                                    New SqlParameter("@TIPO", TipoMensagem),
                                    New SqlParameter("@ANEXO", Anexo),
                                    New SqlParameter("@MENSAGEM", "")
        }
        ClasseConexao.ExecProcedureRetorno("spEnviaEmail", parameters)
    End Function
#End Region
End Class
