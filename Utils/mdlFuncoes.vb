
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop
Imports Microsoft.Online.SharePoint.TenantAdministration
Imports Microsoft.ReportingServices.Rendering.ExcelRenderer

Namespace Modulos.Utils.Funcoes

    Module mdlFuncoes
        ''' <summary>
        ''' Este modulo representa todas as funções disponíveis no sistema
        ''' </summary>
        Public REPORT As Microsoft.Reporting.WinForms.LocalReport
        Public CAMINHOARQUIVO As String
#Region "OBTER DADOS"

        Public Declare Auto Function GetPrivateProfileString Lib "Kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

        Public Declare Auto Function WritePrivateProfileString Lib "Kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer

        Declare Function GetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String,
    ByRef nSize As Integer) As Integer

        Declare Function GetComputerName Lib "kernel32" Alias _
    "GetComputerNameA" (ByVal lpBuffer As String,
    ByRef nSize As Integer) As Integer
        ''' <summary>
        ''' Esta função obtem o nome do usuário do windows autenticado no computador onde o sistema está sendo executado.
        ''' </summary>
        ''' <returns>Retorna o nome do usuário do windows.</returns>
        Public Function GetUserName() As String
            Dim iReturn As Integer
            Dim userName As String

            userName = New String(CChar(" "), 50)
            iReturn = GetUserName(userName, 50)
            GetUserName = userName.Substring(0, userName.IndexOf(Chr(0)))
        End Function
        ''' <summary>
        ''' Esta função obtem o nome do computador onde o sistema está sendo executado.
        ''' </summary>
        ''' <returns>Retorna o nome do computador.</returns>
        Public Function GetComputerName() As String
            Dim iReturn As Integer
            Dim computerName As String

            computerName = New String(CChar(" "), 30)
            iReturn = GetComputerName(computerName, 30)
            GetComputerName = computerName.Substring(0, computerName.IndexOf(Chr(0)))
        End Function
        ''' <summary>
        ''' Este função obtem o endereço de ipv4 do computador onde o sistema está sendo executado.
        ''' </summary>
        ''' <returns>Retorna uma lista de endereço IP.</returns>
        Public Function ObterIP()
            Dim EnderecoIp As String
            Dim hostname As IPHostEntry = Dns.GetHostEntry(My.Computer.Name.ToString)
            Dim ip As IPAddress() = hostname.AddressList
            EnderecoIp = ip.Last.ToString
            Return EnderecoIp
        End Function
#End Region
#Region "FORMATAÇÃO"
        ''' <summary>
        ''' Esta função recebe o CEP e formata com máscara padrão.
        ''' </summary>
        ''' <param name="txtTexto">Representa o número do CEP do tipo object.</param>
        ''' <returns>Retorna o CEP com a máscara aplicada.</returns>
        Public Function FormataCEP(ByVal txtTexto As Object)
            If Len(txtTexto.Text) = 2 Then
                txtTexto.Text = txtTexto.Text & "."
                txtTexto.SelectionStart = Len(txtTexto.Text) + 1
            ElseIf Len(txtTexto.Text) = 6 Then
                txtTexto.Text = txtTexto.Text & "-"
                txtTexto.SelectionStart = Len(txtTexto.Text) + 1
            End If

        End Function
        ''' <summary>
        ''' Este função recebe um número de telefone e aplica a máscara padrão.
        ''' </summary>
        ''' <param name="txtTexto">Representa o número do telefone do tipo object.</param>
        ''' <returns>Retorna o número de telefone com a máscara aplicada.</returns>
        Public Function FormataTelefone(ByVal txtTexto As Object)
            If Len(txtTexto.Text) = 0 Then
                txtTexto.Text = txtTexto.Text & "("
                txtTexto.SelectionStart = Len(txtTexto.Text) + 1
            ElseIf Len(txtTexto.Text) = 3 Then
                txtTexto.Text = txtTexto.Text & ")"
                txtTexto.SelectionStart = Len(txtTexto.Text) + 1
            ElseIf Len(txtTexto.Text) = 8 Then
                txtTexto.Text = txtTexto.Text & "-"
                txtTexto.SelectionStart = Len(txtTexto.Text) + 1
            End If
        End Function
        ''' <summary>
        ''' Este função recebe um número de CNPJ e aplica a máscara padrão.
        ''' </summary>
        ''' <param name="txtTexto">Representa o número do CNPJ do tipo object.</param>
        ''' <returns>Retorna o número de CNPJ com a máscara aplicada.</returns>
        Public Function FormataCNPJ(ByVal txtTexto As Object)
            If Len(txtTexto.Text) = 2 Or Len(txtTexto.Text) = 6 Then
                txtTexto.Text = txtTexto.Text & "."
                txtTexto.SelectionStart = Len(txtTexto.Text) + 1
            End If
            If Len(txtTexto.Text) = 10 Then
                txtTexto.Text = txtTexto.Text & "/"
                txtTexto.SelectionStart = Len(txtTexto.Text) + 1
            End If
            If Len(txtTexto.Text) = 15 Then
                txtTexto.Text = txtTexto.Text & "-"
                txtTexto.SelectionStart = Len(txtTexto.Text) + 1
            End If
        End Function
        ''' <summary>
        ''' Este função recebe um número de CPF e aplica a máscara padrão.
        ''' </summary>
        ''' <param name="txtTexto">Representa o número do CPF do tipo object.</param>
        ''' <returns>Retorna o número de CPF com a máscara aplicada.</returns>
        Public Function FormataCPF(ByVal txtTexto As Object)
            If Len(txtTexto.Text) = 3 Then
                txtTexto.Text = txtTexto.Text & "."
                txtTexto.SelectionStart = Len(txtTexto.Text) + 1
            ElseIf Len(txtTexto.Text) = 7 Then
                txtTexto.Text = txtTexto.Text & "."
                txtTexto.SelectionStart = Len(txtTexto.Text) + 1
            ElseIf Len(txtTexto.Text) = 11 Then
                txtTexto.Text = txtTexto.Text & "-"
                txtTexto.SelectionStart = Len(txtTexto.Text) + 1
            End If
        End Function
        ''' <summary>
        ''' Esta função formata uma mensagem de texto adicionando espaços em branco para alinhar o texto, preenchendo até um determinado número de caracteres.
        ''' </summary>
        ''' <param name="mensagem">Representa a mensagem de texto do tipo string.</param>
        ''' <param name="espacos">Representa a quantidade de espaços em branco do tipo integer.</param>
        ''' <param name="alinha">Representa o alinhamento do mensagem do tipo string.</param>
        ''' <param name="formato">Representa o formato sendo opcional do tipo string.</param>
        ''' <returns></returns>
        Public Function completa(ByVal mensagem As String, ByVal espacos As Integer, Optional ByVal alinha As String = "e", Optional ByVal formato As String = "")
            mensagem = Trim(mensagem)
            If formato <> "" Then

                If Format(formato, "<") = "n" Then
                    mensagem = Format(mensagem, "###,###,##0.00")
                Else
                    If Format(formato, "<") = "data" Then
                        mensagem = Format(mensagem, "dd/mm/yyyy")
                        '            If Format(formato, "<") = "currency" Then
                        '                mensagem = "R$" & CStr(Format(mensagem, "###,###,##0.00"))
                    Else
                        mensagem = Format(mensagem.Trim, formato)
                        '            End If
                    End If
                End If
            End If

            If Len(Trim(mensagem)) > espacos Then
                'espacos = Len(RTrim(mensagem))
                Return Left(mensagem.Trim, espacos)
                Exit Function
            End If

            If Format(alinha, "<") = "d" Then
                Return Space(espacos - Len((mensagem.Trim))) + mensagem.Trim
            Else
                Return mensagem.Trim + Space(espacos - Len((mensagem.Trim)))
            End If

        End Function
#End Region
#Region "VALIDAÇÃO"
        ''' <summary>
        ''' Esta função receber um número de telefone e verifica se o número é válido.
        ''' </summary>
        ''' <param name="txtTexto">Representa o número de telefone do tipo object.</param>
        ''' <returns>Retorna verdadeiro ou falso para o número de telefone.</returns>
        Public Function ValidaTelefones(ByVal txtTexto As Object) As Boolean
            Dim rgxCelular = New Regex("^\(?[1-9]{2}\)??(?:|9[1-9])[0-9]{3}\-?[0-9]{4}$")
            Dim rgxFixo = New Regex("^\(?[1-9]{2}\)??(?:[2-8])[0-9]{3}\-?[0-9]{4}$")

            If Not rgxCelular.IsMatch(txtTexto.text) Then
                MessageBox.Show("Celular Inválido")
                txtTexto.focus()
                Return False
            ElseIf Not rgxFixo.IsMatch(txtTexto.text) Then
                MessageBox.Show("Telefone Fixo Inválido")
                txtTexto.focus()
                Return False
            End If
            Return True
        End Function
        ''' <summary>
        ''' Esta função recebe um número do CNPJ e verifica se este é válido.
        ''' </summary>
        ''' <param name="CNPJ">Representa o número do CNPJ do tipo string.</param>
        ''' <returns>Retorna verdadeiro ou falso para o número de CNPJ.</returns>
        Public Function ValidarCNPJ(ByVal CNPJ As String) As Boolean
            Dim i As Integer
            Dim valida As Boolean

            CNPJ = CNPJ.Trim

            For i = 0 To CNPJ.Length - 1
                If CNPJ.Length <> 14 Or CNPJ(i).Equals(CNPJ) Then
                    Return False
                End If
            Next

            'remove a maskara
            'CPFCNPJ = CPFCNPJ.Substring(0, 2) + CPFCNPJ.Substring(3, 3) + CPFCNPJ.Substring(7, 3) + CPFCNPJ.Substring(11, 4) + CPFCNPJ.Substring(16)
            valida = efetivaValidacao(CNPJ)

            If valida Then
                ValidarCNPJ = True
            Else
                ValidarCNPJ = False
            End If
        End Function
        ''' <summary>
        ''' Esta função efetiva a validação do CNPJ.
        ''' </summary>
        ''' <param name="CNPJ">Representa o número do CNPJ do tipo string.</param>
        ''' <returns>Retorna verdadeiro ou falso para o número de CNPJ.</returns>
        Public Function efetivaValidacao(ByVal cnpj As String)
            Dim Numero(13) As Integer
            Dim soma As Integer
            Dim i As Integer
            'Dim valida As Boolean
            Dim resultado1 As Integer
            Dim resultado2 As Integer
            For i = 0 To Numero.Length - 1
                Numero(i) = CInt(cnpj.Substring(i, 1))
            Next
            soma = Numero(0) * 5 + Numero(1) * 4 + Numero(2) * 3 + Numero(3) * 2 + Numero(4) * 9 + Numero(5) * 8 + Numero(6) * 7 +
                       Numero(7) * 6 + Numero(8) * 5 + Numero(9) * 4 + Numero(10) * 3 + Numero(11) * 2
            soma = soma - (11 * (Int(soma / 11)))
            If soma = 0 Or soma = 1 Then
                resultado1 = 0
            Else
                resultado1 = 11 - soma
            End If

            If resultado1 = Numero(12) Then
                soma = Numero(0) * 6 + Numero(1) * 5 + Numero(2) * 4 + Numero(3) * 3 + Numero(4) * 2 + Numero(5) * 9 + Numero(6) * 8 +
                           Numero(7) * 7 + Numero(8) * 6 + Numero(9) * 5 + Numero(10) * 4 + Numero(11) * 3 + Numero(12) * 2
                soma = soma - (11 * (Int(soma / 11)))
                If soma = 0 Or soma = 1 Then
                    resultado2 = 0
                Else
                    resultado2 = 11 - soma
                End If
                If resultado2 = Numero(13) Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Function
        ''' <summary>
        ''' Esta função recebe um número de CPF e verifica se o mesmo é válido.
        ''' </summary>
        ''' <param name="CPF">Representa o número do CPF do tipo string.</param>
        ''' <returns>Retorna verdadeiro ou falso para o número de CPF.</returns>
        Public Function ValidarCPF(ByVal CPF As String) As Boolean

            Dim CpfValido = True
            Dim I As Integer, J As Byte, N1 As Integer, N2 As Integer
            'remove a maskara
            'CPF = CPF.Substring(0, 3) + CPF.Substring(4, 3) + CPF.Substring(8, 3) + CPF.Substring(12)
            If Len(CPF) <> 11 Then ' CPF tem que te 11 caracteres 
                CpfValido = False
            ElseIf CPF = StrDup(11, "0") Or CPF = StrDup(11, "1") Or CPF = StrDup(11, "2") Or CPF = StrDup(11, "3") Or CPF = StrDup(11, "4") Or CPF = StrDup(11, "5") Or CPF = StrDup(11, "6") Or CPF = StrDup(11, "7") Or CPF = StrDup(11, "8") Or CPF = StrDup(11, "9") Then
                CpfValido = False
            Else
                For I = 1 To Len(CPF)
                    If Not IsNumeric(Mid(CPF, I, 1)) Then 'todos os caractertes tem que ser números
                        CpfValido = False
                        Exit For
                    End If
                Next
            End If

            If CpfValido Then
                'Validar o primeiro número do dígito
                N1 = 0
                J = 1
                For I = 10 To 2 Step -1
                    N1 += Val(Mid(CPF, J, 1)) * I
                    J += 1
                Next
                N1 = (N1 * 10) Mod 11 'o resto da divisão
                If N1 = 10 Then
                    N1 = 0
                End If
                'Verificar o primeiro número do dígito bateu
                If N1 <> Mid(CPF, 10, 1) Then  'posição 10 (penúltima)
                    CpfValido = False
                End If

                'Validar o segundo número do digito
                If CpfValido Then
                    N2 = 0
                    J = 1
                    For I = 11 To 2 Step -1
                        N2 += Val(Mid(CPF, J, 1)) * I
                        J += 1
                    Next
                    N2 = (N2 * 10) Mod 11 'o resto da divisão
                    If N2 = 10 Then
                        N2 = 0
                    End If
                    'Verificar o segundo número do dígito bateu
                    If N2 <> Mid(CPF, 11, 1) Then 'posição 11 (última)
                        CpfValido = False
                    End If
                End If
                'MessageBox.Show(N1 & N2)  ' tire o comentário desta linha para ver o digito cálculado
            End If

            ValidarCPF = CpfValido
        End Function
        ''' <summary>
        ''' Esta função valida o status e define uma cor para o status.
        ''' </summary>
        ''' <param name="txtTexto">Representa o componente que contém o status do tipo object.</param>
        ''' <returns>Retorna o status com uma cor de texto definida.</returns>
        Public Function ValidaStatus(ByVal txtTexto As Object)
            If txtTexto.Text = "EM ABERTO" Then
                txtTexto.ForeColor = Color.Red
            ElseIf txtTexto.Text = "PRODUZIDO" Then
                txtTexto.ForeColor = Color.Yellow
            Else
                txtTexto.ForeColor = Color.Green
            End If
        End Function
        ''' <summary>
        ''' Esta função permite apenas a digitação de valor númerico do componente.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Public Sub ValidaNumero(sender As Object, e As KeyPressEventArgs)
            If Not IsNumeric(e.KeyChar) And Asc(e.KeyChar) <> 8 Then
                e.Handled = True
            End If
        End Sub
        ''' <summary>
        ''' EStá função recebe um e-mail e verifica se o mesmo é valido.
        ''' </summary>
        ''' <param name="Email">Represento o e-mail do tipo string.</param>
        ''' <returns>Retorna verdadeiro ou falso para o e-mail.</returns>
        Public Function ValidarEmail(Email As String) As Boolean

            Dim EmailValido = True
            Dim I As Byte
            Dim QtdeCaracteres As Byte

            If Len(Email) < 5 Then 'Não pode ter menos que 5 caracteres
                EmailValido = False
            ElseIf InStr(Email, "@") = 1 Or InStr(Email, ".") = 1 Then 'Não pode começar com @ ou .
                EmailValido = False
            ElseIf InStr(Email, "@") = Len(Email) Or InStr(Email, ".") = Len(Email) Then 'não pode terminar com @ ou com  .
                EmailValido = False
            ElseIf InStr(Email, ".") = 0 Then 'tem que ter pelo menos um .
                EmailValido = False
            ElseIf InStr(Email, "..") > 0 Then 'não pode ter dois pontos (ou mais) seguidos
                EmailValido = False
            Else
                'Verificando a quantidade de @ no email
                QtdeCaracteres = 0
                For I = 1 To Len(Email)
                    If Mid(Email, I, 1) = "@" Then
                        QtdeCaracteres = QtdeCaracteres + 1 'Quantidade de @
                    End If
                Next
                If QtdeCaracteres <> 1 Then 'Só pode ter um @
                    EmailValido = False
                End If
            End If

            If EmailValido = True Then
                'Verificando se tem mais de dois pontos depois do @
                QtdeCaracteres = 0
                For I = InStr(Email, "@") To Len(Email)
                    If Mid(Email, I, 1) = "." Then
                        QtdeCaracteres = QtdeCaracteres + 1 'Quantidade de . depois do @
                    End If
                Next
                If QtdeCaracteres > 2 Then 'Só pode ter até dois . depois do @
                    EmailValido = False
                End If
            End If

            If EmailValido = True Then
                'Verificar se tem somente caracteres válidos
                Dim Letra As String
                For I = 1 To Len(Email)
                    Letra = Mid(Email, I, 1)
                    If Not (LCase(Letra) Like "[a-z]" Or Letra = "@" Or Letra = "." Or Letra = "-" Or Letra = "_" Or IsNumeric(Letra)) Then
                        'Tem caracter inválido
                        EmailValido = False
                        Exit For
                    End If
                Next
            End If

            ValidarEmail = EmailValido
        End Function
        ''' <summary>
        ''' Esta função verifica se há uma conexão de rede ativa nom computador no qual o sistema está sendo executado. 
        ''' </summary>
        ''' <returns>Retorna verdadeiro ou falso para a conexão.</returns>
        Public Function Conectado() As Boolean
            Try
                Return My.Computer.Network.Ping("www.globo.com")
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' Está função valida um serial, provavelmente associado a algum tipo de licença ou autorização de software. O código tenta baixar um arquivo de texto de um servidor remoto usando My.Computer.Network.DownloadFile. O nome Do arquivo é gerado a partir de um valor retornado pela função serialhd (que provavelmente retorna um identificador único Do computador, possivelmente o serial do HD) e é salvo no diretório "c:\temp".
        ''' Se o download For bem-sucedido, a função retorna True; caso contrário, retorna False.
        ''' Se a operação For diferente de "VALIDAR", o código gera um arquivo de texto no diretório "c:\temp" com base no valor retornado pela função serialhd.Esse arquivo é então enviado para um servidor remoto usando My.Computer.Network.UploadFile.
        ''' </summary>
        ''' <param name="operacao">Representa a operação a ser executada do tipo string.</param>
        ''' <returns></returns>
        Public Function ValidarSerial(Optional ByVal operacao As String = "VALIDAR") As Boolean
            Dim arquivo As String
            If operacao.ToUpper = "VALIDAR" Then
                Try

                    arquivo = "c:\temp\" & Serialhd("c:\") & ".txt"
                    If My.Computer.FileSystem.DirectoryExists("c:\temp") = False Then
                        My.Computer.FileSystem.CreateDirectory("c:\temp")
                    End If
                    If My.Computer.FileSystem.FileExists(arquivo) = True Then
                        My.Computer.FileSystem.DeleteFile(arquivo)
                    End If
                    My.Computer.Network.DownloadFile("http://br.geocities.com/caafs2/" & Serialhd("c:\") & ".txt", arquivo)
                    ValidarSerial = True
                Catch ex As Exception
                    ValidarSerial = False
                End Try
                Exit Function
            Else
                arquivo = "c:\temp\" & Serialhd("c:\") & ".txt"
                If My.Computer.FileSystem.DirectoryExists("c:\temp") = False Then
                    My.Computer.FileSystem.CreateDirectory("c:\temp")
                End If
                If My.Computer.FileSystem.FileExists(arquivo) = True Then
                    My.Computer.FileSystem.DeleteFile(arquivo)
                End If
                Dim arq As StreamWriter = New StreamWriter(arquivo)
                arq.WriteLine(arquivo)
                arq.Close()
                'My.Computer.Network.UploadFile(arquivo, "http://br.geocities.com/caafs2", "caafs2", "a1b1c1d1")
                My.Computer.Network.UploadFile(arquivo, "https://cp1.runhosting.com/ftp_manager.html", "204774", "208bfcae")
            End If
        End Function
#End Region
#Region "ENCRIPTAÇÃO"
        Public Function mcripto(ByVal wvTEXTO As String)
            'faz encriptação de dados informandos em um textbox
            Dim wvTEXTO1, wvTEXTO2, wvRETORNA As String
            Dim X, Y, INDICE As Integer
            Dim CARACTER As String
            wvTEXTO = UCase(wvTEXTO)
            wvRETORNA = ""
            wvTEXTO1 = "ABCDEFGHIJKLMNOPQRSTUVXYZ1234567890 WÇÃÕ"
            wvTEXTO2 = "!@#$%^&*()_+|=\-][{}?/><,.~`®¬½¼¡«»¨©ÇÕÃ"
            For X = 1 To Len(wvTEXTO)
                CARACTER = Right(Left(wvTEXTO, X), 1)
                For Y = 1 To Len(wvTEXTO1)
                    If CARACTER = Right(Left(wvTEXTO1, Y), 1) Then
                        INDICE = Y
                    End If
                Next
                wvRETORNA = wvRETORNA + Right(Left(wvTEXTO2, INDICE), 1)

            Next
            mcripto = wvRETORNA
        End Function
        Public Function geraHash(ByVal valor As String) As String

            'Cria um objeto encoding para assegurar o encoding padrão para o texto fonte
            Dim Ue As New UnicodeEncoding()

            'Retorna um array de bytes baseado no texto fonte
            Dim ByteSourceText() As Byte = Ue.GetBytes(valor)

            'Instancia um objeto MD5
            Dim Md5 As New MD5CryptoServiceProvider()

            'Calcula o valor do hash para o texto
            Dim ByteHash() As Byte = Md5.ComputeHash(ByteSourceText)

            'Converte o valor para o formato string e retorna
            Return Convert.ToBase64String(ByteHash)
        End Function
        Public Function GerarSenha() As String
            Dim rand = New Random()
            Dim caracteres = "!@#$%*+-*0123456789ABCDEFGHIJKLMNOPQRSTUVXYWZabcdefghijklmnopqrstuvxywz"
            Dim numeroMaximo = caracteres.Length
            Dim numero As Integer
            Dim senha = ""

            For i = 0 To 6
                'Gerando um número entre 0 e menor que 8
                numero = rand.Next(numeroMaximo)
                senha = senha & caracteres(numero)
            Next

            Return senha
        End Function
#End Region
#Region "EXPORTAÇÃO DE DADOS"

        Public Function SalvarExcel(ByVal grid As DataGridView)
            Dim XcelApp As New Excel.Application()
            Dim colunasMoeda As New List(Of String) From {"Preço", "sale_price", "Preço de Varejo", "Custo de Compra", "regular_price", "sale_price_effective_date"}


            If grid.Rows.Count > 0 Then
                Try
                    XcelApp.Application.Workbooks.Add(Type.Missing)

                    ' Adiciona os cabeçalhos das colunas
                    For i As Integer = 0 To grid.Columns.Count - 1
                        XcelApp.Cells(1, i + 1) = grid.Columns(i).HeaderText
                    Next

                    ' Adiciona os dados do DataGridView
                    For i As Integer = 0 To grid.Rows.Count - 1
                        For j As Integer = 0 To grid.Columns.Count - 1
                            Dim valorCelula As Object = grid.Rows(i).Cells(j).Value
                            Dim nomeColuna As String = grid.Columns(j).HeaderText

                            If valorCelula IsNot Nothing AndAlso Not DBNull.Value.Equals(valorCelula) Then
                                If colunasMoeda.Contains(nomeColuna) AndAlso IsNumeric(valorCelula) Then
                                    ' Atribuir o valor como Decimal
                                    Dim valorDecimal As Decimal = Convert.ToDecimal(valorCelula)
                                    XcelApp.Cells(i + 2, j + 1) = valorDecimal
                                    XcelApp.Cells(i + 2, j + 1).NumberFormat = "#,##0.00" ' Formato de moeda
                                Else
                                    ' Atribuir valor como está
                                    XcelApp.Cells(i + 2, j + 1) = valorCelula.ToString()
                                End If
                            Else
                                XcelApp.Cells(i + 2, j + 1) = ""
                            End If
                        Next
                    Next

                    ' Ajusta a largura das colunas
                    XcelApp.Columns.AutoFit()

                    ' Torna o Excel visível
                    XcelApp.Visible = True
                Catch ex As Exception
                    MessageBox.Show("Não foi possível gerar a planilha:  " + ex.Message)
                    XcelApp.Quit()
                    XcelApp = Nothing
                End Try
            End If
        End Function
        Public Function ExportarRelatorio(nomeRelatorio As String, dataSourceName As String, data As Object, defaultFileName As String)
            Try
                Using anexo As New SaveFileDialog()
                    Dim REPORT As New Microsoft.Reporting.WinForms.LocalReport()
                    REPORT.ReportEmbeddedResource = nomeRelatorio

                    ' Adicionando o dataSource
                    Dim dataSource As New Microsoft.Reporting.WinForms.ReportDataSource(dataSourceName, data)
                    REPORT.DataSources.Add(dataSource)

                    REPORT.Refresh()

                    anexo.Title = "Salvar Relatório"
                    anexo.DefaultExt = "pdf"
                    anexo.FileName = defaultFileName
                    anexo.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*"

                    If anexo.ShowDialog() = DialogResult.OK Then
                        Dim caminhoArquivo As String = anexo.FileName

                        ' Renderizando o relatório para PDF
                        Dim warnings As Microsoft.Reporting.WinForms.Warning() = Nothing
                        Dim streamids As String() = Nothing
                        Dim mimeType As String = String.Empty
                        Dim encoding As String = String.Empty
                        Dim extension As String = String.Empty
                        Dim bytes As Byte()

                        ' Renderizando o relatório
                        bytes = REPORT.Render("PDF", Nothing, mimeType, encoding, extension, streamids, warnings)

                        ' Escrevendo os bytes no arquivo especificado
                        Using fs As New IO.FileStream(caminhoArquivo, IO.FileMode.Create)
                            fs.Write(bytes, 0, bytes.Length)
                        End Using

                        MessageBox.Show("Relatório gerado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        Exit Function
                    End If
                End Using
            Catch ex As Exception
                MessageBox.Show("Erro: " & ex.Message)
                Throw
            End Try
        End Function
        ''' <summary>
        ''' Esta função carrega a barra de progresso no formulário principal ao emitir um relatório.
        ''' </summary>
        ''' <param name="Barra">Representa a barra de progresso.</param>
        ''' <param name="Mensagem">Representa a mensagem que sera exibida.</param>
        Public Function CarregaBarraProgresso(ByVal Barra As ProgressBar, ByVal Mensagem As Object)
            Dim limite As Integer = 10000
            Dim cntr As Integer = 0
            Dim i As Integer
            Dim status As String

            Barra.Value = 0
            Barra.Step = 1
            For i = 0 To limite
                cntr = cntr + 1
                If cntr Mod 100 = 0 Then
                    status = cntr.ToString()
                    Barra.PerformStep()
                    Application.DoEvents()
                    System.Threading.Thread.Sleep(40)
                    Mensagem.Visible = True
                    Mensagem.Text = "Aguarde.."
                    Barra.Visible = True
                End If
            Next
            Barra.Visible = False
            Mensagem.Visible = False
        End Function
        ''' <summary>
        ''' Este método carrega a barra de progresso ao emitir um relatório.
        ''' </summary>
        ''' <param name="Barra">Representa a barra de progresso.</param>
        ''' <param name="Mensagem">Representa a mensagem que sera exibida.</param>
        Public Sub CarregandoRelatorio(ByVal Barra As ToolStripProgressBar, ByVal Mensagem As Object)
            Dim limite As Integer = 10000
            Dim cntr As Integer = 0
            Dim status As String

            Barra.Value = 0
            Barra.Step = 1
            For i = 0 To limite
                cntr = cntr + 1
                If cntr Mod 100 = 0 Then
                    status = cntr.ToString()
                    Barra.PerformStep()
                    Application.DoEvents()
                    System.Threading.Thread.Sleep(40)
                    Barra.Visible = True
                    Barra.Text = "Aguarde, gerando Relatório..."
                    Barra.Visible = True
                End If
            Next
            Barra.Visible = False
        End Sub
        ''' <summary>
        ''' Este método exporta os dados de um DataGrid para um arquivo CSV.
        ''' </summary>
        ''' <param name="dgv">Representa o objeto DataGrid</param>
        ''' <param name="defaultFileName">Representa o nome com o qual o arquivo será criado.</param>
        Public Sub ExportacaoCSV(ByVal dgv As DataGridView, defaultFileName As String)
            Try
                Using Salvar As New SaveFileDialog()
                    Salvar.Title = "Exportar Arquivo CSV"
                    Salvar.DefaultExt = "csv"
                    Salvar.FileName = defaultFileName
                    Salvar.Filter = "Arquivo CSV|*.csv"

                    If Salvar.ShowDialog() = DialogResult.OK Then
                        Dim caminhoArquivo As String = Salvar.FileName
                        ' Cria ou sobrescreve o arquivo no caminho especificado
                        Using escritor As New StreamWriter(caminhoArquivo)
                            ' Escreve os nomes das colunas
                            For i As Integer = 0 To dgv.Columns.Count - 1
                                escritor.Write(dgv.Columns(i).HeaderText)
                                If i < dgv.Columns.Count - 1 Then
                                    escritor.Write(";") ' Usando vírgula como delimitador
                                End If
                            Next
                            escritor.WriteLine()

                            ' Escreve os dados das linhas
                            For Each linha As DataGridViewRow In dgv.Rows
                                If Not linha.IsNewRow Then
                                    For i As Integer = 0 To dgv.Columns.Count - 1
                                        escritor.Write(linha.Cells(i).Value.ToString())
                                        If i < dgv.Columns.Count - 1 Then
                                            escritor.Write(";")
                                        End If
                                    Next
                                    escritor.WriteLine()
                                End If
                            Next
                        End Using
                    Else
                        Exit Sub
                    End If
                End Using
                MessageBox.Show("Dados exportados com sucesso!", "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Erro ao exportar dados: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
        ''' <summary>
        ''' Esta função realiza a importação de dados de um arquivo CSV.
        ''' </summary>
        ''' <param name="CaminhoArquivo">Representa o arquivo que será importado.</param>
        ''' <returns>Returna uma lista de dados</returns>
        Public Function ImportarCSV(CaminhoArquivo As String) As List(Of String)
            Dim Lista As New List(Of String)

            Using fs As New FileStream(CaminhoArquivo, FileMode.Open)
                Using Leitura As New StreamReader(fs)
                    While Not (Leitura.EndOfStream)
                        Dim Linha As String = Leitura.ReadLine()
                        Dim VetorLinha() As String = Linha.Split(";")

                        For i As Integer = 0 To VetorLinha.Length - 1
                            Lista.Add(VetorLinha(i))
                        Next
                    End While
                End Using
            End Using
            Return Lista
        End Function
#End Region
#Region "MANIPULAÇÃO"
        ''' <summary>
        ''' Esta função lê o conteúdo de um arquivo de texto e adiciona suas linhas a um controle ListBox.
        ''' </summary>
        ''' <param name="cxTexto">Representa o objeto ListBox</param>
        ''' <param name="strFile">Representa o arquivo</param>
        ''' <returns></returns>
        Public Function ProcessFile(ByVal cxTexto As ListBox, ByVal strFile As String)
            cxTexto.Items.AddRange(File.ReadAllLines(strFile))

            cxTexto.Refresh()
            ProcessFile = True
        End Function
        ''' <summary>
        ''' Esta função lê um valor específico de um arquivo INI, procurando na seção e chave especificadas. Se a chave não for encontrada, retorna um valor padrão fornecido.
        ''' </summary>
        ''' <param name="file_name">Representa o nome do arquivo INI do qual você deseja ler do tipo string.</param>
        ''' <param name="section_name">Representa o nome da seção dentro do arquivo INI onde a chave está localizada do tipo string.</param>
        ''' <param name="key_name">Represente o nome da chave cujo valor você deseja ler do tipo string.</param>
        ''' <param name="default_value">Representa o valor padrão a ser retornado caso a chave não seja encontrada do tipo string.</param>
        ''' <returns>Retorna um valor String correspondente ao valor da chave dentro da seção especificada do arquivo INI, ou o valor padrão se a chave não for encontrada.</returns>
        Public Function LeArquivoINI(ByVal file_name As String, ByVal section_name As String, ByVal key_name As String, ByVal default_value As String) As String

            Const MAX_LENGTH As Integer = 500
            Dim string_builder As New StringBuilder(MAX_LENGTH)

            GetPrivateProfileString(section_name, key_name, default_value, string_builder, MAX_LENGTH, file_name)

            Return string_builder.ToString()
        End Function
        ''' <summary>
        ''' Obtém o caminho completo do arquivo INI chamado "Banco_Dados.ini" localizado no diretório imediatamente acima do diretório de inicialização do aplicativo.
        ''' </summary>
        ''' <param name="nomeArquivo">Representa o nome do arquivo que deseja obter o caminho.</param>
        ''' <returns>Uma string que representa o caminho completo do arquivo do arquivo solicitado.</returns>
        ''' <remarks>
        ''' Esta função usa o caminho de inicialização do aplicativo, sobe um diretório e anexa o nome do arquivo INI.
        ''' </remarks>
        Public Function ObterArquivo(nomeArquivo As String) As String
            Dim caminhoDiretorio As String = Application.StartupPath
            Return IO.Path.Combine(caminhoDiretorio, nomeArquivo)

            'Dim caminhoDiretorio As String = Application.StartupPath
            'caminhoDiretorio = caminhoDiretorio.Substring(0, caminhoDiretorio.LastIndexOf("\"))
            'Return caminhoDiretorio & "\" & nomeArquivo
        End Function
        Public Function GravarArquivoBD(nomeArquivo As String, Instancia As String, Banco As String, valorSenha As String, Usuario As String)
            Dim CaminhoArquivo As String = ObterArquivo(nomeArquivo)

            'If chkCifrar.Checked Then
            '    valorSenha = geraHash(txtSenha.Text)
            'Else
            '    valorSenha = txtSenha.Text
            'End If
            WritePrivateProfileString("Geral", "Servidor", Instancia, CaminhoArquivo)
            WritePrivateProfileString("Geral", "Banco", Banco, CaminhoArquivo)
            WritePrivateProfileString("Seguranca", "Password", valorSenha, CaminhoArquivo)
            WritePrivateProfileString("Seguranca", "User", Usuario, CaminhoArquivo)
        End Function
        ''' <summary>
        ''' Esta função encontra o caractere especificado em uma string e retorna o trecho de texto entre a posição inicial e a posição do caractere encontrado na posição especificada.
        ''' </summary>
        ''' <param name="texto">Representa trecho do texto a ser verificado.</param>
        ''' <param name="caract">Represente o caractere da posição inicial.</param>
        ''' <param name="posicao">Representa posição do caractere encontrado na posição especificada.</param>
        ''' <returns></returns>
        Public Function EncontraCaractere(ByVal texto As String, ByVal caract As String, ByVal posicao As Integer) As String
            Dim y As Integer
            Dim posicoes As Integer
            Dim ultima As Integer
            Dim achou As Boolean
            Dim retorno As String
            posicoes = 1
            ultima = 0
            achou = False
            retorno = ""
            For y = 1 To Len(texto)

                If Mid(texto, y, 1) = caract Then
                    If posicoes = posicao Then
                        retorno = Mid(texto, IIf(ultima = 0, 1, ultima), IIf(ultima = 0, y - ultima - 1, y - ultima))
                        achou = True
                        Exit For
                    Else
                        achou = False
                    End If
                    posicoes = posicoes + 1
                    ultima = y + 1
                End If
            Next
            If achou = False Then
                retorno = Mid(texto, IIf(ultima = 0, 1, ultima), IIf(ultima = 0, y - ultima - 1, y - ultima))
            End If
            EncontraCaractere = retorno
        End Function
        ''' <summary>
        ''' Esta função converter uma string de data no formato "DDMMYYYY" em um objeto de data no formato "MM/DD/YYYY". 
        ''' </summary>
        ''' <param name="data">Representa uma data do tipo string.</param>
        ''' <returns></returns>
        Public Function MontaData(ByVal data As String) As Date
            Dim novadata As String
            'rearranja os caracteres da string de data para colocá-la no formato "DD/MM/YYYY
            novadata = Left(data, 2) & "/" & Mid(data, 3, 2) & "/" & Right(data, 4)
            'verifica se a data convertida é uma data válida 
            If IsDate(novadata) = False Then
                MontaData = Date.Today
                Exit Function
            End If
            MontaData = CDate(novadata)
        End Function
        ''' <summary>
        ''' Esta função calcula o número de espaços necessários para centralizar uma frase em um campo de tamanho fixo. 
        ''' </summary>
        ''' <param name="frase">Representa a frase a ser centralizada do tipo string.</param>
        ''' <param name="valor">Representa o valor do tipo integer.</param>
        ''' <returns></returns>
        Public Function Centraliza(ByVal frase As String, ByVal valor As Integer)

            Return (valor - frase.Length) / 2
        End Function
        ''' <summary>
        ''' Esta função converte uma string que representa um valor monetário em um formato numérico padrão.
        ''' </summary>
        ''' <param name="valor">A string que representa o valor monetário a ser convertido.</param>
        ''' <returns>Uma string formatada sem o símbolo de moeda e separadores de milhar, com ponto decimal no lugar da vírgula. Se a string de entrada estiver vazia, retorna "0".</returns>
        ''' <remarks>
        ''' Esta função remove o símbolo de moeda "R$", os pontos que são usados como separadores de milhar,
        ''' substitui a vírgula por um ponto decimal e remove quaisquer espaços em branco no início ou no fim da string.
        ''' </remarks>
        Public Function ConverterMoeda(ByVal valor As String)
            If valor = "" Then
                ConverterMoeda = 0
                Exit Function
            End If
            Dim novovalor As String = valor.Replace("R$", "").Replace(".", "")
            Return novovalor.Replace(",", ".").Trim
        End Function
        ''' <summary>
        ''' Esta função recebe uma data como entrada e retorna um valor inteiro representando o semestre ao qual a data pertence.
        ''' </summary>
        ''' <param name="data">Representa uma data no formato date.</param>
        ''' <returns>Retorna o semestre o qual a data está.</returns>
        Public Function DefineSemestre(ByVal data As Date) As Integer

            If data.Month < 7 Then
                Return 1
            Else
                Return 2
            End If
        End Function
        ''' <summary>
        ''' Este função retorna a data correspondente ao dia útil especificado em um determinado mês e ano.
        ''' </summary>
        ''' <param name="mes">Represente o mês do tipo integer.</param>
        ''' <param name="ano">Representa o ano do tipo integer.</param>
        ''' <param name="diautil">Representa o dia do tipo integer.</param>
        ''' <returns>Retorna a data do dias útil correspondente.</returns>
        Public Function DiasUteis(ByVal mes As Integer, ByVal ano As Integer, ByVal diautil As Integer) As Date
            Dim data, novadata As Date
            Dim x, dias As Integer
            data = CDate("01/" & mes & "/" & ano)
            x = 0
            dias = 0
            While dias < diautil
                novadata = data.AddDays(x)
                If novadata.DayOfWeek = DayOfWeek.Saturday Or novadata.DayOfWeek = DayOfWeek.Sunday Then
                Else
                    dias += 1
                End If
                x += 1
            End While
            Return novadata
        End Function
        ''' <summary>
        ''' Esta função adiciona zeros à esquerda de uma expressão até que ela atinja um determinado comprimento. 
        ''' </summary>
        ''' <param name="expr">Representa a expressão do tipo string.</param>
        ''' <param name="qtd">Representa a quantidade do tipo integer.</param>
        ''' <returns></returns>
        Public Function AcrescentaZeros(ByVal expr As String, ByVal qtd As Integer)
            Dim dd As String
            Dim x As Integer
            dd = ""
            If qtd <= Len(expr) Then
                Return expr
                Exit Function
            End If
            For x = 1 To (qtd - Len(expr))
                dd = dd & "0"
            Next
            Return dd & expr

        End Function
        ''' <summary>
        ''' Esta função é utilizadas para converter uma string que representa uma data em um objeto de data (Date) no formato desejado.
        ''' </summary>
        ''' <param name="data">Representa da data do tipo string.</param>
        ''' <param name="mascara">Representa a máscara da data do tipo string</param>
        ''' <returns></returns>
        Public Function TrataData(ByVal data As String, Optional ByVal mascara As String = "dd/MM/yyyy") As Date
            Dim dataValida As DateTime
            If DateTime.TryParse(data, dataValida) Then
                Return dataValida.ToString(mascara)
            Else
                dataValida = New DateTime(1900, 1, 1)
                Return dataValida.ToString(mascara)
            End If
        End Function
        ''' <summary>
        ''' Retorna o número de série (serial number) do disco rígido do computador.
        ''' </summary>
        ''' <param name="driveletter">Represente a letra da unidade de disco.</param>
        ''' <returns></returns>
        Public Function Serialhd(ByVal driveletter As String)

            Dim fso As Object
            Dim Drv As Object
            Dim DriveSerial As String
            'Cria um objeto FileSystemObject
            fso = CreateObject("Scripting.FileSystemObject")

            'Atribui a letra do drive atual se nada for especificado
            If driveletter <> "" Then
                Drv = fso.GetDrive(driveletter)
            Else
                Drv = fso.GetDrive(fso.GetDriveName("c:\"))
            End If

            With Drv
                If .IsReady Then
                    DriveSerial = Int(.SerialNumber)
                Else '"Drive não esta pronto!"
                    DriveSerial = -1
                End If
            End With

            'libera objetos
            Drv = Nothing
            fso = Nothing

            Serialhd = DriveSerial

        End Function

        ''' <summary>
        '''  Formata uma string de forma que apenas a primeira letra de cada palavra seja maiúscula e as demais sejam minúsculas. 
        ''' </summary>
        ''' <param name="palavra">Representa a palavra que deve ser informada no parâmetro, e deve ser um string.</param>
        ''' <returns>Retorna a palavras com o primeira letra em maiúsuclo.</returns>
        Public Function PrimeiraLetra(ByVal palavra As String)

            If palavra.Length = 0 Then
                Return ""
                Exit Function
            End If
            Dim x As Integer
            Dim nova As String
            nova = ""
            nova = palavra.Substring(0, 1).ToUpper

            For x = 1 To palavra.Length - 1
                If palavra.Substring(x - 1, 1).ToUpper = " " Then
                    nova = nova & palavra.Substring(x, 1).ToUpper
                Else
                    nova = nova & palavra.Substring(x, 1).ToLower
                End If
            Next
            Return nova
        End Function

        ''' <summary>
        ''' Extrair o nome do diretório pai de um determinado caminho de arquivo informado na variavel <paramref name="wcCAMINHO"/>.
        ''' </summary>
        ''' <param name="wcCAMINHO">Representa o caminho completo do diretório e deve ser uma string.</param>
        ''' <returns></returns>
        Public Function mPASTA(ByVal wcCAMINHO As String)
            Dim wctemp
            Dim xx As Integer
            For xx = 1 To Len(wcCAMINHO)
                If Left(Right(wcCAMINHO, xx + 1), 1) = "\" Then
                    wctemp = Len(Right(wcCAMINHO, xx))
                    mPASTA = Left(wcCAMINHO, Len(wcCAMINHO) - wctemp)
                    xx = Len(wcCAMINHO)
                End If
            Next
            Return ""
        End Function
        ''' <summary>
        ''' Retorna o último dia do mês para uma determinada data fornecida como entrada.
        ''' </summary>
        ''' <param name="wcdata">Representa a data da qual deseja saber o último dia.</param>
        ''' <returns>REtorna o último dis do mês.</returns>
        Public Function UltimoDia(ByVal wcdata As Date) As Date
            wcdata = wcdata.AddMonths(1)
            wcdata = wcdata.AddDays(-1)
            Return wcdata
        End Function
        ''' <summary>
        ''' Recebe uma string que pode estar acentuada e vai devolver em string convertendo de UTF para ISO.
        ''' </summary>
        ''' <param name="texto">Representa a strinn que se deseja converter para o formato ISO.</param>
        ''' <returns>Retorna string convertida para o formato ISO.</returns>
        Public Function UTF8toISO(texto As String) As String
            Dim isoEncoding = Encoding.GetEncoding("ISO-8859-1")
            Dim utfEncoding = Encoding.UTF8

            Dim bytesIso As Byte() = utfEncoding.GetBytes(texto)
            Dim bytesUtf As Byte() = Encoding.Convert(utfEncoding, isoEncoding, bytesIso)

            Dim textoISO = utfEncoding.GetString(bytesUtf)

            UTF8toISO = textoISO
        End Function
        ''' <summary>
        ''' Conta quanto checkbox estão marcados em uma datagridview
        ''' </summary>
        ''' <param name="dataGridView">Representa o nome do componente datagrid view.</param>
        ''' <param name="nomeDaColuna">Representa o nome da coluna do datagridview.</param>
        ''' <returns></returns>
        Public Function ContarCheckBoxMarcados(dataGridView As DataGridView, nomeDaColuna As String) As Integer
            Dim count As Integer = 0

            For Each row As DataGridViewRow In dataGridView.Rows
                Dim cell As DataGridViewCheckBoxCell = TryCast(row.Cells(nomeDaColuna), DataGridViewCheckBoxCell)
                If cell IsNot Nothing AndAlso cell.Value IsNot Nothing AndAlso CBool(cell.Value) Then
                    count += 1
                End If
            Next

            Return count
        End Function
        ''' <summary>
        ''' Arredonda o número para o próximo múltiplo do valor fornecido.
        ''' </summary>
        ''' <param name="numero">Número que será arredondado do tipo decimal.</param>
        ''' <param name="multiplicador">Múltiplo para o qual será arredondado do tipo decimal.</param>
        ''' <returns>Retorna o número arredondado para cima.</returns>
        Public Function ArredondarCima(ByVal numero As Decimal, ByVal multiplicador As Decimal) As Decimal
            Dim resto As Decimal = numero Mod multiplicador
            If resto = 0 Then
                Return numero
            Else
                Return numero + (multiplicador - resto)
            End If
        End Function
#End Region
    End Module
End Namespace
