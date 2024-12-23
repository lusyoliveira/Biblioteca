Imports MySql.Data.MySqlClient

Namespace Classes.Conexao

    Public Class ConexaoMYSQl

#Region "PROPRIEDADES"

        Public Property StringConn As String
        Public Property connDB As MySqlConnection

#End Region

#Region "CONSTRUTOR"

        Public Sub New(Servidor As String, vUser As String, vPassword As String, vDataBase As String)

            'server=<SERVIDOR>;database=<SCHEMA>;User Id=<USUARIO>;Password=SENHA"
            StringConn = "server=" + Servidor + ";database=" + vDataBase + ";User Id=" + vUser + ";Password=" + vPassword + ";"
            connDB = New MySqlConnection(StringConn)
            connDB.Open()

        End Sub

#End Region

#Region "MÉTODOS"

        Public Function SQLQuery(SQL As String, ByRef DT As DataTable) As String

            Try
                Dim myCommand As IDbCommand = New MySqlCommand(SQL, connDB)
                myCommand.CommandTimeout = 0
                Dim myReader As IDataReader = myCommand.ExecuteReader
                DT.Load(myReader)
                Return ""

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        End Function

        Public Function SQLCommand(SQL As String) As String

            Try
                Dim myCommand As IDbCommand = New MySqlCommand(SQL, connDB)
                myCommand.CommandTimeout = 0
                myCommand.ExecuteReader()
                Return ""

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        End Function

        Public Sub Close()
            connDB.Close()
        End Sub

#End Region

    End Class

End Namespace


