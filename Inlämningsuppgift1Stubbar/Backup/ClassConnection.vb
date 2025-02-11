Option Explicit On 
Option Strict On
'Upprättar en connection till databasen
'Här ligger sker kopplingen mot databasen vilket innebär att om man 
'byter db behöver man bara ändra på ett ställe
Public Class ClassConnection
    'Mot citrix
    Dim Koppling As New System.Data.SqlClient.SqlConnection("Initial Catalog=EmbeddedVB; Data Source=LUiis02; Integrated Security=true")
    'Mot access
    'Dim Koppling As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.jet.OLEDB.4.0;Data source=Poster.mdb")
    'lokalt mot SQL-server
    'Dim Koppling As New System.Data.SqlClient.SqlConnection("Initial Catalog=Testar;Data Source=(local); Integrated Security=true;")

    Public ReadOnly Property ReturneraKoppling() As SqlClient.SqlConnection
        Get
            Return Koppling
        End Get
    End Property

    Sub open()
        Koppling.Open()
    End Sub

    Sub close()
        Koppling.Close()
    End Sub

End Class