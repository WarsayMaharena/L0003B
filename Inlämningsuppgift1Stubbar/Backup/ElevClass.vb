Public Class ElevClass


    Private mypersonnummer As String
    Private mynamn As String

    Public Sub New(ByVal pnr As String, ByVal namn As String)
        MyBase.New()
        Me.mypersonnummer = pnr
        Me.mynamn = namn
    End Sub
    Public ReadOnly Property namn()
        Get
            Return Me.mynamn
        End Get
    End Property

    Public ReadOnly Property pnr()
        Get
            Return Me.mypersonnummer
        End Get
    End Property

End Class
