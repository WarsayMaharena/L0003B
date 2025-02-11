Imports System.Data.SqlClient
Imports System.Data
Public Class frmKurs
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(40, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(152, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Kurser"
        '
        'ListBox1
        '
        Me.ListBox1.Location = New System.Drawing.Point(40, 32)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(176, 186)
        Me.ListBox1.TabIndex = 1
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(224, 40)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(64, 32)
        Me.btnNew.TabIndex = 2
        Me.btnNew.Text = "Skapa Ny"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(224, 136)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(64, 32)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "Ta Bort"
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(224, 88)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(64, 32)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "Redigera"
        '
        'frmKurs
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnEdit)
        Me.Name = "frmKurs"
        Me.Text = "frmKurs"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private mySqlConnection As New ClassConnection
    Private myAdapter As SqlDataAdapter
    Private mySqlCommand As New SqlCommand
    Private Sub frmKurs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        UpdateList()
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        mySqlConnection.Open()
        Dim nyttKursNamn As String = InputBox("Ange namnet på den kurs du vill skapa", "Skapa ny kurs")
        While nyttKursNamn = "" Or ListBox1.Items.Contains(nyttKursNamn) Or nyttKursNamn.Length > 15
            nyttKursNamn = InputBox("Felaktigt kursnamn! Ange namnet på den kurs du vill skapa (max 15 tecken)", "Skapa ny kurs")
        End While

        mySqlCommand.Connection = mySqlConnection.ReturneraKoppling()
        mySqlCommand.CommandType = CommandType.Text
        mySqlCommand.CommandText = "EGEN KOD: SQL-SATS FÖR ATT LÄGGA IN DEN NYA KURSEN. NAMNET FINNS I VARIABELN nyttKursNamn."

        mySqlCommand.ExecuteNonQuery()

        mySqlConnection.Close()
        UpdateList()
    End Sub
    Private Sub UpdateList()


        mySqlConnection.open()
        mySqlCommand.Connection = mySqlConnection.ReturneraKoppling()
        mySqlCommand.CommandType = CommandType.Text
        mySqlCommand.CommandText = "EGEN KOD: SQL-SATS FÖR ATT VÄLJA ALLA KURSER"

        myAdapter = New SqlDataAdapter(mySqlCommand)
        Dim i As Integer = 0
        While ListBox1.Items.Count <> 0
            ListBox1.Items.RemoveAt(0)
        End While
        Dim myDataSet As New DataSet
        myAdapter.Fill(myDataSet, "Kurs")
        i = 0
        While i < myDataSet.Tables("Kurs").Rows.Count
            ListBox1.Items.Add(Trim(myDataSet.Tables("Kurs").Rows(i).Item("kursnamn")))
            i = i + 1
        End While
        mySqlConnection.close()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        mySqlConnection.open()
        mySqlCommand.Connection = mySqlConnection.ReturneraKoppling()
        mySqlCommand.CommandType = CommandType.Text
        Dim knamn As String = ListBox1.SelectedItem
        mySqlCommand.CommandText = "EGEN KOD: SQL-SATS FÖR ATT VÄLJA UR KURSELEV ALLA FÖREKOMSTER FÖR DEN KURSEN VARS NAMN FINNS I VARIABELN knamn."

        myAdapter = New SqlDataAdapter(mySqlCommand)
        Dim ds As New DataSet
        myAdapter.Fill(ds, "Kurser")
        If ds.Tables("Kurser").Rows.Count > 0 Then
            MsgBox("Tyvärr, det finns personer registrerade på den här kursen. Du kan inte ta bort den då.")
            mySqlConnection.Close()
            Exit Sub
        End If
        mySqlCommand.CommandText = "EGEN KOD: SQL-SATS FÖR ATT TA BORT KURSEN MED KURSNAMN I ListBox1.Items.Item(ListBox1.SelectedIndex)"

        mySqlCommand.ExecuteNonQuery()

        mySqlConnection.Close()
        UpdateList()
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If ListBox1.SelectedIndex < 0 Then
            MsgBox("Markera en kurs du vill redigera", MsgBoxStyle.Critical)
            Exit Sub
        End If
        Dim nyttNamn As String = InputBox("Ange nytt kursnamn", "Redigera kursnamn", ListBox1.SelectedItem)
        While ListBox1.Items.Contains(nyttNamn) Or nyttNamn.Length > 15
            nyttNamn = InputBox("Felaktigt kursnamn! Ange nyttNamn (max 15 tecken)", "Skapa ny kurs", nyttNamn)
        End While
        If nyttNamn = "" Then Exit Sub
        mySqlConnection.open()
        mySqlCommand.Connection = mySqlConnection.ReturneraKoppling()
        mySqlCommand.CommandType = CommandType.Text

        mySqlCommand.CommandText = "EGEN KOD: SQL-SATS FÖR ATT UPPDATERA KURSENS NAMN. NAMNET FINNS I ListBox1.SelectedItem och det nya namnet finns i variabeln nyttNamn"

        mySqlCommand.ExecuteNonQuery()

        mySqlCommand.CommandText = "EGEN KOD: SQL-SATS FÖR ATT UPPDATERA KURSENS NAMN FÖR DE ELEVER SOM HAR DEN KURSEN VARS NAMN ÄNDRATS. NAMNET FINNS I ListBox1.SelectedItem och det nya namnet finns i variabeln nyttNamn."

        mySqlCommand.ExecuteNonQuery()

        mySqlConnection.Close()
        UpdateList()

    End Sub
End Class
