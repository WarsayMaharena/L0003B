
Imports System.Data.SqlClient
Imports System.Data
Public Class frmElev
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        gammaltPnr = ""
        hämtakurser()
    End Sub
    Public Sub New(ByVal personnummer As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        uppdateraelev(personnummer)


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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox2 As System.Windows.Forms.ListBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents txtPnr As System.Windows.Forms.TextBox
    Friend WithEvents TxtFor As System.Windows.Forms.TextBox
    Friend WithEvents txtEfter As System.Windows.Forms.TextBox
    Friend WithEvents txtOrt As System.Windows.Forms.TextBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.txtPnr = New System.Windows.Forms.TextBox()
        Me.TxtFor = New System.Windows.Forms.TextBox()
        Me.txtEfter = New System.Windows.Forms.TextBox()
        Me.txtOrt = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.ListBox2 = New System.Windows.Forms.ListBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtPnr
        '
        Me.txtPnr.Location = New System.Drawing.Point(288, 30)
        Me.txtPnr.Name = "txtPnr"
        Me.txtPnr.Size = New System.Drawing.Size(272, 31)
        Me.txtPnr.TabIndex = 0
        '
        'TxtFor
        '
        Me.TxtFor.Location = New System.Drawing.Point(288, 89)
        Me.TxtFor.Name = "TxtFor"
        Me.TxtFor.Size = New System.Drawing.Size(272, 31)
        Me.TxtFor.TabIndex = 1
        '
        'txtEfter
        '
        Me.txtEfter.Location = New System.Drawing.Point(288, 148)
        Me.txtEfter.Name = "txtEfter"
        Me.txtEfter.Size = New System.Drawing.Size(272, 31)
        Me.txtEfter.TabIndex = 2
        '
        'txtOrt
        '
        Me.txtOrt.Location = New System.Drawing.Point(288, 207)
        Me.txtOrt.Name = "txtOrt"
        Me.txtOrt.Size = New System.Drawing.Size(272, 31)
        Me.txtOrt.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(240, 44)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Personnummer"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 89)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(240, 44)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Förnamn"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 207)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(240, 44)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Ort"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 148)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(240, 44)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Efternamn"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ListBox1
        '
        Me.ListBox1.ItemHeight = 25
        Me.ListBox1.Location = New System.Drawing.Point(112, 310)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(240, 229)
        Me.ListBox1.TabIndex = 5
        '
        'ListBox2
        '
        Me.ListBox2.ItemHeight = 25
        Me.ListBox2.Location = New System.Drawing.Point(384, 310)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(288, 229)
        Me.ListBox2.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(112, 266)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(240, 44)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Existerande Kurser"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(384, 266)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(240, 44)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Valda Kurser"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(96, 591)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(256, 59)
        Me.btnSave.TabIndex = 8
        Me.btnSave.Text = "Spara"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(400, 591)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(256, 59)
        Me.btnCancel.TabIndex = 9
        Me.btnCancel.Text = "Avbryt"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmElev
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(10, 24)
        Me.ClientSize = New System.Drawing.Size(773, 706)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.ListBox2)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtOrt)
        Me.Controls.Add(Me.txtEfter)
        Me.Controls.Add(Me.TxtFor)
        Me.Controls.Add(Me.txtPnr)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label6)
        Me.Name = "frmElev"
        Me.Text = "Elev"
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Private mySqlConnection As New ClassConnection
    Private myAdapter As SqlDataAdapter
    Private mySqlCommand As New SqlCommand
    Private ds As DataSet
    Private gammaltPnr As String

    Private Sub hämtakurser()
        mySqlCommand.Connection = mySqlConnection.ReturneraKoppling()
        mySqlCommand.CommandType = CommandType.Text
        mySqlCommand.CommandText = "SELECT kursnamn FROM Kurs990730"
        '"EGEN KOD: SQL-SATS FÖR ATT HÄMTA KURSER"

        mySqlConnection.open()
        myAdapter = New SqlDataAdapter(mySqlCommand)
        mySqlConnection.close()
        ds = New DataSet

        myAdapter.Fill(ds, "Kurs")
        ListBox1.DataSource = ds.Tables("Kurs")
        ListBox1.ValueMember = "Kursnamn"

    End Sub
    Private Sub uppdateraelev(ByVal personnummer As String)
        gammaltPnr = personnummer
        btnSave.Text = "Uppdatera"
        hämtakurser()

        mySqlConnection.open()
        mySqlCommand.CommandText = "SELECT kursnamn FROM KursElev990730 WHERE pnr = '" & personnummer & "'"
        '"EGEN KOD: SQL-SATS FÖR ATT HÄMTA ELEVS KURSER FÖR DEN ELEV SOM MARKERATS I LISTBOXEN. ELEVENS PERSONNUMMER FINNS I VARIABELN personnummer."

        myAdapter = New SqlDataAdapter(mySqlCommand)


        ds = New DataSet
        myAdapter.Fill(ds, "Elev")
        Dim i As Integer
        For i = 0 To ds.Tables("Elev").Rows.Count - 1
            ListBox2.Items.Add(ds.Tables("Elev").Rows(i).Item("kursnamn"))
        Next
        ListBox2.DisplayMember = "kursNamn"

        mySqlCommand.CommandText = "SELECT pnr, fornamn, efternamn, ort FROM Elev990730 WHERE pnr = '" & personnummer & "'"
        '"EGEN KOD: SQL-SATS FÖR ATT HÄMTA ELEVEN MED PERSONNUMMER SOM FINNS I VARIABLEN personnummer."
        myAdapter = New SqlDataAdapter(mySqlCommand)


        ds = New DataSet
        myAdapter.Fill(ds, "Elev")

        txtPnr.Text = Trim(ds.Tables("Elev").Rows(0).Item("pnr"))
        TxtFor.Text = Trim(ds.Tables("Elev").Rows(0).Item("fornamn"))
        txtEfter.Text = Trim(ds.Tables("Elev").Rows(0).Item("efternamn"))
        txtOrt.Text = Trim(ds.Tables("Elev").Rows(0).Item("ort"))

        mySqlConnection.Close()
    End Sub



    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick
        If Not ListBox2.Items.Contains(ListBox1.SelectedValue) Then
            ListBox2.Items.Add(ListBox1.SelectedValue)
        End If

    End Sub

    Private Sub ListBox2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox2.DoubleClick
        ListBox2.Items.RemoveAt(ListBox2.SelectedIndex)
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim nyttPnr As String = Trim(txtPnr.Text)
        If Not uppgifternaÄrGodkända() Then Exit Sub

        mySqlConnection.Open()
        mySqlCommand.CommandText = "SELECT * FROM Elev990730 WHERE pnr = '" & nyttPnr & "'"
        '"EGEN KOD: SQL-SATS FÖR ATT HÄMTA ELEV MED PERSONNUMMER SOM FYLLTS I I TEXTBOXEN OCH SOM LAGTS IN I VARIABELN nyttPnr ENLIGT OVAN. VI SKA KOLLA ATT PERSONNUMRET INTE FINNS REDAN"

        myAdapter = New SqlDataAdapter(mySqlCommand)

        myAdapter.Fill(ds, "Eleven")

        If ds.Tables("Eleven").Rows.Count <> 0 And nyttPnr <> gammaltPnr Then 'Det finns redan en elev med "nyttPnr" som personnummer'
            MsgBox("Personnumret finns redan registrerat i databasen")
            mySqlConnection.close()
            Exit Sub
        End If

        If Me.Owner.Text = "frmOppna" Then

            mySqlCommand.CommandText = "DELETE FROM Elev990730 WHERE pnr = '" & gammaltPnr & "'"
            '"EGEN KOD: SQL-SATS FÖR ATT TA BORT ELEV MED PERSONUMMER SOM FINNS I VARIABELN gammaltPnr."

            mySqlCommand.ExecuteNonQuery()
            mySqlCommand.CommandText = "DELETE FROM KursElev990730 WHERE pnr = '" & gammaltPnr & "'"
            '"EGEN KOD: SQL-SATS FÖR ATT TA BORT DE ELEVENS ALLA KURSER."

            mySqlCommand.ExecuteNonQuery()
        End If

        mySqlCommand.CommandText = "INSERT INTO Elev990730 (pnr, fornamn, efternamn, ort) " &
                               "VALUES ('" & nyttPnr & "', '" & TxtFor.Text & "', '" & txtEfter.Text & "', '" & txtOrt.Text & "')"
        '"EGEN KOD: SQL-SATS FÖR ATT LÄGGA ELEVEN (pnr, efternamn; förnamn OCH ort. PERSONNUMRET FINNS I VARIABELN nyttPnr OCH ÖVRIGA UPPGIFTER I TEXTBOXARNA."

        mySqlCommand.ExecuteNonQuery()

        If ListBox2.Items.Count = 0 Then
            mySqlConnection.close()
            Me.Close()
            Exit Sub
        End If

        Dim i
        For i = 0 To ListBox2.Items.Count - 1
            mySqlCommand.CommandText = "INSERT INTO KursElev990730 (pnr, kursnamn) " &
                                   "VALUES ('" & nyttPnr & "', '" & ListBox2.Items.Item(i) & "')"
            '"EGEN KOD: SQL-SATS FÖR ATT LÄGGA I ELEVENS KURSER(ListBox2.Items.Item(i)"

            mySqlCommand.ExecuteNonQuery()
        Next
        mySqlConnection.close()
        Me.Close()
    End Sub

    Private Function uppgifternaÄrGodkända() As Boolean
        ErrorProvider1.Dispose()
        If txtPnr.Text = "" Then
            Me.ErrorProvider1.SetError(txtPnr, "Ange ett personnummer.")
            Return False
        End If
        Return True
    End Function

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub


End Class
