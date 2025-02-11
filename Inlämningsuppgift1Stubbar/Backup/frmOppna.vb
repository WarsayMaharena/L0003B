Imports System.Data.SqlClient
Imports System.Data
Public Class frmOppna
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
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.Location = New System.Drawing.Point(8, 8)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(168, 238)
        Me.ListBox1.TabIndex = 0
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(184, 32)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(104, 32)
        Me.btnEdit.TabIndex = 1
        Me.btnEdit.Text = "Redigera"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(184, 72)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(104, 40)
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "Ta Bort"
        '
        'frmOppna
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.ListBox1)
        Me.Name = "frmOppna"
        Me.Text = "frmOppna"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private mySqlConnection As New ClassConnection
    Private myAdapter As SqlDataAdapter
    Private mySqlCommand As New SqlCommand
    

    Private Sub frmOppna_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        uppdateraListan()
    End Sub
    Private Sub uppdateraListan()
        mySqlCommand.Connection = mySqlConnection.ReturneraKoppling()
        mySqlCommand.CommandType = CommandType.Text
        mySqlCommand.CommandText = "EGEN KOD: SQL-SATS FÖR ATT VÄLJA ALLT OM ALLA ELEVER SORTERAT PÅ EFTERNAMN"

        mySqlConnection.open()
        myAdapter = New SqlDataAdapter(mySqlCommand)
        Dim ds As DataSet = New DataSet
        myAdapter.Fill(ds, "Elev")
        mySqlConnection.Close()
        Dim i As Integer, pnr As String, namn As String, elever As New ArrayList

        For i = 0 To ds.Tables("Elev").Rows.Count - 1
            With ds.Tables("Elev").Rows(i)
                namn = Trim(.Item("efternamn")) & ", " & Trim(.Item("fornamn"))
                pnr = Trim(.Item("pnr"))
                elever.Add(New ElevClass(pnr, namn))
            End With
        Next
        ListBox1.DataSource = elever
        ListBox1.DisplayMember = "namn"
        ListBox1.ValueMember = "pnr"

    End Sub


    Private Sub ListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.Click
        ' MsgBox(ListBox1.SelectedValue)

    End Sub



    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If ListBox1.SelectedIndex <> -1 Then
            Dim frmRedigera As New frmElev(ListBox1.SelectedValue)
            frmRedigera.ShowDialog(Me)
            uppdateraListan()
        Else
            MsgBox("Markera en elev")
        End If


    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If ListBox1.SelectedIndex <> -1 Then
            mySqlConnection.Open()
            mySqlCommand.CommandText = "EGEN KOD: SQL-SATS FÖR ATT TA BORT ELEV MED PERSONUMMER SOM FINNS I VARIABELN ListBox1.SelectedValue()."

            mySqlCommand.ExecuteNonQuery()
            mySqlCommand.CommandText = "EGEN KOD: SQL-SATS FÖR ATT TA BORT DEN ELEVENS KURSER, DVS UR TABELLEN KURSELEV."

            mySqlCommand.ExecuteNonQuery()
            mySqlConnection.Close()
            uppdateraListan()


        Else
            MsgBox("Markera den elev du vill ta bort.")
        End If

    End Sub
End Class
