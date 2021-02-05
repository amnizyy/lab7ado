Imports System.Data.OleDb

Public Class Form1
    Dim inc As Integer
    Dim MaxRows As Integer
    Dim con As New OleDb.OleDbConnection
    Dim dbProvider As String
    Dim dbSource As String
    Dim ds As New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Dim sql As String
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
        dbSource = "Data Source = C:\Users\erza6\Documents\Record.mdb"

        con.ConnectionString = dbProvider & dbSource
        con.Open()

        con.close()

        sql = "SELECT * FROM Student"
        da = New OleDb.OleDbDataAdapter(sql, con)
        da.Fill(ds, "Record")
        con.Close()

        MaxRows = ds.Tables("Record").Rows.Count
        inc = -1

    End Sub
    Private Sub navigateRecords()
        txtMatricNum.Text = ds.Tables("Record").Rows(inc).Item(1)
        txtFullName.Text = ds.Tables("Record").Rows(inc).Item(2)
        txtAddress.Text = ds.Tables("Record").Rows(inc).Item(3)
        txtEmail.Text = ds.Tables("Record").Rows(inc).Item(4)

    End Sub

    Private Sub btnForward_Click(sender As Object, e As EventArgs) Handles btnForward.Click
        If inc <> MaxRows - 1 Then
            inc = inc + 1
            navigateRecords()
        Else
            MsgBox("No More Rows")
        End If
    End Sub

    Private Sub btnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click
        If inc > 0 Then
            inc = inc - 1
            navigateRecords()
        ElseIf inc = -1 Then
            MsgBox("No Records Yet")
        ElseIf inc = 0 Then
            MsgBox("First  Record")
        End If
    End Sub

    Private Sub btnForwarddouble_Click(sender As Object, e As EventArgs) Handles btnForwarddouble.Click
        If inc <> MaxRows - 1 Then
            inc = MaxRows - 1
            navigateRecords()

        End If
    End Sub

    Private Sub btnPreviousdouble_Click(sender As Object, e As EventArgs) Handles btnPreviousdouble.Click
        If inc <> 0 Then
            inc = 0
            navigateRecords()

        End If
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Dim cb As New OleDb.OleDbCommandBuilder(da)

        ds.Tables("Record").Rows(inc).Item(1) = txtMatricNum.Text
        ds.Tables("Record").Rows(inc).Item(2) = txtFullName.Text
        ds.Tables("Record").Rows(inc).Item(3) = txtAddress.Text
        ds.Tables("Record").Rows(inc).Item(4) = txtEmail.Text


        da.Update(ds, "Record")
        MsgBox("Data updated successfully")
    End Sub

    Private Sub btnCommit_Click(sender As Object, e As EventArgs) Handles btnCommit.Click

        If inc <> 1 Then

            Dim cb As New OleDb.OleDbCommandBuilder(da)
            Dim dsNewRow As DataRow

            dsNewRow = ds.Tables("Record").NewRow()

            dsNewRow.Item("MatricNum") = txtMatricNum.Text
            dsNewRow.Item("Fullname") = txtFullName.Text
            dsNewRow.Item("Address") = txtAddress.Text
            dsNewRow.Item("Email") = txtEmail.Text


            ds.Tables("Record").Rows.Add(dsNewRow)
            da.Update(ds, "Record")
            MsgBox("New Record added to the database")


            txtMatricNum.Clear()
            txtFullName.Clear()
            txtAddress.Clear()
            txtEmail.Clear()



        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim cb As New OleDb.OleDbCommandBuilder(da)


        If MessageBox.Show("Do you really want to Delete this Record?",
        "Delete", MessageBoxButtons.YesNo,
         MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then

            ds.Tables("Record").Rows(inc).Delete()
            MaxRows = MaxRows - 1
            da.Update(ds, "Record")
            txtMatricNum.Clear()
            txtFullName.Clear()
            txtAddress.Clear()
            txtEmail.Clear()

        Else

            MsgBox("Operation Cancelled")

            Exit Sub

        End If

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        txtMatricNum.Clear()
        txtFullName.Clear()
        txtAddress.Clear()
        txtEmail.Clear()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        btnCommit.Enabled = True
        btnAdd.Enabled = False
        btnUpdate.Enabled = False
        btnDelete.Enabled = False

        txtMatricNum.Clear()
        txtFullName.Clear()
        txtAddress.Clear()
        txtEmail.Clear()

        inc = 0
    End Sub
End Class
