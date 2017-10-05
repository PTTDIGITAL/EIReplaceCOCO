Public Class frmConfirmPassword

    Dim _ExPassword As String = "Export"
    Dim _ImportPassword As String = "Import"
    Dim _Mode As Integer

    Public Property Mode As Integer
        Get
            Return _Mode
        End Get
        Set(ByVal value As Integer)
            _Mode = value
        End Set
    End Property

    Enum ExportImport
        Export = 1
        Import = 2
    End Enum


    Private Sub frmConfirmPassword_Load(sender As Object, e As EventArgs) Handles Me.Load
        txtPassword.PasswordChar = "*"
        lblMode.Text = ""
        Select Case Mode
            Case "1"
                lblMode.Text = "ยืนยันรหัสผ่านสำหรับการนำออกข้อมูล"
            Case "2"
                lblMode.Text = "ยืนยันรหัสผ่านสำหรับการนำเข้าข้อมูล"
        End Select
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Dim password As String = ""
        If _Mode = ExportImport.Export Then
            password = _ExPassword
        ElseIf _Mode = ExportImport.Import
            password = _ImportPassword
        End If

        If txtPassword.Text.Trim = "" Or txtPassword.Text.Trim <> password Then
            Using New Centered_MessageBox(Me)
                Dim confirm As DialogResult = MessageBox.Show("รหัสผ่านไม่ถูกต้อง", "", MessageBoxButtons.OK)
            End Using
            txtPassword.Select()
        Else
            Me.DialogResult = DialogResult.OK
        End If

    End Sub

    Private Sub chkCheckPassword_CheckedChanged(sender As Object, e As EventArgs) Handles chkCheckPassword.CheckedChanged
        If chkCheckPassword.Checked Then
            txtPassword.PasswordChar = ""
        Else
            txtPassword.PasswordChar = "*"
        End If
    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub

    Private Sub txtPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnOK_Click(sender, e)
        End If
    End Sub
End Class