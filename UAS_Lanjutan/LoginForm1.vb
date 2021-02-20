Imports System.Data.OleDb
Public Class LoginForm1
    Dim con As New OleDbConnection
    Dim iFail As Integer
    Sub Open_Koneksi()
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0" & _
                                ";Data Source=uas.accdb;" & _
                                "Persist Security Info=False;"
        con.Open()
    End Sub

    Function CheckLogin(ByVal Username As String, _
                         ByVal Password As String) As Integer
        Dim cmd As New OleDbCommand
        Dim objvalue As Object
        If Not con.State = ConnectionState.Open Then Open_Koneksi()
        Try
            cmd.Connection = con
            cmd.CommandText = "SELECT COUNT(username) AS getin " & _
                "FROM tbluser WHERE Username = " & _
                "'" & Username & "' AND " & _
                "Password = '" & Password & "'"
            objvalue = cmd.ExecuteScalar
            con.Close()
            If objvalue Is Nothing Then
                Return 0
            Else
                Return objvalue.ToString
            End If
        Catch myerror As OleDbException
            MessageBox.Show("Error : " & myerror.Message)
        Finally
            con.Dispose()
        End Try
        Return 0
    End Function

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim strUsername As String = UsernameTextBox.Text
        Dim strPassword As String = PasswordTextBox.Text

        If strUsername = String.Empty Then _
            MsgBox("Username Kosong!") : Exit Sub
        If strPassword = String.Empty Then _
            MsgBox("Password Kosong!") : Exit Sub
        Try
            If CheckLogin(strUsername, strPassword) > 0 Then
                MsgBox("Selamat Datang ADMIN!!")
                Me.Hide()
                FormMaster.Show()
            Else
                iFail = iFail + 1
                If iFail >= 3 Then
                    MsgBox("ANDA GAGAL LOGIN 3x" & vbCrLf & _
                           "APLIKASI DITUTUP!")
                    End
                End If
                MsgBox("USERNAME / PASSWORD SALAH!" & vbCrLf & _
                       "MOHON DI CEK KEMBALI")
            End If
        Catch ex As Exception
            MsgBox("Error Login : " & ex.Message)
        End Try
    End Sub

    Private Sub LoginForm1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim pes As String
        pes = MsgBox("APA KAMU YAKIN", vbYesNo, "LOGOUT")
        If pes = vbYes Then
            End
        ElseIf pes = vbNo Then
            e.Cancel = True
        End If
    End Sub

    Private Sub LoginForm1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
        iFail = 0
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub
End Class
