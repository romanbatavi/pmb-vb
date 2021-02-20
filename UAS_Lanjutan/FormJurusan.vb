Imports System.Data.OleDb
Public Class FormJurusan
    Sub kosong()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox1.Focus()
    End Sub
    Sub Isi()
        TextBox2.Clear()
        TextBox2.Focus()
    End Sub
    Sub TampilkanJenis()
        da = New OleDbDataAdapter("Select * From tbljurusan", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "tbljurusan")
        DataGridView1.DataSource = ds.Tables("tbljurusan")
        DataGridView1.Refresh()
    End Sub
    Sub AturGrid()
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).Width = 100
        DataGridView1.Columns(0).HeaderText = "KODE JURUSAN"
        DataGridView1.Columns(1).HeaderText = "JURUSAN"
        DataGridView1.Columns(2).HeaderText = "SKS"
        DataGridView1.Columns(3).HeaderText = "BPP"
    End Sub

    Private Sub FormJenis_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call TampilkanJenis()
        Call kosong()
        Call AturGrid()
    End Sub
    Private Sub FormJurusan_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Application.Exit()
    End Sub

    Private Sub FormJurusan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        FormMaster.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("DATA BELUM LENGKAP")
            TextBox1.Focus()
            Exit Sub
        Else
            cmd = New OleDbCommand("Select * From tbljurusan where KodeJurusan='" & TextBox1.Text & "'", Conn)
            rd = cmd.ExecuteReader
            rd.Read()
            If Not rd.HasRows Then
                Dim simpan As String = "Insert into tbljurusan(KodeJurusan,Jurusan,SKS,BPP)values " & _
                    "('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
                cmd = New OleDbCommand(simpan, Conn)
                cmd.ExecuteNonQuery()
                MsgBox("SIMPAN DATA BERHASIL", MsgBoxStyle.Information, "Perhatian")
            End If
            Call TampilkanJenis()
            Call kosong()
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("KODE JURUSAN BELUM DIISI")
            TextBox1.Focus()
            Exit Sub
        Else
            Dim Ubah As String = "Update tbljurusan set " & _
                "Jurusan='" & TextBox2.Text & "'," & _
                "SKS='" & TextBox3.Text & "'," & _
                "BPP='" & TextBox4.Text & "'" & _
                "where KodeJurusan='" & TextBox1.Text & "'"
            cmd = New OleDbCommand(Ubah, Conn)
            cmd.ExecuteNonQuery()
            MsgBox("UBAH DATA SUKSES", MsgBoxStyle.Information, "PERHATIAN")
            Call TampilkanJenis()
            Call kosong()
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("KODE JURUSAN BELUM DIISI")
            TextBox1.Focus()
            Exit Sub
        Else
            If MessageBox.Show("JURUSAN AKAN DI HAPUS " & TextBox1.Text & "?", "",
                               MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                cmd = New OleDbCommand("Delete * From tbljurusan where KodeJurusan='" & TextBox1.Text & "'", Conn)
                cmd.ExecuteNonQuery()
                Call kosong()
                Call TampilkanJenis()
            Else
                Call kosong()
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call kosong()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            cmd = New OleDbCommand("Select * From Jurusan where KodeJurusan='" & TextBox1.Text & "'", Conn)
            rd = cmd.ExecuteReader
            rd.Read()
            If rd.HasRows = True Then
                TextBox2.Text = rd.Item(1)
                TextBox2.Focus()
            Else
                Call Isi()
                TextBox2.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            TextBox2.Text = UCase(TextBox2.Text)
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        With DataGridView1.Rows.Item(i)
            Me.TextBox1.Text = .Cells(0).Value
            Me.TextBox2.Text = .Cells(1).Value
            Me.TextBox3.Text = .Cells(2).Value
            Me.TextBox4.Text = .Cells(3).Value
        End With
    End Sub
End Class