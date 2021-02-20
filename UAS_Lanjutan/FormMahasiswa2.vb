Imports System.Data.OleDb
Public Class FormMahasiswa2
    Public Sub TampilAntrian()
        Call Koneksi()
        Try
            ComboBox1.SelectedIndex = -1
            Str = "Select KodeJurusan From tbljurusan ORDER BY KodeJurusan ASC"
            cmd = New OleDbCommand(Str, Conn)
            rd = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    ComboBox1.Items.Add(rd.Item("KodeJurusan"))
                End While
            End If
            Conn.Close()
        Catch ex As Exception
            MsgBox("Error : " & ex.Message)
        End Try
    End Sub
    Sub Kosong()
        TextBox1.Clear()
        ComboBox1.Text = ""
        TextBox2.Clear() 'Roman Batavi 2017230073
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox1.Focus()
    End Sub
    Sub TidakAktif()
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
    End Sub
    Sub Isi()
        ComboBox1.Text = ""
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        ComboBox1.Focus()
    End Sub
    Sub TampilBuku()
        da = New OleDbDataAdapter("Select * From [tblmahasiswa2]", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "tblmahasiswa2")
        DataGridView1.DataSource = ds.Tables("tblmahasiswa2")
        DataGridView1.Refresh()
    End Sub
    'Sub TampilkanJenis()
    '    cmd = New OleDbCommand("Select KodeJurusan From tbljurusan", Conn)
    '    rd = cmd.ExecuteReader
    '    Do While rd.Read
    '        ComboBox1.Items.Add(rd.Item(0))
    '    Loop
    'End Sub
    Sub AturGrid()
        DataGridView1.Columns(0).Width = 90
        DataGridView1.Columns(1).Width = 90 'Roman Batavi 2017230073
        DataGridView1.Columns(2).Width = 90
        DataGridView1.Columns(3).Width = 90
        DataGridView1.Columns(4).Width = 90
        DataGridView1.Columns(5).Width = 90
        DataGridView1.Columns(6).Width = 90
        DataGridView1.Columns(7).Width = 90
        DataGridView1.Columns(0).HeaderText = "NO PENDAFTARAN"
        DataGridView1.Columns(1).HeaderText = "KODE JURUSAN"
        DataGridView1.Columns(2).HeaderText = "NAMA"
        DataGridView1.Columns(3).HeaderText = "ALAMAT"
        DataGridView1.Columns(4).HeaderText = "AGAMA"
        DataGridView1.Columns(5).HeaderText = "SKS"
        DataGridView1.Columns(6).HeaderText = "BPP"
        DataGridView1.Columns(7).HeaderText = "TOTAL"
    End Sub
    Private Sub FormMahasiswa2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Application.Exit()
    End Sub
    Private Sub FormBuku_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TampilAntrian()
        Call Koneksi()
        'Call TampilkanJenis()
        Call TampilBuku()
        Call Kosong()
        Call AturGrid()
        Call TidakAktif()
        Me.CenterToScreen()
    End Sub
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 50
        If e.KeyChar = Chr(13) Then
            cmd = New OleDbCommand("Select * From tblmahasiswa2 where NoPendaftaran='" & TextBox1.Text & "'", Conn)
            rd = cmd.ExecuteReader
            If rd.HasRows = True Then
                ComboBox1.Text = rd.Item(1)
                TextBox2.Text = rd.Item(2)
                TextBox3.Text = rd.Item(3)
                TextBox4.Text = rd.Item(4)
                TextBox5.Text = rd.Item(5)
                TextBox6.Text = rd.Item(6)
                TextBox7.Text = rd.Item(7)
                TextBox2.Focus()
            Else
                Call Isi()
                ComboBox1.Focus()
            End If
        End If
    End Sub
    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        TextBox2.MaxLength = 50
        If e.KeyChar = Chr(13) Then
            TextBox2.Text = UCase(TextBox2.Text) 'Roman Batavi 2017230073
            TextBox3.Focus()
        End If
    End Sub
    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If e.KeyChar = Chr(13) Then
            TextBox3.Text = UCase(TextBox3.Text) 'ROMAN BATAVI
            TextBox4.Focus()
        End If
    End Sub

    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        If e.KeyChar = Chr(13) Then TextBox2.Focus()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call Koneksi()
        Try
            Str = "Select * From tbljurusan Where KodeJurusan= '" & ComboBox1.Text & "'"
            cmd = New OleDbCommand(Str, Conn)
            rd = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read()
                    TextBox8.Text = rd!Jurusan.ToString.Trim()
                    TextBox5.Text = rd!SKS.ToString.Trim()
                    TextBox6.Text = rd!BPP.ToString.Trim() 'ROMAN BATAVI
                End While
            End If
            Conn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Dim total As Double
        total = Val(TextBox5.Text) + Val(TextBox6.Text) 'ROMAN BATAVI
        TextBox7.Text = total
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        With DataGridView1.Rows.Item(i)
            Me.TextBox1.Text = .Cells(0).Value
            Me.ComboBox1.Text = .Cells(1).Value
            Me.TextBox2.Text = .Cells(2).Value
            Me.TextBox3.Text = .Cells(3).Value
            Me.TextBox4.Text = .Cells(4).Value
            Me.TextBox5.Text = .Cells(5).Value
            Me.TextBox6.Text = .Cells(6).Value
            Me.TextBox7.Text = .Cells(7).Value
        End With
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or ComboBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            MsgBox("MOHON LENGKAPI DATA!") 'Roman Batavi 2017230073
            TextBox1.Focus()
            Exit Sub
        Else
            'cmd = New OleDbCommand("Select * From [tblmahasiswa2] where NoPendaftaran='" & TextBox1.Text & "'", Conn)
            'rd = cmd.ExecuteReader
            'rd.Read()
            'If Not rd.HasRows Then
            Dim Simpan As String = "insert into tblmahasiswa2(NoPendaftaran,KodeJurusan,Nama,Alamat,Agama,SKS,BPP,Total) values " & "('" & TextBox1.Text & "','" & ComboBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "')"
            cmd = New OleDbCommand(Simpan, Conn)
            Conn.Open()
            cmd.ExecuteNonQuery()
            MsgBox("SIMPAN DATA SUKSES!", MsgBoxStyle.Information, "PERHATIAN")
            'End If
            Call TampilBuku()
            Call Kosong()
            TextBox1.Focus()
            Conn.Close()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("NOMOR PENDAFTARAN BELUM DIISI")
            TextBox1.Focus()
            Exit Sub
        Else
            Str = "UPDATE `tblmahasiswa2` SET `NoPendaftaran`=@NoPendaftaran,`KodeJurusan`=@KodeJurusan,`Nama`=@Nama,`Alamat`=@Alamat,`Agama`=@Agama,`SKS`=@SKS" & _
                ",`BPP`=@BPP,`Total`=@Total WHERE `NoPendaftaran` = @NoPendaftaran;"
            Dim command As New OleDbCommand(Str, Conn)
            command.Parameters.AddWithValue("@NoPendaftaran", TextBox1.Text)
            command.Parameters.AddWithValue("@KodeJurusan", ComboBox1.Text)
            command.Parameters.AddWithValue("@Nama", TextBox2.Text)
            command.Parameters.AddWithValue("@Alamat", TextBox3.Text)
            command.Parameters.AddWithValue("@Agama", TextBox4.Text)
            command.Parameters.AddWithValue("@SKS", TextBox5.Text)
            command.Parameters.AddWithValue("@BPP", TextBox6.Text)
            command.Parameters.AddWithValue("@Total", TextBox7.Text)
            'KODE DIBAWAH INI TIDAK DIPAKAI KARENA TIDAK COMPATIBLE DENGAN DATABASE ACCESS
            'Dim Ubah As String = "Update tblmahasiswa2 set " & _
            '    "KodeJurusan='" & ComboBox1.Text & "'," & _
            '    "Nama='" & TextBox2.Text & "'," & _
            '    "Alamat='" & TextBox3.Text & "'," & _
            '    "Agama='" & TextBox4.Text & "'," & _
            '    "SKS='" & TextBox5.Text & "'," & _
            '    "BPP='" & TextBox6.Text & "'," & _
            '    "Total='" & TextBox7.Text & "' " & _
            '    "where NoPendaftaran='" & TextBox1.Text & "'"
            'cmd = New OleDbCommand(Ubah, Conn)
            Conn.Open()
            command.ExecuteNonQuery()
            MsgBox("UBAH DATA BERHASIL", MsgBoxStyle.Information, "PERHATIAN")
            Call TampilBuku()
            Call Kosong()
            TextBox1.Focus()
            Conn.Close()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("MOHON ISI NOMOR")
            TextBox1.Focus()
            Exit Sub
        Else
            If MessageBox.Show("YAKIN MAU DI HAPUS " & TextBox1.Text & " ?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Str = "DELETE FROM `tblmahasiswa2` WHERE `NoPendaftaran`=@NoPendaftaran;"
                Dim command As New OleDbCommand(Str, Conn)
                command.Parameters.AddWithValue("@NoPendaftaran", TextBox1.Text)
                'cmd = New OleDbCommand("Delete From tblmahasiswa2 where NoPendaftaran='" & TextBox1.Text & "'", Conn)
                Conn.Open()
                command.ExecuteNonQuery()
                Conn.Close()
                Call Kosong() 'Roman Batavi 2017230073
                Call TampilBuku()
            Else
                Call Kosong()
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call Kosong()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        FormMaster.Show()
        Me.Hide()
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If e.KeyChar = Chr(13) Then
            TextBox4.Text = UCase(TextBox4.Text)
            TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If e.KeyChar = Chr(13) Then TextBox6.Focus()
    End Sub

    Private Sub TextBox6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        If e.KeyChar = Chr(13) Then
            TextBox6.Text = UCase(TextBox6.Text)
            TextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox7.KeyPress
        TextBox7.MaxLength = 220
        If e.KeyChar = Chr(13) Then
            TextBox7.Text = UCase(TextBox7.Text)
            Button1.Focus()
        End If
    End Sub
End Class