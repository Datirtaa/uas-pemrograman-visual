Imports System.IO

Public Class Form1
    ' Struktur data
    Structure DataSiswa
        Dim Nama As String
        Dim NilaiUTS As Double
        Dim NilaiUAS As Double
        Dim NilaiTugas As Double
        Dim NilaiAkhir As Double
        Dim Grade As String
        Dim Status As String
    End Structure

    ' Variabel global
    Dim daftarSiswa As New List(Of DataSiswa)

    ' Komponen Form
    Private WithEvents txtNama As TextBox
    Private WithEvents txtUTS As TextBox
    Private WithEvents txtUAS As TextBox
    Private WithEvents txtTugas As TextBox
    Private WithEvents btnTambah As Button
    Private WithEvents btnSimpanFile As Button
    Private WithEvents btnBacaFile As Button
    Private WithEvents btnHapus As Button
    Private WithEvents lstData As ListBox
    Private WithEvents lblStatistik As Label

    ' Constructor - Buat semua kontrol otomatis
    Public Sub New()
        InitializeComponent()
        BuatKontrol()
    End Sub

    ' Buat semua kontrol secara programmatik
    Private Sub BuatKontrol()
        Me.Text = "Aplikasi Nilai Siswa"
        Me.Size = New Size(900, 600)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Label judul
        Dim lblJudul As New Label With {
            .Text = "APLIKASI PENGELOLA DATA NILAI SISWA",
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .Location = New Point(200, 20),
            .Size = New Size(500, 30)
        }
        Me.Controls.Add(lblJudul)

        ' Label dan TextBox Nama
        Dim lblNama As New Label With {.Text = "Nama Siswa:", .Location = New Point(30, 70), .Size = New Size(100, 20)}
        txtNama = New TextBox With {.Location = New Point(140, 70), .Size = New Size(250, 25)}
        Me.Controls.Add(lblNama)
        Me.Controls.Add(txtNama)

        ' Label dan TextBox UTS
        Dim lblUTS As New Label With {.Text = "Nilai UTS:", .Location = New Point(30, 110), .Size = New Size(100, 20)}
        txtUTS = New TextBox With {.Location = New Point(140, 110), .Size = New Size(100, 25)}
        Me.Controls.Add(lblUTS)
        Me.Controls.Add(txtUTS)

        ' Label dan TextBox UAS
        Dim lblUAS As New Label With {.Text = "Nilai UAS:", .Location = New Point(30, 150), .Size = New Size(100, 20)}
        txtUAS = New TextBox With {.Location = New Point(140, 150), .Size = New Size(100, 25)}
        Me.Controls.Add(lblUAS)
        Me.Controls.Add(txtUAS)

        ' Label dan TextBox Tugas
        Dim lblTugas As New Label With {.Text = "Nilai Tugas:", .Location = New Point(30, 190), .Size = New Size(100, 20)}
        txtTugas = New TextBox With {.Location = New Point(140, 190), .Size = New Size(100, 25)}
        Me.Controls.Add(lblTugas)
        Me.Controls.Add(txtTugas)

        ' Tombol Tambah
        btnTambah = New Button With {.Text = "Tambah Data", .Location = New Point(30, 240), .Size = New Size(120, 35)}
        Me.Controls.Add(btnTambah)

        ' Tombol Simpan
        btnSimpanFile = New Button With {.Text = "Simpan ke File", .Location = New Point(160, 240), .Size = New Size(120, 35)}
        Me.Controls.Add(btnSimpanFile)

        ' Tombol Baca
        btnBacaFile = New Button With {.Text = "Baca dari File", .Location = New Point(290, 240), .Size = New Size(120, 35)}
        Me.Controls.Add(btnBacaFile)

        ' Tombol Hapus
        btnHapus = New Button With {.Text = "Hapus Semua", .Location = New Point(420, 240), .Size = New Size(120, 35)}
        Me.Controls.Add(btnHapus)

        ' ListBox
        lstData = New ListBox With {
            .Location = New Point(30, 290),
            .Size = New Size(830, 220),
            .Font = New Font("Courier New", 9)
        }
        Me.Controls.Add(lstData)

        ' Label Statistik
        lblStatistik = New Label With {
            .Location = New Point(30, 520),
            .Size = New Size(830, 25),
            .Font = New Font("Arial", 10, FontStyle.Bold)
        }
        Me.Controls.Add(lblStatistik)
    End Sub

    ' Function Hitung Nilai Akhir
    Function HitungNilaiAkhir(uts As Double, uas As Double, tugas As Double) As Double
        Return (uts * 0.3) + (uas * 0.4) + (tugas * 0.3)
    End Function

    ' Function Tentukan Grade
    Function TentukanGrade(nilaiAkhir As Double) As String
        If nilaiAkhir >= 85 Then
            Return "A"
        ElseIf nilaiAkhir >= 70 Then
            Return "B"
        ElseIf nilaiAkhir >= 60 Then
            Return "C"
        ElseIf nilaiAkhir >= 50 Then
            Return "D"
        Else
            Return "E"
        End If
    End Function

    ' Function Tentukan Status
    Function TentukanStatus(nilaiAkhir As Double) As String
        If nilaiAkhir >= 60 Then
            Return "LULUS"
        Else
            Return "TIDAK LULUS"
        End If
    End Function

    ' Sub Tampilkan Data
    Sub TampilkanData()
        lstData.Items.Clear()
        lstData.Items.Add("====================================================================")
        lstData.Items.Add(String.Format("{0,-20} {1,8} {2,8} {3,8} {4,8} {5,6} {6,12}", "NAMA", "UTS", "UAS", "TUGAS", "AKHIR", "GRADE", "STATUS"))
        lstData.Items.Add("====================================================================")

        For Each siswa As DataSiswa In daftarSiswa
            lstData.Items.Add(String.Format("{0,-20} {1,8:F2} {2,8:F2} {3,8:F2} {4,8:F2} {5,6} {6,12}", siswa.Nama, siswa.NilaiUTS, siswa.NilaiUAS, siswa.NilaiTugas, siswa.NilaiAkhir, siswa.Grade, siswa.Status))
        Next

        lstData.Items.Add("====================================================================")
        HitungStatistik()
    End Sub

    ' Sub Hitung Statistik
    Sub HitungStatistik()
        If daftarSiswa.Count = 0 Then
            lblStatistik.Text = "Belum ada data"
            Return
        End If

        Dim totalNilai As Double = 0
        Dim jumlahLulus As Integer = 0

        For Each siswa As DataSiswa In daftarSiswa
            totalNilai += siswa.NilaiAkhir
            If siswa.Status = "LULUS" Then
                jumlahLulus += 1
            End If
        Next

        Dim rataRata As Double = totalNilai / daftarSiswa.Count
        Dim persenLulus As Double = (jumlahLulus / daftarSiswa.Count) * 100

        lblStatistik.Text = String.Format("Total: {0} | Rata-rata: {1:F2} | Lulus: {2} ({3:F1}%)", daftarSiswa.Count, rataRata, jumlahLulus, persenLulus)
    End Sub

    ' Event Tombol Tambah
    Private Sub btnTambah_Click(sender As Object, e As EventArgs) Handles btnTambah.Click
        If String.IsNullOrWhiteSpace(txtNama.Text) Or String.IsNullOrWhiteSpace(txtUTS.Text) Or String.IsNullOrWhiteSpace(txtUAS.Text) Or String.IsNullOrWhiteSpace(txtTugas.Text) Then
            MessageBox.Show("Semua data harus diisi!", "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim siswa As New DataSiswa
            siswa.Nama = txtNama.Text
            siswa.NilaiUTS = Convert.ToDouble(txtUTS.Text)
            siswa.NilaiUAS = Convert.ToDouble(txtUAS.Text)
            siswa.NilaiTugas = Convert.ToDouble(txtTugas.Text)
            siswa.NilaiAkhir = HitungNilaiAkhir(siswa.NilaiUTS, siswa.NilaiUAS, siswa.NilaiTugas)
            siswa.Grade = TentukanGrade(siswa.NilaiAkhir)
            siswa.Status = TentukanStatus(siswa.NilaiAkhir)

            daftarSiswa.Add(siswa)
            TampilkanData()

            txtNama.Clear()
            txtUTS.Clear()
            txtUAS.Clear()
            txtTugas.Clear()
            txtNama.Focus()

            MessageBox.Show("Data berhasil ditambahkan!", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Input tidak valid! Pastikan nilai berupa angka.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Event Tombol Simpan File
    Private Sub btnSimpanFile_Click(sender As Object, e As EventArgs) Handles btnSimpanFile.Click
        If daftarSiswa.Count = 0 Then
            MessageBox.Show("Tidak ada data untuk disimpan!", "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim namaFile As String = "DataNilaiSiswa.txt"
            Dim writer As New StreamWriter(namaFile)

            writer.WriteLine("DATA NILAI SISWA")
            writer.WriteLine("====================================================================")

            For Each siswa As DataSiswa In daftarSiswa
                writer.WriteLine(String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}", siswa.Nama, siswa.NilaiUTS, siswa.NilaiUAS, siswa.NilaiTugas, siswa.NilaiAkhir, siswa.Grade, siswa.Status))
            Next

            writer.Close()
            MessageBox.Show("Data berhasil disimpan ke: " & namaFile, "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Gagal menyimpan: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Event Tombol Baca File
    Private Sub btnBacaFile_Click(sender As Object, e As EventArgs) Handles btnBacaFile.Click
        Dim namaFile As String = "DataNilaiSiswa.txt"

        If Not File.Exists(namaFile) Then
            MessageBox.Show("File tidak ditemukan!", "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            daftarSiswa.Clear()
            Dim reader As New StreamReader(namaFile)

            reader.ReadLine()
            reader.ReadLine()

            While Not reader.EndOfStream
                Dim baris As String = reader.ReadLine()
                Dim data() As String = baris.Split("|"c)

                If data.Length = 7 Then
                    Dim siswa As New DataSiswa
                    siswa.Nama = data(0)
                    siswa.NilaiUTS = Convert.ToDouble(data(1))
                    siswa.NilaiUAS = Convert.ToDouble(data(2))
                    siswa.NilaiTugas = Convert.ToDouble(data(3))
                    siswa.NilaiAkhir = Convert.ToDouble(data(4))
                    siswa.Grade = data(5)
                    siswa.Status = data(6)
                    daftarSiswa.Add(siswa)
                End If
            End While

            reader.Close()
            TampilkanData()
            MessageBox.Show("Data berhasil dibaca!", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Gagal membaca file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Event Tombol Hapus
    Private Sub btnHapus_Click(sender As Object, e As EventArgs) Handles btnHapus.Click
        Dim hasil = MessageBox.Show("Hapus semua data?", "Konfirmasi", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If hasil = DialogResult.Yes Then
            daftarSiswa.Clear()
            lstData.Items.Clear()
            lblStatistik.Text = ""
            MessageBox.Show("Data berhasil dihapus!", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class