Imports System.Data.OleDb

Public Class Form1
    Dim baglanti As New OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0;Data Source=musteridata.accdb")


    Private Sub Temizle()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""


    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tablo As New DataTable

        If Button1.Text = "Ekle" Then

            baglanti.Open()
            Dim komut As New OleDbCommand("insert into bilgiler (Ad,Soyad,TC,Telefon,Marka,Kod,Cam,Acıklama,Tarih )values('" + TextBox1.Text + "','" + TextBox2.Text + "','" + TextBox3.Text + "','" + TextBox4.Text + "','" + TextBox5.Text + "','" + TextBox9.Text + "','" + TextBox10.Text + "','" + TextBox11.Text + "','" + TextBox12.Text + "')", baglanti)
            komut.ExecuteNonQuery()
            baglanti.Close()
            MessageBox.Show("Kaydınız Eklendi", "Kayıt")
            Temizle()
            tablo.Clear()
            listele()


        End If

        If Button1.Text = "Güncelle" Then

            baglanti.Open()
            Dim komut2 As New OleDbCommand("Update bilgiler set Ad='" + TextBox1.Text + "',Soyad='" + TextBox2.Text + "',TC='" + TextBox3.Text + "',Telefon='" + TextBox4.Text + "',Marka='" + TextBox5.Text + "',Kod='" + TextBox9.Text + "',Cam='" + TextBox10.Text + "',Acıklama='" + TextBox11.Text + "',Tarih='" + TextBox12.Text + "'where Telefon='" + TextBox4.Text + "' ", baglanti)
            komut2.ExecuteNonQuery()
            baglanti.Close()
            MessageBox.Show("Kaydınız Güncellendi", "Kayıt")
            Temizle()
            tablo.Clear()
            listele()
            Button1.Text = "Ekle"


        End If

    End Sub


    Dim tablo As New DataTable


    Private Sub listele()

        baglanti.Open()
        Dim adtr As New OleDbDataAdapter("select *from bilgiler", baglanti)
        adtr.Fill(tablo)
        DataGridView1.DataSource = tablo
        baglanti.Close()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        listele()
        DataGridView1.AllowUserToAddRows = False
        TextBox1.Visible = False
        TextBox2.Visible = False
        TextBox3.Visible = False
        TextBox4.Visible = False
        TextBox5.Visible = False
        TextBox6.Visible = False
        DataGridView1.Visible = False
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Button1.Visible = False
        Button2.Visible = False
        TextBox9.Visible = False
        TextBox10.Visible = False
        TextBox11.Visible = False
        TextBox12.Visible = False
        Label7.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label12.Visible = False









    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        baglanti.Open()
        Dim sil1 As New OleDbCommand("delete *from bilgiler where Ad='" + DataGridView1.CurrentRow.Cells("Ad").Value.ToString + "'", baglanti)
        sil1.ExecuteNonQuery()
        baglanti.Close()
        tablo.Clear()
        listele()

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        TextBox1.Text = DataGridView1.CurrentRow.Cells("Ad").Value.ToString
        TextBox2.Text = DataGridView1.CurrentRow.Cells("Soyad").Value.ToString
        TextBox3.Text = DataGridView1.CurrentRow.Cells("TC").Value.ToString
        TextBox4.Text = DataGridView1.CurrentRow.Cells("Telefon").Value.ToString
        TextBox5.Text = DataGridView1.CurrentRow.Cells("Marka").Value.ToString
        TextBox9.Text = DataGridView1.CurrentRow.Cells("Kod").Value.ToString
        TextBox10.Text = DataGridView1.CurrentRow.Cells("Cam").Value.ToString
        TextBox11.Text = DataGridView1.CurrentRow.Cells("Acıklama").Value.ToString
        TextBox12.Text = DataGridView1.CurrentRow.Cells("Tarih").Value.ToString
        Button1.Text = "Güncelle"

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged


        baglanti.Open()
        Dim adtr1 As New OleDbDataAdapter("select *from bilgiler where Ad like '" + TextBox6.Text + "%'", baglanti)
        Dim tablo2 As New DataTable
        adtr1.Fill(tablo2)

        DataGridView1.DataSource = tablo2

        baglanti.Close()








    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        TextBox1.Visible = True
            TextBox2.Visible = True
            TextBox3.Visible = True
            TextBox4.Visible = True
            TextBox5.Visible = True
            TextBox6.Visible = True
            DataGridView1.Visible = True
            Label1.Visible = True
            Label2.Visible = True
            Label3.Visible = True
            Label4.Visible = True
            Label5.Visible = True
            Label6.Visible = True
            Button1.Visible = True
            Button2.Visible = True
            TextBox9.Visible = True
            TextBox10.Visible = True
            TextBox11.Visible = True
            TextBox12.Visible = True
            Label7.Visible = True
            Label10.Visible = True
            Label11.Visible = True
            Label12.Visible = True

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Close()

    End Sub


End Class
