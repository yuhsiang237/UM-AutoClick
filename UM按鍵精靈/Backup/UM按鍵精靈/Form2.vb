Public Class Form2

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Form3.Show()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Form3.Close()
    End Sub

    Private Sub Form2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Leave


    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Form3.Show()


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        For i As Integer = 1 To 19 Step 2
            Form1.GroupBox4.Controls("a" & i).Text = Controls("D" & i).Text

        Next


        For j As Integer = 2 To 20 Step 2
            Form1.GroupBox4.Controls("a" & j).Text = Me.Controls("D" & j).Text

        Next
        'Form1.a1.Text = D1.Text
        'Form1.a2.Text = D2.Text
        ' Form1.a3.Text = D3.Text
        '  Form1.a4.Text = D4.Text
        '   Form1.a5.Text = D5.Text
        '    Form1.a6.Text = D6.Text
        '     Form1.a7.Text = D7.Text
        '      Form1.a8.Text = D8.Text
        '       Form1.a9.Text = D9.Text
        '        Form1.a10.Text = D10.Text

        Form3.Close()
        Me.Hide()



    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Label3.Text = Label3.Text - Val(ComboBox1.Text)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Label3.Text = Label3.Text + Val(ComboBox1.Text)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Label2.Text = Label3.Text - Val(ComboBox1.Text)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Label2.Text = Label2.Text + Val(ComboBox1.Text)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        D1.Text = Label2.Text
        D2.Text = Label3.Text
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        D3.Text = Label2.Text
        D4.Text = Label3.Text
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        D5.Text = Label2.Text
        D6.Text = Label3.Text
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        D7.Text = Label2.Text
        D8.Text = Label3.Text
    End Sub
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        D9.Text = Label2.Text
        D10.Text = Label3.Text
    End Sub
    Private Sub Button45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button45.Click
        D11.Text = Label2.Text
        D12.Text = Label3.Text
    End Sub

    Private Sub Button44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button44.Click
        D13.Text = Label2.Text
        D14.Text = Label3.Text
    End Sub

    Private Sub Button43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button43.Click
        D15.Text = Label2.Text
        D16.Text = Label3.Text
    End Sub

    Private Sub Button42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button42.Click
        D17.Text = Label2.Text
        D18.Text = Label3.Text
    End Sub

    Private Sub Button41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button41.Click
        D19.Text = Label2.Text
        D20.Text = Label3.Text
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Form3.Location = New Point(Label2.Text, Label3.Text)
    End Sub
End Class