Public Class Form1

    '==========存檔讀黨========

    Dim filenum As Integer
    Dim filename As String
    '==========停頓============

    Function sleep(ByVal sleep1)
        System.Threading.Thread.Sleep(sleep1)

        Return sleep1
    End Function
    '=========移點===============
    Function moveto(ByVal X_axis, ByVal Y_axis)

        Cursor.Position = New Point(X_axis, Y_axis)

        Return X_axis
        Return Y_axis
    End Function
    ' ========右單擊================
    Function rightclick(ByVal x)


        mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0) ''滑鼠左鍵按下
        mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)   ''滑鼠左鍵上按

        Return x
    End Function
    ' ========左單擊================
    Function leftclick(ByVal x)

        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) ''滑鼠左鍵按下
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)   ''滑鼠左鍵上按
        Return x
    End Function

    '========API===============
    'API 控制滑鼠

    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Integer, _
                                                  ByVal dx As Integer, ByVal dy As Integer, _
                                                  ByVal cButtons As Integer, _
                                                  ByVal dwExtraInfo As Integer)

    Private Const MOUSEEVENTF_LEFTDOWN = &H2   ''滑鼠左鍵按下
    Private Const MOUSEEVENTF_LEFTUP = &H4     ''滑鼠左鍵上按
    Private Const MOUSEEVENTF_RIGHTDOWN = &H8   ''滑鼠右鍵按下
    Private Const MOUSEEVENTF_RIGHTUP = &H10 '滑鼠右鍵上按

    '========API===============

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        '=====運行提示===


        ToolStripStatusLabel1.Text = "成功運行"
        '=迴圈控制=

        '        If CheckBox1.Checked = True Then
        'Timer2.Enabled = True

        'ElseIf ToolStripStatusLabel3.Text < TextBox6.Text Then

        'ToolStripStatusLabel3.Text += 1


        'Else
        'Timer2.Enabled = False

        'End If

        'che點選->TIME開啟->每過幾秒點一次->假如小餘預定次數就停指time

        If CheckBox1.Checked = True Then



            If ToolStripStatusLabel3.Text < Val(TextBox6.Text) - 1 Then


                Timer2.Enabled = True
                ToolStripStatusLabel3.Text += 1
                ProgressBar1.Value += 1

            Else
                Timer2.Enabled = False

                ToolStripStatusLabel3.Text = 0

                ToolStripStatusLabel1.Text = "迴圈完全執行完畢。"
                ProgressBar1.Value = 0

            End If




        End If



        '=====

        Dim ReadData() As String  '不要設陣列大小
        Dim InputText As String = TextBox1.Text

        ReadData = Split(InputText, vbCrLf)


        '======自定功能======
        Dim readstay As String
        Dim readwhere As String
        Dim readhit As String
        Dim rightclick_1 As String
        '-----抓時間----
        Dim catchtime As String
        '-----抓點
        Dim Xread As String
        Dim Yread As String
        '======迴圈控制======

        For i As Integer = 0 To ReadData.Length - 1   '超過大小會產生錯誤

            '=====自定功能延伸抓對象=======
            readwhere = Mid(ReadData(i), 1, 5)
            readstay = Mid(ReadData(i), 1, 5)
            readhit = Mid(ReadData(i), 1, 9)
            rightclick_1 = Mid(ReadData(i), 1, 10)
            '=====自定功能延伸抓對象=======
            If readwhere = "where" Then

                Xread = Mid(ReadData(i), 7, 4)

                Yread = Mid(ReadData(i), 12, 4)

                moveto(Xread, Yread)

            ElseIf readstay = "timer" Then

                catchtime = Mid(ReadData(i), 7, 8)


                sleep(catchtime)
            ElseIf readhit = "LeftClick" Then

                leftclick(1)

            ElseIf rightclick_1 = "RightClick" Then
                rightclick(1)
            End If



          


        Next
        '=========迴圈結束=====




    End Sub

    Private Sub Form1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        Select Case e.KeyCode

            Case Keys.F1
                Button6.PerformClick()
            Case Keys.F2
                Button12.PerformClick()
            Case Keys.F9

                Button1.PerformClick()
            Case Keys.F6
                Button5.PerformClick()
            Case Keys.F12
                Me.Close()
        End Select
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        '        TextBox1.Text += "LeftClick" + vbCrLf
        TextBox1.Focus()
        Clipboard.Clear()
        Clipboard.SetDataObject(vbCrLf + "LeftClick")
        '送出CTRL+V(貼上)
        SendKeys.Send("^V")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Label9.Text = TextBox3.Text.PadLeft(4, "0")
        Label10.Text = TextBox4.Text.PadLeft(4, "0")
        'TextBox1.Text += "where " & Label9.Text & "," & Label10.Text + vbCrLf


        TextBox1.Focus()
        Clipboard.Clear()
        Clipboard.SetDataObject(vbCrLf + "where " & Label9.Text & "," & Label10.Text)
        '送出CTRL+V(貼上)
        SendKeys.Send("^V")

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'TextBox1.Text += "timer " & TextBox2.Text + vbCrLf
        TextBox1.Focus()
        Clipboard.Clear()
        Clipboard.SetDataObject(vbCrLf + "timer " & TextBox2.Text)
        '送出CTRL+V(貼上)
        SendKeys.Send("^V")
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Form2.Show()
        Form3.Show()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TextBox3.Text = a1.Text
        TextBox4.Text = a2.Text
        Button3.PerformClick()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        TextBox3.Text = a3.Text
        TextBox4.Text = a4.Text
        Button3.PerformClick()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        TextBox3.Text = a5.Text
        TextBox4.Text = a6.Text
        Button3.PerformClick()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        TextBox3.Text = a7.Text
        TextBox4.Text = a8.Text
        Button3.PerformClick()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        TextBox3.Text = a9.Text
        TextBox4.Text = a10.Text
        Button3.PerformClick()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        TextBox3.Text = a11.Text
        TextBox4.Text = a12.Text
        Button3.PerformClick()
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        TextBox3.Text = a13.Text
        TextBox4.Text = a14.Text
        Button3.PerformClick()
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        TextBox3.Text = a15.Text
        TextBox4.Text = a16.Text
        Button3.PerformClick()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        TextBox3.Text = a17.Text
        TextBox4.Text = a18.Text
        Button3.PerformClick()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        TextBox3.Text = a19.Text
        TextBox4.Text = a20.Text
        Button3.PerformClick()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        leftclick(1)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Timer1.Enabled = True
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Timer1.Enabled = False
    End Sub

    Private Sub 開啟舊檔ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 開啟舊檔ToolStripMenuItem.Click
        Dim st As String
        OpenFileDialog1.Filter = "文書檔 (*.txt)|*.txt|所有檔案 (*.*)|*.*"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.RestoreDirectory = True
        OpenFileDialog1.DefaultExt = ".txt"

        '檢查否確定

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            filenum = FreeFile()
            filename = OpenFileDialog1.FileName
            '開啟檔讀取模式
            FileOpen(filenum, filename, OpenMode.Input)

            TextBox1.Text = "" '清除存放位置
            Do While Not EOF(filenum)

                Input(filenum, st)

                TextBox1.Text += st '存放位置

            Loop

            FileClose(filenum)

        End If

    End Sub

    Private Sub 另存新檔ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 另存新檔ToolStripMenuItem.Click
        SaveFileDialog1.Filter = "文書檔 (*.txt)|*.txt|所有檔案 (*.*)|*.*"
        SaveFileDialog1.FilterIndex = 1
        SaveFileDialog1.RestoreDirectory = True
        SaveFileDialog1.DefaultExt = ".txt"
        If SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then

            filename = SaveFileDialog1.FileName
            filenum = FreeFile()
            FileOpen(filenum, filename, OpenMode.Output)
            Write(filenum, TextBox1.Text) '寫入存放位置
            FileClose(filenum)

        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Process.Start("http://unromanticman.pixnet.net/blog")
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TextBox1.Text = ""



    End Sub

    Private Sub Button1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Resize

    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        'TextBox1.Text += "LeftClickDouble" + vbCrLf
        ' TextBox1.Text += "LeftClick" + vbCrLf
        'TextBox1.Text += "LeftClick" + vbCrLf


        TextBox1.Focus()
        Clipboard.Clear()
        Clipboard.SetDataObject(vbCrLf + "LeftClick")
        '送出CTRL+V(貼上)
        SendKeys.Send("^V")
        SendKeys.Send("^V")
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Process.Start("http://unromanticman.pixnet.net/blog/post/152344741")
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click

    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        '    TextBox1.Text += "RightClick" + vbCrLf

        TextBox1.Focus()
        Clipboard.Clear()
        Clipboard.SetDataObject(vbCrLf + "RightClick")
        '送出CTRL+V(貼上)
        SendKeys.Send("^V")
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        ' TextBox1.Text += "RightClick" + vbCrLf
        'TextBox1.Text += "RightClick" + vbCrLf
        TextBox1.Focus()
        Clipboard.Clear()
        Clipboard.SetDataObject(vbCrLf + "RightClick")
        '送出CTRL+V(貼上)
        SendKeys.Send("^V")
        SendKeys.Send("^V")
    End Sub

    Private Sub 圈選ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 圈選ToolStripMenuItem.Click
        TextBox1.Focus()
        TextBox1.SelectionStart = 0
        TextBox1.SelectionLength = Me.TextBox1.Text.Length
    End Sub

    Private Sub 剪下ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 剪下ToolStripMenuItem.Click
        Clipboard.Clear()
        Clipboard.SetDataObject(TextBox1.SelectedText)
        TextBox1.SelectedText = ""
    End Sub

    Private Sub 複製ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 複製ToolStripMenuItem.Click
        Clipboard.Clear()
        Clipboard.SetDataObject(TextBox1.SelectedText)
    End Sub

    Private Sub 貼上ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 貼上ToolStripMenuItem.Click
        TextBox1.Focus()
        '送出CTRL+V(貼上)
        SendKeys.Send("^V")
    End Sub

    Private Sub 刪除ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 刪除ToolStripMenuItem.Click
        TextBox1.SelectedText = ""
    End Sub

    Private Sub 抓點功能ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 抓點功能ToolStripMenuItem.Click
        Button5.PerformClick()
    End Sub

    Private Sub 復原ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 復原ToolStripMenuItem.Click
        TextBox1.Focus()

        SendKeys.Send("^Z")
    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Form2.Show()
        Form3.Show()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim st As String
        OpenFileDialog1.Filter = "文書檔 (*.txt)|*.txt|所有檔案 (*.*)|*.*"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.RestoreDirectory = True
        OpenFileDialog1.DefaultExt = ".txt"

        '檢查否確定

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            filenum = FreeFile()
            filename = OpenFileDialog1.FileName
            '開啟檔讀取模式
            FileOpen(filenum, filename, OpenMode.Input)

            TextBox1.Text = "" '清除存放位置
            Do While Not EOF(filenum)

                Input(filenum, st)

                TextBox1.Text += st '存放位置

            Loop

            FileClose(filenum)

        End If

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Button19.PerformClick()
    End Sub

    Private Sub Button19_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        TextBox1.Text = ""
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        SaveFileDialog1.Filter = "文書檔 (*.txt)|*.txt|所有檔案 (*.*)|*.*"
        SaveFileDialog1.FilterIndex = 1
        SaveFileDialog1.RestoreDirectory = True
        SaveFileDialog1.DefaultExt = ".txt"
        If SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then

            filename = SaveFileDialog1.FileName
            filenum = FreeFile()
            FileOpen(filenum, filename, OpenMode.Output)
            Write(filenum, TextBox1.Text) '寫入存放位置
            FileClose(filenum)

        End If
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click



        '=====

        Dim ReadData() As String  '不要設陣列大小
        Dim InputText As String = TextBox1.Text

        ReadData = Split(InputText, vbCrLf)


        '======自定功能======
        Dim readstay As String
        '-----抓時間----
        
        Dim stay_sum As Integer
        '======迴圈控制======

        For i As Integer = 0 To ReadData.Length - 1   '超過大小會產生錯誤

            '=====自定功能延伸抓對象=======



            readstay = Mid(ReadData(i), 1, 5)

            If readstay = "timer" Then

                stay_sum += Mid(ReadData(i), 7, 8)

            End If

            TextBox5.Text = (Val(stay_sum) / 1000) + 1


        Next
        '=========迴圈結束=====


        If Val(TextBox6.Text) <= 0 Then
            MsgBox("請先設定好再確認")

        Else
            Timer2.Interval = Val(TextBox5.Text) * 1000
            MsgBox("設定完成，共執行" & TextBox6.Text & "次")
            ProgressBar1.Maximum = TextBox6.Text
            CheckBox1.Checked = True
        End If

        

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick

        Button1.PerformClick()
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        Form4.Show()
    End Sub

    Private Sub ToolStripProgressBar1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ProgressBar1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProgressBar1.Click

    End Sub
End Class
