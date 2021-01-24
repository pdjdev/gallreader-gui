Imports System.ComponentModel
Imports System.IO
Imports System.Net
Imports System.Runtime.InteropServices
Imports Microsoft.WindowsAPICodePack.Taskbar

Public Class Form1

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, <MarshalAs(UnmanagedType.LPWStr)> ByVal lParam As String) As Int32
    End Function

    '<DllImport("Uxtheme.dll", SetLastError:=True, CharSet:=CharSet.Auto, EntryPoint:="#95")>
    'Public Shared Function GetImmersiveColorFromColorSetEx(ByVal dwImmersiveColorSet As UInteger, ByVal dwImmersiveColorType As UInteger, ByVal bIgnoreHighContrast As Boolean, ByVal dwHighContrastCacheMode As UInteger) As UInteger
    'End Function
    '<DllImport("Uxtheme.dll", SetLastError:=True, CharSet:=CharSet.Auto, EntryPoint:="#96")>
    'Public Shared Function GetImmersiveColorTypeFromName(ByVal pName As IntPtr) As UInteger
    'End Function
    '<DllImport("Uxtheme.dll", SetLastError:=True, CharSet:=CharSet.Auto, EntryPoint:="#98")>
    'Public Shared Function GetImmersiveUserColorSetPreference(ByVal bForceCheckRegistry As Boolean, ByVal bSkipCheckOnFail As Boolean) As UInteger
    'End Function

    Public Function CheckProperColor(ByVal color As Color) As Boolean
        Dim d As Integer = 0
        Dim luminance As Double = (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) / 255

        Return luminance > 0.5
    End Function

    Private Sub DrawTitlePanel(activated As Boolean)
        If activated Then
            TitlePanel.BackColor = Color.FromArgb(77, 87, 170)
            MenuPanel.BackColor = Color.FromArgb(77, 87, 170)
        Else
            TitlePanel.BackColor = Color.FromArgb(87, 97, 178)
            MenuPanel.BackColor = Color.FromArgb(87, 97, 178)
        End If
    End Sub

    Dim CurrentProcessID As Integer
    Dim ScriptDone As Boolean = False

    Dim startnum As Integer = 0
    Dim endnum As Integer = 0

    Dim currentPrintLoc = 0
    Dim dupListUpdated = False
    Dim dupListChanged = False

    Dim WithEvents wc As New Net.WebClient

    Private Sub StartProcess(args As String)
        If Not My.Computer.FileSystem.FileExists("gallreader.exe") Then
            If MsgBox("작업을 위한 gallreader.exe 파일이 존재하지 않습니다. 다운로드 받으시겠습니까?", vbQuestion + vbYesNo) = vbYes Then
                ModeSelect(4)
                CoreDownloadStart()
                SetControls(True)
                Exit Sub
            Else
                MsgBox("작업을 진행하기 위해 gallreader.exe 파일이 필요합니다." + vbCr + vbCr _
                       + "('정보' 탭 > 'gallreader.exe 다운로드/업데이트' 버튼을 눌러 받으실 수 있습니다.)", vbExclamation)
                SetControls(True)
                Exit Sub
            End If
        End If

        StartProcess("gallreader.exe", args)
    End Sub


    Private Sub StartProcess(FileName As String, Arguments As String)

        Dim MyStartInfo As New ProcessStartInfo() With {
            .FileName = FileName,
            .Arguments = Arguments,
            .WorkingDirectory = Path.GetDirectoryName(FileName),
            .RedirectStandardError = True,
            .RedirectStandardOutput = True,
            .UseShellExecute = False,
            .CreateNoWindow = True
        }

        Dim MyProcess As Process = New Process() With {
            .StartInfo = MyStartInfo,
            .EnableRaisingEvents = True,
            .SynchronizingObject = Me
        }

        MyProcess.Start()
        MyProcess.BeginErrorReadLine()
        MyProcess.BeginOutputReadLine()

        CurrentProcessID = MyProcess.Id


        AddHandler MyProcess.OutputDataReceived,
        Sub(sender As Object, e As DataReceivedEventArgs)
            If e.Data IsNot Nothing Then
                BeginInvoke(New MethodInvoker(
                Sub()
                    ProcessLine(e.Data)
                End Sub))
            End If
        End Sub

        AddHandler MyProcess.ErrorDataReceived,
        Sub(sender As Object, e As DataReceivedEventArgs)
            If e.Data IsNot Nothing Then
                BeginInvoke(New MethodInvoker(
                Sub()
                    ProcessLine(e.Data)
                End Sub))
            End If
        End Sub

        AddHandler MyProcess.Exited,
        Sub(source As Object, ev As EventArgs)
            MyProcess.Close()
            If MyProcess IsNot Nothing Then
                MyProcess.Dispose()
            End If

            SetControls(True)
        End Sub
    End Sub

    Sub ProcessLine(line As String)
        If Not line.Contains("<smsg>") Then
            If currentPrintLoc = 0 Then
                RichTextBox1.AppendText(line + Environment.NewLine)
                RichTextBox1.ScrollToCaret()
            ElseIf currentPrintLoc = 1 Then
                RichTextBox2.AppendText(line + Environment.NewLine)
                RichTextBox2.ScrollToCaret()
            ElseIf currentPrintLoc = 2 Then
                RichTextBox3.AppendText(line + Environment.NewLine)
                RichTextBox3.ScrollToCaret()
            End If
        Else
            Dim no As Integer = Convert.ToInt32(getData(line, "no"))
            If currentPrintLoc = 0 Then
                ProgressBar2.Value = no - startnum + 1
                TitlePrevLabel1.Text = getData(line, "title")
                TaskbarManager.Instance.SetProgressValue(ProgressBar1.Value, ProgressBar1.Maximum)
            ElseIf currentPrintLoc = 1 Then
                If line.Contains("<total>") Then
                    If IsNumeric(getData(line, "total")) Then ProgressBar2.Maximum = Convert.ToInt32(getData(line, "total"))
                End If

                If Not no > ProgressBar2.Maximum Then
                    ProgressBar2.Value = no
                Else
                    ProgressBar2.Value = ProgressBar2.Maximum
                End If
                TitlePrevLabel2.Text = getData(line, "title")
                TaskbarManager.Instance.SetProgressValue(ProgressBar2.Value, ProgressBar2.Maximum)
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles StartCollectBT1.Click

        If GallIDTB1.Text = Nothing Then
            MsgBox("갤러리 ID를 입력해 주십시오", vbExclamation)

        ElseIf SaveNameTB1.Text = Nothing Then
            MsgBox("저장 파일명을 입력해 주십시오", vbExclamation)

        ElseIf PostIDNum1.Value <= 0 Then
            MsgBox("시작 글 번호를 입력해 주십시오", vbExclamation)

        ElseIf PostIDNum2.Value <= 0 Then
            MsgBox("끝 글 번호를 입력해 주십시오", vbExclamation)

        ElseIf PostIDNum1.Value > PostIDNum2.Value Then
            MsgBox("시작 번호가 끝 번호보다 클 수 없습니다", vbExclamation)

        Else

            If Not CheckDCURL(GallIDTB1.Text) Then
                If MsgBox("갤러리 ID가 올바르지 않은 것 같습니다." + vbCr + "그래도 계속하시겠습니까?", vbExclamation + vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If


            SetControls(False)

            startnum = PostIDNum1.Value
            endnum = PostIDNum2.Value

            Dim argv As String = ""

            If Not CheckBox2.Checked Then
                argv += "-r "
            End If

            argv += SaveNameTB1.Text + " " + GallIDTB1.Text + " " + PostIDNum1.Value.ToString + " " + PostIDNum2.Value.ToString + " -embed"

            ProgressBar1.Maximum = endnum - startnum + 1

            currentPrintLoc = 0
            StartProcess(argv)
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles StartCollectBT2.Click

        If PageNum1.Value = 0 And PageNum2.Value = 0 Then
            PageNum1.Value = 1
            PageNum2.Value = 1
        End If

        If GallIDTB2.Text = Nothing Then
            MsgBox("갤러리 ID를 입력해 주십시오", vbExclamation)

        ElseIf SaveNameTB2.Text = Nothing Then
            MsgBox("저장 파일명을 입력해 주십시오", vbExclamation)

        ElseIf PageNum1.Value > PageNum2.Value Then
            MsgBox("시작 페이지 번호가 끝 번호보다 클 수 없습니다", vbExclamation)

        ElseIf CheckBox1.Checked And (StartDatePicker.Value - EndDatePicker.Value).TotalSeconds < 0 Then
            MsgBox("앞 날짜 범위가 뒤 날짜 범위보다 나중에 와야 합니다. 순서를 바꿔 주세요.", vbExclamation)

        Else

            If Not CheckDCURL(GallIDTB2.Text) Then
                If MsgBox("갤러리 ID가 올바르지 않은 것 같습니다." + vbCr + "그래도 계속하시겠습니까?", vbExclamation + vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If

            SetControls(False)

            startnum = PostIDNum1.Value
            endnum = PostIDNum2.Value
            Dim argv As String = ""

            If CheckBox3.Checked Then
                argv += "-pa "
            Else
                argv += "-p "
            End If

            argv += SaveNameTB1.Text + " " + GallIDTB1.Text + " " + PageNum1.Value.ToString + " " + PageNum2.Value.ToString

            ' 선택한 시작일자에서 23시간 59분 59초를 추가
            Dim starttime = (StartDatePicker.Value - New DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds + 86400 - 1 - 32400
            ' 뒤범위는 그대로 냅두기
            Dim endtime = (EndDatePicker.Value - New DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds - 32400

            If CheckBox1.Checked Then
                argv += " " + starttime.ToString + " " + endtime.ToString
            End If

            argv += " -embed"

            currentPrintLoc = 1
            StartProcess(argv)
        End If
    End Sub

    Private Sub SetControls(enabled As Boolean)
        Panel3.Enabled = enabled
        Panel6.Enabled = enabled
        Panel10.Enabled = enabled
        Panel14.Enabled = enabled
        GroupBox1.Enabled = enabled
        MenuPanel.Enabled = enabled
    End Sub

    Private Function CheckDCURL(gallid As String)
        Dim is_valid As Boolean = True
        Dim web_response As HttpWebResponse = Nothing

        Try
            Dim web_request As HttpWebRequest = HttpWebRequest.Create("https://gall.dcinside.com/board/lists?id=" + gallid)
            web_response = DirectCast(web_request.GetResponse(), HttpWebResponse)
        Catch ex As Exception
            is_valid = False
        Finally
            If Not (web_response Is Nothing) Then web_response.Close()
        End Try

        Return is_valid
    End Function
    Private Function UrlIsValid(ByVal url As String) As Boolean
        Dim is_valid As Boolean = False
        If url.ToLower().StartsWith("www.") Then url = "http://" & url

        Dim web_response As HttpWebResponse = Nothing
        Try
            Dim web_request As HttpWebRequest =
            HttpWebRequest.Create(url)
            web_response =
            DirectCast(web_request.GetResponse(),
            HttpWebResponse)
            Return True
        Catch ex As Exception
            Return False
        Finally
            If Not (web_response Is Nothing) Then _
            web_response.Close()
        End Try
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SendMessage(GallIDTB1.Handle, &H1501, 0, "갤러리 ID 입력 (예: tree, sultan, ...)")
        SendMessage(SaveNameTB1.Handle, &H1501, 0, "저장명 입력 (예: 1분기결산)")
        SendMessage(GallIDTB2.Handle, &H1501, 0, "갤러리 ID 입력 (예: tree, sultan, ...)")
        SendMessage(SaveNameTB2.Handle, &H1501, 0, "저장명 입력 (예: 1분기결산)")
        SendMessage(ArrangeFileNameTB.Handle, &H1501, 0, "수집한 파일명 하나 이상 입력 (띄어쓰기로 파일 구분, 예: 2020-10.csv 2020-11 ...)")
        SendMessage(SaveNameTB3.Handle, &H1501, 0, "저장명 입력 (예: 1분기집계)")

        StartDatePicker.Enabled = CheckBox1.Checked
        EndDatePicker.Enabled = CheckBox1.Checked
        Label5.Enabled = CheckBox1.Checked

        StartDatePicker.Value = Today.Date
        EndDatePicker.Value = Today.Date

        RichTextBox4.Rtf = My.Resources.infortf

        ModeSelect(1)
    End Sub

    Private Sub Form1_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        DrawTitlePanel(False)
    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        DrawTitlePanel(True)
    End Sub

    Private Sub Button_Click(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button5.Click
        Select Case sender.Name
            Case Button1.Name
                ModeSelect(1)
            Case Button2.Name
                ModeSelect(2)
            Case Button3.Name
                ModeSelect(3)
            Case Button5.Name
                ModeSelect(4)
        End Select
    End Sub

    Private Sub ModeSelect(mode As Integer)
        Dim tmp = New Font(Button1.Font.Name, Button1.Font.Size, FontStyle.Bold)
        Dim tmp2 = New Font(Button1.Font.Name, Button1.Font.Size, FontStyle.Regular)

        MainPanel1.Hide()
        MainPanel2.Hide()
        MainPanel3.Hide()
        MainPanel4.Hide()

        Button1.FlatAppearance.MouseDownBackColor = Color.FromArgb(53, 60, 118)
        Button2.FlatAppearance.MouseDownBackColor = Color.FromArgb(53, 60, 118)
        Button3.FlatAppearance.MouseDownBackColor = Color.FromArgb(53, 60, 118)
        Button5.FlatAppearance.MouseDownBackColor = Color.FromArgb(53, 60, 118)

        Button1.FlatAppearance.MouseOverBackColor = Color.FromArgb(94, 106, 208)
        Button2.FlatAppearance.MouseOverBackColor = Color.FromArgb(94, 106, 208)
        Button3.FlatAppearance.MouseOverBackColor = Color.FromArgb(94, 106, 208)
        Button5.FlatAppearance.MouseOverBackColor = Color.FromArgb(94, 106, 208)

        Select Case mode
            Case 1
                Button1.BackColor = Color.White
                Button2.BackColor = Color.Transparent
                Button3.BackColor = Color.Transparent
                Button5.BackColor = Color.Transparent

                Button1.Font = tmp
                Button2.Font = tmp2
                Button3.Font = tmp2
                Button5.Font = tmp2

                Button1.ForeColor = Color.Black
                Button2.ForeColor = Color.White
                Button3.ForeColor = Color.White
                Button5.ForeColor = Color.White

                Button1.FlatAppearance.MouseDownBackColor = Color.White
                Button1.FlatAppearance.MouseOverBackColor = Color.White

                MainPanel1.Show()

            Case 2
                Button1.BackColor = Color.Transparent
                Button2.BackColor = Color.White
                Button3.BackColor = Color.Transparent
                Button5.BackColor = Color.Transparent

                Button1.Font = tmp2
                Button2.Font = tmp
                Button3.Font = tmp2
                Button5.Font = tmp2

                Button1.ForeColor = Color.White
                Button2.ForeColor = Color.Black
                Button3.ForeColor = Color.White
                Button5.ForeColor = Color.White

                Button2.FlatAppearance.MouseDownBackColor = Color.White
                Button2.FlatAppearance.MouseOverBackColor = Color.White

                MainPanel2.Show()

            Case 3

                If Not dupListUpdated Then UpdateDupList() '다중목록 업데이트

                Button1.BackColor = Color.Transparent
                Button2.BackColor = Color.Transparent
                Button3.BackColor = Color.White
                Button5.BackColor = Color.Transparent

                Button1.Font = tmp2
                Button2.Font = tmp2
                Button3.Font = tmp
                Button5.Font = tmp2

                Button1.ForeColor = Color.White
                Button2.ForeColor = Color.White
                Button3.ForeColor = Color.Black
                Button5.ForeColor = Color.White

                Button3.FlatAppearance.MouseDownBackColor = Color.White
                Button3.FlatAppearance.MouseOverBackColor = Color.White

                MainPanel3.Show()

            Case 4
                Button1.BackColor = Color.Transparent
                Button2.BackColor = Color.Transparent
                Button3.BackColor = Color.Transparent
                Button5.BackColor = Color.White

                Button1.Font = tmp2
                Button2.Font = tmp2
                Button3.Font = tmp2
                Button5.Font = tmp

                Button1.ForeColor = Color.White
                Button2.ForeColor = Color.White
                Button3.ForeColor = Color.White
                Button5.ForeColor = Color.Black

                Button5.FlatAppearance.MouseDownBackColor = Color.White
                Button5.FlatAppearance.MouseOverBackColor = Color.White

                MainPanel4.Show()

        End Select
    End Sub

    Private Sub UpdateDupList()

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()

        Try
            Dim tmp = My.Computer.FileSystem.ReadAllText("dup_list_id.txt").Replace(vbLf, "").Split(vbCr)
            For Each s In tmp
                If Not s.First = "#" Then
                    ListBox1.Items.Add(s)
                End If
            Next
        Catch ex As Exception

        End Try

        Try
            Dim tmp = My.Computer.FileSystem.ReadAllText("dup_list_ip.txt").Replace(vbLf, "").Split(vbCr).ToList
            For Each s In tmp
                If Not s.First = "#" Then
                    ListBox2.Items.Add(s)
                End If
            Next
        Catch ex As Exception

        End Try

        Try
            Dim tmp = My.Computer.FileSystem.ReadAllText("dup_list_nick.txt").Replace(vbLf, "").Split(vbCr).ToList
            For Each s In tmp
                If Not s.First = "#" Then
                    ListBox3.Items.Add(s)
                End If
            Next
        Catch ex As Exception

        End Try

        dupListUpdated = True
        dupListChanged = False

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        StartDatePicker.Enabled = CheckBox1.Checked
        EndDatePicker.Enabled = CheckBox1.Checked
        Label5.Enabled = CheckBox1.Checked
    End Sub

    Private Sub GallIDTB2_TextChanged(sender As Object, e As EventArgs) Handles GallIDTB1.TextChanged, GallIDTB2.TextChanged,
        SaveNameTB1.TextChanged, SaveNameTB2.TextChanged

        Select Case sender.Name
            Case GallIDTB1.Name
                GallIDTB2.Text = GallIDTB1.Text

            Case GallIDTB2.Name
                GallIDTB1.Text = GallIDTB2.Text

            Case SaveNameTB1.Name
                SaveNameTB2.Text = SaveNameTB1.Text

            Case SaveNameTB2.Name
                SaveNameTB1.Text = SaveNameTB2.Text
        End Select
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged, CheckBox3.CheckedChanged
        Select Case sender.Name
            Case CheckBox2.Name
                CheckBox3.Checked = CheckBox2.Checked
            Case CheckBox3.Name
                CheckBox2.Checked = CheckBox3.Checked
        End Select
    End Sub

    Private Sub UpdateToolTipTitle(sender As Object, e As EventArgs) Handles Button1.MouseEnter, Button2.MouseEnter, Button3.MouseEnter, Button5.MouseEnter
        Select Case sender.Name
            Case Button1.Name, Button2.Name, Button3.Name, Button5.Name
                ToolTip1.ToolTipTitle = sender.Text
            Case Else
                ToolTip1.ToolTipTitle = Nothing
        End Select

    End Sub

    Private Sub ListBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseClick, ListBox2.MouseClick, ListBox3.MouseClick
        Select Case sender.Name
            Case ListBox1.Name
                ListBox2.SelectedIndex = -1
                ListBox3.SelectedIndex = -1
            Case ListBox2.Name
                ListBox1.SelectedIndex = -1
                ListBox3.SelectedIndex = -1
            Case ListBox3.Name
                ListBox1.SelectedIndex = -1
                ListBox2.SelectedIndex = -1
        End Select
    End Sub

    Private Sub ListGetBT_Click(sender As Object, e As EventArgs) Handles ListGetBT.Click
        Dim tmp As String()

        If ListBox1.SelectedIndex > -1 Then
            RadioButton1.Checked = True
            tmp = ListBox1.GetItemText(ListBox1.SelectedItem).Split(vbTab)

        ElseIf ListBox2.SelectedIndex > -1 Then
            RadioButton2.Checked = True
            tmp = ListBox2.GetItemText(ListBox2.SelectedItem).Split(vbTab)

        ElseIf ListBox3.SelectedIndex > -1 Then
            RadioButton3.Checked = True
            tmp = ListBox3.GetItemText(ListBox3.SelectedItem).Split(vbTab)

        Else
            Exit Sub

        End If


        Dim count As Integer = 0
        DupListTB.Text = Nothing

        For Each i In tmp
            If count = 0 Then
                DupTitleTB.Text = tmp(count)
            ElseIf count = tmp.Length - 1 Then
                DupListTB.Text += tmp(count)
            Else
                DupListTB.Text += tmp(count) + vbCrLf
            End If
            count += 1
        Next
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles DupListRefreshBT.Click
        If dupListChanged Then
            If MsgBox("저장하지 않은 항목이 있습니다. 그래도 불러오시겠습니까?", vbQuestion + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        UpdateDupList()
    End Sub

    Private Sub ListAddBT_Click(sender As Object, e As EventArgs) Handles ListAddBT.Click
        Dim dupstr As String = DupTitleTB.Text + vbTab + DupListTB.Text.Replace(vbCrLf, vbTab)
        Dim target As ListBox

        If RadioButton1.Checked Then
            target = ListBox1
        ElseIf RadioButton2.Checked Then
            target = ListBox2
        ElseIf RadioButton3.Checked Then
            target = ListBox3
        Else
            Exit Sub
        End If

        Dim index = target.FindString(DupTitleTB.Text)

        ' 못 찾았을때 (신규)
        If index = -1 Then
            target.Items.Add(dupstr)
        Else
            If MsgBox(target.GetItemText(target.Items(index)) + vbCr + vbCr + "위 항목을 덮어씌웁니다. 계속하시겠습니까?", vbQuestion + vbYesNo) = vbYes Then
                target.Items.RemoveAt(index)
                target.Items.Insert(index, dupstr)
            Else
                Exit Sub
            End If
        End If

        dupListChanged = True

    End Sub

    Private Sub DupListSaveBT_Click(sender As Object, e As EventArgs) Handles DupListSaveBT.Click
        Try
            If ListBox1.Items.Count > 0 Then
                Dim tmp As String = "# 다중이 고닉들의 갤로그 ID, 집피를 넣으면 됩니다." _
                         + vbCrLf + "# 맨 처음 항목은 무시되며, Tab 기호로 구분합니다." _
                         + vbCrLf + "# ex) 민수	minsu1	minsu2	123.45" + vbCrLf
                Dim count As Integer = 0

                For Each i In ListBox1.Items
                    If count = ListBox1.Items.Count - 1 Then
                        tmp += ListBox1.GetItemText(i)
                    Else
                        tmp += ListBox1.GetItemText(i) + vbCrLf
                    End If
                Next

                My.Computer.FileSystem.WriteAllText("dup_list_id.txt", tmp, False, System.Text.Encoding.UTF8)
            End If

            If ListBox2.Items.Count > 0 Then
                Dim tmp As String = "# 같은 갤러로 알려진 IP 주소를 입력하면 됩니다." _
                         + vbCrLf + "# 맨 처음 항목은 무시되며, Tab 기호로 구분합니다." _
                         + vbCrLf + "# ex) 민수집피 123.45  67.89   10.11" + vbCrLf
                Dim count As Integer = 0

                For Each i In ListBox2.Items
                    If count = ListBox2.Items.Count - 1 Then
                        tmp += ListBox2.GetItemText(i)
                    Else
                        tmp += ListBox2.GetItemText(i) + vbCrLf
                    End If
                Next

                My.Computer.FileSystem.WriteAllText("dup_list_ip.txt", tmp, False, System.Text.Encoding.UTF8)
            End If

            If ListBox3.Items.Count > 0 Then
                Dim tmp As String = "# 같은 갤러로 알려진 유동 닉네임을 입력하면 됩니다." _
                         + vbCrLf + "# 맨 처음 항목은 무시되며, Tab 기호로 구분합니다." _
                         + vbCrLf + "# ex) 민수	민수	민수아님	김민수" + vbCrLf
                Dim count As Integer = 0

                For Each i In ListBox3.Items
                    If count = ListBox3.Items.Count - 1 Then
                        tmp += ListBox3.GetItemText(i)
                    Else
                        tmp += ListBox3.GetItemText(i) + vbCrLf
                    End If
                Next

                My.Computer.FileSystem.WriteAllText("dup_list_nick.txt", tmp, False, System.Text.Encoding.UTF8)
            End If
        Catch ex As Exception
            MsgBox("오류가 발생하였습니다." + vbCrLf + vbCrLf + ex.Message, vbCritical)
            Exit Sub

        End Try

        MsgBox("저장이 완료되었습니다.", vbInformation)
    End Sub

    Private Sub ListRemoveBT_Click(sender As Object, e As EventArgs) Handles ListRemoveBT.Click
        Dim target As ListBox

        If ListBox1.SelectedIndex > -1 Then
            target = ListBox1
        ElseIf ListBox2.SelectedIndex > -1 Then
            target = ListBox2
        ElseIf ListBox3.SelectedIndex > -1 Then
            target = ListBox3
        Else
            Exit Sub
        End If

        If MsgBox(target.GetItemText(target.SelectedItem) + vbCrLf + vbCrLf + "위 항목을 삭제하시겠습니까?", vbQuestion + vbYesNo) = vbYes Then
            target.Items.RemoveAt(target.SelectedIndex)
        End If
    End Sub

    Private Sub StartArrangeBT_Click(sender As Object, e As EventArgs) Handles StartArrangeBT.Click

        If ArrangeFileNameTB.Text = Nothing Then
            MsgBox("수집한 데이터 파일명을 하나 이상 입력해 주십시오", vbExclamation)

        ElseIf SaveNameTB3.Text = Nothing Then
            MsgBox("저장 파일명을 입력해 주십시오", vbExclamation)
        Else

            SetControls(False)

            currentPrintLoc = 2
            StartProcess("-a " + ArrangeFileNameTB.Text + " " + SaveNameTB3.Text)
        End If
    End Sub

    Private Sub CoreUpdateBT_Click(sender As Object, e As EventArgs) Handles CoreUpdateBT.Click
        If My.Computer.FileSystem.FileExists("gallreader.exe") Then
            If MsgBox("이미 파일이 존재합니다. 덮어 씌우고 다시 받으시겠습니까?", vbQuestion + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If

        CoreDownloadStart()
    End Sub

    Private Sub CoreDownloadStart()
        If wc.IsBusy Then
            MsgBox("이미 다운로드가 진행중입니다. 잠시 후 다시 시도해 주세요.", vbExclamation)
            Exit Sub
        Else
            Try
                My.Computer.FileSystem.DeleteFile("gallreader.exe")
            Catch ex As Exception

            End Try
        End If

        Try
            Dim tmp = webget("https://raw.githubusercontent.com/pdjdev/gallreader/main/latest.txt")
            tmp = getData(tmp, "url")

            wc.DownloadFileAsync(New Uri(tmp), "gallreader.exe")
        Catch ex As Exception
            MsgBox("다운로드에 실패하였습니다." + vbCr + "(인터넷 연결을 확인해 주시고 다시 시도해 주세요.)" + vbCr + vbCr + ex.Message, vbCritical)
            Exit Sub
        End Try

        CoreUpdateBT.Enabled = False
        SetControls(False)
    End Sub

    Private Sub wc_DownloadProgressChanged(sender As Object, e As DownloadProgressChangedEventArgs) Handles wc.DownloadProgressChanged
        CoreUpdateBT.Text = e.ProgressPercentage.ToString + "%"
    End Sub

    Private Sub wc_DownloadFileCompleted(sender As Object, e As AsyncCompletedEventArgs) Handles wc.DownloadFileCompleted
        CoreUpdateBT.Text = "gallreader.exe 다운로드/업데이트"
        MsgBox("다운로드가 완료되었습니다.", vbInformation)

        CoreUpdateBT.Enabled = True
        SetControls(True)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

    End Sub
End Class
