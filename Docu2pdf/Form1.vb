'
'*************************************************************************************************************
'
'           耐火試験室のデータベース作成プログラム MakeData ver 1.0
'
'               2020/10    Coded by kanyama
'
'           (概要）
'
'           NASのフォルダーから報告書ファイル（.xdw .xbd）またはスキャン資料ファイル（.pdf）を検索し、
'           
'           ファイルメーカーデータベースに登録する。
'
'           検索可能な報告書ファイル名の例：3A180001.xdw(.xbd) , 3A18001.xdw , ⅢA180001.xdw , A180001.xdw , 3A-18-0001.xdw , 3A-18-001.xdw , 3A-18-01.xdw
'
'           検索可能な資料ファイル名の例：3A180001_依頼者名_試験名称_1（依頼書） , 3A180001_依頼者名_資料名_2（その他の資料）
'           （試験番号は報告書ファイルに準ずる）
'
'*************************************************************************************************************

Imports System.IO
Imports System.Net
Imports System.Security.AccessControl
Imports FujiXerox.DocuWorks.Toolkit
Imports System.Runtime.InteropServices


Public Class Form1
    Private filename() As String, fname() As String, dir1() As String
    Private DcuPath As String, PdfPath As String
    Private Check() As CheckBox, checkbox_n As Integer
    Private Cansel As Boolean
    Private MyPath As String, MyName As String, username As String, adrList As IPAddress(), MyIP As String, hostname As String
    Private OcrFlag As Boolean
    Private EndFlag As Boolean


    <Flags()>
    Public Enum PlaySoundFlags
        SND_SYNC = &H0
        SND_ASYNC = &H1
        SND_NODEFAULT = &H2
        SND_MEMORY = &H4
        SND_LOOP = &H8
        SND_NOSTOP = &H10
        SND_NOWAIT = &H2000
        SND_ALIAS = &H10000
        SND_ALIAS_ID = &H110000
        SND_FILENAME = &H20000
        SND_RESOURCE = &H40004
        SND_PURGE = &H40
        SND_APPLICATION = &H80
    End Enum

    <System.Runtime.InteropServices.DllImport("winmm.dll",
    CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
    Private Shared Function PlaySound(ByVal pszSound As String,
    ByVal hmod As IntPtr, ByVal fdwSound As PlaySoundFlags) As Boolean
    End Function


    Private TextFileName As String

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If EndFlag = False Then
            If MessageBox.Show("終了しますか？", "終了確認ダイアログ", MessageBoxButtons.YesNo) = DialogResult.No Then
                e.Cancel = True
            Else
                If OcrFlag = True Then
                    Xdwapi.XDW_Finalize()
                End If
            End If
        Else
            e.Cancel = False
        End If
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '
        ' フォームの初期化
        '

        EndFlag = False
        '実行アプリケーションのプロセス名を取得
        Dim strThisProcess As String = System.Diagnostics.Process.GetCurrentProcess().ProcessName

        '取得した同名のプロセスが他に存在するかを確認
        If System.Diagnostics.Process.GetProcessesByName(strThisProcess).Length > 1 Then
            MsgBox("すでに起動中です。", vbOK, "確認")
            EndFlag = True
            Me.Close()
            Me.Dispose()
            AppActivate("MakeData.exe")
            End
        End If

        ' ファイルメーカーサーバーのIPアドレス情報を読み込む
        Dim text As String = ""
        Try
            Dim sr As New StreamReader("c:\ファイル情報設定ファイル\HostIP.txt", System.Text.Encoding.GetEncoding("utf-8"))
            text = sr.ReadLine
            sr.Close()
        Catch
        End Try

        Dim rx = New System.Text.RegularExpressions.Regex("\d+.\d+.\d+.\d+", System.Text.RegularExpressions.RegexOptions.Compiled)
        Dim result As Boolean = rx.IsMatch(text)
        If result = True Then
            TextBox_FileMakerServer.Text = text
        Else
            TextBox_FileMakerServer.Text = FileMakerServer1
        End If


        Me.Icon = My.Resources.auezb_d3bmk_002


        MyPath = My.Application.Info.DirectoryPath
        MyName = My.Application.Info.AssemblyName
        ' ホスト名を取得する
        hostname = Dns.GetHostName()
        username = Environment.UserName
        ' ホスト名からIPアドレスを取得する
        adrList = Dns.GetHostAddresses(hostname)
        MyIP = ""
        For Each address As IPAddress In adrList
            Console.WriteLine(address.ToString())
            If System.Text.RegularExpressions.Regex.IsMatch(address.ToString(), "\d{1,3}(\.\d{1,3}){3}(/\d{1,2})?") = True Then
                Console.WriteLine(address.ToString())
                MyIP = address.ToString()
            End If
        Next

        If MyIP <> "" Then
            If PCInfo(MyIP, username, MyPath, MyName) = False Then
                MsgBox("データベースに接続出来ません!" + vbCrLf + "終了します!", vbOK, "警告")
                End
            End If
        End If
        ' 起動時にプログレスバーを非表示にする。
        Me.ProgressBar1.Visible = False
        Me.ProgressBar2.Visible = False

        ' 分野番号チェックボックスの初期設定

        checkbox_n = 19
        ReDim Check(checkbox_n - 1)
        Check(0) = CheckBox_3A
        Check(1) = CheckBox_3C
        Check(2) = CheckBox_3G
        Check(3) = CheckBox_3J
        Check(4) = CheckBox_3K
        Check(5) = CheckBox_3L
        Check(6) = CheckBox_3M
        Check(7) = CheckBox_3N
        Check(8) = CheckBox_3O
        Check(9) = CheckBox_3P
        Check(10) = CheckBox_3S
        Check(11) = CheckBox_3T
        Check(12) = CheckBox_3U
        Check(13) = CheckBox_3X
        Check(14) = CheckBox_3Y
        Check(15) = CheckBox_3Z
        Check(16) = CheckBox_8A
        Check(17) = CheckBox_8B
        Check(18) = CheckBox_8C

        For i As Integer = 0 To checkbox_n - 1
            Check(i).Checked = True
        Next
        CheckBox_ALL.Checked = True

        CheckBox_kozo.Checked = True
        CheckBox_zairyou.Checked = True
        CheckBox_tobihi.Checked = True
        CheckBox_hyouteika.Checked = True

        'TextBox_FileMakerServer.Text = FileMakerServer1
        Cansel = False

        'TextBox_FolderName1.Text = "\\192.168.0.173\disk1\報告書（耐火）＿業務課から\2000Ⅲ耐火防火試験室"
        'TextBox_FolderName1.Text = "\\192.168.0.173\disk1\報告書（耐火）＿業務課から\test用"
        PdfPath = TextBox_FolderName1.Text

        CheckBox_Input.Checked = True       ' 報告書の入力チェックボックス
        CheckBox_Convert.Checked = True     ' 報告書の変換チェックボックス
        CheckBox_Input2.Checked = True      ' 資料（スキャンデータ）の入力チェックボックス

        OcrFlag = False

        'RadioButton_xdw.Checked = True
        'RadioButton_pdf.Checked = False

        xdwFolderRadioButton.Checked = True
        xdwFileRadioButton.Checked = False
        pdfFolderRadioButton.Checked = True
        pdfFileRadioButton.Checked = False

        成績書OnlyCheckBox.Checked = True
        FolderNameCheckBox1.Checked = True
        スキャンデータOnlyCheckBox.Checked = True
        FolderNameCheckBox2.Checked = True

        Me.Width = 1020
        Me.Height = 720
        Me.CenterToScreen()                 ' Formをモニターの中央に表示

    End Sub



    Private Sub CheckBox_kozo_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_kozo.CheckedChanged
        ' 構造試験のチェックボックスの変更
        Dim b As Boolean = CheckBox_kozo.Checked
        CheckBox_3A.Checked = b
        CheckBox_3J.Checked = b
        CheckBox_3M.Checked = b
        CheckBox_3N.Checked = b
        CheckBox_3S.Checked = b
        CheckBox_3X.Checked = b
    End Sub

    Private Sub CheckBox_zairyou_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_zairyou.CheckedChanged
        ' 材料試験のチェックボックスの変更
        Dim b As Boolean = CheckBox_zairyou.Checked
        CheckBox_3C.Checked = b
        CheckBox_3K.Checked = b
        CheckBox_3O.Checked = b
        CheckBox_3T.Checked = b
        CheckBox_3Y.Checked = b

    End Sub


    Private Sub CheckBox_tobihi_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_tobihi.CheckedChanged
        ' 飛び火試験のチェックボックスの変更
        Dim b As Boolean = CheckBox_tobihi.Checked
        CheckBox_3G.Checked = b
        CheckBox_3L.Checked = b
        CheckBox_3P.Checked = b
        CheckBox_3U.Checked = b
        CheckBox_3Z.Checked = b

    End Sub

    Private Sub CheckBox_ALL_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_ALL.CheckedChanged
        ' すべての試験のチェックボックスの変更
        Dim b As Boolean = CheckBox_ALL.Checked
        For i As Integer = 0 To checkbox_n - 1
            Check(i).Checked = b
        Next
        CheckBox_kozo.Checked = b
        CheckBox_zairyou.Checked = b
        CheckBox_tobihi.Checked = b
        CheckBox_hyouteika.Checked = b
    End Sub

    Private Sub CheckBox_hyouteika_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_hyouteika.CheckedChanged
        ' 性能評価のチェックボックスの変更
        Dim b As Boolean = CheckBox_hyouteika.Checked
        CheckBox_8A.Checked = b
        CheckBox_8B.Checked = b
        CheckBox_8C.Checked = b
    End Sub



    Private Function CountChar(ByVal s As String, ByVal c As Char) As Integer
        ' 文字列 s の中の文字 c の出現回数をカウントする関数
        Return s.Length - s.Replace(c.ToString(), "").Length
    End Function


    Private Sub DocuReadButton_Click(sender As Object, e As EventArgs) Handles DocuReadButton.Click
        '
        '　報告書（xdw,xbd）を読み込んでPDFに変換し、それをFileMakerに登録する。
        '
        If TextBox_FolderName1.Text <> "" Then
            Try
                Dim f1 As New Form2
                f1.title = "ファイルの検索"
                f1.message = "ただいまファイルの検索中！" + vbCrLf + vbCrLf + vbCrLf + "しばらくお待ちください。"
                f1.Show()

                Application.DoEvents()


                Dim fname2 As New List(Of String)
                Dim dir2 As New List(Of String)
                Dim WildCard1() As String
                'Dim Count As Integer = 0
                Dim ff()() As String    ', flag() As Boolean
                Dim ff2() As String, ff3() As String
                Dim ROnly As Boolean = 成績書OnlyCheckBox.Checked

                ReDim WildCard1(1), ff(1)

                WildCard1(0) = "*.xdw"
                WildCard1(1) = "*.xbd"

                Dim nn As Integer = 0

                For i As Integer = 0 To WildCard1.Length - 1
                    ff2 = System.IO.Directory.GetFiles(DcuPath, WildCard1(i), System.IO.SearchOption.AllDirectories)
                    If ROnly Then
                        ReDim ff3(ff2.Length - 1)
                        Dim k As Integer = 0
                        For j As Integer = 0 To ff2.Length - 1

                            If ff2(j).Contains("\成績書\") Then
                                ff3(k) = ff2(j)
                                k += 1
                            End If

                        Next
                        ReDim Preserve ff3(k - 1)
                        ff(i) = ff3
                    Else
                        ff(i) = ff2
                    End If
                    nn += ff(i).Length
                    Application.DoEvents()
                Next

                ReDim filename(nn - 1)

                For i As Integer = 0 To WildCard1.Length - 1
                    If i = 0 Then
                        ff(i).CopyTo(filename, 0)
                    Else
                        ff(i).CopyTo(filename, ff(i - 1).Length)
                    End If
                Next

                Dim Count As Integer = MakeXdwList(filename)

                'fname = fname2.ToArray
                'dir1 = dir2.ToArray

                f1.Close()
                f1.Dispose()

                If Count = 0 Then
                    MsgBox("このフォルダーには報告書ファイルはありません！", vbOK, "確認")
                Else
                    MsgBox("このフォルダーには" + Count.ToString + "個の報告書ファイルがありました。", vbOK, "確認")
                End If

            Catch e1 As Exception
                'Console.WriteLine(e1.Message)
            End Try

        Else
            MsgBox("フォルダーを選択してください！", vbOK, "エラー")

        End If


    End Sub


    Private Function MakeXdwList(ByVal filename As String()) As Integer

        ' ファイルパスをリスト表示する


        Dim Style1 As New DataGridViewCellStyle()
        Style1.BackColor = Color.White
        Style1.Font = New Font("MSゴシック", 9, FontStyle.Regular)
        Style1.Alignment = DataGridViewContentAlignment.MiddleLeft

        Dim Style2 As New DataGridViewCellStyle()
        Style2.BackColor = Color.White
        Style2.Font = New Font("MSゴシック", 9, FontStyle.Regular)
        Style2.Alignment = DataGridViewContentAlignment.MiddleCenter

        Dim Style3 As New DataGridViewCellStyle()
        Style3.BackColor = Color.White
        Style3.Font = New Font("MSゴシック", 9, FontStyle.Regular)
        Style3.Alignment = DataGridViewContentAlignment.MiddleRight

        With Me.DataGridView1
            .Rows.Clear()
            .Columns.Clear()
            .Width = 920
            .Height = 256
            .ColumnCount = 3
            'Col_n = .ColumnCount
            .ColumnHeadersVisible = True
            .ColumnHeadersHeight = 18
            .ScrollBars = ScrollBars.Both


            .ColumnHeadersDefaultCellStyle = Style1
            .Columns(0).Name = "番号"
            .Columns(1).Name = "ファイル名"
            .Columns(2).Name = "Path"

            .RowHeadersVisible = True
            .Columns(0).Width = 60
            .Columns(1).Width = 140
            .Columns(2).Width = 500

            Dim textColumn1 As New DataGridViewTextBoxColumn()
            textColumn1.DataPropertyName = "入力"
            textColumn1.Name = "入力"
            textColumn1.HeaderText = "入力"
            .Columns.Add(textColumn1)

            .Columns(3).Width = 40
            .Columns(3).DefaultCellStyle = Style2

            Dim column1_2 As New DataGridViewCheckBoxColumn
            .Columns.Add(column1_2)
            .Columns(4).Name = "☑️"
            .Columns(4).Width = 25

            Dim textColumn2 As New DataGridViewTextBoxColumn()
            textColumn2.DataPropertyName = "変換"
            textColumn2.Name = "変換"
            textColumn2.HeaderText = "変換"
            .Columns.Add(textColumn2)
            .Columns(5).Width = 40
            .Columns(5).DefaultCellStyle = Style2

            Dim column2_2 As New DataGridViewCheckBoxColumn
            .Columns.Add(column2_2)
            .Columns(6).Name = "☑️"
            .Columns(6).Width = 25

            Dim textColumn3 As New DataGridViewTextBoxColumn()
            textColumn3.DataPropertyName = "読込"
            textColumn3.Name = "読込"
            textColumn3.HeaderText = "読込"
            .Columns.Add(textColumn3)
            .Columns(7).Width = 40
            .Columns(7).DefaultCellStyle = Style2


        End With

        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String

        FileMakerServer = TextBox_FileMakerServer.Text
        db.Connect()

        Dim Count As Integer = 0
        Dim flag() As Boolean

        Dim n As Integer = filename.Length
        ReDim flag(n - 1)
        For i As Integer = 0 To n - 1
            flag(i) = False
        Next

        DataGridView1.Rows.Clear()
        Style1.BackColor = Color.White
        Style1.Font = New Font("MSゴシック", 9, FontStyle.Regular)
        DataGridView1.RowsDefaultCellStyle = Style1

        For j As Integer = 0 To checkbox_n - 1
            If Check(j).Checked = True Then

                If n > 0 Then

                    For i As Integer = 0 To n - 1
                        If flag(i) = False Then

                            For k As Integer = 0 To Ndrive.GetLength(0) - 1
                                filename(i) = filename(i).Replace(Ndrive(k, 0), Ndrive(k, 1))
                            Next
                            'filename(i) = filename(i).Replace("W:\", "\\192.168.37.242\fire\")
                            'filename(i) = filename(i).Replace("X:\", "\\192.168.37.240\fire\")
                            'filename(i) = filename(i).Replace("V:\", "\\192.168.37.242\fire\")
                            'filename(i) = filename(i).Replace("Y:\", "\\192.168.0.173\fire\")

                            Dim f As String = System.IO.Path.GetFileNameWithoutExtension(filename(i))

                            Dim fname As String = System.IO.Path.GetFileName(filename(i))


                            If IsTestNumber(fname, Check(j).Text) Then   ' 試験番号（例えば、3A 3C）を含むファイルかどうかをチェック

                                Count += 1
                                Dim row1() As String
                                row1 = {Count.ToString, fname, filename(i)}
                                DataGridView1.Rows.Add(row1)

                                Sql_Command = "Select ""FilePath"",""PdfPath"",""入力"" FROM """ + Table + """ WHERE (""ファイル名"" = '" & fname.Replace("'", "''") & "')"
                                tb = db.ExecuteSql(Sql_Command)
                                Dim n2 As Integer = tb.Rows.Count
                                Dim st1 As String
                                If n2 > 0 Then
                                    DataGridView1.Rows(Count - 1).Cells(3).Value = "済"
                                    DataGridView1.Rows(Count - 1).Cells(4).Value = False
                                    st1 = tb.Rows(0).Item("PdfPath").ToString()
                                    If st1 <> "" Then
                                        DataGridView1.Rows(Count - 1).Cells(5).Value = "済"
                                        DataGridView1.Rows(Count - 1).Cells(6).Value = False
                                    Else
                                        DataGridView1.Rows(Count - 1).Cells(5).Value = "未"
                                        DataGridView1.Rows(Count - 1).Cells(6).Value = True
                                    End If
                                    st1 = tb.Rows(0).Item("入力").ToString()
                                    If st1 <> "未読" Then
                                        DataGridView1.Rows(Count - 1).Cells(7).Value = "済"
                                    Else
                                        DataGridView1.Rows(Count - 1).Cells(7).Value = "未"
                                    End If
                                Else
                                    DataGridView1.Rows(Count - 1).Cells(3).Value = "未"
                                    DataGridView1.Rows(Count - 1).Cells(4).Value = True
                                    DataGridView1.Rows(Count - 1).Cells(5).Value = "未"
                                    DataGridView1.Rows(Count - 1).Cells(6).Value = True
                                    DataGridView1.Rows(Count - 1).Cells(7).Value = "未"
                                    st1 = ""
                                End If

                                DataGridView1.FirstDisplayedScrollingRowIndex = Count - 1
                                DataGridView1.CurrentCell = DataGridView1(0, Count - 1)

                                flag(i) = True



                            End If
                        End If
                        'Application.DoEvents()

                    Next

                End If
            End If

            Application.DoEvents()

        Next


        db.Disconnect()



        MakeXdwList = Count




    End Function


    Private Sub PdfReadButton_Click(sender As Object, e As EventArgs) Handles PdfReadButton.Click
        '
        '　資料（pdf）を読み込んでFileMakerに登録する。
        '
        If TextBox_FolderName2.Text <> "" Then
            Try
                Dim f1 As New Form2
                f1.title = "ファイルの検索"
                f1.message = "ただいまファイルの検索中！" + vbCrLf + vbCrLf + vbCrLf + "しばらくお待ちください。"
                f1.Show()
                Application.DoEvents()

                Dim fname2 As New List(Of String)
                Dim dir2 As New List(Of String)
                Dim WildCard1() As String

                'Dim Count As Integer = 0
                Dim ff()() As String
                Dim ff2() As String, ff3() As String
                Dim ROnly As Boolean = スキャンデータOnlyCheckBox.Checked

                ReDim WildCard1(0), ff(0)
                WildCard1(0) = "*.pdf"

                Dim nn As Integer = 0

                For i As Integer = 0 To WildCard1.Length - 1
                    ff2 = System.IO.Directory.GetFiles(PdfPath, WildCard1(i), System.IO.SearchOption.AllDirectories)
                    If ROnly Then
                        ReDim ff3(ff2.Length - 1)
                        Dim k As Integer = 0
                        For j As Integer = 0 To ff2.Length - 1

                            If ff2(j).Contains("\スキャンデータ\") Then
                                ff3(k) = ff2(j)
                                k += 1
                            End If

                        Next
                        ReDim Preserve ff3(k - 1)
                        ff(i) = ff3
                    Else
                        ff(i) = ff2
                    End If
                    nn += ff(i).Length
                    Application.DoEvents()
                Next

                ReDim filename(nn - 1)

                For i As Integer = 0 To WildCard1.Length - 1
                    If i = 0 Then
                        ff(i).CopyTo(filename, 0)
                    Else
                        ff(i).CopyTo(filename, ff(i - 1).Length)
                    End If
                Next

                Dim Count As Integer = MakePdfList(filename)



                'Dim Style1 As New DataGridViewCellStyle()
                'Style1.BackColor = Color.White
                'Style1.Font = New Font("MSゴシック", 9, FontStyle.Regular)
                'Style1.Alignment = DataGridViewContentAlignment.MiddleLeft

                'Dim Style2 As New DataGridViewCellStyle()
                'Style2.BackColor = Color.White
                'Style2.Font = New Font("MSゴシック", 9, FontStyle.Regular)
                'Style2.Alignment = DataGridViewContentAlignment.MiddleCenter

                'Dim Style3 As New DataGridViewCellStyle()
                'Style3.BackColor = Color.White
                'Style3.Font = New Font("MSゴシック", 9, FontStyle.Regular)
                'Style3.Alignment = DataGridViewContentAlignment.MiddleRight
                'Dim columnHeaderStyle As New DataGridViewCellStyle()
                'With Me.DataGridView2
                '    .Rows.Clear()
                '    .Columns.Clear()
                '    .Width = 900
                '    .Height = 300
                '    .ColumnCount = 3
                '    .ColumnHeadersVisible = True
                '    .ColumnHeadersHeight = 18
                '    .ScrollBars = ScrollBars.Both

                '    .ColumnHeadersDefaultCellStyle = Style1
                '    .Columns(0).Name = "番号"
                '    .Columns(1).Name = "ファイル名"
                '    .Columns(2).Name = "Path"

                '    .RowHeadersVisible = True
                '    .Columns(0).Width = 60
                '    .Columns(1).Width = 240
                '    .Columns(2).Width = 440

                '    Dim textColumn1 As New DataGridViewTextBoxColumn()
                '    textColumn1.DataPropertyName = "入力"
                '    textColumn1.Name = "入力"
                '    textColumn1.HeaderText = "入力"
                '    .Columns.Add(textColumn1)
                '    .Columns(3).Width = 40
                '    .Columns(3).DefaultCellStyle = Style2

                '    Dim column1_2 As New DataGridViewCheckBoxColumn
                '    .Columns.Add(column1_2)
                '    .Columns(4).Name = "☑️"
                '    .Columns(4).Width = 25

                '    Dim textColumn3 As New DataGridViewTextBoxColumn()
                '    textColumn3.DataPropertyName = "読込"
                '    textColumn3.Name = "読込"
                '    textColumn3.HeaderText = "読込"
                '    .Columns.Add(textColumn3)
                '    .Columns(5).Width = 40
                '    .Columns(5).DefaultCellStyle = Style2

                'End With

                'Dim db As New OdbcDbIf
                'Dim tb As DataTable
                'Dim Sql_Command As String

                'FileMakerServer = TextBox_FileMakerServer.Text
                'db.Connect()

                ''Dim fname2 As New List(Of String)
                ''Dim dir2 As New List(Of String)
                ''Dim WildCard1() As String

                'Dim Count As Integer = 0
                'Dim flag() As Boolean

                ''ReDim WildCard1(0), ff(0)
                ''WildCard1(0) = "*.pdf"

                ''Dim nn As Integer = 0

                ''For i As Integer = 0 To WildCard1.Length - 1
                ''    ff(i) = System.IO.Directory.GetFiles(PdfPath, WildCard1(i), System.IO.SearchOption.AllDirectories)
                ''    nn += ff(i).Length
                ''Next

                ''ReDim filename(nn - 1)

                ''For i As Integer = 0 To WildCard1.Length - 1
                ''    If i = 0 Then
                ''        ff(i).CopyTo(filename, 0)
                ''    Else
                ''        ff(i).CopyTo(filename, ff(i - 1).Length)
                ''    End If
                ''Next

                'Dim n As Integer = filename.Length
                'ReDim flag(n - 1)
                'For i As Integer = 0 To n - 1
                '    flag(i) = False
                'Next

                'DataGridView2.Rows.Clear()

                'DataGridView2.RowsDefaultCellStyle = Style1

                'For j As Integer = 0 To checkbox_n - 1
                '    If Check(j).Checked = True Then

                '        If n > 0 Then

                '            For i As Integer = 0 To n - 1
                '                If flag(i) = False Then
                '                    Dim f As String = System.IO.Path.GetFileNameWithoutExtension(filename(i))

                '                    Dim fname As String = System.IO.Path.GetFileName(filename(i))


                '                    If IsTestNumber(fname, Check(j).Text) Then

                '                        Count += 1
                '                        Dim row1() As String
                '                        row1 = {Count.ToString, fname, filename(i)}
                '                        DataGridView2.Rows.Add(row1)

                '                        Sql_Command = "Select ""FilePath"", ""FilePath"", ""入力"" FROM """ + Table + """ WHERE (""ファイル名"" = '" & fname.Replace("'", "''") & "')"
                '                        tb = db.ExecuteSql(Sql_Command)
                '                        Dim n2 As Integer = tb.Rows.Count
                '                        Dim st1 As String
                '                        If n2 > 0 Then
                '                            DataGridView2.Rows(Count - 1).Cells(3).Value = "済"
                '                            DataGridView2.Rows(Count - 1).Cells(4).Value = False

                '                            st1 = tb.Rows(0).Item("入力").ToString()
                '                            If st1 <> "未読" Then
                '                                DataGridView2.Rows(Count - 1).Cells(5).Value = "済"
                '                            Else
                '                                DataGridView2.Rows(Count - 1).Cells(5).Value = "未"
                '                            End If
                '                        Else
                '                            DataGridView2.Rows(Count - 1).Cells(3).Value = "未"
                '                            DataGridView2.Rows(Count - 1).Cells(4).Value = True
                '                            DataGridView2.Rows(Count - 1).Cells(5).Value = "未"
                '                            st1 = ""
                '                        End If

                '                        DataGridView2.FirstDisplayedScrollingRowIndex = Count - 1
                '                        DataGridView2.CurrentCell = DataGridView2(0, Count - 1)

                '                        flag(i) = True



                '                    End If
                '                End If
                '                'Application.DoEvents()

                '            Next

                '        End If
                '    End If

                '    Application.DoEvents()

                'Next


                'db.Disconnect()

                'fname = fname2.ToArray
                'dir1 = dir2.ToArray

                f1.Close()
                f1.Dispose()

                If Count = 0 Then
                    MsgBox("このフォルダーには資料ファイルはありません！", vbOK, "確認")
                Else
                    MsgBox("このフォルダーには" + Count.ToString + "個の資料ファイルがありました。", vbOK, "確認")
                End If


            Catch e1 As Exception
                'Console.WriteLine(e1.Message)
            End Try
        Else
            MsgBox("フォルダーを選択してください！", vbOK, "エラー")
        End If

    End Sub

    Private Function MakePdfList(ByVal filename As String()) As Integer




        Dim Style1 As New DataGridViewCellStyle()
        Style1.BackColor = Color.White
        Style1.Font = New Font("MSゴシック", 9, FontStyle.Regular)
        Style1.Alignment = DataGridViewContentAlignment.MiddleLeft

        Dim Style2 As New DataGridViewCellStyle()
        Style2.BackColor = Color.White
        Style2.Font = New Font("MSゴシック", 9, FontStyle.Regular)
        Style2.Alignment = DataGridViewContentAlignment.MiddleCenter

        Dim Style3 As New DataGridViewCellStyle()
        Style3.BackColor = Color.White
        Style3.Font = New Font("MSゴシック", 9, FontStyle.Regular)
        Style3.Alignment = DataGridViewContentAlignment.MiddleRight
        Dim columnHeaderStyle As New DataGridViewCellStyle()
        With Me.DataGridView2
            .Rows.Clear()
            .Columns.Clear()
            .Width = 900
            .Height = 262
            .ColumnCount = 3
            .ColumnHeadersVisible = True
            .ColumnHeadersHeight = 18
            .ScrollBars = ScrollBars.Both

            .ColumnHeadersDefaultCellStyle = Style1
            .Columns(0).Name = "番号"
            .Columns(1).Name = "ファイル名"
            .Columns(2).Name = "Path"

            .RowHeadersVisible = True
            .Columns(0).Width = 60
            .Columns(1).Width = 240
            .Columns(2).Width = 440

            Dim textColumn1 As New DataGridViewTextBoxColumn()
            textColumn1.DataPropertyName = "入力"
            textColumn1.Name = "入力"
            textColumn1.HeaderText = "入力"
            .Columns.Add(textColumn1)
            .Columns(3).Width = 40
            .Columns(3).DefaultCellStyle = Style2

            Dim column1_2 As New DataGridViewCheckBoxColumn
            .Columns.Add(column1_2)
            .Columns(4).Name = "☑️"
            .Columns(4).Width = 25

            Dim textColumn3 As New DataGridViewTextBoxColumn()
            textColumn3.DataPropertyName = "読込"
            textColumn3.Name = "読込"
            textColumn3.HeaderText = "読込"
            .Columns.Add(textColumn3)
            .Columns(5).Width = 40
            .Columns(5).DefaultCellStyle = Style2

        End With

        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String

        FileMakerServer = TextBox_FileMakerServer.Text
        db.Connect()

        'Dim fname2 As New List(Of String)
        'Dim dir2 As New List(Of String)
        'Dim WildCard1() As String

        Dim Count As Integer = 0
        Dim flag() As Boolean

        'ReDim WildCard1(0), ff(0)
        'WildCard1(0) = "*.pdf"

        'Dim nn As Integer = 0

        'For i As Integer = 0 To WildCard1.Length - 1
        '    ff(i) = System.IO.Directory.GetFiles(PdfPath, WildCard1(i), System.IO.SearchOption.AllDirectories)
        '    nn += ff(i).Length
        'Next

        'ReDim filename(nn - 1)

        'For i As Integer = 0 To WildCard1.Length - 1
        '    If i = 0 Then
        '        ff(i).CopyTo(filename, 0)
        '    Else
        '        ff(i).CopyTo(filename, ff(i - 1).Length)
        '    End If
        'Next

        Dim n As Integer = filename.Length
        ReDim flag(n - 1)
        For i As Integer = 0 To n - 1
            flag(i) = False
        Next

        DataGridView2.Rows.Clear()

        DataGridView2.RowsDefaultCellStyle = Style1

        For j As Integer = 0 To checkbox_n - 1
            If Check(j).Checked = True Then

                If n > 0 Then

                    For i As Integer = 0 To n - 1
                        If flag(i) = False Then

                            For k As Integer = 0 To Ndrive.GetLength(0) - 1
                                filename(i) = filename(i).Replace(Ndrive(k, 0), Ndrive(k, 1))
                            Next

                            Dim f As String = System.IO.Path.GetFileNameWithoutExtension(filename(i))

                            Dim fname As String = System.IO.Path.GetFileName(filename(i))


                            If IsTestNumber(fname, Check(j).Text) Then

                                Count += 1
                                Dim row1() As String
                                row1 = {Count.ToString, fname, filename(i)}
                                DataGridView2.Rows.Add(row1)

                                Sql_Command = "Select ""FilePath"", ""FilePath"", ""入力"" FROM """ + Table + """ WHERE (""ファイル名"" = '" & fname.Replace("'", "''") & "')"
                                tb = db.ExecuteSql(Sql_Command)
                                Dim n2 As Integer = tb.Rows.Count
                                Dim st1 As String
                                If n2 > 0 Then
                                    DataGridView2.Rows(Count - 1).Cells(3).Value = "済"
                                    DataGridView2.Rows(Count - 1).Cells(4).Value = False

                                    st1 = tb.Rows(0).Item("入力").ToString()
                                    If st1 <> "未読" Then
                                        DataGridView2.Rows(Count - 1).Cells(5).Value = "済"
                                    Else
                                        DataGridView2.Rows(Count - 1).Cells(5).Value = "未"
                                    End If
                                Else
                                    DataGridView2.Rows(Count - 1).Cells(3).Value = "未"
                                    DataGridView2.Rows(Count - 1).Cells(4).Value = True
                                    DataGridView2.Rows(Count - 1).Cells(5).Value = "未"
                                    st1 = ""
                                End If

                                DataGridView2.FirstDisplayedScrollingRowIndex = Count - 1
                                DataGridView2.CurrentCell = DataGridView2(0, Count - 1)

                                flag(i) = True



                            End If
                        End If
                        'Application.DoEvents()

                    Next

                End If
            End If

            Application.DoEvents()

        Next


        db.Disconnect()

        MakePdfList = Count

    End Function


    Private Function IsTestNumber(ByVal fname As String, ByVal checkChr As String) As Boolean
        '
        ' ファイル名に試験番号含まれるかどうかをチャックする関数
        '
        '  A010001,A-01-001,A-01-01,A-01-0001
        '
        '  小文字、全角にも対応


        Dim s(3) As String, w(4) As String

        s(0) = checkChr.Substring(1, 1)                 ' 3A ->A
        s(1) = s(0).ToLower                             ' 3A -> A -> a
        s(2) = StrConv(s(0), VbStrConv.Wide)            ' A -> Ａ（全角）
        s(3) = StrConv(s(1), VbStrConv.Wide)            ' a -> ａ（全角）
        w(0) = "\d\d\d\d\d\d"
        w(1) = "\d\d\d\d\d"
        w(2) = "-\d\d-\d\d\d"
        w(3) = "-\d\d-\d\d"
        w(4) = "-\d\d-\d\d\d\d"

        IsTestNumber = False
        If fname.Substring(0, 1) <> "." Then    ' 隠しファイルを除外する。
            For i As Integer = 0 To s.Length - 1
                For j As Integer = 0 To w.Length - 1
                    IsTestNumber = System.Text.RegularExpressions.Regex.IsMatch(fname, s(i) + w(j)) Or IsTestNumber
                Next
            Next
        End If
    End Function


    Private Sub CheckBox_Input_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_Input.CheckedChanged
        '
        '   すべての入力チャックボックスのON/OFF切替
        '
        Dim n As Integer = DataGridView1.RowCount - 1
        If n > 0 Then
            For i As Integer = 0 To n - 1
                If CheckBox_Input.Checked = True Then
                    If DataGridView1.Rows(i).Cells(3).Value = "未" Then
                        DataGridView1.Rows(i).Cells(4).Value = True
                    End If
                Else
                    If DataGridView1.Rows(i).Cells(3).Value = "未" Then
                        DataGridView1.Rows(i).Cells(4).Value = False
                    End If
                End If

            Next
        End If
    End Sub

    Private Sub CheckBox_Convert_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_Convert.CheckedChanged
        '
        '   すべての変換チャックボックスのON/OFF切替
        '
        Dim n As Integer = DataGridView1.RowCount - 1
        If n > 0 Then
            For i As Integer = 0 To n - 1
                If CheckBox_Convert.Checked = True Then
                    If DataGridView1.Rows(i).Cells(5).Value = "未" Then
                        DataGridView1.Rows(i).Cells(6).Value = True
                    End If
                Else
                    If DataGridView1.Rows(i).Cells(5).Value = "未" Then
                        DataGridView1.Rows(i).Cells(6).Value = False
                    End If
                End If

            Next
        End If
    End Sub



    Private Sub Select_Read_Folder_Button2_Click(sender As Object, e As EventArgs) Handles Select_Read_Folder_Button2.Click
        '
        '   資料（PDF）を検索するフォルダーの選択
        '
        Dim fbd As New FolderBrowserDialog

        '上部に表示する説明テキストを指定する
        fbd.Description = "読み込むフォルダを指定してください。"
        'ルートフォルダを指定する
        'デフォルトでDesktop
        fbd.RootFolder = Environment.SpecialFolder.Desktop
        '最初に選択するフォルダを指定する
        'RootFolder以下にあるフォルダである必要がある
        If TextBox_FolderName2.Text <> "" Then
            fbd.SelectedPath = TextBox_FolderName2.Text
        Else
            fbd.SelectedPath = "\\192.168.0.173\disk1\SCAN\test"
        End If

        'ユーザーが新しいフォルダを作成できるようにする
        'デフォルトでTrue
        fbd.ShowNewFolderButton = True

        'ダイアログを表示する
        If fbd.ShowDialog(Me) = DialogResult.OK Then
            '選択されたフォルダを表示する
            Me.TextBox_FolderName2.Text = fbd.SelectedPath
            PdfPath = fbd.SelectedPath

            'Path2 = PdfSaveFolder + "\" + System.IO.Path.GetFileName(System.IO.Path.GetFileName(Path1))
            'TextBox_FilderName2.Text = Path2
        End If
    End Sub


    Private Function DocuToPdf(ByVal file1 As String, ByVal file2 As String, ByVal Dpi As Integer) As Integer
        '
        '   ドキュワークス(.xdw/.xbd)をPDFに変換する関数
        '
        '   (引数)

        '   file1：ドキュワークスファイルのPath
        '   file2：PDFファイルのPath
        '   Dpi  ：PDFの解像度（DPI）(600DPI以下）
        '
        '   (戻り値)

        '   DocuToPdf：0の場合、変換成功、0以外の場合、変換失敗


        Dim Handle As Xdwapi.XDW_DOCUMENT_HANDLE = New Xdwapi.XDW_DOCUMENT_HANDLE()
        Dim mode As Xdwapi.XDW_OPEN_MODE_EX = New Xdwapi.XDW_OPEN_MODE_EX()
        With mode
            .Option = Xdwapi.XDW_OPEN_READONLY
            .AuthMode = Xdwapi.XDW_AUTH_NODIALOGUE
        End With

        Dim api_result As Integer = Xdwapi.XDW_OpenDocumentHandle(file1, Handle, mode)

        Dim info As Xdwapi.XDW_DOCUMENT_INFO = New Xdwapi.XDW_DOCUMENT_INFO()
        Xdwapi.XDW_GetDocumentInformation(Handle, info)
        Dim end_page As Integer = info.Pages
        Dim start_page As Integer = 1

        Dim pdf1 As Xdwapi.XDW_IMAGE_OPTION_PDF = New Xdwapi.XDW_IMAGE_OPTION_PDF()

        With pdf1
            .Compress = Xdwapi.XDW_COMPRESS_MRC_NORMAL
            .ConvertMethod = Xdwapi.XDW_CONVERT_MRC_OS
            .EndOfMultiPages = end_page
        End With

        Dim Dpi1 As Integer = Dpi
        Dim Color1 As Integer = Xdwapi.XDW_IMAGE_COLOR
        Dim ImageType1 As Integer = Xdwapi.XDW_IMAGE_PDF
        Dim ex1 As Xdwapi.XDW_IMAGE_OPTION_EX = New Xdwapi.XDW_IMAGE_OPTION_EX()
        With ex1
            .Dpi = Dpi1
            .Color = Color1
            .ImageType = ImageType1
            .DetailOption = pdf1
        End With
        Dim api_result2 As Integer = Xdwapi.XDW_ConvertPageToImageFile(Handle, start_page, file2, ex1)

        'Me.TextBox1.Text = api_result2.ToString

        Xdwapi.XDW_CloseDocumentHandle(Handle)

        DocuToPdf = api_result2
    End Function

    Private Function DocuToText(ByVal file1 As String, ByVal page As Integer) As String

        Dim FolderPath As String = Path.GetDirectoryName(file1)
        Dim fileInfo As New FileInfo(FolderPath)
        Dim fileSec As FileSecurity = fileInfo.GetAccessControl()


        ' アクセス権限をEveryoneに対しフルコントロール許可
        Try
            Dim accessRule As New FileSystemAccessRule("Everyone", FileSystemRights.FullControl, AccessControlType.Allow)
            fileSec.AddAccessRule(accessRule)
            fileInfo.SetAccessControl(fileSec)
        Catch ex As Exception

        End Try

        Dim Handle As Xdwapi.XDW_DOCUMENT_HANDLE = New Xdwapi.XDW_DOCUMENT_HANDLE()
        Dim mode As Xdwapi.XDW_OPEN_MODE_EX = New Xdwapi.XDW_OPEN_MODE_EX()
        With mode
            .Option = Xdwapi.XDW_OPEN_READONLY
            .AuthMode = Xdwapi.XDW_AUTH_NODIALOGUE
        End With

        Dim api_result As Integer = Xdwapi.XDW_OpenDocumentHandle(file1, Handle, mode)
        DocuToText = ""

        If api_result >= 0 Then
            Dim info As Xdwapi.XDW_DOCUMENT_INFO = New Xdwapi.XDW_DOCUMENT_INFO()
            Xdwapi.XDW_GetDocumentInformation(Handle, info)
            Dim end_page As Integer = info.Pages
            Dim start_page As Integer = 1


            If page >= 1 And page <= end_page Then

                Try
                    Dim info2 As Xdwapi.XDW_PAGE_INFO_EX = New Xdwapi.XDW_PAGE_INFO_EX()

                    Dim result As Integer = Xdwapi.XDW_GetPageInformation(Handle, page, info2)

                    If result >= 0 Then
                        If info2.PageType = Xdwapi.XDW_PGT_FROMAPPL Or info2.PageType = Xdwapi.XDW_PGT_FROMIMAGE Then
                            Dim text1 As String
                            Dim nDataSize As Integer = 0
                            Dim reserved = Nothing
                            Dim result2 As Integer = Xdwapi.XDW_GetPageTextToMemory(Handle, page, text1)

                            Xdwapi.XDW_CloseDocumentHandle(Handle)  ' ファイルを閉じる

                            If result2 >= 0 Then
                                If text1 <> Nothing Then
                                    ' テキストが正しく読めた場合
                                    DocuToText = "0" + text1

                                Else
                                    ' テキストが読めなかった場合はOCR処理をしてテキストを抽出する。
                                    With mode
                                        .Option = Xdwapi.XDW_OPEN_UPDATE    ' 編集モード
                                        .AuthMode = Xdwapi.XDW_AUTH_NODIALOGUE
                                    End With

                                    api_result = Xdwapi.XDW_OpenDocumentHandle(file1, Handle, mode)     ' 再度ファイルを開く

                                    'Dim result3 As Integer = Xdwapi.XDW_RotatePageAuto(Handle, page)    ' 横書きの場合は90度回
                                    Dim result3 As Integer = 0
                                    If result3 >= 0 Then
                                        Dim ocr_optoin As Xdwapi.XDW_OCR_OPTION_V7 = New Xdwapi.XDW_OCR_OPTION_V7
                                        With ocr_optoin
                                            .NoiseReduction = Xdwapi.XDW_REDUCENOISE_NORMAL
                                            .Language = Xdwapi.XDW_OCR_LANGUAGE_AUTO
                                            .InsertSpaceCharacter = 0
                                            .Form = Xdwapi.XDW_OCR_FORM_AUTO
                                            .Column = Xdwapi.XDW_OCR_COLUMN_AUTO
                                            .EngineLevel = Xdwapi.XDW_OCR_ENGINE_LEVEL_ACCURACY
                                        End With
                                        result3 = Xdwapi.XDW_ApplyOcr(Handle, page, Xdwapi.XDW_OCR_ENGINE_DEFAULT, ocr_optoin)
                                        System.Threading.Thread.Sleep(1000)
                                        OcrFlag = True
                                        If result3 >= 0 Then

                                            result3 = -1
                                            result3 = Xdwapi.XDW_GetPageTextToMemory(Handle, page, text1)
                                        End If

                                        If result3 >= 0 Then
                                            If text1 <> Nothing Then
                                                DocuToText = "1" + text1
                                            End If
                                        End If
                                    End If
                                    Xdwapi.XDW_CloseDocumentHandle(Handle)
                                    'Xdwapi.XDW_Finalize()

                                End If

                            End If

                        End If

                    End If
                Catch ex As Exception

                End Try


            End If



        End If

        'Xdwapi.XDW_CloseDocumentHandle(Handle)

        'If ocr_exec <> 0 Then
        '    Xdwapi.XDW_Finalize()
        'End If

    End Function

    Private Sub Select_Read_Folder_Button_Click(sender As Object, e As EventArgs) Handles Select_Read_Folder_Button.Click
        '
        '   報告書（.xdw/.xbd）を検索するフォルダーの選択
        '
        Dim fbd As New FolderBrowserDialog

        '上部に表示する説明テキストを指定する
        fbd.Description = "読み込むフォルダを指定してください。"
        'ルートフォルダを指定する
        'デフォルトでDesktop
        fbd.RootFolder = Environment.SpecialFolder.Desktop
        '最初に選択するフォルダを指定する
        'RootFolder以下にあるフォルダである必要がある
        If TextBox_FolderName1.Text <> "" Then
            fbd.SelectedPath = TextBox_FolderName1.Text
        Else
            'fbd.SelectedPath = "\\192.168.0.173\disk1\報告書（耐火）＿業務課から"
            'fbd.SelectedPath = "\\192.168.37.242\fire\耐火構造\依頼試験\案件フォルダ【取扱注意】元データのため削除禁止\2015年度"
            fbd.SelectedPath = "W:\耐火構造\依頼試験\案件フォルダ【取扱注意】元データのため削除禁止\2015年度"
        End If

        'ユーザーが新しいフォルダを作成できるようにする
        'デフォルトでTrue
        fbd.ShowNewFolderButton = True

        'ダイアログを表示する
        If fbd.ShowDialog(Me) = DialogResult.OK Then
            '選択されたフォルダを表示する
            Me.TextBox_FolderName1.Text = fbd.SelectedPath
            DcuPath = fbd.SelectedPath

            PdfPath = PdfSaveFolder + "\" + System.IO.Path.GetFileName(System.IO.Path.GetFileName(DcuPath))
            'TextBox_FilderName2.Text = Path2
        End If

    End Sub



    Private Sub Data_Input_Button_Click(sender As Object, e As EventArgs) Handles Data_Input_Button.Click
        '
        '   報告書のデータをデータベースに転送するサブルーチン
        '
        '
        If DataGridView1.RowCount > 1 Then

            Dim db As New OdbcDbIf
            Dim tb As DataTable
            Dim Sql_Command As String
            Dim ROnly = 成績書OnlyCheckBox.Checked And FolderNameCheckBox1.Checked

            Dim n As Integer = DataGridView1.RowCount - 1
            Dim Count As Integer = 0
            For i As Integer = 0 To n - 1
                If DataGridView1.Rows(i).Cells(4).Value = True Then
                    Count += 1
                End If
            Next

            If Count > 0 Then

                ' メッセージボックスに表示するテキスト
                Dim message As String = "未入力のデータが" + Count.ToString + "個あります。" + vbCrLf + "入力しますか？"

                ' タイトルバーに表示するテキスト
                Dim caption As String = "確認"

                ' 表示するボタン([OK]ボタンと[キャンセル]ボタン)
                Dim buttons As MessageBoxButtons = MessageBoxButtons.OKCancel

                Dim result As DialogResult = MessageBox.Show(message, caption, buttons)

                If result = vbOK Then

                    Dim f1 As New Form2
                    f1.title = "PDF変換"
                    f1.message = "ただいまデータをに入力中！" + vbCrLf + vbCrLf + vbCrLf + "しばらくお待ちください。"
                    f1.Show()
                    Application.DoEvents()

                    TextBox_FileLIst2.Text = ""
                    ProgressBar1.Minimum = 0
                    ProgressBar1.Maximum = n
                    ProgressBar1.Visible = True

                    Debug.Print("OK")
                    Try
                        FileMakerServer = TextBox_FileMakerServer.Text
                        db.Connect()

                        If System.IO.File.Exists(CmdFile) = False Then
                            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(CmdFile))
                        End If
                        Dim sw As New StreamWriter(CmdFile, False, System.Text.Encoding.GetEncoding("utf-8"))
                        sw.WriteLine("1")  ' コマンド番号

                        For i As Integer = 0 To n - 1
                            If DataGridView1.Rows(i).Cells(4).Value = True Then

                                Dim fname As String = DataGridView1.Rows(i).Cells(1).Value
                                Dim filename1 As String = DataGridView1.Rows(i).Cells(2).Value
                                Dim FolderPath As String = Path.GetDirectoryName(filename1)

                                Dim Path1 As String, name1 As String
                                Dim text1 As String, text2 As String
                                Dim testname As String
                                Dim username As String

                                text1 = DocuToText(filename1, 1)
                                text1 = text1.Replace(" ", "").Replace("　", "")
                                text2 = DocuToText(filename1, 2)
                                text2 = text2.Replace(" ", "").Replace("　", "")

                                If ROnly Then
                                    Path1 = FolderPath.Replace("\成績書", "")
                                    name1 = Path1.Substring(Path1.LastIndexOf("\") + 1, Path1.Length - Path1.LastIndexOf("\") - 1)
                                    testname = name1
                                    username = name1
                                Else

                                    'text1 = DocuToText(filename1, 1)
                                    'text1 = text1.Replace(" ", "").Replace("　", "")
                                    'text2 = DocuToText(filename1, 2)
                                    'text2 = text2.Replace(" ", "").Replace("　", "")

                                    testname = DataFromText(text1, "試験項目")
                                    'Dim testno As String = DataFromText(text1, "試験番号")
                                    username = DataFromText(text2, "依頼者名")
                                    If username = "" Then
                                        username = DataFromText(text1, "依頼者名")
                                    End If
                                End If

                                Sql_Command = "INSERT INTO """ + Table + """ (""FilePath"",""ファイル名"",""入力"")"
                                Sql_Command += " VALUES ('" + filename1.Replace("'", "''") + "','" + fname.Replace("'", "''") + "','未読')"
                                tb = db.ExecuteSql(Sql_Command)

                                Sql_Command = "UPDATE """ + Table + """ SET ""試験項目2"" = '" + testname + "',""依頼者名2"" = '" + username + "'"
                                Sql_Command += "  WHERE ""ファイル名"" = '" + fname.Replace("'", "''") + "'"
                                tb = db.ExecuteSql(Sql_Command)

                                text1 = text1.Replace("'", "''").Replace(vbCrLf, "")
                                text2 = text2.Replace("'", "''").Replace(vbCrLf, "")

                                Sql_Command = "UPDATE """ + Table + """ SET ""page1"" = '" + text1 + "',""page2"" = '" + text2 + "'"
                                Sql_Command += "  WHERE ""ファイル名"" = '" + fname.Replace("'", "''") + "'"
                                tb = db.ExecuteSql(Sql_Command)

                                Dim filename2 As String = TestNo(fname)
                                If filename2 <> "" Then
                                    Dim kind2 As String = filename2.Substring(0, 2)
                                    Dim year2 As String = filename2.Substring(2, 2)
                                    Dim no2 As String = filename2.Substring(4, 4)
                                    Dim eda2 As String = ""
                                    If filename2.Contains("(") = True Then
                                        eda2 = filename2.Substring(9, 2)
                                    End If
                                    Sql_Command = "UPDATE """ + Table + """ SET ""分類2"" = '" + kind2 + "',""年度2"" = '" + year2 + "',""番号2"" = '" + no2
                                    Sql_Command += "',""枝番2"" = '" + eda2 + "',""試験番号2"" = '" + filename2 + "'"
                                    Sql_Command += "  WHERE ""ファイル名"" = '" + fname.Replace("'", "''") + "'"
                                    tb = db.ExecuteSql(Sql_Command)
                                End If

                                DataGridView1.Rows(i).Cells(3).Value = "済"
                                DataGridView1.Rows(i).Cells(4).Value = False

                                sw.WriteLine(fname)
                                ProgressBar1.Value = i + 1

                                Me.TextBox_FileLIst2.Text += filename1 + vbCrLf
                                Me.TextBox_FileLIst2.SelectionStart = Me.TextBox_FileLIst2.Text.Length
                                Me.TextBox_FileLIst2.Focus()
                                Me.TextBox_FileLIst2.ScrollToCaret()

                                'Count += 1
                            End If
                            Application.DoEvents()
                        Next
                        db.Disconnect()

                        ' Shift-Jisでファイルを作成


                        '２行書き込み


                        'ストリームを閉じる
                        sw.Close()
                        ProgressBar1.Visible = False
                    Catch e1 As Exception
                        f1.Close()
                        f1.Dispose()
                        ProgressBar1.Visible = False
                    End Try

                    f1.Close()
                    f1.Dispose()
                    ProgressBar1.Visible = False
                End If
            Else
                MessageBox.Show("未入力のデータはありません", "警告", MessageBoxButtons.OK)
            End If



        Else
            MessageBox.Show("データがありません!!", "警告", MessageBoxButtons.OK)

        End If
    End Sub

    Private Function TestNo(ByVal FileName As String) As String
        TestNo = ""
        Dim Pattern As String() = {"[3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}\d\d\d\d\d\d\(\d+\)",
                                     "[3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}\d\d\d\d\d\(\d+\)",
                                     "[3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}\d\d\d\d\d\d",
                                     "[3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}\d\d\d\d\d",
                                     "[3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}-\d\d-\d\d\d\d\(\d+\)",
                                     "[3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}-\d\d-\d\d\d\(\d+\)",
                                     "[3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}-\d\d-\d\d\d\d-\d+",
                                     "[3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]]{1}-\d\d-\d\d\d-\d+",
                                     "[3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}-\d\d-\d\d\d\d",
                                     "[3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}-\d\d-\d\d\d"
            }
        Dim Pattern_n = Pattern.Length
        Dim y As Integer
        Dim n1 As Integer, n2 As Integer
        Dim kind As String = ""
        'Imports System.Text.RegularExpressions
        For i As Integer = 0 To Pattern_n - 1
            Dim rx = New System.Text.RegularExpressions.Regex(Pattern(i), System.Text.RegularExpressions.RegexOptions.Compiled)
            Dim result As Boolean = rx.IsMatch(FileName)
            If result = True Then
                Dim r As New System.Text.RegularExpressions.Regex(Pattern(i), System.Text.RegularExpressions.RegexOptions.None)
                'TextBox1.Text内で正規表現と一致する対象をすべて検索 
                Dim mc As System.Text.RegularExpressions.MatchCollection = r.Matches(FileName)

                Dim result2 As String = ""
                For Each m As System.Text.RegularExpressions.Match In mc
                    '正規表現に一致したグループの文字列を表示 
                    result2 = m.Groups(0).Value

                    Exit For
                Next
                If result2 <> "" Then

                    result2 = StrConv(result2, VbStrConv.Narrow).Replace("Ⅲ", "3").Replace("Ⅷ", "8")
                    Select Case i

                        Case 0  ' [3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}\d\d\d\d\d\d\(\d+\)
                            If result2.Substring(0, 1) = "3" Or result2.Substring(0, 1) = "8" Then
                                y = Integer.Parse(result2.Substring(2, 2))
                                n1 = Integer.Parse(result2.Substring(4, 4))
                                n2 = Integer.Parse(result2.Substring(result2.IndexOf("(") + 1, result2.IndexOf(")") - result2.IndexOf("(") - 1))
                                kind = result2.Substring(0, 2)
                            Else
                                y = Integer.Parse(result2.Substring(1, 2))
                                n1 = Integer.Parse(result2.Substring(3, 4))
                                n2 = Integer.Parse(result2.Substring(result2.IndexOf("(") + 1, result2.IndexOf(")") - result2.IndexOf("(") - 1))
                                kind = "3" + result2.Substring(0, 1)
                            End If

                        Case 1  ' [3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}\d\d\d\d\d\(\d+\)
                            If result2.Substring(0, 1) = "3" Or result2.Substring(0, 1) = "8" Then
                                y = Integer.Parse(result2.Substring(2, 2))
                                n1 = Integer.Parse(result2.Substring(4, 3))
                                n2 = Integer.Parse(result2.Substring(result2.IndexOf("(") + 1, result2.IndexOf(")") - result2.IndexOf("(") - 1))
                                kind = result2.Substring(0, 2)
                            Else
                                y = Integer.Parse(result2.Substring(1, 2))
                                n1 = Integer.Parse(result2.Substring(3, 3))
                                n2 = Integer.Parse(result2.Substring(result2.IndexOf("(") + 1, result2.IndexOf(")") - result2.IndexOf("(") - 1))
                                kind = "3" + result2.Substring(0, 1)
                            End If

                        Case 2  ' [3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}\d\d\d\d\d\d
                            If result2.Substring(0, 1) = "3" Or result2.Substring(0, 1) = "8" Then
                                y = Integer.Parse(result2.Substring(2, 2))
                                n1 = Integer.Parse(result2.Substring(4, 4))
                                n2 = 0
                                kind = result2.Substring(0, 2)
                            Else
                                y = Integer.Parse(result2.Substring(1, 2))
                                n1 = Integer.Parse(result2.Substring(3, 4))
                                n2 = 0
                                kind = "3" + result2.Substring(0, 1)
                            End If

                        Case 3  ' [3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}\d\d\d\d\d
                            If result2.Substring(0, 1) = "3" Or result2.Substring(0, 1) = "8" Then
                                y = Integer.Parse(result2.Substring(2, 2))
                                n1 = Integer.Parse(result2.Substring(4, 3))
                                n2 = 0
                                kind = result2.Substring(0, 2)
                            Else
                                y = Integer.Parse(result2.Substring(1, 2))
                                n1 = Integer.Parse(result2.Substring(3, 3))
                                n2 = 0
                                kind = "3" + result2.Substring(0, 1)
                            End If

                        Case 4  ' [3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}-\d\d-\d\d\d\d\(\d+\)
                            If result2.Substring(0, 1) = "3" Or result2.Substring(0, 1) = "8" Then
                                y = Integer.Parse(result2.Substring(3, 2))
                                n1 = Integer.Parse(result2.Substring(6, 4))
                                n2 = Integer.Parse(result2.Substring(result2.IndexOf("(") + 1, result2.IndexOf(")") - result2.IndexOf("(") - 1))
                                kind = result2.Substring(0, 2)
                            Else
                                y = Integer.Parse(result2.Substring(2, 2))
                                n1 = Integer.Parse(result2.Substring(5, 4))
                                n2 = Integer.Parse(result2.Substring(result2.IndexOf("(") + 1, result2.IndexOf(")") - result2.IndexOf("(") - 1))
                                kind = "3" + result2.Substring(0, 1)
                            End If

                        Case 5  ' [3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}-\d\d-\d\d\d\(\d+\)
                            If result2.Substring(0, 1) = "3" Or result2.Substring(0, 1) = "8" Then
                                y = Integer.Parse(result2.Substring(3, 2))
                                n1 = Integer.Parse(result2.Substring(6, 3))
                                n2 = Integer.Parse(result2.Substring(result2.IndexOf("(") + 1, result2.IndexOf(")") - result2.IndexOf("(") - 1))
                                kind = result2.Substring(0, 2)
                            Else
                                y = Integer.Parse(result2.Substring(2, 2))
                                n1 = Integer.Parse(result2.Substring(5, 3))
                                n2 = Integer.Parse(result2.Substring(result2.IndexOf("(") + 1, result2.IndexOf(")") - result2.IndexOf("(") - 1))
                                kind = "3" + result2.Substring(0, 1)
                            End If


                        Case 6  ' [3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}-\d\d-\d\d\d\d-\d+
                            If result2.Substring(0, 1) = "3" Or result2.Substring(0, 1) = "8" Then
                                y = Integer.Parse(result2.Substring(3, 2))
                                n1 = Integer.Parse(result2.Substring(6, 4))
                                n2 = Integer.Parse(result2.Substring(result2.IndexOf("-", 7) + 1, result2.Length - result2.IndexOf("-", 7) - 1))
                                kind = result2.Substring(0, 2)
                            Else
                                y = Integer.Parse(result2.Substring(2, 2))
                                n1 = Integer.Parse(result2.Substring(5, 4))
                                n2 = Integer.Parse(result2.Substring(result2.IndexOf("-", 6) + 1, result2.Length - result2.IndexOf("-", 6) - 1))
                                kind = "3" + result2.Substring(0, 1)
                            End If

                        Case 7  ' [3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}-\d\d-\d\d\d-\d+"
                            If result2.Substring(0, 1) = "3" Or result2.Substring(0, 1) = "8" Then
                                y = Integer.Parse(result2.Substring(3, 2))
                                n1 = Integer.Parse(result2.Substring(6, 3))
                                n2 = Integer.Parse(result2.Substring(result2.IndexOf("-", 6) + 1, result2.Length - result2.IndexOf("-", 6) - 1))
                                kind = result2.Substring(0, 2)
                            Else
                                y = Integer.Parse(result2.Substring(2, 2))
                                n1 = Integer.Parse(result2.Substring(5, 3))
                                n2 = Integer.Parse(result2.Substring(result2.IndexOf("-", 5) + 1, result2.Length - result2.IndexOf("-", 5) - 1))
                                kind = "3" + result2.Substring(0, 1)
                            End If

                        Case 8  ' [3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}-\d\d-\d\d\d\d\
                            If result2.Substring(0, 1) = "3" Or result2.Substring(0, 1) = "8" Then
                                y = Integer.Parse(result2.Substring(3, 2))
                                n1 = Integer.Parse(result2.Substring(6, 4))
                                n2 = 0
                                kind = result2.Substring(0, 2)
                            Else
                                y = Integer.Parse(result2.Substring(2, 2))
                                n1 = Integer.Parse(result2.Substring(5, 4))
                                n2 = 0
                                kind = "3" + result2.Substring(0, 1)
                            End If

                        Case 9  ' [3３8８ⅢⅧ]?[a-zA-Zａ-ｚＡ-Ｚ]{1}-\d\d-\d\d\d\
                            If result2.Substring(0, 1) = "3" Or result2.Substring(0, 1) = "8" Then
                                y = Integer.Parse(result2.Substring(3, 2))
                                n1 = Integer.Parse(result2.Substring(6, 3))
                                n2 = 0
                                kind = result2.Substring(0, 2)
                            Else
                                y = Integer.Parse(result2.Substring(2, 2))
                                n1 = Integer.Parse(result2.Substring(5, 3))
                                n2 = 0
                                kind = "3" + result2.Substring(0, 1)
                            End If


                    End Select
                    If kind <> "" Then Exit For
                End If
            End If
        Next

        If kind <> "" Then
            TestNo = kind + y.ToString("00") + n1.ToString("0000")
            If n2 > 0 Then
                TestNo += "(" + n2.ToString("00") + ")"
            End If
        End If

    End Function



    Private Function DataFromText(ByVal text As String, ByVal kind As String) As String

        DataFromText = ""

        If text <> "" Then
            Select Case kind
                Case "試験項目"
                    Dim pattern As String = "[\p{IsHiragana}\p{IsHiragana}\p{IsCJKUnifiedIdeographs}]+成績書"
                    Dim r As New System.Text.RegularExpressions.Regex(pattern, System.Text.RegularExpressions.RegexOptions.None)

                    'TextBox1.Text内で正規表現と一致する対象をすべて検索 
                    Dim mc As System.Text.RegularExpressions.MatchCollection = r.Matches(text)

                    Dim result As String
                    For Each m As System.Text.RegularExpressions.Match In mc
                        '正規表現に一致したグループの文字列を表示 
                        result = m.Groups(0).Value
                        If result.Substring(0, 1) = "日" Or result.Substring(0, 1) = "号" Then
                            result = result.Substring(1, result.Length - 1)
                        End If
                        Exit For
                    Next

                    If result = Nothing Then
                        Dim pattern2 As String = "[\p{IsHiragana}\p{IsHiragana}\p{IsCJKUnifiedIdeographs}]+結果報告書"
                        Dim r2 As New System.Text.RegularExpressions.Regex(pattern2, System.Text.RegularExpressions.RegexOptions.None)

                        'TextBox1.Text内で正規表現と一致する対象をすべて検索 
                        Dim mc2 As System.Text.RegularExpressions.MatchCollection = r2.Matches(text)

                        For Each m As System.Text.RegularExpressions.Match In mc2
                            '正規表現に一致したグループの文字列を表示 
                            result = m.Groups(0).Value
                            If result.Substring(0, 1) = "日" Or result.Substring(0, 1) = "号" Then
                                result = result.Substring(1, result.Length - 1)
                            End If
                            Exit For
                        Next

                    End If

                    If result = Nothing Then
                        Dim pattern2 As String = "[\p{IsHiragana}\p{IsHiragana}\p{IsCJKUnifiedIdeographs}]+火害調査"
                        Dim r2 As New System.Text.RegularExpressions.Regex(pattern2, System.Text.RegularExpressions.RegexOptions.None)

                        'TextBox1.Text内で正規表現と一致する対象をすべて検索 
                        Dim mc2 As System.Text.RegularExpressions.MatchCollection = r2.Matches(text)

                        For Each m As System.Text.RegularExpressions.Match In mc2
                            '正規表現に一致したグループの文字列を表示 
                            result = m.Groups(0).Value
                            If result.Substring(0, 1) = "日" Or result.Substring(0, 1) = "号" Then
                                result = result.Substring(1, result.Length - 1)
                            End If
                            Exit For
                        Next

                    End If

                    If result <> Nothing Then
                        result = result.Replace("成績書", "").Replace("結果報告書", "")
                        DataFromText = result
                    End If


                Case "試験番号"
                    Dim pattern As String = "[ⅢⅧ]\w+[-－][0-9０-９]+[-－][0-9０-９]+"
                    Dim r As New System.Text.RegularExpressions.Regex(pattern, System.Text.RegularExpressions.RegexOptions.None)

                    'TextBox1.Text内で正規表現と一致する対象をすべて検索 
                    Dim mc As System.Text.RegularExpressions.MatchCollection = r.Matches(text)

                    Dim result As String
                    For Each m As System.Text.RegularExpressions.Match In mc
                        '正規表現に一致したグループの文字列を表示 
                        result = m.Groups(0).Value
                        If result.Substring(0, 1) = "日" Or result.Substring(0, 1) = "号" Then
                            result = result.Substring(1, result.Length - 1)
                        End If
                        Exit For
                    Next

                    If result <> Nothing Then
                        'result = result.Replace("ー", "-")
                        result = StrConv(result, VbStrConv.Narrow).Replace("Ⅲ", "3").Replace("Ⅷ", "8")
                        Dim ind1 As Integer = result.IndexOf("-")
                        Dim ind2 As Integer = result.IndexOf("-", ind1 + 1)
                        Dim y1 As Integer = Integer.Parse(result.Substring(ind1 + 1, ind2 - ind1 - 1))
                        Dim no As Integer = Integer.Parse(result.Substring(ind2 + 1, result.Length - ind2 - 1))
                        'result = StrConv(result, VbStrConv.Narrow).Replace("Ⅲ", "3").Replace("Ⅷ", "8")
                        DataFromText = result.Substring(0, 2) + y1.ToString("00") + no.ToString("0000")
                    End If


                Case "依頼者名"

                    ' 法人名称
                    Dim cname As String() = {"株式会社", "有限会社", "協同組合"}
                    Dim cname_n As Integer = cname.Length

                    ' 会社名の前後に来る可能性が高い単語
                    Dim AddText As String() = {"社名", "依頼者", "報告は", "試験番号", "試験体", "財団法人", "行動記録", "所在地", "による．", "提出資料", ""}
                    Dim AddText_n As Integer = AddText.Length

                    ' 会社名に含まれていた場合に削除する単語
                    Dim EraseText As String() = {"試験機関", "財団法人", "日本建築総合試験所", "依頼者", "試験番号", "試験", "試験体", "所在地", "成績書", "験頼機", "依験頼機", "発熱性", "の"}
                    Dim EraseText_n As Integer = EraseText.Length
                    Dim pattern As String
                    Dim p1 As String = "[\p{IsKatakana}\p{IsHiragana}\p{IsCJKUnifiedIdeographs}\p{IsHalfwidthandFullwidthForms}]+"

                    If text.Substring(0, 1) = "1" Then      ' OCRからのテキストの場合（改行が含まれる）
                        Dim Exit_Flag As Boolean = False
                        For j As Integer = 0 To cname_n - 1
                            If Exit_Flag = True Then Exit For
                            For i As Integer = 0 To 1
                                Select Case i
                                    Case 0
                                        pattern = cname(j) + p1
                                    Case 1
                                        pattern = p1 + cname(j)
                                        'Case 2
                                        '    pattern = p1 + "有限会社"
                                        'Case 3
                                        '    pattern = "有限会社" + p1
                                End Select

                                Dim r As New System.Text.RegularExpressions.Regex(pattern, System.Text.RegularExpressions.RegexOptions.None)

                                'TextBox1.Text内で正規表現と一致する対象をすべて検索 
                                Dim mc As System.Text.RegularExpressions.MatchCollection = r.Matches(text)

                                Dim result As String
                                For Each m As System.Text.RegularExpressions.Match In mc
                                    '正規表現に一致したグループの文字列を表示 
                                    result = m.Groups(0).Value
                                    If (result.Substring(0, 1) = "日" Or result.Substring(0, 1) = "号") And result.Substring(0, 2) <> "日本" Then
                                        result = result.Substring(1, result.Length - 1)
                                    End If
                                    Exit For
                                Next

                                If result <> Nothing Then
                                    'Dim me1 As New MeCab
                                    'Dim t1 As String = me1.Parse(result).Replace(vbLf, vbCrLf)
                                    'me1.Dispose()
                                    'MsgBox(t1)
                                    'result = result.Replace("社名", "").Replace("株式会社", "").Replace("有限会社", "")
                                    DataFromText = result
                                    Exit_Flag = True
                                    Exit For
                                End If

                            Next
                            'If Exit_Flag = True Then Exit For
                        Next

                    ElseIf text.Substring(0, 1) = "0" Then      ' 本文からのテキストの場合（改行が含まれない）
                        Dim Exit_Flag As Boolean = False

                        For j As Integer = 0 To cname_n - 1
                            If Exit_Flag = True Then Exit For

                            For k As Integer = 0 To AddText_n - 1
                                If Exit_Flag = True Then Exit For

                                For i As Integer = 0 To 1
                                    If Exit_Flag = True Then Exit For

                                    Select Case i
                                        Case 0
                                            pattern = AddText(k) + p1 + cname(j)
                                        Case 1
                                            pattern = AddText(k) + cname(j) + p1
                                            'Case 2
                                            '    pattern = "社名" + p1 + "有限会社"
                                            'Case 3
                                            '    pattern = "社名有限会社" + p1
                                            'Case 4
                                            '    pattern = "報告は" + p1 + "株式会社"
                                            'Case 5
                                            '    pattern = "報告は株式会社" + p1
                                            'Case 6
                                            '    pattern = "報告は" + p1 + "有限会社"
                                            'Case 7
                                            '    pattern = "報告は有限会社" + p1
                                            'Case 8
                                            '    pattern = "依頼者" + p1 + "株式会社"
                                            'Case 9
                                            '    pattern = "依頼者株式会社" + p1
                                            'Case 10
                                            '    pattern = "依頼者" + p1 + "有限会社"
                                            'Case 11
                                            '    pattern = "依頼者有限会社" + p1
                                            'Case 12
                                            '    pattern = p1 + "株式会社"
                                            'Case 13
                                            '    pattern = "株式会社" + p1
                                            'Case 14
                                            '    pattern = p1 + "有限会社"
                                            'Case 15
                                            '    pattern = "有限会社" + p1
                                    End Select

                                    Dim r As New System.Text.RegularExpressions.Regex(pattern, System.Text.RegularExpressions.RegexOptions.None)

                                    'TextBox1.Text内で正規表現と一致する対象をすべて検索 
                                    Dim mc As System.Text.RegularExpressions.MatchCollection = r.Matches(text)

                                    Dim result As String
                                    For Each m As System.Text.RegularExpressions.Match In mc
                                        '正規表現に一致したグループの文字列を表示 
                                        result = m.Groups(0).Value
                                        If (result.Substring(0, 1) = "日" Or result.Substring(0, 1) = "号") And result.Substring(0, 2) <> "日本" Then
                                            result = result.Substring(1, result.Length - 1)
                                        End If
                                        Exit For
                                    Next

                                    If result <> Nothing Then

                                        For m As Integer = 0 To AddText_n - 1
                                            If AddText(m) <> "" Then
                                                result = result.Replace(AddText(m), "")
                                            End If
                                        Next
                                        For m As Integer = 0 To EraseText_n - 1
                                            If EraseText(m) <> "" Then
                                                result = result.Replace(EraseText(m), "")
                                            End If
                                        Next
                                        If result <> cname(j) Then
                                            DataFromText = result
                                            Exit_Flag = True
                                            Exit For
                                        End If
                                    End If

                                Next
                                'If Exit_Flag = True Then Exit For
                            Next
                            'If Exit_Flag = True Then Exit For
                        Next
                    End If
            End Select



        End If
    End Function




    Private Sub PDF_Convert_Button_Click(sender As Object, e As EventArgs) Handles PDF_Convert_Button.Click
        '
        '   報告書をPDFに変換し、そのデータをデータベースに転送するサブルーチン
        '
        '
        If DataGridView1.RowCount > 1 Then

            Dim db As New OdbcDbIf
            Dim tb As DataTable
            Dim Sql_Command As String


            Dim n As Integer = DataGridView1.RowCount - 1
            Dim Count As Integer = 0
            For i As Integer = 0 To n - 1
                If DataGridView1.Rows(i).Cells(5).Value = "未" And DataGridView1.Rows(i).Cells(6).Value = True Then
                    Count += 1
                End If
            Next


            If Count > 0 Then

                ' メッセージボックスに表示するテキスト
                Dim message As String = "PDF変換にチェックされたデータが" + Count.ToString + "個あります。" + vbCrLf + "入力しますか？"

                ' タイトルバーに表示するテキスト
                Dim caption As String = "確認"

                ' 表示するボタン([OK]ボタンと[キャンセル]ボタン)
                Dim buttons As MessageBoxButtons = MessageBoxButtons.OKCancel

                Dim result As DialogResult = MessageBox.Show(message, caption, buttons)

                If result = vbOK Then

                    Dim f1 As New Form2
                    f1.title = "PDF変換"
                    f1.message = "ただいまＰＤＦファイルに変換中！" + vbCrLf + vbCrLf + vbCrLf + "しばらくお待ちください。"
                    f1.Show()
                    Application.DoEvents()


                    ProgressBar1.Minimum = 0
                    ProgressBar1.Maximum = n
                    ProgressBar1.Visible = True

                    TextBox_FileLIst2.Text = ""
                    Debug.Print("OK")
                    Try
                        FileMakerServer = TextBox_FileMakerServer.Text
                        db.Connect()

                        'If System.IO.File.Exists(CmdFile) = False Then
                        '    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(CmdFile))
                        'End If
                        'Dim sw As New StreamWriter(CmdFile, False, System.Text.Encoding.GetEncoding("utf-8"))
                        'sw.WriteLine("2")  ' コマンド番号

                        For i As Integer = 0 To n - 1
                            If DataGridView1.Rows(i).Cells(6).Value = True Then

                                Dim fname As String = DataGridView1.Rows(i).Cells(1).Value
                                Dim filename1 As String = DataGridView1.Rows(i).Cells(2).Value

                                Dim FolderPath As String = Path.GetDirectoryName(filename1)
                                Dim PdfFilename As String = Path.GetFileNameWithoutExtension(filename1) + ".pdf"

                                Dim PdfPath = FolderPath + "\" + PdfFilename

                                Try
                                    Dim fileInfo As New FileInfo(FolderPath)
                                    Dim fileSec As FileSecurity = fileInfo.GetAccessControl()

                                    ' アクセス権限をEveryoneに対しフルコントロール許可
                                    Dim accessRule As New FileSystemAccessRule("Everyone", FileSystemRights.FullControl, AccessControlType.Allow)
                                    fileSec.AddAccessRule(accessRule)
                                    fileInfo.SetAccessControl(fileSec)

                                    ' ファイルの読み取り専用属性を削除
                                    If (fileInfo.Attributes And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                                        fileInfo.Attributes = FileAttributes.Normal
                                    End If
                                Catch ex As Exception

                                End Try


                                Dim Ok As Integer = -1
                                Dim AddText As String
                                If System.IO.File.Exists(PdfPath) = False Then
                                    Ok = DocuToPdf(filename1, PdfPath, 600)
                                    If Ok = 0 Then
                                        AddText = "（新規作成）"
                                    Else
                                        AddText = "（失敗）"
                                    End If
                                Else
                                    Ok = 0
                                    AddText = "（既存）"
                                End If

                                If Ok = 0 Then


                                    Sql_Command = "UPDATE """ + Table + """ SET ""PdfPath"" = '" + PdfPath.Replace("'", "''") + "'"
                                    Sql_Command += "  WHERE ""ファイル名"" = '" + fname.Replace("'", "''") + "'"
                                    tb = db.ExecuteSql(Sql_Command)
                                    DataGridView1.Rows(i).Cells(5).Value = "済"
                                    DataGridView1.Rows(i).Cells(6).Value = False

                                    'sw.WriteLine(fname)
                                Else
                                    DataGridView1.Rows(i).Cells(5).Value = "未"
                                    DataGridView1.Rows(i).Cells(6).Value = True
                                End If

                                Me.TextBox_FileLIst2.Text += (i + 1).ToString() + "/" + n.ToString() + ":" + PdfPath + AddText + vbCrLf
                                Me.TextBox_FileLIst2.SelectionStart = Me.TextBox_FileLIst2.Text.Length
                                Me.TextBox_FileLIst2.Focus()
                                Me.TextBox_FileLIst2.ScrollToCaret()

                                ProgressBar1.Value = i + 1


                                'Dim fname As String = DataGridView1.Rows(i).Cells(1).Value
                                'Dim filename As String = DataGridView1.Rows(i).Cells(2).Value
                                'Sql_Command = "INSERT INTO """ + Table + """ (""FilePath"",""ファイル名"",""入力"")"
                                'Sql_Command += " VALUES ('" + filename.Replace("'", "''") + "','" + fname.Replace("'", "''") + "','未読')"
                                'tb = db.ExecuteSql(Sql_Command)
                                'DataGridView1.Rows(i).Cells(4).Value = True

                                'sw.WriteLine(fname)
                                'Count += 1
                            End If
                            Application.DoEvents()
                        Next

                        Me.TextBox_FileLIst2.Text += "変換終了" + vbCrLf
                        Me.TextBox_FileLIst2.SelectionStart = Me.TextBox_FileLIst2.Text.Length
                        Me.TextBox_FileLIst2.Focus()
                        Me.TextBox_FileLIst2.ScrollToCaret()

                        db.Disconnect()

                        ' Shift-Jisでファイルを作成


                        '２行書き込み


                        'ストリームを閉じる
                        'sw.Close()
                        ProgressBar1.Visible = False
                        f1.Close()
                        f1.Dispose()

                    Catch e1 As Exception
                        ProgressBar1.Visible = False
                        f1.Close()
                        f1.Dispose()
                    End Try



                End If
            Else
                MessageBox.Show("未変換のデータはありません", "警告", MessageBoxButtons.OK)
            End If



        Else
            MessageBox.Show("データがありません!!", "警告", MessageBoxButtons.OK)

        End If
    End Sub



    Private Sub Data_Input_Button2_Click(sender As Object, e As EventArgs) Handles Data_Input_Button2.Click
        '
        '   資料のデータをデータベースに転送するサブルーチン
        '
        '

        If DataGridView2.RowCount > 1 Then

            Dim db As New OdbcDbIf
            Dim tb As DataTable
            Dim Sql_Command As String
            Dim ROnly = スキャンデータOnlyCheckBox.Checked And FolderNameCheckBox2.Checked

            Dim n As Integer = DataGridView2.RowCount - 1
            Dim Count As Integer = 0
            For i As Integer = 0 To n - 1
                If DataGridView2.Rows(i).Cells(4).Value = True Then
                    Count += 1
                End If
            Next

            If Count > 0 Then

                ' メッセージボックスに表示するテキスト
                Dim message As String = "未入力のデータが" + Count.ToString + "個あります。" + vbCrLf + "入力しますか？"

                ' タイトルバーに表示するテキスト
                Dim caption As String = "確認"

                ' 表示するボタン([OK]ボタンと[キャンセル]ボタン)
                Dim buttons As MessageBoxButtons = MessageBoxButtons.OKCancel

                Dim result As DialogResult = MessageBox.Show(message, caption, buttons)

                If result = vbOK Then

                    Dim f1 As New Form2
                    f1.title = "PDF変換"
                    f1.message = "ただいまデータをに入力中！" + vbCrLf + vbCrLf + vbCrLf + "しばらくお待ちください。"
                    f1.Show()
                    Application.DoEvents()

                    ProgressBar2.Minimum = 0
                    ProgressBar2.Maximum = n
                    ProgressBar2.Visible = True

                    TextBox_FileLIst3.Text = ""
                    Debug.Print("OK")
                    Try
                        FileMakerServer = TextBox_FileMakerServer.Text
                        db.Connect()

                        'If System.IO.File.Exists(CmdFile) = False Then
                        '    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(CmdFile))
                        'End If
                        'Dim sw As New StreamWriter(CmdFile, False, System.Text.Encoding.GetEncoding("utf-8"))
                        'sw.WriteLine("1")  ' コマンド番号

                        Dim ii As Integer = 0
                        For i As Integer = 0 To n - 1
                            If DataGridView2.Rows(i).Cells(4).Value = True Then

                                Dim fname As String = DataGridView2.Rows(i).Cells(1).Value
                                Dim filename1 As String = DataGridView2.Rows(i).Cells(2).Value

                                Dim FolderPath As String = Path.GetDirectoryName(filename1)

                                Dim Path0 As String, name1 As String
                                Dim text1 As String, text2 As String
                                Dim testname As String
                                Dim username As String

                                Dim ocr_flag1 As Boolean = False
                                Dim input_ok As Boolean = False

                                'Sql_Command = "INSERT INTO """ + Table + """ (""FilePath"",""ファイル名"",""入力"")"
                                'Sql_Command += " VALUES ('" + filename1.Replace("'", "''") + "','" + fname.Replace("'", "''") + "','未読')"
                                'tb = db.ExecuteSql(Sql_Command)

                                Dim fname1 As String = System.IO.Path.GetFileNameWithoutExtension(fname)
                                Dim filekind As String = ""


                                Dim filename2 As String ' ファイル名から試験番号を抽出し、データベースに入力
                                'If filename2 <> "" Then
                                '    Dim kind2 As String = filename2.Substring(0, 2)
                                '    Dim year2 As String = filename2.Substring(2, 2)
                                '    Dim no2 As String = filename2.Substring(4, 4)
                                '    Dim eda2 As String = ""
                                '    If filename2.Contains("(") = True Then
                                '        eda2 = filename2.Substring(9, 2)
                                '    End If
                                '    Sql_Command = "UPDATE """ + Table + """ SET ""分類2"" = '" + kind2 + "',""年度2"" = '" + year2 + "',""番号2"" = '" + no2
                                '    Sql_Command += "',""枝番2"" = '" + eda2 + "',""試験番号2"" = '" + filename2 + "'"
                                '    Sql_Command += "  WHERE ""ファイル名"" = '" + fname.Replace("'", "''") + "'"
                                '    tb = db.ExecuteSql(Sql_Command)
                                'End If

                                If ROnly Then

                                    Sql_Command = "INSERT INTO """ + Table + """ (""FilePath"",""ファイル名"",""入力"")"
                                    Sql_Command += " VALUES ('" + filename1.Replace("'", "''") + "','" + fname.Replace("'", "''") + "','未読')"
                                    tb = db.ExecuteSql(Sql_Command)
                                    input_ok = True

                                    Path0 = FolderPath.Replace("\スキャンデータ", "")
                                    name1 = Path0.Substring(Path0.LastIndexOf("\") + 1, Path0.Length - Path0.LastIndexOf("\") - 1)

                                    filename2 = TestNo(name1) ' ファイル名から試験番号を抽出し、データベースに入力
                                    If filename2 <> "" Then
                                        Dim kind2 As String = filename2.Substring(0, 2)
                                        Dim year2 As String = filename2.Substring(2, 2)
                                        Dim no2 As String = filename2.Substring(4, 4)
                                        Dim eda2 As String = ""
                                        If filename2.Contains("(") = True Then
                                            eda2 = filename2.Substring(9, 2)
                                        End If
                                        Sql_Command = "UPDATE """ + Table + """ SET ""分類2"" = '" + kind2 + "',""年度2"" = '" + year2 + "',""番号2"" = '" + no2
                                        Sql_Command += "',""枝番2"" = '" + eda2 + "',""試験番号2"" = '" + filename2 + "'"
                                        Sql_Command += "  WHERE ""ファイル名"" = '" + fname.Replace("'", "''") + "'"
                                        tb = db.ExecuteSql(Sql_Command)
                                    End If


                                    testname = name1
                                    username = name1
                                    Sql_Command = "UPDATE """ + Table + """ SET ""試験項目2"" = '" + testname + "',""依頼者名2"" = '" + username + "'"
                                    Sql_Command += "  WHERE ""ファイル名"" = '" + fname.Replace("'", "''") + "'"
                                    tb = db.ExecuteSql(Sql_Command)

                                    If fname1.Contains("申請書") Or fname1.Contains("依頼書") Then
                                        ocr_flag1 = True
                                    Else
                                        ocr_flag1 = True
                                    End If
                                Else
                                    fname1 = fname1.Trim
                                    filekind = fname1.Substring(fname1.Length - 1, 1)
                                    If CountChar(fname, "_") = 3 And IsNumeric(filekind) Then

                                        Sql_Command = "INSERT INTO """ + Table + """ (""FilePath"",""ファイル名"",""入力"")"
                                        Sql_Command += " VALUES ('" + filename1.Replace("'", "''") + "','" + fname.Replace("'", "''") + "','未読')"
                                        tb = db.ExecuteSql(Sql_Command)
                                        input_ok = True

                                        filename2 = TestNo(fname) ' ファイル名から試験番号を抽出し、データベースに入力
                                        If filename2 <> "" Then
                                            Dim kind2 As String = filename2.Substring(0, 2)
                                            Dim year2 As String = filename2.Substring(2, 2)
                                            Dim no2 As String = filename2.Substring(4, 4)
                                            Dim eda2 As String = ""
                                            If filename2.Contains("(") = True Then
                                                eda2 = filename2.Substring(9, 2)
                                            End If
                                            Sql_Command = "UPDATE """ + Table + """ SET ""分類2"" = '" + kind2 + "',""年度2"" = '" + year2 + "',""番号2"" = '" + no2
                                            Sql_Command += "',""枝番2"" = '" + eda2 + "',""試験番号2"" = '" + filename2 + "'"
                                            Sql_Command += "  WHERE ""ファイル名"" = '" + fname.Replace("'", "''") + "'"
                                            tb = db.ExecuteSql(Sql_Command)
                                        End If

                                        'If fname.Contains("_") = True Then ' ファイル名から依頼者名と試験項目（または資料名）を抽出し、データベースに入力
                                        '    If CountChar(fname, "_") > 1 Then
                                        '        filekind = fname1.Substring(fname1.Length - 1, 1)
                                        '        If IsNumeric(filekind) Then
                                        Dim p1 As Integer = fname.IndexOf("_", 0)
                                        Dim p2 As Integer = fname.IndexOf("_", p1 + 1)
                                        Dim p3 As Integer = fname.IndexOf("_", p2 + 1)
                                        username = fname.Substring(p1 + 1, p2 - p1 - 1)
                                        testname = fname.Substring(p2 + 1, p3 - p2 - 1)
                                        Sql_Command = "UPDATE """ + Table + """ SET ""試験項目2"" = '" + testname + "',""依頼者名2"" = '" + username + "'"
                                        Sql_Command += "  WHERE ""ファイル名"" = '" + fname.Replace("'", "''") + "'"
                                        tb = db.ExecuteSql(Sql_Command)
                                        '        End If
                                        '    End If
                                        'End If
                                        If filekind = "1" Then
                                            ocr_flag1 = True
                                        Else
                                            ocr_flag1 = False
                                        End If
                                    End If

                                    ' 資料１（依頼書）の場合はPDFをXDWに変換し、OCRをかけてテキストを抽出する。資料２は処理しない
                                    If ocr_flag1 Then
                                        Dim myDodumentFolder = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal)
                                        Dim myUserFolder = Path.GetDirectoryName(myDodumentFolder)

                                        ' XDWファイルは c:\user\(ユーザーフォルダー)\PdfToXdwフォルダーに保存する。
                                        Dim myWdxFolder = myUserFolder + "\PdfToXdw"
                                        If System.IO.Directory.Exists(myWdxFolder) = False Then     ' ディレクトリーが存在しない場合は作成
                                            System.IO.Directory.CreateDirectory(myWdxFolder)
                                        End If

                                        Dim path1 As String = filename1
                                        Dim path2 As String = myWdxFolder + "\" + fname1 + ".xdw"

                                        If System.IO.File.Exists(path2) = False Then     ' すでに同じ.xdwファイルが存在している場合はそれを利用する。
                                            'System.IO.File.Delete(path2)


                                            Dim FolderPath1 As String = Path.GetDirectoryName(path1)
                                            Dim fileInfo1 As New FileInfo(FolderPath1)
                                            Dim fileSec1 As FileSecurity = fileInfo1.GetAccessControl()

                                            'Try
                                            '    ' アクセス権限をEveryoneに対しフルコントロール許可
                                            '    Dim accessRule1 As New FileSystemAccessRule("Everyone", FileSystemRights.FullControl, AccessControlType.Allow)
                                            '    fileSec1.AddAccessRule(accessRule1)
                                            '    fileInfo1.SetAccessControl(fileSec1)
                                            'Catch ex As Exception

                                            'End Try



                                            Dim FolderPath2 As String = Path.GetDirectoryName(path1)
                                            Dim fileInfo2 As New FileInfo(FolderPath2)
                                            Dim fileSec2 As FileSecurity = fileInfo2.GetAccessControl()

                                            'Try
                                            '    ' アクセス権限をEveryoneに対しフルコントロール許可
                                            '    Dim accessRule2 As New FileSystemAccessRule("Everyone", FileSystemRights.FullControl, AccessControlType.Allow)
                                            '    fileSec2.AddAccessRule(accessRule2)
                                            '    fileInfo2.SetAccessControl(fileSec2)
                                            'Catch ex As Exception

                                            'End Try


                                            ' PDFからXDWを作成
                                            Dim r1 As Integer = Xdwapi.XDW_CreateXdwFromImagePdfFile(path1, path2)

                                        End If


                                        'Dim text1 As String
                                        Dim Handle As Xdwapi.XDW_DOCUMENT_HANDLE = New Xdwapi.XDW_DOCUMENT_HANDLE()
                                        Dim mode As Xdwapi.XDW_OPEN_MODE_EX = New Xdwapi.XDW_OPEN_MODE_EX()
                                        With mode
                                            .Option = Xdwapi.XDW_OPEN_UPDATE    ' 編集モード
                                            .AuthMode = Xdwapi.XDW_AUTH_NODIALOGUE
                                        End With
                                        Dim api_result As Integer = Xdwapi.XDW_OpenDocumentHandle(path2, Handle, mode)

                                        'api_result = Xdwapi.XDW_OpenDocumentHandle(path2, Handle, mode)     ' 再度ファイルを開く

                                        'Dim result3 As Integer = Xdwapi.XDW_RotatePageAuto(Handle, page)    ' 横書きの場合は90度回
                                        Dim result3 As Integer = 0
                                        If result3 >= 0 Then
                                            Dim ocr_optoin As Xdwapi.XDW_OCR_OPTION_V7 = New Xdwapi.XDW_OCR_OPTION_V7
                                            With ocr_optoin
                                                .NoiseReduction = Xdwapi.XDW_REDUCENOISE_NORMAL
                                                .Language = Xdwapi.XDW_OCR_LANGUAGE_AUTO
                                                .InsertSpaceCharacter = 0
                                                .Form = Xdwapi.XDW_OCR_FORM_AUTO
                                                .Column = Xdwapi.XDW_OCR_COLUMN_AUTO
                                                .EngineLevel = Xdwapi.XDW_OCR_ENGINE_LEVEL_STANDARD
                                            End With
                                            result3 = Xdwapi.XDW_ApplyOcr(Handle, 1, Xdwapi.XDW_OCR_ENGINE_DEFAULT, ocr_optoin)
                                            System.Threading.Thread.Sleep(1000)
                                            OcrFlag = True
                                            If result3 >= 0 Then

                                                result3 = -1
                                                result3 = Xdwapi.XDW_GetPageTextToMemory(Handle, 1, text1)
                                            End If

                                            'If result3 >= 0 Then
                                            '    If text1 <> Nothing Then
                                            '        'DocuToText = "1" + text1
                                            '    End If
                                            'End If
                                        End If
                                        Xdwapi.XDW_CloseDocumentHandle(Handle)
                                        'Xdwapi.XDW_Finalize()

                                        If text1 <> Nothing Then
                                            text1 = text1.Replace("'", "''").Replace(vbCrLf, "")
                                            'text2 = text2.Replace("'", "''").Replace(vbCrLf, "")

                                            Sql_Command = "UPDATE """ + Table + """ SET ""page1"" = '" + text1.Replace("'", "''") + "'"
                                            Sql_Command += "  WHERE ""ファイル名"" = '" + fname.Replace("'", "''") + "'"
                                            tb = db.ExecuteSql(Sql_Command)
                                        End If

                                        If System.IO.File.Exists(path2) = True Then     ' .xdwファイルを削除する。
                                            System.IO.File.Delete(path2)
                                        End If
                                    End If
                                End If

                                If input_ok Then

                                    DataGridView2.Rows(i).Cells(3).Value = "済"
                                    DataGridView2.Rows(i).Cells(4).Value = False
                                Else
                                    filename1 += "(file name error)"

                                End If
                                'sw.WriteLine(fname)
                                'Count += 1
                                Me.TextBox_FileLIst3.Text += filename1 + vbCrLf
                                Me.TextBox_FileLIst3.SelectionStart = Me.TextBox_FileLIst2.Text.Length
                                Me.TextBox_FileLIst3.Focus()
                                Me.TextBox_FileLIst3.ScrollToCaret()

                                ProgressBar2.Value = i + 1

                            End If
                            Application.DoEvents()
                        Next
                        db.Disconnect()

                        ' Shift-Jisでファイルを作成


                        '２行書き込み


                        'ストリームを閉じる
                        'sw.Close()
                        f1.Close()
                        f1.Dispose()
                        ProgressBar2.Visible = False

                    Catch e1 As Exception
                        f1.Close()
                        f1.Dispose()
                        ProgressBar2.Visible = False
                    End Try

                    'f1.Close()
                    'f1.Dispose()
                    'ProgressBar2.Visible = False
                End If
            Else
                MessageBox.Show("未入力のデータはありません", "警告", MessageBoxButtons.OK)
            End If



        Else
            MessageBox.Show("データがありません!!", "警告", MessageBoxButtons.OK)

        End If
    End Sub


    Private Sub CheckBox_Input2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_Input2.CheckedChanged
        '
        '   すべての資料の入力チャックボックスのON/OFF切替
        '
        Dim n As Integer = DataGridView2.RowCount - 1
        If n > 0 Then
            For i As Integer = 0 To n - 1
                If CheckBox_Input2.Checked = True Then
                    If DataGridView2.Rows(i).Cells(3).Value = "未" Then
                        DataGridView2.Rows(i).Cells(4).Value = True
                    End If
                Else
                    If DataGridView2.Rows(i).Cells(3).Value = "未" Then
                        DataGridView2.Rows(i).Cells(4).Value = False
                    End If
                End If

            Next
        End If
    End Sub


    Private Function PCInfo(ByVal IP As String, ByVal Uname As String, ByVal Path As String, ByVal Pname As String) As Boolean
        '
        '   PCのアドレス、ホストネーム、プログラムのパス、プログラム名、接続日時をデータベースに記録する関数
        '
        '       接続出来ない場合はプログラムを終了する。
        '
        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String
        Dim td1 As DateTime = DateTime.Now
        Dim td2 As String = td1.ToString().Replace("/", "-")

        Try
            FileMakerServer = TextBox_FileMakerServer.Text
            db.Connect()

            Sql_Command = "SELECT ""IP"",""Path"",""ProgramName"",""UserName"" FROM """ + Table2 + """ WHERE (""IP"" = '" + IP + "')"
            tb = db.ExecuteSql(Sql_Command)
            Dim n2 As Integer = tb.Rows.Count
            If n2 > 0 Then
                Sql_Command = "UPDATE """ + Table2 + """ SET ""接続日時"" = TIMESTAMP '" + td2 + "',""Path"" = '" + Path + "'"
                Sql_Command += " WHERE (""IP"" = '" + IP + "')"
                tb = db.ExecuteSql(Sql_Command)
            Else
                Sql_Command = "INSERT INTO """ + Table2 + """ (""IP"",""Path"",""ProgramName"",""UserName"",""接続日時"")"
                Sql_Command += " VALUES ('" + IP + "','" + Path + "','" + Pname + "','" + Uname + "',TIMESTAMP '" + td2 + "')"
                tb = db.ExecuteSql(Sql_Command)
            End If

            db.Disconnect()
            PCInfo = True
        Catch e1 As Exception
            PCInfo = False
        End Try

    End Function


    Private Sub SelectXdwButton_Click(sender As Object, e As EventArgs) Handles SelectXdwButton.Click


        Dim FileFlag As String

        Dim filekind As String
        'If RadioButton_xdw.Checked = True Then
        filekind = "ドキュワークスファイル(*.xdw;*.xbd)|*.xdw;*.xbd"
        FileFlag = "xdw"

        'Else
        '    filekind = "PDFファイル(*.pdf)|*.pdf"
        '    FileFlag = "pdf"
        'End If
        'OpenFileDialogクラスのインスタンスを作成
        Dim ofd As New OpenFileDialog()

        'ofd.FileName = "default.html"
        'はじめに表示されるフォルダを指定する
        '指定しない（空の文字列）の時は、現在のディレクトリが表示される

        If TextBox_FolderName1.Text <> "" Then
            ofd.InitialDirectory = TextBox_FolderName1.Text
        Else
            'ofd.InitialDirectory = "\\192.168.0.173\disk1\報告書（耐火）＿業務課から"
            ofd.InitialDirectory = "W:\耐火構造\依頼試験\案件フォルダ【取扱注意】元データのため削除禁止\2015年度"

        End If
        '[ファイルの種類]に表示される選択肢を指定する
        '指定しないとすべてのファイルが表示される
        ofd.Filter = filekind
        '[ファイルの種類]ではじめに選択されるものを指定する
        '2番目の「すべてのファイル」が選択されているようにする
        ofd.FilterIndex = 2
        'タイトルを設定する
        ofd.Title = "ファイルを選択してください（複数可）"
        'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
        ofd.RestoreDirectory = True
        '存在しないファイルの名前が指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        ofd.CheckFileExists = True
        '存在しないパスが指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        ofd.CheckPathExists = True
        '複数のファイルを選択できるようにする
        ofd.Multiselect = True
        'ダイアログを表示する
        If ofd.ShowDialog() = DialogResult.OK Then

            filename = ofd.FileNames
            Dim Count As Integer = MakeXdwList(filename)

            If Count = 0 Then
                MsgBox("このフォルダーには報告書ファイルはありません！", vbOK, "確認")
            Else
                MsgBox("このフォルダーには" + Count.ToString + "個の報告書ファイルがありました。", vbOK, "確認")
            End If

        End If
    End Sub

    Private Sub SelectPdfButton_Click(sender As Object, e As EventArgs) Handles SelectPdfButton.Click


        Dim FileFlag As String

        Dim filekind As String
        'If RadioButton_xdw.Checked = True Then
        'filekind = "ドキュワークスファイル(*.xdw;*.xbd)|*.xdw;*.xbd"
        'FileFlag = "xdw"

        'Else
        filekind = "PDFファイル(*.pdf)|*.pdf"
        FileFlag = "pdf"
        'End If
        'OpenFileDialogクラスのインスタンスを作成
        Dim ofd As New OpenFileDialog()

        'ofd.FileName = "default.html"
        'はじめに表示されるフォルダを指定する
        '指定しない（空の文字列）の時は、現在のディレクトリが表示される

        If TextBox_FolderName2.Text <> "" Then
            ofd.InitialDirectory = TextBox_FolderName2.Text
        Else
            ofd.InitialDirectory = "\\192.168.0.173\disk1\SCAN"
        End If
        '[ファイルの種類]に表示される選択肢を指定する
        '指定しないとすべてのファイルが表示される
        ofd.Filter = filekind
        '[ファイルの種類]ではじめに選択されるものを指定する
        '2番目の「すべてのファイル」が選択されているようにする
        ofd.FilterIndex = 2
        'タイトルを設定する
        ofd.Title = "ファイルを選択してください（複数可）"
        'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
        ofd.RestoreDirectory = True
        '存在しないファイルの名前が指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        ofd.CheckFileExists = True
        '存在しないパスが指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        ofd.CheckPathExists = True
        '複数のファイルを選択できるようにする
        ofd.Multiselect = True
        'ダイアログを表示する
        If ofd.ShowDialog() = DialogResult.OK Then

            filename = ofd.FileNames
            Dim Count As Integer = MakePdfList(filename)

            If Count = 0 Then
                MsgBox("このフォルダーには資料ファイルはありません！", vbOK, "確認")
            Else
                MsgBox("このフォルダーには" + Count.ToString + "個の資料ファイルがありました。", vbOK, "確認")
            End If

        End If

    End Sub


    Private Sub xdwFolderRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles xdwFolderRadioButton.CheckedChanged, xdwFileRadioButton.CheckedChanged

        If xdwFolderRadioButton.Checked = True Then
            xdwModeChange("folder")
        Else
            xdwModeChange("file")
        End If

    End Sub

    Private Sub xdwModeChange(ByVal flag As String)
        If flag = "folder" Then
            Select_Read_Folder_Button.Visible = True
            FolderMenuButton1.Visible = True
            BeforeFolderButton.Visible = True
            NextFolderButton.Visible = True
            DocuReadButton.Visible = True
            SelectXdwButton.Visible = False
            Label4.Visible = True
            Label5.Visible = True
            Label11.Visible = False

            Label9.Text = "（手順4）データベースへの入力"
            Label10.Text = "（手順5）PDF変換"

        ElseIf flag = "file" Then
            Select_Read_Folder_Button.Visible = False
            FolderMenuButton1.Visible = False
            BeforeFolderButton.Visible = False
            NextFolderButton.Visible = False
            DocuReadButton.Visible = False
            SelectXdwButton.Visible = True
            Label4.Visible = False
            Label5.Visible = False
            Label11.Visible = True

            Label9.Text = "（手順3）データベースへの入力"
            Label10.Text = "（手順4）PDF変換"
        End If
    End Sub

    Private Sub pdfFolderRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles pdfFolderRadioButton.CheckedChanged, pdfFileRadioButton.CheckedChanged
        If pdfFolderRadioButton.Checked = True Then
            pdfModeChange("folder")
        Else
            pdfModeChange("file")
        End If

    End Sub

    Private Sub 成績書OnlyCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles 成績書OnlyCheckBox.CheckedChanged
        If 成績書OnlyCheckBox.Checked Then
            FolderNameCheckBox1.Visible = True
            FolderNameCheckBox1.Checked = True
        Else
            FolderNameCheckBox1.Visible = False
            FolderNameCheckBox1.Checked = False
        End If
    End Sub

    Private Sub スキャンデータOnlyCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles スキャンデータOnlyCheckBox.CheckedChanged
        If スキャンデータOnlyCheckBox.Checked Then
            FolderNameCheckBox2.Visible = True
            FolderNameCheckBox2.Checked = True
        Else
            FolderNameCheckBox2.Visible = False
            FolderNameCheckBox2.Checked = False
        End If
    End Sub

    Private Sub pdfModeChange(ByVal flag As String)
        If flag = "folder" Then
            Select_Read_Folder_Button2.Visible = True
            FolderMenuButton2.Visible = True
            BeforeFolderButton2.Visible = True
            NextFolderButton2.Visible = True
            PdfReadButton.Visible = True
            SelectPdfButton.Visible = False
            Label7.Visible = True
            Label6.Visible = True
            Label12.Visible = False

            Label8.Text = "（手順4）データベースへの入力"


        ElseIf flag = "file" Then
            Select_Read_Folder_Button2.Visible = False
            FolderMenuButton2.Visible = False
            BeforeFolderButton2.Visible = False
            NextFolderButton2.Visible = False
            PdfReadButton.Visible = False
            SelectPdfButton.Visible = True
            Label7.Visible = False
            Label6.Visible = False
            Label12.Visible = True

            Label8.Text = "（手順3）データベースへの入力"

        End If
    End Sub

    Private Sub FolderSaveButton1_Click(sender As Object, e As EventArgs) Handles FolderSaveButton1.Click


        If TextBox_FolderName1.Text <> "" Then
            Try
                Dim db As New OdbcDbIf
                Dim tb As DataTable
                Dim Sql_Command As String

                Dim td1 As DateTime = DateTime.Now
                Dim td2 As String = td1.ToString().Replace("/", "-")

                FileMakerServer = TextBox_FileMakerServer.Text
                db.Connect()

                Dim Path As String = TextBox_FolderName1.Text
                Dim Kind As String = "報告書"

                Sql_Command = "SELECT ""IP"",""Path"" FROM """ + Table3 + """ WHERE (""IP"" = '" + MyIP + "' AND ""Path"" = '" + Path + "' AND ""種類"" = '" + Kind + "')"
                tb = db.ExecuteSql(Sql_Command)

                Dim n2 As Integer = tb.Rows.Count
                If n2 > 0 Then
                    Sql_Command = "UPDATE """ + Table3 + """ SET ""接続日時"" = TIMESTAMP '" + td2 + "'"
                    Sql_Command += " WHERE (""IP"" = '" + MyIP + "' AND ""Path"" = '" + Path + "' AND ""種類"" = '" + Kind + "')"
                    tb = db.ExecuteSql(Sql_Command)
                Else
                    Sql_Command = "INSERT INTO """ + Table3 + """ (""IP"",""Path"",""種類"",""接続日時"")"
                    Sql_Command += " VALUES ('" + MyIP + "','" + Path + "','" + Kind + "',TIMESTAMP '" + td2 + "')"
                    tb = db.ExecuteSql(Sql_Command)
                End If

                db.Disconnect()

            Catch e1 As Exception

            End Try
        End If
    End Sub


    Private Sub FolderSaveButton2_Click(sender As Object, e As EventArgs) Handles FolderSaveButton2.Click


        If TextBox_FolderName2.Text <> "" Then
            Try
                Dim db As New OdbcDbIf
                Dim tb As DataTable
                Dim Sql_Command As String

                Dim td1 As DateTime = DateTime.Now
                Dim td2 As String = td1.ToString().Replace("/", "-")

                FileMakerServer = FileMakerServer1
                db.Connect()

                Dim Path As String = TextBox_FolderName2.Text
                Dim Kind As String = "資料"

                Sql_Command = "SELECT ""IP"",""Path"" FROM """ + Table3 + """ WHERE (""IP"" = '" + MyIP + "' AND ""Path"" = '" + Path + "' AND ""種類"" = '" + Kind + "')"
                tb = db.ExecuteSql(Sql_Command)

                Dim n2 As Integer = tb.Rows.Count
                If n2 > 0 Then
                    Sql_Command = "UPDATE """ + Table3 + """ SET ""接続日時"" = TIMESTAMP '" + td2 + "'"
                    Sql_Command += " WHERE (""IP"" = '" + MyIP + "' AND ""Path"" = '" + Path + "' AND ""種類"" = '" + Kind + "')"
                    tb = db.ExecuteSql(Sql_Command)
                Else
                    Sql_Command = "INSERT INTO """ + Table3 + """ (""IP"",""Path"",""種類"",""接続日時"")"
                    Sql_Command += " VALUES ('" + MyIP + "','" + Path + "','" + Kind + "',TIMESTAMP '" + td2 + "')"
                    tb = db.ExecuteSql(Sql_Command)
                End If

                db.Disconnect()

            Catch e1 As Exception

            End Try
        End If
    End Sub

    Private Sub FolderMenuButton1_Click(sender As Object, e As EventArgs) Handles FolderMenuButton1.Click

        Dim menuForm As New ListForm1
        DataKind = "報告書"
        'menuForm.Show()


        menuForm.StartPosition = FormStartPosition.CenterParent
        If menuForm.ShowDialog = DialogResult.OK Then         '値を受け取る
            TextBox_FolderName1.Text = menuForm.GetValue
            DcuPath = TextBox_FolderName1.Text
            PdfPath = PdfSaveFolder + "\" + System.IO.Path.GetFileName(System.IO.Path.GetFileName(DcuPath))
        End If
        menuForm.Dispose()
    End Sub


    Private Sub FolderMenuButton2_Click(sender As Object, e As EventArgs) Handles FolderMenuButton2.Click
        Dim menuForm As New ListForm1
        DataKind = "資料"
        'menuForm.Show()


        menuForm.StartPosition = FormStartPosition.CenterParent
        If menuForm.ShowDialog = DialogResult.OK Then         '値を受け取る
            TextBox_FolderName2.Text = menuForm.GetValue
            PdfPath = TextBox_FolderName2.Text

        End If
        menuForm.Dispose()
    End Sub



    Private Sub NextFolderButton_Click(sender As Object, e As EventArgs) Handles NextFolderButton.Click


        If TextBox_FolderName1.Text <> "" Then
            Dim Path0 = System.IO.Path.GetDirectoryName(TextBox_FolderName1.Text)
            Dim dname = System.IO.Path.GetFileName(TextBox_FolderName1.Text)
            Dim dir2 As String(), dname2 As String(), dname3 As String()

            If Path0 <> "" And dname <> "" Then
                dir2 = Directory.GetDirectories(Path0, "*", SearchOption.TopDirectoryOnly)
                Dim n As Integer = dir2.Length
                ReDim dname2(n - 1), dname3(n - 1)
                For i As Integer = 0 To n - 1
                    dname2(i) = System.IO.Path.GetFileName(dir2(i))
                Next
                Dim cmp As StringComparer = StringComparer.OrdinalIgnoreCase
                Array.Sort(dname2, cmp)
                Dim nextfolder As String, index As Integer = -1
                Dim n2 As Integer = 0

                For i As Integer = 0 To n - 1
                    If dname2(i).Substring(0, 1) <> "." Then
                        dname3(n2) = dname2(i)
                        n2 += 1
                    End If
                Next


                For i As Integer = 0 To n2 - 1
                    If dname3(i) = dname Then
                        index = i
                        Exit For
                    End If
                Next
                If index >= 0 And index < n2 - 1 Then
                    nextfolder = dname3(index + 1)
                    TextBox_FolderName1.Text = Path0 + "\" + nextfolder
                    DcuPath = TextBox_FolderName1.Text
                    PdfPath = PdfSaveFolder + "\" + System.IO.Path.GetFileName(System.IO.Path.GetFileName(DcuPath))
                Else

                    'PlaySound("SystemHand", IntPtr.Zero, PlaySoundFlags.SND_ALIAS Or PlaySoundFlags.SND_NODEFAULT)
                    MsgBox("次はありません！", vbOK, "エラー")
                End If

            End If
        End If
    End Sub


    Private Sub BeforeFolderButton_Click(sender As Object, e As EventArgs) Handles BeforeFolderButton.Click


        If TextBox_FolderName1.Text <> "" Then
            Dim Path0 = System.IO.Path.GetDirectoryName(TextBox_FolderName1.Text)
            Dim dname = System.IO.Path.GetFileName(TextBox_FolderName1.Text)
            Dim dir2 As String(), dname2 As String(), dname3 As String()

            If Path0 <> "" And dname <> "" Then
                dir2 = Directory.GetDirectories(Path0, "*", SearchOption.TopDirectoryOnly)
                Dim n As Integer = dir2.Length
                ReDim dname2(n - 1), dname3(n - 1)
                For i As Integer = 0 To n - 1
                    dname2(i) = System.IO.Path.GetFileName(dir2(i))
                Next
                Dim cmp As StringComparer = StringComparer.OrdinalIgnoreCase
                Array.Sort(dname2, cmp)
                Dim nextfolder As String, index As Integer = -1
                Dim n2 As Integer = 0

                For i As Integer = 0 To n - 1
                    If dname2(i).Substring(0, 1) <> "." Then
                        dname3(n2) = dname2(i)
                        n2 += 1
                    End If
                Next


                For i As Integer = 0 To n2 - 1
                    If dname3(i) = dname Then
                        index = i
                        Exit For
                    End If
                Next
                If index >= 1 And index < n2 Then
                    nextfolder = dname3(index - 1)
                    TextBox_FolderName1.Text = Path0 + "\" + nextfolder
                    DcuPath = TextBox_FolderName1.Text
                    PdfPath = PdfSaveFolder + "\" + System.IO.Path.GetFileName(System.IO.Path.GetFileName(DcuPath))
                Else
                    'PlaySound("SystemHand", IntPtr.Zero, PlaySoundFlags.SND_ALIAS Or PlaySoundFlags.SND_NODEFAULT)
                    MsgBox("前はありません！", vbOK, "エラー")
                End If

            End If
        End If
    End Sub


    Private Sub BeforeFolderButton2_Click(sender As Object, e As EventArgs) Handles BeforeFolderButton2.Click


        If TextBox_FolderName2.Text <> "" Then
            Dim Path0 = System.IO.Path.GetDirectoryName(TextBox_FolderName2.Text)
            Dim dname = System.IO.Path.GetFileName(TextBox_FolderName2.Text)
            Dim dir2 As String(), dname2 As String(), dname3 As String()

            If Path0 <> "" And dname <> "" Then
                dir2 = Directory.GetDirectories(Path0, "*", SearchOption.TopDirectoryOnly)
                Dim n As Integer = dir2.Length
                ReDim dname2(n - 1), dname3(n - 1)
                For i As Integer = 0 To n - 1
                    dname2(i) = System.IO.Path.GetFileName(dir2(i))
                Next
                Dim cmp As StringComparer = StringComparer.OrdinalIgnoreCase
                Array.Sort(dname2, cmp)
                Dim nextfolder As String, index As Integer = -1
                Dim n2 As Integer = 0

                For i As Integer = 0 To n - 1
                    If dname2(i).Substring(0, 1) <> "." Then
                        dname3(n2) = dname2(i)
                        n2 += 1
                    End If
                Next


                For i As Integer = 0 To n2 - 1
                    If dname3(i) = dname Then
                        index = i
                        Exit For
                    End If
                Next
                If index >= 1 And index < n2 Then
                    nextfolder = dname3(index - 1)
                    TextBox_FolderName2.Text = Path0 + "\" + nextfolder
                    PdfPath = TextBox_FolderName2.Text
                    'PdfPath = PdfSaveFolder + "\" + System.IO.Path.GetFileName(System.IO.Path.GetFileName(DcuPath))
                Else
                    'PlaySound("SystemHand", IntPtr.Zero, PlaySoundFlags.SND_ALIAS Or PlaySoundFlags.SND_NODEFAULT)
                    MsgBox("前はありません！", vbOK, "エラー")
                End If

            End If
        End If
    End Sub

    Private Sub NextFolderButton2_Click(sender As Object, e As EventArgs) Handles NextFolderButton2.Click

        If TextBox_FolderName2.Text <> "" Then
            Dim Path0 = System.IO.Path.GetDirectoryName(TextBox_FolderName2.Text)
            Dim dname = System.IO.Path.GetFileName(TextBox_FolderName2.Text)
            Dim dir2 As String(), dname2 As String(), dname3 As String()

            If Path0 <> "" And dname <> "" Then
                dir2 = Directory.GetDirectories(Path0, "*", SearchOption.TopDirectoryOnly)
                Dim n As Integer = dir2.Length
                ReDim dname2(n - 1), dname3(n - 1)
                For i As Integer = 0 To n - 1
                    dname2(i) = System.IO.Path.GetFileName(dir2(i))
                Next
                Dim cmp As StringComparer = StringComparer.OrdinalIgnoreCase
                Array.Sort(dname2, cmp)
                Dim nextfolder As String, index As Integer = -1
                Dim n2 As Integer = 0

                For i As Integer = 0 To n - 1
                    If dname2(i).Substring(0, 1) <> "." Then
                        dname3(n2) = dname2(i)
                        n2 += 1
                    End If
                Next


                For i As Integer = 0 To n2 - 1
                    If dname3(i) = dname Then
                        index = i
                        Exit For
                    End If
                Next
                If index >= 0 And index < n2 - 1 Then
                    nextfolder = dname3(index + 1)
                    TextBox_FolderName2.Text = Path0 + "\" + nextfolder
                    PdfPath = TextBox_FolderName2.Text
                    'PdfPath = PdfSaveFolder + "\" + System.IO.Path.GetFileName(System.IO.Path.GetFileName(DcuPath))
                Else
                    'PlaySound("SystemHand", IntPtr.Zero, PlaySoundFlags.SND_ALIAS Or PlaySoundFlags.SND_NODEFAULT)
                    MsgBox("次はありません！", vbOK, "エラー")
                End If

            End If
        End If
    End Sub

End Class

Class MeCab
    Implements IDisposable

    <DllImport("libmecab.dll", CallingConvention:=CallingConvention.Cdecl)>
    Public Shared Function mecab_new2(ByVal arg As String) As IntPtr
    End Function

    <DllImport("libmecab.dll", CallingConvention:=CallingConvention.Cdecl)>
    Public Shared Function mecab_sparse_tostr(ByVal m As IntPtr, ByVal str As String) As IntPtr
    End Function

    <DllImport("libmecab.dll", CallingConvention:=CallingConvention.Cdecl)>
    Public Shared Sub mecab_destroy(ByVal m As IntPtr)
    End Sub

    Private ptrMeCab As IntPtr

    Sub New()
        Me.New(String.Empty)
    End Sub

    Sub New(ByVal Arg As String)
        ptrMeCab = mecab_new2(Arg)
    End Sub

    Public Function Parse(ByVal [String] As String) As String
        Dim ptrResult As IntPtr = mecab_sparse_tostr(ptrMeCab, [String])
        Dim strResult As String = Marshal.PtrToStringAnsi(ptrResult)
        Return strResult
    End Function

    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        mecab_destroy(ptrMeCab)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        Dispose()
    End Sub

End Class
