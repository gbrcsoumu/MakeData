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

Public Class Form1
    Private filename() As String, fname() As String, dir1() As String
    Private Path1 As String, Path2 As String
    Private Check() As CheckBox, checkbox_n As Integer
    Private Cansel As Boolean
    Private MyPath As String, MyName As String, hostname As String, adrList As IPAddress()


    Private TextFileName As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '
        ' フォームの初期化
        '
        MyPath = My.Application.Info.DirectoryPath
        MyName = My.Application.Info.AssemblyName
        ' ホスト名を取得する
        hostname = Dns.GetHostName()

        ' ホスト名からIPアドレスを取得する
        adrList = Dns.GetHostAddresses(hostname)
        For Each address As IPAddress In adrList
            Console.WriteLine(address.ToString())
        Next

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

        TextBox_FileMakerServer.Text = FileMakerServer1
        Cansel = False

        'TextBox_FolderName1.Text = "\\192.168.0.173\disk1\報告書（耐火）＿業務課から\2000Ⅲ耐火防火試験室"
        TextBox_FolderName1.Text = "\\192.168.0.173\disk1\報告書（耐火）＿業務課から\test用"
        Path2 = TextBox_FolderName1.Text

        CheckBox_Input.Checked = True       ' 報告書の入力チェックボックス
        CheckBox_Convert.Checked = True     ' 報告書の変換チェックボックス
        CheckBox_Input2.Checked = True      ' 資料（スキャンデータ）の入力チェックボックス

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


    'Private Sub Text_Read_Button_Click(sender As Object, e As EventArgs)

    '    Dim FileMakerOn As Boolean = FileMakerCheckBox.Checked
    '    Dim db As New OdbcDbIf
    '    Dim tb As DataTable
    '    Dim Sql_Command As String

    '    If FileMakerOn = True Then
    '        FileMakerServer = TextBox_FileMakerServer.Text
    '        db.Connect()
    '    End If


    '    If TextFileName <> "" Then
    '        'OKボタンがクリックされたとき、選択されたファイル名を表示する
    '        Console.WriteLine(TextFileName)
    '        Dim sr As New StreamReader(TextFileName, System.Text.Encoding.GetEncoding("Shift_JIS"))

    '        Dim text As String = sr.ReadToEnd()

    '        sr.Close()
    '        'TextBox3.Text = ""
    '        Dim textLine() As String = Split(text, vbCrLf)
    '        'TextBox3.Text = text
    '        Dim n As Long = textLine.Length - 3
    '        Dim dir(20) As String
    '        Dim WildCard1() As String, WildCard2() As String
    '        If RadioButton_xdw.Checked = True Then
    '            ReDim WildCard1(1), WildCard2(1)
    '            WildCard1(0) = ".xdw"
    '            WildCard1(1) = ".xbd"
    '            WildCard2(0) = WildCard1(0).ToUpper
    '            WildCard2(1) = WildCard1(1).ToUpper
    '        Else
    '            ReDim WildCard1(0), WildCard2(0)
    '            WildCard1(0) = ".pdf"
    '            WildCard2(0) = WildCard1(0).ToUpper
    '        End If
    '        'WildCard2 = WildCard1.ToUpper

    '        For k As Integer = 0 To checkbox_n - 1
    '            If Check(k).Checked = True Then

    '                Dim Path1 As String = textLine(0)
    '                Dim Path2 As String
    '                Dim TabCount As Integer
    '                Dim name As String

    '                Dim s1 As String = Check(k).Text.ToUpper.Replace("3", "")
    '                Dim s2 As String = Check(k).Text.ToLower.Replace("3", "")
    '                Dim s3 As String = StrConv(s1, VbStrConv.Wide)
    '                Dim s4 As String = StrConv(s2, VbStrConv.Wide)
    '                Dim w1 As String = "\d\d\d\d\d\d"
    '                Dim w2 As String = "-\d\d-\d\d\d"
    '                Dim w3 As String = "-\d\d-\d\d"
    '                Dim w4 As String = "-\d\d-\d\d\d\d"

    '                For i As Integer = 3 To n + 2
    '                    If textLine(i) <> "" Then

    '                        TabCount = CountChar(textLine(i), vbTab)
    '                        name = textLine(i).Replace(vbTab, "")

    '                        If name.Substring(name.Length - 5, 5) = "(dir)" Then
    '                            dir(TabCount - 1) = name.Replace(" (dir)", "").Replace("・ ", "")
    '                        Else

    '                            Path2 = Path1
    '                            If TabCount > 1 Then
    '                                For j = 0 To TabCount - 2
    '                                    Path2 += "\" + dir(j)
    '                                Next
    '                            End If
    '                            Dim name2 As String = name.Replace(" (file)", "").Replace("・ ", "")
    '                            Path2 += "\" + name2

    '                            'zenkaku = StrConv("ﾊﾝｶｸﾉﾓｼﾞﾚﾂ", VbStrConv.Wide) '戻り値：ハンカクノモジレツ

    '                            If System.Text.RegularExpressions.Regex.IsMatch(name2, s1 + w1) Or System.Text.RegularExpressions.Regex.IsMatch(name2, s2 + w1) _
    '                                Or System.Text.RegularExpressions.Regex.IsMatch(name2, s1 + w2) Or System.Text.RegularExpressions.Regex.IsMatch(name2, s2 + w2) _
    '                                Or System.Text.RegularExpressions.Regex.IsMatch(name2, s1 + w3) Or System.Text.RegularExpressions.Regex.IsMatch(name2, s2 + w3) _
    '                                Or System.Text.RegularExpressions.Regex.IsMatch(name2, s1 + w4) Or System.Text.RegularExpressions.Regex.IsMatch(name2, s2 + w4) _
    '                                Or System.Text.RegularExpressions.Regex.IsMatch(name2, s3 + w1) Or System.Text.RegularExpressions.Regex.IsMatch(name2, s4 + w1) _
    '                                Or System.Text.RegularExpressions.Regex.IsMatch(name2, s3 + w2) Or System.Text.RegularExpressions.Regex.IsMatch(name2, s4 + w2) _
    '                                Or System.Text.RegularExpressions.Regex.IsMatch(name2, s3 + w3) Or System.Text.RegularExpressions.Regex.IsMatch(name2, s4 + w3) _
    '                                Or System.Text.RegularExpressions.Regex.IsMatch(name2, s3 + w4) Or System.Text.RegularExpressions.Regex.IsMatch(name2, s4 + w4) _
    '                                Then


    '                                For kk As Integer = 0 To WildCard1.Length - 1
    '                                    If System.Text.RegularExpressions.Regex.IsMatch(name2, WildCard1(kk)) Or System.Text.RegularExpressions.Regex.IsMatch(name2, WildCard2(kk)) Then
    '                                        Dim st1 As String = ""
    '                                        If FileMakerOn = True Then

    '                                            If Path2.Contains("'") Then Path2 = Path2.Replace("'", "''")
    '                                            If name2.Contains("'") Then name2 = name2.Replace("'", "''")
    '                                            'Dim name3 As String = name2.Replace(".xdw", ".pdf").Replace(".XDW", ".pdf")

    '                                            Sql_Command = "SELECT ""FilePath"",""PdfPath"" FROM """ + Table + """ WHERE (""ファイル名"" = '" & name2 & "')"
    '                                            tb = db.ExecuteSql(Sql_Command)
    '                                            Dim n2 As Integer = tb.Rows.Count

    '                                            If n2 > 0 Then
    '                                                st1 = "(済)"
    '                                            End If



    '                                        End If

    '                                        'TextBox3.Text += Path2 + st1 + vbCrLf
    '                                        'Me.TextBox3.SelectionStart = Me.TextBox3.Text.Length
    '                                        'Me.TextBox3.Focus()
    '                                        'Me.TextBox3.ScrollToCaret()







    '                                    End If
    '                                Next




    '                                'Dim RN = DataGridView1.Rows.Count - 2
    '                                'If RN >= 0 Then
    '                                '    For i As Integer = 0 To RN
    '                                '        DataGridView1.Rows.RemoveAt(0)
    '                                '    Next
    '                                'End If
    '                                'Dim row1() As String
    '                                'Dim _No As Integer, _X As Double, _Y As Double, _Z1 As Double, _Z2 As Double
    '                                'For i As Integer = 0 To n - 2
    '                                '    _No = Val(Data(i + 1, 0))
    '                                '    _X = Val(Data(i + 1, 1))
    '                                '    _Y = Val(Data(i + 1, 2))
    '                                '    _Z1 = Val(Data(i + 1, 3))
    '                                '    _Z2 = Val(Data(i + 1, 4))

    '                                '    loadAry(i) = New XYZData()
    '                                '    loadAry(i).No = _No
    '                                '    loadAry(i).X = _X
    '                                '    loadAry(i).Y = _Y
    '                                '    loadAry(i).Z1 = _Z1
    '                                '    loadAry(i).Z2 = _Z2

    '                                '    row1 = {_No.ToString, _X.ToString, _Y.ToString, _Z1.ToString, _Z2.ToString}
    '                                '    DataGridView1.Rows.Add(row1)

    '                                '    Dim columnHeaderStyle As New DataGridViewCellStyle()
    '                                '    columnHeaderStyle.BackColor = Color.White
    '                                '    columnHeaderStyle.Font = New Font("MSゴシック", 20, FontStyle.Bold)
    '                                '    DataGridView1.RowsDefaultCellStyle = columnHeaderStyle
    '                                '    '       R1 = R1 + 1
    '                                '    '       no = R1.ToString
    '                                '    '       row1 = {no, "", "", ""}
    '                                '    '       DataGridView1.Rows.Add(row1)

    '                                '    DataGridView1.Rows(i).Height = 30
    '                                '    DataGridView1.FirstDisplayedScrollingRowIndex = i
    '                                '    DataGridView1.CurrentCell = DataGridView1(0, i)
    '                                'Next
    '                                'loadAry2 = loadAry
    '                                'PointN = loadAry2.Length
    '                                'Me.EndPoint1.Text = PointN.ToString

    '                                'For i As Integer = 0 To loadAry.Length - 1
    '                                '    DataGridView1.Rows(i).Cells(5).Value = True
    '                                '    DataGridView1.Rows(i).Cells(6).Value = True
    '                                'Next

    '                            End If
    '                        End If
    '                    End If
    '                    'End If

    '                    Application.DoEvents()
    '                Next
    '            End If
    '        Next

    '        'TextBox3.Text += "== END ==" + vbCrLf
    '        'Me.TextBox3.SelectionStart = Me.TextBox3.Text.Length
    '        'Me.TextBox3.Focus()
    '        'Me.TextBox3.ScrollToCaret()
    '    End If

    '    If FileMakerOn = True Then
    '        db.Disconnect()
    '    End If


    'End Sub

    Private Function CountChar(ByVal s As String, ByVal c As Char) As Integer
        ' 文字列 s の中の文字 c の出現回数をカウントする関数
        Return s.Length - s.Replace(c.ToString(), "").Length
    End Function


    Private Sub DocuReadButton_Click(sender As Object, e As EventArgs) Handles DocuReadButton.Click
        '
        '　報告書（xdw,xbd）を読み込んでPDFに変換し、それをFileMakerに登録する。
        '

        Try

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
                .Height = 300
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


            'Application.DoEvents()

            Dim db As New OdbcDbIf
            Dim tb As DataTable
            Dim Sql_Command As String

            FileMakerServer = TextBox_FileMakerServer.Text
            db.Connect()

            Dim fname2 As New List(Of String)
            Dim dir2 As New List(Of String)
            Dim WildCard1() As String
            Dim Count As Integer = 0
            Dim ff()() As String, flag() As Boolean

            ReDim WildCard1(1), ff(1)
            WildCard1(0) = "*.xdw"
            WildCard1(1) = "*.xbd"


            Dim nn As Integer = 0

            For i As Integer = 0 To WildCard1.Length - 1
                ff(i) = System.IO.Directory.GetFiles(Path1, WildCard1(i), System.IO.SearchOption.AllDirectories)
                nn += ff(i).Length
            Next

            ReDim filename(nn - 1)

            For i As Integer = 0 To WildCard1.Length - 1
                If i = 0 Then
                    ff(i).CopyTo(filename, 0)
                Else
                    ff(i).CopyTo(filename, ff(i - 1).Length)
                End If
            Next

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
                                Dim f As String = System.IO.Path.GetFileNameWithoutExtension(filename(i))

                                Dim fname As String = System.IO.Path.GetFileName(filename(i))


                                If IsTestNumber(fname, Check(j).Text) Then   ' 試験番号（例えば、3A 3C）を含むファイルかどうかをチェック

                                    Count += 1
                                    Dim row1() As String
                                    row1 = {Count.ToString, fname, filename(i)}
                                    DataGridView1.Rows.Add(row1)

                                    Sql_Command = "SELECT ""FilePath"",""PdfPath"",""入力"" FROM """ + Table + """ WHERE (""ファイル名"" = '" & fname.Replace("'", "''") & "')"
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

            fname = fname2.ToArray
            dir1 = dir2.ToArray

            If Count = 0 Then
                MsgBox("このフォルダーには報告書ファイルはありません！", vbOK, "確認")
            Else
                MsgBox("このフォルダーには" + Count.ToString + "個の報告書ファイルがありました。", vbOK, "確認")
            End If

        Catch e1 As Exception
            'Console.WriteLine(e1.Message)
        End Try




    End Sub


    Private Sub PdfReadButton_Click(sender As Object, e As EventArgs) Handles PdfReadButton.Click
        '
        '　資料（pdf）を読み込んでFileMakerに登録する。
        '

        Try

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
                .Height = 300
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

            Dim fname2 As New List(Of String)
            Dim dir2 As New List(Of String)
            Dim WildCard1() As String

            Dim Count As Integer = 0
            Dim ff()() As String, flag() As Boolean

            ReDim WildCard1(0), ff(0)
            WildCard1(0) = "*.pdf"

            Dim nn As Integer = 0

            For i As Integer = 0 To WildCard1.Length - 1
                ff(i) = System.IO.Directory.GetFiles(Path2, WildCard1(i), System.IO.SearchOption.AllDirectories)
                nn += ff(i).Length
            Next

            ReDim filename(nn - 1)

            For i As Integer = 0 To WildCard1.Length - 1
                If i = 0 Then
                    ff(i).CopyTo(filename, 0)
                Else
                    ff(i).CopyTo(filename, ff(i - 1).Length)
                End If
            Next

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

            fname = fname2.ToArray
            dir1 = dir2.ToArray

            If Count = 0 Then
                MsgBox("このフォルダーには資料ファイルはありません！", vbOK, "確認")
            Else
                MsgBox("このフォルダーには" + Count.ToString + "個の資料ファイルがありました。", vbOK, "確認")
            End If


        Catch e1 As Exception
            'Console.WriteLine(e1.Message)
        End Try


    End Sub


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
        If TextBox3.Text <> "" Then
            fbd.SelectedPath = TextBox3.Text
        Else
            fbd.SelectedPath = "\\192.168.0.173\disk1\SCAN\test"
        End If

        'ユーザーが新しいフォルダを作成できるようにする
        'デフォルトでTrue
        fbd.ShowNewFolderButton = True

        'ダイアログを表示する
        If fbd.ShowDialog(Me) = DialogResult.OK Then
            '選択されたフォルダを表示する
            Me.TextBox3.Text = fbd.SelectedPath
            Path2 = fbd.SelectedPath

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
            fbd.SelectedPath = "\\192.168.0.173\disk1\報告書（耐火）＿業務課から"
        End If

        'ユーザーが新しいフォルダを作成できるようにする
        'デフォルトでTrue
        fbd.ShowNewFolderButton = True

        'ダイアログを表示する
        If fbd.ShowDialog(Me) = DialogResult.OK Then
            '選択されたフォルダを表示する
            Me.TextBox_FolderName1.Text = fbd.SelectedPath
            Path1 = fbd.SelectedPath

            Path2 = PdfSaveFolder + "\" + System.IO.Path.GetFileName(System.IO.Path.GetFileName(Path1))
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
                    TextBox_FileLIst2.Text = ""
                    ProgressBar1.Minimum = 0
                    ProgressBar1.Maximum = Count
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
                                Sql_Command = "INSERT INTO """ + Table + """ (""FilePath"",""ファイル名"",""入力"")"
                                Sql_Command += " VALUES ('" + filename1.Replace("'", "''") + "','" + fname.Replace("'", "''") + "','未読')"
                                tb = db.ExecuteSql(Sql_Command)
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

                    End Try

                End If
            Else
                MessageBox.Show("未入力のデータはありません", "警告", MessageBoxButtons.OK)
            End If



        Else
            MessageBox.Show("データがありません!!", "警告", MessageBoxButtons.OK)

        End If
    End Sub


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
                    ProgressBar1.Minimum = 0
                    ProgressBar1.Maximum = Count
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

                                Me.TextBox_FileLIst2.Text += PdfPath + AddText + vbCrLf
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
                        db.Disconnect()

                        ' Shift-Jisでファイルを作成


                        '２行書き込み


                        'ストリームを閉じる
                        'sw.Close()
                        ProgressBar1.Visible = False

                    Catch e1 As Exception

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

                    ProgressBar2.Minimum = 0
                    ProgressBar2.Maximum = Count
                    ProgressBar2.Visible = True

                    TextBox_FileLIst3.Text = ""
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
                            If DataGridView2.Rows(i).Cells(4).Value = True Then

                                Dim fname As String = DataGridView2.Rows(i).Cells(1).Value
                                Dim filename1 As String = DataGridView2.Rows(i).Cells(2).Value
                                Sql_Command = "INSERT INTO """ + Table + """ (""FilePath"",""ファイル名"",""入力"")"
                                Sql_Command += " VALUES ('" + filename1.Replace("'", "''") + "','" + fname.Replace("'", "''") + "','未読')"
                                tb = db.ExecuteSql(Sql_Command)
                                DataGridView2.Rows(i).Cells(3).Value = "済"
                                DataGridView2.Rows(i).Cells(4).Value = False

                                sw.WriteLine(fname)
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
                        sw.Close()

                        ProgressBar2.Visible = False

                    Catch e1 As Exception

                    End Try

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

End Class
