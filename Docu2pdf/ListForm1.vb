



Imports System.Net

Public Class ListForm1

    'Private filename() As String, fname() As String, dir1() As String
    'Private Path1 As String, Path2 As String
    'Private Check() As CheckBox, checkbox_n As Integer
    'Private Cansel As Boolean
    Private MyPath As String, MyName As String, hostname As String, adrList As IPAddress(), MyIP As String
    Private anser As String, _Kind As String


    Private Sub Cansel_Button1_Click(sender As Object, e As EventArgs) Handles Cancel_Button1.Click
        '
        '  Canselボタン処理
        '
        DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        MyPath = My.Application.Info.DirectoryPath
        MyName = My.Application.Info.AssemblyName
        ' ホスト名を取得する
        hostname = Dns.GetHostName()

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
            .Width = 870
            .Height = 220
            '.ColumnCount = 3
            'Col_n = .ColumnCount
            .ColumnHeadersVisible = True
            .ColumnHeadersHeight = 28
            .ScrollBars = ScrollBars.Both


            '.ColumnHeadersDefaultCellStyle = Style1
            '.Columns(0).Name = "IP"
            '.Columns(1).Name = "日時"
            '.Columns(2).Name = "Path"

            '.RowHeadersVisible = True
            '.Columns(0).Width = 60
            '.Columns(1).Width = 140
            '.Columns(2).Width = 500

            Dim ButtonColumn1 As New DataGridViewButtonColumn()
            '列の名前を設定
            ButtonColumn1.Name = "削除"
            '全てのボタンに"詳細閲覧"と表示する
            ButtonColumn1.UseColumnTextForButtonValue = True
            ButtonColumn1.Text = "削除"
            ButtonColumn1.Width = 50
            ButtonColumn1.DefaultCellStyle = Style1
            'DataGridViewに追加する
            .Columns.Add(ButtonColumn1)


            Dim textColumn1 As New DataGridViewTextBoxColumn()
            textColumn1.DataPropertyName = "IP"
            textColumn1.Name = "IP"
            textColumn1.HeaderText = "IP"
            textColumn1.Width = 100
            textColumn1.DefaultCellStyle = Style1
            .Columns.Add(textColumn1)

            Dim textColumn2 As New DataGridViewTextBoxColumn()
            textColumn2.DataPropertyName = "日時"
            textColumn2.Name = "日時"
            textColumn2.HeaderText = "日時"
            textColumn2.Width = 140
            textColumn2.DefaultCellStyle = Style1
            .Columns.Add(textColumn2)

            Dim textColumn3 As New DataGridViewTextBoxColumn()
            textColumn3.DataPropertyName = "Path"
            textColumn3.Name = "Path"
            textColumn3.HeaderText = "Path"
            textColumn3.Width = 480
            textColumn3.DefaultCellStyle = Style1
            .Columns.Add(textColumn3)

            Dim ButtonColumn2 As New DataGridViewButtonColumn()
            '列の名前を設定
            ButtonColumn2.Name = "選択"
            '全てのボタンに"詳細閲覧"と表示する
            ButtonColumn2.UseColumnTextForButtonValue = True
            ButtonColumn2.Text = "選択"
            ButtonColumn2.Width = 50
            ButtonColumn2.DefaultCellStyle = Style2
            'DataGridViewに追加する
            .Columns.Add(ButtonColumn2)

            Dim db As New OdbcDbIf
            Dim tb As DataTable
            Dim Sql_Command As String

            Dim Kind As String = DataKind

            'Dim td1 As DateTime = DateTime.Now
            'Dim td2 As String = td1.ToString().Replace("/", "-")

            FileMakerServer = FileMakerServer1
            db.Connect()

            Sql_Command = "SELECT ""IP"",""Path"",""接続日時"" FROM """ + Table3 + """ WHERE (""IP"" = '" + MyIP + "' AND ""種類"" = '" + Kind + "')"
            tb = db.ExecuteSql(Sql_Command)


            Dim n2 As Integer = tb.Rows.Count
            Dim IP As String, Path As String, td1 As String

            If n2 > 0 Then
                For i As Integer = 0 To n2 - 1
                    IP = tb.Rows(i).Item("IP").ToString()
                    Path = tb.Rows(i).Item("Path").ToString()
                    td1 = tb.Rows(i).Item("接続日時").ToString()
                    .Rows.Add()
                    .Rows(i).Height = 24
                    .Rows(i).Cells(1).Value = IP
                    .Rows(i).Cells(2).Value = td1
                    .Rows(i).Cells(3).Value = Path
                Next
            Else
                MsgBox("登録データがありません", vbOK, "警告")
                DialogResult = DialogResult.Cancel
                Me.Close()
            End If

            db.Disconnect()




        End With

        Me.CenterToScreen()
    End Sub

    'CellContentClickイベントハンドラ
    Private Sub DataGridView1_CellContentClick(ByVal sender As Object,
            ByVal e As DataGridViewCellEventArgs) _
            Handles DataGridView1.CellContentClick
        Dim dgv As DataGridView = CType(sender, DataGridView)
        '"Button"列ならば、ボタンがクリックされた
        If dgv.Columns(e.ColumnIndex).Name = "選択" Then
            'MessageBox.Show((e.RowIndex.ToString() +
            '    "行の選択ボタンがクリックされました。"))
            anser = DataGridView1.Rows(e.RowIndex).Cells(3).Value
            DialogResult = DialogResult.OK
            Me.Close()
        End If

        If dgv.Columns(e.ColumnIndex).Name = "削除" Then
            MessageBox.Show((e.RowIndex.ToString() +
                "行の削除ボタンがクリックされました。"))
        End If

    End Sub


    Public Function GetValue() As String
        Return anser
    End Function



End Class