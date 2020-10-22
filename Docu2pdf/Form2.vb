Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Label1.Text = "ただいまファイルの検索中！" + vbCrLf + vbCrLf + vbCrLf + "しばらくお待ちください。"
        Me.CenterToScreen()
    End Sub
End Class