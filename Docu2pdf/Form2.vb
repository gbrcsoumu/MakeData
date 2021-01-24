Public Class Form2
    Public message As String
    Public title As String
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If title <> "" Then
            Me.Text = title
        End If

        If message <> "" Then
            Me.Label1.Text = message
        Else
            Me.Label1.Text = "ただいまファイルの検索中！" + vbCrLf + vbCrLf + vbCrLf + "しばらくお待ちください。"
        End If

        'Me.CenterToScreen()
        Me.CenterToParent()
    End Sub
End Class