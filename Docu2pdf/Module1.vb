Module Module1
    Public Const FileMakerServer1 As String = "192.168.37.228"
    Public FileMakerServer As String
    Public Const Table = "ファイル情報"
    Public Const Table2 = "PC情報"
    Public Const Table3 = "フォルダー情報1"
    Public Const Table4 = "フォルダー情報2"
    Public Const PdfSaveFolder = "\\192.168.0.173\disk1\報告書（耐火＿PDF）"
    Public Const CmdFile = "C:\CMD\cmd.txt"
    Public DataKind As String
    Public Ndrive As String(,) = {{"W:\", "\\192.168.37.242\fire\"},
                                  {"X:\", "\\192.168.37.240\fire\"},
                                  {"V:\", "\\192.168.37.241\fire\"},
                                  {"Y:\", "\\192.168.0.173\disk1\"}}

    'Public Const PdfSaveFolder = "\\192.168.32.90\Win共有\PDF"

    ' 文字の出現回数をカウント
    Public Function CountChar(ByVal s As String, ByVal c As Char) As Integer
        Return s.Length - s.Replace(c.ToString(), "").Length
    End Function

End Module
