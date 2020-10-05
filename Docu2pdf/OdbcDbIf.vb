Imports System.Data.Odbc


Public Class OdbcDbIf
    ''' 
    ''' SQLコネクション
    ''' 
    Private _con As OdbcConnection = Nothing

    ''' 
    ''' トランザクション・オブジェクト
    ''' 
    ''' 
    Private _trn As OdbcTransaction = Nothing

    ''' 
    ''' DB接続
    ''' 
    ''' データソース名
    ''' データベース名
    ''' ユーザーID
    ''' パスワード
    ''' タイムアウト値
    ''' 
    Public Sub Connect(Optional ByVal dbn As String = "ファイル情報",
                       Optional ByVal uid As String = "Admin",
                       Optional ByVal pas As String = "taika1",
                       Optional ByVal tot As Integer = -1)

        Try
            If _con Is Nothing Then
                _con = New OdbcConnection
            End If

            Dim cst As String = ""
            'cst = cst & ";DSN=" & dsn
            cst = cst & ";DRIVER={FileMaker ODBC};HST=" + FileMakerServer + ";PRT=2399"
            cst = cst & ";Database=" & dbn
            cst = cst & ";UID=" & uid
            cst = cst & ";PWD=" & pas
            If tot > -1 Then
                '_con.ConnectionTimeout = tot
                cst = cst & ";Connect Timeout=" & tot.ToString
            End If
            'cst = "DRIVER={FileMaker ODBC};UID=Admin;PWD=taifu1;HST=192.168.33.250;PRT=2399;database=見積管理;"
            _con.ConnectionString = cst

            _con.Open()
        Catch ex As Exception
            Throw New Exception("Connect Error", ex)
        End Try
    End Sub

    ''' 
    ''' DB切断
    ''' 
    Public Sub Disconnect()
        Try
            If _con IsNot Nothing Then
                _con.Close()
            End If
        Catch ex As Exception
            Throw New Exception("Disconnect Error", ex)
        End Try
    End Sub

    ''' 
    ''' SQLの実行
    ''' 
    ''' SQL文
    ''' タイムアウト値
    ''' 
    ''' 
    Public Function ExecuteSql(ByVal sql As String, _
                               Optional ByVal tot As Integer = -1) As DataTable
        Dim dt As New DataTable

        Try
            Dim sqlCommand As New OdbcCommand(sql, _con, _trn)

            If tot > -1 Then
                sqlCommand.CommandTimeout = tot
            End If

            Dim adapter As New OdbcDataAdapter(sqlCommand)

            adapter.Fill(dt)
            adapter.Dispose()
            sqlCommand.Dispose()
        Catch ex As Exception
            Throw New Exception("ExecuteSql Error", ex)
        End Try

        Return dt
    End Function

    ''' 
    ''' トランザクション開始
    ''' 
    ''' 
    Public Sub BeginTransaction()
        Try
            _trn = _con.BeginTransaction()
        Catch ex As Exception
            Throw New Exception("BeginTransaction Error", ex)
        End Try
    End Sub

    ''' 
    ''' コミット
    ''' 
    ''' 
    Public Sub CommitTransaction()
        Try
            If _trn Is Nothing = False Then
                _trn.Commit()
            End If
        Catch ex As Exception
            Throw New Exception("CommitTransaction Error", ex)
        Finally
            _trn = Nothing
        End Try
    End Sub

    ''' 
    ''' ロールバック
    ''' 
    ''' 
    Public Sub RollbackTransaction()
        Try
            If _trn Is Nothing = False Then
                _trn.Rollback()
            End If
        Catch ex As Exception
            Throw New Exception("RollbackTransaction Error", ex)
        Finally
            _trn = Nothing
        End Try
    End Sub

    ''' 
    ''' ファイナライズ
    ''' 
    ''' 
    Protected Overrides Sub Finalize()
        Disconnect()
        MyBase.Finalize()
    End Sub
End Class
