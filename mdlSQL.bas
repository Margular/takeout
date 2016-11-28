Attribute VB_Name = "mdlSQL"
Private DBConnectionString As String    '数据库连接串
Private conn As ADODB.Connection        '定义以ADO方式连接到数据库的对象变量
Private IsConnected As Boolean          '是否连接了数据库

'定义连接到数据库系统的全局函数
Public Function DBConnect(Optional tmpDBConnectionString As String = "") As Boolean
    '是否已经连接
    If IsConnected = True Then
        DBConnect = True
        Exit Function
    End If
    
    On Error GoTo sql_error
    
    Set conn = New ADODB.Connection
    '是否更新数据库连接串
    If tmpDBConnectionString <> "" Then
        DBConnectionString = tmpDBConnectionString
    End If
    
    conn.ConnectionString = DBConnectionString
    conn.Open
    IsConnected = True
    DBConnect = True
    Exit Function
sql_error:
    On Error Resume Next
    MsgBox "数据库连接失败：" & Err.Description, vbOKOnly + vbCritical, "错误"
    DBConnect = False
End Function

'断开与数据库的连接
Public Sub DBDisConnect()
    If IsConnected = False Then
        Exit Sub
    Else
        conn.Close
        Set conn = Nothing
        IsConnected = False
    End If
End Sub

'执行SQL-DML语句操作, 如INSERT、UPDATE、DELETE等，不需要返回结果集
Public Sub SQLDML(ByVal SQL_DMLStr As String)
    Dim cmd As New ADODB.Command
    
    On Error GoTo sql_error
    
    DBConnect
    Set cmd.ActiveConnection = conn     '设置cmd所关联的数据库
    cmd.CommandText = SQL_DMLStr
    cmd.Execute
sql_exit:
    On Error Resume Next
    Set cmd = Nothing
    DBDisConnect
    Exit Sub
sql_error:
    MsgBox "数据库更新操作失败：" & Err.Description, vbOKOnly + vbCritical, "错误"
    Resume sql_exit
End Sub

'执行SQL-SELECT语句操作, 并返回数据查询结果集
Public Function SQLQRY(ByVal SQL_QRYStr As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    DBConnect
    rs.Open SQL_QRYStr, conn, adOpenStatic, adLockOptimistic
    Set SQLQRY = rs
End Function
