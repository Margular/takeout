Attribute VB_Name = "mdlCommon"
'定义全局变量
Public CurUserName As String        '当前登录用户名
Public CurUserType As String        '当前登录用户类型
Public CurUserTypeString As String  '当前登录用户类型描述

Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

'程序入口
Public Sub Main()
    Dim DBConnectionString As String

    DBConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;" & _
        "Initial Catalog=takeout;Data Source=localhost"
    If DBConnect(DBConnectionString) = False Then
        End
    Else
        frmLogin.Show
    End If
End Sub

'自动调整DataGrid列宽
Public Sub AutoFitWidth(ByRef grd As DataGrid)
    Dim tmprs As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim col, width As Integer
    
    Set tmprs = grd.DataSource
    If tmprs Is Nothing Then Exit Sub
    If tmprs.State = adStateClosed Then Exit Sub
    If tmprs.RecordCount = 0 Then Exit Sub
    
    Set rs = tmprs.Clone
    
    For col = 0 To grd.Columns.Count - 1
        width = lstrlen(grd.Columns(col).Caption)
        rs.MoveFirst
        Do Until rs.EOF
            If lstrlen(rs(col)) > width Then width = lstrlen(rs(col))
            rs.MoveNext
        Loop
        grd.Columns(col).width = width * 100
    Next
End Sub
