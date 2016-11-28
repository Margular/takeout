Attribute VB_Name = "mdlCommon"
'����ȫ�ֱ���
Public CurUserName As String        '��ǰ��¼�û���
Public CurUserType As String        '��ǰ��¼�û�����
Public CurUserTypeString As String  '��ǰ��¼�û���������

Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

'�������
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

'�Զ�����DataGrid�п�
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
