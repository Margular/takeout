Attribute VB_Name = "mdlSQL"
Private DBConnectionString As String    '���ݿ����Ӵ�
Private conn As ADODB.Connection        '������ADO��ʽ���ӵ����ݿ�Ķ������
Private IsConnected As Boolean          '�Ƿ����������ݿ�

'�������ӵ����ݿ�ϵͳ��ȫ�ֺ���
Public Function DBConnect(Optional tmpDBConnectionString As String = "") As Boolean
    '�Ƿ��Ѿ�����
    If IsConnected = True Then
        DBConnect = True
        Exit Function
    End If
    
    On Error GoTo sql_error
    
    Set conn = New ADODB.Connection
    '�Ƿ�������ݿ����Ӵ�
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
    MsgBox "���ݿ�����ʧ�ܣ�" & Err.Description, vbOKOnly + vbCritical, "����"
    DBConnect = False
End Function

'�Ͽ������ݿ������
Public Sub DBDisConnect()
    If IsConnected = False Then
        Exit Sub
    Else
        conn.Close
        Set conn = Nothing
        IsConnected = False
    End If
End Sub

'ִ��SQL-DML������, ��INSERT��UPDATE��DELETE�ȣ�����Ҫ���ؽ����
Public Sub SQLDML(ByVal SQL_DMLStr As String)
    Dim cmd As New ADODB.Command
    
    On Error GoTo sql_error
    
    DBConnect
    Set cmd.ActiveConnection = conn     '����cmd�����������ݿ�
    cmd.CommandText = SQL_DMLStr
    cmd.Execute
sql_exit:
    On Error Resume Next
    Set cmd = Nothing
    DBDisConnect
    Exit Sub
sql_error:
    MsgBox "���ݿ���²���ʧ�ܣ�" & Err.Description, vbOKOnly + vbCritical, "����"
    Resume sql_exit
End Sub

'ִ��SQL-SELECT������, ���������ݲ�ѯ�����
Public Function SQLQRY(ByVal SQL_QRYStr As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    DBConnect
    rs.Open SQL_QRYStr, conn, adOpenStatic, adLockOptimistic
    Set SQLQRY = rs
End Function
