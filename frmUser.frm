VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����û�����"
   ClientHeight    =   8460
   ClientLeft      =   4260
   ClientTop       =   1185
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   15645
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox cmbScore 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmUser.frx":0000
      Left            =   1680
      List            =   "frmUser.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   7200
      Width           =   3975
   End
   Begin VB.ComboBox cmbMethod 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmUser.frx":0046
      Left            =   1680
      List            =   "frmUser.frx":0050
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox txtHistoryId 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   16
      Top             =   6720
      Width           =   3975
   End
   Begin VB.Frame fraHistory 
      Caption         =   "��ʷ��¼"
      Height          =   3015
      Left            =   240
      TabIndex        =   15
      Top             =   3480
      Width           =   5415
      Begin MSDataGridLib.DataGrid grdHistory 
         Height          =   2655
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4683
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdScore 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   7800
      Width           =   1575
   End
   Begin VB.TextBox txtBalance 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Text            =   "0"
      Top             =   2280
      Width           =   3975
   End
   Begin VB.ComboBox cmbType 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmUser.frx":0060
      Left            =   1680
      List            =   "frmUser.frx":0076
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Frame fraData 
      Caption         =   "��Ʒ�б�"
      Height          =   8175
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   9615
      Begin MSDataGridLib.DataGrid grdData 
         Height          =   7815
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   13785
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtMenuId 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.TextBox txtTelephone 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblMethod 
      Caption         =   "��ͷ�ʽ:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblHistoryId 
      Caption         =   "�� ¼ ID:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label lblBalance 
      Caption         =   "��    ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblType 
      Caption         =   "��    ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblMenuId 
      Caption         =   "�� �� ID:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblTelephone 
      Caption         =   "�� �� ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblScore 
      Caption         =   "��    ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   7320
      Width           =   1335
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim sqlStr As String
    Dim rs As ADODB.Recordset
    Dim Id As Integer
    
    'ɾ��֮ǰ�е�
    cmdDelete_Click
    GetUserType
    
    '��ȡid
    sqlStr = "select max(id) from [user]"
    Set rs = SQLQRY(sqlStr)

    If IsNull(rs.Fields(0).Value) Then
        Id = 1
    Else
        Id = rs.Fields(0).Value + 1
    End If

    '���
    sqlStr = "insert into [user] (id, username, password, type, balance) values (" & Id & _
        ", '" & Trim(txtUsername.Text) & "', '" & MD5(Trim(txtPassword.Text), 32) & "', '" & CurUserType & "', " & _
        Trim(txtBalance.Text) & ")"

    SQLDML (sqlStr)
    cmdRefresh_Click
End Sub

Private Sub cmdDelete_Click()
    Dim sqlStr As String
    
    sqlStr = "delete from [user] where username = '" & Trim(txtUsername.Text) & "'"
    SQLDML (sqlStr)
    cmdRefresh_Click
End Sub

Private Sub cmbType_Click()
    cmdRefresh_Click
End Sub

Private Sub cmdExit_Click()
    Unload Me
    frmLogin.Show
End Sub

Private Sub cmdOrder_Click()
    Dim sqlStr As String
    Dim Id As Integer
    Dim Price As Double
    Dim rs As ADODB.Recordset
    
    '��ȡid
    sqlStr = "select max(id) from history"
    Set rs = SQLQRY(sqlStr)

    If IsNull(rs.Fields(0).Value) Then
        Id = 1
    Else
        Id = rs.Fields(0).Value + 1
    End If
    
    sqlStr = "insert into history (id, telephone, menu_id, method, score, datetime) values (" & Id & ", " & _
        Trim(txtTelephone.Text) & ", '" & Trim(txtMenuId.Text) & "', '" & cmbMethod.Text & "', 0, '" & _
        Format(Now, "YYYY-MM-DD hh:mm:ss") & "')"
    SQLDML (sqlStr)
    '��Ʒ��������һ
    sqlStr = "update menu set total = total + 1 where id = " & Trim(txtMenuId.Text)
    
    '��ü۸�
    sqlStr = "select price from menu where id = " & Trim(txtMenuId.Text)
    Set rs = SQLQRY(sqlStr)
    Price = rs.Fields(0).Value
    
    '�������
    sqlStr = "update [user] set balance = balance - " & Price & " where username = '" & CurUserName & "'"
    SQLDML (sqlStr)
    cmdRefresh_Click
End Sub

Private Sub cmdRefresh_Click()
    Dim rs, rsBalance, rsHistory As ADODB.Recordset
    Dim sqlStr As String
    
    '��ʾ����
    If cmbType.Text = "����" Then
        sqlStr = "select id as ��ID, seller_name as ����, name as ����, price as �۸�, type as ����, list as ����, score as ����, " & _
            "count as ��������, total as �������� from menu"
    Else
        sqlStr = "select id as ��ID, seller_name as ����, name as ����, price as �۸�, type as ����, list as ����, score as ����, " & _
            "count as ��������, total as �������� from menu where type = '" & cmbType.Text & "'"
    End If
    
    Set rs = SQLQRY(sqlStr)
    Set grdData.DataSource = rs
    
    '�������
    sqlStr = "select balance from [user] where username = '" & CurUserName & "'"
    Set rsBalance = SQLQRY(sqlStr)
    txtBalance.Text = rsBalance.Fields(0).Value
    
    '������ʷ��¼
    sqlStr = "select id as ��ʷ��¼ID, menu_id as ��ID, method as ��ͷ�ʽ, score as ����, datetime as ����ʱ�� from history where telephone = " & _
        Trim(txtTelephone.Text)
    Set rsHistory = SQLQRY(sqlStr)
    Set grdHistory.DataSource = rsHistory
    AutoFitWidth grdHistory
    AutoFitWidth grdData
End Sub

Private Sub cmdScore_Click()
    Dim sqlStr As String
    Dim rs As ADODB.Recordset
    Dim Score As Double
    Dim MenuId As Integer
    
    '��ʷ��¼�Ƿ�Ϊ��
    If txtHistoryId.Text = "" Then
        MsgBox "��������ʷ��¼ID", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    sqlStr = "select score from history where id = " & txtHistoryId.Text & " and telephone = " & txtTelephone.Text
    Set rs = SQLQRY(sqlStr)
    '��ʷ��¼ID���
    If rs.EOF Then
        MsgBox "��������ȷ����ʷ��¼ID", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    Score = rs.Fields(0).Value
    '�Ƿ��Ѿ������
    If Score <> 0 Then
        MsgBox "�벻Ҫ�ظ����!", vbOKOnly + vbExclamation, "��ʾ"
        Exit Sub
    End If
    '���
    sqlStr = "update history set score = " & cmbScore.Text & " where id = " & txtHistoryId.Text
    SQLDML (sqlStr)
    Score = cmbScore.Text
    
    '���²�Ʒ����ͳ��
    sqlStr = "select menu_id from history where id = " & txtHistoryId.Text
    Set rs = SQLQRY(sqlStr)
    MenuId = rs.Fields(0).Value
    
    sqlStr = "update menu set score = (score * count + " & Score & ") / (count + 1), count = count + 1 " & _
        "where id = " & MenuId
    SQLDML (sqlStr)
    cmdRefresh_Click
End Sub

Private Sub Form_Load()
    txtTelephone.Text = CurUserName
    cmbType.ListIndex = 0
    cmbMethod.ListIndex = 0
    cmbScore.ListIndex = cmbScore.ListCount - 1
    cmdRefresh_Click
End Sub

Private Sub txtBalance_Change()
    If txtBalance.Text = "" Then
        txtBalance.Text = "0"
    End If
End Sub
