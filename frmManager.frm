VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������"
   ClientHeight    =   6570
   ClientLeft      =   5475
   ClientTop       =   3405
   ClientWidth     =   15825
   BeginProperty Font 
      Name            =   "����"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   15825
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtRepeat 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox txtUsername 
      Height          =   375
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.ComboBox cmbUserType 
      Height          =   405
      ItemData        =   "frmManager.frx":0000
      Left            =   1680
      List            =   "frmManager.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox txtBalance 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Text            =   "0"
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "���/����"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ  ��"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "��  ��"
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame fraData 
      Caption         =   "�û��б�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   5280
      TabIndex        =   6
      Top             =   0
      Width           =   10455
      Begin MSDataGridLib.DataGrid grdData 
         Height          =   6135
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   10821
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
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ  ��"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblRepeat 
      Caption         =   "�ظ�����:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblUsername 
      Caption         =   "�� �� ��:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblPassword 
      Caption         =   "��    ��:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblUserType 
      Caption         =   "�û�����:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblBalance 
      Caption         =   "��    ��:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "frmManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'����ȫ�ֱ���CurUserType�Լ�CurUserTypeString
Private Sub GetUserType()
    CurUserTypeString = Trim(cmbUserType.Text)
    If CurUserTypeString = "�����û�" Then
        CurUserType = "U"
    ElseIf CurUserTypeString = "�̼�" Then
        CurUserType = "S"
    ElseIf CurUserTypeString = "����Ա" Then
        CurUserType = "A"
    End If
End Sub

Private Sub cmbUserType_Click()
    GetUserType
End Sub

Private Sub cmdAdd_Click()
    Dim sqlStr As String
    Dim Id As Integer
    Dim Username, Password As String
    Dim rs As ADODB.Recordset
    
    'ɾ��֮ǰ�е�
    cmdDelete_Click
    
    '����û����������Ƿ�Ϊ��
    Username = Trim(txtUsername.Text)
    If Username = "" Or txtPassword.Text = "" Then
        MsgBox "�������û���������", vbOKOnly + vbExclamation, "����"
        txtUsername.SetFocus
        Exit Sub
    End If
    
    '����Ƕ����û�����Ƿ����ֻ���
    If CurUserType = "U" And (Not IsNumeric(Username) Or Len(Username) <> 11) Then
        MsgBox "������Ϸ����ֻ�����15367893593", vbOKOnly + vbExclamation, "����"
        txtUsername.SetFocus
        Exit Sub
    End If
    
    '��������Ƿ�һ��
    If txtPassword.Text <> txtRepeat.Text Then
        MsgBox "���벻һ�£�", vbOKOnly + vbExclamation, "����"
        txtRepeat.SetFocus
        Exit Sub
    End If
    
    '����û�
    Password = MD5(txtPassword.Text, 32)
    sqlStr = "select max(id) from [user]"
    Set rs = SQLQRY(sqlStr)
    Id = rs.Fields(0).Value + 1
    sqlStr = "insert into [user] (id, username, password, type, balance) values (" & Id & _
        ", '" & Username & "', '" & Password & "', '" & CurUserType & "', " & Trim(txtBalance.Text) & ")"
    SQLDML (sqlStr)
    cmdRefresh_Click
End Sub

Private Sub cmdDelete_Click()
    Dim sqlStr As String
    
    sqlStr = "delete from [user] where username = '" & Trim(txtUsername.Text) & "' and type = '" & CurUserType & "'"
    SQLDML (sqlStr)
    cmdRefresh_Click
End Sub

Private Sub cmdExit_Click()
    Unload Me
    frmLogin.Show
End Sub

Private Sub cmdRefresh_Click()
    Dim rs As ADODB.Recordset
    Dim sqlStr As String
    
    sqlStr = "select username as �û���, password as ����, " & _
        "case type when 'A' then '����Ա' when 'U' then '�����û�' when 'S' then '�̼�' end as �û�����, " & _
        "balance as ��� from [user]"
    Set rs = SQLQRY(sqlStr)
    Set grdData.DataSource = rs
    AutoFitWidth grdData
End Sub

Private Sub Form_Load()
    cmbUserType.ListIndex = 0
    GetUserType
    cmdRefresh_Click
End Sub

Private Sub txtBalance_Change()
    If txtBalance.Text = "" Then
        txtBalance.Text = "0"
    End If
End Sub
