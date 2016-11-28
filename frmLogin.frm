VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ô��¼����----Made by Margular"
   ClientHeight    =   3135
   ClientLeft      =   10095
   ClientTop       =   5220
   ClientWidth     =   5430
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
   ScaleHeight     =   3135
   ScaleWidth      =   5430
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdRegistry 
      Caption         =   "ע  ��"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox cmbUserType 
      Height          =   405
      ItemData        =   "frmLogin.frx":0000
      Left            =   1560
      List            =   "frmLogin.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "��  ��"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "��  ¼"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtUsername 
      Height          =   345
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblUserType 
      Caption         =   "��  ��:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblPassword 
      Caption         =   "��  ��:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblUsername 
      Caption         =   "�û���:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdExit_Click()
    Unload Me
End Sub

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

Private Sub cmdLogin_Click()
    Dim sqlStr As String
    Dim Username, Password As String
    Dim rs As ADODB.Recordset
    
    Username = Trim(txtUsername.Text)
    Password = MD5(txtPassword.Text, 32)
    GetUserType
    
    '����Ƿ������û���������
    If Username = "" Or txtPassword.Text = "" Then
        MsgBox "�������û���������", vbOKOnly + vbExclamation, "����"
        txtUsername.SetFocus
        Exit Sub
    End If
    
    '����Ƿ�����û���
    sqlStr = "select * from [user] where username = '" & Username & "' and type = '" & CurUserType & "'"
    Set rs = SQLQRY(sqlStr)
    If rs.EOF Then
        MsgBox "�û���������", vbOKOnly + vbExclamation, "����"
        txtUsername.SetFocus
        Exit Sub
    End If
    
    '��������Ƿ���ȷ
    sqlStr = "select * from [user] where username = '" & Username & "' and password = '" & Password & _
            "' and type = '" & CurUserType & "'"
    Set rs = SQLQRY(sqlStr)
        
    If rs.EOF Then
        MsgBox "�������", vbOKOnly + vbInformation, "��ʾ"
    Else
        Unload Me
        CurUserName = Username
        If CurUserType = "U" Then
            frmUser.Show
        ElseIf CurUserType = "S" Then
            frmSeller.Show
        ElseIf CurUserType = "A" Then
            frmManager.Show
        End If
    End If
End Sub

Private Sub cmdRegistry_Click()
    GetUserType
    If CurUserType = "A" Then
        MsgBox "������ע�����Ա", vbOKOnly + vbCritical, "����"
        Exit Sub
    End If
    
    Unload Me
    frmRegistry.Show
End Sub

Private Sub Form_Load()
    cmbUserType.ListIndex = 0
End Sub
