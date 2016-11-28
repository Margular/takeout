VERSION 5.00
Begin VB.Form frmRegistry 
   Caption         =   "注册用户"
   ClientHeight    =   3420
   ClientLeft      =   4935
   ClientTop       =   2670
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   5160
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtUserType 
      Enabled         =   0   'False
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   9
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtRepeat 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "返  回"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdRegistry 
      Caption         =   "注  册"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtUsername 
      Height          =   390
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblRepeat 
      Caption         =   "重  复:"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblUserType 
      Caption         =   "类  型:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblPassword 
      Caption         =   "密  码:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblUsername 
      Caption         =   "用户名:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
    frmLogin.Show
End Sub

Private Sub cmdRegistry_Click()
    Dim sqlStr As String
    Dim Id As Integer
    Dim Username, Password As String
    Dim rs As ADODB.Recordset
    
    '检查用户名和密码是否为空
    Username = Trim(txtUsername.Text)
    If Username = "" Or txtPassword.Text = "" Then
        MsgBox "请输入用户名和密码", vbOKOnly + vbExclamation, "警告"
        txtUsername.SetFocus
        Exit Sub
    End If
    
    '如果是订餐用户检查是否是手机号
    If CurUserType = "U" And (Not IsNumeric(Username) Or Len(Username) <> 11) Then
        MsgBox "请输入合法的手机号如15367893593", vbOKOnly + vbExclamation, "警告"
        txtUsername.SetFocus
        Exit Sub
    End If
    
    '检查用户名是否存在
    sqlStr = "select * from [user] where username = '" & Username & "' and type = '" & CurUserType & "'"
    Set rs = SQLQRY(sqlStr)
    If Not rs.EOF Then
        MsgBox "用户名已存在！", vbOKOnly + vbInformation, "提示"
        txtUsername.SetFocus
        Exit Sub
    End If
    
    '检查密码是否一致
    If txtPassword.Text <> txtRepeat.Text Then
        MsgBox "密码不一致！", vbOKOnly + vbExclamation, "警告"
        txtRepeat.SetFocus
        Exit Sub
    End If
    
    '添加用户
    Password = MD5(txtPassword.Text, 32)
    sqlStr = "select max(id) from [user]"
    Set rs = SQLQRY(sqlStr)
    Id = rs.Fields(0).Value + 1
    
    sqlStr = "insert into [user] (id, username, password, type, balance) values (" & Id & _
        ", '" & Username & "', '" & Password & "', '" & CurUserType & "', 0)"
    SQLDML (sqlStr)
    MsgBox "注册成功！", vbOKOnly + vbInformation, "提示"
    Unload Me
    frmLogin.Show
End Sub

Private Sub Form_Load()
    txtUserType.Text = CurUserTypeString
End Sub

