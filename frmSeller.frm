VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSeller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�̼ҹ���"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   15930
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtSellerName 
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
      TabIndex        =   14
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ  ��"
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
      Left            =   1680
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Frame fraData 
      Caption         =   "�ҵĲ���"
      Height          =   5775
      Left            =   6000
      TabIndex        =   11
      Top             =   120
      Width           =   9855
      Begin MSDataGridLib.DataGrid grdData 
         Height          =   5415
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   9551
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
      Left            =   4560
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
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
      Left            =   3120
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "���/����"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtList 
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
      TabIndex        =   6
      Top             =   2160
      Width           =   4215
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
      ItemData        =   "frmSeller.frx":0000
      Left            =   1680
      List            =   "frmSeller.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txtPrice 
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
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox txtName 
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
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label lblSellerName 
      Caption         =   "�̼�����:"
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
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblList 
      Caption         =   "�ײͲ���:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblType 
      Caption         =   "�ײ�����:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblPrice 
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
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblName 
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
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmSeller"
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
    
    '��ȡid
    sqlStr = "select max(id) from menu"
    Set rs = SQLQRY(sqlStr)

    If IsNull(rs.Fields(0).Value) Then
        Id = 1
    Else
        Id = rs.Fields(0).Value + 1
    End If

    '���
    sqlStr = "insert into menu (id, seller_name, name, price, type, list, score, count, total) values (" & _
        Id & ", '" & CurUserName & "', '" & Trim(txtName.Text) & "', " & Trim(txtPrice.Text) & ", '" & cmbType.Text & "', '" & _
        Trim(txtList.Text) & "', 0, 0, 0)"
    
    SQLDML (sqlStr)
    cmdRefresh_Click
End Sub

Private Sub cmdDelete_Click()
    Dim sqlStr As String
    
    sqlStr = "delete from menu where seller_name = '" & CurUserName & "' and name = '" & Trim(txtName.Text) & "'"
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
    
    sqlStr = "select name as ����, price as �۸�, type as ����, list as ����, score as ����, " & _
        "count as ��������, total as �������� from menu where seller_name = '" & CurUserName & "'"
    Set rs = SQLQRY(sqlStr)
    Set grdData.DataSource = rs
    AutoFitWidth grdData
End Sub

Private Sub Form_Load()
    cmbType.ListIndex = 0
    txtSellerName.Text = CurUserName
    cmdRefresh_Click
End Sub

