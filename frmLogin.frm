VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   9810
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   19155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5796.072
   ScaleMode       =   0  'User
   ScaleWidth      =   17985.51
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   10170
      TabIndex        =   1
      Top             =   4575
      Width           =   3885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8175
      TabIndex        =   4
      Top             =   6540
      Width           =   1980
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   10860
      TabIndex        =   5
      Top             =   6540
      Width           =   2100
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      IMEMode         =   3  'DISABLE
      Left            =   10170
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5685
      Width           =   4005
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   25
      Height          =   3975
      Left            =   7560
      Top             =   3840
      Width           =   7095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   0
      Left            =   7905
      TabIndex        =   0
      Top             =   4590
      Width           =   2400
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   630
      Index           =   1
      Left            =   7905
      TabIndex        =   2
      Top             =   5460
      Width           =   2040
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOK_Click()
If txtUserName.Text = "user" And txtPassword.Text = "john" Then
    MsgBox "Welcome  Valid User"
                Me.Hide
                Form1.Show
    Else
        MsgBox "Invalid UserName or Password, try again!", , "Login"
        txtUserName.SetFocus
       End If
End Sub
