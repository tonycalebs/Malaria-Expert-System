VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00404000&
   Caption         =   "Form3"
   ClientHeight    =   10665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form3"
   ScaleHeight     =   10665
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "PROCEED TO LOGIN FORM<<<<<<<<<>>>>>>>>>>>>>>>>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   9720
      Width           =   20655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EXPERT SYSTEM FOR DIAGNOSING MALARIA AND TYPHOID FEVER USING RETE ALGORITHM"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   68.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   9615
      Left            =   2640
      TabIndex        =   0
      Top             =   600
      Width           =   14535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
frmLogin.Show
End Sub
