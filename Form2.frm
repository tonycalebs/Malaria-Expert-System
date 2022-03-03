VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form2"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14040
   LinkTopic       =   "Form2"
   ScaleHeight     =   10620
   ScaleWidth      =   14040
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   10695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      Begin VB.CommandButton Command5 
         BackColor       =   &H8000000D&
         Caption         =   "REFRESH"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H8000000D&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   7200
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "LEARN ABOUT MALARIA/TYPHOID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         MaskColor       =   &H0000FF00&
         TabIndex        =   33
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000D&
         Caption         =   "back"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   8640
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6480
         TabIndex        =   31
         Top             =   9840
         Width           =   6135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "DIAGNOSE SYMPTOMS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1800
         TabIndex        =   30
         Top             =   9840
         Width           =   4695
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00000000&
         Height          =   1215
         Left            =   1800
         TabIndex        =   26
         Top             =   8640
         Width           =   10815
         Begin VB.OptionButton Option14 
            BackColor       =   &H00000000&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   5760
            TabIndex        =   28
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option13 
            BackColor       =   &H00000000&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3000
            TabIndex        =   27
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "7. Do You Experience Feverish Signs"
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
            Height          =   615
            Left            =   2280
            TabIndex        =   29
            Top             =   120
            Width           =   7095
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00000000&
         Height          =   1215
         Left            =   1800
         TabIndex        =   22
         Top             =   7320
         Width           =   10815
         Begin VB.OptionButton Option12 
            BackColor       =   &H00000000&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3000
            TabIndex        =   24
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option11 
            BackColor       =   &H00000000&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   5760
            TabIndex        =   23
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "6. Do You Experience Vomiting"
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
            Height          =   615
            Left            =   2400
            TabIndex        =   25
            Top             =   120
            Width           =   6735
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00000000&
         Height          =   1215
         Left            =   1800
         TabIndex        =   18
         Top             =   6000
         Width           =   10815
         Begin VB.OptionButton Option10 
            BackColor       =   &H00000000&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   5760
            TabIndex        =   20
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00000000&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3000
            TabIndex        =   19
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "5. Do You Experience Cough"
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
            Height          =   615
            Left            =   2520
            TabIndex        =   21
            Top             =   120
            Width           =   5775
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Height          =   1215
         Left            =   1800
         TabIndex        =   14
         Top             =   4680
         Width           =   10815
         Begin VB.OptionButton Option8 
            BackColor       =   &H00000000&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3000
            TabIndex        =   16
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00000000&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   5760
            TabIndex        =   15
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "4. Do You Experience Diarrhea or constipation"
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
            Height          =   615
            Left            =   1440
            TabIndex        =   17
            Top             =   120
            Width           =   9615
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Height          =   1215
         Left            =   1800
         TabIndex        =   10
         Top             =   3360
         Width           =   10815
         Begin VB.OptionButton Option6 
            BackColor       =   &H00000000&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   5760
            TabIndex        =   12
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00000000&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3000
            TabIndex        =   11
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "3. Do You Experience Headache"
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
            Height          =   615
            Left            =   2400
            TabIndex        =   13
            Top             =   120
            Width           =   6975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Height          =   1215
         Left            =   1800
         TabIndex        =   6
         Top             =   2040
         Width           =   10815
         Begin VB.OptionButton Option4 
            BackColor       =   &H00000000&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3000
            TabIndex        =   8
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00000000&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   5760
            TabIndex        =   7
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "2. Do You Experience  Stomach Pain"
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
            Height          =   615
            Left            =   1920
            TabIndex        =   9
            Top             =   120
            Width           =   6975
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Height          =   1215
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   10815
         Begin VB.OptionButton Option2 
            BackColor       =   &H00000000&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   5760
            TabIndex        =   5
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00000000&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3000
            TabIndex        =   4
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "1. Do You Experience  General Weakness of the Body"
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
            Height          =   615
            Left            =   480
            TabIndex        =   3
            Top             =   120
            Width           =   9975
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   570
         Left            =   12720
         Top             =   9000
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1005
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483635
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DIAGNOSIS_DB.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DIAGNOSIS_DB.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "DIAGNOSIS_TB"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4560
         TabIndex        =   1
         Text            =   "Select Patient Name"
         Top             =   120
         Width           =   6015
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
On Error Resume Next
If Combo1.Text <> "" Then
  With Adodc1.Recordset
    .MoveFirst
    .Find "Patient_Name='" & Combo1.Text & " ' "
    Exit Sub
   End With
   End If
End Sub

Private Sub Command1_Click()
If Option1.Value = True Then
Text1.Text = "YOU HAVE TYPHOID"
ElseIf Option4.Value = True Then
Text1.Text = "TYPHOID"
ElseIf Option5.Value = True Then
Text1.Text = "YOU HAVE MALARIA"
ElseIf Option7.Value = True Then
Text1.Text = "YOU DON'T HAVE MALARIA AND TYPHOID"
ElseIf Option8.Value = True Then
Text1.Text = "TYPHOID"
ElseIf Option9.Value = True Then
Text1.Text = "YOU HAVE MALARIA"
ElseIf Option12.Value = True Then
Text1.Text = "TYPHOID"
ElseIf Option13.Value = True Then
Text1.Text = "YOU HAVE MALARIA"
Else
Text1.Text = "YOU HAVE BOTH MALARIA AND TYPOID"
End If
End Sub

Private Sub Command2_Click()
Me.Hide
Form1.Show
End Sub

Private Sub Command4_Click()
Form2.PrintForm
End Sub

Private Sub Form_Load()
Adodc1.Refresh
With Adodc1.Recordset
Do Until .EOF
Combo1.AddItem ![Patient_Name]
.MoveNext
Loop
End With
End Sub
