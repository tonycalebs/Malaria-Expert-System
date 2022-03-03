VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18990
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   18990
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
      Height          =   8775
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   17415
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Height          =   7215
         Left            =   3240
         TabIndex        =   1
         Top             =   720
         Width           =   11175
         Begin VB.CommandButton Command6 
            Caption         =   "Get Card No.:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   3135
         End
         Begin VB.TextBox Text4 
            DataField       =   "Date_of_Reg"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            TabIndex        =   23
            Text            =   " "
            Top             =   6360
            Width           =   4095
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Get Tod ay's Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   22
            Top             =   6360
            Width           =   2775
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   8760
            Top             =   5880
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
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
            ForeColor       =   -2147483640
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
         Begin VB.CommandButton Command3 
            Caption         =   "PATIENT DIAGNOSIS"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   8160
            TabIndex        =   21
            Top             =   3120
            Width           =   2535
         End
         Begin VB.CommandButton Command2 
            Caption         =   "REFRESH"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   8160
            TabIndex        =   20
            Top             =   1920
            Width           =   2535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "SAVE PATIENT RECORD"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   8160
            TabIndex        =   19
            Top             =   600
            Width           =   2535
         End
         Begin VB.ComboBox Combo2 
            DataField       =   "Marital Status"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            ItemData        =   "Form1.frx":0000
            Left            =   3480
            List            =   "Form1.frx":000D
            TabIndex        =   18
            Text            =   "Select Marital Status"
            Top             =   2760
            Width           =   4095
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "Sex"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            ItemData        =   "Form1.frx":002A
            Left            =   3480
            List            =   "Form1.frx":0034
            TabIndex        =   17
            Text            =   "Select Gender"
            Top             =   1560
            Width           =   4095
         End
         Begin VB.TextBox Text10 
            DataField       =   "Reference_No"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            TabIndex        =   16
            Text            =   " "
            Top             =   5640
            Width           =   4095
         End
         Begin VB.TextBox Text8 
            DataField       =   "Reference_Name"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            TabIndex        =   15
            Text            =   " "
            Top             =   4800
            Width           =   4095
         End
         Begin VB.TextBox Text7 
            DataField       =   "Phone_No"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            TabIndex        =   14
            Text            =   " "
            Top             =   4080
            Width           =   4095
         End
         Begin VB.TextBox Text6 
            DataField       =   "Contact_Address"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            TabIndex        =   13
            Text            =   " "
            Top             =   3360
            Width           =   4095
         End
         Begin VB.TextBox Text3 
            DataField       =   "Age"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            TabIndex        =   12
            Text            =   " "
            Top             =   2160
            Width           =   4095
         End
         Begin VB.TextBox Text2 
            DataField       =   "Patient_Name"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            TabIndex        =   11
            Text            =   " "
            Top             =   960
            Width           =   4095
         End
         Begin VB.TextBox Text1 
            DataField       =   "Card_No"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            TabIndex        =   10
            Text            =   " "
            Top             =   360
            Width           =   4095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Reference No.:"
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
            Height          =   615
            Index           =   8
            Left            =   240
            TabIndex        =   9
            Top             =   5640
            Width           =   3015
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Reference Name:"
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
            Height          =   495
            Index           =   7
            Left            =   240
            TabIndex        =   8
            Top             =   4800
            Width           =   3135
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No.:"
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
            Height          =   615
            Index           =   6
            Left            =   240
            TabIndex        =   7
            Top             =   3960
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Address:"
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
            Height          =   615
            Index           =   5
            Left            =   240
            TabIndex        =   6
            Top             =   3360
            Width           =   3255
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Marital Status:"
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
            Height          =   615
            Index           =   4
            Left            =   240
            TabIndex        =   5
            Top             =   2760
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
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
            Height          =   615
            Index           =   3
            Left            =   240
            TabIndex        =   4
            Top             =   2160
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:"
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
            Height          =   615
            Index           =   2
            Left            =   240
            TabIndex        =   3
            Top             =   1560
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Name:"
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
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   960
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
MsgBox "Patient Record Saved Successfully", vbInformation, "RECORD HAS BEEN SAVED"
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text10.Text = ""
Combo1.Text = Empty
End Sub

Private Sub Command3_Click()
Me.Hide
Form2.Show
End Sub

Private Sub Command5_Click()
Text4.Text = Format$(Now, "dddd, mmmm dd, yyyy")
End Sub

Private Sub Command6_Click()
A = Int(1000 * Rnd) + 1000
Text1.Text = "Hosp_" & A
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
End Sub

