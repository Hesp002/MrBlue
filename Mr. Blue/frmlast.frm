VERSION 5.00
Begin VB.Form frmlast 
   Caption         =   "Congratulations!"
   ClientHeight    =   2175
   ClientLeft      =   3795
   ClientTop       =   2115
   ClientWidth     =   3720
   Icon            =   "frmlast.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   3720
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Black Chancery"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2160
      TabIndex        =   3
      Top             =   1560
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Black Chancery"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Play Again?"
      BeginProperty Font 
         Name            =   "Black Chancery"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Congratulations!"
      BeginProperty Font 
         Name            =   "Black Chancery"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3105
   End
End
Attribute VB_Name = "frmlast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BY: KYLE HESPE
Dim cindex As Integer
Dim valwin As Integer
Dim valmove As Integer
Dim valpowerup As Integer

Private Sub Form_Load()
Unload frmlvl7
End Sub




Private Sub Label3_Click()
Unload Me
frmmaze.Show 1
End Sub


Private Sub Label4_Click()
'Copywrite 2006 Hesp Corp.
MsgBox "Copywrite 2006 Hesp Corp.", vbCritical, "Copywrite 2006"
End
End Sub


