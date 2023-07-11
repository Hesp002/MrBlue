VERSION 5.00
Begin VB.Form frmmaze 
   Caption         =   "Level 1"
   ClientHeight    =   6375
   ClientLeft      =   3315
   ClientTop       =   2415
   ClientWidth     =   6285
   Icon            =   "mrblue.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   6285
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "`"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Black Chancery"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   810
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   35
      Left            =   4440
      Top             =   4560
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   34
      Left            =   4440
      Top             =   3960
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   33
      Left            =   4440
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   32
      Left            =   4440
      Top             =   2760
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   31
      Left            =   4440
      Top             =   2160
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   30
      Left            =   4440
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   29
      Left            =   3840
      Top             =   4560
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   28
      Left            =   3840
      Top             =   3960
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   27
      Left            =   3840
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   26
      Left            =   3840
      Top             =   2760
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   25
      Left            =   3840
      Top             =   2160
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   24
      Left            =   3840
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   23
      Left            =   3240
      Top             =   4560
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   22
      Left            =   3240
      Top             =   3960
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   21
      Left            =   3240
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   20
      Left            =   3240
      Top             =   2760
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   19
      Left            =   3240
      Top             =   2160
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   18
      Left            =   3240
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   17
      Left            =   2640
      Top             =   4560
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   16
      Left            =   2640
      Top             =   3960
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   15
      Left            =   2640
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   14
      Left            =   2640
      Top             =   2760
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   13
      Left            =   2640
      Top             =   2160
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   12
      Left            =   2640
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000001&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   11
      Left            =   2040
      Top             =   4560
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   10
      Left            =   2040
      Top             =   3960
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   9
      Left            =   2040
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   8
      Left            =   2040
      Top             =   2760
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   7
      Left            =   2040
      Top             =   2160
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   6
      Left            =   2040
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   5
      Left            =   1440
      Top             =   4560
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   4
      Left            =   1440
      Top             =   3960
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   3
      Left            =   1440
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   1440
      Top             =   2760
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   1440
      Top             =   2160
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   1440
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "frmmaze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BY: KYLE HESPE
'cindex holds the value of Mr. Blue. checks to see if there are walls or is a winning spot. If neither of it is true, allows to move to another space and changes cindex
Dim cindex As Integer
Dim valwin As Integer
Dim valmove As Integer
Dim valpowerup As Integer


Private Sub Command1_Click()
frmlvl2.Show 1
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
For i = 0 To 35
    If Shape1(i).FillColor = vbBlue Then
        cindex = i
    End If
    'sets cindex
Next i
'gets keypress and moves color
If KeyCode = vbKeyUp Then
    If cindex = 0 Or cindex = 6 Or cindex = 12 Or cindex = 18 Or cindex = 24 Or cindex = 30 Then Exit Sub
    If Shape1(cindex - 1).FillColor = vbBlack Then Exit Sub
    If Shape1(cindex - 1).FillColor = vbGreen Then
        MsgBox "You won!", vbInformation, "Winner!"
        MsgBox "Next Level", vbInformation, "Winner!"
        valwin = 1
    End If
    cindex = cindex - 1
    Shape1(cindex).FillColor = vbBlue
    Shape1(cindex + 1).FillColor = vbWhite
ElseIf KeyCode = vbKeyDown Then
    If cindex = 5 Or cindex = 11 Or cindex = 17 Or cindex = 23 Or cindex = 29 Or cindex = 35 Then Exit Sub
    If Shape1(cindex + 1).FillColor = vbBlack Then Exit Sub
    If Shape1(cindex + 1).FillColor = vbGreen Then
        MsgBox "You won!", vbInformation, "Winner!"
        MsgBox "Next Level", vbInformation, "Winner!"
        valwin = 1
    End If
    cindex = cindex + 1
    Shape1(cindex).FillColor = vbBlue
    Shape1(cindex - 1).FillColor = vbWhite
ElseIf KeyCode = vbKeyRight Then
    If cindex = 30 Or cindex = 31 Or cindex = 32 Or cindex = 33 Or cindex = 34 Or cindex = 35 Then Exit Sub
    If Shape1(cindex + 6).FillColor = vbBlack Then Exit Sub
    If Shape1(cindex + 6).FillColor = vbGreen Then
        MsgBox "You won!", vbInformation, "Winner!"
        MsgBox "Next Level", vbInformation, "Winner!"
        valwin = 1
    End If
    cindex = cindex + 6
    Shape1(cindex).FillColor = vbBlue
    Shape1(cindex - 6).FillColor = vbWhite
ElseIf KeyCode = vbKeyLeft Then
    If cindex = 0 Or cindex = 1 Or cindex = 2 Or cindex = 3 Or cindex = 4 Or cindex = 5 Then Exit Sub
    If Shape1(cindex - 6).FillColor = vbBlack Then Exit Sub
    If Shape1(cindex - 6).FillColor = vbGreen Then
        MsgBox "You won!", vbInformation, "Winner!"
        MsgBox "Next Level", vbInformation, "Winner!"
        valwin = 1
    End If
    cindex = cindex - 6
    Shape1(cindex).FillColor = vbBlue
    Shape1(cindex + 6).FillColor = vbWhite
End If
If valwin = 1 Then
    frmlvl2.Show 1
    Unload frmmaze
End If
End Sub




Private Sub Form_Load()
'set color
'Flash! Ah-ah!
'Copywrite 2006 Hesp Corp.
MsgBox "Copywrite 2006 Hesp Corp.", vbCritical, "Copywrite 2006"
Shape1(5).FillColor = vbBlue
Shape1(15).FillColor = vbGreen
valwin = 0
End Sub





Private Sub Label1_Click()
End
End Sub


Private Sub Label3_Click()
frmlvl2.Show 1
End Sub


