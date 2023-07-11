VERSION 5.00
Begin VB.Form frmlvl2 
   Caption         =   "Level 2"
   ClientHeight    =   7080
   ClientLeft      =   2400
   ClientTop       =   2055
   ClientWidth     =   6585
   Icon            =   "frmlvl2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   6585
   Begin VB.Label Label2 
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
      Height          =   510
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "`"
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   45
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   99
      Left            =   5640
      Top             =   6120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   98
      Left            =   5640
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   97
      Left            =   5640
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   96
      Left            =   5640
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   95
      Left            =   5640
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   94
      Left            =   5640
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   93
      Left            =   5640
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   92
      Left            =   5640
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   91
      Left            =   5640
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   90
      Left            =   5640
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   89
      Left            =   5040
      Top             =   6120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   88
      Left            =   5040
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   87
      Left            =   5040
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   86
      Left            =   5040
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   85
      Left            =   5040
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   84
      Left            =   5040
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   83
      Left            =   5040
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   82
      Left            =   5040
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   81
      Left            =   5040
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   80
      Left            =   5040
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   79
      Left            =   4440
      Top             =   6120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   78
      Left            =   4440
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   77
      Left            =   4440
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   76
      Left            =   4440
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   75
      Left            =   4440
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   74
      Left            =   4440
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   73
      Left            =   4440
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   72
      Left            =   4440
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   71
      Left            =   4440
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   70
      Left            =   4440
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   69
      Left            =   3840
      Top             =   6120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   68
      Left            =   3840
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   67
      Left            =   3840
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   66
      Left            =   3840
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   65
      Left            =   3840
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   64
      Left            =   3840
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   63
      Left            =   3840
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   62
      Left            =   3840
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   61
      Left            =   3840
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   60
      Left            =   3840
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   59
      Left            =   3240
      Top             =   6120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   58
      Left            =   3240
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   57
      Left            =   3240
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   56
      Left            =   3240
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   55
      Left            =   3240
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   54
      Left            =   3240
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   53
      Left            =   3240
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   52
      Left            =   3240
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   51
      Left            =   3240
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   50
      Left            =   3240
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   49
      Left            =   2640
      Top             =   6120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   48
      Left            =   2640
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   47
      Left            =   2640
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   46
      Left            =   2640
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   45
      Left            =   2640
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   44
      Left            =   2640
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   43
      Left            =   2640
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   42
      Left            =   2640
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   41
      Left            =   2640
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   40
      Left            =   2640
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   39
      Left            =   2040
      Top             =   6120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   38
      Left            =   2040
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   37
      Left            =   2040
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   36
      Left            =   2040
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   35
      Left            =   2040
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   34
      Left            =   2040
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   33
      Left            =   2040
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   32
      Left            =   2040
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   31
      Left            =   2040
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   30
      Left            =   2040
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   29
      Left            =   1440
      Top             =   6120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   28
      Left            =   1440
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   27
      Left            =   1440
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   26
      Left            =   1440
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   25
      Left            =   1440
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   24
      Left            =   1440
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   23
      Left            =   1440
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   22
      Left            =   1440
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   21
      Left            =   1440
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   20
      Left            =   1440
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   19
      Left            =   840
      Top             =   6120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   18
      Left            =   840
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   17
      Left            =   840
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   16
      Left            =   840
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   15
      Left            =   840
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   14
      Left            =   840
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   13
      Left            =   840
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   12
      Left            =   840
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   11
      Left            =   840
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   10
      Left            =   840
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   9
      Left            =   240
      Top             =   6120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   8
      Left            =   240
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   7
      Left            =   240
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   6
      Left            =   240
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   5
      Left            =   240
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   4
      Left            =   240
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   3
      Left            =   240
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   240
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   240
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   240
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmlvl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BY: KYLE HESPE
Dim cindex As Integer
Dim valwin As Integer
Dim valmove As Integer
Dim valpowerup As Integer

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
    If cindex = 0 Or cindex = 10 Or cindex = 20 Or cindex = 30 Or cindex = 40 Or cindex = 50 Or cindex = 60 Or cindex = 70 Or cindex = 80 Or cindex = 90 Then Exit Sub
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
    If cindex = 9 Or cindex = 19 Or cindex = 29 Or cindex = 39 Or cindex = 49 Or cindex = 59 Or cindex = 69 Or cindex = 79 Or cindex = 89 Or cindex = 99 Then Exit Sub
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
    If cindex = 90 Or cindex = 91 Or cindex = 92 Or cindex = 93 Or cindex = 94 Or cindex = 95 Or cindex = 96 Or cindex = 97 Or cindex = 98 Or cindex = 99 Then Exit Sub
    If Shape1(cindex + 10).FillColor = vbBlack Then Exit Sub
    If Shape1(cindex + 10).FillColor = vbGreen Then
        MsgBox "You won!", vbInformation, "Winner!"
        MsgBox "Next Level", vbInformation, "Winner!"
        valwin = 1
    End If
    cindex = cindex + 10
    Shape1(cindex).FillColor = vbBlue
    Shape1(cindex - 10).FillColor = vbWhite
ElseIf KeyCode = vbKeyLeft Then
    If cindex = 0 Or cindex = 1 Or cindex = 2 Or cindex = 3 Or cindex = 4 Or cindex = 5 Or cindex = 6 Or cindex = 7 Or cindex = 8 Or cindex = 9 Then Exit Sub
    If Shape1(cindex - 10).FillColor = vbBlack Then Exit Sub
    If Shape1(cindex - 10).FillColor = vbGreen Then
        MsgBox "You won!", vbInformation, "Winner!"
        MsgBox "Next Level", vbInformation, "Winner!"
        valwin = 1
    End If
    cindex = cindex - 10
    Shape1(cindex).FillColor = vbBlue
    Shape1(cindex + 10).FillColor = vbWhite
End If

If Shape1(27).FillColor = vbBlue Then
    MsgBox "POWER UP!!!", vbCritical, "Power-up"
End If
If Shape1(27).FillColor = vbWhite Then
    Shape1(27).FillColor = vbYellow
End If


If Shape1(79).FillColor = vbBlue And Shape1(55).FillColor = vbWhite Then
    Shape1(55).FillColor = vbBlack
End If
If Shape1(79).FillColor = vbWhite Then
    Shape1(79).FillColor = vbRed
End If

If Shape1(79).FillColor = vbBlue And Shape1(55).FillColor = vbBlack Then
    Shape1(19).FillColor = vbWhite
    Shape1(29).FillColor = vbWhite
    Shape1(39).FillColor = vbWhite
End If

If valwin = 1 Then
    frmlvl3.Show 1
End If
End Sub



Private Sub Form_Load()
Unload frmmaze

Shape1(9).FillColor = vbBlue
Shape1(99).FillColor = vbGreen
Shape1(27).FillColor = vbYellow
Shape1(79).FillColor = vbRed
valwin = 0
End Sub




Private Sub Label1_Click()
frmlvl3.Show 1
End Sub


Private Sub Label2_Click()
End
End Sub


