VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reversi"
   ClientHeight    =   5055
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7185
   ForeColor       =   &H00B75820&
   Icon            =   "Reversi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1680
      Top             =   4680
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Setup"
      Height          =   4815
      Left            =   5040
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "End Setup"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   4200
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080C0FF&
         Caption         =   "White Player"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Black Player"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Side to play:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   6240
      Top             =   3720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Status"
      Height          =   4815
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton Command3 
         Caption         =   "Stop Auto Play"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   840
         Top             =   3600
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Good Computer"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Human"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "White Player:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Black Player:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Black discs:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "White discs:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Black"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Next turn:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   4320
         Width           =   855
      End
   End
   Begin VB.Line Line2 
      Index           =   8
      X1              =   328
      X2              =   328
      Y1              =   8
      Y2              =   328
   End
   Begin VB.Line Line2 
      Index           =   7
      X1              =   288
      X2              =   288
      Y1              =   8
      Y2              =   328
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   248
      X2              =   248
      Y1              =   8
      Y2              =   328
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   208
      X2              =   208
      Y1              =   8
      Y2              =   328
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   168
      X2              =   168
      Y1              =   8
      Y2              =   328
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   128
      X2              =   128
      Y1              =   8
      Y2              =   328
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   88
      X2              =   88
      Y1              =   8
      Y2              =   328
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   48
      X2              =   48
      Y1              =   8
      Y2              =   328
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   8
      X2              =   8
      Y1              =   8
      Y2              =   328
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   8
      X1              =   8
      X2              =   328
      Y1              =   328
      Y2              =   328
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   7
      X1              =   8
      X2              =   328
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   6
      X1              =   8
      X2              =   328
      Y1              =   248
      Y2              =   248
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   5
      X1              =   8
      X2              =   328
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   4
      X1              =   8
      X2              =   328
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   3
      X1              =   8
      X2              =   328
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   2
      X1              =   8
      X2              =   328
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   8
      X2              =   328
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   8
      X2              =   328
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   63
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   62
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   61
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   60
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   59
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   58
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   57
      Left            =   720
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   56
      Left            =   120
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   55
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   54
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   53
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   52
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   51
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   50
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   49
      Left            =   720
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   48
      Left            =   120
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   47
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   46
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   45
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   44
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   43
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   42
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   41
      Left            =   720
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   40
      Left            =   120
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   39
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   38
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   37
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   36
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   35
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   34
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   33
      Left            =   720
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   32
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   31
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   30
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   29
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   28
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   27
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   26
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   25
      Left            =   720
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   24
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   23
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   22
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   21
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   20
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   19
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   18
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   17
      Left            =   720
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   16
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   15
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   14
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   13
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   12
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   11
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   10
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   9
      Left            =   720
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   8
      Left            =   120
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   7
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   6
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   5
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   4
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   3
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   2
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   1
      Left            =   720
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   0
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin ComctlLib.ImageList imgMain 
      Left            =   4320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Reversi.frx":3B6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Reversi.frx":4E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Reversi.frx":618E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Reversi.frx":74A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnugame 
      Caption         =   "Game"
      Begin VB.Menu mnunew 
         Caption         =   "New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuseggest 
         Caption         =   "Suggest Move"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnusetup 
         Caption         =   "Setup Mode"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "Options"
      Begin VB.Menu mnublack 
         Caption         =   "Black Player"
         Begin VB.Menu blackoptions 
            Caption         =   "Human"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu blackoptions 
            Caption         =   "Beginner Computer"
            Index           =   1
         End
         Begin VB.Menu blackoptions 
            Caption         =   "Novice Computer"
            Index           =   2
         End
         Begin VB.Menu blackoptions 
            Caption         =   "Good Computer"
            Index           =   3
         End
         Begin VB.Menu blackoptions 
            Caption         =   "Expert Computer"
            Index           =   4
         End
      End
      Begin VB.Menu mnuwhite 
         Caption         =   "White Player"
         Index           =   0
         Begin VB.Menu whiteoptions 
            Caption         =   "Human"
            Index           =   0
         End
         Begin VB.Menu whiteoptions 
            Caption         =   "Beginner Computer"
            Index           =   1
         End
         Begin VB.Menu whiteoptions 
            Caption         =   "Novice Computer"
            Index           =   2
         End
         Begin VB.Menu whiteoptions 
            Caption         =   "Good Computer"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu whiteoptions 
            Caption         =   "Expert Computer"
            Index           =   4
         End
      End
      Begin VB.Menu seperate1 
         Caption         =   "-"
      End
      Begin VB.Menu mnushowpossible 
         Caption         =   "Show possible moves"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'ensure all variables are declared

'stop for a while
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim Grid(1 To 8, 1 To 8) As Integer 'keeps the board information(0 = nothing , 1 = black , 2 = white)
Dim turn As String 'which player is going to play now
Dim WhiteCount As Integer, BlackCount As Integer 'number of discs for each player
Dim LegalMoves(40, 2) As Integer, LegalMovesNum
Dim NewDiscs(1 To 30, 1 To 2) As Integer, FlippedNum As Integer
Dim SelMove(1 To 2) As Integer 'the move that the computer thinks its the best (i for x coordinate,2 for y coordinate)

Dim Nodes As Long 'number of nodes
Dim SearchEnd As Boolean 'true if we are at the end of game and will calculate the exact final score
Dim startdepth As Integer 'the depth we are gring to search to
Dim MidDepth As Integer, EndDepth As Integer 'the depth for the mid game and for the end game(depends of the computer level)
Dim Player1Type As Integer, Player2Type As Integer 'whether the player is a human or a computerat a certain level
Dim freezed As Boolean

Private Sub DrawBoard()
'draws the board
Dim imgnum As Integer, theplayer As Integer, i As Integer, j As Integer
Select Case turn
Case "black"
theplayer = 1
Case "white"
theplayer = 2
End Select
For i = 1 To 8 'loop over in x direction
    For j = 1 To 8 'loop over cells in y direction
    Select Case Grid(i, j)
    Case 0
    imgnum = 3 'image number 3 in the imagelist
    Case 1
    imgnum = 2 'image number 2 in the imagelist
    Case 2
    imgnum = 1 'image number 1 in the imagelist
    End Select
    'show possible moves
    If IsValid(theplayer, i, j) = True And mnushowpossible.Checked = True And mnusetup.Checked = False Then imgnum = 4
    'draw the image
    Image1((j - 1) * 8 + i - 1).Picture = imgMain.ListImages.Item(imgnum).Picture
    Next
Next
End Sub
Private Sub MakeMove(player As Integer, x As Integer, y As Integer)
Dim i As Integer, j As Integer, xx As Integer, Flipped As Integer
'changes the grid information after making a move and updates the number of discs for each player


Grid(x, y) = player

'check for flippedd discs above the new placed disc
Flipped = 0
For i = y - 1 To 1 Step -1
    If Grid(x, i) = 0 Then Exit For
    If Grid(x, i) = player Then
        For j = y - 1 To i + 1 Step -1
            Grid(x, j) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'check for flippedd discs below the new placed disc
For i = y + 1 To 8 Step 1
    If Grid(x, i) = 0 Then Exit For
    If Grid(x, i) = player Then
        For j = y + 1 To i - 1 Step 1
            Grid(x, j) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'check for flippedd discs at the right of the new placed disc
For i = x + 1 To 8 Step 1
    If Grid(i, y) = 0 Then Exit For
    If Grid(i, y) = player Then
        For j = x + 1 To i - 1 Step 1
            Grid(j, y) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next


'check for flippedd discs at the left of the new placed disc\
For i = x - 1 To 1 Step -1
    If Grid(i, y) = 0 Then Exit For
    If Grid(i, y) = player Then
        For j = x - 1 To i + 1 Step -1
            Grid(j, y) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'check up left
For i = y - 1 To 1 Step -1
    xx = x - (y - i)
    If xx < 1 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = y - 1 To i + 1 Step -1
            xx = x - (y - j)
            Grid(xx, j) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'check up right
For i = y - 1 To 1 Step -1
    xx = x + (y - i)
    If xx > 8 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = y - 1 To i + 1 Step -1
            xx = x + (y - j)
            Grid(xx, j) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'check down left
For i = y + 1 To 8 Step 1
    xx = x - (i - y)
    If xx < 1 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = y + 1 To i - 1 Step 1
            xx = x - (j - y)
            Grid(xx, j) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'check down right
For i = y + 1 To 8 Step 1
    xx = x + (i - y)
    If xx > 8 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = y + 1 To i - 1 Step 1
            xx = x + (j - y)
            Grid(xx, j) = player
            Flipped = Flipped + 1
        Next
        Exit For
    End If
Next

'update the number of discs for each player
Select Case player
Case 2
WhiteCount = WhiteCount + Flipped + 1
BlackCount = BlackCount - Flipped
Case 1
WhiteCount = WhiteCount - Flipped
BlackCount = BlackCount + Flipped + 1
End Select
End Sub

Private Sub GetFlipped(player As Integer, x As Integer, y As Integer)
Dim i As Integer, j As Integer, xx As Integer
'this sub if used in move ordering, if fills the (NewDiscs) array
'with the discs that will be flipped after
'making a move and if works almost the same as the (MakeMove) sub

FlippedNum = 1

NewDiscs(1, 1) = x
NewDiscs(1, 2) = y

For i = y - 1 To 1 Step -1
    If Grid(x, i) = 0 Then Exit For
    If Grid(x, i) = player Then
        For j = y - 1 To i + 1 Step -1
        FlippedNum = FlippedNum + 1
        NewDiscs(FlippedNum, 1) = x
        NewDiscs(FlippedNum, 2) = j
        Next
        Exit For
    End If
Next

For i = y + 1 To 8 Step 1
    If Grid(x, i) = 0 Then Exit For
    If Grid(x, i) = player Then
        For j = y + 1 To i - 1 Step 1
        FlippedNum = FlippedNum + 1
        NewDiscs(FlippedNum, 1) = x
        NewDiscs(FlippedNum, 2) = j
        Next
        Exit For
    End If
Next

For i = x + 1 To 8 Step 1
    If Grid(i, y) = 0 Then Exit For
    If Grid(i, y) = player Then
        For j = x + 1 To i - 1 Step 1
        FlippedNum = FlippedNum + 1
        NewDiscs(FlippedNum, 1) = j
        NewDiscs(FlippedNum, 2) = y
        Next
        Exit For
    End If
Next

For i = x - 1 To 1 Step -1
    If Grid(i, y) = 0 Then Exit For
    If Grid(i, y) = player Then
        For j = x - 1 To i + 1 Step -1
        FlippedNum = FlippedNum + 1
        NewDiscs(FlippedNum, 1) = j
        NewDiscs(FlippedNum, 2) = y
        Next
        Exit For
    End If
Next

For i = y - 1 To 1 Step -1
    xx = x - (y - i)
    If xx < 1 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = y - 1 To i + 1 Step -1
            xx = x - (y - j)
            FlippedNum = FlippedNum + 1
            NewDiscs(FlippedNum, 1) = xx
            NewDiscs(FlippedNum, 2) = j
        Next
        Exit For
    End If
Next

For i = y - 1 To 1 Step -1
    xx = x + (y - i)
    If xx > 8 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = y - 1 To i + 1 Step -1
            xx = x + (y - j)
            FlippedNum = FlippedNum + 1
            NewDiscs(FlippedNum, 1) = xx
            NewDiscs(FlippedNum, 2) = j
        Next
        Exit For
    End If
Next

For i = y + 1 To 8 Step 1
    xx = x - (i - y)
    If xx < 1 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = y + 1 To i - 1 Step 1
            xx = x - (j - y)
            FlippedNum = FlippedNum + 1
            NewDiscs(FlippedNum, 1) = xx
            NewDiscs(FlippedNum, 2) = j
        Next
        Exit For
    End If
Next

For i = y + 1 To 8 Step 1
    xx = x + (i - y)
    If xx > 8 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player Then
        For j = y + 1 To i - 1 Step 1
            xx = x + (j - y)
            FlippedNum = FlippedNum + 1
            NewDiscs(FlippedNum, 1) = xx
            NewDiscs(FlippedNum, 2) = j
        Next
        Exit For
    End If
Next


End Sub

Private Function IsValid(player As Integer, x As Integer, y As Integer) As Boolean
Dim i As Integer, xx As Integer
'this function checkes whether a move if valid in the current position and for a certain player
'it is similar to the(MakeMove)sub but if doedn't care about how many discs are flipped, it breaks
'and returnes true at the first time if finds a disc that wil be flipped

IsValid = False
If Grid(x, y) <> 0 Then Exit Function
For i = y - 1 To 1 Step -1
    If Grid(x, i) = 0 Then Exit For
    If Grid(x, i) = player And (y - i) >= 2 Then IsValid = True: Exit Function
    If Grid(x, i) = player Then Exit For
Next

For i = y + 1 To 8 Step 1
    If Grid(x, i) = 0 Then Exit For
    If Grid(x, i) = player And (i - y) >= 2 Then IsValid = True: Exit Function
    If Grid(x, i) = player Then Exit For
Next

For i = x + 1 To 8 Step 1
    If Grid(i, y) = 0 Then Exit For
    If Grid(i, y) = player And (i - x) >= 2 Then IsValid = True: Exit Function
    If Grid(i, y) = player Then Exit For
Next

For i = x - 1 To 1 Step -1
    If Grid(i, y) = 0 Then Exit For
    If Grid(i, y) = player And (x - i) >= 2 Then IsValid = True: Exit Function
    If Grid(i, y) = player Then Exit For
Next

For i = y - 1 To 1 Step -1
    xx = x - (y - i)
    If xx < 1 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player And (y - i) >= 2 Then IsValid = True: Exit Function
    If Grid(xx, i) = player Then Exit For
Next

For i = y - 1 To 1 Step -1
    xx = x + (y - i)
    If xx > 8 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player And (y - i) >= 2 Then IsValid = True: Exit Function
    If Grid(xx, i) = player Then Exit For
Next

For i = y + 1 To 8 Step 1
    xx = x - (i - y)
    If xx < 1 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player And (i - y) >= 2 Then IsValid = True: Exit Function
    If Grid(xx, i) = player Then Exit For
Next

For i = y + 1 To 8 Step 1
    xx = x + (i - y)
    If xx > 8 Then Exit For
    If Grid(xx, i) = 0 Then Exit For
    If Grid(xx, i) = player And (i - y) >= 2 Then IsValid = True: Exit Function
    If Grid(xx, i) = player Then Exit For
Next

End Function

Private Function CountBad(x As Integer, y As Integer)
Dim Bad As Integer
'counts how many empty cells are around a cell,it is used for move ordering
'so the more empty discs around a cell, the worse it is
Bad = 0
If x > 1 Then
If Grid(x - 1, y) = 0 Then Bad = Bad + 1
End If

If x < 8 Then
If Grid(x + 1, y) = 0 Then Bad = Bad + 1
End If

If y > 1 Then
If Grid(x, y - 1) = 0 Then Bad = Bad + 1
End If

If y < 8 Then
If Grid(x, y + 1) = 0 Then Bad = Bad + 1
End If

If x > 1 And y > 1 Then
If Grid(x - 1, y - 1) = 0 Then Bad = Bad + 1
End If

If x < 8 And y < 8 Then
If Grid(x + 1, y + 1) = 0 Then Bad = Bad + 1
End If

If x > 1 And y < 8 Then
If Grid(x - 1, y + 1) = 0 Then Bad = Bad + 1
End If

If x < 8 And y > 1 Then
If Grid(x + 1, y - 1) = 0 Then Bad = Bad + 1
End If

CountBad = Bad

End Function

Private Sub GetLegalMoves(player As Integer)
'fills the (LegalMoves) array with the legal moves for a certain
'player in the current position(used for move generation)
Dim i As Integer, j As Integer
LegalMovesNum = 0
For i = 1 To 8
    For j = 1 To 8
    If IsValid(player, i, j) = True Then
    LegalMovesNum = LegalMovesNum + 1
    LegalMoves(LegalMovesNum, 1) = i
    LegalMoves(LegalMovesNum, 2) = j
    End If
    Next
Next
End Sub

Private Sub NewGame()
'starts a new game, resets everything to default
Dim i As Integer, j As Integer
For i = 1 To 8
    For j = 1 To 8
    Grid(i, j) = 0
    Next
Next
Grid(4, 4) = 2
Grid(5, 5) = 2
Grid(5, 4) = 1
Grid(4, 5) = 1
WhiteCount = 2
BlackCount = 2
Label5.Caption = "2"
Label6.Caption = "2"

turn = "black"
Label2.Caption = "black"
DrawBoard

Timer1.Enabled = True
Timer2.Enabled = True
freezed = False
End Sub
Private Function EvaluateEnd(player As Integer)
'the evaluation function used in the end of the game, it returnes
'the difference between our discs and the opponents discs
Dim Value As Integer
Value = WhiteCount - BlackCount
If player = 2 Then EvaluateEnd = Value Else: EvaluateEnd = -Value
End Function

Private Function Evaluate(player As Integer)
'the evaluation function used in the begining and the middle of the game, it is
'based on mobility(number of legal moves) for each player and also on the board
'positions that each player occupy
Dim Value As Integer, cellval As Integer, cellnum As Integer
Dim i As Integer, j As Integer
'at first the function evaluates for the white player and then changes the sign if we are evaluating for black
Value = 0
For i = 1 To 8
    For j = 1 To 8
    If IsValid(2, i, j) = True Then Value = Value + 18 'add 18 for each legal move for white
    If IsValid(1, i, j) = True Then Value = Value - 18 'subtract 18 for each legal move for black
    If Grid(i, j) <> 0 Then
        cellnum = (j - 1) * 8 + i
        Select Case cellnum
        Case 1, 8, 57, 64 'a corner
        cellval = 250
        Case 2, 7, 9, 16, 49, 56, 58, 63 'C square(a square that is adjacent to a corner horizontally or vertically
        cellval = -30
        Case 10, 15, 50, 55 'X square(a square that is adjacent to a corner diagonally
        cellval = -40
        Case 3, 4, 5, 6, 17, 25, 33, 41, 24, 32, 40, 48, 59, 60, 61, 62 ' sides other than C squares
        cellval = 15
        Case Else 'any other place(in the middle of the board)
        cellval = 0
        End Select
        'adds or subtracts the score depending on which player occupies that cell
        If Grid(i, j) = 2 Then Value = Value + cellval Else: Value = Value - cellval
    End If
    Next
Next
'negate the value if we are evaluating for black and return the normal value if we are evaluating for white
If player = 2 Then Evaluate = Value Else: Evaluate = -Value
        
End Function


Private Function Search(depth As Integer, alpha As Integer, beta As Integer)
'the searching function(finds the best move for the computer)
Dim NewScore As Integer
Dim LegalMovesNow(40, 2) As Integer 'keeps the legal moves in this procedure
Dim player As Integer 'which player is going to play at the current depth
Dim TempGrid(1 To 8, 1 To 8) As Integer 'stores the grid information before making a move in order to be able to unmake it
Dim a As Integer, b As Integer
Dim tempwhite As Integer, tempblack As Integer 'stores the number of discs for each player before making a move in order to be able to unmake it
Dim i As Integer

'determine which player is to go at the current depth
Select Case turn
Case "white"
If (startdepth - depth) Mod 2 = 0 Then
player = 2
Else
player = 1
End If
Case "black"
If (startdepth - depth) Mod 2 = 0 Then
player = 1
Else
player = 2
End If
End Select

Erase LegalMovesNow()

If depth = 0 Then 'if we are at the end of the tree then return evaluation
    If SearchEnd = True Then
    alpha = EvaluateEnd(player)
    Else
    alpha = Evaluate(player)
    End If
    Search = alpha
    Nodes = Nodes + 1
    Exit Function
End If

GetLegalMoves (player) 'generates the legal moves for the current player

If LegalMovesNum > 0 Then

OrderMoves (player) ' move ordering to make the search faster(by producing cutoffs earlier)

'copy the legal moves array to the current procedure
For i = 1 To LegalMovesNum
LegalMovesNow(i, 1) = LegalMoves(i, 1)
LegalMovesNow(i, 2) = LegalMoves(i, 2)
Next


For i = 1 To LegalMovesNum
 
'store grid information in order to be able to unmake the move
For a = 1 To 8
    For b = 1 To 8
    TempGrid(a, b) = Grid(a, b)
    Next
Next
tempwhite = WhiteCount
tempblack = BlackCount

 Call MakeMove(player, LegalMovesNow(i, 1), LegalMovesNow(i, 2))
 
 
 NewScore = -Search(depth - 1, -beta, -alpha) 'calculate the score after making this move
  
'unmake the move
For a = 1 To 8
    For b = 1 To 8
    Grid(a, b) = TempGrid(a, b)
    Next
Next
 WhiteCount = tempwhite
 BlackCount = tempblack
 
 If NewScore >= beta Then
 'make a cutoff because the opponent won't let me to get in this too good position for me (he already knows a strategy to avoid this)
 Search = beta
 Exit Function
 End If
 
 
 If NewScore > alpha Then 'if the new score is better than the best one found so far
 alpha = NewScore 'make if the best score
    If depth = startdepth Then ' if we are at the root nodes then record this move because it is the best one found so far
    SelMove(1) = LegalMovesNow(i, 1)
    SelMove(2) = LegalMovesNow(i, 2)
    End If
 End If


Next

Search = alpha

Else ' if LegalMovesNum = 0

Search = -Search(depth - 1, -beta, -alpha)

End If
End Function

Private Sub OrderMoves(player As Integer)
'orders the legal moves by giving quick values for them and then sorting
Dim Values(40) As Integer
Dim TempGrid(1 To 8, 1 To 8) As Integer
Dim tempwhite As Integer, tempblack As Integer
Dim i As Integer, j As Integer, a As Integer, b As Integer
Dim temp1 As Integer, temp2 As Integer
Dim cellnum As Integer, cellval As Integer, score As Integer
Dim Counted As Integer

For i = 1 To LegalMovesNum

score = 0
Call GetFlipped(player, LegalMoves(i, 1), LegalMoves(i, 2))

'the value of a move is the sum of the values for all the discs that if flipps
For j = 1 To FlippedNum

a = NewDiscs(j, 1)
b = NewDiscs(j, 2)
 
 ' there is something very similar to this in the evaluation function
        cellnum = (b - 1) * 8 + a
        Select Case cellnum
        Case 1, 8, 57, 64
        cellval = 250
        Case 2, 7, 9, 16, 49, 56, 58, 63
        cellval = -30
        Case 10, 15, 50, 55
        cellval = -40
        Case 3, 4, 5, 6, 17, 25, 33, 41, 24, 32, 40, 48, 59, 60, 61, 62
        cellval = 15
        Case Else
        cellval = 0
        End Select
 
Counted = CountBad(a, b)
cellval = cellval - (Counted + 8 * Sgn(Counted)) ' the value of the cell also depends on how many empty cells are surrounding it
Next
 
score = score + cellval
Values(i) = score
  


Next

'sort the moves
Dim LargestVal As Integer, BestNow As Integer
For i = 1 To LegalMovesNum
LargestVal = -10000
    For j = i To LegalMovesNum
    If Values(j) > LargestVal Then BestNow = j: LargestVal = Values(j)
    Next
    temp1 = LegalMoves(BestNow, 1)
    temp2 = LegalMoves(BestNow, 2)
    LegalMoves(BestNow, 1) = LegalMoves(i, 1)
    LegalMoves(BestNow, 2) = LegalMoves(i, 2)
    LegalMoves(i, 1) = temp1
    LegalMoves(i, 2) = temp2
    Values(BestNow) = Values(i)
Next
End Sub
Private Sub ChangeTurn()
'changes the turn for the other player or passes turns or ends the game
Dim player As Integer, opponent As Integer, i As Integer, j As Integer
Dim whattodo As String
Dim PlayerMayPass As String
DrawBoard

Select Case turn
Case "black"
PlayerMayPass = "White"
player = 2
opponent = 1
Case "white"
PlayerMayPass = "Black"
player = 1
opponent = 2
End Select

whattodo = ""
For i = 1 To 8
    For j = 1 To 8
    'if there are legal moves for the othar player then only change the turn
    If IsValid(player, i, j) = True Then whattodo = "normal": Exit For
    Next
Next
'if there are no legal moves for the other player then check if we are going
'to pass the turn and give the player already moves another turn
If whattodo = "" Then

For i = 1 To 8
    For j = 1 To 8
    If IsValid(opponent, i, j) = True Then whattodo = "pass": Exit For
    Next
Next
End If

'if both players have no moves then end the game
If whattodo = "" Then whattodo = "endgame"

Select Case whattodo
Case "normal"
    Select Case turn
    Case "black"
    turn = "white"
    Case "white"
    turn = "black"
    End Select
Case "pass"
    Dim s As String
    s = PlayerMayPass + " passes"
    MsgBox (s)
Case "endgame"
    EndGame
End Select

Label2.Caption = turn
End Sub

Private Sub EndGame()
'ends the game
DrawBoard
If BlackCount > WhiteCount Then
MsgBox ("The black player wins the game")
ElseIf BlackCount < WhiteCount Then
MsgBox ("The white player wins the game")
Else
MsgBox ("The game is a draw")
End If
Timer1.Enabled = False
Timer2.Enabled = False

End Sub
Private Sub TakeMove(Index As Integer)
'takes a move for the computer or for the human player
Dim a As Integer, b As Integer, theplayer As Integer
a = (Index Mod 8) + 1
b = Int(Index / 8) + 1
Select Case turn
Case "black"
theplayer = 1
Case "white"
theplayer = 2
End Select
If IsValid(theplayer, a, b) = True Then
Call MakeMove(theplayer, a, b)
Label5.Caption = Str(WhiteCount)
Label6.Caption = Str(BlackCount)
ChangeTurn
End If

DrawBoard

End Sub

Private Sub CompMove()
'makes the computer move
Dim TheDepth As Integer, best As Integer, a As Integer, b As Integer, c As Integer
freezed = True 'make the player not able to make a move
Form1.MousePointer = 11 'change the shape of the mouse ponter
Nodes = 0
If (BlackCount + WhiteCount) > (64 - EndDepth) Then 'if we are at the end of the game then make a search to play the perfect move
SearchEnd = True
TheDepth = 20 'the depth is higher than the number of empty discs to handle passes
Else 'nor at the end of the game, make a normal search
SearchEnd = False
TheDepth = MidDepth
End If
startdepth = TheDepth
best = Search(TheDepth, -5000, 5000) 'make the search
Form1.MousePointer = 1

a = SelMove(1)
b = SelMove(2)
c = 8 * (b - 1) + a - 1 'calculate the index of the square from and x and y coordinates
TakeMove (c)
freezed = False
End Sub


Private Sub blackoptions_Click(Index As Integer)
Dim i As Integer
For i = 0 To 4
blackoptions(i).Checked = False
Next
blackoptions(Index).Checked = True
Player1Type = Index
Label9.Caption = blackoptions(Index).Caption
End Sub


Private Sub Command1_Click()
'the (End Setup) command button
Dim theplayer As Integer, opponent As Integer
Dim DoWhat As String
BlackCount = 0
WhiteCount = 0
Dim i As Integer, j As Integer
For i = 1 To 8
    For j = 1 To 8
    If Grid(i, j) = 1 Then BlackCount = BlackCount + 1
    If Grid(i, j) = 2 Then WhiteCount = WhiteCount + 1
    Next
Next

Label5.Caption = Str(WhiteCount)
Label6.Caption = Str(BlackCount)

If Option1.Value = True Then theplayer = 1: opponent = 2 Else: theplayer = 2: opponent = 1

DoWhat = ""

GetLegalMoves (theplayer)
If LegalMovesNum > 0 Then
DoWhat = "normal"
Select Case theplayer
Case 1
turn = "black"
Case 2
turn = "white"
End Select
Label2.Caption = turn
End If

If DoWhat = "" Then
    GetLegalMoves (opponent)
    If LegalMovesNum > 0 Then DoWhat = "pass"
End If

If DoWhat = "" Then EndGame
If DoWhat = "pass" Then ChangeTurn

mnusetup_Click
End Sub

Private Sub Command2_Click()
'the (clear) command button
Dim i As Integer, j As Integer
For i = 1 To 8
    For j = 1 To 8
    Grid(i, j) = 0
    Next
Next
Grid(4, 4) = 2
Grid(5, 5) = 2
Grid(5, 4) = 1
Grid(4, 5) = 1
DrawBoard
End Sub

Private Sub Command3_Click()
blackoptions_Click (0)
End Sub

Private Sub Form_Load()
NewGame
Player1Type = 0
Player2Type = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_Click(Index As Integer)
If mnusetup.Checked = True Then 'if we are in the setup mode
    Dim a As Integer, b As Integer
    a = (Index Mod 8) + 1
    b = Int(Index / 8) + 1
    Select Case Grid(a, b)
    Case 0
    Grid(a, b) = 1
    Case 1
    Grid(a, b) = 2
    Case 2
    Grid(a, b) = 0
    End Select
    DrawBoard
Else ' we are in the normal mode
    If freezed = True Then Exit Sub
    TakeMove (Index)
End If

End Sub


Private Sub mnuexit_Click()
End
End Sub

Private Sub mnunew_Click()
NewGame
End Sub

Private Sub mnuseggest_Click()
'similar to the (CompMove) sub
If (Player1Type <> 0 And turn = "black") Or (Player2Type <> 0 And turn = "white") Then Exit Sub 'exit sub if it is computer to play
If Timer1.Enabled = False And Timer2.Enabled = False Then Exit Sub 'exit sub if the game is over
Dim OppType As Integer, imgnum As Integer
mnuseggest.Enabled = False
freezed = True
If turn = "black" Then OppType = Player2Type: imgnum = 2
If turn = "white" Then OppType = Player1Type: imgnum = 1
If OppType = 0 Then OppType = 3
Select Case OppType
Case 1
MidDepth = 1
Case 2
MidDepth = 2
Case 3
MidDepth = 4
Case 4
MidDepth = 6
End Select
EndDepth = MidDepth * 2

Dim TheDepth As Integer, best As Integer, a As Integer, b As Integer, c As Integer
Form1.MousePointer = 11
If (BlackCount + WhiteCount) > (64 - EndDepth) Then
SearchEnd = True
TheDepth = 20
Else
SearchEnd = False
TheDepth = MidDepth
End If
startdepth = TheDepth
best = Search(TheDepth, -5000, 5000)
Form1.MousePointer = 1

a = SelMove(1)
b = SelMove(2)
c = 8 * (b - 1) + a - 1

Dim i As Integer
For i = 1 To 3 'make the disc blink 3 times
    Image1(c).Picture = imgMain.ListImages.Item(imgnum).Picture
    DoEvents
    Sleep 200
    Image1(c).Picture = imgMain.ListImages.Item(4).Picture
    DoEvents
    Sleep 200
Next
freezed = False
mnuseggest.Enabled = True
End Sub

Private Sub mnusetup_Click()
mnugame.Enabled = mnusetup.Checked
mnuoptions.Enabled = mnusetup.Checked
Timer1.Enabled = mnusetup.Checked
Timer2.Enabled = mnusetup.Checked
mnusetup.Checked = Not (mnusetup.Checked)
Frame2.Visible = mnusetup.Checked
DrawBoard
End Sub

Private Sub mnushowpossible_Click()
mnushowpossible.Checked = Not (mnushowpossible.Checked)
DrawBoard
End Sub

Private Sub Timer1_Timer()
'make a move for the computer if the black player is a computer

'if two computers are playing then show the (Stop Aubo Play) button
If Timer1.Enabled = True And Timer2.Enabled = True And Player1Type <> 0 And Player2Type <> 0 Then
Command3.Visible = True
Else
Command3.Visible = False
End If

If turn = "black" And Player1Type <> 0 Then
'determine the search depth according to the computer level
Select Case Player1Type
Case 1
MidDepth = 1
Case 2
MidDepth = 2
Case 3
MidDepth = 4
Case 4
MidDepth = 6
End Select
EndDepth = MidDepth * 2
CompMove
End If
End Sub

Private Sub Timer2_Timer()
'similar to (timer1) but for the white player
If Timer1.Enabled = True And Timer2.Enabled = True And Player1Type <> 0 And Player2Type <> 0 Then
Command3.Visible = True
Else
Command3.Visible = False
End If

If turn = "white" And Player2Type <> 0 Then
Select Case Player2Type
Case 1
MidDepth = 1
Case 2
MidDepth = 2
Case 3
MidDepth = 4
Case 4
MidDepth = 6
End Select
EndDepth = MidDepth * 2
CompMove
End If
End Sub

Private Sub Timer3_Timer()
If Timer1.Enabled = True And Timer2.Enabled = True And Player1Type <> 0 And Player2Type <> 0 Then
Command3.Visible = True
Else
Command3.Visible = False
End If
End Sub

Private Sub whiteoptions_Click(Index As Integer)
Dim i As Integer
For i = 0 To 4
whiteoptions(i).Checked = False
Next
whiteoptions(Index).Checked = True
Player2Type = Index
Label10.Caption = whiteoptions(Index).Caption
End Sub
