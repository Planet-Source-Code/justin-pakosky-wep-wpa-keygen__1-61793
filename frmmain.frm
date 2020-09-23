VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "keygen"
   ClientHeight    =   6240
   ClientLeft      =   1380
   ClientTop       =   825
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Copy to Clipboard"
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox txtkey 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   360
      MaxLength       =   64
      TabIndex        =   3
      Top             =   5280
      Width           =   6015
   End
   Begin VB.CommandButton cmdgen 
      Caption         =   "Generate"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encryption options"
      Height          =   4455
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   6495
      Begin VB.OptionButton Option6 
         Caption         =   "8-63 characters (WPA-PSK) "
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   2880
         Width           =   2295
      End
      Begin VB.OptionButton Option5 
         Height          =   615
         Left            =   480
         TabIndex        =   8
         Top             =   3600
         Width           =   5055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "5 characters (11,881,376 possibilities)"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   5055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "13 characters (2,481,152,873,203,736,576 possibilities)"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1080
         Width           =   5415
      End
      Begin VB.OptionButton Option3 
         Caption         =   "10 characters (1,099,511,627,776 possibilities)"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1680
         Width           =   5055
      End
      Begin VB.OptionButton Option4 
         Caption         =   "26 characters (20,282,409,603,651,670,423,947,251,286,016 possibilities)"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   2040
         Width           =   5655
      End
      Begin VB.Frame Frame2 
         Caption         =   "WEP"
         Height          =   2175
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6135
         Begin VB.Label Label1 
            Caption         =   "Characters A-Z"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Characters A-F, 0-9"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "WPA"
         Height          =   1935
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   6135
         Begin VB.CommandButton cmdpos 
            Caption         =   "possibilities"
            Height          =   375
            Left            =   4800
            TabIndex        =   20
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtchar 
            Height          =   285
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   18
            Text            =   "8"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "Possibilities:"
            Height          =   255
            Left            =   600
            TabIndex        =   21
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Number of charcters:"
            Height          =   255
            Left            =   2760
            TabIndex        =   19
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblpos 
            Caption         =   "Press possibilities"
            Height          =   495
            Left            =   1560
            TabIndex        =   17
            Top             =   840
            Width           =   3975
         End
         Begin VB.Label Label3 
            Caption         =   "Characters A-F, 0-9"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin VB.Label Label4 
      Caption         =   "By Justin Pakosky"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   15
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label lbl1 
      Caption         =   "Wireless key generator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Made by justin pakosky
''I made it to have a keygen so when i set up networks i dont just have to hit keys then count them
''plus it proves i have to much time, lol
''I hope you find it use full i just added the WPA option
Option Explicit
Public bytnumofchar As Byte ''I forgot how to pass values (i even forgot dim) but,
Public bytrannum As Byte    ''i remembered dim so i just made form levels

Private Sub cmdpos_Click()
lblpos = 16 ^ txtchar
End Sub

Private Sub Form_Load()
Option4.Value = True
''I am not experanced enough to know why i get an error when i leave the option
''5 caption in it, it gives me an error with strings so i just asigned it on load
Option5.Caption = "64 characters (more then 11,579,208,923,731,619,542,357,098,500,869,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000 possibilities)"
End Sub

Private Sub cmdgen_Click()
If Option6 = True Then
lblpos = 16 ^ txtchar
End If
Dim bytnum As Byte
txtkey = ""
Call numofchar
Do While bytnum < bytnumofchar
    Call getkey
    bytnum = bytnum + 1
Loop
End Sub

Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText txtkey, vbCFText
End Sub

Public Sub getkey()
    
    If bytnumofchar = 5 Or bytnumofchar = 13 Then
    Randomize
    bytrannum = Rnd * 26
    Else
    Randomize
    bytrannum = Rnd * 16
    End If
    Call Char
End Sub
Public Sub Char()
    Dim strchar As String
    If bytnumofchar = 5 Or bytnumofchar = 13 Then
    Select Case bytrannum
        Case Is = 1
            strchar = "a"
        Case Is = 2
            strchar = "b"
        Case Is = 3
            strchar = "c"
        Case Is = 4
            strchar = "d"
        Case Is = 5
            strchar = "e"
        Case Is = 6
            strchar = "f"
        Case Is = 7
            strchar = "g"
        Case Is = 8
            strchar = "h"
        Case Is = 9
            strchar = "i"
        Case Is = 10
            strchar = "j"
        Case Is = 11
            strchar = "k"
        Case Is = 12
            strchar = "l"
        Case Is = 13
            strchar = "m"
        Case Is = 14
            strchar = "n"
        Case Is = 15
            strchar = "o"
        Case Is = 16
            strchar = "p"
        Case Is = 17
            strchar = "q"
        Case Is = 18
            strchar = "r"
        Case Is = 19
            strchar = "s"
        Case Is = 20
            strchar = "t"
        Case Is = 21
            strchar = "u"
        Case Is = 22
            strchar = "v"
        Case Is = 23
            strchar = "w"
        Case Is = 24
            strchar = "x"
        Case Is = 25
            strchar = "y"
        Case Is = 26
            strchar = "z"
        End Select
        Else
        Select Case bytrannum
        Case Is = 1
            strchar = "a"
        Case Is = 2
            strchar = "b"
        Case Is = 3
            strchar = "c"
        Case Is = 4
            strchar = "d"
        Case Is = 5
            strchar = "e"
        Case Is = 6
            strchar = "f"
        Case Is = 7
            strchar = "1"
        Case Is = 8
            strchar = "2"
        Case Is = 9
            strchar = "3"
        Case Is = 10
            strchar = "4"
        Case Is = 11
            strchar = "5"
        Case Is = 12
            strchar = "6"
        Case Is = 13
            strchar = "7"
        Case Is = 14
            strchar = "8"
        Case Is = 15
            strchar = "9"
        Case Is = 16
            strchar = "0"
        End Select
    End If
       txtkey.Text = txtkey.Text + strchar
End Sub
Public Sub numofchar()
Select Case True
    Case Option1
        bytnumofchar = 5
    Case Option2
        bytnumofchar = 13
    Case Option3
        bytnumofchar = 10
    Case Option4
        bytnumofchar = 26
    Case Option5
        bytnumofchar = 64
    Case Option6
        bytnumofchar = txtchar
End Select
End Sub

