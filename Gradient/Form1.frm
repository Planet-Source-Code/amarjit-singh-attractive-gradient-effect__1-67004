VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   23
      Left            =   1200
      TabIndex        =   24
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   22
      Left            =   2235
      TabIndex        =   23
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   21
      Left            =   3270
      TabIndex        =   22
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   20
      Left            =   4305
      TabIndex        =   21
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   19
      Left            =   5160
      TabIndex        =   20
      Top             =   600
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   18
      Left            =   5160
      TabIndex        =   19
      Top             =   1080
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   17
      Left            =   5160
      TabIndex        =   18
      Top             =   1560
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   16
      Left            =   5160
      TabIndex        =   17
      Top             =   2040
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   15
      Left            =   5160
      TabIndex        =   16
      Top             =   2520
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   14
      Left            =   5160
      TabIndex        =   15
      Top             =   3000
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   13
      Left            =   5160
      TabIndex        =   14
      Top             =   3480
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   12
      Left            =   5160
      TabIndex        =   13
      Top             =   3960
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   11
      Left            =   4305
      TabIndex        =   12
      Top             =   4440
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   10
      Left            =   3270
      TabIndex        =   11
      Top             =   4440
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   9
      Left            =   2235
      TabIndex        =   10
      Top             =   4440
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   8
      Left            =   1200
      TabIndex        =   9
      Top             =   4440
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effect 1"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1035
   End
   Begin VB.PictureBox gradient 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   1320
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   254
      TabIndex        =   0
      Top             =   600
      Width           =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim fire(255) As Long, w As Integer, h As Integer, brun As Boolean
'initialize the pallete
Sub initialize()
Dim i As Integer
For i = 0 To 63
fire(i) = RGB(i * 4, 0, 0)
Next
For i = 64 To 127
fire(i) = RGB(255, (i - 64) * 4, 0)

Next
For i = 128 To 191

fire(i) = RGB(255, 255, (i - 128) * 4)
Next
For i = 192 To 255
fire(i) = RGB((i - 192) * 4, (i - 192) * 4, (i - 192) * 4)

Next
w = gradient.ScaleWidth
h = gradient.ScaleHeight
End Sub
'Shift the Pallete
Sub change()
Dim temp As Long, i As Integer
temp = fire(128)
For i = 128 To 1 Step -1
fire(i) = fire(i - 1)
Next
fire(0) = temp

End Sub





Private Sub Command1_Click(Index As Integer)
Dim str As String
brun = False
effect Index

End Sub

Private Sub Form_Click()
brun = False
End Sub

Private Sub Form_Load()
initialize
brun = False
Dim i As Integer
For i = 0 To 23
Command1(i).Caption = "Effect " & i
Next
End Sub
Sub effect(idx As Integer)
gradient.Cls
brun = True
Dim i As Integer
Select Case idx
Case 0:

        Do While brun
        change
        For i = 0 To h
        gradient.ForeColor = fire(i Mod 128)
        gradient.Line (h - i, 0)-(h, i)
        gradient.Line (0, h - i)-(i, h)
        Next
        DoEvents
        Loop

Case 1:

        Do While brun
        change

        For i = 1 To h
        gradient.Line (i, 1)-(i, h), fire(i Mod 128)
        Next
        DoEvents
        Loop

Case 2:

        Do While brun
        change
        For i = 0 To h
        gradient.Line (0, 0)-(i, i), fire(i Mod 128), B
        Next
        DoEvents
        Loop

Case 3:
        Do While brun
        change
        For i = 0 To h
        gradient.ForeColor = fire(i Mod 128)
        gradient.Line (i, 0)-(h, h - i)
        gradient.Line (0, i)-(h - i, h)
        Next
        DoEvents
        Loop

Case 4:
        Do While brun
        change
        For i = 0 To h
        gradient.ForeColor = fire(i Mod 128)
        gradient.Line (h \ 2 - i \ 2, h \ 2 - i \ 2)-(h \ 2 + i \ 2, h \ 2 + i \ 2), , B
        Next
        DoEvents
        Loop

Case 5:

        Do While brun
        change
        For i = 0 To h
        gradient.ForeColor = fire(i Mod 128)
        gradient.Line (i / 2, i / 2)-(h - i / 2, h - i / 2), , B
        Next
        DoEvents
        Loop

Case 6:
        Do While brun
        change
        For i = 0 To h
        gradient.ForeColor = fire(i Mod 128)
        gradient.Line (0, 0)-(i \ 2, i \ 2), , B
        gradient.Line (h, h)-(h - i \ 2, h - i \ 2), , B
        gradient.Line (h, 0)-(h - i \ 2, i \ 2), , B
        gradient.Line (0, h)-(i \ 2, h - i \ 2), , B
        Next
        DoEvents
        Loop
Case 7:
        Do While brun
        change
        For i = 0 To h
        gradient.ForeColor = fire((h - i) Mod 128)
        gradient.Line (0, 0)-(i \ 2, i \ 2), , B
        gradient.Line (h, h)-(h - i \ 2, h - i \ 2), , B
        gradient.Line (h, 0)-(h - i \ 2, i \ 2), , B
        gradient.Line (0, h)-(i \ 2, h - i \ 2), , B
        Next
        DoEvents
        Loop

Case 8:
        Do While brun
        change
        For i = 0 To h
        gradient.ForeColor = fire(i Mod 128)
        gradient.Line (h \ 4 - i \ 4, h \ 4 - i \ 4)-(h \ 4 + i \ 4, h \ 4 + i \ 4), , B
        gradient.Line (3 * h \ 4 - i / 4, h \ 4 - i / 4)-(3 * h \ 4 + i / 4, h \ 4 + i / 4), , B
        gradient.Line (h \ 4 - i / 4, 3 * h \ 4 - i / 4)-(h \ 4 + i / 4, 3 * h \ 4 + i / 4), , B
        gradient.Line (3 * h / 4 - i / 4, 3 * h \ 4 - i / 4)-(3 * h \ 4 + i / 4, 3 * h \ 4 + i / 4), , B
        Next
        DoEvents
        Loop
        
Case 9:
        Do While brun
        change
        For i = 0 To h
        gradient.ForeColor = fire(i Mod 128)
        gradient.Line (i / 4, i / 4)-(h \ 2 - i / 4, h \ 2 - i / 4), , B
        gradient.Line (h \ 2 + i / 4, i / 4)-(h - i / 4, h \ 2 - i / 4), , B
        gradient.Line (i / 4, h \ 2 + i / 4)-(h \ 2 - i / 4, h - i / 4), , B
        gradient.Line (h \ 2 + i / 4, h \ 2 + i / 4)-(h - i / 4, h - i / 4), , B
        Next
        DoEvents
        Loop

Case 10:
        Do While brun
        change
        For i = 0 To h
        If i >= 10 Then
        gradient.Line (h \ 2 - i \ 2, i)-(h \ 2 + i \ 2, i), fire(i Mod 128)
        End If
        Next
        DoEvents
        Loop

Case 11:
        Do While brun
        change
        For i = 0 To h
        If i >= 10 Then
        gradient.ForeColor = fire(i Mod 128)
        gradient.Line (h \ 2 - i \ 2, i)-(h \ 2 + i \ 2, i)
        gradient.Line (h \ 2 - i \ 2, h - i)-(h \ 2 + i \ 2, h - i)
        End If
        Next
        DoEvents
        Loop

Case 12:
        Do While brun
        change
        For i = 0 To h
        If i >= 10 Then
        gradient.Line (h - i, h \ 2 - i \ 2)-(h - i, h \ 2 + i \ 2), fire(i Mod 128)
        End If
        Next
        DoEvents
        Loop

Case 13:
        Do While brun
        change
        For i = 0 To h
        If i >= 10 Then
        gradient.ForeColor = fire(i Mod 128)
        gradient.Line (i, h \ 2 - i \ 2)-(i, h \ 2 + i \ 2)
        gradient.Line (h - i, h \ 2 - i \ 2)-(h - i, h \ 2 + i \ 2)
        End If
        Next
        DoEvents
        Loop

Case 14:
        Do While brun
        change
        For i = 0 To h
        gradient.Line (i, i)-(h - i, i), fire(i Mod 128)
        Next
        DoEvents
        Loop

Case 15:
        Do While brun
        change
        For i = 0 To h
        gradient.Line (i, i)-(i, h - i), fire(i Mod 128)
        Next
        DoEvents
        Loop

Case 16:
        Do While brun
        change
        For i = 0 To h
        gradient.ForeColor = fire(i Mod 128)
        gradient.Line (0, 0)-(i, h)
        gradient.Line (0, 0)-(h, h - i)
        Next
        DoEvents
        Loop

Case 17:
        Do While brun
        change
        For i = 0 To h
        gradient.Line (0, i)-(i, 0), fire(i Mod 128)
        gradient.Line (h, i)-(i, h), fire(i Mod 128)
        Next
        DoEvents
        Loop

Case 18:
        Dim hl As Long
        hl = gradient.hDC
        'gradient.FillStyle = 0
        Do While brun
        change
        For i = 0 To h
        gradient.ForeColor = fire(128 - i Mod 128)
        gradient.FillColor = fire(128 - i Mod 128)
        'Ellipse hl, h \ 2 - i \ 2, h \ 2 - i \ 2, h \ 2 + i \ 2, h \ 2 + i \ 2
        gradient.Circle (h \ 2, h \ 2), h \ 2 - i \ 2
        Next
        'gradient(18).Refresh
        DoEvents
        Loop
       'gradient.FillStyle = 1
Case 19:
        'gradient.FillStyle = 0
        Do While brun
        change
        For i = 0 To h
        gradient.ForeColor = fire(i Mod 128)
        gradient.FillColor = fire(i Mod 128)
        gradient.Circle (h \ 2, h \ 2), h \ 2 - i \ 2
        Next

        DoEvents
        Loop
        
Case 20:
        'gradient.FillStyle = 0
        Do While brun
        change
        For i = 0 To h

        gradient.ForeColor = fire(i Mod 128)

        gradient.Circle (h \ 2, h \ 4), i / 2, , , , 0.5

        gradient.Circle (h \ 2, 3 * h \ 4), i / 2, , , , 0.5

        'gradient(20).Circle (32, 64), i / 2, , , , 2

        'gradient(20).Circle (96, 64), i / 2, , , , 2

        Next

        DoEvents
        Loop

Case 21:
        Do While brun
        change
        For i = 0 To h

        gradient.ForeColor = fire(i Mod 128)

        gradient.Circle (h \ 4, h \ 2), i / 2, , , , 2

        gradient.Circle (3 * h \ 4, h \ 2), i / 2, , , , 2
        Next
        DoEvents
        Loop

Case 22:
        'gradient.FillStyle = 0
        Do While brun
        change
        For i = 0 To h
        
        gradient.ForeColor = fire(i Mod 128)
        gradient.Line (i, 0)-(h - i, h)
        gradient.Line (0, h - i)-(h, i)

        Next

        DoEvents
        Loop
        
Case 23:
        Do While brun
        change
        For i = 0 To h
        gradient.ForeColor = fire(i Mod 128)
        gradient.Line (i, 0)-(h - i, h)
        gradient.Line (0, h - i)-(h, i), fire(128 - i Mod 128)
        Next

        DoEvents
        Loop
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
brun = False
End Sub
