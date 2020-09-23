VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculator"
   ClientHeight    =   4200
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   2325
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtResult 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sqr"
      Height          =   495
      Left            =   1200
      TabIndex        =   19
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "%"
      Height          =   495
      Left            =   600
      TabIndex        =   18
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton CmdNegative 
      Caption         =   "-/+"
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox TxtInputNum 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      MaxLength       =   17
      TabIndex        =   16
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Signs 
      Caption         =   "/"
      Height          =   495
      Index           =   3
      Left            =   1800
      TabIndex        =   15
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Signs 
      Caption         =   "*"
      Height          =   495
      Index           =   2
      Left            =   1800
      TabIndex        =   14
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Signs 
      Caption         =   "-"
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Signs 
      Caption         =   "+"
      Height          =   495
      Index           =   0
      Left            =   1800
      TabIndex        =   12
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "="
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "."
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   1200
      TabIndex        =   9
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   600
      TabIndex        =   8
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   3000
      Width           =   495
   End
   Begin VB.Menu mnuClear 
      Caption         =   "Clear"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I've just seen heaps of calculators on this site with heaps
'and heaps of unnecessary code so I decided to make my own
'which is very simple.
'If you have any comments like - something that I could have
'done better or more efficiently. Then please tell me because I am
'still learning to programme and need feedback. I have only been coding for a
'month so go easy if ya decide to comment.
Option Explicit
    Dim DecFlag     As Boolean 'Decimal Flag
    Dim Calculate   As Boolean 'Calculate flag
    Dim Opt1        As Variant 'First Number
    Dim Opt2        As Variant 'Second Number
    Dim StrSign     As String * 1 'A string that can only accept 1 char
Private Sub cmdCalculate_Click() 'If the equals button was pressed
     If StrSign = "+" Then
        TxtInputNum.Text = Val(Opt1) + Val(TxtInputNum.Text)
     ElseIf StrSign = "-" Then
        TxtInputNum.Text = Val(Opt1) - Val(TxtInputNum.Text)
     ElseIf StrSign = "*" Then
        TxtInputNum.Text = Val(Opt1) * Val(TxtInputNum.Text)
     ElseIf StrSign = "/" Then
        TxtInputNum.Text = Val(Opt1) / Val(TxtInputNum.Text)
     End If
     
     If Calculate = True Then Call mnuClear_Click
     Calculate = True
End Sub
Private Sub cmdDecimal_Click()
    If DecFlag = True Then 'Check if a decimal has already being pressed
        Exit Sub
    Else
        TxtInputNum.Text = TxtInputNum.Text & "."
        DecFlag = True
    End If
End Sub
Private Sub CmdNegative_Click()
   TxtInputNum.Text = Val(TxtInputNum) * -1 'Nice trick eh
End Sub

Private Sub Command1_Click()
    TxtInputNum.Text = Val(TxtInputNum.Text) / 100
End Sub

Private Sub Command2_Click()
    TxtInputNum.Text = Sqr(TxtInputNum.Text)
End Sub

Private Sub mnuClear_Click() 'Clear everything
    DecFlag = False
    Opt1 = ""
    Opt2 = ""
    StrSign = ""
    TxtInputNum.Text = ""
    TxtResult = ""
End Sub

Private Sub Number_Click(Index As Integer) 'Put a corresponding number in the text box
    If TxtInputNum.Text = "" And Number(0) Then Exit Sub
    TxtInputNum.Text = TxtInputNum.Text & Number(Index).Caption
End Sub

Private Sub Signs_Click(Index As Integer) 'Put and calculate for the sign
    Calculate = False
    Select Case Signs(Index)
        Case Signs(0) 'Addition
                StrSign = "+"
                If Opt1 = "" Then
                    Opt1 = TxtInputNum.Text
                Else
                    Opt2 = Val(Opt1) + Val(TxtInputNum.Text)
                    Opt1 = Opt2
                End If
                TxtInputNum.Text = ""
         Case Signs(1) 'Subtraction
                StrSign = "-"
                If Opt1 = "" Then
                    Opt1 = TxtInputNum.Text
                Else
                    Opt2 = Val(Opt1) - Val(TxtInputNum.Text)
                    Opt1 = Opt2
                End If
                TxtInputNum.Text = ""
        Case Signs(2) 'Multiplication
                StrSign = "*"
                If Opt1 = "" Then
                    Opt1 = TxtInputNum.Text
                Else
                    Opt2 = Val(Opt1) * Val(TxtInputNum.Text)
                    Opt1 = Opt2
                End If
                TxtInputNum.Text = ""
        Case Signs(3) 'Division
                StrSign = "/"
                If Opt1 = "" Then
                    Opt1 = TxtInputNum.Text
                Else
                    Opt2 = Val(Opt1) / Val(TxtInputNum.Text)
                    Opt1 = Opt2
                End If
                TxtInputNum.Text = ""
    End Select
    TxtResult.Text = Opt1
End Sub
