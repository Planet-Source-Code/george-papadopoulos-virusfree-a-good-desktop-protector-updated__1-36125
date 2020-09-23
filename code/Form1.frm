VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   3120
   ClientTop       =   2850
   ClientWidth     =   10785
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   450
      Left            =   3360
      Top             =   4560
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   252
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10815
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Desktop Protector - By George Papapadopouos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   6015
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Caption         =   " OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   7920
      TabIndex        =   5
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Access Denied"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Pass :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   3840
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stopit As Boolean
Const pass = "abc" 'THIS IS THE PASSWORD
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        If LCase(Label4.Tag) = LCase(pass) Then
            stopit = True
            SystemParametersInfo 97, False, Waste, 0
            showmouse
            DoEvents
            Unload Me
            End
        Else
            Label4.Caption = ""
            Label4.Tag = ""
            Timer1.Enabled = True
        End If
    Case 8
        If Len(Label4.Caption) > 0 Then
            Label4.Caption = Mid(Label4.Caption, 1, Len(Label4.Caption) - 1)
            Label4.Tag = Mid(Label4.Tag, 1, Len(Label4.Tag) - 1)
        End If
Case Else
    Label4.Caption = Label4.Caption & "*"
    Label4.Tag = Label4.Tag & LCase(Chr(KeyCode))
End Select


End Sub

Private Sub Form_Load()
On Error Resume Next
Randomize
MsgBox "The Password Is " & pass 'Remove This Command
App.Title = ""  'this hides our program from the XP,NT,2k endtask list

'this remove our program from the windows 9x endtask list
Dim process As Long
process = GetCurrentProcessId()
Call RegisterServiceProcess(process, RSP_SIMPLE_SERVICE)

'set it to realtime priority so windows endtask won't allow closing it
Call SetPriorityClass(GetCurrentProcess(), HIGH_PRIORITY_CLASS)


Me.Show
'stay ontop
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
stopit = False
SystemParametersInfo 97, True, Waste, 0
hidemouse
Me.Show
Do
'make sure the mouse is stucked in the middle of the form
DoEvents
SetCursorPos (Me.Left / Screen.TwipsPerPixelX) + (Me.Width / 2) / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY + (Me.Height / 2) / Screen.TwipsPerPixelY
DoEvents
SetCursorPos (Me.Left / Screen.TwipsPerPixelX) + (Me.Width / 2) / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY + (Me.Height / 2) / Screen.TwipsPerPixelY
SetCursorPos (Me.Left / Screen.TwipsPerPixelX) + (Me.Width / 2) / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY + (Me.Height / 2) / Screen.TwipsPerPixelY
DoEvents
SetCursorPos (Me.Left / Screen.TwipsPerPixelX) + (Me.Width / 2) / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY + (Me.Height / 2) / Screen.TwipsPerPixelY
SetCursorPos (Me.Left / Screen.TwipsPerPixelX) + (Me.Width / 2) / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY + (Me.Height / 2) / Screen.TwipsPerPixelY
DoEvents
SetCursorPos (Me.Left / Screen.TwipsPerPixelX) + (Me.Width / 2) / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY + (Me.Height / 2) / Screen.TwipsPerPixelY
SetCursorPos (Me.Left / Screen.TwipsPerPixelX) + (Me.Width / 2) / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY + (Me.Height / 2) / Screen.TwipsPerPixelY
DoEvents
SetCursorPos (Me.Left / Screen.TwipsPerPixelX) + (Me.Width / 2) / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY + (Me.Height / 2) / Screen.TwipsPerPixelY
SetCursorPos (Me.Left / Screen.TwipsPerPixelX) + (Me.Width / 2) / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY + (Me.Height / 2) / Screen.TwipsPerPixelY
DoEvents
SetCursorPos (Me.Left / Screen.TwipsPerPixelX) + (Me.Width / 2) / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY + (Me.Height / 2) / Screen.TwipsPerPixelY
SetCursorPos (Me.Left / Screen.TwipsPerPixelX) + (Me.Width / 2) / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY + (Me.Height / 2) / Screen.TwipsPerPixelY
Loop Until stopit = True
End Sub



Private Sub Form_Unload(Cancel As Integer)
If stopit = False Then Cancel = 1
End Sub



Private Sub Timer1_Timer()
Label3.Visible = Not (Label3.Visible)
End Sub

Private Sub Timer2_Timer()

End Sub
