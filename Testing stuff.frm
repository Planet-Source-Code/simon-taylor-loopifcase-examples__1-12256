VERSION 5.00
Begin VB.Form frmTesting 
   Caption         =   "Tester stuff"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdIfState 
      Caption         =   "If Else If Statement"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton CmddoWhile 
      Caption         =   "Do while loop"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdCaseState 
      Caption         =   "Case Statement"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdForLoop 
      Caption         =   "For Loop"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdDoLoop 
      Caption         =   "Do Loop"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCaseState_Click()
Dim Age As Integer


Age = InputBox("How old are you?")

Select Case Age


Case Age = 1 To 11
    MsgBox " To young to do shit", vbInformation + vbOKOnly

Case Age = 12 To 17
    MsgBox " If your a girl you could probably get away with drinking in Ashton!", vbExclamation, vbOKOnly

Case Age = 17 To 21
    MsgBox "Ahh 17 - 21, Nice age to be at", vbCritical

Case Age = 21 To 30
    MsgBox "Your getting passed it now", vbInformation + vbAbortRetryIgnore
    
Case Age = 31 To 100
    MsgBox "It's all down hill from here matey", vbInformation + vbCritical
    
Case Else
    MsgBox "Somethings gone wrong, are you alive?", vbInformation
    
End Select


End Sub

Private Sub cmdDoLoop_Click()
Dim msg As String

Do
msg = InputBox("Type 'End' to stop the loop")

If msg = "End" Then

Exit Do

End If

Loop

MsgBox "You just entered : " & msg, vbInformation + vbOKOnly
End Sub

Private Sub CmddoWhile_Click()

Calories = InputBox("How many calories have you eaten")


Do While Calories < 3200

MsgBox " Keep on eating, you have eaten " & Calories, vbCritical


Calories = Calories * 2


Loop

MsgBox "Start Dieting Fatty!!! You have gained :" & Calories & " Calories", vbInformation


End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdForLoop_Click()
Dim Counter As Integer

For I = 1 To 15
Counter = I

MsgBox "This has repeated " & Counter & " times", vbInformation + vbOKOnly

Next

End Sub

Private Sub CmdIfState_Click()
Dim Message As String

Message = InputBox("Do you like School")

If Message = "Yes" Then

MsgBox "You sad monkey", vbCritical + vbOKOnly

ElseIf Message = "No" Then

MsgBox "Congratulations, your a cool dude", vbInformation + vbOKOnly

Else

MsgBox "Errr, Enter 'Yes' or 'No' please!", vbInformation

End If




End Sub
