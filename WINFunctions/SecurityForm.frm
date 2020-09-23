VERSION 5.00
Begin VB.Form SecurityForm 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-Security-"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox WrongCounter 
      Height          =   375
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "3"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox PasswordText 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Enter the Password!"
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton ExitButton 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1680
      MousePointer    =   10  'Up Arrow
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton EnterButton 
      Caption         =   "Enter"
      Height          =   495
      Left            =   120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "SecurityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------
Rem Made ßy: Mike Canejo
Rem E-mail me at: IamMikeC@aol.com
'--------------------------------------------------------------------
Private Sub EnterButton_Click() 'if the Enter button is clicked then do the below
If PasswordText = LCase$("password") Then
'If the password is X then do something
MsgBox "Correct password!", vbSystemModal + vbInformation + vbOKOnly, "Valid"
'If its right, display this message
MsgBox "Put your code here to do something if its right", vbSystemModal + vbInformation + vbOKOnly, "Command here" 'replace this line with your choice if it's right
Else
If WrongCounter = "1" Then
'if the counter is 1 then do this
This& = MsgBox("Your system will now crash!", vbSystemModal + vbOKCancel, "Confirm")
'display message
If This& = vbCancel Then MsgBox "-Sorry...You cant Cancel!", vbSystemModal + vbCritical + vbOKOnly, "Error"
'if the user clicks Cancel then do this
MsgBox "Your System is now crashing. Please wait..", vbSystemModal + vbInformation + vbOKOnly, "Info"
'displays message'i did this just for a prank or joke..it doesnt really harm your computer in any way :)

'ScreenBlackOut Me
'WinShutdown
'WinReboot
'WinForceClose'---------------------------[Pick a command to use]
'WinLogUserOff
'HideTaskbar
'HideWindowsToolBar
'HideStartButton
'Use one of these Commands if the password is wrong 3 times

WrongCounter = "3" 'Counter = "3" to start over
End 'Ends program
End If 'Ends the If
WrongCounter = Val(WrongCounter) - 1 'make the counter 1 number less than it was
If WrongCounter = "1" Then 'i used this to make the message make sense
MsgBox "The password you've entered is invalid! You have " & WrongCounter & " try left until System Failure!", vbSystemModal + vbCritical + vbOKOnly, "Error"
'displays error message
Else
MsgBox "The password you've entered is invalid! You have " & WrongCounter & " tries left until System Failure!", vbSystemModal + vbCritical + vbOKOnly, "Error"
'displays error message
End If 'End the If
End If 'End the If
End Sub

Private Sub ExitButton_Click() 'If the Exit button is clciked then do the below
Unload Me 'Unload the program
End Sub

Private Sub Form_Load() ' when the form loads, do this
CenterForm Me 'Centers the form in the screen
StayOnTop Me 'Keeps the form OnTop of everything
'PreventFromClosing
'DisableCtrlAltDel
'EnableCtrlAltDel
Rem Use a command for the loading part
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) 'when the user clicks unload or the X button
'use this incase the user clicks the unload from the taskbar
Value& = MsgBox("Are you sure you want to exit?", vbSystemModal + vbInformation + vbYesNo, "Exit")
'hold outcome of the message
If Value& = vbYes Then 'if the user clicks Yes then do something
End 'do this when the user clicks Yes
End If 'Ends the If
Cancel = 1 'makes the form not unload automatically ..only unloads from the above Yes or No statement
End Sub

Private Sub PasswordText_Click() 'if the textbox is clicked then check for text
If PasswordText = "Enter the Password!" Then PasswordText = ""
'clears the textbox if it has the X text in it
End Sub
