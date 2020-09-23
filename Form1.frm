VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   180
      TabIndex        =   0
      Top             =   1035
      Width           =   1410
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents cM   As cSubclassMsg
Attribute cM.VB_VarHelpID = -1


Private Sub cM_Error(errDescription As String)

 MsgBox errDescription
 
End Sub

Private Sub cM_msgString(strMsg As String, wParam As Long, lParam As Long, bdiscard_msg As Boolean)

 Debug.Print strMsg

End Sub

Private Sub Command1_Click()
  
  cM.about_help
  
End Sub

Private Sub Form_Load()
  
  Set cM = New cSubclassMsg
  cM.add_msg_to_track aWM_ALL_MESSAGES
  cM.StartSubclass hwnd, True
 
End Sub

Private Sub Form_Terminate()
  
  cM.clean_up
  Set cM = Nothing
  
End Sub
