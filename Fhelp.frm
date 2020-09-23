VERSION 5.00
Begin VB.Form Fhelp 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   2535
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   -90
      Visible         =   0   'False
      Width           =   3300
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   360
         Width           =   45
      End
      Begin VB.Image Image1 
         Height          =   210
         Left            =   0
         Picture         =   "Fhelp.frx":0000
         Top             =   90
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   3300
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Text            =   "select a topic"
         Top             =   315
         Width           =   3165
      End
   End
End
Attribute VB_Name = "Fhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 

Dim help_info(2)    As String


Private Sub Combo1_Click()

  Label1 = FormatLineLen(help_info(Combo1.ListIndex), 32)
  Frame1(1).Visible = True
  Height = Label1.Height + 750
  Frame1(1).Height = Label1.Height + 450
  
End Sub



 
Private Sub Form_Load()

  Label1.AutoSize = True
  'add help topics to combo box
  With Combo1
     .AddItem "[method] add_msg_to_track"
     .AddItem "[method] clean_up"
     .AddItem "[method] StartSubclass"
  End With
 
  'specify the strings for the help items
  'the index matches the listindex of the combobox
  help_info(0) = "Using this method you select the messages that you wish to track/intercept.  Call this method for each message you wish to track.  When this message is intercepted it will trigger either the [cMsg_msgLong] event  or  [cMsg_msgString] event depending your setting for  [show_msg_as_stringconst] in the [start_subclass] method"
  help_info(1) = "Call this method from your [Form_Unload] or [Form_Terminate] subroutine.  Calling this method safely transfers control of messages directly back to your form for safe un-subclassing.  If you forget to do this, cleanup code is called, by default, in the classes [Terminate] event however it is a good practice, and safer to call this method yourself"
  help_info(2) = "This starts the subclass process. Before the subclass actually starts this method checks to make sure that you specified at least one message to intercept. If at least one wasnt specified and error is raised in the [Error] event. NOTE: IF YOU SPECIFY TO SUBCLASS ALL MESSAGES..DO NOT PLACE THE RETURN MESSAGE IN A LABEL, OR A FORMS CAPTION OR ANY CONTROL.  IF YOU DESIRE THE VISUAL FEEDBACK OF THE MESSAGES USE THE DEBUG WINDOW ONLY!!"
  
End Sub

Private Sub Image1_Click()
 
  Frame1(1).Visible = False
  Height = Frame1(0).Height + 300

End Sub
 
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

 Image1.BorderStyle = 1

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 Image1.BorderStyle = 0

End Sub

 

'###############################################################
'PURPOSE:
'
'     take a string of any length, and format it as paragraph with
'     a maximum character width [LINE_LEN], and optionally add a
'     a left spacing or padding to each line...this comment block
'     was formatted with this function...see how neat and organized
'     it looks :)
'
'     PART             DESCRIPTION
'     -----------------------------------------------------------
'     strToFormat      [required | string]
'                      The string to input that will be formatted
'     -----------------------------------------------------------
'     LINE_LEN         [optional | long]
'                      The maximum allowable character length of
'                      any line. (30=default)
'     -----------------------------------------------------------
'     LeftPad          [optional | long]
'                      How many spaces to left pad each returned
'                      line with
'     -----------------------------------------------------------
'
Private Function FormatLineLen(strToFormat$, Optional LINE_LEN& = 30, _
                                   Optional LeftPad As Long) As String
 
    Dim sparts() As String
    Dim tempLine$, strPad$
    Dim lcnt&
    
    On Error GoTo Err_Handler:
    
    strPad = String(LeftPad, " ")
   'break strToFormat$ word by word
1    sparts = Split(strToFormat)
2    For lcnt = 0 To UBound(sparts)
       'if the linelen including this spart < max line len
3       If (Len(tempLine$ & sparts(lcnt))) <= LINE_LEN& Then
4          tempLine$ = (tempLine$ & sparts(lcnt) & " ")
        Else
5          FormatLineLen = (FormatLineLen & strPad & tempLine$ & vbCrLf)
6          tempLine = sparts(lcnt) & " "
        End If
     Next lcnt
7    FormatLineLen = (FormatLineLen & strPad & tempLine)
    
    Exit Function
Err_Handler:
   Debug.Print Err.Number & vbTab & "  line: " & Erl() & vbTab & Err.Description
End Function











