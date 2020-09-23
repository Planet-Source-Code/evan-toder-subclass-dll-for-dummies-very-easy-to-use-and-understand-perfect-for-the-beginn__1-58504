Attribute VB_Name = "code"
Option Explicit





Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public PrevProc As Long

Public arr_msg()       As Variant
Public calling_class   As cSubclassMsg
Public tag_all_msgs    As Boolean

Public Sub HookForm(your_hwnd As Long)
    
    On Error Resume Next
    
    PrevProc = SetWindowLong(your_hwnd, GWL_WNDPROC, AddressOf WindowProc)
    
End Sub
Public Sub UnHookForm(your_hwnd As Long)

    On Error Resume Next
    
    SetWindowLong your_hwnd, GWL_WNDPROC, PrevProc
    
End Sub
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error Resume Next
    
    Dim l          As Long
    Dim upp        As Long
    Dim bdiscard   As Boolean
    
    
   If tag_all_msgs = True Then
       'if were intercepting all messages obviously were not going to
       'discard them all
       WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
       calling_class.friend_event_notify uMsg, wParam, lParam, bdiscard
   Else
       upp = UBound(arr_msg, 2)
    
       'search the array of messages you set in form_load
       'and see if this is one of those msges
       For l = 0 To upp
         If uMsg = arr_msg(0, l) Then
           'is it a message we want to discard
           If arr_msg(1, l) Then
              bdiscard = True
              'create the event
              calling_class.friend_event_notify uMsg, wParam, lParam, bdiscard
              Exit For
           Else
             'create the event
              calling_class.friend_event_notify uMsg, wParam, lParam, bdiscard
           End If
         End If
       Next l
    
       'if were discarding the message dont pass it on
       If bdiscard Then
          WindowProc = 0
       Else
          'pass the message on to where its suppose to go
          WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
       End If
    End If
    
End Function



'-- check to see if item is array or array is initialized
Function IsArray(varArray As Variant) As Boolean

Dim Upper As Integer
On Error Resume Next
 
  Upper = UBound(varArray)
  
  If Err.Number Then
     If Err.Number = 9 Then
       IsArray = False
     End If
  Else
     IsArray = True
  End If

End Function



 
