VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubclassMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 


Enum enMsg
  aWM_ALL_MESSAGES = -1
  WM_USER = &H400
  WM_CAP_START = WM_USER
  WM_ACTIVATE = &H6
  WM_ACTIVATEAPP = &H1C
  WM_ADSPROP_NOTIFY_APPLY = (WM_USER + 1104)
  WM_ADSPROP_NOTIFY_CHANGE = (WM_USER + 1103)
  WM_ADSPROP_NOTIFY_ERROR = (WM_USER + 1110)
  WM_ADSPROP_NOTIFY_EXIT = (WM_USER + 1107)
  WM_ADSPROP_NOTIFY_FOREGROUND = (WM_USER + 1106)
  WM_ADSPROP_NOTIFY_PAGEHWND = (WM_USER + 1102)
  WM_ADSPROP_NOTIFY_PAGEINIT = (WM_USER + 1101)
  WM_ADSPROP_NOTIFY_SETFOCUS = (WM_USER + 1105)
  WM_ADSPROP_NOTIFY_SHOW_ERROR_DIALOG = (WM_USER + 1111)
  WM_AFXFIRST = &H360
  WM_AFXLAST = &H37F
  WM_APP = &H8000
  WM_APPCOMMAND = &H319
  WM_ASKCBFORMATNAME = &H30C
  WM_CANCELJOURNAL = &H4B
  WM_CANCELMODE = &H1F
  WM_CAPTURECHANGED = &H215
  WM_CHANGECBCHAIN = &H30D
  WM_CHANGEUISTATE = &H127
  WM_CHAR = &H102
  WM_CHARTOITEM = &H2F
  WM_CHILDACTIVATE = &H22
  WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
  WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)
  WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
  WM_CLEAR = &H303
  WM_CLOSE = &H10
  WM_COMMAND = &H111
  WM_COMMNOTIFY = &H44
  WM_COMPACTING = &H41
  WM_COMPAREITEM = &H39
  WM_CONTEXTMENU = &H7B
  WM_CONVERTREQUEST = &H10A
  WM_CONVERTREQUESTEX = &H108
  WM_CONVERTRESULT = &H10B
  WM_COPY = &H301
  WM_COPYDATA = &H4A
  WM_CPL_LAUNCH = (WM_USER + 1000)
  WM_CPL_LAUNCHED = (WM_USER + 1001)
  WM_CREATE = &H1
  WM_CTLCOLOR = &H19
  WM_CTLCOLORBTN = &H135
  WM_CTLCOLORDLG = &H136
  WM_CTLCOLOREDIT = &H133
  WM_CTLCOLORLISTBOX = &H134
  WM_CTLCOLORMSGBOX = &H132
  WM_CTLCOLORSCROLLBAR = &H137
  WM_CTLCOLORSTATIC = &H138
  WM_CUT = &H300
  WM_DDE_FIRST = &H3E0
  WM_DDE_ACK = (WM_DDE_FIRST + 4)
  WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
  WM_DDE_DATA = (WM_DDE_FIRST + 5)
  WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
  WM_DDE_INITIATE = (WM_DDE_FIRST)
  WM_DDE_LAST = (WM_DDE_FIRST + 8)
  WM_DDE_POKE = (WM_DDE_FIRST + 7)
  WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
  WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
  WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
  WM_DEADCHAR = &H103
  WM_DELETEITEM = &H2D
  WM_DESTROY = &H2
  WM_DESTROYCLIPBOARD = &H307
  WM_DEVICECHANGE = &H219
  WM_DEVMODECHANGE = &H1B
  WM_DISPLAYCHANGE = &H7E
  WM_DRAWCLIPBOARD = &H308
  WM_DRAWITEM = &H2B
  WM_DROPFILES = &H233
  WM_ENABLE = &HA
  WM_ENDSESSION = &H16
  WM_ENTERIDLE = &H121
  WM_ENTERMENULOOP = &H211
  WM_ENTERSIZEMOVE = &H231
  WM_ERASEBKGND = &H14
  WM_EXITMENULOOP = &H212
  WM_EXITSIZEMOVE = &H232
  WM_FONTCHANGE = &H1D
  WM_FORWARDMSG = &H37F
  WM_GETDLGCODE = &H87
  WM_GETFONT = &H31
  WM_GETHOTKEY = &H33
  WM_GETICON = &H7F
  WM_GETMINMAXINFO = &H24
  WM_GETOBJECT = &H3D
  WM_GETTEXT = &HD
  WM_GETTEXTLENGTH = &HE
  WM_HANDHELDFIRST = &H358
  WM_HANDHELDLAST = &H35F
  WM_HELP = &H53
  WM_HOTKEY = &H312
  WM_HSCROLL = &H114
  WM_HSCROLLCLIPBOARD = &H30E
  WM_ICONERASEBKGND = &H27
  WM_IME_CHAR = &H286
  WM_IME_COMPOSITION = &H10F
  WM_IME_COMPOSITIONFULL = &H284
  WM_IME_CONTROL = &H283
  WM_IME_ENDCOMPOSITION = &H10E
  WM_IME_KEYDOWN = &H290
  WM_IME_KEYLAST = &H10F
  WM_IME_KEYUP = &H291
  WM_IME_NOTIFY = &H282
  WM_IME_REPORT = &H280
  WM_IME_REQUEST = &H288
  WM_IME_SELECT = &H285
  WM_IME_SETCONTEXT = &H281
  WM_IME_STARTCOMPOSITION = &H10D
  WM_IMEKEYDOWN = &H290
  WM_IMEKEYUP = &H291
  WM_INITDIALOG = &H110
  WM_INITMENU = &H116
  WM_INITMENUPOPUP = &H117
  WM_INPUTLANGCHANGE = &H51
  WM_INPUTLANGCHANGEREQUEST = &H50
  WM_INTERIM = &H10C
  WM_KEYDOWN = &H100
  WM_KEYFIRST = &H100
  WM_KEYLAST = &H108
  WM_KEYUP = &H101
  WM_KILLFOCUS = &H8
  WM_LBUTTONDBLCLK = &H203
  WM_LBUTTONDOWN = &H201
  WM_LBUTTONUP = &H202
  WM_MBUTTONDBLCLK = &H209
  WM_MBUTTONDOWN = &H207
  WM_MBUTTONUP = &H208
  WM_MDIACTIVATE = &H222
  WM_MDICASCADE = &H227
  WM_MDICREATE = &H220
  WM_MDIDESTROY = &H221
  WM_MDIGETACTIVE = &H229
  WM_MDIICONARRANGE = &H228
  WM_MDIMAXIMIZE = &H225
  WM_MDINEXT = &H224
  WM_MDIREFRESHMENU = &H234
  WM_MDIRESTORE = &H223
  WM_MDISETMENU = &H230
  WM_MDITILE = &H226
  WM_MEASUREITEM = &H2C
  WM_MENUCHAR = &H120
  WM_MENUCOMMAND = &H126
  WM_MENUDRAG = &H123
  WM_MENUGETOBJECT = &H124
  WM_MENURBUTTONUP = &H122
  WM_MENUSELECT = &H11F
  WM_MOUSEACTIVATE = &H21
  WM_MOUSEFIRST = &H200
  WM_MOUSEHOVER = &H2A1
  WM_MOUSELAST = &H209
  WM_MOUSELEAVE = &H2A3
  WM_MOUSEMOVE = &H200
  WM_MOUSEWHEEL = &H20A
  WM_MOVE = &H3
  WM_MOVING = &H216
  WM_NCACTIVATE = &H86
  WM_NCCALCSIZE = &H83
  WM_NCCREATE = &H81
  WM_NCDESTROY = &H82
  WM_NCHITTEST = &H84
  WM_NCLBUTTONDBLCLK = &HA3
  WM_NCLBUTTONDOWN = &HA1
  WM_NCLBUTTONUP = &HA2
  WM_NCMBUTTONDBLCLK = &HA9
  WM_NCMBUTTONDOWN = &HA7
  WM_NCMBUTTONUP = &HA8
  WM_NCMOUSEHOVER = &H2A0
  WM_NCMOUSELEAVE = &H2A2
  WM_NCMOUSEMOVE = &HA0
  WM_NCPAINT = &H85
  WM_NCRBUTTONDBLCLK = &HA6
  WM_NCRBUTTONDOWN = &HA4
  WM_NCRBUTTONUP = &HA5
  WM_NCXBUTTONDBLCLK = &HAD
  WM_NCXBUTTONDOWN = &HAB
  WM_NCXBUTTONUP = &HAC
  WM_NEXTDLGCTL = &H28
  WM_NEXTMENU = &H213
  WM_NOTIFY = &H4E
  WM_NOTIFYFORMAT = &H55
  WM_NULL = &H0
  WM_OTHERWINDOWCREATED = &H42
  WM_OTHERWINDOWDESTROYED = &H43
  WM_PAINT = &HF
  WM_PAINTCLIPBOARD = &H309
  WM_PAINTICON = &H26
  WM_PALETTECHANGED = &H311
  WM_PALETTEISCHANGING = &H310
  WM_PARENTNOTIFY = &H210
  WM_PASTE = &H302
  WM_PENWINFIRST = &H380
  WM_PENWINLAST = &H38F
  WM_POWER = &H48
  WM_POWERBROADCAST = &H218
  WM_PRINT = &H317
  WM_PRINTCLIENT = &H318
  WM_QUERYDRAGICON = &H37
  WM_QUERYENDSESSION = &H11
  WM_QUERYNEWPALETTE = &H30F
  WM_QUERYOPEN = &H13
  WM_QUERYUISTATE = &H129
  WM_QUEUESYNC = &H23
  WM_QUIT = &H12
  WM_RASDIALEVENT = &HCCCD
  WM_RBUTTONDBLCLK = &H206
  WM_RBUTTONDOWN = &H204
  WM_RBUTTONUP = &H205
  WM_RENDERALLFORMATS = &H306
  WM_RENDERFORMAT = &H305
  WM_SETCURSOR = &H20
  WM_SETFOCUS = &H7
  WM_SETFONT = &H30
  WM_SETHOTKEY = &H32
  WM_SETICON = &H80
  WM_SETREDRAW = &HB
  WM_WININICHANGE = &H1A
  WM_SETTEXT = &HC
  WM_SETTINGCHANGE = WM_WININICHANGE
  WM_SHOWWINDOW = &H18
  WM_SIZE = &H5
  WM_SIZECLIPBOARD = &H30B
  WM_SIZING = &H214
  WM_SPOOLERSTATUS = &H2A
  WM_STYLECHANGED = &H7D
  WM_STYLECHANGING = &H7C
  WM_SYNCPAINT = &H88
  WM_SYSCHAR = &H106
  WM_SYSCOLORCHANGE = &H15
  WM_SYSCOMMAND = &H112
  WM_SYSDEADCHAR = &H107
  WM_SYSKEYDOWN = &H104
  WM_SYSKEYUP = &H105
  WM_TCARD = &H52
  WM_TIMECHANGE = &H1E
  WM_TIMER = &H113
  WM_UNDO = &H304
  WM_UNINITMENUPOPUP = &H125
  WM_UPDATEUISTATE = &H128
  WM_USERCHANGED = &H54
  WM_VKEYTOITEM = &H2E
  WM_VSCROLL = &H115
  WM_VSCROLLCLIPBOARD = &H30A
  WM_WINDOWPOSCHANGED = &H47
  WM_WINDOWPOSCHANGING = &H46
  WM_WNT_CONVERTREQUESTEX = &H109
  WM_XBUTTONDBLCLK = &H20D
  WM_XBUTTONDOWN = &H20B
  WM_XBUTTONUP = &H20C
End Enum
     
     
     
Dim m_your_hwnd                As Long
Dim m_show_msg_as_stringconst  As Boolean

Event Error(errDescription$)
Event msgLong(uMsg As Long, wParam As Long, lParam As Long, bdiscard_msg As Boolean)
Event msgString(strMsg$, wParam As Long, lParam As Long, bdiscard_msg As Boolean)

Private Function convert_smsg_to_string(ByVal uMsg As Long)
  
  
  Select Case uMsg
      Case Is = WM_USER: convert_smsg_to_string = "WM_USER"
      Case Is = WM_CAP_START: convert_smsg_to_string = "WM_CAP_START"
      Case Is = WM_ACTIVATE: convert_smsg_to_string = "WM_ACTIVATE"
      Case Is = WM_ADSPROP_NOTIFY_APPLY: convert_smsg_to_string = "WM_ADSPROP_NOTIFY_APPLY"
      Case Is = WM_ADSPROP_NOTIFY_CHANGE: convert_smsg_to_string = "WM_ADSPROP_NOTIFY_CHANGE"
      Case Is = WM_ADSPROP_NOTIFY_ERROR: convert_smsg_to_string = "WM_ADSPROP_NOTIFY_ERROR"
      Case Is = WM_ADSPROP_NOTIFY_EXIT: convert_smsg_to_string = "WM_ADSPROP_NOTIFY_EXIT"
      Case Is = WM_ADSPROP_NOTIFY_FOREGROUND: convert_smsg_to_string = "WM_ADSPROP_NOTIFY_FOREGROUND"
      Case Is = WM_ADSPROP_NOTIFY_PAGEHWND: convert_smsg_to_string = "WM_ADSPROP_NOTIFY_PAGEHWND"
      Case Is = WM_ADSPROP_NOTIFY_PAGEINIT: convert_smsg_to_string = "WM_ADSPROP_NOTIFY_PAGEINIT"
      Case Is = WM_ADSPROP_NOTIFY_SETFOCUS: convert_smsg_to_string = "WM_ADSPROP_NOTIFY_SETFOCUS"
      Case Is = WM_ADSPROP_NOTIFY_SHOW_ERROR_DIALOG: convert_smsg_to_string = "WM_ADSPROP_NOTIFY_SHOW_ERROR_DIALOG"
      Case Is = WM_AFXFIRST: convert_smsg_to_string = "WM_AFXFIRST"
      Case Is = WM_AFXLAST: convert_smsg_to_string = "WM_AFXLAST"
      Case Is = WM_APP: convert_smsg_to_string = "WM_APP"
      Case Is = WM_APPCOMMAND: convert_smsg_to_string = "WM_APPCOMMAND"
      Case Is = WM_ASKCBFORMATNAME: convert_smsg_to_string = "WM_ASKCBFORMATNAME"
      Case Is = WM_CANCELJOURNAL: convert_smsg_to_string = "WM_CANCELJOURNAL"
      Case Is = WM_CANCELMODE: convert_smsg_to_string = "WM_CANCELMODE"
      Case Is = WM_CAPTURECHANGED: convert_smsg_to_string = "WM_CAPTURECHANGED"
      Case Is = WM_CHANGECBCHAIN: convert_smsg_to_string = "WM_CHANGECBCHAIN"
      Case Is = WM_CHANGEUISTATE: convert_smsg_to_string = "WM_CHANGEUISTATE"
      Case Is = WM_CHAR: convert_smsg_to_string = "WM_CHAR"
      Case Is = WM_CHARTOITEM: convert_smsg_to_string = "WM_CHARTOITEM"
      Case Is = WM_CHILDACTIVATE: convert_smsg_to_string = "WM_CHILDACTIVATE"
      Case Is = WM_CHOOSEFONT_GETLOGFONT: convert_smsg_to_string = "WM_CHOOSEFONT_GETLOGFONT"
      Case Is = WM_CHOOSEFONT_SETFLAGS: convert_smsg_to_string = "WM_CHOOSEFONT_SETFLAGS"
      Case Is = WM_CHOOSEFONT_SETLOGFONT: convert_smsg_to_string = "WM_CHOOSEFONT_SETLOGFONT"
      Case Is = WM_CLEAR: convert_smsg_to_string = "WM_CLEAR"
      Case Is = WM_CLOSE: convert_smsg_to_string = "WM_CLOSE"
      Case Is = WM_COMMAND: convert_smsg_to_string = "WM_COMMAND"
      Case Is = WM_COMMNOTIFY: convert_smsg_to_string = "WM_COMMNOTIFY"
      Case Is = WM_COMPACTING: convert_smsg_to_string = "WM_COMPACTING"
      Case Is = WM_COMPAREITEM: convert_smsg_to_string = "WM_COMPAREITEM"
      Case Is = WM_CONTEXTMENU: convert_smsg_to_string = "WM_CONTEXTMENU"
      Case Is = WM_CONVERTREQUEST: convert_smsg_to_string = "WM_CONVERTREQUEST"
      Case Is = WM_CONVERTREQUESTEX: convert_smsg_to_string = "WM_CONVERTREQUESTEX"
      Case Is = WM_CONVERTRESULT: convert_smsg_to_string = "WM_CONVERTRESULT"
      Case Is = WM_COPY: convert_smsg_to_string = "WM_COPY"
      Case Is = WM_COPYDATA: convert_smsg_to_string = "WM_COPYDATA"
      Case Is = WM_CPL_LAUNCH: convert_smsg_to_string = "WM_CPL_LAUNCH"
      Case Is = WM_CPL_LAUNCHED: convert_smsg_to_string = "WM_CPL_LAUNCHED"
      Case Is = WM_CREATE: convert_smsg_to_string = "WM_CREATE"
      Case Is = WM_CTLCOLOR: convert_smsg_to_string = "WM_CTLCOLOR"
      Case Is = WM_CTLCOLORBTN: convert_smsg_to_string = "WM_CTLCOLORBTN"
      Case Is = WM_CTLCOLORDLG: convert_smsg_to_string = "WM_CTLCOLORDLG"
      Case Is = WM_CTLCOLOREDIT: convert_smsg_to_string = "WM_CTLCOLOREDIT"
      Case Is = WM_CTLCOLORLISTBOX: convert_smsg_to_string = "WM_CTLCOLORLISTBOX"
      Case Is = WM_CTLCOLORMSGBOX: convert_smsg_to_string = "WM_CTLCOLORMSGBOX"
      Case Is = WM_CTLCOLORSCROLLBAR: convert_smsg_to_string = "WM_CTLCOLORSCROLLBAR"
      Case Is = WM_CTLCOLORSTATIC: convert_smsg_to_string = "WM_CTLCOLORSTATIC"
      Case Is = WM_CUT: convert_smsg_to_string = "WM_CUT"
      Case Is = WM_DDE_FIRST: convert_smsg_to_string = "WM_DDE_FIRST"
      Case Is = WM_DDE_ACK: convert_smsg_to_string = "WM_DDE_ACK"
      Case Is = WM_DDE_ADVISE: convert_smsg_to_string = "WM_DDE_ADVISE"
      Case Is = WM_DDE_DATA: convert_smsg_to_string = "WM_DDE_DATA"
      Case Is = WM_DDE_EXECUTE: convert_smsg_to_string = "WM_DDE_EXECUTE"
      Case Is = WM_DDE_INITIATE: convert_smsg_to_string = "WM_DDE_INITIATE"
      Case Is = WM_DDE_LAST: convert_smsg_to_string = "WM_DDE_LAST"
      Case Is = WM_DDE_POKE: convert_smsg_to_string = "WM_DDE_POKE"
      Case Is = WM_DDE_REQUEST: convert_smsg_to_string = "WM_DDE_REQUEST"
      Case Is = WM_DDE_TERMINATE: convert_smsg_to_string = "WM_DDE_TERMINATE"
      Case Is = WM_DDE_UNADVISE: convert_smsg_to_string = "WM_DDE_UNADVISE"
      Case Is = WM_DEADCHAR: convert_smsg_to_string = "WM_DEADCHAR"
      Case Is = WM_DELETEITEM: convert_smsg_to_string = "WM_DELETEITEM"
      Case Is = WM_DESTROY: convert_smsg_to_string = "WM_DESTROY"
      Case Is = WM_DESTROYCLIPBOARD: convert_smsg_to_string = "WM_DESTROYCLIPBOARD"
      Case Is = WM_DEVICECHANGE: convert_smsg_to_string = "WM_DEVICECHANGE"
      Case Is = WM_DEVMODECHANGE: convert_smsg_to_string = "WM_DEVMODECHANGE"
      Case Is = WM_DISPLAYCHANGE: convert_smsg_to_string = "WM_DISPLAYCHANGE"
      Case Is = WM_DRAWCLIPBOARD: convert_smsg_to_string = "WM_DRAWCLIPBOARD"
      Case Is = WM_DRAWITEM: convert_smsg_to_string = "WM_DRAWITEM"
      Case Is = WM_DROPFILES: convert_smsg_to_string = "WM_DROPFILES"
      Case Is = WM_ENABLE: convert_smsg_to_string = "WM_ENABLE"
      Case Is = WM_ENDSESSION: convert_smsg_to_string = "WM_ENDSESSION"
      Case Is = WM_ENTERIDLE: convert_smsg_to_string = "WM_ENTERIDLE"
      Case Is = WM_ENTERMENULOOP: convert_smsg_to_string = "WM_ENTERMENULOOP"
      Case Is = WM_ENTERSIZEMOVE: convert_smsg_to_string = "WM_ENTERSIZEMOVE"
      Case Is = WM_ERASEBKGND: convert_smsg_to_string = "WM_ERASEBKGND"
      Case Is = WM_EXITMENULOOP: convert_smsg_to_string = "WM_EXITMENULOOP"
      Case Is = WM_EXITSIZEMOVE: convert_smsg_to_string = "WM_EXITSIZEMOVE"
      Case Is = WM_FONTCHANGE: convert_smsg_to_string = "WM_FONTCHANGE"
      Case Is = WM_FORWARDMSG: convert_smsg_to_string = "WM_FORWARDMSG"
      Case Is = WM_GETDLGCODE: convert_smsg_to_string = "WM_GETDLGCODE"
      Case Is = WM_GETFONT: convert_smsg_to_string = "WM_GETFONT"
      Case Is = WM_GETHOTKEY: convert_smsg_to_string = "WM_GETHOTKEY"
      Case Is = WM_GETICON: convert_smsg_to_string = "WM_GETICON"
      Case Is = WM_GETMINMAXINFO: convert_smsg_to_string = "WM_GETMINMAXINFO"
      Case Is = WM_GETOBJECT: convert_smsg_to_string = "WM_GETOBJECT"
      Case Is = WM_GETTEXT: convert_smsg_to_string = "WM_GETTEXT"
      Case Is = WM_GETTEXTLENGTH: convert_smsg_to_string = "WM_GETTEXTLENGTH"
      Case Is = WM_HANDHELDFIRST: convert_smsg_to_string = "WM_HANDHELDFIRST"
      Case Is = WM_HANDHELDLAST: convert_smsg_to_string = "WM_HANDHELDLAST"
      Case Is = WM_HELP: convert_smsg_to_string = "WM_HELP"
      Case Is = WM_HOTKEY: convert_smsg_to_string = "WM_HOTKEY"
      Case Is = WM_HSCROLL: convert_smsg_to_string = "WM_HSCROLL"
      Case Is = WM_HSCROLLCLIPBOARD: convert_smsg_to_string = "WM_HSCROLLCLIPBOARD"
      Case Is = WM_ICONERASEBKGND: convert_smsg_to_string = "WM_ICONERASEBKGND"
      Case Is = WM_IME_CHAR: convert_smsg_to_string = "WM_IME_CHAR"
      Case Is = WM_IME_COMPOSITION: convert_smsg_to_string = "WM_IME_COMPOSITION"
      Case Is = WM_IME_COMPOSITIONFULL: convert_smsg_to_string = "WM_IME_COMPOSITIONFULL"
      Case Is = WM_IME_CONTROL: convert_smsg_to_string = "WM_IME_CONTROL"
      Case Is = WM_IME_ENDCOMPOSITION: convert_smsg_to_string = "WM_IME_ENDCOMPOSITION"
      Case Is = WM_IME_KEYDOWN: convert_smsg_to_string = "WM_IME_KEYDOWN"
      Case Is = WM_IME_KEYLAST: convert_smsg_to_string = "WM_IME_KEYLAST"
      Case Is = WM_IME_KEYUP: convert_smsg_to_string = "WM_IME_KEYUP"
      Case Is = WM_IME_NOTIFY: convert_smsg_to_string = "WM_IME_NOTIFY"
      Case Is = WM_IME_REPORT: convert_smsg_to_string = "WM_IME_REPORT"
      Case Is = WM_IME_REQUEST: convert_smsg_to_string = "WM_IME_REQUEST"
      Case Is = WM_IME_SELECT: convert_smsg_to_string = "WM_IME_SELECT"
      Case Is = WM_IME_SETCONTEXT: convert_smsg_to_string = "WM_IME_SETCONTEXT"
      Case Is = WM_IME_STARTCOMPOSITION: convert_smsg_to_string = "WM_IME_STARTCOMPOSITION"
      Case Is = WM_IMEKEYDOWN: convert_smsg_to_string = "WM_IMEKEYDOWN"
      Case Is = WM_IMEKEYUP: convert_smsg_to_string = "WM_IMEKEYUP"
      Case Is = WM_INITMENU: convert_smsg_to_string = "WM_INITMENU"
      Case Is = WM_INITMENUPOPUP: convert_smsg_to_string = "WM_INITMENUPOPUP"
      Case Is = WM_INPUTLANGCHANGE: convert_smsg_to_string = "WM_INPUTLANGCHANGE"
      Case Is = WM_INPUTLANGCHANGEREQUEST: convert_smsg_to_string = "WM_INPUTLANGCHANGEREQUEST"
      Case Is = WM_INTERIM: convert_smsg_to_string = "WM_INTERIM"
      Case Is = WM_KEYDOWN: convert_smsg_to_string = "WM_KEYDOWN"
      Case Is = WM_KEYLAST: convert_smsg_to_string = "WM_KEYLAST"
      Case Is = WM_KEYUP: convert_smsg_to_string = "WM_KEYUP"
      Case Is = WM_KILLFOCUS: convert_smsg_to_string = "WM_KILLFOCUS"
      Case Is = WM_LBUTTONDBLCLK: convert_smsg_to_string = "WM_LBUTTONDBLCLK"
      Case Is = WM_LBUTTONDOWN: convert_smsg_to_string = "WM_LBUTTONDOWN"
      Case Is = WM_MBUTTONDBLCLK: convert_smsg_to_string = "WM_MBUTTONDBLCLK"
      Case Is = WM_MBUTTONDOWN: convert_smsg_to_string = "WM_MBUTTONDOWN"
      Case Is = WM_MBUTTONUP: convert_smsg_to_string = "WM_MBUTTONUP"
      Case Is = WM_MDIACTIVATE: convert_smsg_to_string = "WM_MDIACTIVATE"
      Case Is = WM_MDICASCADE: convert_smsg_to_string = "WM_MDICASCADE"
      Case Is = WM_MDICREATE: convert_smsg_to_string = "WM_MDICREATE"
      Case Is = WM_MDIDESTROY: convert_smsg_to_string = "WM_MDIDESTROY"
      Case Is = WM_MDIGETACTIVE: convert_smsg_to_string = "WM_MDIGETACTIVE"
      Case Is = WM_MDIICONARRANGE: convert_smsg_to_string = "WM_MDIICONARRANGE"
      Case Is = WM_MDIMAXIMIZE: convert_smsg_to_string = "WM_MDIMAXIMIZE"
      Case Is = WM_MDINEXT: convert_smsg_to_string = "WM_MDINEXT"
      Case Is = WM_MDIREFRESHMENU: convert_smsg_to_string = "WM_MDIREFRESHMENU"
      Case Is = WM_MDIREFRESHMENU: convert_smsg_to_string = "WM_MDIREFRESHMENU"
      Case Is = WM_MDIRESTORE: convert_smsg_to_string = "WM_MDIRESTORE"
      Case Is = WM_MDISETMENU: convert_smsg_to_string = "WM_MDISETMENU"
      Case Is = WM_MDITILE: convert_smsg_to_string = "WM_MDITILE"
      Case Is = WM_MEASUREITEM: convert_smsg_to_string = "WM_MEASUREITEM"
      Case Is = WM_MENUCHAR: convert_smsg_to_string = "WM_MENUCHAR"
      Case Is = WM_MENUCOMMAND: convert_smsg_to_string = "WM_MENUCOMMAND"
      Case Is = WM_MENUDRAG: convert_smsg_to_string = "WM_MENUDRAG"
      Case Is = WM_MENUGETOBJECT: convert_smsg_to_string = "WM_MENUGETOBJECT"
      Case Is = WM_MENURBUTTONUP: convert_smsg_to_string = "WM_MENURBUTTONUP"
      Case Is = WM_MENUSELECT: convert_smsg_to_string = "WM_MENUSELECT"
      Case Is = WM_MOUSEACTIVATE: convert_smsg_to_string = "WM_MOUSEACTIVATE"
      Case Is = WM_MOUSEMOVE: convert_smsg_to_string = "WM_MOUSEMOVE"
      Case Is = WM_MOUSEHOVER: convert_smsg_to_string = "WM_MOUSEHOVER"
      Case Is = WM_MOUSELAST: convert_smsg_to_string = "WM_MOUSELAST"
      Case Is = WM_MOUSELEAVE: convert_smsg_to_string = "WM_MOUSELEAVE"
      Case Is = WM_MOUSEWHEEL: convert_smsg_to_string = "WM_MOUSEWHEEL"
      Case Is = WM_MOVE: convert_smsg_to_string = "WM_MOVE"
      Case Is = WM_MOVING: convert_smsg_to_string = "WM_MOVING"
      Case Is = WM_NCACTIVATE: convert_smsg_to_string = "WM_NCACTIVATE"
      Case Is = WM_NCCALCSIZE: convert_smsg_to_string = "WM_NCCALCSIZE"
      Case Is = WM_NCCREATE: convert_smsg_to_string = "WM_NCCREATE"
      Case Is = WM_NCDESTROY: convert_smsg_to_string = "WM_NCDESTROY"
      Case Is = WM_NCHITTEST: convert_smsg_to_string = "WM_NCHITTEST"
      Case Is = WM_NCLBUTTONDBLCLK: convert_smsg_to_string = "WM_NCLBUTTONDBLCLK"
      Case Is = WM_NCLBUTTONDOWN: convert_smsg_to_string = "WM_NCLBUTTONDOWN"
      Case Is = WM_NCLBUTTONUP: convert_smsg_to_string = "WM_NCLBUTTONUP"
      Case Is = WM_NCMBUTTONDBLCLK: convert_smsg_to_string = "WM_NCMBUTTONDBLCLK"
      Case Is = WM_NCMBUTTONDOWN: convert_smsg_to_string = "WM_NCMBUTTONDOWN"
      Case Is = WM_NCMBUTTONUP: convert_smsg_to_string = "WM_NCMBUTTONUP"
      Case Is = WM_NCMOUSEHOVER: convert_smsg_to_string = "WM_NCMOUSEHOVER"
      Case Is = WM_NCMOUSELEAVE: convert_smsg_to_string = "WM_NCMOUSELEAVE"
      Case Is = WM_NCMOUSEMOVE: convert_smsg_to_string = "WM_NCMOUSEMOVE"
      Case Is = WM_NCPAINT: convert_smsg_to_string = "WM_NCPAINT"
      Case Is = WM_NCRBUTTONDBLCLK: convert_smsg_to_string = "WM_NCRBUTTONDBLCLK"
      Case Is = WM_NCRBUTTONDOWN: convert_smsg_to_string = "WM_NCRBUTTONDOWN"
      Case Is = WM_NCRBUTTONUP: convert_smsg_to_string = "WM_NCRBUTTONUP"
      Case Is = WM_NCXBUTTONDBLCLK: convert_smsg_to_string = "WM_NCXBUTTONDBLCLK"
      Case Is = WM_NCXBUTTONDOWN: convert_smsg_to_string = "WM_NCXBUTTONDOWN"
      Case Is = WM_NCXBUTTONUP: convert_smsg_to_string = "WM_NCXBUTTONUP"
      Case Is = WM_NEXTDLGCTL: convert_smsg_to_string = "WM_NEXTDLGCTL"
      Case Is = WM_NEXTMENU: convert_smsg_to_string = "WM_NEXTMENU"
      Case Is = WM_NOTIFY: convert_smsg_to_string = "WM_NOTIFY"
      Case Is = WM_NOTIFYFORMAT: convert_smsg_to_string = "WM_NOTIFYFORMAT"
      Case Is = WM_NULL: convert_smsg_to_string = "WM_NULL"
      Case Is = WM_OTHERWINDOWCREATED: convert_smsg_to_string = "WM_OTHERWINDOWCREATED"
      Case Is = WM_OTHERWINDOWDESTROYED: convert_smsg_to_string = "WM_OTHERWINDOWDESTROYED"
      Case Is = WM_PAINT: convert_smsg_to_string = "WM_PAINT"
      Case Is = WM_PAINTCLIPBOARD: convert_smsg_to_string = "WM_PAINTCLIPBOARD"
      Case Is = WM_PAINTCLIPBOARD: convert_smsg_to_string = "WM_PAINTCLIPBOARD"
      Case Is = WM_PAINTICON: convert_smsg_to_string = "WM_PAINTICON"
      Case Is = WM_PALETTECHANGED: convert_smsg_to_string = "WM_PALETTECHANGED"
      Case Is = WM_PALETTEISCHANGING: convert_smsg_to_string = "WM_PALETTEISCHANGING"
      Case Is = WM_PARENTNOTIFY: convert_smsg_to_string = "WM_PARENTNOTIFY"
      Case Is = WM_PASTE: convert_smsg_to_string = "WM_PASTE"
      Case Is = WM_PENWINFIRST: convert_smsg_to_string = "WM_PENWINFIRST"
      Case Is = WM_PENWINLAST: convert_smsg_to_string = "WM_PENWINLAST"
      Case Is = WM_POWER: convert_smsg_to_string = "WM_POWER"
      Case Is = WM_PRINT: convert_smsg_to_string = "WM_PRINT"
      Case Is = WM_PRINTCLIENT: convert_smsg_to_string = "WM_PRINTCLIENT"
      Case Is = WM_QUERYDRAGICON: convert_smsg_to_string = "WM_QUERYDRAGICON"
      Case Is = WM_QUERYENDSESSION: convert_smsg_to_string = "WM_QUERYENDSESSION"
      Case Is = WM_QUERYNEWPALETTE: convert_smsg_to_string = "WM_QUERYNEWPALETTE"
      Case Is = WM_QUERYOPEN: convert_smsg_to_string = "WM_QUERYOPENQUERYDRAGICON"
      Case Is = WM_QUERYUISTATE: convert_smsg_to_string = "WM_QUERYUISTATE"
      Case Is = WM_QUEUESYNC: convert_smsg_to_string = "WM_QUEUESYNC"
      Case Is = WM_QUIT: convert_smsg_to_string = "WM_QUIT"
      Case Is = WM_RASDIALEVENT: convert_smsg_to_string = "WM_RASDIALEVENT"
      Case Is = WM_RBUTTONDBLCLK: convert_smsg_to_string = "WM_RBUTTONDBLCLK"
      Case Is = WM_RBUTTONDOWN: convert_smsg_to_string = "WM_RBUTTONDOWN"
      Case Is = WM_RBUTTONUP: convert_smsg_to_string = "WM_RBUTTONUP"
      Case Is = WM_RENDERALLFORMATS: convert_smsg_to_string = "WM_RENDERALLFORMATS"
      Case Is = WM_SETCURSOR: convert_smsg_to_string = "WM_SETCURSOR"
      Case Is = WM_SETFOCUS: convert_smsg_to_string = "WM_SETFOCUS"
      Case Is = WM_SETFONT: convert_smsg_to_string = "WM_SETFONT"
      Case Is = WM_SETHOTKEY: convert_smsg_to_string = "WM_SETHOTKEY"
      Case Is = WM_SETICON: convert_smsg_to_string = "WM_SETICON"
      Case Is = WM_SETREDRAW: convert_smsg_to_string = "WM_SETREDRAW"
      Case Is = WM_WININICHANGE: convert_smsg_to_string = "WM_WININICHANGE"
      Case Is = WM_SETTEXT: convert_smsg_to_string = "WM_SETTEXT"
      Case Is = WM_SHOWWINDOW: convert_smsg_to_string = "WM_SHOWWINDOW"
      Case Is = WM_SIZE: convert_smsg_to_string = "WM_SIZE"
      Case Is = WM_SIZECLIPBOARD: convert_smsg_to_string = "WM_SIZECLIPBOARD"
      Case Is = WM_SIZING: convert_smsg_to_string = "WM_SIZING"
      Case Is = WM_SPOOLERSTATUS: convert_smsg_to_string = "WM_SPOOLERSTATUS"
      Case Is = WM_STYLECHANGED: convert_smsg_to_string = "WM_STYLECHANGED"
      Case Is = WM_STYLECHANGING: convert_smsg_to_string = "WM_STYLECHANGING"
      Case Is = WM_SYNCPAINT: convert_smsg_to_string = "WM_SYNCPAINT"
      Case Is = WM_SYSCHAR: convert_smsg_to_string = "WM_SYSCHAR"
      Case Is = WM_SYSCOLORCHANGE: convert_smsg_to_string = "WM_SYSCOLORCHANGE"
      Case Is = WM_SYSCOMMAND: convert_smsg_to_string = "WM_SYSCOMMAND"
      Case Is = WM_SYSDEADCHAR: convert_smsg_to_string = "WM_SYSDEADCHAR"
      Case Is = WM_SYSKEYDOWN: convert_smsg_to_string = "WM_SYSKEYDOWN"
      Case Is = WM_SYSKEYUP: convert_smsg_to_string = "WM_SYSKEYUP"
      Case Is = WM_TCARD: convert_smsg_to_string = "WM_TCARD"
      Case Is = WM_TIMECHANGE: convert_smsg_to_string = "WM_TIMECHANGE"
      Case Is = WM_TIMER: convert_smsg_to_string = "WM_TIMER"
      Case Is = WM_UNDO: convert_smsg_to_string = "WM_UNDO"
      Case Is = WM_UNINITMENUPOPUP: convert_smsg_to_string = "WM_UNINITMENUPOPUP"
      Case Is = WM_UPDATEUISTATE: convert_smsg_to_string = "WM_UPDATEUISTATE"
      Case Is = WM_USERCHANGED: convert_smsg_to_string = "WM_USERCHANGED"
      Case Is = WM_VKEYTOITEM: convert_smsg_to_string = "WM_VKEYTOITEM"
      Case Is = WM_VSCROLL: convert_smsg_to_string = "WM_VSCROLL"
      Case Is = WM_VSCROLLCLIPBOARD: convert_smsg_to_string = "WM_VSCROLLCLIPBOARD"
      Case Is = WM_WINDOWPOSCHANGED: convert_smsg_to_string = "WM_WINDOWPOSCHANGED"
      Case Is = WM_WINDOWPOSCHANGING: convert_smsg_to_string = "WM_WINDOWPOSCHANGING"
      Case Is = WM_WNT_CONVERTREQUESTEX: convert_smsg_to_string = "WM_WNT_CONVERTREQUESTEX"
      Case Is = WM_XBUTTONDBLCLK: convert_smsg_to_string = "WM_XBUTTONDBLCLK"
      Case Is = WM_XBUTTONDOWN: convert_smsg_to_string = "WM_XBUTTONDOWN"
      Case Is = WM_XBUTTONUP: convert_smsg_to_string = "WM_XBUTTONUP"
  End Select

End Function
 
'
'here we add messages to track to an array
'as messages pass through the window proc in
'module(code) it will go through the array looking for
'a match. If there is one found you will be notified
'
Public Sub add_msg_to_track(msg As enMsg, _
                     Optional b_discard_message As Boolean = False)
 
 Dim upp    As Long
 
 'if we want to track all messages
 If msg = aWM_ALL_MESSAGES Then
    code.tag_all_msgs = True
    Exit Sub
 End If
 
 'get the ubound of the array
 If code.IsArray(code.arr_msg) Then
   upp = UBound(code.arr_msg, 2) + 1
 Else
   upp = 0
 End If
  
 ReDim Preserve code.arr_msg(1, upp)
 code.arr_msg(0, upp) = msg
 code.arr_msg(1, upp) = b_discard_message
 
End Sub
 
'
'unhook
'
Sub clean_up()

  On Error Resume Next
  
  'if were already unhooked then dont bother
  If code.PrevProc = 0 Then Exit Sub
  'if the form was never set dont bother
  If m_your_hwnd = 0 Then Exit Sub
  code.UnHookForm m_your_hwnd
  Set code.calling_class = Nothing
  Erase code.arr_msg
  
End Sub
 
Sub StartSubclass(your_hwnd&, Optional show_msg_as_stringconst As Boolean)
   
   On Error Resume Next
  
   If code.tag_all_msgs = False And Not (code.IsArray(code.arr_msg)) Then
      RaiseEvent Error("You must specify at least one message to track. Subclass has NOT been started")
      Exit Sub
   End If
   
   m_your_hwnd = your_hwnd
   m_show_msg_as_stringconst = show_msg_as_stringconst
   Set code.calling_class = Me
   If code.PrevProc = 0 Then code.HookForm m_your_hwnd
   
End Sub
 
 
'
'if on of the messags specified in [add_msg_to_track]
'match, then [code] sends the notification here so we can raiseevent
Friend Sub friend_event_notify(uMsg&, wParam&, lParam&, bdiscard_msg As Boolean)
  
  If Not (m_show_msg_as_stringconst) Then
    RaiseEvent msgLong(uMsg&, wParam&, lParam&, bdiscard_msg)
  Else
    RaiseEvent msgString(convert_smsg_to_string(uMsg&), wParam&, lParam&, bdiscard_msg)
  End If
  
End Sub
 
Sub about_help()
  
  Fhelp.Show vbModeless

End Sub
'
'unhook if not already done in sub "clean_up"
'
Private Sub Class_Terminate()
  
  On Error Resume Next
  Call clean_up
  
End Sub
