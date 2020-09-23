VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
'
' Class Name:  cTreeView
' Author:      Graeme Grant        (a.k.a. Slider)
' Date:        23/09/2002
' Version:     02.03.01
' Description: Advanced TreeView Handler
' History:     Too big to put here - See cTreeviewHistory.Txt file
' Compiling:   DLL : set class Instancing = 5 - MultiUse
'              APP : set class Instancing = 1 - Private
'
' Notes:       NODLL conditional compilation flag is used in the project's
'              'Conditional Compiling Arguments' property (see
'              tutorial?_NODLL.PRJ properties to see usage).
'
' Original Authors: Credit where it's due:- MultiNode Selection, HitTestEx,
'                   and NodeFindEx2 - www.aboutvb.de; TrueDragImage,
'                   SetOverlayImage, CutIconState, Locked, & Enabled - Brad
'                   Martinez; FirstVisibleNode - www.vb2themax.com;
'                   ScrollView - Microsoft (Articles Q177743 & Q244312);
'                   Subclassing & cTile - Steve McMahon.
'
' Special Thanks:   Zhu JinYong   - Finding alot of bugs + his suggestions
'                   Elim LaRoche  - Suggestions on expanding NodeMove
'                                   functionality
'                   Denis Chollet - Finding problem with NodeMove expanding
'                                   collapsed nodes when it shouldn't.
'                   Michel        - Finding more bugs.
'                   PSC users     - for rewarding my hours of effort with
'                                   votes - without these guys, I probably
'                                   wouldn't have pushed this project as
'                                   far as I have!
'
'===========================================================================

Option Explicit

'===========================================================================
' Debugging... Saves adding the debug statements to the form events
'
#Const DEBUGMODE = 1                    '## 0=No debug
                                        '   1=debug
#Const MOUSEEVENTS = 1                  '## 0=No mouse events
                                        '   1=Mouse Up & Mouse Down
                                        '   2=All Mouse events
#If DEBUGMODE = 1 Then
    Private dbgCtrlName  As String
#End If

'===========================================================================
' Private: API Delclarations & variables
'
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Type Size
  cx As Long
  cy As Long
End Type

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
                                        (ByVal hwnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" _
                                        (ByVal hwnd As Long, _
                                         ByVal wMsg As Long, _
                                         ByVal wParam As Long, _
                                               lParam As Any) As Long

Private Declare Function SendMessageAAny Lib "user32" Alias "SendMessageA" _
                                        (ByVal hwnd As Long, _
                                         ByVal wMsg As Long, _
                                               wParam As Any, _
                                               lParam As Any) As Long

Private Declare Function SetWindowPos Lib "user32" _
                                        (ByVal hwnd As Long, _
                                         ByVal hWndInsertAfter As Long, _
                                         ByVal x As Long, _
                                         ByVal y As Long, _
                                         ByVal cx As Long, _
                                         ByVal cy As Long, _
                                         ByVal wFlags As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                                        (ByVal hwnd As Long, _
                                         ByVal nIndex As Long, _
                                         ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                                        (ByVal hwnd As Long, _
                                         ByVal nIndex As Long) As Long

Private Declare Function ImageList_SetOverlayImage Lib "comctl32.dll" _
                                        (ByVal hIml As Long, _
                                         ByVal iImage As Long, _
                                         ByVal iOverlay As Long) As Boolean

Private Declare Function GetSystemMetrics Lib "user32" _
                                        (ByVal nIndex As Long) As Long

Private Declare Function GetCursorPos Lib "user32" _
                                        (lpPoint As POINTAPI) As Long

Private Declare Function ScreenToClient Lib "user32" _
                                        (ByVal hwnd As Long, _
                                         lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long

Private Declare Function GetClientRect Lib "user32" _
                                        (ByVal hwnd As Long, _
                                         lpRect As RECT) As Long

Private Declare Function PtInRect Lib "user32" _
                                        (lprc As RECT, _
                                         ByVal x As Long, _
                                         ByVal y As Long) As Long

Private Declare Function GetScrollPos Lib "user32" _
                                        (ByVal hwnd As Long, _
                                         ByVal nBar As Long) As Long

Private Declare Function GetScrollInfo Lib "user32" _
                                        (ByVal hwnd As Long, _
                                         ByVal fnBar As SB_Type, _
                                         lpsi As SCROLLINFO) As Boolean

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                                        (lpvDest As Any, _
                                         lpvSource As Any, _
                                         ByVal cbCopy As Long)    '@@ v02.02.00

'{ @@ v02.03.00
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' GDI object functions:
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Const DSna = &H220326
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As Any) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As Any) As Long

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Type NMHDR
    hwndFrom As Long            '## Window handle of control sending message
    idfrom   As Long            '## Identifier of control sending message
    code     As Long            '## Specifies the notification code
End Type

Private Type PAINTSTRUCT
    hdc                  As Long
    fErase               As Long
    rcPaint              As RECT
    fRestore             As Long
    fIncUpdate           As Long
    rgbReserved(0 To 31) As Byte
End Type
'}

Private Const WM_SETREDRAW        As Long = &HB

Private Const TV_FIRST            As Long = &H1100

Private Const TVM_DELETEITEM      As Long = (TV_FIRST + 1)
Private Const TVM_GETITEMRECT     As Long = (TV_FIRST + 4)
Private Const TVM_GETNEXTITEM     As Long = (TV_FIRST + 10)
Private Const TVM_SELECTITEM      As Long = (TV_FIRST + 11)
Private Const TVM_GETITEM         As Long = (TV_FIRST + 12)
Private Const TVM_SETITEM         As Long = (TV_FIRST + 13)
Private Const TVM_GETEDITCONTROL  As Long = (TV_FIRST + 15)
Private Const TVM_HITTEST         As Long = (TV_FIRST + 17)     '@@ v02.02.00

Private Const TVN_FIRST           As Long = -400&               '@@ v02.03.00
Private Const TVN_ITEMEXPANDINGA  As Long = (TVN_FIRST - 5)     '@@
Private Const TVN_ITEMEXPANDEDA   As Long = (TVN_FIRST - 6)     '@@

Private Const TVN_SELCHANGING = (TVN_FIRST - 1)
Private Const TVN_SELCHANGED = (TVN_FIRST - 2)
Private Const TVN_GETDISPINFO = (TVN_FIRST - 3)
Private Const TVN_SETDISPINFO = (TVN_FIRST - 4)
Private Const TVN_DELETEITEM = (TVN_FIRST - 9)
Private Const TVN_BEGINLABELEDIT = (TVN_FIRST - 10)
Private Const TVN_ENDLABELEDIT = (TVN_FIRST - 11)


Private Const EM_SETLIMITTEXT     As Long = &HC5    '## EditBox '@@ v02.02.00
Private Const EM_GETLIMITTEXT     As Long = &HD5    '## EditBox '@@ v02.02.00

' TVM_GETNEXTITEM wParam values
Private Enum TVGN_Flags
    TVGN_ROOT = &H0
    TVGN_NEXT = &H1
    TVGN_PREVIOUS = &H2
    TVGN_PARENT = &H3
    TVGN_CHILD = &H4
    TVGN_FIRSTVISIBLE = &H5
    TVGN_NEXTVISIBLE = &H6
    TVGN_PREVIOUSVISIBLE = &H7
    TVGN_DROPHILITE = &H8
    TVGN_CARET = &H9
'#If (WIN32_IE >= &H400) Then   ' >= Comctl32.dll v4.71
    TVGN_LASTVISIBLE = &HA
'#End If
End Enum

Private Type TVITEM   ' was TV_ITEM
  Mask           As Long
  hItem          As Long
  State          As Long
  StateMask      As Long
  pszText        As Long    ' pointer
  cchTextMax     As Long
  iImage         As Long
  iSelectedImage As Long
  cChildren      As Long
  lParam         As Long
End Type

Private Enum TVITEM_mask
    TVIF_TEXT = &H1
    TVIF_IMAGE = &H2
    TVIF_PARAM = &H4
    TVIF_STATE = &H8
    TVIF_HANDLE = &H10
    TVIF_SELECTEDIMAGE = &H20
    TVIF_CHILDREN = &H40
#If (WIN32_IE >= &H400) Then   ' WIN32_IE = 1024 (>= Comctl32.dll v4.71)
    TVIF_INTEGRAL = &H80
#End If
    TVIF_DI_SETITEM = &H1000   ' Notification
End Enum

Private Enum TVITEM_state
    TVIS_SELECTED = &H2
    TVIS_CUT = &H4
    TVIS_DROPHILITED = &H8
    TVIS_BOLD = &H10
    TVIS_EXPANDED = &H20
    TVIS_EXPANDEDONCE = &H40
#If (WIN32_IE >= &H300) Then
    TVIS_EXPANDPARTIAL = &H80
#End If
    
    TVIS_OVERLAYMASK = &HF00
    TVIS_STATEIMAGEMASK = &HF000
    TVIS_USERMASK = &HF000
End Enum

Private Const TVS_NOTOOLTIPS    As Long = &H80

Private Type TVHITTESTINFO              '@@ v02.02.00
    pt As POINTAPI
    flags As Long
    hItem As Long
End Type

Private Enum eTVHITTESTFLAGS            '@@ v02.02.00
    TVHT_NOWHERE = &H1
    TVHT_ONITEMICON = &H2
    TVHT_ONITEMLABEL = &H4
    TVHT_ONITEMINDENT = &H8
    TVHT_ONITEMBUTTON = &H10
    TVHT_ONITEMRIGHT = &H20
    TVHT_ONITEMSTATEICON = &H40
    TVHT_ABOVE = &H100
    TVHT_BELOW = &H200
    TVHT_TORIGHT = &H400
    TVHT_TOLEFT = &H800
End Enum

' scrollbar definitions
Private Type SCROLLINFO
  cbSize    As Long
  fMask     As SIF_Mask
  nMin      As Long
  nMax      As Long
  nPage     As Long
  nPos      As Long
  nTrackPos As Long
End Type

Private Enum SIF_Mask
  SIF_RANGE = &H1
  SIF_PAGE = &H2
  SIF_POS = &H4
  SIF_DISABLENOSCROLL = &H8
  SIF_TRACKPOS = &H10
  SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
End Enum

Private Enum ScrollDirectionFlags
    sdLeft = &H1
    sdUp = &H2
    sdRight = &H4
    sdDown = &H8
End Enum

Private Enum SB_Type
  SB_HORZ = 0
  SB_VERT = 1
  SB_CTL = 2
  SB_BOTH = 3
End Enum

Private Const WS_HSCROLL = &H100000
Private Const WS_VSCROLL = &H200000

' Scroll Bar Commands
Private Const SB_LINEUP         As Long = 0
Private Const SB_LINELEFT       As Long = 0
Private Const SB_LINEDOWN       As Long = 1
Private Const SB_LINERIGHT      As Long = 1
Private Const SB_PAGEUP         As Long = 2
Private Const SB_PAGELEFT       As Long = 2
Private Const SB_PAGEDOWN       As Long = 3
Private Const SB_PAGERIGHT      As Long = 3
Private Const SB_THUMBPOSITION  As Long = 4
Private Const SB_THUMBTRACK     As Long = 5
Private Const SB_TOP            As Long = 6
Private Const SB_LEFT           As Long = 6
Private Const SB_BOTTOM         As Long = 7
Private Const SB_RIGHT          As Long = 7
Private Const SB_ENDSCROLL      As Long = 8

Private Const WM_HSCROLL        As Long = &H114
Private Const WM_VSCROLL        As Long = &H115

Private Const GW_CHILD          As Long = 5

'## Border Styles
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const GWL_STYLE         As Long = (-16)
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_FRAMECHANGED  As Long = &H20       '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOOWNERZORDER As Long = &H200      '  Don't do owner Z ordering

Private Const WS_EX_CLIENTEDGE  As Long = &H200
Private Const WS_EX_STATICEDGE  As Long = &H20000

Private Const SM_CXDRAG         As Long = &H44
Private Const SM_CYDRAG         As Long = &H45

Private Enum RectFlags
  rfLeft = &H1
  rfTop = &H2
  rfRight = &H4
  rfBottom = &H8
End Enum

'===========================================================================
' Public: Variables and Declarations
'
Public Enum eWinKeys            '## Left out by Microsoft!
    vbKeyWinLeftLogo = &H5B
    vbKeyWinRightLogo = &H5C
    vbKeyWinMenu = &H5D
End Enum

Public Enum eContextMenuMode    '## Right-Click Menu Timing
    [Before Click] = 0
    [After Click] = 1
End Enum

Public Enum eIsScrollbarVisible '## Scroll visible
    [Horizontal] = 1
    [Vertical] = 2
End Enum

Public Enum eCodeScrollView     '## Scroll Treeview
    [Home] = 0
    [Page Up] = 1
    [Up] = 2
    [Down] = 3
    [Page Down] = 4
    [End] = 5
    [Left] = 6
    [Page Left] = 7
    [Line Left] = 8
    [Line Right] = 9
    [Page Right] = 10
    [Right] = 11
End Enum

Public Enum eNodeShuffle        '## Branch level node shuffle
    [Node Up] = 0
    [Node Down] = 1
    [Node First] = 2
    [Node Last] = 3
End Enum

Public Enum eNodeDepth          '## IsChild Node Depth
    [First Branch] = 0
    [All Branches] = 1
End Enum

Public Enum eNodeUniquePrefixNumberMode '@@ v02.02.00
    [Numbers] = 0
    [Prefix] = 1
End Enum

Public Enum eNodeCoords         '## NodeCoord method    v02.02.00
    [By Node] = 0               '## (Default) Width = node label editbox width
    [By TreeView] = 1           '##           Width = Node label left to TreeView window right
End Enum

Public Enum eNodeFindEx          '## NodeFindEx method  v02.02.00
    [Find Node Text] = 1         '## (Default)
    [Find Node Key] = 2          '##
    [Find Node Text And Key] = 3 '##
End Enum

Public Enum eAppearance         '## Treeview Border Appearance
    [None] = 0
    [Fixed Single] = 1
    [Thin] = 2
End Enum

Public Enum eOpMode             '## Normal or Multi-Node Select
    [Single Select] = 0
    [Multi Select] = 1
End Enum

Public Enum eFileMode           '## Load/Save
    [Binary] = 0
    [xml] = 1
End Enum

Public Enum eItemOverlayIndices '## Overlay Icons. [Note: Additional Overlay modes can be used...]
  [No Overlay] = 0
  [Share Overlay] = 1
  [Shortcut Overlay] = 2
  [Custom 1] = 3
  [Custom 2] = 4
  [Custom 3] = 5
End Enum

Public Enum eHitTestEx                                                      '@@ v02.02.00
    [Hit Client Nowhere] = TVHT_NOWHERE
    [Hit Item Icon] = TVHT_ONITEMICON
    [Hit Item Label] = TVHT_ONITEMLABEL
    [Hit Item Indent] = TVHT_ONITEMINDENT
    [Hit Item Button] = TVHT_ONITEMBUTTON
    [Hit Item Right] = TVHT_ONITEMRIGHT
    [Hit Item Check] = TVHT_ONITEMSTATEICON
    [Hit Client Above] = TVHT_ABOVE
    [Hit Client Below] = TVHT_BELOW
    [Hit Client Right] = TVHT_TORIGHT
    [Hit Client Left] = TVHT_TOLEFT
    [Hit Client LeftAbove] = TVHT_TOLEFT Or TVHT_ABOVE
    [Hit Client LeftBelow] = TVHT_TOLEFT Or TVHT_BELOW
    [Hit Client RightAbove] = TVHT_TORIGHT Or TVHT_ABOVE
    [Hit Client RightBelow] = TVHT_TORIGHT Or TVHT_BELOW
End Enum

Public Enum eBackMode                                                       '@@ v02.03.00
    bmDefault = 0
    bmColor = 1
    bmGraphic = 2
End Enum

'===========================================================================
' Private: Variables and Declarations
'
Private Const csDefaultNodePrefix As String = "Copy of "   '@@ v02.02.00 NodeTextUnique function

Private WithEvents oTree  As MSComctlLib.TreeView
Attribute oTree.VB_VarHelpID = -1

Private Type udtNodeInfo                        '## Mimics the Node object + state info.
                                                '   Used by Binary Load/Save function.
    Index        As Long
    LinkId       As Long                        '## Parent Node Index
    Text         As String
    Key          As Variant
    Image        As Variant
    SelImage     As Variant
    ExpImage     As Variant
    Expanded     As Boolean                     '## Expanded State
    Checked      As Boolean                     '## Checked state
    Bold         As Boolean                     '## Bold state
    BackColor    As Long
    ForeColor    As Long
    Sorted       As Long
    Tag          As Variant
    Selected     As Boolean
    FirstVisible As Boolean                     '## Whether node is the first visible node
End Type
    
Private Type udtMemDC                                                       ' @@ v02.03.00
   hdc     As Long
   hBmp    As Long
   hBmpOld As Long
   lWidth  As Long
   lHeight As Long
End Type

Private meAppearance      As eAppearance        '## TreeView Appearance
Private meContextMenuMode As eContextMenuMode   '## Before/After ContextMenu State
Private mbEnabled         As Boolean            '## TreeView Enabled state
Private mbLocked          As Boolean            '## TreeView Locked state

'{ @@ v02.03.00
Private meBackMode        As eBackMode          '## Type of TreeView background
Private mlBackColor       As OLE_COLOR          '## TreeView background color   (meBackMode = bmColor)
Private moBackPicture     As StdPicture         '## TreeView background picture (meBackMode = bmGraphic)
Private msBackFileName    As String             '## TreeView background picture filename
                                                '   NOTE: Set to 'PIC' if a picture object is passed.
Private mtMemDC(1 To 3)   As udtMemDC           '## TreeView background drawing canvas.
Private mlXOffset         As Long               '## TreeView background bitmap tiling offset
Private mcTile            As cTile              '## Class to tile a TreeView background bitmap
'}

''Private mlNodeHeight      As Long               '## Physical node height (pixels)
Private mbToolTips        As Boolean            '## TreeView ToolTip Enabled/Disabled state
Private meOpMode          As eOpMode            '## Normal or Multi-Node Select state

Private moHoverNode       As MSComctlLib.Node   '## Used by Hover Event code to stop cascading events
Private moFocusNode       As MSComctlLib.Node   '## Used in Multi-Node Select mode

''---
'@@ v02.02.00 - Removed & replaced with subclassed code
''Private mbNodeClickEvent  As Boolean            '## Used by new TreeView Locked state to
''                                                '   disable the Node Check feature.
''Private moCheckNode       As MSComctlLib.Node   '## Used by new TreeView Locked state to
''                                                '   disable the Node Check feature.
''---

'===========================================================================
' Private: Variables and Declarations for Drag'n'Drop
'
Private WithEvents moDragScroll As XTimer
Attribute moDragScroll.VB_VarHelpID = -1
Private WithEvents moDragExpand As XTimer
Attribute moDragExpand.VB_VarHelpID = -1

Private mbDragEnabled As Boolean            '## Allow Drag'n'Drop functionality?
Private mbStartDrag   As Boolean            '## Drag operation flag
Private mbInDrag      As Boolean            '## Drag operation flag
Private mlAutoScroll  As Long               '## Distance in which auto-scrolling happens
Private mptBtnDown    As POINTAPI           '## Use to determine if MouseDown+MouseMove is
                                            '   outside start drag region.
Private mszDrag       As Size               '## X and Y distance cursor moves before
                                            '   dragging begins, in pixels.

Private moDragNode    As MSComctlLib.Node   '## Selected Drag Node

'===========================================================================
' Private: Variables and Declarations for DisableNodes  '@@ v02.02.00
'
Private mbDisAllowSel  As Boolean           '## Allow disabled nodes to be selected
Private mlDisBackColor As OLE_COLOR         '## Custom selection Back Colour
Private mlDisForeColor As OLE_COLOR         '## Custom selection Fore Colour
Private mbDisBold      As Boolean           '## Selection bold state
Private meDisAdjSelect As eDisableAdjustSelectMode
Private moDisNodes     As cTags             '## Disabled Node collection
                                            '   object tracks actual node states
Private mDisabledNodes As Collection        '## Selected node collection object

'## moDisNodes variant fields same as moSelNodes variant fields

'===========================================================================
' Private: Variables and Declarations for iMultiSelect
'
Private mlSelBackColor As OLE_COLOR         '## Custom selection Back Colour
Private mlSelForeColor As OLE_COLOR         '## Custom selection Fore Colour
Private mbSelNoClear   As Boolean           '## SpaceBar operation flag
Private mbSelNoDefault As Boolean           '## Focus node select state flag
Private mbSelBold      As Boolean           '## Selection bold state

Private mbSelKbdMode   As Boolean           '## Keyboard selection state flag
Private mbSelMode      As Boolean           '## Selection mode state flag
Private moSelNodes     As cTags             '## Selection Node collection
                                            '   object tracks actual node states
Private mSelectedNodes As Collection        '## Selected node collection object

Private Enum eNodeTags                      '## moSelNodes variant fields
    [Node Key] = 0
    [Node ForeColor] = 1
    [Node BackColor] = 2
    [Node Bold] = 3
End Enum

'===========================================================================
' Implemented Interfaces
'
Implements iMultiSelect
Implements iDisable                                             '@@ v02.02.00

'===========================================================================
' Private: Subclassing                                          '@@ v02.02.00
'
Private Const WM_LBUTTONDOWN   As Long = &H201
Private Const WM_LBUTTONUP     As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203

Private Const WM_DESTROY       As Long = &H2                    '@@ v02.03.00
Private Const WM_PAINT         As Long = &HF                    '@@
Private Const WM_ERASEBKGND    As Long = &H14                   '@@
Private Const WM_NOTIFY        As Long = &H4E                   '@@

Private m_emr                  As EMsgResponse
Private moSubclass             As GSubclass
Private mlHwnd                 As Long
Private mlHWndParent           As Long                          '@@ v02.03.00

'===========================================================================
' Private: cTreeView Internal Error Declarations
'
Private Const csSOURCE_ERR    As String = "cTreeView"
Private Const clTREECOPY_ERR1 As Long = vbObjectError + 2048 + 201
Private Const csTREECOPY_ERR1 As String = "Destination child node does not exist."
Private Const clTREEMODE_ERR1 As Long = vbObjectError + 2048 + 201
Private Const csTREEMODE_ERR1 As String = "MultiNode Selection disabled."
Private Const clTREEMODE_ERR2 As Long = vbObjectError + 2048 + 202
Private Const csTREEMODE_ERR2 As String = "Drag'n'Drop operations cannot be used with MultiNode Selection enabled."
Private Const clTREEHOOK_ERR  As Long = vbObjectError + 2048 + 203                              '@@ v02.01.00
Private Const csTREEHOOK_ERR  As String = "Invalid control! Must be a VB6 TreeView control"     '@@ v02.01.00

'===========================================================================
' Public: cTreeView Event Declarations
'
Public Event Selected(Node As MSComctlLib.Node)
Attribute Selected.VB_Description = "More reliable than NodeClick Event (preferred method)."
Public Event Hover(Node As MSComctlLib.Node)
Attribute Hover.VB_Description = "Generated if Treeview's HotTracking is enabled and returns the Node under the cursor."
Public Event ContextMenu(Node As MSComctlLib.Node, x As Single, y As Single)
Attribute ContextMenu.VB_Description = "Occurs either before or after the KeyPress or mouse Click events."

Public Event StartDrag(SourceNode As MSComctlLib.Node)
Attribute StartDrag.VB_Description = " Occurs when a drag-and-drop operation is initiated."
Public Event Dragging(SourceNode As MSComctlLib.Node, TargetParent As MSComctlLib.Node)
Public Event Dropped(SourceNode As MSComctlLib.Node, TargetParent As MSComctlLib.Node)
Attribute Dropped.VB_Description = "Occurs when a drag-and-drop operation is completed."

Public Event CopyNode(DestNode As MSComctlLib.Node, SrcNode As MSComctlLib.Node, Cancel As Boolean)
Attribute CopyNode.VB_Description = "Generated during the NodeCopy Function."

'===========================================================================
' Public: MSComctlLib.TreeView Event Declarations
'
Public Event AfterLabelEdit(Cancel As Integer, NewString As String)
Attribute AfterLabelEdit.VB_Description = "Occurs after a user edits the label of the currently selected Node or ListItem object."
Public Event BeforeLabelEdit(Cancel As Integer)
Attribute BeforeLabelEdit.VB_Description = "Occurs when a user attempts to edit the label of the currently selected ListItem or Node object."
Public Event Click()
Attribute Click.VB_Description = " Occurs when the user presses and then releases a mouse button over an object."
Public Event Collapse(ByVal Node As MSComctlLib.Node)
Attribute Collapse.VB_Description = "Generated when any Node object in a TreeView control is collapsed."
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Public Event Expand(ByVal Node As MSComctlLib.Node)
Attribute Expand.VB_Description = "Occurs when a Node object in a TreeView control is expanded; that is, when its child nodes become visible."
Public Event KeyDown(KeyCode, Shift)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Public Event KeyPress(KeyAscii)
Attribute KeyPress.VB_Description = " Occurs when the user presses and releases an ANSI key."
Public Event KeyUp(KeyCode, Shift)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Public Event MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
Attribute MouseDown.VB_Description = " Occurs when the user presses the mouse button while an object has the focus with Optional Cancel feature."
Public Event MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Public Event MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Public Event NodeClick(ByVal Node As MSComctlLib.Node)
Attribute NodeClick.VB_Description = "Occurs when a Node object is clicked."
Public Event SelChange()
Attribute SelChange.VB_Description = "MultiSelect Event. Selection change notification."
Public Event StateChange()                                                  '@@ v02.02.00                                                     '@@ v02.02.00

'===========================================================================
' Subclassing                                                               '@@ v02.02.00
'
Friend Property Let MsgResponse(ByVal RHS As EMsgResponse)
    m_emr = RHS
End Property

Friend Property Get MsgResponse() As EMsgResponse
    MsgResponse = m_emr
End Property

Friend Function WindowProc(ByVal hwnd As Long, _
                           ByVal iMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

    Dim oNode As MSComctlLib.Node
    Dim lR    As Long
    Dim bR    As Boolean
    Dim bOK   As Boolean

    bOK = ((meBackMode = bmGraphic) And (mcTile.Filename <> "")) Or (meBackMode = bmColor)

    Select Case iMsg
        '
        '{ Disabled Node code
        '
        Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_LBUTTONDBLCLK
            With oTree
                Set oNode = .HitTest(.Parent.ScaleX(pLoWord(lParam), vbPixels, .Parent.ScaleMode), _
                                     .Parent.ScaleY(pHiWord(lParam), vbPixels, .Parent.ScaleMode))
                If Not (oNode Is Nothing) Then
                    If moDisNodes.Exist(ObjPtr(oNode)) And (mbDisAllowSel = False) Then
                        '
                        '## Force the Disabled Node to be painted if selected
                        '
                        LockWindowUpdate mlHWndParent                           '@@ v02.03.01
                        SendMessageAny mlHwnd, WM_PAINT, &H0, ByVal &H0         '@@
                        SendMessageAny mlHwnd, WM_ERASEBKGND, &H0, ByVal &H0    '@@
                        LockWindowUpdate 0                                      '@@
                        '
                        '## Don't let the node be selected as it's disabled
                        '
                        m_emr = emrConsume
                    Else
                        m_emr = emrPostProcess
                    End If
                    '
                    '## Disable the modification of a Node's check status if control
                    '   is locked.
                    '
                    If Not (iMsg = WM_LBUTTONDBLCLK) Then
                        '
                        '## Only handle the Single-Click Mouse event
                        '
                        If (mbLocked = True) And (.Checkboxes = True) Then
                            '
                            '## Don't allow the check state to be modified
                            '
                            m_emr = emrConsume
                        End If
                    End If
                Else
                    m_emr = emrPostProcess
                End If
            End With
        '}
        '
        '{ Background color/Image                                       '@@ v02.03.00
        '
        Case WM_PAINT, WM_ERASEBKGND
            Select Case iMsg
                Case WM_PAINT
                    #If DEBUGMODE = 1 Then
                        Debug.Print "WM_PAINT", wParam, lParam
                    #End If
                Case Else
                    #If DEBUGMODE = 1 Then
                        Debug.Print "WM_ERASEBKGND", wParam, lParam
                    #End If
            End Select
            If bOK Then
                pTreeViewMessage hwnd, iMsg, wParam, lParam, lR, bR
                If bR Then
                   WindowProc = lR
                Else
                   WindowProc = moSubclass.CallOldWindowProc(hwnd, iMsg, wParam, lParam)
                End If
            Else
                WindowProc = moSubclass.CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            End If
            m_emr = emrPostProcess

        Case WM_NOTIFY
            If bOK Then
                Dim tNHM As NMHDR
                CopyMemory tNHM, ByVal lParam, Len(tNHM)
                If tNHM.hwndFrom = mlHwnd Then
                    Select Case tNHM.code
                        Case TVN_ITEMEXPANDINGA
                            #If DEBUGMODE = 1 Then
                                Debug.Print "WM_NOTIFY::TVN_ITEMEXPANDINGA"
                            #End If
                            LockWindowUpdate mlHWndParent
                        Case TVN_ITEMEXPANDEDA
                            #If DEBUGMODE = 1 Then
                                Debug.Print "WM_NOTIFY::TVN_ITEMEXPANDEDA"
                            #End If
                            InvalidateRect mlHwnd, ByVal 0&, 0
                            LockWindowUpdate 0
                        Case TVN_SELCHANGING
                            #If DEBUGMODE = 1 Then
                                Debug.Print "WM_NOTIFY::TVN_SELCHANGING"
                            #End If
                        Case TVN_SELCHANGED
                            #If DEBUGMODE = 1 Then
                                Debug.Print "WM_NOTIFY::TVN_SELCHANGED"
                            #End If
'                        Case TVN_GETDISPINFO
'                            'Debug.Print "TVN_GETDISPINFO"
'                        Case TVN_SETDISPINFO
'                            Debug.Print "TVN_SETDISPINFO"
'                        Case TVN_DELETEITEM
'                            'Debug.Print "TVN_DELETEITEM"
                        Case TVN_BEGINLABELEDIT
                            #If DEBUGMODE = 1 Then
                                Debug.Print "WM_NOTIFY::TVN_BEGINLABELEDIT"
                            #End If
                        Case TVN_ENDLABELEDIT
                            #If DEBUGMODE = 1 Then
                                Debug.Print "WM_NOTIFY::TVN_ENDLABELEDIT"
                            #End If
'                        Case Else
'                            'Debug.Print "MSG=" + CStr(tNHM.code)
                    End Select
                End If
            End If
        '}

    End Select

End Function

'===========================================================================
' TreeView: Events
'
Private Sub oTree_AfterLabelEdit(Cancel As Integer, NewString As String)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::AfterLabelEdit -> Cancel="; CStr(Cancel); "  NewString="; NewString
    #End If
    RaiseEvent AfterLabelEdit(Cancel, NewString)
End Sub

Private Sub oTree_BeforeLabelEdit(Cancel As Integer)

    Dim lMaxEditLen    As Long

    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::BeforeLabelEdit -> Cancel="; CStr(Cancel)
    #End If
    '{ MultiSelect
    If mbSelMode Then
        Cancel = True
    Else
    '}
        RaiseEvent BeforeLabelEdit(Cancel)
    End If

End Sub

Private Sub oTree_Click()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Click"
    #End If

'@@ v02.02.00 This code was buggy - now uses subclassing to implement this feature.
''    '## If Locked state = true, then try and stop user from changing
''    '   a Node's check state.
''    If mbLocked Then
''        If oTree.Checkboxes Then
''            If Not mbNodeClickEvent Then
''                If Not (moCheckNode Is Nothing) Then
''                    moCheckNode.Checked = (Not moCheckNode.Checked)
''                    Set moCheckNode = Nothing
''                End If
''            End If
''        End If
''    End If
''    mbNodeClickEvent = False
    RaiseEvent Click

End Sub

Private Sub oTree_Collapse(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Collapse -> Node="; Node.Text
    #End If
    RaiseEvent Collapse(Node)

End Sub

Private Sub oTree_DblClick()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::DblClick"
    #End If
    RaiseEvent DblClick

End Sub

Private Sub oTree_DragDrop(Source As Control, x As Single, y As Single)

    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::DragDrop -> Source="; Source.Name; "  X="; CStr(x); "  Y="; CStr(y)
    #End If

    If mbDragEnabled Then
        With oTree
            .DropHighlight = .HitTest(x, y)
            If Not (.DropHighlight Is Nothing) Then                 '## Did we drop a node?
                If moDragNode <> .DropHighlight Then                '## Yes. Did we drag the node onto itself?
                    RaiseEvent Dropped(moDragNode, .DropHighlight)  '## Notify programmer & Reset
                End If
            End If
            '## Reset
            Set .DropHighlight = Nothing
            Set moDragNode = Nothing
            mbInDrag = False
            mbStartDrag = False
            '.Drag vbEndDrag                                         '!! Moved to otree_MouseUp
        End With
    End If

End Sub

Private Sub oTree_DragOver(Source As Control, x As Single, y As Single, State As Integer)

    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::DragOver -> Source="; Source.Name; "  X="; CStr(x); "  Y="; CStr(y)
    #End If

    If mbDragEnabled Then
        With oTree
            Set .DropHighlight = .HitTest(x, y)
            If .DropHighlight Is Nothing Then
                #If NODLL = 0 Then
                    .DragIcon = LoadResPicture(1, vbResCursor)
                #Else
                    .DragIcon = LoadPicture(App.Path + "\NoSelect.CUR")
                #End If
            Else
                .DragIcon = moDragNode.CreateDragImage
            End If
        End With
        pDoDrag
    End If

End Sub

Private Sub oTree_Expand(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Expand -> Node="; Node.Text
    #End If
    RaiseEvent Expand(Node)

End Sub

Private Sub oTree_GotFocus()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::GotFocus"
    #End If

End Sub

Private Sub oTree_KeyDown(KeyCode As Integer, Shift As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::KeyDown -> KeyCode="; CStr(KeyCode); "  Shift="; CStr(Shift)
    #End If

    Dim rc       As RECT
    Dim snOutX   As Single
    Dim snOutY   As Single
    Dim oSelNode As MSComctlLib.Node
    Dim oResNode As MSComctlLib.Node
    Dim bEval    As Boolean

    RaiseEvent KeyDown(KeyCode, Shift)

    '{ MultiNode Select code
    If meOpMode = [Multi Select] Then
        mbSelKbdMode = True
    End If
    '}

    '## Fire ContextMenu event before click event?
    If meContextMenuMode = [Before Click] Then
        Select Case KeyCode
            Case vbKeyWinMenu
                '
                '## Retrieve selected node's co-ordinates and covert to the treeview's parent
                '   container's scale mode.
                '
                rc = pGetSelectedNodeCoords
                pScale CSng(rc.Left + (rc.Right - rc.Left) \ 4), _
                       CSng(rc.Top + (rc.Bottom - rc.Top) \ 2), _
                       snOutX, _
                       snOutY, _
                       vbPixels, _
                       oTree.Parent.ScaleMode
                RaiseEvent ContextMenu(oTree.SelectedItem, snOutX, snOutY)
        End Select
    End If

    '{ Disable Node code    '@@ v02.02.00
    '
    ' ## Note: Managing the cursor control in the KeyDown event, instead of
    '          KeyUp, stops flickering of the selected item.
    '
    If mbDisAllowSel = False Then
        Set oSelNode = oTree.SelectedItem
        Select Case KeyCode

            Case vbKeyLeft
                If (oSelNode.Children = 0) Or _
                  ((oSelNode.Children > 0) And (oSelNode.Expanded = False)) Then
                    Set oResNode = pGetPrevEnabledParent(oSelNode)
                    bEval = True
                End If

            Case vbKeyDown, vbKeyRight
                Set oResNode = pGetNextEnabled(oSelNode)
                If (oResNode Is Nothing) And (oTree.Nodes.Count > 0) Then
                    Set oResNode = oSelNode
                End If
                bEval = True

            Case vbKeyPageUp
                Set oResNode = FirstVisibleNode
                If oResNode.Key = oSelNode.Key Then
                    Set oResNode = pGetPrevEnabled(oResNode)
                End If
                If Not (oResNode Is Nothing) Then
                    Debug.Print oResNode.Text
                    If moDisNodes.Exist(ObjPtr(oResNode)) Then
                        Set oResNode = pGetPrevEnabled(oResNode)
                        bEval = True
                    End If
                Else
                    bEval = True
                End If

            Case vbKeyPageDown
                Set oResNode = LastVisibleNode
                If oResNode.Key = oSelNode.Key Then
                    Set oResNode = pGetNextEnabled(oResNode)
                End If
                If Not (oResNode Is Nothing) Then
                    Debug.Print oResNode.Text
                    If moDisNodes.Exist(ObjPtr(oResNode)) Then
                        Set oResNode = pGetNextEnabled(oResNode)
                        bEval = True
                    End If
                Else
                    bEval = True
                End If

            Case vbKeyUp
                Set oResNode = pGetPrevEnabled(oSelNode)
                If (oResNode Is Nothing) And (oTree.Nodes.Count > 0) Then
                    Set oResNode = oSelNode
                End If
                bEval = True

            Case vbKeyHome
                Set oResNode = pGetFirstNode
                If moDisNodes.Exist(ObjPtr(oResNode)) Then
                    Set oResNode = pGetNextEnabled(oResNode)
                    bEval = True
                End If

            Case vbKeyEnd
                Set oResNode = pGetLastNode
                If moDisNodes.Exist(ObjPtr(oResNode)) Then
                    Set oResNode = pGetPrevEnabled(oResNode)
                    bEval = True
                End If

        End Select
    End If

    '## Cursor control needs to override built-in functionality
    If bEval Then
        If Not (oResNode Is Nothing) Then
            If Not (oSelNode Is Nothing) Then
                If meOpMode = [Multi Select] Then
                    pSetFocusNode oResNode
                Else
                    Set oTree.SelectedItem = oResNode
                End If
                oResNode.EnsureVisible
                KeyCode = 0
                RaiseEvent SelChange
            End If
        Else
            Set oTree.SelectedItem = Nothing
            KeyCode = 0
        End If
    End If
    '}

End Sub

Private Sub oTree_KeyPress(KeyAscii As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::KeyPress -> KeyAscii="; CStr(KeyAscii)
    #End If

    '{ MultiNode Select code
    If meOpMode = [Multi Select] Then
        Select Case KeyAscii
            Case vbKeyEscape
                KeyAscii = 0
            Case vbKeySpace
                KeyAscii = 0
        End Select
    End If
    '}

    If KeyAscii Then
        RaiseEvent KeyPress(KeyAscii)
    End If

End Sub

Private Sub oTree_KeyUp(KeyCode As Integer, Shift As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::KeyUp -> KeyCode="; CStr(KeyCode); "  Shift="; CStr(Shift)
    #End If

    Dim rc     As RECT
    Dim snOutX As Single
    Dim snOutY As Single

    '{ MultiNode Select code
    If meOpMode = [Multi Select] Then
        Select Case KeyCode
            Case vbKeyEscape
                If Shift = vbShiftMask Then
                    iMultiSelect_ClearSelection
                    KeyCode = 0
                End If
            Case vbKeySpace
                Select Case Shift
                    Case 0
                        iMultiSelect_ToggleNode , CBool(Shift = vbShiftMask)
                        KeyCode = False
                    Case vbShiftMask
                        iMultiSelect_ToggleNode moFocusNode, CBool(Shift = vbShiftMask)
                        KeyCode = False
                End Select
            Case vbKeyA
                If Shift = vbCtrlMask Then
                    iMultiSelect_SelectAllNodes
                    KeyCode = 0
                End If
            Case vbKeyT
                If Shift = vbCtrlMask Then
                    iMultiSelect_ToggleSelection
                    KeyCode = 0
                End If
        End Select
    End If
    '}

    If KeyCode Then
        RaiseEvent KeyUp(KeyCode, Shift)
    End If

    '## Fire ContextMenu event after keypress event?
    If meContextMenuMode = [After Click] Then
        Select Case KeyCode
            Case vbKeyWinMenu
                '
                '## Retrieve selected node's co-ordinates and covert to the treeview's parent
                '   container's scale mode.
                '
                rc = pGetSelectedNodeCoords
                pScale CSng(rc.Left + (rc.Right - rc.Left) \ 4), CSng(rc.Top + (rc.Bottom - rc.Top) \ 2), _
                       snOutX, snOutY, _
                       vbPixels, _
                       oTree.Parent.ScaleMode
                RaiseEvent ContextMenu(oTree.SelectedItem, snOutX, snOutY)
        End Select
    End If

End Sub

Private Sub oTree_LostFocus()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::LostFocus"
    #End If
End Sub

Private Sub oTree_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 1 Or MOUSEEVENTS = 2 Then
            Debug.Print dbgCtrlName; "::MouseDown -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If

    Dim oNode   As MSComctlLib.Node
    Dim bCancel As Boolean
    
    With oTree
        '## Is drag'n'drop allowed?
        If mbDragEnabled Then
            GetCursorPos mptBtnDown                         '## Update cursor position for Drag'n'Drop testing
            If Button = vbLeftButton Then
                Set moDragNode = .HitTest(x, y)             '## Capture the node to be dragged
            End If
        End If

        '## Fire ContextMenu event before click event?
        If meContextMenuMode = [Before Click] Then
            If Button = vbRightButton Then
                Set .SelectedItem = .HitTest(x, y)
                RaiseEvent ContextMenu(.SelectedItem, x, y)
            End If
        End If
    End With

    RaiseEvent MouseDown(Button, Shift, x, y, bCancel)

    '{ MultiNode Select code
    mbSelMode = False
    mbSelKbdMode = False
    If meOpMode = [Multi Select] Then
        If Not bCancel Then
            Set oNode = oTree.HitTest(x, y)
            If (Shift And vbCtrlMask) = vbCtrlMask Then
                mbSelMode = True
            ElseIf (Not mbSelNoClear) Or (Not mbSelKbdMode) Then
                iMultiSelect_ClearSelection
                RaiseEvent SelChange
            End If
        End If
    End If
    '}

End Sub

Private Sub oTree_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 2 Then
            Debug.Print dbgCtrlName; "::MouseMove -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If

    Dim oNode As MSComctlLib.Node
    Dim pt As POINTAPI

    Set oNode = oTree.HitTest(x, y)
    Select Case Button
        Case vbLeftButton
        Case vbRightButton
        Case Else
            '
            '## Hover event
            '
            If oTree.HotTracking Then
                '## Requires HotTracking
                If Not (oNode Is Nothing) Then
                    '## Is the cursor is over a node
                    Select Case True
                        Case (moHoverNode Is Nothing), (oNode.Key <> moHoverNode.Key)
                            '## New node selected
                            Set moHoverNode = oNode
                            RaiseEvent Hover(oNode)
                    End Select
                End If
            End If

    End Select

    On Error GoTo ErrorHandler

    If mbDragEnabled Then                                   '## Is drag'n'drop allowed?
        If Button = vbLeftButton Then                       '## Yes. Signal a Drag operation.
            With oTree
                If Not (.HitTest(x, y) Is Nothing) Then     '## Do we have a node selected?
                    If mbStartDrag = True Then
                        mbInDrag = True                         '## Yes. Set the flag to true.
                        '.DragIcon = moDragNode.CreateDragImage '!! Moved to otree_DragOver
                        .Drag vbBeginDrag                       '## Signal VB to start drag operation.
                    Else
                        If Not (moDragNode Is Nothing) Then
                            GetCursorPos pt
                            If (Abs(pt.x - mptBtnDown.x) >= mszDrag.cx) Or (Abs(pt.y - mptBtnDown.y) >= mszDrag.cy) Then
                                RaiseEvent StartDrag(moDragNode)    '## Notify programmer starting drag operation
                                mbStartDrag = True
                            End If
                        End If
                    End If
                End If
            End With
        End If
    End If

    RaiseEvent MouseMove(Button, Shift, x, y)
    Exit Sub

ErrorHandler:
    mbStartDrag = False
    mbInDrag = False
End Sub

Private Sub oTree_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 1 Or MOUSEEVENTS = 2 Then
            Debug.Print dbgCtrlName; "::MouseUp -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If
    Dim oNode As MSComctlLib.Node

    Set oNode = oTree.HitTest(x, y)
    Select Case Button
        Case vbLeftButton
            '{ MultiNode Select code
            If meOpMode = [Multi Select] Then
                If Not (oNode Is Nothing) Then
                    If (Shift And vbCtrlMask) = vbCtrlMask Then
                        iMultiSelect_ToggleNode oNode, CBool((Shift And vbShiftMask) = vbShiftMask)
                        oNode.Selected = False
                    End If
                End If
            End If
            '}

        Case vbRightButton

        Case Else
            '{ MultiNode Select code
            If mbSelMode Then
                If Not (oNode Is Nothing) Then
                    oNode.Selected = False
                End If
            End If
            '}

    End Select
    '{ MultiNode Select & Disabled code
    Set moFocusNode = oNode
    '}
    '## Fire ContextMenu event after click event?
    If meContextMenuMode = [After Click] Then
        If Button = vbRightButton Then
            With oTree
                Set .SelectedItem = .HitTest(x, y)
                RaiseEvent ContextMenu(.SelectedItem, x, y)
            End With
        End If
    End If

    If mbDragEnabled Then                                   '## Is drag'n'drop allowed?
        mbStartDrag = False
        mbInDrag = False

        oTree.Drag vbEndDrag
        moDragExpand.Enabled = False
        moDragScroll.Enabled = False
    End If

    RaiseEvent MouseUp(Button, Shift, x, y)

End Sub

Private Sub oTree_NodeCheck(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::NodeCheck -> Node="; Node.Text
    #End If

    'Set moCheckNode = Node                     '@@ v02.02.00 - Removed & replaced with subclassed code

End Sub

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::NodeClick -> Node="; Node.Text
    #End If
    RaiseEvent NodeClick(Node)

''    '## Node selected, raise event            '@@ v02.02.00 - Removed & replaced with subclassed code
''    mbNodeClickEvent = True

    '{ MultiNode Select code
    If meOpMode = [Multi Select] Then
        If Not mbSelMode And Not mbSelKbdMode Then
            If mbSelNoDefault Then
                Node.Selected = False
            End If
            With Node
                '## Reset collections and select node
                Set moSelNodes = New cTags
                Set mSelectedNodes = New Collection
                pSetNodeSelect Node, True
            End With
            RaiseEvent SelChange
        End If
        If mbSelKbdMode And mbSelNoDefault Then
            Node.Selected = False
        End If
        Set moFocusNode = Node
    End If
    '}

    RaiseEvent NodeClick(Node)
    RaiseEvent Selected(Node)

End Sub

Private Sub oTree_OLECompleteDrag(Effect As Long)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLECompleteDrag -> Effect="; CStr(Effect)
    #End If

End Sub

Private Sub oTree_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEDragDrop -> Effect="; CStr(Effect); "  Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
    #End If

End Sub

Private Sub oTree_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEDragOver -> Effect="; CStr(Effect); "  Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y); "  State="; CStr(State)
    #End If

End Sub

Private Sub oTree_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEGiveFeedback -> Effect="; CStr(Effect); "  DefaultCursors="; CStr(DefaultCursors)
    #End If

End Sub

Private Sub oTree_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLESetData -> Effect="; CStr(DataFormat)
    #End If

End Sub

Private Sub oTree_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEStartDrag -> AllowedEffects="; CStr(AllowedEffects)
    #End If

End Sub

Private Sub oTree_Validate(Cancel As Boolean)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Validate -> Cancel="; CStr(Cancel)
    #End If

End Sub

'===========================================================================
' Internal Class Subroutines
'
Private Sub Class_Initialize()
    '## Set defaults...
    meContextMenuMode = [After Click]
    meOpMode = [Single Select]          '## Normal TreeView operation mode
    mbToolTips = True

    '{ Disable Node code                                        '@@ v02.02.00
    meDisAdjSelect = [eiDis Focus Rectangle]
    mbDisAllowSel = False
    mlDisBackColor = vbWindowBackground
    mlDisForeColor = vbGrayText
    Set moDisNodes = New cTags
    Set mDisabledNodes = New Collection
    '}

    '{ MultiNode Select code
    mbSelNoDefault = True
    mlSelBackColor = vbHighlight
    mlSelForeColor = vbWindowBackground
    Set moSelNodes = New cTags
    Set mSelectedNodes = New Collection
    '}

    '{ TreeView Background color/gradient/picture               '@@ v02.03.00
    meBackMode = bmDefault
    mlBackColor = vbInfoBackground
    msBackFileName = ""
    Set mcTile = New cTile
    '}

End Sub

Private Sub Class_Terminate()
    '## release control
    UnHookCtrl
    '{ MultiNode Select code
    Set moSelNodes = Nothing
    Set mSelectedNodes = Nothing
    '}
    Set mcTile = Nothing                                        '@@ v02.03.00
    
    If Not (moDragScroll Is Nothing) Then Set moDragScroll = Nothing
    If Not (moDragExpand Is Nothing) Then Set moDragExpand = Nothing
End Sub

'===========================================================================
' Timers: Events
'
Private Sub moDragExpand_Tick()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::moDragExpand -> Tick!"
    #End If
    With oTree
        Select Case True
            Case (.DropHighlight Is Nothing), (moDragNode Is Nothing)
                '## Avoid possible error - should not get here!
            Case (.DropHighlight.Children > 0) And (.DropHighlight.Expanded = False)
                '## The highlighted node has children and is collapsed
                .DropHighlight.Expanded = True
        End Select
    End With
    '## Kill Timer
    moDragExpand.Enabled = False
End Sub

Private Sub moDragScroll_Tick()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::moDragScroll -> Tick!"
    #End If

    Dim pt            As POINTAPI
    Dim rcClient      As RECT
    Dim dwRectFlags   As RectFlags
    Dim dwScrollFlags As ScrollDirectionFlags

    If mbInDrag = False Then
        '## We shouldn't be here!
        moDragScroll.Enabled = False
        Exit Sub
    End If
    '
    '## Get the cursor postion in TreeView client coords
    '
    With oTree
        GetCursorPos pt
        ScreenToClient .hwnd, pt
        GetClientRect .hwnd, rcClient
    End With
    '
    '## If the cursor is within an auto scroll region in the TreeView's client area...
    '
    dwRectFlags = pPtInRectRegion(rcClient, mlAutoScroll, pt)
    If dwRectFlags Then
        '
        '## Determine which direction the TreeView can be scrolled...
        '
        dwScrollFlags = pIsWindowScrollable(oTree.hwnd)
        '
        '## If the cursor is within the respective drag region specified by the
        '   mlAutoScroll distance, and if the TreeView can be scrolled
        '   in that direction, send the TreeView that respective scroll message.
        '
        Select Case True
            Case (dwRectFlags And rfLeft) And (dwScrollFlags And sdLeft)
                Debug.Print "Left"
                ScrollView [Line Left]
            Case (dwRectFlags And rfRight) And (dwScrollFlags And sdRight)
                Debug.Print "Right"
                ScrollView [Line Right]
            Case (dwRectFlags And rfTop) And (dwScrollFlags And sdUp)
                Debug.Print "Up"
                ScrollView [Up]
            Case (dwRectFlags And rfBottom) And (dwScrollFlags And sdDown)
                Debug.Print "Down"
                ScrollView [Down]
            Case Else
                moDragScroll.Enabled = False
        End Select
    End If

End Sub

'===========================================================================
' Public Properties
'
Public Property Get Appearance() As eAppearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D, Flat 3D  or no effects."
    Appearance = meAppearance
End Property

Public Property Let Appearance(ByVal New_Appearance As eAppearance)

    Dim lStyle As Long

    meAppearance = New_Appearance

    '## Retrieve current style
    lStyle = GetWindowLong(oTree.hwnd, GWL_EXSTYLE)

    '## Toggle style flags
    Select Case New_Appearance
        Case [None]:         lStyle = lStyle And Not WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
        Case [Fixed Single]: lStyle = lStyle Or WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
        Case [Thin]:         lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    End Select

    '## Apply changes
    SetWindowLong oTree.hwnd, GWL_EXSTYLE, lStyle
    SetWindowPos oTree.hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or _
                                            SWP_NOSIZE Or SWP_NOMOVE

End Property

Public Property Get BackColor() As OLE_COLOR                                    '@@ v02.03.00
    BackColor = mlBackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)                     '@@ v02.03.00
    mlBackColor = New_Color
    BackMode = BackMode
End Property

Public Property Get BackFileName() As String                                    '@@ v02.03.00
    BackFileName = msBackFileName
End Property

Public Property Let BackFileName(ByVal New_FileName As String)                  '@@ v02.03.00
    msBackFileName = New_FileName
    mcTile.Picture = LoadPicture()      '## Reset image
    BackMode = BackMode
End Property

Public Property Get BackMode() As eBackMode                                     '@@ v02.03.00
    BackMode = meBackMode
End Property

Public Property Let BackMode(ByVal New_Mode As eBackMode)                       '@@ v02.03.00

    If Not (meBackMode = New_Mode) Then
        pClearUpDC 2
        meBackMode = New_Mode
    End If
    
    If meBackMode = bmGraphic Then
        If Not (moBackPicture Is Nothing) Then
            mcTile.Picture = moBackPicture
        ElseIf Len(msBackFileName) Then
            mcTile.Filename = msBackFileName
        Else
            Exit Property
        End If
    End If
    '
    '## Force the control to repaint. Note: there may be another method to trick the
    '   TreeView control into thinking that something has changed but this'll do for
    '   now due to the time spent pulling my hair out!! I haven't much let - need
    '   to let it grow... :)
    '
    LockWindowUpdate mlHWndParent
    SendMessageAny mlHwnd, WM_PAINT, &H0, ByVal &H0
    SendMessageAny mlHwnd, WM_ERASEBKGND, &H0, ByVal &H0
    LockWindowUpdate 0

End Property

Public Property Get BackPicture() As StdPicture                                 '@@ v02.03.00
    Set BackPicture = mcTile.Picture
End Property

Public Property Set BackPicture(ByVal New_Pic As StdPicture)                    '@@ v02.03.00
    Set moBackPicture = New_Pic
    mcTile.Filename = ""
    BackMode = BackMode
End Property

Public Property Get ChildParentKeys(Optional ByVal oNode As MSComctlLib.Node) As String()
Attribute ChildParentKeys.VB_Description = "Returns a string array of node keys."

    Dim skeys() As String
    Dim sTmp()  As String
    Dim iMax    As Integer
    Dim iloop   As Integer

    '## Check and set Node
    If (oNode Is Nothing) Then
        If (oTree.SelectedItem Is Nothing) Then
            Exit Property
        Else
            Set oNode = oTree.SelectedItem
        End If
    Else
        Set oNode = oNode
    End If

    '## Traverse backwards and store keys
    Do
        iMax = iMax + 1
        ReDim Preserve sTmp(1 To iMax)
        sTmp(iMax) = oNode.Key
        If Not (oNode.FullPath = oNode.Text) Then
            Set oNode = oNode.Parent
        Else
            '## Root node reached.
            '   Swap keys to correct sequence and return results
            ReDim Preserve skeys(1 To iMax)
            For iloop = 1 To iMax
                skeys(iloop) = sTmp(iMax - (iloop - 1))
            Next
            ChildParentKeys = skeys()
            Exit Do
        End If
    Loop

End Property

Public Property Get ContextMenuMode() As eContextMenuMode
Attribute ContextMenuMode.VB_Description = "Set whether the ContextMenu Event fires before or after KeyPress or Click Events [Default = After]."
    ContextMenuMode = meContextMenuMode
End Property

Public Property Let ContextMenuMode(ByVal Mode As eContextMenuMode)
    meContextMenuMode = Mode
End Property

#If NODLL = 0 Then
    Public Property Get Ctrl() As Object
Attribute Ctrl.VB_Description = "Exposes the TreeView's properties and methods that are not wrapped by the cTreeView class"
        '## Default Object
        Set Ctrl = oTree
#Else
    Public Property Get Ctrl() As MSComctlLib.TreeView
        '## Default Object
        Set Ctrl = oTree
#End If

End Property

Public Property Get CutIcon(Optional ByVal Node As MSComctlLib.Node) As Boolean
Attribute CutIcon.VB_Description = "Gets/Sets a Node's cut icon state."

    Dim tvi   As TVITEM
    Dim hItem As Long

    '## Get Node handle
    hItem = pGetNodeHandle(Node)
    '## Process if valid handle
    If hItem Then
        '## Retrieve Cut Icon Flag
        tvi.Mask = TVIF_STATE
        tvi.hItem = hItem
        tvi.StateMask = TVIS_CUT
        If SendMessageAny(oTree.hwnd, TVM_GETITEM, 0&, tvi) Then
            CutIcon = (tvi.State And TVIS_CUT) / (2 ^ 8)
        End If
    End If

End Property

Public Property Let CutIcon(Optional ByVal Node As MSComctlLib.Node, ByVal SetState As Boolean)

    Dim tvi   As TVITEM
    Dim hItem As Long
  
    '## Get Node handle
    hItem = pGetNodeHandle(Node)
    '## Process if valid handle
    If hItem Then
        With tvi
            .hItem = hItem
            .Mask = TVIF_STATE
            .State = SetState And TVIS_CUT
            '## Indicate what state bits we're changing
            .StateMask = TVIS_CUT
        End With
        '## Old docs say returns 0 on success, -1 on failure. New docs say
        '   returns TRUE if successful, or FALSE otherwise we'll go new...
        SendMessageAny oTree.hwnd, TVM_SETITEM, 0&, tvi
    End If

End Property

Public Property Get DragEnabled() As Boolean
Attribute DragEnabled.VB_Description = "Enables/disables Drag'n'Drop operation. [Note: Cannot be used with MultiSelect feature]"
    DragEnabled = mbDragEnabled
End Property

Public Property Let DragEnabled(ByVal Enabled As Boolean)
    mbDragEnabled = Enabled
End Property

Public Property Get DragExpandTime() As Long
Attribute DragExpandTime.VB_Description = "Auto Expand Parent node timer (milliseconds) during Drag'n'Drop operation [Default = 1000ms (1 seconds)]"
    DragExpandTime = moDragExpand.Interval
End Property

Public Property Let DragExpandTime(ByVal New_DragExpandTime As Long)
    moDragExpand.Interval = New_DragExpandTime
End Property

Public Property Get DragScrollTime() As Long
Attribute DragScrollTime.VB_Description = "Auto scroll timer (milliseconds) during Drag'n'Drop operation [Default = 200ms (0.2 seconds)]"
    DragScrollTime = moDragScroll.Interval
End Property

Public Property Let DragScrollTime(ByVal New_DragScrollTime As Long)
    moDragScroll.Interval = New_DragScrollTime
End Property

Public Property Get EditLabelMaxLength() As Long                            '@@ v02.02.00

    Dim hwnd As Long

    hwnd = SendMessageAny(oTree.hwnd, TVM_GETEDITCONTROL, 0, 0)
    EditLabelMaxLength = SendMessageAny(hwnd, EM_GETLIMITTEXT, 0&, 0)

End Property

Public Property Let EditLabelMaxLength(ByVal NewLength As Long)             '@@ v02.02.00

    Dim hwnd As Long

    If NewLength < 0 Then NewLength = 0
    hwnd = SendMessageAny(oTree.hwnd, TVM_GETEDITCONTROL, 0, 0)
    SendMessageAny hwnd, EM_SETLIMITTEXT, NewLength, 0

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "True disabled appearance! Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = mbEnabled
End Property

Public Property Let Enabled(ByVal bState As Boolean)

    Dim tvi   As TVITEM
    Dim hItem As Long

    If Not mbLocked Then
        tvi.Mask = TVIF_STATE
        hItem = SendMessageLong(oTree.hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0&)
        Do While hItem <> 0
            With tvi
                .hItem = hItem
                .Mask = TVIF_STATE
                .StateMask = TVIS_CUT
                If (bState) Then                            '## Toggle the cut icon states
                    .State = .StateMask And Not TVIS_CUT
                Else
                    .State = .StateMask Or TVIS_CUT
                End If
                SendMessageAny oTree.hwnd, TVM_SETITEM, 0&, tvi
            End With
            hItem = SendMessageLong(oTree.hwnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, hItem)
        Loop
    End If
    oTree.Enabled = bState
    mbEnabled = bState
End Property

Public Property Get FirstVisibleNode() As MSComctlLib.Node
Attribute FirstVisibleNode.VB_Description = "Returns the first viewable node"
    '
    '## Returns the first viewable node
    '
    Dim hItem   As Long
    Dim selNode As MSComctlLib.Node
    
    LockWindowUpdate mlHWndParent
    With oTree
        '## Remember the node currently selected
        Set selNode = .SelectedItem
        '## Get the handle of the first visible Node
        hItem = SendMessageAny(.hwnd, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, ByVal 0&)
        '## Make it the selected Node
        SendMessageAny .hwnd, TVM_SELECTITEM, TVGN_CARET, ByVal hItem
        '## Return the result as a Node object
        Set FirstVisibleNode = .SelectedItem
        '## Restore node that was selected
        Set .SelectedItem = selNode
    End With
    LockWindowUpdate 0

End Property

Public Property Set FirstVisibleNode(ByVal Node As MSComctlLib.Node)
    '
    '## Set the first visible Node of a TreeView control
    '
    Dim hItem   As Long
    Dim selNode As MSComctlLib.Node
    
    LockWindowUpdate mlHWndParent
    With oTree
        '## Remember the node currently selected
        Set selNode = .SelectedItem
        '## Make the Node the select Node in the control
        Set .SelectedItem = Node
        '## Now we can get its handle
        hItem = SendMessageAny(.hwnd, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
        '## Restore node that was selected
        Set .SelectedItem = selNode
        '## Make it the fist visible Node
        SendMessageAny .hwnd, TVM_SELECTITEM, TVGN_FIRSTVISIBLE, ByVal hItem
    End With
    LockWindowUpdate 0

End Property

Public Property Get FullKeyPath(Optional ByVal oNode As MSComctlLib.Node) As String
Attribute FullKeyPath.VB_Description = "Returns a path full of keys instead of node.text seperated with the TreeView's Seperator."
    '
    '## Returns a node's path of keys
    '
    Dim skeys() As String
    Dim iMax    As Integer
    Dim iloop   As Integer

    '## Retrieve keys for node passed
    skeys = ChildParentKeys(oNode)
    '## Build path string
    iMax = UBound(skeys)
    For iloop = 1 To iMax
        FullKeyPath = FullKeyPath + skeys(iloop) + oTree.PathSeparator
    Next
    FullKeyPath = VBA.Left$(FullKeyPath, Len(FullKeyPath) - 1)

End Property

Public Property Get LastVisibleNode() As MSComctlLib.Node
Attribute LastVisibleNode.VB_Description = "Returns the last visible node."
    '
    '## Returns the last visible node
    '
    Dim lLoop    As Long
    Dim hItem    As Long
    Dim hItemOld As Long
    Dim selNode  As MSComctlLib.Node
    
    LockWindowUpdate mlHWndParent
    With oTree
        '## Remember the node currently selected
        Set selNode = .SelectedItem
        '## Get the handle of the first visible Node
        hItem = SendMessageAny(.hwnd, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, ByVal 0&)
        For lLoop = 1 To .GetVisibleCount - 1
            '## Get the handle of the next visible Node
            hItem = SendMessageAny(.hwnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, ByVal hItem)
            '## Was a Node handle returned?
            If hItem Then
                hItemOld = hItem
            Else
                '## Less items visible than .GetVisibleCount
                hItem = hItemOld
                Exit For
            End If
        Next
        '## Make it the selected Node
        SendMessageAny .hwnd, TVM_SELECTITEM, TVGN_CARET, ByVal hItem
        '## Return the result as a Node object
        Set LastVisibleNode = .SelectedItem
        '## Restore node that was selected
        Set .SelectedItem = selNode
    End With
    LockWindowUpdate 0

End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = " Determines whether a control can be edited."
   Locked = mbLocked
End Property

Public Property Let Locked(ByVal bState As Boolean)

    Dim tvi   As TVITEM
    Dim hItem As Long
   
    mbLocked = bState

    If oTree.Enabled Then
        tvi.Mask = TVIF_STATE
        hItem = SendMessageLong(oTree.hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0&)
        Do While hItem <> 0
            With tvi
                .hItem = hItem
                .Mask = TVIF_STATE
                .StateMask = TVIS_CUT
                If (bState) Then                            '## Toggle the cut icon states
                    .State = .StateMask Or TVIS_CUT
                Else
                    .State = .StateMask And Not TVIS_CUT
                End If
                SendMessageAny oTree.hwnd, TVM_SETITEM, 0&, tvi
            End With
            hItem = SendMessageLong(oTree.hwnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, hItem)
        Loop
    End If

End Property

Public Property Get NodeNestingLevel(ByVal Node As MSComctlLib.Node) As Integer
Attribute NodeNestingLevel.VB_Description = "Returns number of nesting/branch levels for a specific node."
    '
    '## Returns the nesting level of a TreeView's Node object
    '   (returns zero for root nodes.)
    '
    Do Until (Node.Parent Is Nothing)
        NodeNestingLevel = NodeNestingLevel + 1
        Set Node = Node.Parent
    Loop

End Property

Public Property Get OverlayIcon(Optional ByVal Node As MSComctlLib.Node) As eItemOverlayIndices
Attribute OverlayIcon.VB_Description = "Returns/Sets a Node's Overlay Icon."

    Dim tvi   As TVITEM
    Dim hItem As Long

    '## Get Node handle
    hItem = pGetNodeHandle(Node)
    With oTree
        If hItem Then
            '## Retrieve Overlay Icon Flag
            tvi.Mask = TVIF_STATE
            tvi.hItem = hItem
            tvi.StateMask = TVIS_OVERLAYMASK
            If SendMessageAny(.hwnd, TVM_GETITEM, 0&, tvi) Then
                OverlayIcon = CByte((tvi.State And TVIS_OVERLAYMASK) / (2 ^ 8))
            End If
        End If
    End With

End Property

Public Property Let OverlayIcon(Optional ByVal Node As MSComctlLib.Node, Index As eItemOverlayIndices)

    Dim tvi   As TVITEM
    Dim hItem As Long

    '## Get Node handle
    hItem = pGetNodeHandle(Node)
    With oTree
        If hItem Then
            '## Set Overlay Icon Flag
            tvi.Mask = TVIF_STATE
            tvi.hItem = hItem
            tvi.State = CLng(Index) * (2 ^ 8)
            tvi.StateMask = TVIS_OVERLAYMASK
            SendMessageAny .hwnd, TVM_SETITEM, 0&, tvi
        End If
    End With

End Property

Public Property Get ParentNodeText(ByVal Node As MSComctlLib.Node) As String
Attribute ParentNodeText.VB_Description = "Returns the parent node text (Safe routine to stop error condition if no parent node)."
    '
    '## Returns the parent text
    '   NOTE: If a Node has no parent then VB raises an error.
    '         This is a much safer Node.Parent.Text routine.
    With Node
        ParentNodeText = IIf(.FullPath = .Text, "", .Parent.Text)
    End With
End Property

Public Property Get ToolTips() As Boolean
Attribute ToolTips.VB_Description = "Enable/Disable Treeview ToolTips."
    ToolTips = mbToolTips
End Property

Public Property Let ToolTips(ByVal NewState As Boolean)
    mbToolTips = NewState
    With oTree
        If NewState Then
            '## Enable ToolTips (Default )
            SetWindowLong .hwnd, GWL_STYLE, GetWindowLong(.hwnd, GWL_STYLE) And Not TVS_NOTOOLTIPS
        Else
            '## Disable ToolTips
            SetWindowLong .hwnd, GWL_STYLE, GetWindowLong(.hwnd, GWL_STYLE) Or TVS_NOTOOLTIPS
        End If
    End With
End Property

'===========================================================================
' Public subroutines and functions
'
Public Sub ClearTreeView()
Attribute ClearTreeView.VB_Description = "A very fast way of clearing all nodes."
    '
    '## Clear the treeview
    '
    Dim hItem As Long
    Dim hwnd  As Long

    Redraw False
    hwnd = oTree.hwnd
    Do
        hItem = SendMessageLong(hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
        If hItem <= 0 Then Exit Do
        SendMessageLong hwnd, TVM_DELETEITEM, 0&, hItem
    Loop
    Redraw True

End Sub

Public Sub CollapseAll()
Attribute CollapseAll.VB_Description = "Expands all parent and sub-parent nodes of the entire TreeView."
    '
    '## Collapse all nodes
    '
    Dim Node As MSComctlLib.Node

    For Each Node In oTree.Nodes
        With Node
            If .Children Then .Expanded = False
        End With
    Next

End Sub

Public Sub CollapseChildNodes(ByVal Node As MSComctlLib.Node)
Attribute CollapseChildNodes.VB_Description = "Cascading effect. Hides all children for each parent node sub-branch."

    '@@ v02.02.00 Old code was buggy - replaced with new efficient routine!
    pSubNodesExpandState Node, False
'    Dim lCount As Long
'    Dim lLoop  As Long
'
'    If Not Node Is Nothing Then
'        With Node
'            .Expanded = False
'
'            If .Children Then               '## Does node have children?
'                lCount = .Children
'                Set Node = .Child
'                For lLoop = 1 To lCount
'                    CollapseChildNodes Node '## Yes, collapse them as well (recursive)
'                    If lLoop < lCount Then
'                        Set Node = .Next
'                    End If
'                Next
'            End If
'        End With
'    End If

End Sub

Public Sub ExpandAll()
Attribute ExpandAll.VB_Description = "Expands all parent and sub-parent nodes of the entire TreeView."
    '
    '## Expand all nodes
    '
    Dim Node As MSComctlLib.Node

    On Error GoTo ErrorHandler                      '@@ v01.00.01
    For Each Node In oTree.Nodes
        With Node
            If .Children Then .Expanded = True
        End With
    Next
    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 35606: ExpandAll
    End Select
End Sub

Public Sub ExpandChildNodes(ByVal Node As MSComctlLib.Node)
Attribute ExpandChildNodes.VB_Description = "Cascading effect. Expand all parent & sub-parent nodes that belong to Node."

    '@@ v02.02.00 Old code was buggy - replaced with new efficient routine!
    pSubNodesExpandState Node, True
'    Dim lCount As Long
'    Dim lLoop  As Long
'
'    With Node
'        If Not Node Is Nothing Then
'            .Expanded = True
'        End If
'
'        If .Children Then               '## Does node have children?
'            lCount = .Children
'            Set Node = .Child
'            For lLoop = 1 To lCount
'                ExpandChildNodes Node   '## Yes, expand them as well (recursive)
'                If lLoop < lCount Then
'                    Set Node = .Next
'                End If
'            Next
'        End If
'    End With

End Sub

'@@ v02.02.00
Public Function HitTestEx(ByVal x As Single, _
                          ByVal y As Single, _
                 Optional ByRef Part As eHitTestEx, _
                 Optional ByVal Button As Integer, _
                 Optional ByVal Capture As Boolean = False, _
                 Optional ByVal TestRow As Boolean = True) As MSComctlLib.Node

    Dim tHitTestInfo As TVHITTESTINFO
    Dim oHitNode     As Node
    Dim bHit         As Boolean
    Dim snX          As Single
    Dim tRect        As RECT
    
    With oTree
        If Capture And CBool(Button) Then
            SetCapture .hwnd
        End If
        Set oHitNode = .HitTest(x, y)
    End With

    With tHitTestInfo
        With .pt
            .x = x \ Screen.TwipsPerPixelX
            .y = y \ Screen.TwipsPerPixelX
        End With
        bHit = CBool(SendMessageAny(oTree.hwnd, TVM_HITTEST, 0, tHitTestInfo))
        Part = .flags
    End With

    If bHit Then
        If oHitNode Is Nothing Then
            If TestRow Then
                Select Case Part
                    Case [Hit Item Button], [Hit Item Indent]
                        With oTree
                            GetClientRect .hwnd, tRect
                            For snX = x To tRect.Right * Screen.TwipsPerPixelX Step Abs(TestRow) * Screen.TwipsPerPixelX
                                Set oHitNode = .HitTest(snX, y)
                                If Not (oHitNode Is Nothing) Then
                                    Set HitTestEx = oHitNode
                                    Exit For
                                End If
                            Next
                        End With
                    Case [Hit Item Right]
                        With oTree
                            For snX = x To 0 Step Abs(TestRow) * -Screen.TwipsPerPixelX
                                Set oHitNode = .HitTest(snX, y)
                                If Not (oHitNode Is Nothing) Then
                                    Set HitTestEx = oHitNode
                                    Exit For
                                End If
                            Next
                        End With
                End Select
            End If
        Else
            Set HitTestEx = oHitNode
            If (Part And TVHT_ONITEMSTATEICON) = TVHT_ONITEMSTATEICON Then
                Part = [Hit Item Check]
            ElseIf (Part And TVHT_ONITEMICON) = TVHT_ONITEMICON Then
                Part = [Hit Item Icon]
            ElseIf (Part And TVHT_ONITEMLABEL) = TVHT_ONITEMLABEL Then
                Part = [Hit Item Label]
            End If
        End If
    End If

    If Capture Then
        If Part >= TVHT_ABOVE Then
            If Button = 0 Then
                ReleaseCapture
            End If
        End If
    End If

End Function

#If NODLL = 0 Then
    Public Sub HookCtrl(ByRef Ctrl As Object, Optional OpMode As eOpMode = [Single Select])
Attribute HookCtrl.VB_Description = "Primes the cTreeview class to manage the  Treeview control"

    If Not (TypeOf Ctrl Is MSComctlLib.TreeView) Then
        Err.Raise clTREEHOOK_ERR, csSOURCE_ERR, csTREEHOOK_ERR  '@@ v02.01.00
        Exit Sub
    End If

#Else
    
    Public Sub HookCtrl(ByRef Ctrl As MSComctlLib.TreeView, Optional OpMode As eOpMode = [Single Select])

#End If

    Dim bState As Boolean
    Dim rc     As RECT

    Set oTree = Ctrl                                '## Capture the Treeview control.
                                                    '   Class won't work if this isn't called first.
    With oTree
        #If DEBUGMODE = 1 Then
            dbgCtrlName = .Parent.Name + "." + .Name
        #End If
        '
        '## calculate the height (pixels) of a node
        '
        bState = .Scroll                            '## Remember if scrollbars are enabled
        .Scroll = False                             '## Turn them off
        ''mlNodeHeight = .Height \ .GetVisibleCount   '## calculate height
        .Scroll = bState                            '## restore scrollbar state
        mlHwnd = .hwnd
        mlHWndParent = .Parent.hwnd                 '@@ v02.03.00
    End With
    meOpMode = OpMode

    Dim lStyle As Long
    lStyle = GetWindowLong(oTree.hwnd, GWL_EXSTYLE)
    '
    '## Drag'n'Drop primer
    '
    mbDragEnabled = False
    '
    '## Establish the distance in which auto-scrolling happens within
    '   the TreeView's client area (we need a root item for these calls)
    '
    rc.Left = SendMessageAny(oTree.hwnd, TVM_GETNEXTITEM, ByVal TVGN_ROOT, ByVal 0&)
    If SendMessageAny(oTree.hwnd, TVM_GETITEMRECT, ByVal 1, rc) Then
        mlAutoScroll = (rc.Bottom - rc.Top) * 2
    Else
        mlAutoScroll = 32
    End If
    '
    '## Get system Start Drag Zone metrics
    '
    mszDrag.cx = GetSystemMetrics(SM_CXDRAG)
    mszDrag.cy = GetSystemMetrics(SM_CYDRAG)
    '
    '## Initialise timers & their default timings
    '
    Set moDragScroll = New XTimer
    moDragScroll.Interval = 200     '## 0.2 seconds
    Set moDragExpand = New XTimer
    moDragExpand.Interval = 1000    '## 1 second

    Set moSubclass = New GSubclass                              '@@ v02.02.00

    With moSubclass                                             '@@
        '
        '## Used for Disabled Nodes & Lock mode
        '
        .AttachMessage Me, mlHwnd, WM_LBUTTONDOWN               '@@
        .AttachMessage Me, mlHwnd, WM_LBUTTONUP                 '@@
        .AttachMessage Me, mlHwnd, WM_LBUTTONDBLCLK             '@@
        '
        '## Used for painting background color
        '
        .AttachMessage Me, mlHwnd, WM_PAINT                     '@@ v02.03.00
        .AttachMessage Me, mlHwnd, WM_ERASEBKGND                '@@
        .AttachMessage Me, mlHWndParent, WM_NOTIFY              '@@
        .AttachMessage Me, mlHwnd, WM_DESTROY                   '@@
    End With

End Sub

Public Function IsChildNode(ByVal TestNode As MSComctlLib.Node, _
                            ByVal ParentNode As MSComctlLib.Node, _
                   Optional ByVal NodeDepth As eNodeDepth = [First Branch]) As Boolean
Attribute IsChildNode.VB_Description = "Returns if a node is a direct child of another node or whether a node exists on another node's branch."

    If NodeDepth = [First Branch] Then
        If ParentNode.Children Then
            If Not IsRootNode(TestNode) Then
                IsChildNode = (TestNode.Parent = ParentNode)
            End If
        End If
    Else
        IsChildNode = IsParentNode(TestNode, ParentNode)
    End If

End Function

'@@ v02.02.00
Public Function IsNodeTextUnique(ByVal Node As MSComctlLib.Node, _
                                 ByVal Text As String, _
                        Optional ByVal CompareMethod As VbCompareMethod = vbTextCompare) As Boolean
Attribute IsNodeTextUnique.VB_Description = "Check's whether Text passed is already used by a node or by nodes."

    Dim oNode As MSComctlLib.Node

    For Each oNode In oTree.Nodes
        If Not (oNode Is Node) Then
            If StrComp(oNode.Text, Text, CompareMethod) = 0 Then
                Exit Function
            End If
        End If
    Next
    IsNodeTextUnique = True

End Function

Public Function IsParentNode(ByVal ChildNode As MSComctlLib.Node, _
                             ByVal ParentNode As MSComctlLib.Node) As Boolean
Attribute IsParentNode.VB_Description = "Checks if one node is the parent of another."
    '
    '## Checks if one node is the parent of another.
    '   This is a recursive routine that steps down through
    '   the branches of the parent node.
    '
    Dim lNDX As Long

    If (ParentNode Is Nothing) Or (ChildNode Is Nothing) Then
        Exit Function
    End If
    If ParentNode.Children Then             '## Does the parent node have children?
        lNDX = ParentNode.Child.Index       '## Yes, remember the first child
        Do                                  '## Step through all child nodes
            If lNDX = ChildNode.Index Then  '## is ChildNode the test node?
                IsParentNode = True         '## ParentNode is the parent of ChildNode.
                Exit Do
            End If
            If IsParentNode(ChildNode, oTree.Nodes(lNDX)) Then  '## Step down through the branches
                IsParentNode = True         '## ParentNode is the parent of ChildNode.
                Exit Do
            End If
            If lNDX <> ParentNode.Child.LastSibling.Index Then  '## Have we tested the last child node?
                lNDX = oTree.Nodes(lNDX).Next.Index             '## No. Point to the next child node
            Else
                Exit Do                                         '## Yes.
            End If
        Loop
    End If

End Function

Public Function IsPathExpanded(ByVal Node As MSComctlLib.Node) As Boolean   '@@ v02.02.00
Attribute IsPathExpanded.VB_Description = "Returns whether a node is expanded without affecting the node's expand state."
    '
    '## Checks (recursively) if a Node's path is expanded (back to the Root-level node
    '
    Dim oParentNode As MSComctlLib.Node

    With Node
        If IsRootNode(Node) Then
            IsPathExpanded = .Expanded
        Else
            Set oParentNode = Node.Parent
            If oParentNode.Expanded Then
                IsPathExpanded = IsPathExpanded(oParentNode)
            End If
        End If
    End With

End Function

Public Function IsRootNode(ByVal Node As MSComctlLib.Node) As Boolean
Attribute IsRootNode.VB_Description = "Check is selected node is the root node."
    '
    '## Check if selected node is a root node.
    '
    With Node
        IsRootNode = (.FullPath = .Text)
    End With

End Function

'@@ v02.01.00 Added new method
Public Function IsScrollbarVisible(WhichScroll As eIsScrollbarVisible) As Boolean
Attribute IsScrollbarVisible.VB_Description = "Returns whether the designated scrollbar(s) visible."

'    IsScrollbarVisible = (CByte(Abs((GetWindowLong(oTree.hWnd, GWL_STYLE) And WS_HSCROLL) <> 0) * 1) + _
'                               (Abs((GetWindowLong(oTree.hWnd, GWL_STYLE) And WS_VSCROLL) <> 0) * 2) _
'                          And WhichScroll) = WhichScroll
'!! v02.01.01 Optimised function

    Dim lTest As Long

    lTest = WhichScroll * WS_HSCROLL
    IsScrollbarVisible = ((GetWindowLong(oTree.hwnd, GWL_STYLE) And lTest) = lTest)

End Function

Public Function Load(ByVal Filename As String, _
            Optional ByVal FileMode As eFileMode) As Boolean
Attribute Load.VB_Description = "Load Treeview from Binary or XML stored file."
    Select Case FileMode
        Case [Binary]: Load = pLoadBIN(Filename)
        Case [xml]:    Load = pLoadXML(Filename)
    End Select
End Function

'@@ v02.02.00
Public Function NodeAbsoluteIndex(ByVal Node As MSComctlLib.Node) As Long

    Dim oNode  As MSComctlLib.Node
    Dim lCount As Long
    
    lCount = 1
    If Not (Node Is Node.FirstSibling) Then
        Set oNode = Node.Previous
        Do While Not (oNode Is Nothing)
            lCount = lCount + pNodeCountChildren(oNode, False, True) + 1
            Set oNode = oNode.Previous
        Loop
    End If
    Set oNode = Node.Parent
    If Not (oNode Is Nothing) Then
        lCount = lCount + NodeAbsoluteIndex(oNode)
    End If
    NodeAbsoluteIndex = lCount

End Function

'@@ v02.02.00
Public Function NodeAbsolutePosition(ByVal Node As MSComctlLib.Node, _
                            Optional ByVal VisibleOnly As Boolean = True) As Long

    If Node.Visible Or Not VisibleOnly Then
        NodeAbsolutePosition = pNodeAbsolutePosition(Node, VisibleOnly)
    End If

End Function

Public Function NodeAdd(Optional ByVal vRelative As Variant, _
                        Optional ByVal eRelationship As MSComctlLib.TreeRelationshipConstants, _
                        Optional ByVal sKey As String, _
                        Optional ByVal sText As String, _
                        Optional ByVal vImage As Variant, _
                        Optional ByVal vSelectedImage As Variant, _
                        Optional ByVal vTag As Variant, _
                        Optional ByVal bBold As Boolean = False, _
                        Optional ByVal bChecked As Boolean = False, _
                        Optional ByVal bEnsureVisible As Boolean = False, _
                        Optional ByVal bExpanded As Boolean = True, _
                        Optional ByVal bSelected As Boolean = False, _
                        Optional ByVal bSorted As Boolean = False, _
                        Optional ByVal lForeColor As OLE_COLOR, _
                        Optional ByVal lBackColor As OLE_COLOR, _
                        Optional ByVal vExpandedImage As Variant) As MSComctlLib.Node
Attribute NodeAdd.VB_Description = "Adds node to TreeView and sets properties in one call."

    Dim Node As MSComctlLib.Node

    On Error GoTo ErrorHandler

    Set Node = oTree.Nodes.Add(vRelative, eRelationship, sKey, sText, vImage, vSelectedImage)
    With Node
        #If DEBUGMODE = 1 Then
            'Debug.Print "ADD-> Text = "; .Text; "   Index = "; CStr(.Index)
        #End If
        .Tag = vTag
        .Bold = bBold
        .Checked = bChecked
        If bEnsureVisible Then .EnsureVisible
        .Expanded = bExpanded
        .Selected = bSelected
        .Sorted = bSorted
        If lForeColor = 0 Then
            .ForeColor = vbWindowText
        Else
            .ForeColor = lForeColor
        End If
        If lBackColor = 0 Then
            .BackColor = vbWindowBackground
        Else
            .BackColor = lBackColor
        End If
        If Not IsMissing(vExpandedImage) Then
            .ExpandedImage = vExpandedImage
        End If
    End With
    
    Set NodeAdd = Node
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'@@ v02.02.00
Public Sub NodeCoord(ByRef Left As Single, _
                     ByRef Top As Single, _
                     ByRef Width As Single, _
                     ByRef Height As Single, _
                  Optional HostType As eNodeCoords = [By Node])
Attribute NodeCoord.VB_Description = "Return's a Nodes coordinates"

    Dim rcNode   As RECT
    Dim rcClient As RECT
    '
    '## Retrieve selected node's co-ordinates and covert to the treeview's parent
    '   container's scale mode.
    '
    rcNode = pGetSelectedNodeCoords
    With rcNode
        pScale CSng(.Left), _
               CSng(.Top), _
               Left, _
               Top, _
               vbPixels, _
               oTree.Parent.ScaleMode
        pScale CSng(.Right - .Left), _
               CSng(.Bottom - .Top), _
               Width, _
               Height, _
               vbPixels, _
               oTree.Parent.ScaleMode
    End With
    '
    '## Set coordinates according to Host Type
    '
    If HostType = [By TreeView] Then
        With oTree
            GetClientRect .hwnd, rcClient
            With .Parent
                Width = .ScaleX((rcClient.Right - rcNode.Left) + 1, vbPixels, .ScaleMode)
            End With
        End With
    End If
End Sub

Public Function NodeCopy(ByVal DestNode As MSComctlLib.Node, _
                         ByVal SrcNode As MSComctlLib.Node, _
                Optional ByVal IncludeChildren As Boolean = True) As Boolean    '@@ v01.00.03
Attribute NodeCopy.VB_Description = "Framework to copy one node to another. Generates CopyNode event for user controlled copying."

    Dim oSNode  As MSComctlLib.Node
    Dim oDNode  As MSComctlLib.Node
    Dim bCancel As Boolean
    '
    '## Ensure that we don't overwrite the original nodes values
    '
    Set oSNode = SrcNode
    Set oDNode = DestNode

    If oSNode = oDNode Then
        '## Same node - therefore no point trying to copy
        Exit Function
    End If
    '
    '## Force the Nodes to be expanded before the copy operation is performed. This
    '   will allow child nodes to be loaded if done on demand
    '
    oDNode.Expanded = True
    oSNode.Expanded = True
    '
    '## Let the actual process happen externally to cTREEVIEW - a must if working with a database
    '   or multiple types of data.
    '
    RaiseEvent CopyNode(oDNode, oSNode, bCancel)     '## Get the calling routine to copy the node.
    If bCancel Then                                  '## Was the process cancelled?
        NodeCopy = False
        Exit Function                                '## Yes.
    End If
    If IncludeChildren Then                                     '## Include all child nodes?
        If oSNode.Children Then                                 '## Yes. Are there child nodes?
            Set oSNode = oSNode.Child.FirstSibling              '## Yes. Select the first child.
            Do
                If (oDNode.Child Is Nothing) Then
                    NodeCopy = False
                    Err.Raise clTREECOPY_ERR1, "COPY NODE::" + csSOURCE_ERR, csTREECOPY_ERR1
                    Exit Function
                End If
                If NodeCopy(oDNode.Child.LastSibling, oSNode, IncludeChildren) Then '## Copy child node
                    If Not (oSNode = oSNode.LastSibling) Then   '## Are there more child nodes?
                        Set oSNode = oSNode.Next                '## Yes.
                    Else
                        Exit Do
                    End If
                Else
                    NodeCopy = False
                    Exit Function
                End If
            Loop
        End If
    End If
    NodeCopy = True                                         '## Success!

End Function

'@@ v02.02.00
Public Function NodeCountChildren(ByVal Node As MSComctlLib.Node, _
                         Optional ByVal VisibleOnly As Boolean = False, _
                         Optional ByVal ForceExpanded As Boolean = False) As Long
    
    NodeCountChildren = pNodeCountChildren(Node, VisibleOnly, ForceExpanded)

End Function

Public Function NodeDelete(ByVal Node As MSComctlLib.Node, _
                  Optional ByVal bSelect As Boolean = True, _
                  Optional ByVal bIncludeChilden = True) As Boolean
Attribute NodeDelete.VB_Description = "Deletes a Node and optionally it's child nodes"

    Dim lCount As Long
    Dim lLoop  As Long

    With Node
        If Not IsRootNode(Node) Then    '## Is this a root node?
            If Not bIncludeChilden Then
                '## Move the node's children to the Node's parent
                lCount = .Children
                For lLoop = 1 To lCount
                    Set oTree.Nodes(.Child.Index).Parent = .Parent
                Next
            End If
            With .Parent                '## No.
                .EnsureVisible
                .Selected = bSelect
            End With
        End If
        '## Delete node
        oTree.Nodes.Remove .Index       '## Delete the node
    End With
    NodeDelete = True

End Function

Function NodeExist(ByVal Key As String) As Boolean
Attribute NodeExist.VB_Description = "Returns whether a specified node exists."

    Dim sTmp As String

    On Local Error GoTo ErrorHandler

    sTmp = TypeName(oTree.Nodes(Key))
    NodeExist = True
    Exit Function

ErrorHandler:
'
End Function

Public Function NodeFind(ByRef oResultNode As MSComctlLib.Node, _
                         ByVal NodeText As String, _
                Optional ByVal NodeKey As String, _
                Optional ByVal bSelect As Boolean = True) As Boolean
Attribute NodeFind.VB_Description = "Finds a Node and optionally selects & ensures visibility. (Now replaced by NodeFindEx & kept for backward compatibility)"

    Dim Node      As MSComctlLib.Node
    Dim bCheckKey As Boolean
    Dim bKeyOK    As Boolean

    bCheckKey = (Len(Trim$(NodeKey)) > 0)
    bKeyOK = (Not bCheckKey)
    For Each Node In oTree.Nodes
        With Node
            If UCase$(.Text) = UCase$(NodeText) Then
                If bCheckKey Then bKeyOK = (UCase$(.Key) = UCase$(NodeKey))
                If bKeyOK Then
                    If bSelect Then
                        .Selected = True
                        .EnsureVisible
                    End If
                    Set oResultNode = Node
                    NodeFind = True
                    Exit For
                End If
            End If
        End With
    Next

End Function

'@@ v02.02.00
Public Function NodeFindEx(ByVal Text As String, _
                  Optional ByVal Mode As eNodeFindEx = [Find Node Text], _
                  Optional ByVal CaseSensitive As Boolean = False) As Collection
Attribute NodeFindEx.VB_Description = "Returns a collection of Nodes found based on search criteria. (Replaces NodeFind)"

    Dim oNode  As MSComctlLib.Node
    Dim bFound As Boolean
    Dim oNodes As Collection

    Set oNodes = New Collection
    For Each oNode In oTree.Nodes
        bFound = False
        With oNode
            If (Mode And [Find Node Text]) = [Find Node Text] Then
                If CaseSensitive Then
                    If InStr(.Text, Text) Then
                        oNodes.Add oNode, .Key
                        bFound = True
                    End If
                Else
                    If InStr(UCase$(.Text), UCase$(Text)) Then
                        oNodes.Add oNode, oNode.Key
                        bFound = True
                    End If
                End If
            End If
            If (Mode And [Find Node Key]) = [Find Node Key] Then
                If CaseSensitive Then
                    If InStr(.Key, Text) Then
                        If Not bFound Then
                            oNodes.Add oNode, .Key
                        End If
                    End If
                Else
                    If InStr(UCase$(.Key), UCase$(Text)) Then
                        If Not bFound Then
                            oNodes.Add oNode, .Key
                        End If
                    End If
                End If
            End If
        End With
    Next
    If oNodes.Count Then
        Set NodeFindEx = oNodes
    End If

End Function

'@@ v02.02.00
Public Function NodeFindEx2(ByVal LikeText As String, _
                   Optional ByVal StartNode As MSComctlLib.Node, _
                   Optional ByVal FindMode As eNodeFindEx = [Find Node Text], _
                   Optional ByVal CompareMethod As VbCompareMethod = vbTextCompare) As Collection

    Dim oNodes As Collection
    Dim oNode  As MSComctlLib.Node

    Set oNodes = New Collection
    If StartNode Is Nothing Then
        If oTree.Nodes.Count Then
            Set oNode = pGetFirstNode
            Select Case CompareMethod
                Case vbBinaryCompare
                    Select Case FindMode
                        Case [Find Node Text]
                            Do
                                pNodeFindTextCompareBinary oNode, LikeText, oNodes
                                Set oNode = oNode.Next
                            Loop Until oNode Is Nothing

                        Case [Find Node Key]
                            Do
                                pNodeFindKeyCompareBinary oNode, LikeText, oNodes
                                Set oNode = oNode.Next
                            Loop Until oNode Is Nothing

                        Case [Find Node Text And Key]
                            Do
                                pNodeFindBothCompareBinary oNode, LikeText, oNodes
                                Set oNode = oNode.Next
                            Loop Until oNode Is Nothing

                        Case Else
                            Exit Function

                    End Select

                Case vbTextCompare
                    Select Case FindMode
                        Case [Find Node Text]
                            Do
                                pNodeFindTextCompareText oNode, LikeText, oNodes
                                Set oNode = oNode.Next
                            Loop Until oNode Is Nothing

                        Case [Find Node Key]
                            Do
                                pNodeFindKeyCompareText oNode, LikeText, oNodes
                                Set oNode = oNode.Next
                            Loop Until oNode Is Nothing

                        Case [Find Node Text And Key]
                            Do
                                pNodeFindBothCompareText oNode, LikeText, oNodes
                                Set oNode = oNode.Next
                            Loop Until oNode Is Nothing
                        Case Else
                            Exit Function

                    End Select

                Case Else
                    Exit Function

            End Select
        Else
            Exit Function
        End If
    Else
        Select Case CompareMethod
            Case vbBinaryCompare
                Select Case FindMode
                    Case [Find Node Text]
                        pNodeFindTextCompareBinary StartNode, LikeText, oNodes

                    Case [Find Node Key]
                        pNodeFindKeyCompareBinary StartNode, LikeText, oNodes

                    Case [Find Node Text And Key]
                        pNodeFindBothCompareBinary StartNode, LikeText, oNodes

                    Case Else
                        Exit Function

                End Select

            Case vbTextCompare
                Select Case FindMode
                    Case [Find Node Text]
                        pNodeFindTextCompareText StartNode, LikeText, oNodes

                    Case [Find Node Key]
                        pNodeFindKeyCompareText StartNode, LikeText, oNodes

                    Case [Find Node Text And Key]
                        pNodeFindBothCompareText StartNode, LikeText, oNodes

                    Case Else
                        Exit Function

                End Select

            Case Else
                Exit Function

        End Select
    End If
    Set NodeFindEx2 = oNodes

End Function

'@@ v02.02.00
Public Function NodeFindObjectInTag(ByVal FindObject As Object, _
                           Optional ByVal StartNode As MSComctlLib.Node, _
                           Optional ByVal OnlyFirstOccurance As Boolean = True) As Object

    Dim oNodes As Object            '## Either a node object or collection object
    Dim oNode  As MSComctlLib.Node

    If Not OnlyFirstOccurance Then
        Set oNodes = New Collection
    End If

    If StartNode Is Nothing Then
        If oTree.Nodes.Count Then
            Set oNode = pGetFirstNode
            Do
                pNodeFindObjectInTag oNode, FindObject, oNodes, OnlyFirstOccurance
                If OnlyFirstOccurance And Not (oNodes Is Nothing) Then
                    Exit Do
                End If
                Set oNode = oNode.Next
            Loop Until oNode Is Nothing
        Else
            Exit Function
        End If
    Else
        pNodeFindObjectInTag StartNode, FindObject, oNodes, OnlyFirstOccurance
    End If
    Set NodeFindObjectInTag = oNodes

End Function

Public Function NodeMove(ByVal TargetNode As MSComctlLib.Node, _
                         ByRef SourceNode As MSComctlLib.Node, _
                Optional ByVal bSelect As Boolean = True, _
                Optional ByVal eRelationship As TreeRelationshipConstants = tvwChild) As Boolean
Attribute NodeMove.VB_Description = "Moves a node within the treeview. (Can be used to turn a child node into a root node)"

    Dim lLoop    As Long
    Dim lCount   As Long
    Dim oFVNode  As MSComctlLib.Node
    Dim oTmpNode As MSComctlLib.Node
    Dim oNewNode As MSComctlLib.Node
    Dim bIsXpd1 As Boolean                                                  '@@ v02.03.01
    Dim bIsXpd2 As Boolean                                                  '@@

    If (TargetNode Is Nothing) Then Exit Function
    If (SourceNode Is Nothing) Then Exit Function
    If TargetNode.Index = SourceNode.Index Then Exit Function

    If IsParentNode(TargetNode, SourceNode) Then
        NodeMove = pNodeSwap(TargetNode, SourceNode, bSelect)
        Exit Function
    End If

    '
    '## Remember the first visible node
    '
    Set oFVNode = FirstVisibleNode
    '
    '## Hide control whilst moving nodes around to stop flickering
    '
    Redraw False
    '
    '## 1. Create a temp node for the kids
    '
    If IsRootNode(SourceNode) Then
        With SourceNode
            Set oTmpNode = NodeAdd(, , _
                                   .Key + "[Tmp]", _
                                   .Text, _
                                   .Image, _
                                   .SelectedImage, _
                                   .Tag, _
                                   .Bold, _
                                   .Checked, , , , _
                                   .Sorted, _
                                   .ForeColor, _
                                   .BackColor, _
                                   .ExpandedImage)
        End With
    Else
        With SourceNode
            Set oTmpNode = NodeAdd(.Parent.Key, _
                                   tvwChild, _
                                   .Key + "[Tmp]", _
                                   .Text, _
                                   .Image, _
                                   .SelectedImage, _
                                   .Tag, _
                                   .Bold, _
                                   .Checked, , , , _
                                   .Sorted, _
                                   .ForeColor, _
                                   .BackColor, _
                                   .ExpandedImage)
        End With
    End If
    bIsXpd1 = SourceNode.Expanded                                           '@@ v02.03.01
    '
    '## 2. Move the kids to the foster parent
    '
    lCount = SourceNode.Children
    For lLoop = 1 To lCount
        Set oTree.Nodes(SourceNode.Child.Index).Parent = oTmpNode
    Next
    '
    '## 3. Delete the old parent
    '
    oTree.Nodes.Remove SourceNode.Index
    '
    '## 4. Add a new parent
    '
    With oTmpNode
        Set oNewNode = NodeAdd(TargetNode.Key, _
                               eRelationship, _
                               Replace(.Key, "[Tmp]", ""), _
                               .Text, _
                               .Image, _
                               .SelectedImage, _
                               .Tag, _
                               .Bold, _
                               .Checked, , , _
                               True, _
                               .Sorted, _
                               .ForeColor, _
                               .BackColor, _
                               .ExpandedImage)
    End With
    '
    '## 5. Give the kids to the new parent
    '
    lCount = oTmpNode.Children
    For lLoop = 1 To lCount
        With oTree.Nodes(oTmpNode.Child.Index)
            bIsXpd2 = .Expanded                                             '@@ v02.03.01
            Set .Parent = oNewNode
            .Expanded = bIsXpd2                                             '@@ v02.03.01
        End With
    Next
    oNewNode.Expanded = bIsXpd1                                             '@@ v02.03.01
    '
    '## 6. Delete the foster parent
    '
    oTree.Nodes.Remove oTmpNode.Index
    '
    '## Make sure that the control only scrolls if it's necessary.
    '
    On Error Resume Next    '## If oFVNode = oNewNode then an error will be
                            '   raised as the reference node was deleted but
                            '   VB doesn't destory oFVNode. So the test If
                            '   Not (oFVNode is Nothing) then ... will
                            '   generate an error!
    oFVNode.EnsureVisible
    With oNewNode
        If .Children > 0 And bIsXpd1 = True Then                            '@@ v02.03.01
            .Child.LastSibling.EnsureVisible
        End If
        .EnsureVisible
    End With
    '
    '## Show changes made
    '
    Redraw True
    NodeMove = True

End Function

'@@ v02.02.00
Public Function NodePosition(ByVal Node As MSComctlLib.Node) As Long

    Dim lIndex As Long
    
    lIndex = 1
    Set Node = Node.Previous
    Do While Not (Node Is Nothing)
        lIndex = lIndex + 1
        Set Node = Node.Previous
    Loop
    NodePosition = lIndex

End Function

Public Function NodeRename(ByVal Node As MSComctlLib.Node, _
                           ByVal NewNodeText As String, _
                  Optional ByVal bSelect As Boolean = True) As Boolean
Attribute NodeRename.VB_Description = "Changes the text of a node and optionally selects & ensures visibility."

    With Node
        If Len(NewNodeText) Then            '## Check if we have new text
            .Text = NewNodeText
            If Not IsRootNode(Node) Then
                .EnsureVisible
                .Selected = bSelect
            End If
            NodeRename = True
        End If
    End With

End Function

'@@ v02.02.00
Public Function NodeUniqueText(ByVal Node As MSComctlLib.Node, _
                               ByVal NewText As String, _
                      Optional ByVal CompareMethod As VbCompareMethod = vbTextCompare, _
                      Optional ByVal PrefixNumberMode As eNodeUniquePrefixNumberMode = Numbers, _
                      Optional ByVal Prefix As String = csDefaultNodePrefix) As String
Attribute NodeUniqueText.VB_Description = "Returns unique text based on parameters passed."

    Dim sText  As String
    Dim lCount As Long
    
    sText = NewText
    Select Case PrefixNumberMode
        Case [Numbers]
            lCount = 1
            Do
                If IsNodeTextUnique(Node, sText, CompareMethod) Then
                    Exit Do
                Else
                    lCount = lCount + 1
                    sText = NewText & "(" + CStr(lCount) + ")"
                End If
            Loop
        Case [Prefix]
            Do
                If IsNodeTextUnique(Node, sText, CompareMethod) Then
                    Exit Do
                Else
                    sText = Prefix & sText
                End If
            Loop
    End Select
    NodeUniqueText = sText

End Function

Public Sub Redraw(ByVal Enabled As Boolean)
Attribute Redraw.VB_Description = "Enable/disable screen updates."

    Dim hwnd  As Long

    If Not (oTree Is Nothing) Then
        hwnd = oTree.hwnd
        If Enabled Then
            '## Lock the Parent window update to avoid flickering
            LockWindowUpdate 0
            'SendMessageLong hwnd, WM_SETREDRAW, False, &H0
        Else
            '## Unlock the Parent window
            LockWindowUpdate mlHWndParent
            'SendMessageLong hwnd, WM_SETREDRAW, True, &H0
        End If
    End If

End Sub

Public Function Save(ByVal Filename As String, _
            Optional ByVal FileMode As eFileMode) As Boolean
Attribute Save.VB_Description = "Save Treeview to Binary or XML file."
    Select Case FileMode
        Case [Binary]: Save = pSaveBIN(Filename)
        Case [xml]:    Save = pSaveXML(Filename)
    End Select
End Function

Public Sub ScrollView(ByVal Dir As eCodeScrollView)
Attribute ScrollView.VB_Description = "Scrolls the treview using code in all four directions."
    '
    '## Scrolls the treview using code
    '
    Dim lHwnd As Long
    Dim lPos  As Long
    Dim lBar1 As Long
    Dim lBar2 As Long
    Dim lDir  As Long

    lHwnd = oTree.hwnd
    Select Case Dir
        Case [Home]:       SendMessageAny lHwnd, WM_VSCROLL, SB_TOP, vbNull
        Case [Page Up]:    SendMessageAny lHwnd, WM_VSCROLL, SB_PAGEUP, vbNull
        Case [Up]:         SendMessageAny lHwnd, WM_VSCROLL, SB_LINEUP, vbNull
        Case [Down]:       SendMessageAny lHwnd, WM_VSCROLL, SB_LINEDOWN, vbNull
        Case [Page Down]:  SendMessageAny lHwnd, WM_VSCROLL, SB_PAGEDOWN, vbNull
        Case [End]:        SendMessageAny lHwnd, WM_VSCROLL, SB_BOTTOM, vbNull
        Case [Left]:       SendMessageAny lHwnd, WM_HSCROLL, SB_LEFT, vbNull
        Case [Page Left]:  SendMessageAny lHwnd, WM_HSCROLL, SB_PAGELEFT, vbNull
        Case [Line Left]:  SendMessageAny lHwnd, WM_HSCROLL, SB_LINELEFT, vbNull
        Case [Line Right]: SendMessageAny lHwnd, WM_HSCROLL, SB_LINERIGHT, vbNull
        Case [Page Right]: SendMessageAny lHwnd, WM_HSCROLL, SB_PAGERIGHT, vbNull
        Case [Right] ':    SendMessageAny lHwnd, WM_HSCROLL, SB_RIGHT, vbNull
            '
            '## For some reason, the treeview doesn't respond to the above commented
            '   out message. Therefore a work-around is required.
            '
            '## To stop flickering, the control is hidden temporarily.
            Redraw False
            ' ## Loop until we've scrolled to the far right side
            Do
                lPos = GetScrollPos(lHwnd, 0&)
                SendMessageAny lHwnd, WM_HSCROLL, SB_PAGERIGHT, vbNull
            Loop Until (lPos = GetScrollPos(lHwnd, 0&))
            '## Now show the control
            Redraw True
    End Select

End Sub

Public Sub SetFocusNode(ByVal KeyIndex As Variant)
Attribute SetFocusNode.VB_Description = "Shows the focus rectangle without actually selecting the node."
    With oTree.Nodes(KeyIndex)
        .Selected = True
        .Selected = False
    End With
End Sub

Public Sub SetOverlayImage(ByVal Imaglelisthandle As OLE_HANDLE, _
                           ByVal ImageListIndex As Long, _
                           ByVal ShareOverlay As eItemOverlayIndices)
Attribute SetOverlayImage.VB_Description = "Initialises Overlay Icon Images. Must be used before setting the Node's OverlayIcon state image."
    If Imaglelisthandle Then
        ImageList_SetOverlayImage Imaglelisthandle, ImageListIndex, ShareOverlay
    End If
End Sub

'@@ v02.02.00
Public Sub SubNodesSorted(Optional ByVal Node As MSComctlLib.Node, _
                          Optional ByVal State As Boolean = True)
Attribute SubNodesSorted.VB_Description = "Sets the sort properties of the specified node and all child nodes for all sub-branch levels.."

    Dim oChildNode As Node

    If (Node Is Nothing) Then
        Set Node = oTree.Nodes(1)
    End If
    With Node
        .Sorted = State
        If .Children Then
            Set oChildNode = .Child
            Do
                SubNodesSorted oChildNode, State
                Set oChildNode = oChildNode.Next
            Loop Until oChildNode Is Nothing
        End If
    End With
End Sub

Public Sub UnHookCtrl()
Attribute UnHookCtrl.VB_Description = "Releases the hooked TreeView control."

    Dim lLoop As Long
    '
    '## Release the hooked control
    '
    Set oTree = Nothing
    '
    '## Shut down subclassing
    '
    If Not (moSubclass Is Nothing) Then
        With moSubclass
            .DetachMessage Me, mlHwnd, WM_LBUTTONDOWN               '@@ v02.02.00
            .DetachMessage Me, mlHwnd, WM_LBUTTONUP                 '@@
            .DetachMessage Me, mlHwnd, WM_LBUTTONDBLCLK             '@@
            .DetachMessage Me, mlHwnd, WM_DESTROY                   '@@ v02.03.00
            .DetachMessage Me, mlHwnd, WM_PAINT                     '@@
            .DetachMessage Me, mlHwnd, WM_ERASEBKGND                '@@
            .DetachMessage Me, mlHWndParent, WM_NOTIFY              '@@
        End With
    End If
    '
    '## Clean up memory DCs
    '
    For lLoop = 1 To 3                                          '@@ v02.03.00
        pClearUpDC lLoop                                        '@@
    Next                                                        '@@

End Sub

'===========================================================================
' Disable Node: Public properties, subroutines and functions    '@@ v02.02.00
'
Private Property Let iDisable_BackColor(ByVal RHS As stdole.OLE_COLOR)

    Dim vItem As Variant
    
    Select Case RHS
        Case mlDisBackColor
            '## Already set
        Case Else
            Redraw False
            mlDisBackColor = RHS
            On Error Resume Next
            '## Update all disabled items
            For Each vItem In moDisNodes
                oTree.Nodes(vItem([Node Key])).BackColor = mlDisBackColor
                If Err.Number Then
                    Redraw True
                    Exit Property
                End If
            Next
            Redraw True
    End Select

End Property

Private Property Get iDisable_AllowSelect() As Boolean
    iDisable_AllowSelect = mbDisAllowSel
End Property

Private Property Let iDisable_AllowSelect(ByVal RHS As Boolean)
    mbDisAllowSel = RHS
End Property

Private Property Get iDisable_BackColor() As stdole.OLE_COLOR
    iDisable_BackColor = mlDisBackColor
End Property

Private Property Let iDisable_Bold(ByVal RHS As Boolean)

    Dim vItem As Variant
    
    Select Case RHS
        Case mbDisBold
            '## Already set
        Case Else
            Redraw False
            mbDisBold = RHS
            On Error Resume Next
            '## Update all disabled items
            For Each vItem In moDisNodes
                oTree.Nodes(vItem([Node Key])).Bold = mbDisBold
                If Err.Number Then
                    Redraw True
                    Exit Property
                End If
            Next
            Redraw True
    End Select

End Property

Private Property Get iDisable_Bold() As Boolean
    iDisable_Bold = mbDisBold
End Property

Private Property Let iDisable_ForeColor(ByVal RHS As stdole.OLE_COLOR)

    Dim vItem As Variant

    Select Case RHS
        Case mlDisForeColor
            '## Already set
        Case Else
            Redraw False
            mlDisForeColor = RHS
            On Error Resume Next
            '## Update all disabled items
            For Each vItem In moDisNodes
                oTree.Nodes(vItem([Node Key])).ForeColor = mlDisForeColor
                If Err.Number Then
                    Redraw True
                    Exit Property
                End If
            Next
            Redraw True
    End Select

End Property

Private Property Get iDisable_ForeColor() As stdole.OLE_COLOR
    iDisable_ForeColor = mlDisForeColor
End Property

Private Property Get iDisable_FocusNode() As MSComctlLib.INode
    Set iDisable_FocusNode = oTree.SelectedItem
End Property

Private Property Get iDisable_DisabledNode(ByVal Index As Long) As Collection

    If moDisNodes.Count Then
        '## Return the requested Node from the collection
        Set iDisable_DisabledNode = oTree.Nodes(moDisNodes(Index, [Node Key]))
    End If

End Property

Private Property Get iDisable_SelectMode() As eDisableAdjustSelectMode
    iDisable_SelectMode = meDisAdjSelect
End Property

Private Property Let iDisable_SelectMode(ByVal RHS As eDisableAdjustSelectMode)
    meDisAdjSelect = RHS
    With oTree
        If Not (.SelectedItem Is Nothing) Then
            pAdjustSelection .SelectedItem
        End If
    End With
End Property

Private Sub iDisable_DisableAllNodes(Optional Node As MSComctlLib.INode, Optional ByVal State As Boolean = True)

    Dim oNode  As MSComctlLib.Node

    If Node Is Nothing Then
        '## Select the first node as none were passed
        Set oNode = pGetFirstNode()
    Else
        Set oNode = Node
    End If
    If oNode Is Nothing Then
        '## Treeview has no nodes
        Exit Sub
    End If

    If (moDisNodes.Count = 0) And (State = False) Then                      '@@ v02.03.01
        '## Nothing to do...                                                '@@
        Exit Sub                                                            '@@
    End If                                                                  '@@

    Redraw False
    pDisableAllNodes oNode, State
    pAdjustSelection oNode
    Redraw True
    RaiseEvent StateChange

End Sub

Private Sub iDisable_DisableMultiSelect(Optional ByVal State As Boolean = True) '@@ v02.03.00

    Dim oNode As MSComctlLib.Node

    If meOpMode = [Multi Select] Then
        If mSelectedNodes.Count Then                                '## Do we have any nodes selected?
            Redraw False
            For Each oNode In mSelectedNodes                        '## For every selected node
                Select Case State
                    Case True
                        If Not moDisNodes.Exist(ObjPtr(oNode)) Then '## If not disabled:-
                            pSetNodeSelect oNode, (Not State)       '## a. remove selected node
                            pSetNodeState oNode, State              '## b. then make it so.
                        End If
                    Case False
                        If moDisNodes.Exist(ObjPtr(oNode)) Then     '## If disabled,
                            pSetNodeState oNode, State              '## then enable it.
                        End If
                End Select
            Next
            If State = True Then
                pAdjustSelection moFocusNode
            End If
            Redraw True
            RaiseEvent StateChange
        End If
    Else
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Sub
    End If

End Sub

Private Sub iDisable_DisableNode(Node As MSComctlLib.INode, Optional ByVal State As Boolean = True)

    Dim lPtr As Long

    If Node Is Nothing Then
        Exit Sub
    End If

    Redraw False
    pSetNodeState Node, State
    pAdjustSelection Node
    Redraw True
    RaiseEvent StateChange

End Sub

Private Sub iDisable_ToggleDisabled(Optional Node As MSComctlLib.INode)

    Dim oNode  As MSComctlLib.Node

    If (Node Is Nothing) Then
        '## Select the first node as none were passed
        Set oNode = pGetFirstNode()
    Else
        Set oNode = Node
    End If
    If oNode Is Nothing Then
        '## Treeview has no nodes
        Exit Sub
    End If

    '## hide the Treeview to speed up the operation
    Redraw False
    pToggleDisabled oNode
    pAdjustSelection oNode
    Redraw True
    RaiseEvent StateChange

End Sub

Private Sub iDisable_ToggleNodeState(Optional Node As MSComctlLib.INode, Optional ByVal Children As Boolean)

    Dim oNode     As MSComctlLib.Node
    Dim oChild    As MSComctlLib.Node
    'Dim bDisabled As Boolean
    Dim lPtr      As Long

    If (Node Is Nothing) Then
        '## Select the focus node as none were passed
        Set oNode = oTree.SelectedItem
    Else
        Set oNode = Node
    End If
    If oNode Is Nothing Then
        '## Treeview has no nodes
        Exit Sub
    End If

    Redraw False
    lPtr = ObjPtr(oNode)    '## Tag collection node ID
    On Error Resume Next
    Select Case moDisNodes.Exist(lPtr)
        Case True
            pSetNodeState oNode, False
        Case False
            pSetNodeState oNode, True
            'bDisabled = True
    End Select
    If moDisNodes.Count Then
        If CBool(oNode.Children) And Children Then
            Set oNode = oNode.Child
            Do While Not (oNode Is Nothing)
                pToggleDisabled oNode
                'pDisableNode oNode, bDisabled, True
                Set oNode = oNode.Next
            Loop
        End If
    End If
    pAdjustSelection Node
    Redraw True
    RaiseEvent StateChange

End Sub

Private Function iDisable_NewEnum() As stdole.IUnknown

    '## Pass the enumeration pointers to the Implementation class. Implementation class
    '   will then expose the enumeration to the calling routine. Who said that the
    '   Interface class had no functionality...
    Set iDisable_NewEnum = mDisabledNodes.[_NewEnum]

End Function

'===========================================================================
' MultiNode Selection: Public properties, subroutines and functions
'
Private Property Let iMultiSelect_NoClearOnSpaceClick(ByVal RHS As Boolean)

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Property
    End If

    mbSelNoClear = RHS

End Property

Private Property Get iMultiSelect_NoClearOnSpaceClick() As Boolean
    iMultiSelect_NoClearOnSpaceClick = mbSelNoClear
End Property

Private Property Let iMultiSelect_NoDefaultSel(ByVal RHS As Boolean)

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Property
    End If

    mbSelNoDefault = RHS

End Property

Private Property Get iMultiSelect_NoDefaultSel() As Boolean
    iMultiSelect_NoDefaultSel = mbSelNoDefault
End Property

Private Property Let iMultiSelect_SelBackColor(ByVal RHS As stdole.OLE_COLOR)

    Dim vItem As Variant
    
    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Property
    End If

    Select Case RHS
        Case mlSelBackColor
            '## Already set
        Case Else
            Redraw False
            mlSelBackColor = RHS
            On Error Resume Next
            '## Update all selected items
            For Each vItem In moSelNodes
                oTree.Nodes(vItem([Node Key])).BackColor = mlSelBackColor
                If Err.Number Then
                    Redraw True
                    Exit Property
                End If
            Next
            Redraw True
    End Select

End Property

Private Property Get iMultiSelect_SelBackColor() As stdole.OLE_COLOR
    iMultiSelect_SelBackColor = mlSelBackColor
End Property

Private Property Let iMultiSelect_SelBold(ByVal RHS As Boolean)

    Dim vItem As Variant
    
    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Property
    End If

    Select Case RHS
        Case mbSelBold
            '## Already set
        Case Else
            Redraw False
            mbSelBold = RHS
            On Error Resume Next
            '## Update all selected items
            For Each vItem In moSelNodes
                oTree.Nodes(vItem([Node Key])).Bold = mbSelBold
                If Err.Number Then
                    Redraw True
                    Exit Property
                End If
            Next
            Redraw True
    End Select

End Property

Private Property Get iMultiSelect_SelBold() As Boolean
    iMultiSelect_SelBold = mbSelBold
End Property

Private Property Let iMultiSelect_SelForeColor(ByVal RHS As stdole.OLE_COLOR)

    Dim vItem As Variant
    
    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Property
    End If

    Select Case RHS
        Case mlSelForeColor
            '## Already set
        Case Else
            Redraw False
            mlSelForeColor = RHS
            On Error Resume Next
            '## Update all selected items
            For Each vItem In moSelNodes
                oTree.Nodes(vItem([Node Key])).ForeColor = mlSelForeColor
                If Err.Number Then
                    Redraw True
                    Exit Property
                End If
            Next
            Redraw True
    End Select

End Property

Private Property Get iMultiSelect_SelForeColor() As stdole.OLE_COLOR
    iMultiSelect_SelForeColor = mlSelForeColor
End Property

Private Property Let iMultiSelect_FocusNode(ByVal RHS As MSComctlLib.INode) '@@ v02.02.00
    pSetFocusNode RHS
End Property

Private Property Get iMultiSelect_FocusNode() As MSComctlLib.INode
    Set iMultiSelect_FocusNode = moFocusNode
End Property

Private Property Get iMultiSelect_SelectedNode(ByVal Index As Long) As Collection

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Property
    End If

    If moSelNodes.Count Then
        '## Return the requested Node from the collection
        Set iMultiSelect_SelectedNode = oTree.Nodes(moSelNodes(Index, [Node Key]))
    End If

End Property

Private Sub iMultiSelect_ClearSelection(Optional ByRef Node As MSComctlLib.INode)

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Sub
    End If

    pClearSelection Node        '@@ v02.02.00 - Optimised for speed by minimising refresh
    RaiseEvent SelChange

End Sub

Private Sub iMultiSelect_SelectAllNodes(Optional ByRef Node As MSComctlLib.INode)

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Sub
    End If

    Redraw False
    pSelectAllNodes Node        '@@ v02.02.00 - Optimised for speed by minimising refresh
    Redraw True
    RaiseEvent SelChange

End Sub

Private Sub iMultiSelect_SelectNode(ByRef Node As MSComctlLib.INode, _
                                    ByVal Selected As Boolean)

    Dim lPtr As Long

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Sub
    End If

    If Node Is Nothing Then
        Exit Sub
    End If

    'redraw False
    pSetNodeSelect Node, Selected
    'redraw True
    RaiseEvent SelChange

End Sub

Private Sub iMultiSelect_ToggleNode(Optional ByRef Node As MSComctlLib.INode, _
                                    Optional ByVal Children As Boolean)

    Dim oNode     As MSComctlLib.Node
    Dim oChild    As MSComctlLib.Node
    Dim bSelected As Boolean
    Dim lPtr      As Long

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Sub
    End If

    If (Node Is Nothing) Then
        '## Select the focus node as none were passed
        Set oNode = moFocusNode
    Else
        Set oNode = Node
    End If
    If oNode Is Nothing Then
        '## Treeview has no nodes
        Exit Sub
    End If

    Redraw False
    lPtr = ObjPtr(oNode)    '## Tag collection node ID
    On Error Resume Next
    Select Case moSelNodes.Exist(lPtr)              '@@ v02.02.00
        Case True                                   '@@
            pSetNodeSelect oNode, False             '@@ Fixes problem if:-
        Case False                                  '@@   normal backcolor = selection back color
            pSetNodeSelect oNode, True              '@@
    End Select                                      '@@
'    Select Case oNode.BackColor                    '@@
'        Case mlSelBackColor                        '@@
'            pSetNodeSelect oNode, False            '@@
'        Case moSelNodes(lPtr, [Node BackColor])         '@@
'            pSetNodeSelect oNode, True             '@@
'            bSelected = True                       '@@
'    End Select                                     '@@
    If moSelNodes.Count Then
        If CBool(oNode.Children) And Children Then
            Set oNode = oNode.Child
            Do While Not (oNode Is Nothing)
                pSelectNode oNode, bSelected, True
                Set oNode = oNode.Next
            Loop
        End If
    End If
    Redraw True
    RaiseEvent SelChange

End Sub

Private Sub iMultiSelect_ToggleSelection(Optional ByRef Node As MSComctlLib.INode)

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Sub
    End If

    '## hide the Treeview to speed up the operation
    Redraw False
    pToggleSelection Node           '@@ v02.02.00 - Optimised for speed by minimising refresh
    Redraw True
    RaiseEvent SelChange

End Sub

Private Function iMultiSelect_NewEnum() As stdole.IUnknown

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Function
    End If

    '## Pass the enumeration pointers to the Implementation class. Implementation class
    '   will then expose the enumeration to the calling routine. Who said that the
    '   Interface class had no functionality...
    Set iMultiSelect_NewEnum = mSelectedNodes.[_NewEnum]

End Function

'===========================================================================
' Disable Node: Private subroutines and functions                               '@@ v02.02.00
'
Private Sub pDisableAllNodes(Optional ByVal Node As MSComctlLib.Node, Optional ByVal State As Boolean = True)

    Dim oChild As MSComctlLib.Node
    Dim lPtr   As Long

    Do
        lPtr = ObjPtr(Node)    '## Tag collection node ID
        Select Case State
            Case True   '## Disable
                If Not moDisNodes.Exist(lPtr) Then
                    '## Node not already Disabled
                    pSetNodeState Node, True
                End If

            Case False  '## Enable
                If moDisNodes.Exist(lPtr) Then
                    '## Node not already Disabled
                    pSetNodeState Node, False
                End If

        End Select
        If moDisNodes.Count Then
            If Node.Children Then
                '## Node has child nodes, repeat for all child nodes
                Set Node = Node.Child
                Do
                    pDisableAllNodes Node, State
                    If Not (Node = Node.LastSibling) Then
                        Set Node = Node.Next
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If Not (Node Is Nothing) Then
            Exit Do
        Else
            If Not (Node = Node.LastSibling) Then
                Set Node = Node.Next
            Else
                Exit Do
            End If
        End If
    Loop

End Sub

Private Sub pDisableNode(ByRef Node As MSComctlLib.Node, _
                         ByVal State As Boolean, _
               Optional ByVal ChildState As Boolean = False)

    Dim oChild As MSComctlLib.Node

    On Error Resume Next
    '## Set Disable Node state
    pSetNodeState Node, State
    If ChildState Then
        '## Set all child nodes disabled state as well
        Set oChild = Node.Child
        Do While Not (oChild Is Nothing)
            pDisableNode oChild, State, ChildState
            Set oChild = oChild.Next
        Loop
    End If

End Sub

Private Sub pSetNodeState(ByRef Node As MSComctlLib.Node, _
                          ByVal State As Boolean)

    Dim lPtr As Long

    lPtr = ObjPtr(Node) '## Tag collection node ID
    With Node
        Select Case State
            Case False
                On Error Resume Next
                '## Get original settings
                Me.CutIcon(Node) = False
                .BackColor = moDisNodes(lPtr, [Node BackColor])
                .ForeColor = moDisNodes(lPtr, [Node ForeColor])
                .Bold = moDisNodes(lPtr, [Node Bold])
                If Not Err.Number Then
                    '## No longer required - therefore remove
                    moDisNodes.Remove lPtr
                    mDisabledNodes.Remove CStr(lPtr)
                End If
            Case True
                '## Save original settings before selecting
                moDisNodes.Add lPtr, .Key, .ForeColor, .BackColor, .Bold
                mDisabledNodes.Add Node, CStr(lPtr)
                Debug.Print .Text
                Me.CutIcon(Node) = True
                .BackColor = mlDisBackColor
                .ForeColor = mlDisForeColor
                .Bold = mbDisBold
        End Select
    End With

End Sub

'@@ v02.02.00
Private Sub pAdjustSelection(ByVal Node As MSComctlLib.Node, _
                    Optional ByVal LookForward As Boolean = True)
    '
    '## Tells the treeview what to do if a Node's disabled state was altered...
    '
    If mbDisAllowSel Then Exit Sub      '## Selection is allowed, so igore...

    Select Case meDisAdjSelect
        Case [eiDis Select Direction]
            If LookForward Then
                Set Node = pGetNextEnabled(Node)
            Else
                Set Node = pGetPrevEnabled(Node)
            End If
            If Node Is Nothing Then
                Set oTree.SelectedItem = Nothing
            Else
                Set oTree.SelectedItem = Node
                '
                '## Better tell the Multi-node select system that a node
                '   was selected...
                '
                If meOpMode = [Multi Select] Then       '## v02.03.00
                    pSetNodeSelect Node, True
                End If
                RaiseEvent SelChange
            End If
        Case [eiDis Focus Rectangle]
            SetFocusNode Node.Key
            RaiseEvent SelChange
        Case [eiDis Select None]
            Set oTree.SelectedItem = Nothing
            RaiseEvent SelChange
    End Select

End Sub

Private Sub pToggleDisabled(Optional ByVal Node As MSComctlLib.Node)

    Dim oChild As MSComctlLib.Node
    Dim lPtr   As Long

    Do
        lPtr = ObjPtr(Node)    '## Tag collection node ID
        On Error Resume Next
        Select Case moDisNodes.Exist(lPtr)
            Case True:  pSetNodeState Node, False
            Case False: pSetNodeState Node, True
        End Select
        If moDisNodes.Count Then
            If Node.Children Then
                Set Node = Node.Child
                Do
                    pToggleDisabled Node
                    If Not (Node = Node.LastSibling) Then
                        Set Node = Node.Next
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If Not (Node Is Nothing) Then
            Exit Do
        Else
            If Not (Node = Node.LastSibling) Then
                Set Node = Node.Next
            Else
                Exit Do
            End If
        End If
    Loop

End Sub

'===========================================================================
' MultiNode Selection: Private subroutines and functions
'
Private Sub pClearSelection(Optional ByVal Node As MSComctlLib.Node)            '@@ v02.02.00

    Dim oChild As MSComctlLib.Node
    Dim lPtr   As Long

    Dim oNode  As MSComctlLib.Node

    If Node Is Nothing Then
        '## Select the first node as none were passed
        'Set oNode = pGetFirstNode()
        For Each oChild In mSelectedNodes
            pSetNodeSelect oChild, False
        Next
        Exit Sub
    Else
        Set oNode = Node
    End If
'    If oNode Is Nothing Then
'        '## Treeview has no nodes
'        Exit Sub
'    End If

    Do
        lPtr = ObjPtr(oNode)    '## Tag collection node ID
        If moSelNodes.Exist(lPtr) Then
            '## Node exists so remove node
            pSetNodeSelect oNode, False
        End If
        If moSelNodes.Count Then
            If oNode.Children Then
                '## Node has child nodes, repeat for all child nodes
                Set oNode = oNode.Child
                Do
                    pClearSelection oNode
                    If Not (oNode = oNode.LastSibling) Then
                        Set oNode = oNode.Next
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If Not (Node Is Nothing) Then
            Exit Do
        Else
            If Not (oNode = oNode.LastSibling) Then
                Set oNode = oNode.Next
            Else
                Exit Do
            End If
        End If
    Loop

End Sub

Private Sub pSelectAllNodes(Optional ByRef Node As MSComctlLib.Node)            '@@ v02.02.00

    Dim oNode  As MSComctlLib.Node
    Dim oChild As MSComctlLib.Node
    Dim lPtr   As Long

    If Node Is Nothing Then
        '## Select the first node as none were passed
        Set oNode = pGetFirstNode()
    Else
        Set oNode = Node
    End If
    If oNode Is Nothing Then
        '## Treeview has no nodes
        Exit Sub
    End If

    Do
        lPtr = ObjPtr(oNode)    '## Tag collection node ID
        If Not moSelNodes.Exist(lPtr) Then
            '## Node not already selected
            pSetNodeSelect oNode, True
        End If
        If moSelNodes.Count Then
            If oNode.Children Then
                '## Node has child nodes, repeat for all child nodes
                Set oNode = oNode.Child
                Do
                    pSelectAllNodes oNode
                    If Not (oNode = oNode.LastSibling) Then
                        Set oNode = oNode.Next
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If Not (Node Is Nothing) Then
            Exit Do
        Else
            If Not (oNode = oNode.LastSibling) Then
                Set oNode = oNode.Next
            Else
                Exit Do
            End If
        End If
    Loop

End Sub

Private Sub pSelectNode(ByRef Node As MSComctlLib.Node, _
                        ByVal Selected As Boolean, _
               Optional ByVal ChildSelect As Boolean = False)

    Dim oChild As MSComctlLib.Node

    On Error Resume Next
    '## Set select Node state
    pSetNodeSelect Node, Selected
    If ChildSelect Then
        '## Set all child nodes select state as well
        Set oChild = Node.Child
        Do While Not (oChild Is Nothing)
            pSelectNode oChild, Selected, True
            Set oChild = oChild.Next
        Loop
    End If

End Sub

Private Sub pSetFocusNode(Node As MSComctlLib.Node)                             '@@ v02.02.00
    'redraw False
    With Node
        .Selected = True
        .Selected = False
    End With
    Set moFocusNode = Node
    'redraw True
End Sub

Private Sub pSetNodeSelect(ByRef Node As MSComctlLib.Node, _
                           ByVal State As Boolean)

    Dim lPtr As Long

    lPtr = ObjPtr(Node) '## Tag collection node ID
    With Node
        Select Case State
            Case False
                On Error Resume Next
                '## Get original settings
                .BackColor = moSelNodes(lPtr, [Node BackColor])
                .ForeColor = moSelNodes(lPtr, [Node ForeColor])
                .Bold = moSelNodes(lPtr, [Node Bold])
                If Not Err.Number Then
                    '## No longer required - therefore remove
                    moSelNodes.Remove lPtr
                    mSelectedNodes.Remove CStr(lPtr)
                End If
            Case True
                '## Save original settings before selecting
                moSelNodes.Add lPtr, .Key, .ForeColor, .BackColor, .Bold
                mSelectedNodes.Add Node, CStr(lPtr)
                .BackColor = mlSelBackColor
                .ForeColor = mlSelForeColor
                .Bold = mbSelBold
        End Select
    End With

End Sub

Private Sub pToggleSelection(Optional ByRef Node As MSComctlLib.Node)           '@@ v02.02.00

    Dim oNode  As MSComctlLib.Node
    Dim oChild As MSComctlLib.Node
    Dim lPtr   As Long

    If (Node Is Nothing) Then
        '## Select the first node as none were passed
        Set oNode = pGetFirstNode()
    Else
        Set oNode = Node
    End If
    If oNode Is Nothing Then
        '## Treeview has no nodes
        Exit Sub
    End If

    Do
        lPtr = ObjPtr(oNode)    '## Tag collection node ID
        On Error Resume Next
        Select Case moSelNodes.Exist(lPtr)              '@@ v02.02.00
            Case True:  pSetNodeSelect oNode, False     '@@ Fixes problem if:-
            Case False: pSetNodeSelect oNode, True      '@@   normal backcolor = selection back color
        End Select                                      '@@
'        Select Case oNode.BackColor                                            '@@
'            Case mlSelBackColor:                pSetNodeSelect oNode, False    '@@
'            Case moSelNodes(lPtr, [Node BackColor]): pSetNodeSelect oNode, True     '@@
'        End Select                                                             '@@
        If moSelNodes.Count Then
            If oNode.Children Then
                Set oNode = oNode.Child
                Do
                    'iMultiSelect_ToggleSelection oNode
                    pToggleSelection oNode                  '@@ v02.02.00 Fixes control flickering
                    If Not (oNode = oNode.LastSibling) Then
                        Set oNode = oNode.Next
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If Not (Node Is Nothing) Then
            Exit Do
        Else
            If Not (oNode = oNode.LastSibling) Then
                Set oNode = oNode.Next
            Else
                Exit Do
            End If
        End If
    Loop

End Sub

'===========================================================================
' Drag'n'Drop: Private subroutines and functions
'
Private Sub pDoDrag()

    Dim pt         As POINTAPI
    Dim rcClient   As RECT
    Static lOldNdx As Long

    With oTree
        If mbStartDrag = True Then
            If mbInDrag = True Then
                '
                '## If the cursor is still over same item as it was on the previous call,
                '   start the auto expand timer, disable the timer otherwise.
                '
                If Not (.DropHighlight Is Nothing) Then
                    If lOldNdx <> .DropHighlight.Index Then
                        If (.DropHighlight.Children > 0) And (.DropHighlight.Expanded = False) Then
                            moDragExpand.Enabled = True
                        Else
                            moDragExpand.Enabled = False
                        End If
                    End If
                    lOldNdx = .DropHighlight.Index
                End If
                '
                '## If the window is scrollable, and the cursor is within that auto scroll
                '   distance, start the auto scroll timer, disable the timer otherwise.
                '
                GetCursorPos pt
                ScreenToClient .hwnd, pt
                GetClientRect .hwnd, rcClient
                If (pIsWindowScrollable(.hwnd) And pPtInRectRegion(rcClient, mlAutoScroll, pt)) Then
                    moDragScroll.Enabled = True
                Else
                    moDragScroll.Enabled = False
                End If
                If Not (.DropHighlight Is Nothing) Then
                    '## We're over a node
                    RaiseEvent Dragging(moDragNode, .DropHighlight)
                End If
            End If
        End If
    End With

End Sub

'===========================================================================
' General: Private BIN/XML Save/Load subroutines and functions
'
Private Sub pInitXML(ByRef oRs As ADODB.Recordset)
    Set oRs = New ADODB.Recordset
    With oRs.Fields
        .Append "PkID", adInteger, , adFldRowID '## Auto Increment field
        .Append "Text", adVarChar, 255
        .Append "Key", adVarChar, 255
        .Append "Tag", adVariant
        .Append "BackColor", adInteger
        .Append "Bold", adBoolean
        .Append "Checked", adBoolean
        .Append "Expanded", adBoolean
        .Append "ExpandedImage", adInteger
        .Append "ForeColor", adInteger
        .Append "Image", adInteger
        .Append "SelectedImage", adInteger
        .Append "Sorted", adBoolean
        .Append "Selected", adBoolean
        .Append "LinkID", adVarChar, 255
        .Append "FirstVisible", adBoolean
    End With
    oRs.Open
End Sub

Private Function pLoadBIN(ByVal Filename As String) As Boolean

    Dim oNode    As MSComctlLib.Node
    Dim tNodes() As udtNodeInfo
    Dim fHndl    As Long
    Dim lRecs    As Long
    Dim lLoop    As Long
    Dim sFirst   As String

    On Error GoTo ErrorHandler

    fHndl = FreeFile(0)
    Open Filename For Binary As fHndl
    Get #fHndl, , lRecs
    ReDim tNodes(1 To lRecs)
    Get #fHndl, , tNodes()
    Close #fHndl

    If lRecs Then
        oTree.Nodes.Clear
        For lLoop = 1 To lRecs
            If tNodes(lLoop).LinkId Then
                Set oNode = NodeAdd(tNodes(lLoop).LinkId, _
                                    tvwChild, _
                                    CStr(tNodes(lLoop).Key), _
                                    CStr(tNodes(lLoop).Text), _
                                    CInt(tNodes(lLoop).Image), _
                                    CInt(tNodes(lLoop).SelImage), _
                                    tNodes(lLoop).Tag, _
                                    tNodes(lLoop).Bold, _
                                    tNodes(lLoop).Checked, _
                                    tNodes(lLoop).Sorted, _
                                    tNodes(lLoop).Expanded, _
                                    tNodes(lLoop).Selected, , _
                                    tNodes(lLoop).ForeColor, _
                                    tNodes(lLoop).BackColor, _
                                    CInt(tNodes(lLoop).ExpImage))
            Else
                Set oNode = NodeAdd(, , _
                                    CStr(tNodes(lLoop).Key), _
                                    CStr(tNodes(lLoop).Text), _
                                    CInt(tNodes(lLoop).Image), _
                                    CInt(tNodes(lLoop).SelImage), _
                                    tNodes(lLoop).Tag, _
                                    tNodes(lLoop).Bold, _
                                    tNodes(lLoop).Checked, _
                                    tNodes(lLoop).Sorted, _
                                    tNodes(lLoop).Expanded, _
                                    tNodes(lLoop).Selected, , _
                                    tNodes(lLoop).ForeColor, _
                                    tNodes(lLoop).BackColor, _
                                    CInt(tNodes(lLoop).ExpImage))
            End If
            If tNodes(lLoop).FirstVisible Then
                sFirst = oNode.Key
            End If
        Next
        If Len(sFirst) Then
            oTree.Nodes(sFirst).EnsureVisible
            Set FirstVisibleNode = oTree.Nodes(sFirst)
        End If
        Set oNode = Nothing
    End If
    pLoadBIN = True
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function pLoadXML(ByVal Filename As String) As Boolean

    Dim oNode   As MSComctlLib.Node
    Dim oRs     As ADODB.Recordset
    Dim fHndl   As Long
    Dim sData   As String
    Dim oStream As Stream
    Dim sFirst  As String

    On Error GoTo ErrorHandler

    Set oStream = New Stream
    Set oRs = New ADODB.Recordset
    pInitXML oRs

    fHndl = FreeFile(0)
    Open Filename For Binary As fHndl
    sData = Space$(LOF(fHndl))
    Get #fHndl, , sData
    Close #fHndl

    With oStream
        .Open
        .WriteText sData
        .Position = 0
    End With
    With oRs
        .Close
        .Open oStream
        .MoveFirst
        If .RecordCount Then
            oTree.Nodes.Clear
            Do
                If Len(!LinkId) Then
                    Set oNode = NodeAdd(CStr(!LinkId), _
                                        tvwChild, _
                                        CStr(!Key), _
                                        CStr(!Text), _
                                        CInt(!Image), _
                                        CInt(!SelectedImage), _
                                        !Tag, _
                                        !Bold, _
                                        !Checked, _
                                        !Sorted, _
                                        !Expanded, _
                                        !Selected, , _
                                        !ForeColor, _
                                        !BackColor, _
                                        CInt(!ExpandedImage))
                Else
                    Set oNode = NodeAdd(, , _
                                        CStr(!Key), _
                                        CStr(!Text), _
                                        CInt(!Image), _
                                        CInt(!SelectedImage), _
                                        !Tag, _
                                        !Bold, _
                                        !Checked, _
                                        !Sorted, _
                                        !Expanded, _
                                        !Selected, , _
                                        !ForeColor, _
                                        !BackColor, _
                                        CInt(!ExpandedImage))
                End If
                If !FirstVisible Then
                    sFirst = oNode.Key
                End If
                .MoveNext
            Loop Until .EOF
            If Len(sFirst) Then
                oTree.Nodes(sFirst).EnsureVisible
                Set FirstVisibleNode = oTree.Nodes(sFirst)
            End If
        End If
        .Close
    End With
    oStream.Close
    Set oStream = Nothing
    Set oRs = Nothing
    pLoadXML = True
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function pSaveBIN(ByVal Filename As String) As Boolean

    Dim oNode    As MSComctlLib.Node
    Dim tNodes() As udtNodeInfo
    Dim fHndl    As Long
    Dim lNDX     As Long

    On Error GoTo ErrorHandler

    ReDim tNodes(1 To oTree.Nodes.Count) As udtNodeInfo
    For Each oNode In oTree.Nodes
        lNDX = lNDX + 1
        With oNode
            tNodes(lNDX).Index = .Index
            tNodes(lNDX).Text = .Text
            tNodes(lNDX).Key = .Key
            tNodes(lNDX).Tag = .Tag
            tNodes(lNDX).BackColor = .BackColor
            tNodes(lNDX).Bold = .Bold
            tNodes(lNDX).Checked = .Checked
            tNodes(lNDX).Expanded = .Expanded
            tNodes(lNDX).ExpImage = .ExpandedImage
            tNodes(lNDX).ForeColor = .ForeColor
            tNodes(lNDX).Image = .Image
            tNodes(lNDX).SelImage = .SelectedImage
            tNodes(lNDX).Sorted = .Sorted
            tNodes(lNDX).Selected = (.Key = oTree.SelectedItem.Key)
            If IsParentNode(oNode, .Parent) Then
                tNodes(lNDX).LinkId = .Parent.Index
            Else
                tNodes(lNDX).LinkId = 0
            End If
            tNodes(lNDX).FirstVisible = (.Key = FirstVisibleNode.Key)
        End With
    Next

    fHndl = FreeFile(0)
    Open Filename For Binary As fHndl
    Put #fHndl, , CLng(oTree.Nodes.Count)
    Put #fHndl, , tNodes()
    Close #fHndl

    Set oNode = Nothing
    pSaveBIN = True
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function pSaveXML(ByVal Filename As String) As Boolean

    Dim oNode   As MSComctlLib.Node
    Dim oRs     As ADODB.Recordset
    Dim fHndl   As Long
    Dim oStream As Stream
    
    On Error GoTo ErrorHandler

    Set oStream = New Stream
    Set oRs = New ADODB.Recordset
    pInitXML oRs

    With oRs
        For Each oNode In oTree.Nodes
            .AddNew
            !Text = oNode.Text
            !Key = oNode.Key
            !Tag = oNode.Tag
            !BackColor = oNode.BackColor
            !Bold = oNode.Bold
            !Checked = oNode.Checked
            !Expanded = oNode.Expanded
            !ExpandedImage = oNode.ExpandedImage
            !ForeColor = oNode.ForeColor
            !Image = oNode.Image
            !SelectedImage = oNode.SelectedImage
            !Sorted = oNode.Sorted
            !Selected = (oNode.Key = oTree.SelectedItem.Key)
            If IsParentNode(oNode, oNode.Parent) Then
                !LinkId = oNode.Parent.Key
            End If
            !FirstVisible = (oNode.Key = FirstVisibleNode.Key)
            .Update
        Next

        fHndl = FreeFile(0)
        Open Filename For Binary As fHndl
        .Save oStream, adPersistXML
        Put #fHndl, , oStream.ReadText
        Close #fHndl
        .Close
    End With
    oStream.Close
    Set oStream = Nothing
    Set oRs = Nothing
    pSaveXML = True
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'===========================================================================
' Node Find: Private Methods                                            '@@ v02.02.00
'
Private Sub pNodeFindTextCompareBinary(ByVal Node As MSComctlLib.Node, _
                                       ByVal FindText As String, _
                                       ByVal Nodes As Collection)

    Dim oNode As MSComctlLib.Node

    With Node
        If LikeCompBinary(.Text, FindText) Then
            Nodes.Add Node
        End If
        If .Children Then
            Set oNode = .Child
            Do
                pNodeFindTextCompareBinary oNode, FindText, Nodes
                Set oNode = oNode.Next
            Loop Until oNode Is Nothing
        End If
    End With

End Sub

Private Sub pNodeFindKeyCompareBinary(ByVal Node As MSComctlLib.Node, _
                                      ByVal FindText As String, _
                                      ByVal Nodes As Collection)

    Dim oNode As MSComctlLib.Node

    With Node
        If LikeCompBinary(.Key, FindText) Then
            Nodes.Add Node
        End If
        If .Children Then
            Set oNode = .Child
            Do
                pNodeFindKeyCompareBinary oNode, FindText, Nodes
                Set oNode = oNode.Next
            Loop Until oNode Is Nothing
        End If
    End With

End Sub

Private Sub pNodeFindBothCompareBinary(ByVal Node As MSComctlLib.Node, _
                                       ByVal FindText As String, _
                                       ByVal Nodes As Collection)

    Dim oNode As MSComctlLib.Node

    With Node
        If LikeCompBinary(.Text, FindText) Or LikeCompBinary(.Key, FindText) Then
            Nodes.Add Node
        End If
        If .Children Then
            Set oNode = .Child
            Do
                pNodeFindBothCompareBinary oNode, FindText, Nodes
                Set oNode = oNode.Next
            Loop Until oNode Is Nothing
        End If
    End With

End Sub

Private Sub pNodeFindTextCompareText(ByVal Node As MSComctlLib.Node, _
                                     ByVal FindText As String, _
                                     ByVal Nodes As Collection)

    Dim oNode As MSComctlLib.Node

    With Node
        If LikeCompText(.Text, FindText) Then
            Nodes.Add Node
        End If
        If .Children Then
            Set oNode = .Child
            Do
                pNodeFindTextCompareText oNode, FindText, Nodes
                Set oNode = oNode.Next
            Loop Until oNode Is Nothing
        End If
    End With

End Sub

Private Sub pNodeFindKeyCompareText(ByVal Node As MSComctlLib.Node, _
                                    ByVal FindText As String, _
                                    ByVal Nodes As Collection)

    Dim oNode As MSComctlLib.Node

    With Node
        If LikeCompText(.Key, FindText) Then
            Nodes.Add Node
        End If
        If .Children Then
            Set oNode = .Child
            Do
                pNodeFindKeyCompareText oNode, FindText, Nodes
                Set oNode = oNode.Next
            Loop Until oNode Is Nothing
        End If
    End With

End Sub

Private Sub pNodeFindBothCompareText(ByVal Node As MSComctlLib.Node, _
                                     ByVal FindText As String, _
                                     ByVal Nodes As Collection)

    Dim oNode As MSComctlLib.Node

    With Node
        If LikeCompText(.Text, FindText) Or LikeCompText(.Key, FindText) Then
            Nodes.Add Node
        End If
        If .Children Then
            Set oNode = .Child
            Do
                pNodeFindBothCompareText oNode, FindText, Nodes
                Set oNode = oNode.Next
            Loop Until oNode Is Nothing
        End If
    End With

End Sub

Private Sub pNodeFindObjectInTag(ByVal Node As MSComctlLib.Node, _
                                 ByVal FindObject As Object, _
                                 ByVal Nodes As Object, _
                                 ByVal OnlyFirstOccurance As Boolean)

    Dim oNode As MSComctlLib.Node

    With Node
        If IsObject(.Tag) Then
            If .Tag Is FindObject Then
                If OnlyFirstOccurance Then
                    Set Nodes = Node
                    Exit Sub
                End If
            Else
                Nodes.Add Node
            End If
        End If
        If .Children Then
            Set oNode = .Child
            Do
                pNodeFindObjectInTag oNode, FindObject, Nodes, OnlyFirstOccurance
                Set oNode = oNode.Next
            Loop Until oNode Is Nothing
        End If
    End With

End Sub

'===========================================================================
' General: Private subroutines and functions
'
'@@ v02.03.00
Private Sub pClearUpDC(ByVal lIndex As Long)

    With mtMemDC(lIndex)
        If .hBmpOld <> 0 Then
            SelectObject .hdc, .hBmpOld
        End If
        If .hBmp <> 0 Then
            DeleteObject .hBmp
        End If
        If .hdc <> 0 Then
            DeleteDC .hdc
        End If
    End With

End Sub

'@@ v02.03.00
Private Sub pEnsureMemDC(ByVal hDCC As Long, ByVal lIndex As Long, ByVal lWidth As Long, ByVal lHeight As Long, Optional BackColor As OLE_COLOR = vbWindowBackground)

    With mtMemDC(lIndex)
        If lWidth > .lWidth Or lHeight > .lHeight Or .hdc = 0 Then
            pClearUpDC lIndex
            .hdc = CreateCompatibleDC(hDCC)
            If .hdc <> 0 Then
                If hDCC = 0 Then hDCC = .hdc
                .hBmp = CreateCompatibleBitmap(hDCC, lWidth, lHeight)
                If .hBmp <> 0 Then
                    .hBmpOld = SelectObject(.hdc, .hBmp)
                    Dim tR As RECT, hBr As Long
                    tR.Right = lWidth
                    tR.Bottom = lHeight
                    hBr = CreateSolidBrush(pTranslateColor(BackColor))
                    FillRect .hdc, tR, hBr
                    DeleteObject hBr
                Else
                    pClearUpDC lIndex
                End If
            End If
        End If
    End With
   
End Sub

Private Function pGetFirstNode() As MSComctlLib.Node
    With oTree.Nodes
        If .Count Then
            Set pGetFirstNode = .Item(1).root.FirstSibling
        End If
    End With
End Function

Private Function pGetItemFromNode(ByVal Node As MSComctlLib.Node) As Long
    '
    '## If successful, returns the treeview item handle represented by
    '   the specified Node, returns 0 otherwise.
    '
    Dim anSiblingPos() As Integer  '## Contains the sibling position of the node and all it's parents
    Dim iLevel         As Integer  '## Hierarchical level of the node
    Dim hItem          As Long
    Dim iloop          As Integer
    Dim iPos           As Integer

    '
    '## Continually work backwards from the current node to the current node's
    '   first sibling, caching the current node's sibling position in the one-based
    '   array. Then get the first sibling's parent node and start over. Keep going
    '   until the postion of the specified node's top level parent item is obtained...
    '
    Do While (Node Is Nothing) = False
        iLevel = iLevel + 1
        ReDim Preserve anSiblingPos(iLevel)
        anSiblingPos(iLevel) = pGetNodeSiblingPos(Node)
        Set Node = Node.Parent
    Loop

    '## Get the hItem of the first item in the treeview
    hItem = SendMessageAAny(oTree.hwnd, TVM_GETNEXTITEM, ByVal TVGN_ROOT, 0&)
    If hItem Then
        '
        '## Now work backwards through the cached node positions in the array
        '   (from the first treeview node to the specified node), obtaining the respective
        '   item handle for each node at the cached position. When we get to the
        '   specified node's position (the value of the first element in the array), we
        '   got it's hItem...
        '
        For iloop = iLevel To 1 Step -1
            iPos = anSiblingPos(iloop)
            Do While iPos > 1
                hItem = SendMessageAAny(oTree.hwnd, TVM_GETNEXTITEM, ByVal TVGN_NEXT, ByVal hItem)
                iPos = iPos - 1
            Loop
            If (iloop > 1) Then hItem = SendMessageAAny(oTree.hwnd, TVM_GETNEXTITEM, ByVal TVGN_CHILD, ByVal hItem)
        Next
        pGetItemFromNode = hItem
    End If

End Function

'@@ v02.02.00
Private Function pGetLastNode(Optional ByVal Node As MSComctlLib.Node) As MSComctlLib.Node
    '
    '## Recursively retrieves the very last node of the TreeView's expanded branches
    '
    If oTree.Nodes.Count = 0 Then Exit Function                         '## Nothing to process!
    If (Node Is Nothing) Then                                           '## Need somewhere to start!
        Set Node = pGetFirstNode                                        '   Assume all branches instead of a sub-branch
    End If
    If Not (Node = Node.LastSibling) Then Set Node = Node.LastSibling   '## Adjust to last node at same branch level
    If Node.Children And Node.Expanded Then                             '## We have an expanded Parent Node
        Set pGetLastNode = pGetLastNode(Node.Child.LastSibling)         '## Step down to the next branch level & try again
    Else
        Set pGetLastNode = Node                                         '## Found the very last node
    End If

End Function

'@@ v02.02.00
Private Function pGetNextEnabled(ByVal Node As MSComctlLib.Node) As MSComctlLib.Node

    Dim oNode As MSComctlLib.Node

    If (Node Is Nothing) Then Exit Function
    If (Node.Key = pGetLastNode.Key) Then
        If Not moDisNodes.Exist(ObjPtr(Node)) Then
            Set pGetNextEnabled = Node
        End If
        Exit Function
    End If

    If Node = Node.LastSibling Then
        If (Node.Key = pGetLastNode(Node).Key) Then
            Do
                Set Node = Node.Parent
                If Not (Node = Node.LastSibling) Then
                    Set Node = Node.Next
                    If Not moDisNodes.Exist(ObjPtr(Node)) Then
                        Set pGetNextEnabled = Node
                        Exit Do
                    Else
                        Set pGetNextEnabled = pGetNextEnabled(Node)
                        Exit Do
                    End If
                End If
            Loop Until IsRootNode(Node)
        Else
            If Node.Expanded And Node.Children Then
                Set Node = Node.Child.FirstSibling
                If Not moDisNodes.Exist(ObjPtr(Node)) Then
                    Set pGetNextEnabled = Node
                Else
                    Set pGetNextEnabled = pGetNextEnabled(Node)
                End If
            Else
                Set Node = Node.Next
                If Not (Node Is Nothing) Then
                    If Not moDisNodes.Exist(ObjPtr(Node)) Then
                        Set pGetNextEnabled = Node
                    Else
                        Set pGetNextEnabled = pGetNextEnabled(Node)
                    End If
                End If
            End If
        End If
    Else
        If Node.Expanded And Node.Children Then
            Set Node = Node.Child.FirstSibling
            If Not moDisNodes.Exist(ObjPtr(Node)) Then
                Set pGetNextEnabled = Node
            Else
                Set pGetNextEnabled = pGetNextEnabled(Node)
            End If
        Else
            Set Node = Node.Next
            If Not (Node Is Nothing) Then
                If Not moDisNodes.Exist(ObjPtr(Node)) Then
                    Set pGetNextEnabled = Node
                Else
                    Set pGetNextEnabled = pGetNextEnabled(Node)
                End If
            End If
        End If
    End If

End Function

Private Function pGetNodeHandle(Optional ByVal Node As MSComctlLib.Node) As Long
    If (Node Is Nothing) Then
        pGetNodeHandle = SendMessageAny(oTree.hwnd, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
    Else
        pGetNodeHandle = pGetItemFromNode(Node)
    End If
End Function

Private Function pGetNodeSiblingPos(ByVal Node As MSComctlLib.Node) As Integer
    '
    '## Returns the one-base position of the specified node
    '   with respect to it's sibling order.
    '
    Dim nPos As Integer

    '## Keep counting up from one until the node has no more previous siblings
    Do While (Node Is Nothing) = False
        nPos = nPos + 1
        Set Node = Node.Previous
    Loop
    pGetNodeSiblingPos = nPos

End Function

'@@ v02.02.00
Private Function pGetPrevEnabled(ByVal Node As MSComctlLib.Node) As MSComctlLib.Node
    '
    '## Walks recursively backward through the nodes to find the previous 'enabled' node.
    '
    If (Node Is Nothing) Then Exit Function         '## No node passed - error - return nothing

    If Node = Node.FirstSibling Then
        If IsRootNode(Node) Then
            If Node = pGetFirstNode Then
                If Not moDisNodes.Exist(ObjPtr(Node)) Then
                    Set pGetPrevEnabled = Node
                End If
            Else
                Set Node = Node.Previous
                If Node.Children And Node.Expanded Then
                    Set Node = pGetLastNode(Node.Child)
                End If
                If Not moDisNodes.Exist(ObjPtr(Node)) Then
                    Set pGetPrevEnabled = Node
                Else
                    Set pGetPrevEnabled = pGetPrevEnabled(Node)
                End If
            End If
        Else
            Set Node = Node.Parent
            If Not moDisNodes.Exist(ObjPtr(Node)) Then
                Set pGetPrevEnabled = Node
            Else
                Set pGetPrevEnabled = pGetPrevEnabled(Node)
            End If
        End If
    Else
        Set Node = Node.Previous
        If Node.Children And Node.Expanded Then
            Set Node = pGetLastNode(Node.Child)
        End If
        If Not moDisNodes.Exist(ObjPtr(Node)) Then
            Set pGetPrevEnabled = Node
        Else
            Set pGetPrevEnabled = pGetPrevEnabled(Node)
        End If
    End If

End Function

'@@ v02.02.00
Private Function pGetPrevEnabledParent(ByVal Node As MSComctlLib.Node) As MSComctlLib.Node
    '
    '## Walks recursively backward through the Parent nodes to find the previous 'enabled'
    '   parent node. If root node is not enabled, then look back through all root nodes
    '   until all root nodes are exhausted or an enabled root node is found.
    '
    If (Node Is Nothing) Then Exit Function         '## No node passed - error - return nothing

    If IsRootNode(Node) Then
        If Node = pGetFirstNode Then
            If Not moDisNodes.Exist(ObjPtr(Node)) Then
                Set pGetPrevEnabledParent = Node
            End If
        Else
            Set Node = Node.Previous
            If Not moDisNodes.Exist(ObjPtr(Node)) Then
                Set pGetPrevEnabledParent = Node
            Else
                Set pGetPrevEnabledParent = pGetPrevEnabledParent(Node)
            End If
        End If
    Else
        Set Node = Node.Parent
        If Not moDisNodes.Exist(ObjPtr(Node)) Then
            Set pGetPrevEnabledParent = Node
        Else
            Set pGetPrevEnabledParent = pGetPrevEnabledParent(Node)
        End If
    End If

End Function

'@@ v02.01.00
Private Function pGetSelectedNodeCoords() As RECT

    Dim hNode As Long
    Dim rcBox As RECT

    With oTree
        If Not .SelectedItem Is Nothing Then
            '## Node must be fully visible to get coords.
            If .SelectedItem.Visible Then

                '## Get a handle to the currently selected node
                hNode = SendMessageAny(.hwnd, TVM_GETNEXTITEM, TVGN_CARET, ByVal TVGN_ROOT)

                '## Set the first four bytes (Long) of the Rect to the handle of the node, before
                '   calling TVM_GETITEMRECT.
                rcBox.Left = hNode

                '## Get the coordinates of the node
                If SendMessageAny(oTree.hwnd, TVM_GETITEMRECT, True, rcBox) = 0 Then
                    '## Zero the rect to show failure
                    With rcBox
                        .Left = 0
                        .Top = 0
                        .Right = 0
                        .Bottom = 0
                    End With
                End If
            End If
        End If
    End With

    '## Return the rect
    pGetSelectedNodeCoords = rcBox

End Function

Private Function pIsWindowScrollable(hwnd As Long) As ScrollDirectionFlags  '@@ v02.01.00
    '
    '## Returns a set of bit flags indicating whether the specified
    '   window can be scrolled in any given direction.

    Dim si            As SCROLLINFO
    Dim dwScrollFlags As ScrollDirectionFlags

    si.cbSize = Len(si)
    si.fMask = SIF_ALL
    '
    '## Get the horizontal scrollbar's info (GetScrollInfo returns
    '   TRUE after a scrollbar has been added to a window,
    '   even if the respective style bit is not set...)
    '
    If (GetWindowLong(hwnd, GWL_STYLE) And WS_HSCROLL) Then
        If GetScrollInfo(hwnd, SB_HORZ, si) Then
            dwScrollFlags = (sdLeft And (si.nPos > 0))
            dwScrollFlags = dwScrollFlags Or (sdRight And (si.nPos < (((si.nMax - si.nMin) + 1) - si.nPage)))
        End If
    End If
    '
    '## Get the vertical scrollbar's info.
    '
    If (GetWindowLong(hwnd, GWL_STYLE) And WS_VSCROLL) Then
        If GetScrollInfo(hwnd, SB_VERT, si) Then
            dwScrollFlags = dwScrollFlags Or (sdUp And (si.nPos > 0))
            dwScrollFlags = dwScrollFlags Or (sdDown And (si.nPos < (((si.nMax - si.nMin) + 1) - si.nPage)))
        End If
    End If

    pIsWindowScrollable = dwScrollFlags

End Function

Private Function pLoWord(ByVal Value As Long) As Integer                '@@ v02.02.00
    CopyMemory pLoWord, Value, 2
End Function

Private Function pHiWord(ByVal Value As Long) As Integer                '@@ v02.02.00
    CopyMemory pHiWord, ByVal VarPtr(Value) + 2, 2
End Function

'@@ v02.02.00
Private Function pNodeAbsolutePosition(Node As MSComctlLib.Node, ByVal VisibleOnly As Boolean) As Long

    Dim oNode  As MSComctlLib.Node
    Dim lCount As Long

    lCount = Abs(Node.Visible Or Not VisibleOnly)
    If Not (Node Is Node.FirstSibling) Then
        Set oNode = Node.Previous
        Do While Not (oNode Is Nothing)
            lCount = lCount + pNodeCountChildren(oNode, VisibleOnly, False) + Abs(oNode.Visible Or Not VisibleOnly)
            Set oNode = oNode.Previous
        Loop
    End If
    Set oNode = Node.Parent
    If Not (oNode Is Nothing) Then
        lCount = lCount + pNodeAbsolutePosition(oNode, VisibleOnly)
    End If
    pNodeAbsolutePosition = lCount

End Function

'@@ v02.02.00
Private Function pNodeCountChildren(ByVal Node As MSComctlLib.Node, ByVal VisibleOnly As Boolean, ByVal ForceExpanded As Boolean) As Long

    Dim lCount As Long

    If Node.Expanded Or ForceExpanded Then
        Set Node = Node.Child
        Do While Not (Node Is Nothing)
            lCount = lCount + pNodeCountChildren(Node, VisibleOnly, ForceExpanded) + Abs(Node.Visible Or Not VisibleOnly)
            Set Node = Node.Next
        Loop
    End If
    pNodeCountChildren = lCount

End Function

Private Function pNodeSwap(ByVal ParentNode As MSComctlLib.Node, _
                           ByVal ChildNode As MSComctlLib.Node, _
                  Optional ByVal bSelect As Boolean = True) As Boolean    '@@ v02.01.00

    Dim lNDX   As Long
    Dim lCount As Long
    Dim lLoop  As Long

    Redraw False
    With ChildNode
        If IsParentNode(ParentNode, ChildNode) Then '## Are we moving a parent node?
            If IsRootNode(ChildNode) Then           '## Yes. Is it a root node?
                Redraw True
                Exit Function                       '## Yes. Can't move a root node.
            End If
            '## move the children before moving the designated node
            lCount = .Children
            For lLoop = 1 To lCount
                lNDX = .Child.Index
                Set oTree.Nodes(lNDX).Parent = .Parent
            Next
        End If
        '## Force the ParentNode to be expanded before the move
        ParentNode.Expanded = True                  '@@ v01.00.03
        '## Give the child a new parent
        Set .Parent = ParentNode
        If bSelect Then
            .EnsureVisible
            .Selected = bSelect
        End If
    End With
    Redraw True
    pNodeSwap = True

End Function

Private Function pPtInRectRegion(rc As RECT, cxyRegion As Long, pt As POINTAPI) As RectFlags
    '
    '## Returns a set of bit flags indicating whether the specified point resides in
    '   the specified size region with the perimeter of the specified rect. cxyRegion
    '   defines the rectangular region within rc, and must be a positive value

    Dim dwFlags As RectFlags

    If PtInRect(rc, pt.x, pt.y) Then
        dwFlags = (rfLeft And (pt.x <= (rc.Left + cxyRegion)))
        dwFlags = dwFlags Or (rfRight And (pt.x >= (rc.Right - cxyRegion)))
        dwFlags = dwFlags Or (rfTop And (pt.y <= (rc.Top + cxyRegion)))
        dwFlags = dwFlags Or (rfBottom And (pt.y >= (rc.Bottom - cxyRegion)))
    End If

    pPtInRectRegion = dwFlags

End Function

Private Sub pScale(ByVal InX As Single, ByVal InY As Single, _
                    ByRef OutX As Single, ByRef OutY As Single, _
           Optional ByVal InScaleMode As VBRUN.ScaleModeConstants = vbPixels, _
           Optional ByVal OutScaleMode As VBRUN.ScaleModeConstants = vbTwips)
    '
    ' ## Converts from Pixels to nominated scale mode
    '
    With oTree.Parent
        OutX = .ScaleX(InX, InScaleMode, OutScaleMode)
        OutY = .ScaleY(InY, InScaleMode, OutScaleMode)
    End With

End Sub

'Private Sub pScale(ByVal InX As Single, ByVal InY As Single, _
'                   ByRef OutX As Single, ByRef OutY As Single, _
'                   ByVal ScaleMode As VBRUN.ScaleModeConstants)
'    '
'    ' ## Converts from Pixels to nominated scale mode. Use this
'    '    if host doesn't support ScaleX/ScaleY functions
'    '
'    Dim snTwipX As Single
'    Dim snTwipY As Single
'
'    '## Return calculations
'    snTwipX = InX * Screen.TwipsPerPixelX
'    snTwipY = InY * Screen.TwipsPerPixelY
'
'    Select Case ScaleMode
'        Case vbUser
'        Case vbTwips
'            OutX = snTwipX
'            OutY = snTwipY
'        Case vbPoints
'            OutX = snTwipX / 20
'            OutY = snTwipY / 20
'        Case vbPixels
'            OutX = InX
'            OutY = InY
'        Case vbCharacters
'            OutX = snTwipX / 120
'            OutY = snTwipY / 240
'        Case vbInches
'            OutX = snTwipX / 1440
'            OutY = snTwipY / 1440
'        Case vbMillimeters
'            OutX = snTwipX / 56.7
'            OutY = snTwipY / 56.7
'        Case vbCentimeters
'            OutX = snTwipX / 567
'            OutY = snTwipY / 567
'        Case vbHimetric
'            '##      Twip to Inch to Himetric
'            OutX = snTwipX / 1440 / 2540
'            OutY = snTwipY / 1440 / 2540
'        Case vbContainerPosition
'        Case vbContainerSize
'    End Select
'
'End Sub

'@@ v02.02.00
Private Sub pSubNodesExpandState(ByRef Node As MSComctlLib.Node, _
                        Optional ByVal ExpandAll As Boolean)

    Dim oNode As Node

    With Node
        If .Children Then
            .Expanded = ExpandAll
            Set oNode = .Child
            Do
                pSubNodesExpandState oNode, ExpandAll
                Set oNode = oNode.Next
            Loop Until oNode Is Nothing
        End If
    End With

End Sub

'@@ v02.03.00
Private Function pTranslateColor(lColor As Long, Optional ByVal hPal As Long = 0) As Long

    Dim lRColor As Long

    OleTranslateColor lColor, hPal, lRColor
    pTranslateColor = lRColor

End Function

'@@ v02.03.00
Private Sub pTreeViewMessage(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
                             ByVal lParam As Long, ByRef RetVal As Long, ByRef UseRetVal As Boolean)

    Static InProc  As Boolean        '## Prevent recursion with this variable
    Dim ps         As PAINTSTRUCT
    Dim TVDC       As Long
    Dim hDCC       As Long
    Dim trcTV      As RECT
    Dim tRTV       As RECT
    Dim tR         As RECT
    Dim hItemFirst As Long
    Dim hBr        As Long
    Dim lPtr       As Long
    Dim lWidth     As Long
    Dim lHeight    As Long

    Select Case wMsg
        Case WM_PAINT
            If InProc = True Then
                Exit Sub            '## We're busy, come back later!
            End If
            InProc = True
            GetClientRect hwnd, tRTV
            '
            '## Begin painting. This API must be called in response to the WM_PAINT message:
            '
            BeginPaint hwnd, ps

            TVDC = ps.hdc
            LSet trcTV = ps.rcPaint
            '
            '## Create a few canvases in memory to draw on:
            '
            lWidth = trcTV.Right - trcTV.Left
            lHeight = trcTV.Bottom - trcTV.Top
            pEnsureMemDC TVDC, 1, tRTV.Right - tRTV.Left, tRTV.Bottom - tRTV.Top
            pEnsureMemDC TVDC, 2, tRTV.Right - tRTV.Left, tRTV.Bottom - tRTV.Top
            '
            '## This actually causes the TreeView to paint itself onto our memory DC!
            '
            CallOldWindowProc hwnd, WM_PAINT, mtMemDC(1).hdc, 0&
            '
            '## Tile the bitmap or set Backcolor, then draw the TreeView over it transparently
            '
            hItemFirst = SendMessageAny(hwnd, TVM_GETNEXTITEM, 0, ByVal TVGN_ROOT)
            tR.Left = hItemFirst
            SendMessageAny hwnd, TVM_GETITEMRECT, 1, tR
            If tR.Left > mlXOffset Then mlXOffset = tR.Left

            Select Case True
                Case meBackMode = bmGraphic
                    mcTile.XOriginOffset = -tR.Left + mlXOffset
                    mcTile.YOriginOffset = -tR.Top
                    mcTile.TileArea mtMemDC(2).hdc, trcTV.Left, trcTV.Top, lWidth, lHeight

                Case meBackMode = bmColor
                    hBr = CreateSolidBrush(pTranslateColor(mlBackColor))
                    FillRect mtMemDC(2).hdc, trcTV, hBr
                    DeleteObject hBr

            End Select
            '
            '## Create a mono DC of the size to be painted this time:
            '
            pEnsureMemDC 0, 3, lWidth, lHeight
            '
            '## Set BkColor of the TreeView Mem DC to match transparent colour:
            '
            SetBkColor mtMemDC(1).hdc, pTranslateColor(vbWindowBackground)
            '
            '## Copy from TreeView Mem DC -> mono DC:
            '
            BitBlt mtMemDC(3).hdc, 0, 0, lWidth, lHeight, mtMemDC(1).hdc, trcTV.Left, trcTV.Top, vbSrcCopy
            '
            '## Now the DC(3) is a mask.
            '
            SetTextColor mtMemDC(2).hdc, vbBlack
            SetBkColor mtMemDC(2).hdc, vbWhite
            '
            '## AND the mask to the Background so we go white where the treeview is black:
            '
            BitBlt mtMemDC(2).hdc, trcTV.Left, trcTV.Top, lWidth, lHeight, mtMemDC(3).hdc, 0, 0, vbSrcAnd
            '
            '## OR the mask to the TreeView so it goes black where the background is:
            '
            SetTextColor mtMemDC(1).hdc, vbBlack
            SetBkColor mtMemDC(1).hdc, vbWhite
            BitBlt mtMemDC(1).hdc, trcTV.Left, trcTV.Top, lWidth, lHeight, mtMemDC(3).hdc, 0, 0, DSna
            '
            '## Now OR the treeview with the unmasked background:
            '
            BitBlt mtMemDC(2).hdc, trcTV.Left, trcTV.Top, lWidth, lHeight, mtMemDC(1).hdc, trcTV.Left, trcTV.Top, vbSrcPaint
            '
            '## Draw to the target DC
            '
            BitBlt TVDC, trcTV.Left, trcTV.Top, lWidth, lHeight, mtMemDC(2).hdc, trcTV.Left, trcTV.Top, vbSrcCopy
            EndPaint hwnd, ps
            RetVal = 0
            UseRetVal = True
            InProc = False

        Case WM_ERASEBKGND
            'Return TRUE
            RetVal = 1
            UseRetVal = True

        Case WM_DESTROY
            With moSubclass
                .DetachMessage Me, mlHwnd, WM_PAINT
                .DetachMessage Me, mlHwnd, WM_ERASEBKGND
                .DetachMessage Me, mlHWndParent, WM_NOTIFY
                .DetachMessage Me, mlHwnd, WM_DESTROY
            End With

    End Select

End Sub
