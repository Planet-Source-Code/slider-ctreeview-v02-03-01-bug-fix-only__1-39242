VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fDragNDrop2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tutorial 2b : Drag and Drop - Same Node Level / By Parent Type"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fDragNDrop2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDialog 
      Caption         =   "Drag'n'Drop Settings:"
      Height          =   2010
      Left            =   3465
      TabIndex        =   1
      Top             =   105
      Width           =   5790
      Begin VB.ComboBox cboDialog 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1470
         Width           =   2220
      End
      Begin VB.TextBox txtDialog 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   3360
         TabIndex        =   4
         Top             =   315
         Width           =   960
      End
      Begin VB.TextBox txtDialog 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   7
         Top             =   735
         Width           =   960
      End
      Begin VB.CheckBox chkDialog 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Allow Drag'n'Drop"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1995
         TabIndex        =   8
         Top             =   1155
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Node Drop Mode:"
         Height          =   315
         Index           =   4
         Left            =   1680
         TabIndex        =   9
         Top             =   1470
         Width           =   1590
      End
      Begin VB.Label lblDialog 
         Caption         =   "milliseconds"
         Height          =   210
         Index           =   3
         Left            =   4410
         TabIndex        =   6
         Top             =   780
         Width           =   855
      End
      Begin VB.Label lblDialog 
         Caption         =   "milliseconds"
         Height          =   210
         Index           =   2
         Left            =   4410
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Auto-Expand Timer:"
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Top             =   735
         Width           =   1590
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Auto-Scroll Timer:"
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   2
         Top             =   315
         Width           =   1590
      End
   End
   Begin VB.ListBox lstEvents 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   3465
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2625
      Width           =   5790
   End
   Begin MSComctlLib.TreeView tvwDialog 
      Height          =   6630
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   11695
      _Version        =   393217
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilDialog 
      Left            =   5985
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483628
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":27A2
            Key             =   "Closed1"
            Object.Tag             =   "Closed Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":2D3C
            Key             =   "Open1"
            Object.Tag             =   "Open Folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":32D6
            Key             =   "Selected"
            Object.Tag             =   "Selected Folder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":3870
            Key             =   "Group"
            Object.Tag             =   "Group Folder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":3E0A
            Key             =   "Closed2"
            Object.Tag             =   "Closed Network Folder"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":43A4
            Key             =   "Open2"
            Object.Tag             =   "Open Network Folder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":493E
            Key             =   "Clock"
            Object.Tag             =   "Clock"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":4D90
            Key             =   "Barcode"
            Object.Tag             =   "Barcode"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":4EEA
            Key             =   "Agent"
            Object.Tag             =   "Agent"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":7024
            Key             =   "Diary"
            Object.Tag             =   "Diary"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":CC46
            Key             =   "Item"
            Object.Tag             =   "Card Item"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":CF60
            Key             =   "ShareOverlay"
            Object.Tag             =   "Share Overlay"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":D4FA
            Key             =   "ShortcutOverlay"
            Object.Tag             =   "Shortcut Overlay"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":D654
            Key             =   "Custom1Overlay"
            Object.Tag             =   "Custom Overlay 1"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":D96E
            Key             =   "Custom2Overlay"
            Object.Tag             =   "Custom Overlay 2"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fDragNDrop2.frx":DAC8
            Key             =   "Custom3Overlay"
            Object.Tag             =   "Custom Overlay 3"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblEvents 
      Caption         =   "Drag and Drop Events:"
      Height          =   225
      Left            =   3465
      TabIndex        =   11
      Top             =   2310
      Width           =   2430
   End
   Begin VB.Menu mnuPopNode 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuTest 
         Caption         =   "Test Popup"
      End
   End
End
Attribute VB_Name = "fDragNDrop2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fDragDrop2 [Tutorial 2b]
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         1/05/2002
' Version:      01.01.00
' Description:  Application showing the Drag and Drop feature of vbTree DLL
'               (cTreeView class).
' Edit History: 01.00.00 23/03/2002 Initial Release
'               01.01.00 01/05/2002 Broke Tutorial 2 into two tutorials.
'                                   This tutorial illustrates the following:-
'                                   ... Overriding the defalt drag cursor
'                                   ... Tie nodes to their own branch level
'                                   ... Modify behaviour of dropping a node
'               01.01.01 23/09/2002 Fixed flickering problem - Thanks to
'                                   Denis Chollet for pointing this out.
'
'===========================================================================

Option Explicit

Private Enum eProperty
    [Auto Scroll] = 0
    [Auto Expand] = 1
End Enum

#Const OLDDATA = 0      '## 1 = Old Test Data

#If NODLL = 0 Then
    Private WithEvents moTree As vbTree.cTreeView
Attribute moTree.VB_VarHelpID = -1
#Else
    Private WithEvents moTree As cTreeView
Attribute moTree.VB_VarHelpID = -1
#End If

Private moDestNode     As MSComctlLib.Node
Private mbOkToDrop     As Boolean
Private meRelationship As TreeRelationshipConstants

Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" _
                                        (ByVal hwnd As Long, _
                                         ByVal wMsg As Long, _
                                         ByVal wParam As Long, _
                                               lParam As Any) As Long

Private Const WM_VSCROLL As Long = &H115
Private Const SB_BOTTOM  As Long = 7

'===========================================================================
' Form Events
'
Private Sub Form_KeyPress(KeyAscii As Integer)
    '
    '## Make Return/Enter key act like the Tab key...
    '
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys "{TAB}"
    End Select

End Sub

Private Sub Form_Load()

#If NODLL = 0 Then
    Set moTree = New vbTree.cTreeView   '## Used to manage the TreeView
#Else
    Set moTree = New cTreeView          '## Used to manage the TreeView
#End If

    With moTree
        '
        '## Hook treeview control
        '
        .HookCtrl tvwDialog, [Single Select]
        '
        '## Set TreeView features
        '
        With .Ctrl
            .Style = tvwTreelinesPlusMinusPictureText
            .LineStyle = tvwRootLines
            .Indentation = 10
            .ImageList = ilDialog
            .FullRowSelect = False
            .HideSelection = False
            .HotTracking = True
            '
            '## Build TreeView data
            '
            .Visible = False
            pInitData
            .Visible = True
        End With
        .BackMode = bmColor 'bmGraphic
        .BackColor = vbInfoBackground
        '.BackFileName = "C:\cTreeview v020300\Tutorial 01 - Basic Functions\Sample Backgrounds\Can.gif"
        '
        '## Show focus rectangle over first node but don't select
        '
        .SetFocusNode (1)
        '
        '## Load Drag timer default settigs
        '
        txtDialog([Auto Scroll]).Text = CStr(.DragScrollTime)
        txtDialog([Auto Expand]).Text = CStr(.DragExpandTime)
        '
        '## Show default setting"
        '
        pShowEvent "** Auto Scroll Timer default setting is " + CStr(.DragScrollTime) + "ms"
        pShowEvent "** Auto Expand Timer default setting is " + CStr(.DragExpandTime) + "ms"
        pShowEvent "** Drag'n'Drop Operation is enabled"
        pShowEvent "-----------------------------------------"
        .DragEnabled = True
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moTree = Nothing
End Sub

'===========================================================================
' Form Control Events
'
Private Sub cboDialog_Click()
    With cboDialog
        meRelationship = .ItemData(.ListIndex)
        pShowEvent "** Node drop mode set to '" + .List(.ListIndex) + "'"
    End With
End Sub

Private Sub chkDialog_Click()
    With moTree
        .DragEnabled = CBool(chkDialog.Value = vbChecked)
        If .DragEnabled Then
            pShowEvent "** Drag'n'Drop Operation is enabled"
        Else
            pShowEvent "** Drag'n'Drop Operation is disabled"
        End If
    End With
End Sub

Private Sub lstEvents_GotFocus()
    tvwDialog.SetFocus          '## We don't want the item to be selected
End Sub

Private Sub lstEvents_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lstEvents.ListIndex = -1    '## We don't want the item to be selected
End Sub

Private Sub txtDialog_GotFocus(Index As Integer)
    pHiLite Index
End Sub

Private Sub txtDialog_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
            '## Valid keystokes (numeric only) - Do nothing
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtDialog_Validate(Index As Integer, Cancel As Boolean)
    '
    '## Timer value changed
    '
    Dim lValue As Long

    lValue = CLng(txtDialog(Index).Text)
    With moTree
        Select Case Index
            Case [Auto Scroll]
                If Not (.DragScrollTime = lValue) Then
                    .DragScrollTime = lValue
                    pShowEvent "** Auto Scroll Timer now set to " + CStr(.DragScrollTime) + "ms"
                End If
            Case [Auto Expand]
                If Not (.DragExpandTime = lValue) Then
                    .DragExpandTime = lValue
                    pShowEvent "** Auto Expand Timer now set to " + CStr(.DragExpandTime) + "ms"
                End If
        End Select
    End With
End Sub

'===========================================================================
' cTreeView Class Events
'
Private Sub moTree_StartDrag(SourceNode As MSComctlLib.Node)
    '
    '## We've started dragging a node
    '
    pShowEvent "Start Drag Node = '" + SourceNode.Text + "'"

End Sub

Private Sub moTree_Dragging(SourceNode As MSComctlLib.Node, TargetParent As MSComctlLib.Node)
    '
    '## Node being dragged
    '
    Dim sTag As String
    Dim lTLvl As Long
    Dim lSLvl As Long

    lTLvl = moTree.NodeNestingLevel(TargetParent)
    lSLvl = moTree.NodeNestingLevel(SourceNode)
    If (lTLvl < lSLvl - 1) Or (lTLvl > lSLvl) Then
        '## no drag - invalid node level!
        tvwDialog.DragIcon = LoadPicture(App.Path + "\NoSelect.CUR")
        mbOkToDrop = False
    Else
        mbOkToDrop = True
    End If
    Select Case True
        Case (moDestNode Is Nothing), moDestNode <> TargetParent
            Set moDestNode = TargetParent
            pShowEvent "Dragging '" + SourceNode.Text + "' over '" + TargetParent.Text + "'"
    End Select

End Sub

Private Sub moTree_Dropped(SourceNode As MSComctlLib.Node, TargetParent As MSComctlLib.Node)

    Dim eRelationship As TreeRelationshipConstants
    Dim oMarker       As MSComctlLib.Node

    '
    '## Node has been dropped - Now what to do with it...
    '
    Debug.Print "++ Dropped Node = '"; SourceNode.Text; "'"
    '
    '## Move the dragged node
    '
    If mbOkToDrop Then
        If moTree.NodeNestingLevel(TargetParent) = moTree.NodeNestingLevel(SourceNode) Then
            '## Same branch level
            Set oMarker = TargetParent
        Else
            '## Dropped on parent level - get first valid child
            Set oMarker = TargetParent.Child.FirstSibling
            If oMarker.Index = SourceNode.Index Then
                '## can't set drop target to self, select next available
                Set oMarker = oMarker.Next
            End If
        End If
        pShowEvent "Dropped '" + SourceNode.Text + "' on '" + TargetParent.Text + "'"
        tvwDialog.Visible = False                                       '@@ 01.01.01
        If Not moTree.NodeMove(oMarker, SourceNode, , meRelationship) Then
            '
            '## Problems with moving the node. Most likely a root node was dragged!
            '
            MsgBox "Unable to move the selected node.", _
                   vbApplicationModal + vbExclamation + vbOKOnly, _
                   App.Title
        End If
        tvwDialog.Visible = True                                        '@@ 01.01.01
    Else
        '## Invalid node...
        pShowEvent "!! Drag disallowed !!"
    End If
    pShowEvent "-----------------------------------------"

End Sub

'===========================================================================
' Internal Functions
'
Private Sub pHiLite(Index As Integer)
    With txtDialog(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub pInitData()

#If OLDDATA = 1 Then
    With moTree
        .NodeAdd , , "A", "Basic Functions", 1, 3, , , , True, , True, , , , 2
        .NodeAdd , , "B", "Drag and Drop", 1, 3, , , , , , , , , , 2
        .NodeAdd , , "C", "MultiSelection", 1, 3, , , , , , , , , , 2
        .NodeAdd , , "D", "Load On Demand", 1, 3, , , , , , , , , , 2
        .NodeAdd , , "E", "ADO Integration", 1, 3, , , , , , , , , , 2

        Dim lLoop As Long
        .NodeAdd , , "X1", "Node Item 1", 1, 3, , , , , , , , , , 2
        For lLoop = 2 To 50
            .NodeAdd tvwDialog.Nodes("X" + CStr(lLoop - 1)), tvwChild, "X" + CStr(lLoop), "Node Item " + CStr(lLoop), 1, 3, , , , , , , , , , 2
        Next
    End With
#Else
    Dim lLoop1 As Long
    Dim lLoop2 As Long
    Dim lLoop3 As Long
    Dim lLoop4 As Long
    Dim sChar1 As String * 1
    Dim sChar2 As String * 2
    Dim sChar3 As String * 3
    Dim sText  As String

    With moTree
        For lLoop1 = 65 To 70
            sChar1 = Chr$(lLoop1)
            .NodeAdd , , sChar1, "Folder " + sChar1, 5, 6, , , , , , , , , , 2
            For lLoop2 = 1 To 5
                sChar2 = sChar1 + CStr(lLoop2)
                .NodeAdd tvwDialog.Nodes(sChar1), tvwChild, sChar2, "Sub Folder " + sChar2, 1, 3, , , , , , , , &HFF&, , 2
                For lLoop3 = 1 To 3
                    sChar3 = sChar2 + CStr(lLoop3)
                    sText = "Book ID " + CStr(lLoop3)
                    .NodeAdd tvwDialog.Nodes(sChar2), tvwChild, sChar3, sText, 10, 8, , , , , , , , &HFF0000, , 10
                    For lLoop4 = 1 To 3
                        .NodeAdd tvwDialog.Nodes(sChar3), tvwChild, sChar3 + "-" + CStr(lLoop4), "Chapter " + CStr(lLoop4) + "[" + sText + "]", 11, 9, , , , , , , , &H800080, , 11
                    Next
                Next
            Next
        Next
    End With
#End If

    '
    '## Fill Combobox with TreeRelationshipConstants
    '
    With cboDialog
        .AddItem "First child node"
        .ItemData(.NewIndex) = tvwFirst
        .AddItem "Before highlighted node"
        .ItemData(.NewIndex) = tvwPrevious
        .AddItem "After highlighted node"
        .ItemData(.NewIndex) = tvwNext
        .AddItem "Last child node"
        .ItemData(.NewIndex) = tvwLast
        .ListIndex = 1
    End With

End Sub

Private Sub pShowEvent(sText As String)
    With lstEvents
        .AddItem sText
        If .ListCount > 200 Then
            '## Don't let the list get too long (maimum of 200 entries) ...
            .RemoveItem (0)
        End If
        SendMessageAny .hwnd, WM_VSCROLL, SB_BOTTOM, vbNull '## Scroll to bottom
    End With
End Sub
