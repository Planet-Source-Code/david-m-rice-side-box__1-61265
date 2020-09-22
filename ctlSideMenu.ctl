VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.UserControl ctlSideMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   ScaleHeight     =   2790
   ScaleWidth      =   3600
   ToolboxBitmap   =   "ctlSideMenu.ctx":0000
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   2790
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      IntegralHeight  =   0   'False
      ItemData        =   "ctlSideMenu.ctx":0312
      Left            =   1800
      List            =   "ctlSideMenu.ctx":0314
      TabIndex        =   1
      Top             =   1485
      Visible         =   0   'False
      Width           =   780
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1890
      Top             =   315
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   31
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlSideMenu.ctx":0316
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlSideMenu.ctx":093A
            Key             =   "PointLeft"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlSideMenu.ctx":0E0E
            Key             =   "PointRight"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   0
      ScaleHeight     =   1770
      ScaleWidth      =   570
      TabIndex        =   0
      Top             =   0
      Width           =   570
   End
   Begin VB.Label lblResize 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   945
      TabIndex        =   2
      Top             =   2115
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "ctlSideMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ImageWidth As Long, ImageHeight As Long, X As Long, Y As Long
Dim ListMax As Long, PointHeight As Long, PointWidth As Long
Dim CenterHeight As Long, CurrentLeft As Long, MoveBack As Boolean

Dim m_ShowMenu As Boolean, m_Align As Integer, m_Font As Variant
Dim M_CloseOnDblClick As Boolean, M_Border As Integer

Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event DblClick(List As String, ListIndex As Long)
Event Click(List As String, ListIndex As Long)
Event KeyPress(KeyAscii As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub List1_Click()
    RaiseEvent Click(List1.List(List1.ListIndex), List1.ListIndex)
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_ShowMenu = False Then
        m_ShowMenu = True
    Else
        m_ShowMenu = False
    End If

    DrawControl
End Sub

Private Sub UserControl_Initialize()
    '
    '   Arrow Bitmaps: width=15, height=24
    '   Center Bitmap: width=15, height=31
    '
    PointWidth = Screen.TwipsPerPixelX * 15
    PointHeight = Screen.TwipsPerPixelY * 24
    CenterHeight = Screen.TwipsPerPixelY * 31

    ImageWidth = PointWidth
    ImageHeight = PointHeight + PointHeight + CenterHeight

    List1.Left = 0
    List1.Top = 0
    List1.Height = Picture1.Height
    List1.Width = ImageWidth * 2
    List1.Visible = False
    Picture1.Left = 0
    Picture1.Top = 0
    
    m_ShowMenu = False
    MoveBack = False
    'M_CloseOnDblClick = True
End Sub

Private Sub UserControl_InitProperties()
    '
End Sub

Public Property Get Align() As Integer
Attribute Align.VB_Description = "Returns/sets a value that determines where an object is displayed on a form."
    Align = m_Align
End Property

Public Property Let Align(ByVal New_Align As Integer)
    If New_Align = 0 Or New_Align = 1 Then
        m_Align = New_Align
        PropertyChanged "Align"
        UserControl_Resize
    End If
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Align = PropBag.ReadProperty("Align", 0)
    m_ShowMenu = PropBag.ReadProperty("SnowMenu", False)
    Set List1.Font = PropBag.ReadProperty("Font", m_Font)
    Set lblResize.Font = List1.Font
    Set UserControl.Font = List1.Font
    Set List1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    List1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    List1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    List1.FontSize = PropBag.ReadProperty("FontSize", 12)
    List1.FontName = PropBag.ReadProperty("FontName", Ambient.Font.Name)
    List1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    List1.FontBold = PropBag.ReadProperty("FontBold", 0)
    List1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    List1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    List1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    M_CloseOnDblClick = PropBag.ReadProperty("M_CloseOnDblClick", True)

    DrawControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Align", m_Align, 0)
    Call PropBag.WriteProperty("ShowMenu", m_ShowMenu, False)
    Call PropBag.WriteProperty("Font", m_Font, m_Font)
    Call PropBag.WriteProperty("Font", List1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackColor", List1.BackColor, &H80000005)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FontStrikethru", List1.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontSize", List1.FontSize, 0)
    Call PropBag.WriteProperty("FontName", List1.FontName, Ambient.Font.Name)
    Call PropBag.WriteProperty("FontItalic", List1.FontItalic, 0)
    Call PropBag.WriteProperty("FontBold", List1.FontBold, 0)
    Call PropBag.WriteProperty("FontUnderline", List1.FontUnderline, 0)
    Call PropBag.WriteProperty("ForeColor", List1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("MousePointer", List1.MousePointer, 0)
    Call PropBag.WriteProperty("CloseOnDblClick", M_CloseOnDblClick, True)
End Sub

Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
    List1.AddItem Item, Index
    If List1.ListCount = 1 Then List1.ListIndex = 0
    lblResize.Caption = Item

    If lblResize.Width > ListMax Then
        ListMax = lblResize.Width
        List1.Width = ListMax + (Screen.TwipsPerPixelX * 6) + SysInfo1.ScrollBarSize
    End If
End Sub

Private Sub List1_DblClick()
    RaiseEvent DblClick(List1.List(List1.ListIndex), List1.ListIndex)
    If M_CloseOnDblClick = True Then
        m_ShowMenu = False
        DrawControl
    End If
End Sub

Public Property Set Font(ByVal New_Font As Font)
    Dim i As Long

    Set List1.Font = New_Font
    Set lblResize.Font = New_Font
    Set UserControl.Font = New_Font
    Set m_Font = New_Font

    ListMax = PointWidth * 3
    For i = 0 To List1.ListCount - 1
        lblResize.Caption = List1.List(i)
        If lblResize.Width > ListMax Then
            ListMax = lblResize.Width
        End If
    Next

    PropertyChanged "Font"

End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = List1.Font
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = List1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    List1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Sub Cls()
Attribute Cls.VB_Description = "Clears the contents of a control or the system Clipboard."
    List1.Clear
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get CloseOnDblClick() As Boolean
    CloseOnDblClick = M_CloseOnDblClick
End Property

Public Property Let CloseOnDblClick(ByVal New_CloseOnDblClick As Boolean)
    M_CloseOnDblClick = New_CloseOnDblClick
    PropertyChanged "CloseOnDblClick"
End Property

Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = List1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    List1.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = List1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    List1.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = List1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    List1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = List1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    List1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = List1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    List1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = List1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    List1.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = List1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    List1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = List1.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    List1.ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
End Property

Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = List1.ListCount
End Property

Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
    List = List1.List(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
    List1.List(Index) = New_List
    PropertyChanged "List"
End Property

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = List1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    List1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Public Sub RemoveItem(ByVal Index As Integer)
Attribute RemoveItem.VB_Description = "Removes an item from a ListBox or ComboBox control or a row from a Grid control."
    List1.RemoveItem Index
End Sub

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = List1.Sorted
End Property

Private Sub UserControl_Resize()
    DrawControl
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Sub DrawControl()
    Dim Temp As Long, CenterSpan As Long, i As Integer, X As Long
    Dim PointWhere As String, PointBack As String
    
    If UserControl.Height < (PointHeight * 2) Then UserControl.Height = (PointHeight * 2)
    Picture1.Height = UserControl.Height
    List1.Height = UserControl.Height

    Picture1.Cls
    Picture1.Left = 0
    Picture1.Width = PointWidth

    Temp = (Picture1.Height - PointHeight) - (Screen.TwipsPerPixelX * 2)
    CenterSpan = Picture1.Height - (PointHeight * 2)

    If m_Align = 0 Then
        PointWhere = "PointRight"
        PointBack = "PointLeft"
        List1.Left = Picture1.Width
    Else
        PointWhere = "PointLeft"
        PointBack = "PointRight"
        List1.Left = 0
    End If

    If m_ShowMenu = False Then
        If UserControl.Width <> PointWidth Then UserControl.Width = PointWidth
        Picture1.PaintPicture ImageList1.ListImages(PointWhere).Picture, 0, 0
        
        For i = 0 To (CenterSpan \ CenterHeight)
            Picture1.PaintPicture ImageList1.ListImages("Center").Picture, 0, PointHeight + (i * CenterHeight)
        Next
        
        Picture1.PaintPicture ImageList1.ListImages(PointWhere).Picture, 0, Temp
        List1.Visible = False
        
        If MoveBack = True Then
            UserControl.Extender.Left = CurrentLeft
            MoveBack = False
        End If
    Else
        CurrentLeft = UserControl.Extender.Left
        
        UserControl.Width = List1.Width + PointWidth
        
        If m_Align = 1 Then
            UserControl.Extender.Left = CurrentLeft - List1.Width
            MoveBack = True
            Picture1.Left = List1.Width
        End If

        
        Picture1.PaintPicture ImageList1.ListImages(PointBack).Picture, 0, 0
        
        For i = 0 To (CenterSpan \ CenterHeight)
            Picture1.PaintPicture ImageList1.ListImages("Center").Picture, 0, PointHeight + (i * CenterHeight)
        Next
        
        Picture1.PaintPicture ImageList1.ListImages(PointBack).Picture, 0, Temp
        List1.Visible = True
    End If
End Sub
