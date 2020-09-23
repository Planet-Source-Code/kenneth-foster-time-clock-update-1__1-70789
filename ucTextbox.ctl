VERSION 5.00
Begin VB.UserControl ucTextbox 
   BackColor       =   &H007AAC4C&
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   ScaleHeight     =   600
   ScaleWidth      =   5970
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   2025
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   45
      Width           =   3315
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   5385
      Picture         =   "ucTextbox.ctx":0000
      Top             =   60
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   1815
      Picture         =   "ucTextbox.ctx":014A
      Stretch         =   -1  'True
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "ucText"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   75
      Width           =   600
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5715
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   90
      TabIndex        =   0
      Top             =   285
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "ucTextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'by Ken Foster

Const m_def_Caption = "ucTextbox"
Const m_def_BorderColor = vbBlack
Const m_def_TextBackColor = vbWhite
Const m_def_CaptionColor = vbBlack
Const m_def_Text = ""
Const m_def_MaxLength = 0
Const m_def_StretchDivider = False
Const m_def_PointerLeft = True
Const m_def_PointerRight = False
Const m_def_BoldText = False
Const m_def_PassWordChar = ""

Dim m_PassWordChar As String
Dim m_Caption As String
Dim m_BorderColor As OLE_COLOR
Dim m_TextBackColor As OLE_COLOR
Dim m_CaptionColor As OLE_COLOR
Dim m_Text As String
Dim m_MaxLength As Integer
Dim m_StretchDivider As Boolean
Dim m_PointerLeft As Boolean
Dim m_PointerRight As Boolean
Dim m_BoldText As Boolean

Event Change()
Event Click()
Event DblClick()
Event KeyPress(KeyAscii As Integer)
Event KeyDown()
Event KeyUp()

Private Sub Text1_Change()
   Text = Text1.Text
   RaiseEvent Change
End Sub

Private Sub Text1_Click()
   RaiseEvent Click
   Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Text1_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub Text1_GotFocus()
   Dim TxtLen As Integer
   'put carot at end of text
   TxtLen = Len(Text1.Text)
   Text1.SelStart = TxtLen
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp
End Sub

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
   m_Caption = m_def_Caption
   m_BorderColor = m_def_BorderColor
   m_TextBackColor = m_def_TextBackColor
   m_CaptionColor = m_def_CaptionColor
   m_Text = m_def_Text
   m_MaxLength = m_def_MaxLength
   m_StretchDivider = m_def_StretchDivider
   m_PointerLeft = m_def_PointerLeft
   m_PointerRight = m_def_PointerRight
   m_BoldText = m_def_BoldText
   m_PassWordChar = m_def_PassWordChar
End Sub

Private Sub UserControl_InitProperties()
   Caption = Extender.Name
   BorderColor = m_BorderColor
   TextBackColor = m_TextBackColor
   CaptionColor = m_CaptionColor
   Text = m_Text
   MaxLength = m_MaxLength
   StretchDivider = m_StretchDivider
   PointerLeft = m_PointerLeft
   PointerRight = m_PointerRight
   BoldText = m_BoldText
   PassWordChar = m_PassWordChar
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp
End Sub

Private Sub UserControl_Resize()
   Label2.Caption = Caption   'presizes label1 width
   Label1.Left = 20
   Label1.Width = Label2.Width
   Label1.Caption = Label2.Caption
   'position and size all the components
   If StretchDivider = True Then
      Image1.Top = -50
      Image1.Width = 210
      Image1.Height = 465
      Image2.Top = -50
      Image2.Width = 210
      Image2.Height = 465
   Else
      Image1.Top = 65
      Image1.Width = 240
      Image1.Height = 240
      Image2.Top = 65
      Image2.Width = 240
      Image2.Height = 240
   End If
   Image1.Left = Label1.Width + Label1.Left
   Shape1.Left = 0
   Shape1.Width = UserControl.Width
   Text1.Top = 80
   Text1.Left = Image1.Width + Image1.Left + 100
   Text1.Width = UserControl.Width - Label1.Width - Image2.Width - 285
   Image2.Left = UserControl.Width - Image2.Width
   Text1.Height = Shape1.Height - 100
   UserControl.Height = Shape1.Height
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Text to left of divider."
   Caption = m_Caption
End Property

Public Property Let Caption(NewCaption As String)
   m_Caption = NewCaption
   Label2.Caption = m_Caption
   PropertyChanged "Caption"
   UserControl_Resize
End Property

Public Property Get BorderColor() As OLE_COLOR
   BorderColor = m_BorderColor
End Property

Public Property Get CaptionColor() As OLE_COLOR
   CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(NewCaptionColor As OLE_COLOR)
   m_CaptionColor = NewCaptionColor
   Label1.ForeColor = m_CaptionColor
   Text1.ForeColor = m_CaptionColor
   PropertyChanged "CaptionColor"
   UserControl_Resize
End Property

Public Property Let BorderColor(NewBorderColor As OLE_COLOR)
   m_BorderColor = NewBorderColor
   Shape1.BorderColor = BorderColor
   PropertyChanged "BorderColor"
   UserControl_Resize
End Property

Public Property Get BoldText() As Boolean
   BoldText = m_BoldText
End Property

Public Property Let BoldText(NewBoldText As Boolean)
   m_BoldText = NewBoldText
   Text1.FontBold = m_BoldText
   PropertyChanged "BoldText"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Text to right of divider."
   Text = m_Text
End Property

Public Property Let Text(NewText As String)
   m_Text = NewText
   Text1.Text = m_Text
   PropertyChanged "Text"
End Property

Public Property Get TextBackColor() As OLE_COLOR
   TextBackColor = m_TextBackColor
End Property

Public Property Let TextBackColor(NewTextBackColor As OLE_COLOR)
   m_TextBackColor = NewTextBackColor
   Text1.BackColor = m_TextBackColor
   Shape1.FillColor = m_TextBackColor
   Label1.BackColor = m_TextBackColor
   PropertyChanged "TextBackColor"
   UserControl_Resize
End Property

Public Property Get MaxLength() As Integer
Attribute MaxLength.VB_Description = "Maximum number of charactors that can be typed in. 0 = no limit."
   MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(NewMaxLength As Integer)
   m_MaxLength = NewMaxLength
   Text1.MaxLength = m_MaxLength
   PropertyChanged "MaxLength"
End Property

Public Property Get StretchDivider() As Boolean
Attribute StretchDivider.VB_Description = "Makes pointers taller if true."
   StretchDivider = m_StretchDivider
End Property

Public Property Let StretchDivider(NewStretchDivider As Boolean)
   m_StretchDivider = NewStretchDivider
   Image1.Stretch = m_StretchDivider
   Image2.Stretch = m_StretchDivider
   PropertyChanged "StretchDivider"
   UserControl_Resize
End Property

Public Property Get PointerLeft() As Boolean
   PointerLeft = m_PointerLeft
End Property

Public Property Let PointerLeft(NewPointerLeft As Boolean)
   m_PointerLeft = NewPointerLeft
   Image1.Visible = m_PointerLeft
   PropertyChanged "PointerLeft"
End Property

Public Property Get PointerRight() As Boolean
Attribute PointerRight.VB_Description = "Right hand divider. Set MaxLength to appropriate number of charactors."
   PointerRight = m_PointerRight
End Property

Public Property Let PointerRight(NewPointerRight As Boolean)
   m_PointerRight = NewPointerRight
   Image2.Visible = m_PointerRight
   PropertyChanged "PointerRight"
End Property

Public Property Get PassWordChar() As String
   Let PassWordChar = m_PassWordChar
End Property

Public Property Let PassWordChar(ByVal NewPassWordChar As String)
   Let m_PassWordChar = NewPassWordChar
   Text1.PassWordChar = m_PassWordChar
   PropertyChanged "PassWordChar"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Caption = PropBag.ReadProperty("Caption", m_def_Caption)
   BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
   TextBackColor = PropBag.ReadProperty("TextBackColor", m_def_TextBackColor)
   CaptionColor = PropBag.ReadProperty("CaptionColor", m_def_CaptionColor)
   Text = PropBag.ReadProperty("Text", m_def_Text)
   MaxLength = PropBag.ReadProperty("MaxLength", m_def_MaxLength)
   StretchDivider = PropBag.ReadProperty("StretchDivider", m_def_StretchDivider)
   PointerLeft = PropBag.ReadProperty("PointerLeft", m_def_PointerLeft)
   PointerRight = PropBag.ReadProperty("PointerRight", m_def_PointerRight)
   BoldText = PropBag.ReadProperty("BoldText", m_def_BoldText)
   PassWordChar = PropBag.ReadProperty("PassWordChar", m_def_PassWordChar)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
   Call .WriteProperty("Caption", m_Caption, m_def_Caption)
   Call .WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
   Call .WriteProperty("TextBackColor", m_TextBackColor, m_def_TextBackColor)
   Call .WriteProperty("CaptionColor", m_CaptionColor, m_def_CaptionColor)
   Call .WriteProperty("Text", m_Text, m_def_Text)
   Call .WriteProperty("MaxLength", m_MaxLength, m_def_MaxLength)
   Call .WriteProperty("StretchDivider", m_StretchDivider, m_def_StretchDivider)
   Call .WriteProperty("PointerLeft", m_PointerLeft, m_def_PointerLeft)
   Call .WriteProperty("PointerRight", m_PointerRight, m_def_PointerRight)
   Call .WriteProperty("BoldText", m_BoldText, m_def_BoldText)
   Call .WriteProperty("PassWordChar", m_PassWordChar, m_def_PassWordChar)
   End With
End Sub
