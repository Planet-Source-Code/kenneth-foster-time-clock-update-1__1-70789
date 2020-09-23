VERSION 5.00
Begin VB.UserControl StrokeText 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   1245
   ScaleWidth      =   6360
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   2355
      Width           =   480
   End
End
Attribute VB_Name = "StrokeText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'by Ken Foster

Enum eeStyle
   Stroke = 0
   Fill = 1
   StrokeandFill = 2
End Enum

Const m_def_ColorFill = vbRed
Const m_def_ColorStroke = vbBlack
Const m_def_Style = 1
Const m_def_StrokeWidth = 1
Const m_def_Shadow = False

Private m_ColorStroke As OLE_COLOR
Private m_ColorFill As OLE_COLOR
Private m_Caption As String
Private m_Style As eeStyle
Private m_StrokeWidth As Long
Private m_Shadow As Boolean

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long

Dim hBrush As Long, oldBrush As Long
Dim sText As String

Private Sub DrawText()
    'KPD-Team 2000
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
   
    UserControl.Cls
    UserControl.ScaleMode = vbPixels
    sText = m_Caption
    
     If Shadow = True Then
       hBrush = CreateSolidBrush(&HC0C0C0)
       oldBrush = SelectObject(UserControl.hdc, hBrush)
       UserControl.ForeColor = &HC0C0C0
       BeginPath UserControl.hdc
       TextOut UserControl.hdc, 5, 4, sText, Len(sText)
       EndPath UserControl.hdc
       If TStyle = 0 Then
          StrokePath UserControl.hdc
       Else
          StrokeAndFillPath UserControl.hdc
       End If
       
        hBrush = CreateSolidBrush(&H808080)
       oldBrush = SelectObject(UserControl.hdc, hBrush)
       UserControl.ForeColor = &H808080
       BeginPath UserControl.hdc
       TextOut UserControl.hdc, 5, 3, sText, Len(sText)
       EndPath UserControl.hdc
       If TStyle = 0 Then
          StrokePath UserControl.hdc
       Else
          StrokeAndFillPath UserControl.hdc
       End If
       
       hBrush = CreateSolidBrush(vbBlack)
       'replace the current brush with the new  brush
       oldBrush = SelectObject(UserControl.hdc, hBrush)
       'set the fore color to black
       UserControl.ForeColor = vbBlack
       'begin a new path
       BeginPath UserControl.hdc
       TextOut UserControl.hdc, 5, 2, sText, Len(sText)
       EndPath UserControl.hdc
       If TStyle = 0 Then
          StrokePath UserControl.hdc
       Else
          StrokeAndFillPath UserControl.hdc
       End If
     End If
   
    'create a new brush
    hBrush = CreateSolidBrush(m_ColorFill)
    'replace the current brush with the new white brush
    oldBrush = SelectObject(UserControl.hdc, hBrush)
    'set the fore color
    UserControl.ForeColor = m_ColorStroke
    'begin a new path
    BeginPath UserControl.hdc
    TextOut UserControl.hdc, 3, 0, sText, Len(sText)
    EndPath UserControl.hdc
    
    If TStyle = 0 Then StrokePath UserControl.hdc
    If TStyle = 1 Then FillPath UserControl.hdc
    If TStyle = 2 Then StrokeAndFillPath UserControl.hdc
    
    'replace this form's brush with the original one
    SelectObject UserControl.hdc, oldBrush
    'delete  brush
    DeleteObject hBrush

    UserControl.MaskPicture = UserControl.Image  ' this line of code is important, makes uc transparent
End Sub

Private Sub UserControl_Initialize()
   m_ColorStroke = m_def_ColorStroke
   m_ColorFill = m_def_ColorFill
   m_StrokeWidth = m_def_StrokeWidth
   m_Shadow = m_def_Shadow
   UserControl.FontName = "Tahoma"
   UserControl.FontSize = 48
   UserControl.DrawWidth = StrokeWidth

End Sub

Private Sub UserControl_InitProperties()
     Caption = Extender.Name                                 'assigns Caption name of usercontrol
     ColorFill = vbRed
     ColorStroke = vbBlack
     Shadow = False
End Sub

Private Sub UserControl_Resize()
   UserControl.ScaleMode = 1
   With Label1
       UserControl.ScaleMode = 1
       .Caption = m_Caption
       .FontName = UserControl.FontName
       .FontSize = UserControl.FontSize
       UserControl.Width = .Width + 200
       UserControl.Height = .Height
    End With
    
   DrawText
End Sub

Public Property Get Caption() As String
     Caption = m_Caption
End Property

Public Property Let Caption(NewCaption As String)
     m_Caption = NewCaption
     PropertyChanged "Caption"
     DrawText
End Property

Public Property Get Font() As Font
     Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal NewFont As Font)
     Set UserControl.Font = NewFont
     PropertyChanged "Font"
     DrawText
End Property

Public Property Get ColorStroke() As OLE_COLOR
     ColorStroke = m_ColorStroke
End Property

Public Property Let ColorStroke(ByVal NewColorStroke As OLE_COLOR)
     m_ColorStroke = NewColorStroke
     PropertyChanged "ColorStroke"
     DrawText
End Property

Public Property Get ColorFill() As OLE_COLOR
     ColorFill = m_ColorFill
End Property

Public Property Let ColorFill(ByVal NewColorFill As OLE_COLOR)
     m_ColorFill = NewColorFill
     PropertyChanged "ColorFill"
     DrawText
End Property

Public Property Get StrokeWidth() As Long
   StrokeWidth = m_StrokeWidth
End Property

Public Property Let StrokeWidth(ByVal newStrokeWidth As Long)
   m_StrokeWidth = newStrokeWidth
   UserControl.DrawWidth = m_StrokeWidth
   PropertyChanged "StrokeWidth"
   DrawText
End Property

Public Property Get Shadow() As Boolean
   Shadow = m_Shadow
End Property

Public Property Let Shadow(ByVal newShadow As Boolean)
   m_Shadow = newShadow
   PropertyChanged "Shadow"
   DrawText
End Property

Public Property Get TStyle() As eeStyle
   TStyle = m_Style
End Property

Public Property Let TStyle(ByVal NewStyle As eeStyle)
   m_Style = NewStyle
   PropertyChanged "TStyle"
   DrawText
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     With PropBag
          Caption = .ReadProperty("Caption", Extender.Name)
          ColorStroke = .ReadProperty("ColorStroke", m_def_ColorStroke)
          ColorFill = .ReadProperty("ColorFill", m_def_ColorFill)
          StrokeWidth = .ReadProperty("StrokeWidth", m_def_StrokeWidth)
          Shadow = .ReadProperty("Shadow", m_def_Shadow)
          TStyle = .ReadProperty("TStyle", m_def_Style)
          Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
     End With
     
     DrawText
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     With PropBag
          Call .WriteProperty("Caption", m_Caption, Extender.Name)
          Call .WriteProperty("ColorStroke", m_ColorStroke, m_def_ColorStroke)
          Call .WriteProperty("ColorFill", m_ColorFill, m_def_ColorFill)
          Call .WriteProperty("StrokeWidth", m_StrokeWidth, m_def_StrokeWidth)
          Call .WriteProperty("Shadow", m_Shadow, m_def_Shadow)
          Call .WriteProperty("TStyle", m_Style, m_def_Style)
          Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
     End With
End Sub
