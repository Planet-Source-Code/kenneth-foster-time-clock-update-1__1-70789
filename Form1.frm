VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00CEE0BC&
   Caption         =   "                                                                        Time Clock Demo"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ccXPButton cmdExit 
      Default         =   -1  'True
      Height          =   435
      Left            =   9375
      TabIndex        =   19
      Top             =   6330
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   767
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.StrokeText StrokeText2 
      Height          =   345
      Left            =   3525
      TabIndex        =   11
      Top             =   6225
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
      Caption         =   "by Ken Foster"
      ColorFill       =   8039500
      TStyle          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   11085
      TabIndex        =   10
      Top             =   225
      Visible         =   0   'False
      Width           =   1470
   End
   Begin Project1.PanelFx PanelFx2 
      Height          =   2220
      Left            =   5745
      TabIndex        =   4
      Top             =   2880
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3916
      TileHeight      =   25
      TitleCaption    =   "New Employee Info / Change Pin"
      TitleForeColor  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleAlignment  =   1
      RoundEdge       =   15
      CanCollapse     =   -1  'True
      BackGroundStyle =   1
      gCTitleStart    =   8039500
      gCTitleEnd      =   12769963
      gCPanelStart    =   8039500
      gCPanelEnd      =   12769963
      Begin Project1.ccXPButton cmdSaveInfo 
         Height          =   345
         Left            =   2985
         TabIndex        =   9
         Top             =   780
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         Caption         =   "SAVE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.ucTextbox txtNewEmpPin 
         Height          =   345
         Left            =   285
         TabIndex        =   6
         Top             =   780
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   609
         Caption         =   " Pin"
         MaxLength       =   6
         BoldText        =   -1  'True
      End
      Begin Project1.ucTextbox txtNewEmpName 
         Height          =   345
         Left            =   300
         TabIndex        =   5
         Top             =   405
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   609
         Caption         =   " Name"
         BoldText        =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "To change pin, type in correct name or choose from above combo box then dbl click on Name box and enter a new pin. Press Save."
         Height          =   630
         Left            =   720
         TabIndex        =   16
         Top             =   1410
         Width           =   3405
      End
   End
   Begin Project1.PanelFx PanelFx1 
      Height          =   2340
      Left            =   5730
      TabIndex        =   2
      Top             =   150
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   4128
      TileHeight      =   25
      TitleCaption    =   "Punch In/Punch Out"
      TitleForeColor  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleAlignment  =   1
      RoundEdge       =   15
      BackGroundStyle =   1
      gCTitleStart    =   8039500
      gCTitleEnd      =   12769963
      gCPanelStart    =   8039500
      gCPanelEnd      =   12769963
      Begin Project1.ccXPButton cmdCancel 
         Height          =   345
         Left            =   3285
         TabIndex        =   22
         Top             =   870
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         Caption         =   "Cancel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.ccXPButton cmdEnter 
         Height          =   780
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   1376
         Caption         =   "Enter"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboNameList 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Text            =   " Select Name"
         Top             =   450
         Width           =   4350
      End
      Begin Project1.ucTextbox txtEmpPin 
         Height          =   345
         Left            =   240
         TabIndex        =   3
         Top             =   885
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   609
         Caption         =   " Pin"
         MaxLength       =   6
         BoldText        =   -1  'True
         PassWordChar    =   "*"
      End
   End
   Begin Project1.StrokeText StrokeText1 
      Height          =   870
      Left            =   1065
      TabIndex        =   1
      Top             =   4395
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   1535
      Caption         =   "Time Clock"
      ColorFill       =   8039500
      Shadow          =   -1  'True
      TStyle          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   720
      Top             =   6900
   End
   Begin VB.PictureBox picClock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3465
      Left            =   1125
      ScaleHeight     =   3465
      ScaleWidth      =   3465
      TabIndex        =   0
      Top             =   570
      Width           =   3465
   End
   Begin Project1.PanelFx PanelFx3 
      Height          =   1725
      Left            =   5745
      TabIndex        =   12
      Top             =   3330
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3043
      TileHeight      =   25
      TitleCaption    =   "Clear Files / Delete Record"
      TitleForeColor  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleAlignment  =   1
      RoundEdge       =   15
      CanCollapse     =   -1  'True
      BackGroundStyle =   1
      gCTitleStart    =   8039500
      gCTitleEnd      =   12769963
      gCPanelStart    =   8039500
      gCPanelEnd      =   12769963
      Begin Project1.ccXPButton cmdDeleteRec 
         Height          =   450
         Left            =   420
         TabIndex        =   14
         Top             =   960
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   794
         Caption         =   "Delete Single Record"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.ccXPButton cmdClear 
         Height          =   495
         Left            =   420
         TabIndex        =   13
         Top             =   405
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   873
         Caption         =   "Clear all Records for New Week"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Name to delete from above combo box. No pin needed."
         Height          =   255
         Left            =   210
         TabIndex        =   15
         Top             =   1425
         Width           =   4455
      End
   End
   Begin Project1.PanelFx PanelFx4 
      Height          =   1275
      Left            =   5745
      TabIndex        =   17
      Top             =   3765
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2249
      TileHeight      =   25
      TitleCaption    =   "Reports"
      TitleForeColor  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleAlignment  =   1
      RoundEdge       =   15
      CanCollapse     =   -1  'True
      BackGroundStyle =   1
      gCTitleStart    =   8039500
      gCTitleEnd      =   12769963
      gCPanelStart    =   8039500
      gCPanelEnd      =   12769963
      Begin Project1.ccXPButton cmdShowReport 
         Height          =   495
         Left            =   90
         TabIndex        =   18
         Top             =   525
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   873
         Caption         =   "Show Reports"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "pin numbers must be at least 4 digits"
      Height          =   210
      Left            =   6150
      TabIndex        =   23
      Top             =   6180
      Width           =   2685
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Also there is very little or no error trapping , so its up to you."
      Height          =   450
      Left            =   6150
      TabIndex        =   21
      Top             =   6420
      Width           =   2340
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: For this demo, the Pin numbers are the employ digit times 4. ex: Emp1--- 1111"
      Height          =   660
      Left            =   6150
      TabIndex        =   20
      Top             =   5580
      Width           =   2550
   End
   Begin VB.Image Image1 
      Height          =   6630
      Left            =   120
      Picture         =   "Form1.frx":0000
      Top             =   135
      Width           =   5505
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 'Assumes for each start and end time period
      'if there is a start time then there is an end time ex. 8:00 / 12:00  --- lunch---  1:00 / 5:00
      'no more than one 24 hr period. ex. Starttime on a Monday and Endtime on a Wednesday (not good)
      'very little or no error trapping
      'pin numbers must be at least 4 digits
      
   Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

   'Analog clock code is not mine

   Dim hourHand As Single
   Dim minuteHand As Single
   Dim secondHand As Single

   Dim sizeX As Long
   Dim sizeY As Long
   Dim bCollapse2 As Boolean
   Dim bCollapse3 As Boolean
   Dim bCollapse4 As Boolean
   Dim arr() As String   'temp storage of currently selected employee's info

Private Sub Form_Load()
   ' Initialize form and timer
   picClock.FontSize = 18
   picClock.FontBold = True
   Timer1.Interval = 250
   Timer1.Enabled = True
   PanelFx2.Collapse True
   PanelFx3.Collapse True
   PanelFx4.Collapse True
   LoadList
   bCollapse2 = True
   bCollapse3 = True
   bCollapse4 = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Timer1.Enabled = False
   Unload Form2
   Unload Me
End Sub

Private Sub cboNameList_Click()
   LoadArray     'load select employees info into List1. Basically used to verify pin #
   txtEmpPin.SetFocus
End Sub

Private Sub cmdCancel_Click()
    cboNameList.Text = " Select Name"
    txtEmpPin.Text = ""
End Sub

Private Sub cmdClear_Click()
Dim x As Integer
Dim y As Integer
Dim respon As String

'clears all employee info except for the pin #
respon = MsgBox("Are you sure?", vbYesNo, "Clear all Records")
If respon = vbNo Then Exit Sub
For x = 0 To cboNameList.ListCount - 1
   cboNameList.Text = cboNameList.List(x)
   LoadArray
   WriteText App.Path & "\DataFolder\" & cboNameList.Text & ".txt", List1.List(0) & ",", False
   List1.Clear
Next x
cboNameList.Text = " Select Name"
MsgBox "DONE", , "Done"
End Sub

Private Sub cmdDeleteRec_Click()
   If cboNameList.Text = " Select Name" Then Exit Sub
   DeleteFile App.Path & "\DataFolder\" & cboNameList.Text & ".txt"
   MsgBox cboNameList.Text & "   Deleted", , "File Deleted"
   'clear everything so we can reload a new list
   cboNameList.Clear
   Form2.cboRptEmpList.Clear
   cboNameList.Text = " Select Name"
   Form2.cboRptEmpList.Text = " Select Name"
   LoadList
End Sub

Private Sub cmdEnter_Click()
   If cboNameList.Text = " Select Name" Or txtEmpPin.Text = "" Then Exit Sub    'no entries made
   If List1.List(0) <> txtEmpPin.Text Then
      MsgBox "Sorry, Wrong Pin Number. Try Again", , "Try Again"
      txtEmpPin.Text = ""
      cboNameList.Text = " Select Name"
      Exit Sub
   End If
   'if pin # match
   WriteText App.Path & "\DataFolder\" & cboNameList.Text & ".txt", Format(Now, "mm/dd/yyyy  hh:mm:ss AM/PM") & ",", True
   LoadArray
   cboNameList.Text = " Select Name"
   txtEmpPin.Text = ""
   MsgBox "Saved  " & Format(Now, "mm/dd/yyyy  hh:mm:ss AM/PM"), , "Time Saved"
End Sub

Private Sub cmdExit_Click()
   Timer1.Enabled = False
   Unload Form2
   Unload Me
End Sub

Private Sub cmdSaveInfo_Click()
   Dim respon As String
   
   If txtNewEmpPin.Text = "" Then
      MsgBox "No pin number entered", , "No Pin"
      Exit Sub
   End If
   'save new employee info
   If Dir(App.Path & "\DataFolder\" & txtNewEmpName.Text & ".txt", vbNormal) = "" Then
      'File doesn't exist
   Else
     respon = MsgBox("File already exists, overwrite anyways?", vbYesNo)
     If respon = vbNo Then Exit Sub
   End If

   
   WriteText App.Path & "\DataFolder\" & txtNewEmpName.Text & ".txt", txtNewEmpPin.Text & "," & vbNewLine, False
   txtNewEmpName.Text = ""
   txtNewEmpPin.Text = ""
   cboNameList.Clear
   cboNameList.Text = " Select Name"
   Form2.cboRptEmpList.Clear
   Form2.cboRptEmpList.Text = " Select Name"
   LoadList
   MsgBox "Saved", , "Info Saved"
End Sub

Private Sub cmdShowReport_Click()
   Form2.Show
End Sub

Private Sub PanelFx2_PanelClick()
   PanelFx2_TileClick
End Sub

Private Sub PanelFx2_TileClick()
   bCollapse2 = Not bCollapse2
   PanelFx2.Collapse bCollapse2
   txtNewEmpName.Text = ""
End Sub

Private Sub PanelFx3_PanelClick()
   PanelFx3_TileClick
End Sub

Private Sub PanelFx3_TileClick()
   bCollapse3 = Not bCollapse3
   PanelFx3.Collapse bCollapse3
End Sub

Private Sub PanelFx4_PanelClick()
   PanelFx4_TileClick
End Sub

Private Sub PanelFx4_TileClick()
   bCollapse4 = Not bCollapse4
   PanelFx4.Collapse bCollapse4
End Sub

Private Sub picClock_Paint()
   ' Draws the clock face when needed
   Dim num As String
   Dim Angle As Single, sinX As Single, cosY As Single
   Dim tick As Long, fontX As Long, fontY As Long
   Static Busy As Boolean
   
   ' Avoid re-entry problems (resizing form can cause many paints)
   If Busy Then Exit Sub
   Busy = True
   
   ' Fit to current form size
   sizeX = picClock.ScaleWidth / 2
   sizeY = picClock.ScaleHeight / 2
   
   ' Start with a blank screen
   picClock.AutoRedraw = True
   picClock.Cls
   picClock.DrawWidth = 3
   
   ' Loop through clock circle (starts at 1 o'clock)
   For Angle = 8.9 To 2.7 Step -0.1047198
      ' Draw tick marks
      tick = tick + 1
      sinX = Sin(Angle)
      cosY = Cos(Angle)
      picClock.Line (sizeX + sinX * (sizeX * 0.9), _
      sizeY + cosY * (sizeY * 0.9))- _
      (sizeX + sinX * sizeX, _
      sizeY + cosY * sizeY), _
      vbBlack
      ' Make every 5th tick darker
      Select Case tick Mod 5
         Case 0
            picClock.DrawWidth = 3
         Case 1
            ' Center number where it belongs
            num = CStr(tick \ 5 + 1)
            fontX = picClock.TextWidth(num) / 2
            fontY = picClock.TextHeight(num) / 2
            picClock.PSet (sizeX + sinX * (sizeX * 0.75) - fontX, _
            sizeY + cosY * (sizeY * 0.75) - fontY), _
            picClock.BackColor
            picClock.Print num
            picClock.DrawWidth = 1
         Case Else
            picClock.DrawWidth = 1
      End Select
      Next
      
      ' Save image to memory
      picClock.AutoRedraw = False
      ' Force redraw of hands
      secondHand = -1
      
      Busy = False
End Sub

Private Sub LoadList()
   Dim sFile As String
   Dim lgh As Integer
   'load names into combo boxes
   sFile = Dir$(App.Path & "\DataFolder\*.txt", vbNormal)                       'Get first entry
   Do Until Len(sFile) = 0                             'Loop until we run out (sFile will be empty)
      lgh = Len(sFile)
      cboNameList.AddItem Left$(sFile, lgh - 4)                             'Display what we found, minus extension
      Form2.cboRptEmpList.AddItem Left$(sFile, lgh - 4)
      sFile = Dir$                                    'Get next entry
      Loop
End Sub

Private Sub LoadArray()
   Dim textfile As String
   Dim i As Integer
   Dim ff As Integer
   List1.Clear
   ff = FreeFile
   Open App.Path & "\DataFolder\" & cboNameList.Text & ".txt" For Input As #ff
   Do While Not EOF(ff)
      Line Input #ff, textfile
      arr = Split(textfile, ",")
      For i = 0 To UBound(arr)
         If Not Trim(arr(i)) = "" Then
            List1.AddItem arr(i)
         End If
         Next
         Loop
         Close #ff
End Sub

Private Function WriteText(FileName As String, Text As String, Optional Append As Boolean = False)
   Dim ff As Integer
   
   On Error GoTo Handle
   
   ff = FreeFile
   If Append = True Then
      Open FileName For Append As #ff
   Else
      Open FileName For Output As #ff
   End If
   Print #ff, Text '; - doesn't include the trailing newline
   Close #ff
   Exit Function
Handle:
   WriteText = False
   MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
End Function

Private Sub Timer1_Timer()
   ' Draws the clock hands when seconds have changed
   Dim sec As Single, cir As Single
   
   ' FYI:
   ' pi = Atn(1) * 4
   ' date = Int(Now)
   ' time = Now - Int(Now)
   
   ' hour hand makes 2 revolutions per day (pi * 4)
   cir = Atn(1) * 16
   hourHand = (Now - Int(Now)) * cir
   ' and it needs to go clockwise (radians increase counter-clockwise)
   hourHand = cir - hourHand
   ' minute hand is 12 times as fast as the hour hand
   minuteHand = hourHand * 12
   ' second hand is 60 times as fast as the minute hand
   sec = minuteHand * 60
   
   ' only re-draw if the seconds have changed
   If secondHand <> sec Then
      secondHand = sec
      picClock.Cls
      picClock.DrawWidth = 7
      picClock.Line (sizeX, sizeY)-Step(-Sin(hourHand) * (sizeX * 0.5), _
      -Cos(hourHand) * (sizeY * 0.5)), _
      vbBlack
      picClock.DrawWidth = 3
      picClock.Line (sizeX, sizeY)-Step(-Sin(minuteHand) * (sizeX * 0.85), _
      -Cos(minuteHand) * (sizeY * 0.85)), _
      vbBlack
      picClock.DrawWidth = 1
      picClock.Line (sizeX, sizeY)-Step(-Sin(secondHand) * sizeX, _
      -Cos(secondHand) * sizeY), _
      vbRed
   End If
End Sub

Private Sub txtNewEmpName_DblClick()
   If cboNameList.Text = " Select Name" Then Exit Sub
   txtNewEmpName.Text = cboNameList.Text
End Sub
