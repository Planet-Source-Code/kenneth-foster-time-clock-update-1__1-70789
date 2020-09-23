VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00CEE0BC&
   Caption         =   "Time Report"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   9795
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin Project1.PanelFx PanelFx1 
      Height          =   9630
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   16986
      TitleCaption    =   "Time Report"
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
      BackGroundStyle =   1
      gCTitleStart    =   8039500
      gCTitleEnd      =   12769963
      gCPanelStart    =   8039500
      gCPanelEnd      =   12769963
      Begin VB.ListBox List1 
         Height          =   8445
         Left            =   105
         TabIndex        =   4
         Top             =   1035
         Width           =   3375
      End
      Begin VB.ListBox List2 
         Height          =   6300
         Left            =   105
         TabIndex        =   5
         Top             =   1035
         Width           =   2175
      End
      Begin VB.ComboBox cboRptEmpList 
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
         Left            =   105
         TabIndex        =   3
         Text            =   " Select Name"
         Top             =   315
         Width           =   3405
      End
      Begin Project1.ccXPButton cmdPrint 
         Height          =   495
         Left            =   3645
         TabIndex        =   2
         Top             =   6720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "Print"
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
      Begin Project1.ccXPButton cmdClose 
         Height          =   495
         Left            =   3675
         TabIndex        =   1
         Top             =   8850
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   873
         Caption         =   "Close"
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Time Card"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   1050
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 'Assumes for each start and end time period
      'if there is a start time then there is an end time ex. 8:00 / 12:00  --- lunch---  1:00 / 5:00
      'no more than one 24 hr period. ex. Starttime on a Monday and Endtime on a Wednesday (not good)
      'very little or no error trapping
      ' pin number must be at least 4 digits
Dim arrRpt() As String
Dim arrCal() As String

Private Sub cboRptEmpList_Click()
   GetitArray
   CalTime
End Sub

Private Sub cmdClose_Click()
   List1.Clear
   List2.Clear
   cboRptEmpList.Text = " Select Name"
   Form2.Hide
End Sub

Private Sub GetitArray()
   Dim textfile As String
   Dim i As Integer
   Dim ff As Integer
   Dim lgh As Integer
   
   List1.Clear
   List2.Clear    ' hidden
  
   ff = FreeFile
   Open App.Path & "\DataFolder\" & cboRptEmpList.Text & ".txt" For Input As #ff
   Do While Not EOF(ff)
      Line Input #ff, textfile
      arrRpt = Split(textfile, ",")
      
      For i = 0 To UBound(arrRpt)
         If Not Trim(arrRpt(i)) = "" Then
            lgh = Len(arrRpt(i))
            If lgh <= 4 Then GoTo here   'skip the pin number
            List2.AddItem (arrRpt(i))    'extract time and date
here:
         End If
      Next i
      Loop
      Close #ff
   End Sub
   
Private Sub CalTime()
   'Assumes for each start and end time period
      'if there is a start time then there is an end time ex. 8:00 / 12:00  --- lunch---  1:00 / 5:00
      'no more than one 24 hr period. ex. Starttime on a Monday and Endtime on a Wednesday (not good)
      'very little or no error trapping
      
   Dim starttimeHour As String
   Dim starttimeMin As String
   Dim endtimeHour As String
   Dim endtimeMin As String
   Dim x As Integer
   Dim stg1 As String
   Dim stg2 As String
   
   Static TotHour As Integer
   Static TotMin As Integer
   Static S_Hour As Integer
   Static S_Min As Integer
   Static E_Hour As Integer
   Static E_Min As Integer
   Static Hr_Min_String As String
   
   TotHour = 0
   TotMin = 0
   For x = 1 To List2.ListCount - 1 Step 2
      
      starttimeHour = Mid$(List2.List(x - 1), 13, 2)
      starttimeMin = Mid$(List2.List(x - 1), 16, 2)
      
      endtimeHour = Mid$(List2.List(x), 13, 2)
      endtimeMin = Mid$(List2.List(x), 16, 2)
      
      S_Hour = starttimeHour
      S_Min = starttimeMin
      E_Hour = endtimeHour
      E_Min = endtimeMin
      
   If E_Hour < S_Hour Then E_Hour = E_Hour + 12   'Starting hour is > than end hour
      
   If E_Min >= S_Min Then      'ending min > than start min
      Hr_Min_String = Format(E_Hour - S_Hour, "0#") & " hrs  " & Format(E_Min - S_Min, "0#") & " min"
   Else   ' starting min > than end min
      E_Min = E_Min + 60
      Hr_Min_String = Format(E_Hour - S_Hour, "0#") & " hrs  " & Format(E_Min - S_Min, "0#") & " min"
   End If
   
   If E_Hour >= S_Hour Then    'Starting hour <=  to end hour
      If E_Min >= S_Min Then    'ending min > than start min
         Hr_Min_String = Format(E_Hour - S_Hour, "0#") & " hrs  " & Format(E_Min - S_Min, "0#") & " min"
      Else   'starting min > than end min
         Hr_Min_String = Format(E_Hour - S_Hour, "0#") & " hrs  " & Format(S_Min - E_Min, "0#") & " min"
      End If
   End If
   
'keeps running total of hours and minutes
stg1 = Left$(Hr_Min_String, 2)
TotHour = TotHour + Int(stg1)
stg2 = Mid$(Hr_Min_String, 9, 2)
TotMin = TotMin + Int(stg2)

List1.AddItem ""                              ' add a line space
List1.AddItem List2.List(x - 1)
List1.AddItem List2.List(x) & "    " & Hr_Min_String       'add hours and minutes to List1

Next x

'if minutes are greater than 60
If TotMin >= 60 Then
   Do Until TotMin < 60
      TotMin = TotMin - 60  'subtract 60 from minutes
      TotHour = TotHour + 1        'add one to hours
   Loop
End If

List1.AddItem "                                 --------------------------------------"
List1.AddItem "                                 Total" & "     " & Format(TotHour, "0#") & " hrs  " & Format(TotMin, "0#") & " mins"
End Sub

Private Sub cmdPrint_Click()
    Dim lngCount As Long

   On Error GoTo ErrorExit

   Printer.FontBold = True
   Printer.FontSize = 10
   Printer.Print
   Printer.Print "File printed: " & Now   'datetime stamp
   Printer.Print ""     'blank line
   Printer.Print cboRptEmpList.Text        'print employees name
   Printer.Print
   
   For lngCount = 0 To List1.ListCount - 1
      Printer.Print List1.List(lngCount)     'send text to printer object
   Next lngCount
   Printer.Print
   
   Printer.EndDoc       'release printer object to printer
   Exit Sub

ErrorExit:
   Dim strErrMsg As String
   strErrMsg = "Error number: " & Err.Number & vbCrLf & "Error Desc: " & Err.Description
   MsgBox strErrMsg

End Sub

