VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~1.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Music"
   ClientHeight    =   5505
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1920
      Width           =   6495
   End
   Begin RichTextLib.RichTextBox r 
      Height          =   735
      Left            =   3480
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   3360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3960
      Top             =   1320
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   735
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      _Version        =   786432
      _ExtentX        =   1931
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Stop"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0084
      Left            =   480
      List            =   "Form1.frx":0086
      TabIndex        =   3
      Text            =   "Accordion"
      Top             =   480
      Width           =   2535
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   735
      Left            =   3960
      TabIndex        =   2
      Top             =   360
      Width           =   1095
      _Version        =   786432
      _ExtentX        =   1931
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Start"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Text            =   "500"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7440
      Top             =   3480
   End
   Begin VB.Label Label1 
      Caption         =   "Tempo : "
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin WMPLibCtl.WindowsMediaPlayer w 
      Height          =   15
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   15
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "invisible"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   26
      _cy             =   26
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
      Begin VB.Menu h 
         Caption         =   "-"
      End
      Begin VB.Menu help 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim q As Integer, qq As Integer
Dim z As Integer
Dim s As String


Private Sub help_Click()
Form2.Show
End Sub

Private Sub open_Click()
On Error Resume Next
C.Filter = "Music Master(*.MUM)|*.MUM|"

C.ShowOpen
r.LoadFile C.FileName
Text1.Text = r.Text
End Sub

Private Sub PushButton1_Click()
z = Combo1.ListIndex
Combo1.Enabled = False
PushButton1.Enabled = False
PushButton2.Enabled = True
Timer1.Interval = Text2.Text
Timer1.Enabled = True
Timer2.Interval = 10
qq = 0
q = 0
s = ""
a = ""
End Sub

Private Sub PushButton2_Click()
PushButton2.Enabled = False
PushButton1.Enabled = True
Combo1.Enabled = True
w.Close
a = ""
q = 0
Timer1.Enabled = False
End Sub

Private Sub save_Click()
On Error Resume Next
C.Filter = "Music Master(*.MUM)|*.MUM|"

C.ShowSave
r.Text = Text1.Text
r.SaveFile C.FileName
End Sub

Private Sub Text2_Change()
Timer1.Interval = Text2.Text
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = Text2.Text
Select Case z
Case -1
If Len(Text1.Text) < q Then
Timer1.Enabled = False
q = 0
a = ""
w.Controls.stop
PushButton2_Click
GoTo L:
End If
If q = 0 Then
q = q + 1

Else

q = q + 2
End If
a = Mid(Text1.Text, q, 2)

Select Case Trim(a)

Case "m"
w.URL = App.Path & "\" & "Untitled.wma"
Case "p"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (2).wma"
Case "d"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (3).wma"
Case "n"
w.URL = App.Path & "\" & "Untitled (4).wma"
Case "S"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (5).wma"
Case "R"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (6).wma"
Case "G"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (7).wma"
Case "M"
w.URL = App.Path & "\" & "Untitled (8).wma"
Case "P"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (9).wma"
Case "D"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (10).wma"
Case "N"
w.URL = App.Path & "\" & "Untitled (11).wma"
Case "S'"
w.URL = App.Path & "\" & "Untitled (12).wma"
Case "R'"
w.URL = App.Path & "\" & "Untitled (13).wma"
Case "G'"
w.URL = App.Path & "\" & "Untitled (14).wma"
Case "M'"
w.URL = App.Path & "\" & "Untitled (15).wma"
Case "P'"
w.URL = App.Path & "\" & "Untitled (16).wma"
Case "D'"
w.URL = App.Path & "\" & "Untitled (17).wma"
Case "N'"
w.URL = App.Path & "\" & "Untitled (18).wma"
Case "m."
w.URL = App.Path & "\" & "Untitled (19).wma"
Case "d."
w.URL = App.Path & "\" & "Untitled (20).wma"
Case "n."
w.URL = App.Path & "\" & "Untitled (21).wma"
Case "R."
w.URL = App.Path & "\" & "Untitled (22).wma"
Case "G."
w.URL = App.Path & "\" & "Untitled (23).wma"
Case "M."
w.URL = App.Path & "\" & "Untitled (24).wma"
Case "D."
w.URL = App.Path & "\" & "Untitled (25).wma"
Case "N."
w.URL = App.Path & "\" & "Untitled (26).wma"
Case "R;"
w.URL = App.Path & "\" & "Untitled (27).wma"
Case "G;"
w.URL = App.Path & "\" & "Untitled (28).wma"
Case "M;"
w.URL = App.Path & "\" & "Untitled (29).wma"
Case "D;"
w.URL = App.Path & "\" & "Untitled (30).wma"
Case "N;"
w.URL = App.Path & "\" & "Untitled (31).wma"
Case "+"
w.Controls.stop
Case "\" & 2
Timer2.Enabled = True
Timer1.Enabled = False
s = Mid(Text1.Text, q + 2, 4)
Timer2.Interval = Int(Timer1.Interval / 2)
Case "\" & 3
Timer2.Enabled = True
Timer1.Enabled = False
s = Mid(Text1.Text, q + 2, 6)

Timer2.Interval = Int(Timer1.Interval / 3)

End Select

End Select
L:
End Sub

Private Sub Timer2_Timer()

If Len(s) = qq + 1 Then
Timer1.Interval = 4
Timer1.Enabled = True
Timer2.Enabled = False
q = q + qq + 3

Timer2.Interval = 1
qq = 0
s = ""
GoTo L:
End If
If qq = 0 Then
qq = qq + 1
Else
qq = qq + 2
End If
a = Mid(s, qq, 2)
Select Case Trim(a)

Case "m"
w.URL = App.Path & "\" & "Untitled.wma"
Case "p"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (2).wma"
Case "d"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (3).wma"
Case "n"
w.URL = App.Path & "\" & "Untitled (4).wma"
Case "S"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (5).wma"
Case "R"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (6).wma"
Case "G"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (7).wma"
Case "M"
w.URL = App.Path & "\" & "Untitled (8).wma"
Case "P"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (9).wma"
Case "D"
w.Controls.stop
w.URL = App.Path & "\" & "Untitled (10).wma"
Case "N"
w.URL = App.Path & "\" & "Untitled (11).wma"
Case "S'"
w.URL = App.Path & "\" & "Untitled (12).wma"
Case "R'"
w.URL = App.Path & "\" & "Untitled (13).wma"
Case "G'"
w.URL = App.Path & "\" & "Untitled (14).wma"
Case "M'"
w.URL = App.Path & "\" & "Untitled (15).wma"
Case "P'"
w.URL = App.Path & "\" & "Untitled (16).wma"
Case "D'"
w.URL = App.Path & "\" & "Untitled (17).wma"
Case "N'"
w.URL = App.Path & "\" & "Untitled (18).wma"
Case "m."
w.URL = App.Path & "\" & "Untitled (19).wma"
Case "d."
w.URL = App.Path & "\" & "Untitled (20).wma"
Case "n."
w.URL = App.Path & "\" & "Untitled (21).wma"
Case "R."
w.URL = App.Path & "\" & "Untitled (22).wma"
Case "G."
w.URL = App.Path & "\" & "Untitled (23).wma"
Case "M."
w.URL = App.Path & "\" & "Untitled (24).wma"
Case "D."
w.URL = App.Path & "\" & "Untitled (25).wma"
Case "N."
w.URL = App.Path & "\" & "Untitled (26).wma"
Case "R;"
w.URL = App.Path & "\" & "Untitled (27).wma"
Case "G;"
w.URL = App.Path & "\" & "Untitled (28).wma"
Case "M;"
w.URL = App.Path & "\" & "Untitled (29).wma"
Case "D;"
w.URL = App.Path & "\" & "Untitled (30).wma"
Case "N;"
w.URL = App.Path & "\" & "Untitled (31).wma"
Case "+"
w.Controls.stop

End Select

If Len(s) = qq + 1 Then
Timer2.Interval = Timer2.Interval - 5
End If
L:
End Sub
