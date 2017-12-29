VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{062F78EB-5E1D-43A7-9135-96E1392B3950}#1.0#0"; "MoviePlayer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "X-Control"
   ClientHeight    =   10650
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14985
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   10650
   ScaleWidth      =   14985
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   120
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   14400
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   14400
      Top             =   720
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   $"Form1.frx":3D9A7
   End
   Begin VB.Frame Frame2 
      Caption         =   "Song's"
      Height          =   4335
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   12255
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   6600
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   0
         Top             =   3720
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\DataBases\Database1.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\DataBases\Database1.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Karokeurl"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Command7 
         Caption         =   "End Select"
         Height          =   495
         Left            =   10560
         TabIndex        =   15
         Top             =   3720
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form1.frx":3DA36
         Height          =   2775
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   4895
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Nama"
            Caption         =   "Nama"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Url"
            Caption         =   "Url"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   6089,953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5490,142
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4200
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Song"
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Singer"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Room 1 Monitoring"
      Height          =   4815
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   13935
      Begin VB.CommandButton CmdPlay 
         Caption         =   "Play"
         Height          =   495
         Left            =   9120
         TabIndex        =   25
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox CmdSoundOut 
         Height          =   360
         Left            =   10920
         TabIndex        =   23
         Text            =   "Default DirectSound Device"
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CommandButton CmdPause 
         Caption         =   "Pause"
         Height          =   495
         Left            =   10680
         TabIndex        =   22
         Top             =   1080
         Width           =   1335
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   4320
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   7080
         Max             =   1000
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MOVIEPLAYERLib.MoviePlayer MoviePlayer1 
         Height          =   3855
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   6375
         _Version        =   65536
         _ExtentX        =   11245
         _ExtentY        =   6800
         _StockProps     =   0
      End
      Begin VB.HScrollBar ScrollVolume 
         Height          =   255
         LargeChange     =   10
         Left            =   10920
         Max             =   10000
         Min             =   6000
         TabIndex        =   18
         Top             =   3960
         Value           =   8000
         Width           =   2415
      End
      Begin VB.ListBox LUrl 
         Height          =   3180
         Left            =   6840
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CmdEmpty 
         Caption         =   "Empty List"
         Height          =   495
         Left            =   10680
         TabIndex        =   8
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton CmdSong 
         Caption         =   "Song's"
         Height          =   495
         Left            =   9120
         TabIndex        =   7
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton CmdVocal 
         BackColor       =   &H8000000D&
         Caption         =   "Vocal On"
         Height          =   495
         Left            =   10680
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton CmdNextSong 
         Caption         =   "Next Song"
         Height          =   495
         Left            =   9120
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "Stop"
         Height          =   495
         Left            =   12240
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ListBox LNama 
         Enabled         =   0   'False
         Height          =   4140
         Left            =   6840
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Audio Output"
         Height          =   255
         Left            =   9120
         TabIndex        =   24
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume"
         Height          =   255
         Left            =   9120
         TabIndex        =   19
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Played Song's"
         BeginProperty Font 
            Name            =   "OCR A Extended"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   17
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Playlist"
         Height          =   255
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Menu rom 
      Caption         =   "Room Active"
      Begin VB.Menu rom1 
         Caption         =   "Room 1"
         Checked         =   -1  'True
         Shortcut        =   {F2}
      End
      Begin VB.Menu rom2 
         Caption         =   "Room 2"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu set 
      Caption         =   "Setting"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdEmpty_Click()
LNama.Clear
LUrl.Clear
End Sub

Private Sub CmdNextSong_Click()
If LUrl.List(0) = "" Then
MsgBox "Empty Playlist"
Else
Me.MoviePlayer1.stop
Form2.MoviePlayer1.stop
Me.MoviePlayer1.FileName = LUrl.List(0)
Form2.MoviePlayer1.FileName = LUrl.List(0)
Me.MoviePlayer1.play
Me.MoviePlayer1.SoundVolume = Me.HScroll2.Value - 10000
Form2.MoviePlayer1.play
LNama.RemoveItem (0)
LUrl.RemoveItem (0)
End If
End Sub

Private Sub CmdPause_Click()
If CmdPause.Caption = "Pause" Then
Me.MoviePlayer1.pause
Form2.MoviePlayer1.pause
CmdPause.Caption = "Resume"
Else
CmdPause.Caption = "Pause"
Me.MoviePlayer1.play
Form2.MoviePlayer1.play
End If
End Sub

Private Sub CmdPlay_Click()
If LUrl.List(0) = "" Then
MsgBox "Empty Playlist"
Else
Label4.Caption = LNama.List(0)
Me.MoviePlayer1.FileName = LUrl.List(0)
Me.MoviePlayer1.SoundVolume = Me.HScroll2.Value - 10000
Form2.MoviePlayer1.UseVMR9 = True
Form2.MoviePlayer1.UseVolumeBoost = True
Form2.MoviePlayer1.VolumeGain = 1
Form2.MoviePlayer1.VolumeAutoGain = True
Form2.MoviePlayer1.FileName = LUrl.List(0)
Form2.MoviePlayer1.SetMPEG1AudioChannel 2
Form2.MoviePlayer1.SoundVolume = Me.ScrollVolume.Value - 10000
Me.MoviePlayer1.play
Form2.MoviePlayer1.play
LNama.RemoveItem (0)
LUrl.RemoveItem (0)
CmdStop.Visible = True
Me.Slider1.Max = Me.MoviePlayer1.duration
Timer2.Enabled = True
Timer1.Enabled = True
End If
Form2.MoviePlayer1.DrawText 0, 0, 2, "X-KarokeProject ~ BakaMomo48 ~", 15, "Arial", False, False, False, RGB(0, 0, 0), RGB(255, 255, 255), RGB(0, 0, 0), 1, 250, 0
MoviePlayer1.DrawText 0, 0, 2, "X-KarokeProject ~ BakaMomo48 ~", 5, "Arial", False, False, False, RGB(0, 0, 0), RGB(255, 255, 255), RGB(0, 0, 0), 1, 250, 0
End Sub

Private Sub CmdSong_Click()
Frame2.Visible = True
End Sub

Private Sub CmdStop_Click()
Me.MoviePlayer1.stop
Form2.MoviePlayer1.stop
End Sub

Private Sub CmdVocal_Click()
If CmdVocal.Caption = "Vocal Off" Then
Form2.MoviePlayer1.VolumeAudioChannel = 1
CmdVocal.Caption = "Vocal On"
Else
Form2.MoviePlayer1.VolumeAudioChannel = 0
CmdVocal.Caption = "Vocal Off"
End If
End Sub

Private Sub Command7_Click()
Frame2.Visible = False
End Sub

Private Sub DataGrid1_Click()
LNama.AddItem (Adodc1.Recordset.Fields!Nama)
LUrl.AddItem (Adodc1.Recordset.Fields!URL)
End Sub

Private Sub exit_Click()
Unload Me
Unload Form2
End Sub

Private Sub Form_Load()
Dim lihat As New Recordset
Dim sql As String
buka
con.CursorLocation = adUseClient
Set lihat = New Recordset
sql = "Select * From Karokeurl"
lihat.Open sql, con, adOpenStatic, adLockReadOnly
lihat.Sort = "Nama"
iAudioRendererCount = MoviePlayer1.GetAudioRendererCount

For i = 0 To iAudioRendererCount - 1
    CmdSoundOut.AddItem MoviePlayer1.GetAudioRendererName(i)
        
    ' the default audio renderer should Default DirectSound Device, another renderer may affect the audio syn
        
    If MoviePlayer1.GetAudioRendererName(i) = "Default DirectSound Device" Then
            CmdSoundOut.ListIndex = i
    End If
Next
Adodc1.Recordset.Sort = "Nama"
End Sub

Private Sub MoviePlayer1_OnCompleted()
If LNama.List(0) = "" Then
MsgBox "Empty Playlist"
Else
Me.MoviePlayer1.FileName = LUrl.List(0)
Form2.MoviePlayer1.FileName = LUrl.List(0)
Me.MoviePlayer1.play
Form2.MoviePlayer1.play
LNama.RemoveItem (0)
LUrl.RemoveItem (0)
End If
End Sub

Private Sub MoviePlayer1_OnPlaying(ByVal iCurrent As Double, ByVal strTime As String)
Me.Slider1.Value = iCurrent
End Sub

Private Sub rom1_Click()
Frame1.Visible = True
Form2.Show
End Sub

Private Sub ScrollVolume_Change()
Form2.MoviePlayer1.SoundVolume = ScrollVolume.Value - 10000
End Sub

Private Sub Text1_Change()
Call buka
rst.Sort = "Nama"
rst.CursorLocation = adUseClient
rst.Open "Select * from Karokeurl where Nama like '%" & Text1 & "%'", con
If Not rst.EOF Then
    With rst
        With DataGrid1
            Set .DataSource = rst
                .Refresh
        End With
    End With
End If
End Sub

Private Sub Timer1_Timer()
Label4.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

Private Sub Timer2_Timer()
a = Label4.Caption
b = Left(a, 1)
c = Right(a, Len(a) - 1)
Label4.Caption = c + b

End Sub
