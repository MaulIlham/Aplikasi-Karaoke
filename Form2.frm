VERSION 5.00
Object = "{062F78EB-5E1D-43A7-9135-96E1392B3950}#1.0#0"; "MoviePlayer.ocx"
Begin VB.Form Form2 
   Caption         =   "X-KarokeProject Room 1"
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13500
   LinkTopic       =   "Form2"
   ScaleHeight     =   9180
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin MOVIEPLAYERLib.MoviePlayer MoviePlayer1 
      Height          =   8535
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   12855
      _Version        =   65536
      _ExtentX        =   22675
      _ExtentY        =   15055
      _StockProps     =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   2880
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MoviePlayer1_Click()
Me.MoviePlayer1.ShowFullScreen True
End Sub

