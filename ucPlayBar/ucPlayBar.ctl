VERSION 5.00
Begin VB.UserControl ucPlayBar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   HasDC           =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   6690
   Begin MPlayBarDemo.DSPicButton cmdDSPause 
      Height          =   465
      Left            =   -1120
      TabIndex        =   4
      ToolTipText     =   "Play"
      Top             =   60
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   820
      ImageDown       =   "ucPlayBar.ctx":0000
      ImageHot        =   "ucPlayBar.ctx":0BF2
      ImageDisabled   =   "ucPlayBar.ctx":17E4
      ImageUp         =   "ucPlayBar.ctx":23D6
   End
   Begin MPlayBarDemo.DSPicButton cmdDSPlay 
      Height          =   465
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Play"
      Top             =   60
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   820
      ImageDown       =   "ucPlayBar.ctx":2FC8
      ImageHot        =   "ucPlayBar.ctx":3BBA
      ImageDisabled   =   "ucPlayBar.ctx":47AC
      ImageUp         =   "ucPlayBar.ctx":539E
   End
   Begin MPlayBarDemo.ucSlider ucSlider1 
      Height          =   480
      Left            =   1920
      Top             =   15
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   847
      SliderImage     =   "ucPlayBar.ctx":5F90
      Orientation     =   0
      RailImage       =   "ucPlayBar.ctx":649A
      RailStyle       =   99
      Max             =   100
   End
   Begin MPlayBarDemo.DSPicButton cmdDSStop 
      Height          =   465
      Left            =   600
      TabIndex        =   1
      Top             =   60
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   820
      ImageDown       =   "ucPlayBar.ctx":706C
      ImageHot        =   "ucPlayBar.ctx":7C5E
      ImageDisabled   =   "ucPlayBar.ctx":8850
      ImageUp         =   "ucPlayBar.ctx":9442
   End
   Begin MPlayBarDemo.DSPicButton cmdDSMute 
      Height          =   465
      Left            =   1440
      TabIndex        =   2
      Top             =   45
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   820
      ImageDown       =   "ucPlayBar.ctx":A034
      ImageHot        =   "ucPlayBar.ctx":AC86
      ImageUp         =   "ucPlayBar.ctx":B8D8
      Style           =   1
   End
   Begin VB.Label lbPlayBar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Station Playing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3000
      TabIndex        =   3
      Top             =   165
      Width           =   1275
   End
   Begin VB.Image ImgBack 
      Height          =   540
      Left            =   0
      Picture         =   "ucPlayBar.ctx":C10A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2.45745e5
   End
End
Attribute VB_Name = "ucPlayBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ======================================================================================
' File  :     ucPlayBar
' Author:     DracullSoft
' Date  :     01-04-2008
' --------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------
' Copyright © 2008 DracullSoft
' Author or copyright holders can not be held responsible for anything.
' Application: MediaPlaybarDemo
' --------------------------------------------------------------------------------------
' Purpose:
'   I wrote this control to mimic the Media player 10 because i wanted to
'   have the same look no matter what mediaplayer version 9,10,11 installed on the machine
'   The control used is Charles P.V ucSlider and a modified version of the
'   Simple ActiveButton by Gene Martynov
'   This was also for use with FLV player that looked like Media Player 10
'
' ======================================================================================
Option Explicit

Private pv_lTPPx                                As Long    ' TwipsPerPixelX
Private pv_lTPPy                                As Long    ' TwipsPerPixelY
Private m_Caption                               As String

Public Event ButtonClick(ByVal Key As String)
Public Event Volume(ByVal val As Long)

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    ' __ no storring of caption property only dynamic' PropertyChanged "Caption"
    lbPlayBar.Caption = m_Caption
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

'__ used for resetting the play button if the stream can not be read
Public Sub SetState(ByVal Key As String)
  Select Case LCase(Key)
    Case "stop"
      'cmdDSStop_Click
      cmdDSStop.Enabled = False
      cmdDSPlay.Left = 120
      cmdDSPause.Left = -2000
  End Select
End Sub




Private Sub cmdDSPause_Click()
  cmdDSPlay.Left = 120
  cmdDSPause.Left = -2000
  RaiseEvent ButtonClick("Pause")
End Sub

Private Sub cmdDSPlay_Click()
  cmdDSPlay.Left = -2000
  cmdDSPause.Left = 120
  cmdDSStop.Enabled = True
  RaiseEvent ButtonClick("Play")
End Sub

Private Sub cmdDSStop_Click()
  cmdDSStop.Enabled = False
  cmdDSPlay.Left = 120
  cmdDSPause.Left = -2000
  RaiseEvent ButtonClick("Stop")
End Sub

Private Sub cmdDSMute_Click()
  If cmdDSMute.Value = abUnPressed Then
    RaiseEvent ButtonClick("NoMute")
  Else
    RaiseEvent ButtonClick("Mute")
  End If
End Sub

Private Sub ucSlider1_Change()
  RaiseEvent Volume(ucSlider1.Value)
End Sub




Private Sub UserControl_Initialize()
    pv_lTPPx = Screen.TwipsPerPixelX
    pv_lTPPy = Screen.TwipsPerPixelY
    
End Sub



'UserControl.Ambient.UserMode can be used to detect if its running in
'either readproperties or show or resize.. not in initialize
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  '__test pause button
  
  If UserControl.Ambient.UserMode Then
  '  cmdDSPause.Left = 120
    cmdDSStop.Enabled = False
    lbPlayBar.Caption = ""
  End If
  
'  If UserControl.Ambient.UserMode Then
'    ' setup image for buttons if runtime
'    ImgPlay(0).Top = 60
'    ImgPlay(1).Top = 60
'    ImgPlay(2).Top = 60
'    ImgPlay(0).Visible = True
'    ImgPlay(1).Visible = False
'    ImgPlay(2).Visible = False
'
'    ImgStop(0).Top = 60
'    ImgStop(1).Top = 60
'    ImgStop(0).Visible = True
'    ImgStop(1).Visible = False
'
'    ImgS(0).Top = 60
'    ImgS(1).Top = 60
'    ImgS(0).Visible = True
'    ImgS(1).Visible = False
'  End If
  
End Sub

Private Sub UserControl_Resize()
  ImgBack.Width = UserControl.Width
End Sub


