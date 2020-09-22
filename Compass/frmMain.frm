VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Compass"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCompass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   3840
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   8
      Top             =   60
      Width           =   3030
      Begin VB.Shape ArrawHead 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0080FF80&
         FillColor       =   &H00008000&
         Height          =   135
         Left            =   2700
         Shape           =   3  'Circle
         Top             =   2460
         Width           =   135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         X1              =   1560
         X2              =   2220
         Y1              =   1740
         Y2              =   2520
      End
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtLongitude 
      Height          =   315
      Left            =   1260
      TabIndex        =   4
      Text            =   "31.3"
      Top             =   1500
      Width           =   735
   End
   Begin VB.TextBox txtLatitude 
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Text            =   "30.1"
      Top             =   900
      Width           =   735
   End
   Begin VB.Label lblQibla 
      Height          =   255
      Left            =   1260
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Angle:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "Longitude:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Latitude:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Compass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   3675
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================
'           Copyright Information
'==========================================================
'Program Name     : Mewsoft Qibla Direction Compass
'Program Author   : Elsheshtawy, A. A.
'Home Page        : http://www.mewsoft.com
'Home Page        : http://www.islamware.com
'Copyrights © 2006 Mewsoft Corporation. All rights reserved.
'==========================================================
'Muslim Qibla Direction Compass, Great Circle Distance
'and Great Circle Direction Calculator.

'Qibla is an Arabic word referring to the direction that should be
'faced when a Muslim prays. This program calculates the Qibla direction
'from any point on the Earth. It also uses and claculates
'the Great Circle Distance and the Great Circle Direction.
'==========================================================
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    
    Dim Latitude As Single
    Dim Longitude  As Single
    Dim QiblaDir  As Single
    
    'Cairo: Lat=30.1, Long=31.3
    Latitude = 30.1
    Longitude = 31.3
    
    ' Calculates the direction of the Qibla from any point on
    ' the Earth From North Clocklwise
    QiblaDir = QiblaDirection(Latitude, Longitude)
    
    cmdCalculate_Click
    
    lblQibla.Caption = QiblaDir
End Sub

Private Sub cmdCalculate_Click()
    
    Dim QiblaDir  As Single
    
    If txtLatitude.Text = "" Or txtLongitude.Text = "" Then Exit Sub
    
    QiblaDir = QiblaDirection(CSng(txtLatitude.Text), CSng(txtLongitude.Text))
    
    lblQibla.Caption = QiblaDir
    
    Dim W As Long, H As Long, R As Long
    Dim X2 As Long, Y2 As Long
        
    W = picCompass.ScaleWidth \ 2
    H = picCompass.ScaleHeight \ 2
    Line1.X1 = W
    Line1.Y1 = H
    
    QiblaDir = QiblaDir - 90
    R = W * 0.71
    'X = R Cos(theta), Y = R Sin(theta)
    X2 = H + (R * Cos(QiblaDir * (3.14 / 180)))
    Y2 = W + (R * Sin(QiblaDir * (3.14 / 180)))
    Line1.X2 = X2
    Line1.Y2 = Y2
    
    ArrawHead.Move X2 - (ArrawHead.Width / 2), Y2 - (ArrawHead.Height / 2)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
End Sub

Private Sub txtLatitude_Change()
    cmdCalculate_Click
End Sub

Private Sub txtLongitude_Change()
    cmdCalculate_Click
End Sub
