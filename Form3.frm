VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6330
   ClientLeft      =   4905
   ClientTop       =   3525
   ClientWidth     =   6270
   LinkTopic       =   "Form3"
   ScaleHeight     =   6330
   ScaleWidth      =   6270
   Begin VB.Image Image6 
      Height          =   135
      Left            =   4700
      Top             =   570
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   4520
      Top             =   570
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   4330
      Top             =   570
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   4890
      Top             =   570
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   360
      Top             =   5640
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   360
      Top             =   4560
      Width           =   5295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rgnBasic As New Region
Dim rgnExtended As New Region
Dim CurrentRgn As Long
Dim pic(0 To 1) As New StdPicture

Private Sub Form_Load()
    ' Load pictures from file
    Set pic(0) = LoadPicture(App.Path & "\Sonique Large Extended.bmp", 0, 0, 0, 0)
    Set pic(1) = LoadPicture(App.Path & "\Sonique Large.bmp", 0, 0, 0, 0)
    ' Scan Shape from Green Screen Style Image
    Call rgnExtended.ScanPicture(pic(0))
    Call rgnBasic.ScanPicture(pic(1))
    ' Offset the Shape to allow for the form header.
    Call rgnBasic.OffsetHeader(Me)
    Call rgnExtended.OffsetHeader(Me)
    
    Me.picture = pic(1)
    Call rgnBasic.ApplyRgn(Me.hWnd)
    'CurrentRgn = rgnBasic.hndRegion
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Public Sub RestoreView()
    If Me.WindowState = vbMinimized Then Exit Sub
    Me.picture = pic(1)
    Call rgnBasic.ApplyRgn(Me.hWnd)
    'CurrentRgn = rgnBasic.hndRegion
End Sub

Private Sub ExtendView()
    If Me.WindowState = vbMinimized Then Exit Sub
    Me.picture = pic(0)
    Call rgnExtended.ApplyRgn(Me.hWnd)
    'CurrentRgn = rgnExtended.hndRegion
End Sub

Private Sub Image1_Click()
    Call ExtendView
End Sub

Private Sub Image2_Click()
    Call RestoreView
End Sub

Private Sub Image3_Click()
    Unload Me
    End
End Sub

Private Sub Image4_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub Image5_Click()
    Unload Me
    Form2.Show
End Sub

Private Sub Image6_Click()
    Unload Me
    Form1.Show
End Sub
