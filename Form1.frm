VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   172
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4920
      Top             =   480
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   2175
      Top             =   420
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   1995
      Top             =   420
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   1800
      Top             =   420
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   2370
      Top             =   420
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rgnBasic As New Region
Dim rgnExtended As New Region
Dim CurrentRgn As Long
Dim pic(0 To 1) As New StdPicture

Private Sub Form_DblClick()
    Unload Me
    End
End Sub

Private Sub Form_Load()
    ' Load pictures from file
    Set pic(0) = LoadPicture(App.Path & "\Sonique Small Extended.bmp", 0, 0, 0, 0)
    Set pic(1) = LoadPicture(App.Path & "\Sonique Small.bmp", 0, 0, 0, 0)
    ' Scan Shape from Green Screen Style Image
    Call rgnExtended.ScanPicture(pic(0))
    Call rgnBasic.ScanPicture(pic(1))
    ' Offset the Shape to allow for the form header.
    Call rgnBasic.OffsetHeader(Me)
    Call rgnExtended.OffsetHeader(Me)
    
    Me.picture = pic(1) ' Set the Form Background
    Call rgnBasic.ApplyRgn(Me.hWnd) ' Set the Form Shape
    CurrentRgn = rgnBasic.hndRegion ' Set the Current Shape
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ExtendView
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rgnExtended = Nothing
    Set rgnBasic = Nothing
End Sub

Private Sub Image1_Click()
    Unload Me
    End
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ExtendView
End Sub

Private Sub Image2_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ExtendView
End Sub

Private Sub Image3_Click()
    Unload Me
    Form2.Show
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ExtendView
End Sub

Private Sub Image4_Click()
    Unload Me
    Form3.Show
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ExtendView
End Sub

Private Sub Timer1_Timer()
    If Me.WindowState = vbMinimized Then Exit Sub
    If CurrentRgn <> rgnBasic.hndRegion Then ' If it is not already the Current Shape
        Me.picture = pic(1) ' Set the Form Background
        Call rgnBasic.ApplyRgn(Me.hWnd) ' Set the Form Shape
        CurrentRgn = rgnBasic.hndRegion ' Set the Current Shape
    End If
    Timer1.Enabled = False
End Sub

Private Sub ExtendView()
    If Me.WindowState = vbMinimized Then Exit Sub
    Timer1.Enabled = False
    If CurrentRgn <> rgnExtended.hndRegion Then ' If it is not already the Current Shape
        Me.picture = pic(0) ' Set the Form Background
        Call rgnExtended.ApplyRgn(Me.hWnd) ' Set the Form Shape
        CurrentRgn = rgnExtended.hndRegion ' Set the Current Shape
    End If
    Timer1.Enabled = True
End Sub
