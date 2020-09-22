VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4425
   ClientLeft      =   4095
   ClientTop       =   2760
   ClientWidth     =   2715
   LinkTopic       =   "Form2"
   ScaleHeight     =   4425
   ScaleWidth      =   2715
   Begin VB.Image Image4 
      Height          =   135
      Left            =   1440
      Top             =   360
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   1590
      Top             =   450
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   1710
      Top             =   570
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   1800
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim pic(0) As StdPicture
    Dim rgnBasic As New Region
    
    ' Load pictures from file
    Set pic(0) = LoadPicture(App.Path & "\Sonique Medium.bmp", 0, 0, 0, 0)
    ' Scan Shape from Green Screen Style Image
    Call rgnBasic.ScanPicture(pic(0))
    ' Offset the Shape to allow for the form header.
    Call rgnBasic.OffsetHeader(Me)
    
    Me.picture = pic(0)
    Call rgnBasic.ApplyRgn(Me.hWnd)
    
    Set rgnBasic = Nothing
    Set pic(0) = Nothing
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub Image1_Click()
    Unload Me
    End
End Sub

Private Sub Image2_Click()
    Unload Me
    Form1.Show
End Sub

Private Sub Image3_Click()
    Unload Me
    Form3.Show
End Sub

Private Sub Image4_Click()
    Me.WindowState = vbMinimized
End Sub
