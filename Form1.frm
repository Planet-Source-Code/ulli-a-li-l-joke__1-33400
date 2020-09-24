VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'Kein
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   221
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   2  'Maximiert
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = Asc("Q") Then
        Unload Me
        Do Until ShowCursor(True) = 0
        Loop
    End If

End Sub

Private Sub Form_Resize()

  Dim DesktopDC As Long

    DesktopDC = GetDC(0)
    Debug.Print StretchBlt(hDC, ScaleWidth, ScaleHeight, -ScaleWidth, -ScaleHeight, DesktopDC, 0, 0, ScaleWidth, ScaleHeight, SRCCOPY)
    ReleaseDC 0, DesktopDC
    Do While ShowCursor(False) = 0
    Loop

End Sub

':) Ulli's VB Code Formatter V2.11.3 (03.04.2002 23:29:04) 7 + 24 = 31 Lines
