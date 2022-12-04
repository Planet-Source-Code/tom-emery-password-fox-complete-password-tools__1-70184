Attribute VB_Name = "modForm"
Option Explicit

Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Function FormRegion(ByRef frm As Form, ByVal TrnsColor As Long) As Long

     Dim ScaleSize As Long
     Dim Width, Height As Long
     Dim rgnMain As Long
     Dim X, Y As Long
     Dim rgnPixel As Long
     Dim RGBColor As Long
     Dim dcMain As Long
     Dim bmpMain As Long
     
     ScaleSize = frm.ScaleMode
     frm.ScaleMode = 3
     Width = frm.ScaleX(frm.Picture.Width, vbHimetric, vbPixels)
     Height = frm.ScaleY(frm.Picture.Height, vbHimetric, vbPixels)
     frm.Width = Width * Screen.TwipsPerPixelX
     frm.Height = Height * Screen.TwipsPerPixelY
     rgnMain = CreateRectRgn(0, 0, Width, Height)
     dcMain = CreateCompatibleDC(frm.hDC)
     bmpMain = SelectObject(dcMain, frm.Picture.Handle)
     
     For Y = 0 To Height
        For X = 0 To Width
            RGBColor = GetPixel(dcMain, X, Y)
            If RGBColor = TrnsColor Then
                rgnPixel = CreateRectRgn(X, Y, X + 1, Y + 1)
                CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR
                DeleteObject rgnPixel
            End If
        Next X
    Next Y
    
    SelectObject dcMain, bmpMain
    DeleteDC dcMain
    DeleteObject bmpMain
    
    If rgnMain <> 0 Then
        SetWindowRgn frm.hwnd, rgnMain, True
        FormRegion = rgnMain
    End If
    
    frm.ScaleMode = ScaleSize
End Function


Public Sub FormMove(TheHwnd As Long)
    'Drag the form with the mouse
    ReleaseCapture
    SendMessage TheHwnd, &HA1, 2, 0&
End Sub
