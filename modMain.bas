Attribute VB_Name = "modMain"
Option Explicit
'declares
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Const RGN_COPY = 5
Public Const RGN_AND = 1
Public Const RGN_XOR = 3
Public Const RGN_OR = 2
Public Const RGN_MIN = RGN_AND
Public Const RGN_MAX = RGN_COPY
Public Const RGN_DIFF = 4
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOP = 0

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Sub Main()
Dim lRgn As Long

    'loading form
    Load frmMain
    'loading bitmap and calculate region
    frmMain.picMain.Picture = LoadResPicture(1, vbResBitmap)
    lRgn = lGetRegion(frmMain.picMain, RGB(255, 0, 255))
    'attach region to window and then delete this region
    SetWindowRgn frmMain.hwnd, lRgn, True
    DeleteObject lRgn
    frmMain.Show
    'always on top
    SetFormPosition frmMain.hwnd, True
End Sub

'always on top
Public Sub SetFormPosition(hwnd As Long, TopPosition As Boolean)
    If TopPosition Then
         SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
     Else
         SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
     End If
End Sub

Public Function lGetRegion(pic As PictureBox, lBackColor As Long) As Long
Dim lRgn As Long
Dim lSkinRgn As Long
Dim lStart As Long
Dim lX As Long
Dim lY As Long
Dim lHeight As Long
Dim lWidth As Long
Dim ms As Long

'create blank region
lSkinRgn = CreateRectRgn(0, 0, 0, 0)

With pic
    'count size of bitmap
    lHeight = .Height / Screen.TwipsPerPixelY
    lWidth = .Width / Screen.TwipsPerPixelX
    For lX = 0 To lHeight - 1
        lY = 0
        Do While lY < lWidth
            'find required pixel
            Do While lY < lWidth And GetPixel(.hDC, lY, lX) = lBackColor
                lY = lY + 1
            Loop

            If lY < lWidth Then
                lStart = lY
                Do While lY < lWidth And GetPixel(.hDC, lY, lX) <> lBackColor
                    lY = lY + 1
                Loop
                If lY > lWidth Then lY = lWidth
                'add required pixel to region
                lRgn = CreateRectRgn(lStart, lX, lY, lX + 1)
                CombineRgn lSkinRgn, lSkinRgn, lRgn, RGN_OR
                'delete unnecessary object
                DeleteObject lRgn
            End If
        Loop
    Next
End With
lGetRegion = lSkinRgn
End Function

