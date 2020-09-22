Attribute VB_Name = "Module1"
' Drag Form Declaration
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
'rounded
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


' MACOS Titlebar
Public Function CreateMacOSTitleBar(pict As PictureBox, title As String, frm As Form)
    pict.FontTransparent = False
    pict.AutoRedraw = True
    pict.ScaleMode = 3
    pict.BackColor = &HCCCCCC
    pict.BorderStyle = 0
    pict.ForeColor = QBColor(0)
    pict.Font = "Tahoma"
    pict.FontBold = False
    pict.FontSize = 10
    pict.Left = 100: pict.Top = 100: pict.Width = frm.Width - 200
    If (pict.ScaleWidth / 2) - (pict.TextWidth(title) / 2) <= 8 Then title = ""
    If title = "" Then
        lhs_left = 8
        lhs_right = pict.ScaleWidth - 8
        l_top = pict.ScaleHeight / 2 - 6
        dorhs = False
            dolhs = True
                GoTo drawit
            End If
            l_top = pict.ScaleHeight / 2 - 6
            lhs_left = 8
            sc = pict.ScaleWidth
            lhs_right = ((sc / 2) - (pict.TextWidth(title) / 2)) - 4
            lhs_right = Int(lhs_right)
            rhs_left = ((sc / 2) + (pict.TextWidth(title) / 2)) + 4
            rhs_left = Int(rhs_left)
            rhs_right = pict.ScaleWidth - 8
            dolhs = True
                dorhs = True
drawit:
                    If dolhs = True Then
                        For X = l_top To l_top + 10 Step 2
                            pict.Line (lhs_left - 1, X)-(lhs_right, X), &HFFFFFF
                            pict.Line (lhs_left, X + 1.5)-(lhs_right + 1, X + 1.5), &H666666
                        Next X
                    End If
                    If dorhs = True Then
                        For X = l_top To l_top + 10 Step 2
                            pict.Line (rhs_left - 1, X)-(rhs_right, X), &HFFFFFF
                            pict.Line (rhs_left, X + 1.5)-(rhs_right + 1, X + 1.5), &H666666
                        Next X
                    End If
                    pict.Line (0, pict.ScaleHeight - 1)-(pict.ScaleWidth, pict.ScaleHeight - 1), &H666666
                    maclefttext = (pict.ScaleWidth / 2) - (pict.TextWidth(title) / 2)
                    pict.CurrentX = maclefttext
                    mactoptext = (pict.ScaleHeight / 2) - (pict.TextHeight(title) / 2)
                    pict.CurrentY = mactoptext
                    pict.Print title
End Function
' 3D FORM SETTINGS
Function ColForm(Obj As PictureBox, r%, G%, B%, Step%)
    Dim R1%, G1%, B1%, R2%, G2%, B2%
    Obj.ScaleMode = 3
    Obj.AutoRedraw = True
    Obj.BorderStyle = 0
    Obj.BackColor = RGB(r%, G%, B%)
    R1% = r% + Step%: If R1% > 255 Then R1% = 255
    G1% = G% + Step%: If G1% > 255 Then G1% = 255
    B1% = B% + Step%: If B1% > 255 Then B1% = 255
    R2% = r% - Step%: If R2% < 0 Then R2% = 0
    G2% = G% - Step%: If G2% < 0 Then G2% = 0
    B2% = B% - Step%: If B2% < 0 Then B2% = 0
    Obj.Line (2, 2)-(Obj.ScaleWidth - 2, Obj.ScaleHeight - 2), RGB(R1%, G1%, B1%), B
    Obj.Line (Obj.ScaleWidth - 2, 2)-(Obj.ScaleWidth - 2, Obj.ScaleHeight - 1), RGB(R2%, G2%, B2%)
    Obj.Line (1, Obj.ScaleHeight - 2)-(Obj.ScaleWidth - 2, Obj.ScaleHeight - 2), RGB(R2%, G2%, B2%)
    Obj.Line (5, 5)-(Obj.ScaleWidth - 5, Obj.ScaleHeight - 5), RGB(R2%, G2%, B2%), B
    Obj.Line (Obj.ScaleWidth - 5, 6)-(Obj.ScaleWidth - 5, Obj.ScaleHeight - 4), RGB(R1%, G1%, B1%)
    Obj.Line (5, Obj.ScaleHeight - 5)-(Obj.ScaleWidth - 5, Obj.ScaleHeight - 5), RGB(R1%, G1%, B1%)
End Function

Function DragForm(frm As Form)
  Dim ret As Long
  ret = ReleaseCapture()
  ret = SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, 2&, 0&)
End Function


