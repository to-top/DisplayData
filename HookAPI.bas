Attribute VB_Name = "HookAPI"
Option Explicit

Public Type POINTL
    X As Long
    Y As Long
End Type

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
     
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, xyPoint As POINTL) As Long

'获取系统注册表中配置的caption和border高度象素值
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CYBORDER = 6
Public Const SM_CYCAPTION = 4

Public Const GWL_WNDPROC = -4
Public Const SPI_GETWHEELSCROLLLINES = 104
Public Const WM_MOUSEWHEEL = &H20A
Public WHEEL_SCROLL_LINES As Long
       
Global lpPrevWndProc As Long

'是否应用按钮的皮肤效果
Public Const gConApplayCmdSkin = 1
Public Const gConCmdSkinType = 1

Public Sub Hook(ByVal hWnd As Long)
    lpPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
    '获取"控制面板"中的滚动行数值
    Call SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, WHEEL_SCROLL_LINES, 0)
    If WHEEL_SCROLL_LINES > FrmShowData.CurrentDBGrid.VisibleRows Then
        WHEEL_SCROLL_LINES = FrmShowData.CurrentDBGrid.VisibleRows
    End If
End Sub

Public Sub UnHook(ByVal hWnd As Long)
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(hWnd, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, _
                    ByVal uMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long
    Dim pt As POINTL
    Dim wzDelta As Integer, wKeys As Integer
    
    Select Case uMsg
        Case WM_MOUSEWHEEL
            
            wzDelta = HIWORD(wParam)
            wKeys = LOWORD(wParam)
            pt.X = LOWORD(lParam)
            pt.Y = HIWORD(lParam)
            '将屏幕坐标转换为frmshowdata.窗口坐标
            ScreenToClient FrmShowData.SSTabResults.hWnd, pt
            With FrmShowData.CurrentDBGrid

                '判断坐标是否在frmshowdata.CurrentDBGrid窗口内
                If pt.X > .Left / Screen.TwipsPerPixelX And pt.X < (.Left + .Width) / Screen.TwipsPerPixelX And _
                   pt.Y > .Top / Screen.TwipsPerPixelY And pt.Y < (.Top + .Height) / Screen.TwipsPerPixelY Then
                    
                    If wKeys = 16 Then
                    '滚动键按下，水平滚动CurrentDBGrid
                        If Sgn(wzDelta) = 1 Then
                            FrmShowData.CurrentDBGrid.Scroll -1, 0
                        Else
                            FrmShowData.CurrentDBGrid.Scroll 1, 0
                        End If
                    Else
                        '垂直滚动grdDataGrid
                        If Sgn(wzDelta) = 1 Then
                            FrmShowData.CurrentDBGrid.Scroll 0, 0 - WHEEL_SCROLL_LINES
                        Else
                            FrmShowData.CurrentDBGrid.Scroll 0, WHEEL_SCROLL_LINES
                        End If
                    End If
                End If
            End With
        Case Else
            WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
    End Select
End Function

Public Function HIWORD(LongIn As Long) As Integer
    ' 取出32位值的高16位
    HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

Public Function LOWORD(LongIn As Long) As Integer
    ' 取出32位值的低16位
    LOWORD = LongIn And &HFFFF&
End Function




