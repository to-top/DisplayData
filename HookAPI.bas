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

'��ȡϵͳע��������õ�caption��border�߶�����ֵ
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CYBORDER = 6
Public Const SM_CYCAPTION = 4

Public Const GWL_WNDPROC = -4
Public Const SPI_GETWHEELSCROLLLINES = 104
Public Const WM_MOUSEWHEEL = &H20A
Public WHEEL_SCROLL_LINES As Long
       
Global lpPrevWndProc As Long

'�Ƿ�Ӧ�ð�ť��Ƥ��Ч��
Public Const gConApplayCmdSkin = 1
Public Const gConCmdSkinType = 1

Public Sub Hook(ByVal hWnd As Long)
    lpPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
    '��ȡ"�������"�еĹ�������ֵ
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
            '����Ļ����ת��Ϊfrmshowdata.��������
            ScreenToClient FrmShowData.SSTabResults.hWnd, pt
            With FrmShowData.CurrentDBGrid

                '�ж������Ƿ���frmshowdata.CurrentDBGrid������
                If pt.X > .Left / Screen.TwipsPerPixelX And pt.X < (.Left + .Width) / Screen.TwipsPerPixelX And _
                   pt.Y > .Top / Screen.TwipsPerPixelY And pt.Y < (.Top + .Height) / Screen.TwipsPerPixelY Then
                    
                    If wKeys = 16 Then
                    '���������£�ˮƽ����CurrentDBGrid
                        If Sgn(wzDelta) = 1 Then
                            FrmShowData.CurrentDBGrid.Scroll -1, 0
                        Else
                            FrmShowData.CurrentDBGrid.Scroll 1, 0
                        End If
                    Else
                        '��ֱ����grdDataGrid
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
    ' ȡ��32λֵ�ĸ�16λ
    HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

Public Function LOWORD(LongIn As Long) As Integer
    ' ȡ��32λֵ�ĵ�16λ
    LOWORD = LongIn And &HFFFF&
End Function




