Attribute VB_Name = "CaretPosition"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const EM_GETSEL = &HB0
Const EM_LINEFROMCHAR = &HC9
Const EM_LINEINDEX = &HBB

Public Sub GetCaretPos(ByVal TextHwnd As Long, LineNo As Long, ColNo As Long)
    Dim i As Long, j As Long
    Dim lParam As Long, wParam As Long
    Dim k As Long

    '首先向文本框传递EM_GETSEL消息以获取从起始位置到
    '光标所在位置的字符数
    '注释：取得目前Caret所在前面有多少个byte
    i = SendMessage(TextHwnd, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16

    '再向文本框传递EM_LINEFROMCHAR消息根据获得的字符
    '数确定光标以获取所在行数
    '注释：取得前面有多少行
    LineNo = SendMessage(TextHwnd, EM_LINEFROMCHAR, j, 0)
    LineNo = LineNo + 1

    '向文本框传递EM_LINEINDEX消息以获取所在列数
    '注释: 取得目前caret所在行前面有多少个byte
    'k = SendMessage(TextHwnd, EM_LINEINDEX, -1, 0)
    'ColNo = j - k + 1
End Sub





