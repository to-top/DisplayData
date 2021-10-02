Attribute VB_Name = "CaretPosition"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const EM_GETSEL = &HB0
Const EM_LINEFROMCHAR = &HC9
Const EM_LINEINDEX = &HBB

Public Sub GetCaretPos(ByVal TextHwnd As Long, LineNo As Long, ColNo As Long)
    Dim i As Long, j As Long
    Dim lParam As Long, wParam As Long
    Dim k As Long

    '�������ı��򴫵�EM_GETSEL��Ϣ�Ի�ȡ����ʼλ�õ�
    '�������λ�õ��ַ���
    'ע�ͣ�ȡ��ĿǰCaret����ǰ���ж��ٸ�byte
    i = SendMessage(TextHwnd, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16

    '�����ı��򴫵�EM_LINEFROMCHAR��Ϣ���ݻ�õ��ַ�
    '��ȷ������Ի�ȡ��������
    'ע�ͣ�ȡ��ǰ���ж�����
    LineNo = SendMessage(TextHwnd, EM_LINEFROMCHAR, j, 0)
    LineNo = LineNo + 1

    '���ı��򴫵�EM_LINEINDEX��Ϣ�Ի�ȡ��������
    'ע��: ȡ��Ŀǰcaret������ǰ���ж��ٸ�byte
    'k = SendMessage(TextHwnd, EM_LINEINDEX, -1, 0)
    'ColNo = j - k + 1
End Sub





