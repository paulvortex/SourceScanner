Attribute VB_Name = "OpenFolder32"
Public Const MAX_PATH = 260 '������������ ����� �������� � ������������ �����

Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib "Shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHSimpleIDListFromPath Lib "Shell32" Alias "#162" (ByVal szPath As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Function SelectFolder(Form1 As Form, tTitle As String, tStartDir As String) As String
    Dim bi As BROWSEINFO '��������� ��� SHBrowseForFolder
    Dim sPath As String '���������� ��� ��� �����
    Dim pID As Long '���������� ��� PIDL

    bi.hOwner = Form1.hwnd '��������� ����������� ���� ��� ����������� ����������� � ZOrder
    bi.pszDisplayName = String$(MAX_PATH, 0) '����� ��� ��� �����
    bi.lpszTitle = tTitle '��������� ����
    '������� ���������� �� ��� ��������� �����
    '� � PIDL (Point to ID list), �.�. ���������
    '� ��������� ������. ��� ������� �� ����. �������
    bi.pidlRoot = SHSimpleIDListFromPath(tStartDir) '���� � ������ ���� � ����
    '���� �������� ��� ���� ������, �� ��������� ����� "������� ����"
    '�������� ��������� ���������� � SHBrowseForFolder �
    '�� �� �������� ��������� ������������� �����
    pID = SHBrowseForFolder(bi)
    '������� ����� ��� ��� ������������ �����
    sPath = String$(MAX_PATH, 0)
    If SHGetPathFromIDList(ByVal pID, ByVal sPath) Then
        '�������� ��� ����� � ���������� ���, �.�. �������� ����
        'sPath = StrZToStr(sPath)
    End If
    '����������� ��������� ������
    Call CoTaskMemFree(pID)
    SelectFolder = sPath '���������� ��� ��������� �����
End Function
