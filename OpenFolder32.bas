Attribute VB_Name = "OpenFolder32"
Public Const MAX_PATH = 260 'Максимальное число символов в наименовании папки

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
    Dim bi As BROWSEINFO 'структура для SHBrowseForFolder
    Dim sPath As String 'переменная под имя папки
    Dim pID As Long 'переменная для PIDL

    bi.hOwner = Form1.hwnd 'описатель вызывающего окна для правильного отображения в ZOrder
    bi.pszDisplayName = String$(MAX_PATH, 0) 'буфер под имя папки
    bi.lpszTitle = tTitle 'заголовак окна
    'функции передается не имя начальной папки
    'а её PIDL (Point to ID list), т.е. указатель
    'в системном списке. Его получам от спец. функции
    bi.pidlRoot = SHSimpleIDListFromPath(tStartDir) 'диск С должен быть у всех
    'если оставить это поле пустым, то начальная папка "Рабочий стол"
    'передаем начальную информацию в SHBrowseForFolder и
    'от неё получаем выбранную пользователем папку
    pID = SHBrowseForFolder(bi)
    'создаем буфер под имя возвращаемой папки
    sPath = String$(MAX_PATH, 0)
    If SHGetPathFromIDList(ByVal pID, ByVal sPath) Then
        'приводим имя папки в нормальный вид, т.е. отсекаем нули
        'sPath = StrZToStr(sPath)
    End If
    'обязательно почистите память
    Call CoTaskMemFree(pID)
    SelectFolder = sPath 'возвращаем имя выбранной папки
End Function
