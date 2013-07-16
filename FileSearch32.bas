Attribute VB_Name = "FileSearch32"
 Declare Function FindFirstFile Lib "kernel32" Alias _
   "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
   As WIN32_FIND_DATA) As Long

   Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
   (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

   Declare Function GetFileAttributes Lib "kernel32" Alias _
   "GetFileAttributesA" (ByVal lpFileName As String) As Long

   Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) _
   As Long

   Declare Function FileTimeToLocalFileTime Lib "kernel32" _
   (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
     
   Declare Function FileTimeToSystemTime Lib "kernel32" _
   (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

   Public Const MAX_PATH = 260
   Public Const MAXDWORD = &HFFFF
   Public Const INVALID_HANDLE_VALUE = -1
   Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
   Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
   Public Const FILE_ATTRIBUTE_HIDDEN = &H2
   Public Const FILE_ATTRIBUTE_NORMAL = &H80
   Public Const FILE_ATTRIBUTE_READONLY = &H1
   Public Const FILE_ATTRIBUTE_SYSTEM = &H4
   Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

   Type FILETIME
     dwLowDateTime As Long
     dwHighDateTime As Long
   End Type

   Type WIN32_FIND_DATA
     dwFileAttributes As Long
     ftCreationTime As FILETIME
     ftLastAccessTime As FILETIME
     ftLastWriteTime As FILETIME
     nFileSizeHigh As Long
     nFileSizeLow As Long
     dwReserved0 As Long
     dwReserved1 As Long
     cFileName As String * MAX_PATH
     cAlternate As String * 14
   End Type

   Type SYSTEMTIME
     wYear As Integer
     wMonth As Integer
     wDayOfWeek As Integer
     wDay As Integer
     wHour As Integer
     wMinute As Integer
     wSecond As Integer
     wMilliseconds As Integer
   End Type

   Public Function StripNulls(OriginalStr As String) As String
      If (InStr(OriginalStr, Chr(0)) > 0) Then
         OriginalStr = Left(OriginalStr, _
          InStr(OriginalStr, Chr(0)) - 1)
      End If
      StripNulls = OriginalStr
   End Function
   
   
'
' API
'

Function FindFiles(tList As ListBox, path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)
   Dim FileName As String   ' Walking filename variable...
   Dim DirName As String    ' SubDirectory Name
   Dim dirNames() As String ' Buffer for directory name entries
   Dim nDir As Integer   ' Number of directories in this path
   Dim i As Integer      ' For-loop counter...
   Dim hSearch As Long   ' Search Handle
   Dim WFD As WIN32_FIND_DATA
   Dim Cont As Integer
   Dim FT As FILETIME
   Dim ST As SYSTEMTIME
   Dim DateCStr As String, DateMStr As String
     
   If Right(path, 1) <> "\" Then path = path & "\"
   ' Search for subdirectories.
   nDir = 0
   ReDim dirNames(nDir)
   Cont = True
   hSearch = FindFirstFile(path & "*", WFD)
   If hSearch <> INVALID_HANDLE_VALUE Then
      Do While Cont
         DirName = StripNulls(WFD.cFileName)
         ' Ignore the current and encompassing directories.
         If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(path & DirName) And _
             FILE_ATTRIBUTE_DIRECTORY Then
               dirNames(nDir) = DirName
               DirCount = DirCount + 1
               nDir = nDir + 1
               ReDim Preserve dirNames(nDir)
               ' Uncomment the next line to list directories
               'List1.AddItem path & FileName
            End If
         End If
         Cont = FindNextFile(hSearch, WFD)  ' Get next subdirectory.
      Loop
      Cont = FindClose(hSearch)
   End If

   ' Walk through this directory and sum file sizes.
   hSearch = FindFirstFile(path & SearchStr, WFD)
   Cont = True
   If hSearch <> INVALID_HANDLE_VALUE Then
      While Cont
         FileName = StripNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") And _
              ((GetFileAttributes(path & FileName) And _
               FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
            FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * _
             MAXDWORD) + WFD.nFileSizeLow
            FileCount = FileCount + 1
            ' To list files w/o dates, uncomment the next line
            ' and remove or Comment the lines down to End If
            tList.AddItem path & FileName
   
           ' Include Creation date...
           ' FileTimeToLocalFileTime WFD.ftCreationTime, FT
           'FileTimeToSystemTime FT, ST
           'DateCStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
           '   " " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
           ' and Last Modified Date
           'FileTimeToLocalFileTime WFD.ftLastWriteTime, FT
          ' FileTimeToSystemTime FT, ST
          ' DateMStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
           '   " " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
          ' List1.AddItem path & FileName & vbTab & _
           '   Format(DateCStr, "mm/dd/yyyy hh:nn:ss") _
          '    & vbTab & Format(DateMStr, "mm/dd/yyyy hh:nn:ss")
          End If
         Cont = FindNextFile(hSearch, WFD)  ' Get next file
      Wend
      Cont = FindClose(hSearch)
   End If

   ' If there are sub-directories...
    If nDir > 0 Then
      ' Recursively walk into them...
      For i = 0 To nDir - 1
        FindFiles = FindFiles + FindFiles(tList, path & dirNames(i) _
         & "\", SearchStr, FileCount, DirCount)
      Next i
   End If
End Function

