VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'-----------------------
' auth : t.ichi
'-----------------------

Private mFullname As String
Private mBaseDir As String
Private mName As String
Private mExtension As String

'--- 初期化・コンストラクタとして使用
Public Function Init(ByVal FullName As String) As Boolean
    Init = False
    mFullname = FullName
    mBaseDir = ""
    mName = ""
    mExtension = ""
    
    If mFullname = "" Then: Exit Function
    
    Dim tmp() As String

    tmp = Split(mFullname, "\")
    If UBound(tmp) > 0 Then
        mName = tmp(UBound(tmp))
        mBaseDir = Left(mFullname, Len(mFullname) - Len(mName) - 1)
    Else
        mName = mFullname
    End If
    
    tmp = Split(mName, ".")
    If UBound(tmp) > 0 Then
        mExtension = tmp(UBound(tmp))
        mName = Left(mName, Len(mName) - Len(mExtension) - 1)
    Else
        mExtension = ""
    End If
    
    Init = True
End Function


'//////////////////////////////////////

'--- 親フォルダPath
Public Function BaseDirectory(Optional ByVal needDirectorySeparator As Boolean = False) As String
    BaseDirectory = mBaseDir & IIf(mBaseDir <> "" And needDirectorySeparator, "\", "")
End Function

'--- 親フォルダ DirectoryInfo
Public Function BaseDirectoryInfo() As DirectoryInfo
    Dim d As New DirectoryInfo
    Call d.Init(mBaseDir)
    Set BaseDirectoryInfo = d
End Function

'--- フルPath
Public Function FullName() As String
    FullName = mFullname
End Function


'--- ファイル名
Public Function name(Optional ByVal needExtension As Boolean = True) As String
    If needExtension Then
        name = mName & Extension(True)
    Else
        name = mName
    End If
End Function

'--- 拡張子
Public Function Extension(Optional ByVal needDot As Boolean = True) As String
    Extension = IIf(needDot, ".", "") & mExtension
End Function


'//////////////////////////////////////
'--- 存在判定
Public Function IsExists() As Boolean
On Error GoTo ex:
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    IsExists = fs.FileExists(mFullname)
ex:
    Set fs = Nothing
End Function

'--- ファイル削除
Public Function Delete() As Boolean
On Error GoTo ex:
    Delete = False
    If Not IsExists Then: Exit Function
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Call fs.DeleteFile(mFullname, True)
    Delete = True
ex:
    Set fs = Nothing
End Function

'--- ファイルコピー
Public Function CopyTo(ByVal path As String, Optional overwrite As Boolean = False) As Boolean
On Error GoTo ex:
    CopyTo = False
    If Not IsExists Then: Exit Function
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Call fs.CopyFile(mFullname, path, overwrite)
    CopyTo = True
ex:
End Function

'--- ファイル移動
Public Function MoveTo(path) As Boolean
On Error GoTo ex:
    MoveTo = False
    If Not IsExists Then: Exit Function
    If mName = path Then: Exit Function
    Name mFullname As path
    MoveTo = Me.Init(path)
ex:
End Function

'--- ファイル名変更（ディレクトリはそのまま）
Public Function Rename(name) As Boolean
On Error GoTo ex:
    Rename = False
    If Not IsExists Then: Exit Function
    
    Dim newPath As New FileInfo
    If Not newPath.Init(mBaseDir & "\" & name) Then: Exit Function
    If newPath.IsExists Then: Exit Function
    Name mFullname As newPath.FullName
    Rename = Me.Init(newPath.FullName)
ex:
End Function


'--- ファイル取得
Public Function ReadText(Optional ByVal utf8 As Boolean = True) As String
On Error GoTo ex:
    ReadText = ""
    If Not IsExists Then: Exit Function
    
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    s.Type = 2 ' 1:バイナリ / 2:TEXT
    s.Charset = IIf(utf8, "UTF-8", "Shift-JIS")
    s.Open
    
    s.LoadFromFile = mFullname
    
    ReadText = s.ReadText
ex:
    If Not s Is Nothing Then: s.Close
    Set s = Nothing
End Function

'--- ファイル書き込み
Public Function WriteText(ByVal text As String, Optional ByVal utf8 As Boolean = True, Optional ByVal bom As Boolean = False) As Boolean
    Dim data As New Collection
    Call data.Add(text)
    WriteText = FileWriter(data, False, False, utf8, bom)
    Set data = Nothing
End Function
Public Function WriteLine(ByVal line As String, Optional ByVal utf8 As Boolean = True, Optional ByVal bom As Boolean = False) As Boolean
    Dim data As New Collection
    Call data.Add(line)
    WriteLine = FileWriter(data, False, True, utf8, bom)
    Set data = Nothing
End Function
Public Function WriteLines(ByVal data As Collection, Optional ByVal utf8 As Boolean = True, Optional ByVal bom As Boolean = False) As Boolean
    WriteLines = FileWriter(data, False, True, utf8, bom)
End Function

'--- ファイル追加書き込み
Public Function AppendText(ByVal text As String, Optional ByVal utf8 As Boolean = True, Optional ByVal bom As Boolean = False) As Boolean
    Dim data As New Collection
    Call data.Add(text)
    AppendText = FileWriter(data, True, False, utf8, bom)
    Set data = Nothing
End Function
Public Function AppendLine(ByVal line As String, Optional ByVal utf8 As Boolean = True, Optional ByVal bom As Boolean = False) As Boolean
    Dim data As New Collection
    Call data.Add(line)
    AppendLine = FileWriter(data, True, True, utf8, bom)
    Set data = Nothing
End Function
Public Function AppendLines(ByVal data As Collection, Optional ByVal utf8 As Boolean = True, Optional ByVal bom As Boolean = False) As Boolean
    AppendLines = FileWriter(data, True, True, utf8, bom)
End Function

Private Function FileWriter(ByVal data As Collection, ByVal isAppend As Boolean, ByVal isLine As Boolean, ByVal utf8 As Boolean, Optional ByVal bom As Boolean)
On Error GoTo ex:
    FileWriter = False
    Dim writeOpt As Integer: writeOpt = IIf(isLine, 1, 0) '後ろに改行をつけるか
    
    Dim posSize As Long
    posSize = 0
    
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    s.Type = 2 ' 1:バイナリ / 2:TEXT
    s.Charset = IIf(utf8, "UTF-8", "Shift-JIS")
    
    s.Open
    
    '追記の場合
    If isAppend Then
        s.LoadFromFile = mFullname
        s.Position = s.size
        If s.size > 0 Then
            posSize = s.size
        End If
    End If
    
    Dim v As Variant
    For Each v In data
        Call s.WriteText(CStr(v), writeOpt)
    Next
    
    'BOM消し
    If Not bom And utf8 And posSize = 0 Then
        s.Position = 0
        s.Type = 1 ' 1:バイナリ / 2:TEXT
        s.Position = 3
        
        Dim bin: bin = s.Read
        s.Close
        
        s.Open
        s.Write bin
    End If
    
    Call s.SaveToFile(mFullname, 2) '2: 新規作成 & 上書き
    
    FileWriter = True
ex:
    If Not s Is Nothing Then: s.Close
    Set s = Nothing
End Function
