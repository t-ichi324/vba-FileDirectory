VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DirectoryInfo"
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

'--- 初期化・コンストラクタとして使用
Public Function Init(ByVal FullName As String) As Boolean
    Init = False
    mFullname = FullName
    mBaseDir = ""
    mName = ""
    
    If mFullname = "" Then: Exit Function
    
    Dim tmp() As String

    tmp = Split(mFullname, "\")
    If UBound(tmp) > 0 Then
        mName = tmp(UBound(tmp))
        mBaseDir = Left(mFullname, Len(mFullname) - Len(mName) - 1)
    Else
        mName = mFullname
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

'--- フォルダ名
Public Function name() As String
    name = mName
End Function

'//////////////////////////////////////
'--- 存在判定
Public Function IsExists() As Boolean
On Error GoTo ex:
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    IsExists = fs.FolderExists(mFullname)
ex:
    Set fs = Nothing
End Function

'--- フォルダ削除
Public Function Delete() As Boolean
On Error GoTo ex:
    Delete = False
    If Not IsExists Then: Exit Function
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Call fs.DeleteFolder(mFullname, True)
    Delete = True
ex:
    Set fs = Nothing
End Function

'--- フォルダコピー
Public Function CopyTo(ByVal path As String, Optional overwrite As Boolean = False) As Boolean
On Error GoTo ex:
    CopyTo = False
    If Not IsExists Then: Exit Function
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Call fs.CopyFolder(mFullname, path, overwrite)
    CopyTo = True
ex:
End Function

'--- フォルダ移動
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
    
    Dim newPath As New DirectoryInfo
    If Not newPath.Init(mBaseDir & "\" & name) Then: Exit Function
    If newPath.IsExists Then: Exit Function
    Name mFullname As newPath.FullName
    Rename = Me.Init(newPath.FullName)
ex:
End Function
