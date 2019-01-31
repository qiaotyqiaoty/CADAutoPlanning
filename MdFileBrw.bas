Option Explicit

Public Type BROWSEINFO
	hOwner As Long
	pidlRoot As Long
	pszDisplayName As String
	lpszTitle As String
	ulFlags As Long
	lpfn As Long
	lParam As Long
	iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const BIF_NEWDIALOGSTYLE = &H40
Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

'---------------------
'确保返回后面带反斜杠
Public Function EnsurePath(ByVal sPath As String) As String
	If Right(sPath,1) <> "\" Then
		EnsurePath = sPath & "\"
	Else
		EnsurePath = sPath
	End If
End Function

'主函数
Public Function GetFolder(ByVal sTitle As String) As String
	Dim bInf As BROWSEINFO
	Dim retval As Long
	Dim PathID As Long
	Dim RetPath As String
	Dim Offset As Integer
	
	bInf.lpszTitle = sTitle
	bInf.ulFlags = BIF_NEWDIALOGSTYLE
	PathID = SHBrowseForFolder(bInf)
	RetPath = Space$(512)
	
	retval = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
	
	If retval Then
		Offset = InStr(RetPath, Chr$(0))
		GetFolder = Left$(RetPath, Offset - 1)
	End If
End Function

Public Sub ListFilesFSO(ByVal sPath As String)
	Dim oFSO As Object
	Dim oFolder As Object
	Dim oFile As Object
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oFolder = oFSO.GetFolder(sPath)
	
	For Each oFile In oFolder.Files
		Debug.Print oFile.Name
	Next 'oFile

	Set oFile = Nothing
	Set oFolder = Nothing
	Set oFSO = Nothing
End Sub