'PowerShell Scriptをドラッグ＆ドロップして起動するためのショートカット作成

Const C_LINK = ".lnk"
Const C_PS = "powershell.exe"
Const C_PS_ARGS = "-NoProfile -ExecutionPolicy Unrestricted -File "

Call Main()

'本Scriptのメイン
Public Function Main()
	Dim strFileNameWPath

	Set objWS = WScript.CreateObject("WScript.Shell")
	strFileNameWPath = GetArguments			'ドラッグ＆ドロップしたファイルのパス取得
	Call CreatePSShortcut(strFileNameWPath)	'PowerShellおｎショートカット生成
End Function

'ドラッグ＆ドロップしたしたパスを含むファイル名を取得する。
'ドラッグ＆ドロップしていない、もしくは２つ以上のファイルをドロップした場合にはエラーで応答する。
'
'@return string ドラッグ＆ドロップしたファイルのパス。
Private Function GetArguments()
	Dim i				'ドラッグ＆ドロップしたファイルのパスを確認するループ用カウンタ
	Dim objFSO

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	If WScript.Arguments.Count = 0 Then
		MsgBox("対象のファイルをドラッグ＆ドロップしてください。")
		WScript.Quit()
	End If

	If WScript.Arguments.Count > 1 Then
		MsgBox("ドラッグ＆ドロップファイルは１つだけにしてください。")
		WScript.Quit()
	End If

	If objFSO.FolderExists(WScript.Arguments.Item(0)) Then
		MsgBox("フォルダではなくファイルをドラッグ＆ドロップしてください。")
		WScript.Quit()
	End If

	GetArguments = WScript.Arguments.Item(0)
End Function

'strFileNameWPathをもとにPowerShellのショートカットを、strFileNameWPathのあるフォルダに生成する。
'
'@strFileNameWPath string パスを含むファイル名。
Private Function CreatePSShortcut(strFileNameWPath)
	Dim objFSO, objWS
	Dim objLink
	Dim strFileLink
	Dim strLinkPath

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set objWS = WScript.CreateObject("WScript.Shell")

	strFileLink = objFSO.GetParentFolderName(strFileNameWPath) + "\" _
				+ objFSO.GetBaseName(strFileNameWPath) _
				+ C_LINK
	Set objLink = objWS.CreateShortcut(strFileLink)
	objLink.TargetPath = C_PS
	objLink.Arguments = C_PS_ARGS + strFileNameWPath
	objLink.Save
End Function