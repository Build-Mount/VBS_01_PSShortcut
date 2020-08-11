'PowerShell Script���h���b�O���h���b�v���ċN�����邽�߂̃V���[�g�J�b�g�쐬

Const C_LINK = ".lnk"
Const C_PS = "powershell.exe"
Const C_PS_ARGS = "-NoProfile -ExecutionPolicy Unrestricted -File "

Call Main()

'�{Script�̃��C��
Public Function Main()
	Dim strFileNameWPath

	Set objWS = WScript.CreateObject("WScript.Shell")
	strFileNameWPath = GetArguments			'�h���b�O���h���b�v�����t�@�C���̃p�X�擾
	Call CreatePSShortcut(strFileNameWPath)	'PowerShell�����V���[�g�J�b�g����
End Function

'�h���b�O���h���b�v���������p�X���܂ރt�@�C�������擾����B
'�h���b�O���h���b�v���Ă��Ȃ��A�������͂Q�ȏ�̃t�@�C�����h���b�v�����ꍇ�ɂ̓G���[�ŉ�������B
'
'@return string �h���b�O���h���b�v�����t�@�C���̃p�X�B
Private Function GetArguments()
	Dim i				'�h���b�O���h���b�v�����t�@�C���̃p�X���m�F���郋�[�v�p�J�E���^
	Dim objFSO

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	If WScript.Arguments.Count = 0 Then
		MsgBox("�Ώۂ̃t�@�C�����h���b�O���h���b�v���Ă��������B")
		WScript.Quit()
	End If

	If WScript.Arguments.Count > 1 Then
		MsgBox("�h���b�O���h���b�v�t�@�C���͂P�����ɂ��Ă��������B")
		WScript.Quit()
	End If

	If objFSO.FolderExists(WScript.Arguments.Item(0)) Then
		MsgBox("�t�H���_�ł͂Ȃ��t�@�C�����h���b�O���h���b�v���Ă��������B")
		WScript.Quit()
	End If

	GetArguments = WScript.Arguments.Item(0)
End Function

'strFileNameWPath�����Ƃ�PowerShell�̃V���[�g�J�b�g���AstrFileNameWPath�̂���t�H���_�ɐ�������B
'
'@strFileNameWPath string �p�X���܂ރt�@�C�����B
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