Option Explicit

Dim objShell, objFSO, objFile, prefix

'# youtube-dl�R�}���h�ݒ�
'# �I�[�f�B�I�̂݉ߋ�5���ȓ��ɓo�^���ꂽ���̂��Ώ�
prefix = "youtube-dl.exe " &_
		"--no-check-certificate " &_
		"--download-archive Downloaded.txt " &_
		"-x --audio-format ""m4a"" " &_
		"--prefer-ffmpeg " &_
		"--audio-quality 0 " &_
		"--dateafter now-5day " &_
		"-o ""Files\%(title)s.%(ext)s"" "

'# �^�[�Q�b�g�ɋL�ڂ��ꂽURL���_�E�����[�h
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("target.txt")

Do While objFile.AtEndOfStream <> True
	Dim command
	command = prefix & objFile.ReadLine
	Call objShell.Run(command, 1, true)
Loop
objFile.Close

'# �Đ��ʒu���L�����������̂ŃI�[�f�B�I�u�b�N�Ɋg���q��ύX�im4a��m4b�j
Call objShell.Run("cmd /c ren Files\*.m4a *.m4b", 1, true)

'# iTunes�Ɏ����I�ɒǉ��t�H���_�ֈړ�
Call objFSO.MoveFile("Files\*", "L:\data\MyMusic\iTunes Music\iTunes �Ɏ����I�ɒǉ�\")


Set objShell = Nothing
Set objFile = Nothing
Set objFSO = Nothing
