'#################################################
'# Youtube����w�肵��URL�̃I�[�f�B�I���擾����X�N���v�g
'# "youtube-dl.exe","ffmpeg.exe","ffprobe.exe"�K�{
'# "target.txt"�ɑΏ�URL���L�ڂ��ē��K�w�ɕۑ����邱��
'#################################################
Option Explicit

'# iTunes�t�H���_
Const iTunesFolder = "L:\data\MyMusic\iTunes Music\iTunes �Ɏ����I�ɒǉ�\"

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
On Error Resume Next
Call objFSO.MoveFile("Files\*", iTunesFolder)
On Error GoTo 0


Set objShell = Nothing
Set objFile = Nothing
Set objFSO = Nothing
