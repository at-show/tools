'#################################################
'# Youtubeから指定したURLのオーディオを取得するスクリプト
'# "youtube-dl.exe","ffmpeg.exe","ffprobe.exe"必須
'# "target.txt"に対象URLを記載して同階層に保存すること
'#################################################
Option Explicit

Dim objShell, objFSO, objFile, prefix, iTunesFolder

'# iTunesフォルダ
iTunesFolder = "L:\data\MyMusic\iTunes Music\iTunes に自動的に追加\"

'# youtube-dlコマンド設定
'# オーディオのみ過去5日以内に登録されたものが対象
prefix = "youtube-dl.exe " &_
	"--no-check-certificate " &_
	"--download-archive Downloaded.txt " &_
	"-x --audio-format ""m4a"" " &_
	"--prefer-ffmpeg " &_
	"--audio-quality 0 " &_
	"--dateafter now-5day " &_
	"-o ""Files\%(title)s.%(ext)s"" "

'# ターゲットに記載されたURLをダウンロード
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("target.txt")
Do While objFile.AtEndOfStream <> True
	Dim command
	command = prefix & objFile.ReadLine
	Call objShell.Run(command, 1, true)
Loop
objFile.Close

'# 再生位置を記憶させたいのでオーディオブックに拡張子を変更（m4a→m4b）
Call objShell.Run("cmd /c ren Files\*.m4a *.m4b", 1, true)

'# iTunesに自動的に追加フォルダへ移動
Call objFSO.MoveFile("Files\*", iTunesFolder)


Set objShell = Nothing
Set objFile = Nothing
Set objFSO = Nothing
