Option Explicit

Dim objShell
Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd /c Dir /b ""C:\ProgramData\Microsoft\Windows\Start Menu\Programs"" >C:\Users\****\Desktop\file.txt",0,false

Dim msg 
msg = "�f�X�N�g�b�v�ɒ��o���܂����B"
msgbox msg,,"�C���X�g�[��APL"