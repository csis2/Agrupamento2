#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; unzipping_srag versão 1.0 - CSIS Software - 03/03/2021.
; Descompacta arquivos enfileirados na pasta c:\agrupamento2\data.

GLOBAL argui2 := 0 ; armazena o argumento escolhido pelo usuário.
GLOBAL nFiles := 0 ; armazena a quantidade de arquivos gerados após a descompactacao.

; Testa se o script está rodando com argumento.
if 0 < 1  
{
ExitApp ; se o argumento não foi encontrado, finaliza o script.
}

loop
{
	switch := %a_index%
	if switch =
	{
	ExitApp ; se o argumento não foi encontrado, finaliza o script.
	}
	If switch = run
	{
     	argui2 = 1 ; atribui o valor 1 se o argumento passado for run.
		FileDelete, c:\agrupamento2\data\*.dbf
		Sleep, 3000
		break
	}
}

if (argui2 == 1 and !FileExist("c:\agrupamento2\data\*.zip")) ; procura por arquivos compactados na subpasta 'data' do agrupamento2.
{
ExitApp ; 
}

FileList := []

if (argui2 == 1) ; nesse argumento, coloca no array FileList os arquivos na pasta c:\agrupamento2\data.
{
Loop, Files, c:\agrupamento2\data\*.zip
{
FileList.Insert(A_LoopFileFullPath)
}
}

FileList.MaxIndex()

if (argui2 == 1) ; nesse argumento, os arquivos serão descompactados na pasta c:\agrupamento2\data.
Output_Path := "c:\agrupamento2\data\"

For index, value in FileList
{
Compressed_File := % FileList[A_Index]

Sleep 2000

RunWait %comspec% /c c:\agrupamento2\bin\7za e %Compressed_File% -o%Output_Path% -y, , Hide
}

Sleep 2000
