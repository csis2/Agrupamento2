PROCEDURE Main()

cls
set color to gr+/
@ 1,0 say "CSIS Software 2021 - 2txt.exe"
@ 2,0 say "------------------------------------------------------------------------"
@ 3,0 say "Usa o arquivo 'file_manager.dbf' para gerar um arquivo txt"
@ 4,0 say "dos arquivos que serao processados pelo 'agrupamento2.exe'."
@ 5,0 say "Os arquivos DBF que não tem registros ficam de fora da listagem"
@ 6,0 say "no arquivo txt."
@ 7,0 say "------------------------------------------------------------------------"
@ 8,0 say "Compilado no Harbour Make (hbmk2) versao 3.0.0"
@ 9,0 say "Roda somente com argumento de entrada."
@ 10,0 say "Exemplo: classificador.exe run"
@ 11,0 say "------------------------------------------------------------------------"
set color to g+/

* Verifica se existe argumento.
if empty(HB_ArgV(1)) = .T.
set color to r+/
? "Erro! O programa precisa que voce forneca argumento de entrada."
__Quit()
endif

* Verifica se o argumento é valido.
set exact on
if (HB_ArgV(1)) <> "run"
set color to r+/
? "Erro! O argumento escolhido nao e valido."
__Quit()
endif
set exact off

use "c:\agrupamento2\tmp\file_manager.dbf"
copy FIELDS filename to "c:\agrupamento2\tmp\DBF_files" FOR processa <> "N" DELIMITED WITH BLANK

return nil