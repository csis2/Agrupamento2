PROCEDURE Main()

cls
set color to bg+/
@ 1,0 say "CSIS Software 2021 - classificador.exe"
@ 2,0 say "------------------------------------------------------------------------"
@ 3,0 say "Verifica um a um os arquivos DBF criados pelo 'separador.exe'."
@ 4,0 say "Se possuirem registros invalidos, exclui esses registros."
@ 5,0 say "Classifica na tabela 'file_manager.dbf' se os arquivos vao ser processados"
@ 6,0 say "na proxima etapa ou nao."
@ 7,0 say "------------------------------------------------------------------------"
@ 8,0 say "Compilado no Harbour Make (hbmk2) versao 3.0.0"
@ 9,0 say "Roda somente com argumento de entrada."
@ 10,0 say "Exemplo: classificador.exe run"
@ 11,0 say "------------------------------------------------------------------------"
set color to w+/

* Verifica se existe argumento.
if empty(HB_ArgV(1)) = .T.
set color to r+/
@ 12,0 say "Erro! O programa precisa que voce forneca argumento de entrada."
__Quit()
endif

* Verifica se o argumento é valido.
set exact on
if (HB_ArgV(1)) <> "run"
set color to r+/
@ 12,0 say "Erro! O argumento escolhido nao e valido."
__Quit()
endif
set exact off

@ 12,0 say "Criando array com os registros de 'file_manager'."
aArray := {}
use "c:\agrupamento2\tmp\file_manager.dbf"
nRecs := reccount()
      FOR x := 1 TO nRecs
	  AAdd( aArray, {alltrim( filename )} )
	  skip	
	  NEXT
	  close

for y := 1 to nRecs
cFile_zero := alltrim( aArray [y,1] )
cFile := "c:\agrupamento2\tab\"
cFile_all := cFile + cFile_zero

@ 13,0 say "Deleta arquivo por arquivo, os registros invalidos."
* Entenda-se registros invalidos como aqueles em que nao há numero da notificação.
use ( cFile_all )
delete all for empty( nu_notific ) = .T.
goto top
pack
nRecs2 := reccount()
close

@ 14,0 say "Classificando os arquivos para processamento."
use "c:\agrupamento2\tmp\file_manager.dbf"
goto (y)
if nRecs2 = 0
replace processa with "N"
else
replace processa with "S"
endif
close

next

@ 15,0 say "Fim do processamento."

return nil