PROCEDURE Main()

cls
set color to w+/
? "CSIS Software 2021 - listador.exe"
? ""
? ""
? "------------------------------------------------------------------------"
? "Compilado no Harbour Make (hbmk2) versao 3.0.0"
? "Roda somente com argumento de entrada."
? "Exemplo: listador.exe run"
? "------------------------------------------------------------------------"
set color to g+/

/*
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
*/

   criaDBF:= {}
   AAdd(criaDBF, {"filename", "C", 45, 0})
   AAdd(criaDBF, {"records", "N", 10, 0})
   dbcreate( "c:\agrupamento2\tmp\file_manager", criaDBF )

use "c:\agrupamento2\tmp\file_manager.dbf"

   aFFList := HB_DirScan( "c:\agrupamento2\tab\" , "*.dbf" )

        x1Row := 0 

         FOR EACH x1Row IN aFFList
		 append blank
		 replace filename with PAD( x1Row[ 1 ], 23 )		 
            ? PAD( x1Row[ 1 ], 23 )
         NEXT x1Row

close

return nil