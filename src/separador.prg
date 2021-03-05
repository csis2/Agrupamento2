PROCEDURE Main()

set date british
set century on
cFileName := ""
cSemIni := ""
cSemFin := ""

cls
set color to w+/
@ 1,0 say "CSIS Software 2021 - separador.exe"
@ 2,0 say "Separa os registros do arquivo 'sraghosp2.dbf' por semana epidemiologica de"
@ 3,0 say "acordo com as datas de obito. No final do processamento havera varios ar-  "
@ 4,0 say "quivos DBF, cada qual com o nome e os registros de uma semana epidemiolo-  "
@ 5,0 say "gica especifica."
@ 6,0 say "Tambem gera por ordem de criacao, um registro no arquivo 'file_manager.dbf',"
@ 7,0 say "com o nome do arquivo e a ordem em que ele foi criado."
@ 8,0 say "---------------------------------------------------------------------------"
@ 9,0 say "Compilado no Harbour Make (hbmk2) versao 3.0.0"
@ 10,0 say "Roda somente com argumento de entrada."
@ 11,0 say "Exemplo: separador.exe run"
@ 12,0 say "---------------------------------------------------------------------------"
set color to g+/

* Verifica se existe argumento.
if empty(HB_ArgV(1)) = .T.
set color to r+/
@ 13,0 say "Erro! O programa precisa que voce forneca argumento de entrada."
__Quit()
endif

* Verifica se o argumento é valido.
set exact on
if (HB_ArgV(1)) <> "run"
set color to r+/
@ 13,0 say "Erro! O argumento escolhido nao e valido."
__Quit()
endif
set exact off

@ 14,0 say "Cria a tabela 'weeks_table.dbf'."
   matstru0:= {}
   AAdd(matstru0, {"filename", "C", 9, 0})
   AAdd(matstru0, {"iniweek", "C", 10, 0})
   AAdd(matstru0, {"finweek", "C", 10, 0})
   dbcreate("c:\agrupamento2\tmp\weeks_table", matstru0)

@ 15,0 say "Cria a tabela 'file_manager.dbf'."
   criaDBF:= {}
   AAdd(criaDBF, {"filename", "C", 45, 0})
   AAdd(criaDBF, {"class", "N", 5, 0})
   AAdd(criaDBF, {"processa", "C", 1, 0})   
   dbcreate( "c:\agrupamento2\tmp\file_manager", criaDBF )

@ 16,0 say "Recebe no arquivo DBF os registros do arquivo hr_weeks.txt."
use "c:\agrupamento2\tmp\weeks_table"
append from "c:\agrupamento2\tab\hr_weeks.txt" delimited with ','

@ 17,0 say "Cria um array com os registros do arquivo 'weeks_table.dbf'."
goto top
aArray := {}
nRecs := reccount()
for x := 1 to nRecs
	  AAdd( aArray, {alltrim(filename),alltrim(iniweek),alltrim(finweek)} )
	  skip
endfor

close

for y := 1 to nRecs
cFileName := aArray [y,1]
cSemIni := aArray [y,2]
cSemFin := aArray [y,3]

@ 18,0 say "Cria um arquivo DBF com cada semana epidemiologica do calendario."
use "c:\agrupamento2\data\sraghosp2.dbf"
temp := "c:\agrupamento2\tab\" + "_" + alltrim(str(Y)) + "_" + (cFileName) + ".dbf"
? "Gerando arquivo " + (temp) + ".dbf..."
copy to (temp) for ( (dt_evoluc2 >= ctod(cSemIni)) .and. (dt_evoluc2 <= ctod(cSemFin)) )
? "ok."
close
* Também gera por ordem de criação, um registro no arquivo 'file_manager.dbf', com o nome do arquivo
* e a ordem em que ele foi criado.
use "c:\agrupamento2\tmp\file_manager.dbf"
append blank
replace filename with ( "_" + alltrim(str(Y)) + "_" + (cFileName) + ".dbf" )
replace class with (y)
close

endfor

? "Fim do processo."

return nil