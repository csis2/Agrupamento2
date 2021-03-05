PROCEDURE Main()

cls
set color to w+/
? "CSIS Software 2021 - converter.exe"
? ""
? ""
? "------------------------------------------------------------------------"
? "Compilado no Harbour Make (hbmk2) versao 3.0.0"
? "Converte os registros em formato de strings no arquivo 'consolidado_percent.dbf' em"
? "valores tipo double no arquivo 'consolidado_percentual.dbf'."
? "Roda somente com argumento de entrada."
? "Exemplo: converter.exe run"
? "------------------------------------------------------------------------"
set color to g+/

? "Iniciando o script..."
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

   ? "Criando nova tabela."
   criaDBF:= {}
   AAdd(criaDBF, {"periodo", "C", 20, 0})
   AAdd(criaDBF, {"oportuno", "B", 3, 2})
   AAdd(criaDBF, {"d_2_e_7", "B", 3, 2})
   AAdd(criaDBF, {"d_8_e_14", "B", 3, 2})
   AAdd(criaDBF, {"d_15_mais", "B", 3, 2})
   AAdd(criaDBF, {"d_total", "B", 3, 2})
   dbcreate( "c:\agrupamento2\tmp\consolidado_percentual.dbf", criaDBF )

? "Criando um array do campo 'periodo'."
aArray0 := {}
use c:\agrupamento2\tmp\consolidado_percent.dbf
nRecs0 := reccount()
do while .not. eof()
AAdd( aArray0, {periodo} )
skip
enddo
close

? "Criando um array do campo 'oportuno'."
aArray1 := {}
use c:\agrupamento2\tmp\consolidado_percent.dbf
nRecs := reccount()
do while .not. eof()
if oportuno = "100,0"
oportuno_d = val ("100")
else
oportuno_d = val (substr(oportuno,1,2) + "." + substr(oportuno,4,2))
endif
AAdd( aArray1, {oportuno_d} )
skip
enddo
close

? "Criando um array do campo 'd_2_e_7'."
aArray2 := {}
use c:\agrupamento2\tmp\consolidado_percent.dbf
nRecs2 := reccount()
do while .not. eof()
if d_2_e_7 = "100,0"
d_2_e_7_d = val ("100")
else
d_2_e_7_d = val (substr(d_2_e_7,1,2) + "." + substr(d_2_e_7,4,2))
endif
AAdd( aArray2, {d_2_e_7_d} )
skip
enddo
close

? "Criando um array do campo 'd_8_e_14'."
aArray3 := {}
use c:\agrupamento2\tmp\consolidado_percent.dbf
nRecs3 := reccount()
do while .not. eof()
if d_8_e_14 = "100,0"
d_8_e_14_d = val ("100")
else
d_8_e_14_d = val (substr(d_8_e_14,1,2) + "." + substr(d_8_e_14,4,2))
endif
AAdd( aArray3, {d_8_e_14_d} )
skip
enddo
close

? "Criando um array do campo 'd_15_mais'."
aArray4 := {}
use c:\agrupamento2\tmp\consolidado_percent.dbf
nRecs4 := reccount()
do while .not. eof()
if d_15_mais = "100,0"
d_15_mais_d = val ("100")
else
d_15_mais_d = val (substr(d_15_mais,1,2) + "." + substr(d_15_mais,4,2))
endif
AAdd( aArray4, {d_15_mais_d} )
skip
enddo
close

? "Criando um array do campo 'total'."
aArray5 := {}
use c:\agrupamento2\tmp\consolidado_percent.dbf
nRecs5 := reccount()
do while .not. eof()
if total= "100,0"
d_total = val ("100")
else
d_total_d = val (total)
endif
AAdd( aArray5, {d_total_d} )
skip
enddo
close

? "Povoando os valores na nova tabela."
use c:\agrupamento2\tmp\consolidado_percentual.dbf
for x := 1 to nRecs
append blank
replace periodo with (aArray0[x,1])
replace oportuno with (aArray1[x,1])
replace d_2_e_7 with (aArray2[x,1])
replace d_8_e_14 with (aArray3[x,1])
replace d_15_mais with (aArray4[x,1])
replace d_total with (aArray5[x,1])
next
close

? "Exportando para um arquivo no formato CSV"
use c:\agrupamento2\tmp\consolidado_percentual.dbf
copy to c:\agrupamento2\tmp\consolidado_percentual.txt delimited with ';'
close

? "Fim do script."
return nil