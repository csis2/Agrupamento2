PROCEDURE Main()

set color to g+/
cls

@ 1,0 say "CSIS Software 2021                            "
@ 2,0 say "Converte o campo 'dt_evoluc' de tipo caracter "
@ 3,0 say "para tipo data.                               "
@ 4,0 say "----------------------------------------------"
@ 5,0 say "Compilado no Harbour Make (hbmk2) versao 3.0.0"
@ 6,0 say "Roda somente com argumento de entrada.        "
@ 7,0 say "Exemplo: conv_dt_evoluc.exe run               "
@ 8,0 say "----------------------------------------------"

* Verifica se existe argumento.
if empty(HB_ArgV(1)) = .T.
set color to r+/
@ 9,0 say "Erro! O programa precisa que voce forneca argumento de entrada."
__Quit()
endif

* Verifica se o argumento é valido.
set exact on
if (HB_ArgV(1)) <> "run"
set color to r+/
@ 9,0 say "Erro! O argumento escolhido nao e valido."
__Quit()
endif
set exact off

* Configurndo variaveis de ambiente.
set century on
set date british

close all
cDBFFile = "c:\agrupamento2\data\sraghosp.dbf"
cDBFFile2 = "c:\agrupamento2\data\sraghosp2.dbf"

set color to bg+/n
@ 10,0 say "Criando novo arquivo com um novo campo."
* Copia a estrutura do arquivo "sraghosp.dbf" e acrescenta campos novos.
	   use ( cDBFFile ) new
	   aStruct := dbStruct( cDBFFile )
	   aAdd( aStruct, { "dt_evoluc2", "D", 8, 0 } )
	   
* Cria uma tabela nova a partir da estrutura da tabela selecionada como arquivo de entrada.
* A tabela nova já contém os campos criados pela aplicação.
	   dbCreate( (cDBFFile2 ) , aStruct )	   
	   close

@ 11,0 say "Transferindo os dados para a nova tabela."
* Transferindo os dados da tabela original para a tabela recem criada.
	   use ( cDBFFile2 ) new
	   append from ( cDBFFile )
	   close

@ 12,0 say "Converte tipo caracter para tipo data."
use "c:\agrupamento2\data\sraghosp2.dbf"
goto top
do while .not. eof()
dDateX15 := ctod(dt_evoluca)
replace dt_evoluc2 with dDateX15
skip
enddo

close

set color to w+/
@ 13,0 say "Fim do processo..."

return nil