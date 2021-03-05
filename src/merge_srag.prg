PROCEDURE Main()

clear
set console off
set status off
set scoreboard off
set century on
set date british
set confirm off
set exact on

set color to w+/n
DispBox (1,2,6,74,1)
set color to gr+/n 
@ 2,3 say " merge_srag versao 1.0                                 "
@ 3,3 say " Uniao de arquivos SRAG hospitalizado                  "
@ 4,3 say " CSIS Software                                         "
@ 5,3 say " Marco/2021                                            "
set color to g+/n 

PUBLIC nPos := 13
PUBLIC aName, tmp
PUBLIC cFile := ""

@ 7,1 say "Inicializando o script..."

* Verifica se os argumentos são válidos. 
if HB_ArgV( 1 ) <> "run"
@ 8,1 say "Esses parametros nao sao validos."
__Quit()
else
@ 8,1 say "Argumento valido."
endif

* verifica se há arquivos DBF na pasta especifica.
if HB_ArgV ( 1 ) = "run"
if (File( 'c:\agrupamento2\data\*.dbf' ) = .T.)
@ 9,1 say "Arquivos dbf presentes na pasta 'data' para uniao de arquivos..."
else
@ 9,1 say "Arquivos dbf nao foram detectados na pasta 'data' para a uniao de arquivos..."
__Quit()
endif
endif

if HB_ArgV( 1 ) = "run"
nLen = ADir( "c:\agrupamento2\data\*.dbf" )
@ 10,1 say "Copiando o arquivo modelo..."
copy file "c:\agrupamento2\mod\srag_model.dbf" to "c:\agrupamento2\data\srag_model.dbf"
endif

if HB_ArgV( 1 ) = "run" 
@ 11,1 say "Anexando os arquivos baixados no arquivo srag_model.dbf"
endif

IF nLen > 0 // coloca em um array os arquivos dbf descompactados no script 'unzipping_srag'.
   aName := Array( nLen ) 
   if HB_ArgV( 1 ) = "run"
   ADir( "c:\agrupamento2\data\SRAGHOSPITALIZADO*.dbf", aName)
   endif
      
   FOR tmp := 1 TO nLen
   
   * seleciona no array criado anteriormente, o arquivo que será anexado.
   if HB_ArgV( 1 ) = "run"
   cFile = "c:\agrupamento2\data\"+aName[tmp]
   endif
   
   @ nPos, 1 say (aName [ tmp ]) + "-----" + cFile

* Um a um os arquivos descompactados são anexados no arquivo srag_model.dbf.
if HB_ArgV( 1 ) = "run"
use "c:\agrupamento2\data\srag_model.dbf"
append from (cFile)
close all
endif

nPos = nPos + 1
   
NEXT
ENDIF

* Após a anexação dos registros, verifica se existem registros no arquivo srag_model.dbf.
if HB_ArgV( 1 ) = "run"
use "c:\agrupamento2\data\srag_model.dbf"
if reccount()>0
@ nPos,1 say "Identificado registros dentro da tabela srag_model.dbf."
@ nPos+1, 1 say "Renomeando arquivo srag_model.dbf..."
close all
rename "c:\agrupamento2\data\srag_model.dbf" to "c:\agrupamento2\data\sraghosp.dbf"
if (File( 'c:\agrupamento2\data\sraghosp.dbf' ) = .T.)
@ nPos+2, 1 say "Arquivo srag_model.dbf renomeado para sraghosp.dbf..."
@ nPos+3,1 say "Finalizando o script."
else
@ nPos,1 say "Nao foi identificado registros dentro da tabela srag_model.dbf. O processo falhou."
@ nPos+1, 1 say "Finalizando o script..."
endif
endif
endif

RETURN