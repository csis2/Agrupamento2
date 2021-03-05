PROCEDURE Main()

cls
set color to w+/
? "CSIS Software 2020 - filler.exe"
? "Preenche o campo 'escala' dos arquivos DBF criados em 'Agrupamento.exe'."
? "Agrega os dados gerados no arquivo 'escala.dbf'."
? "------------------------------------------------------------------------"
? "Compilado no Harbour Make (hbmk2) versao 3.0.0"
? "Roda somente com argumento de entrada."
? "Exemplo: filler.exe run"
? "------------------------------------------------------------------------"
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

* Verifica se os arquivos que serao trabalhados, de fato existem.
if file("c:\agrupamento2\tmp\scale_M0.dbf") = .F.
set color to r+/
? "Erro! O arquivo 'scale_M0.dbf' nao existe."
__Quit()
endif

if file("c:\agrupamento2\tmp\scale_0_1.dbf") = .F.
set color to r+/
? "Erro! O arquivo 'scale_0_1.dbf' nao existe."
__Quit()
endif

if file("c:\agrupamento2\tmp\scale_2_7.dbf") = .F.
set color to r+/
? "Erro! O arquivo 'scale_2_7.dbf' nao existe."
__Quit()
endif

if file("c:\agrupamento2\tmp\scale_8_14.dbf") = .F.
set color to r+/
? "Erro! O arquivo 'scale_8_14.dbf' nao existe."
__Quit()
endif

if file("c:\agrupamento2\tmp\scale_M15.dbf") = .F.
set color to r+/
? "Erro! O arquivo 'scale_M15.dbf' nao existe."
__Quit()
endif

set color to g+/
? "Preenche scale_M0.dbf"
USE "c:\agrupamento2\tmp\scale_M0.dbf"
REPLACE escala WITH "FICHA DIGITADA ANTES DO OBITO" WHILE .not. EOF()
CLOSE
? "Ok."

? "Preenche scale_0_1.dbf"
USE "c:\agrupamento2\tmp\scale_0_1.dbf"
REPLACE escala WITH "EM ATE 1 DIA" WHILE .not. EOF()
CLOSE
? "Ok."

? "Preenche scale_2_7.dbf"
USE "c:\agrupamento2\tmp\scale_2_7.dbf"
REPLACE escala WITH "ENTRE 2 E 7 DIAS" WHILE .not. EOF()
CLOSE
? "Ok."

? "Preenche scale_8_14.dbf"
USE "c:\agrupamento2\tmp\scale_8_14.dbf"
REPLACE escala WITH "ENTRE 8 E 14 DIAS" WHILE .not. EOF()
CLOSE
? "Ok."

? "Preenche scale_M15.dbf"
USE "c:\agrupamento2\tmp\scale_M15.dbf"
REPLACE escala WITH "15 OU MAIS DIAS" WHILE .not. EOF()
CLOSE
? "Ok."

? "Agrupando os dados gerados no arquivo 'escala.dbf'."
use "c:\agrupamento2\tmp\escala.dbf"
append from "c:\agrupamento2\tmp\scale_M0.dbf"
append from "c:\agrupamento2\tmp\scale_0_1.dbf" 
append from "c:\agrupamento2\tmp\scale_2_7.dbf"
append from "c:\agrupamento2\tmp\scale_8_14.dbf" 
append from "c:\agrupamento2\tmp\scale_M15.dbf"

? "Fim do script."

return nil