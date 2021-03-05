PROCEDURE Main()

cls
set color to g+/
? "CSIS Software 2020 - intervalo.exe"
set color to w+/
? "Subtrai a data de digitacao com a data da evolucao do paciente."
? "Os registros em que nao foi possivel efetuar o calculo sao excluidos."
set color to g+/
? "Compilado no Harbour Make (hbmk2) versao 3.0.0"
? "Roda somente com argumento de entrada."
? "Exemplo: intervalo.exe run"

set century on
set date british

* Verifica se existe argumento.
if empty(HB_ArgV(1)) = .T.
set color to r+/
? "Erro! O programa precisa que voce forneca argumento de entrada."
nErro = 1
__Quit()
endif

* Verifica se o argumento é valido.
set exact on
if (HB_ArgV(1)) <> "run"
set color to r+/
? "Erro! O argumento escolhido nao e valido."
nErro = 1
__Quit()
endif
set exact off

* Subtrai as datas relativas aos campos "dt_digita" e "dt_evoluca".
* Detalhe: o campo "dt_digita" é do tipo date, enquanto o campo "dt_evoluca" é do tipo caractere.

? "------------------------------------------------"
set color to bg+/
? "Iniciando..."

BEGIN SEQUENCE WITH {| oError | oError:cargo := { ProcName( 1 ), ProcLine( 1 ) }, Break( oError ) }
use c:\agrupamento2\data\sraghosp2.dbf

do while .not. eof()

dDigita := (dt_digita)
dEvoluc := ctod(dt_evoluca)

if (empty( dDigita ) = .F.) .and. (empty( dEvoluc ) = .F.)
nDias := (dDigita) - (dEvoluc)
replace intervalo with nDias
endif

skip
enddo

RECOVER USING oError
set color to r+/
? "Erro! Preenchimento campo intervalo falhou."
errorlevel(1)
quit
END

? "Campo intervalo preenchido...........................OK."

* Exclui os registros que nao tiveram o campo intervalo calculado.
* Isso acontece porque o campo dt_digita ou o campo dt_evoluca estão vazios. Ou ambos.

BEGIN SEQUENCE WITH {| oError | oError:cargo := { ProcName( 1 ), ProcLine( 1 ) }, Break( oError ) }
goto top
delete for (empty(intervalo) = .T.) while .not. EOF()
pack

RECOVER USING oError
set color to r+/
? "Erro! Exclusao de registros falhou."
errorlevel(1)
quit
END

? "Exclusao dos registros nao calculados.................OK."
? "Fim do script."

return nil