PROCEDURE Main()

set century on
set date british

use "c:\agrupamento2\data\sraghosp.dbf"
do while .not. eof()
cDateX := ctod(dt_notific)
replace dt_not2 with cDateX
skip
enddo

return nil