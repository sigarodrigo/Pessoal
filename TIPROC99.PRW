//exemplo do Processa() e incremento de régua

#INCLUDE 'PROTHEUS.CH'

User Function TIPROC99()
Processa( {|| xProcessa() }, "Aguarde","Iniciando o processamento",.F.)
return

************************************************************

Static Function xProcessa()
Local lAborta := .F. 
SetKey( VK_ESCAPE, { || lAborta := .T.  } )

ProcRegua(0) //seta o total de registros a processar

while !lAborta
	IncProc("Processando...") //incrementa a régua
EndDo

return
