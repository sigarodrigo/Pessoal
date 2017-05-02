#include "ap5mail.ch"
#INCLUDE "PROTHEUS.CH"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณMANDAEMAILบAutor  ณLucas Silva         บ Data ณ  10/05/07   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ                                                            บฑฑ
ฑฑบ          ณ                                                            บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ AP                                                        บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
User Function EnvEmail(l_Alerta,cPara,cAssunto,cCorpo,cAnexo,cServer, cCC)
local lResult   := .f.
Local cError    := ""
Local cPassword := GetMV("MV_RELAPSW")
Local cServer	:= GetMV("MV_RELSERV") 
Local cAccount	:= GetMV("MV_RELACNT")
Local cFrom		:= GetMV("MV_RELAUSR")
Local lAuth    	:= GetNewPar("MV_RELAUTH",.F.)
Local lConnect	:= .T.

CONNECT SMTP SERVER cServer ACCOUNT cAccount PASSWORD cPassword RESULT lConnect
 
If lAuth
   lRetAuth := MailAuth(cAccount,cPassword)
Else
   lRetAuth := .T.
EndIf

//ConOut(" ==== Enviando E-Mail ==== ")
//ConOut("De   : " + _cDe)
//ConOut("Para : " + _cPara)

If U_Validaemail(cPara)
	If lRetAuth 	                                                      
		If Empty(cAnexo)
			SEND MAIL FROM cFrom TO cPara CC cCC SUBJECT cAssunto BODY cCorpo RESULT lResult   
		Else
			SEND MAIL FROM cFrom TO cPara CC cCC SUBJECT cAssunto BODY cCorpo ATTACHMENT cAnexo RESULT lResult
		EndIf
	EndIf
Endif

If l_Alerta
	
	If lResult
		Alert("Email enviado com Sucesso!!")
	Else
		GET MAIL ERROR cError
		//		Help(" ", 1, "ATENCAO",,cError,4,5)
		Alert("Problema no envio do E-mail: "+ cError + " !")
	EndIf
Endif

DISCONNECT SMTP SERVER	


Return

user Function ValidaEmail(email)
local lValida := .t.
local c_Email := Alltrim(email)

if at('.',c_Email) = 0 .or. at('@',c_Email) = 0 .or. empty(c_Email)
	lValida := .f.
else
	c_Carac := "!#$%&*()+=/\|?'{}[]ชบ:,งฐ<>"+'"'
	
	For i := 1 to len(c_Email)
		if SubStr(c_Email,i,1)$c_Carac .and. lValida
			lValida := .f.
		endif
	Next
endif

return(lvalida)


