#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'TBICONN.CH'

/*
_____________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+----------+----------+-------+-----------------------+------+----------+¦¦
¦¦¦ Programa ¦INFO03 	¦ Autor ¦ RAMILSON SOBRAL       ¦ Data ¦ 02/11/13 ¦¦¦
¦¦+----------+----------+-------+-----------------------+------+----------+¦¦
¦¦¦Descrição ¦ Atualização da Tabela ZU1010 com os Novos Usuários 		  ¦¦¦
¦¦¦          ¦ Cadastrados.												  ¦¦¦
¦¦+----------+------------------------------------------------------------+¦¦
¦¦¦ Uso      ¦ Loja/Schedule	                                          ¦¦¦
¦¦+----------+------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
*/

User Function INFO03()

Private lAuto 		:= If(Select("SX2") == 0,.T.,.F.)
Private dPB_ATUZU1 	:= U_CriaSX6( {"  ", "D", "PB_ATUZU1","Indica quando Atualizar a ZU1.","","","20140404"} )

If lAuto
	If dDataBase > dPB_ATUZU1
		RpcSetType(3)
		RpcSetEnv("01","02",,,"LOJ", "Importacao dos Usuários para ZU1",{},,,,)
		
		cFilImp := ""
		cDtIni := cDtFin := DTOS(Date())
		Conout("INFO03 - " + Time() + " - Inicio do Processo" )
	Else
		Return
	EndIf	
EndIf
	
aUsers  := AllUsers()
aGroups := AllGroups()

DbSelectArea("ZU1")
ZU1->(DbSetOrder(1))

DbSelectArea("ZU2")
ZU2->(DbSetOrder(1))

ProcRegua(Len(aUsers))
nCont := 0

For nI := 1 To len(aUsers)
	
	nCont++
	If !lAuto
		IncProc("Atualizando Usuários..." + Padr(Str(nCont),8) )
	EndIf	
	
	If ZU1->(!DbSeek(xFilial("ZU1") + aUsers[nI][1][1] ))
		RecLock("ZU1",.T.)
	Else
		RecLock("ZU1",.F.)
	EndIf
	
	ZU1->ZU1_FILIAL 	:= xFilial("ZU1")
	ZU1->ZU1_COD		:= aUsers[nI][1][1]
	ZU1->ZU1_LOGIN		:= aUsers[nI][1][2]
	ZU1->ZU1_NOME		:= aUsers[nI][1][4]
	ZU1->ZU1_SUP		:= aUsers[nI][1][11]
	ZU1->ZU1_DEPART		:= aUsers[nI][1][12]
	ZU1->ZU1_CARGO		:= aUsers[nI][1][13]
	ZU1->ZU1_EMAIL		:= aUsers[nI][1][14]
	ZU1->ZU1_NACESS		:= aUsers[nI][1][15]
	ZU1->ZU1_BLOQ		:= aUsers[nI][1][17]
	ZU1->ZU1_DTVAL    	:= aUsers[nI][1][6]
	ZU1->ZU1_DIGANO		:= aUsers[nI][1][18]
	ZU1->ZU1_SPOOL 		:= aUsers[nI][2][3]
	ZU1->ZU1_ACESS		:= aUsers[nI][2][5]	

	cEmpFil := ""
	For nJ:=1 To Len(aUsers[nI][2][6])
		cEmpFil += If(nJ = 1,"","|") + aUsers[nI][2][6][nJ] 
	Next
	
	ZU1->ZU1_EMP := cEmpFil

	ZU1->(MsUnlock())
	
	For nJ := 1 To Len(aUsers[nI][3])
		
		If Substr(aUsers[nI][3][nJ],3,1) != 'X'
		
			If ZU2->(!DbSeek(xFilial("ZU2") + aUsers[nI][1][1] + aUsers[nI][3][nJ]))
				RecLock("ZU2",.T.)
			Else
				RecLock("ZU2",.F.)
			EndIf
			ZU2->ZU2_FILIAL	:= xFilial("ZU2")
			ZU2->ZU2_COD	:= aUsers[nI][1][1]
			ZU2->ZU2_MODULO	:= Substr(aUsers[nI][3][nJ],1,2)
			ZU2->ZU2_NIVEL	:= Substr(aUsers[nI][3][nJ],3,1)
			ZU2->ZU2_MENU 	:= Substr(aUsers[nI][3][nJ],4,len(aUsers[nI][3][nJ]))
			ZU2->(MsUnlock())
	
		EndIf
		
	Next nJ	

Next nI

If lAuto
	Conout("INFO03 - " + Time() + " - Fim do Processo" )
	Processa( {|| AlteraPar() } )	
Else
	MsgInfo("Tabela Atualizada!")
	AlteraPar()
EndIf
     
Return 

/*
_____________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+----------+----------+-------+-----------------------+------+----------+¦¦
¦¦¦ Programa ¦AlteraPar	¦ Autor ¦ RAMILSON SOBRAL       ¦ Data ¦ 02/11/13 ¦¦¦
¦¦+----------+----------+-------+-----------------------+------+----------+¦¦
¦¦¦Descrição ¦ Grava a data da Última Atualização da ZU1010 no Parâmetro  ¦¦¦
¦¦¦          ¦ PB_ATUZU1.												  ¦¦¦
¦¦+----------+------------------------------------------------------------+¦¦
¦¦¦ Uso      ¦ Loja/Schedule	                                          ¦¦¦
¦¦+----------+------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
*/
Static Function AlteraPar()

DbSelectArea("SX6")
DbSetOrder(1)

If SX6->(DbSeek(Space(2)+"PB_ATUZU1")) .And. RecLock("SX6",.F.)
	SX6->X6_CONTEUD := DtoC(dDataBase)
EndIf
	
Return