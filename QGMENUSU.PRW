#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'TOPCONN.CH'


/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³LOJR69    º Autor ³ JOELDSON PEREIRA   º Data ³  09/01/12   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescricao ³ Codigo gerado pelo AP6 IDE.                                º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP6 IDE                                                    º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/

User Function QGMENUSU


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Declaracao de Variaveis                                             ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

Local cDesc1         := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2         := "de acordo com os parametros informados pelo usuario."
Local cDesc3         := "Relatório de Menus de Usuários"
Local cPict          := ""
Local titulo       := "Relatório de Menus de Usuários"
Local nLin         := 80

Local Cabec1       := ""
Local Cabec2       := ""
Local imprime      := .T.
Local aOrd := {}
Private lEnd         := .F.
Private lAbortPrint  := .F.
Private CbTxt        := ""
Private limite           := 80
Private tamanho          := "P"
Private nomeprog         := "LOJR69" // Coloque aqui o nome do programa para impressao no cabecalho
Private nTipo            := 18
Private aReturn          := { "Zebrado", 1, "Administracao", 2, 2, 1, "", 1}
Private nLastKey        := 0
Private cbtxt      := Space(10)
Private cbcont     := 00
Private CONTFL     := 01
Private m_pag      := 01
Private wnrel      := "QGMENUSU" // Coloque aqui o nome do arquivo usado para impressao em disco

Private cPerg		:= "QGMENUSU"
Private cString 	:= "ZU1"
Private aMod	 	:= {} 

AADD(aMod,{"01","Ativo Fixo"})
AADD(aMod,{"02","Compras"})
AADD(aMod,{"03","Contabilidade"})
AADD(aMod,{"04","Estoque/Custos"})
AADD(aMod,{"05","Faturamento"})
AADD(aMod,{"06","Financeiro"})
AADD(aMod,{"07","Gestão de Pessoal"})
AADD(aMod,{"08","Faturamento Serviço"})
AADD(aMod,{"09","Livros Fiscais"})
AADD(aMod,{"10","Planej.Contr.Produção"})
AADD(aMod,{"11","Veículos"})
AADD(aMod,{"12","Controle de Lojas"})
AADD(aMod,{"13","Call Center"})
AADD(aMod,{"14","Oficina"})
AADD(aMod,{"15","Protheus T Report Utility"})
AADD(aMod,{"16","Ponto Eletrônico"})
AADD(aMod,{"17","Easy Import Control"})
AADD(aMod,{"20","Recrutamento e Seleção Pessoal"})
AADD(aMod,{"23","Front Loja"})
AADD(aMod,{"34","Contabilidade Gerencial"})
AADD(aMod,{"40","Cargos e Salários"})
AADD(aMod,{"42","WMS - Gestão de Armazenagem"})
AADD(aMod,{"43","TMS - Gestão de Transporte"})

dbSelectArea("ZU1")
dbSetOrder(1)

dbSelectArea("ZU2")
dbSetOrder(1) 

ValidPerg(cPerg)
pergunte(cPerg,.F.)


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Monta a interface padrao com o usuario...                           ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

wnrel := SetPrint(cString,NomeProg,cPerg,@titulo,cDesc1,cDesc2,cDesc3,.T.,aOrd,.T.,Tamanho,,.T.)

If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

If nLastKey == 27
	Return
Endif

nTipo := If(aReturn[4]==1,15,18)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Processamento. RPTSTATUS monta janela com a regua de processamento. ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

RptStatus({|| RunReport(Cabec1,Cabec2,Titulo,nLin) },Titulo)
Return

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFun‡„o    ³RUNREPORT º Autor ³ AP6 IDE            º Data ³  09/01/12   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescri‡„o ³ Funcao auxiliar chamada pela RPTSTATUS. A funcao RPTSTATUS º±±
±±º          ³ monta a janela com a regua de processamento.               º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ Programa principal                                         º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/

Static Function RunReport(Cabec1,Cabec2,Titulo,nLin)

//Local nOrdem

cQuery := "SELECT ZU1_FILIAL FIL,ZU1_COD COD, ZU1_NOME NOME,ZU1_LOGIN LOGIN FROM ZU1010 "
cQuery += "WHERE ZU1_FILIAL = '" + xFilial("ZU1") + "' "
cQuery += "AND ZU1_COD BETWEEN '" + MV_PAR01 + "' AND '" + MV_PAR02 + "' "
cQuery += "AND D_E_L_E_T_ = ' ' "
cQuery += "ORDER BY ZU1_COD "

cARQSQL := Criatrab( , .F. )
Memowrit( cARQSQL, cQuery )

TcQuery cQuery Alias TMP NEW


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ SETREGUA -> Indica quantos registros serao processados para a regua ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

SetRegua(Contar("TMP","COD <> ''"))


TMP->(dbGoTop())
While TMP->(!EOF())
	
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ Verifica o cancelamento pelo usuario...                             ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	
	If lAbortPrint
		@nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
		Exit
	Endif
	
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ Impressao do cabecalho do relatorio. . .                            ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	
	Cabec1 := "Codigo: " + TMP->COD + " - Login:  " + TMP->LOGIN
	Cabec2 := "Nome:   " + TMP->NOME + Space(37) + "Status"
	
	If nLin > 60 // Salto de Página. Neste caso o formulario tem 55 linhas...
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin := 9
	Endif
	
	If ZU2->(DbSeek(TMP->FIL + TMP->COD))
		
		While TMP->FIL + TMP->COD == ZU2->ZU2_FILIAL + ZU2->ZU2_COD .And. TMP->(!Eof()) .And. ZU2->(!Eof())
			
			If nLin > 60 // Salto de Página. Neste caso o formulario tem 55 linhas...
				Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
				nLin := 9
			Endif
			
			nAbriu := FT_FUse(Alltrim(ZU2->ZU2_MENU))
			
			If nAbriu = -1
				@nLin,10 PSAY "ARQUIVO DE MENU NAO ENCONTRADO:" + Alltrim(ZU2->ZU2_MENU)
			Else
				nPos := aScan( aMod, { |X| X[1] = ZU2->ZU2_MODULO } )
				@nLin,0 PSAY "Módulo: " + ZU2->ZU2_MODULO +"-"+If(nPos>0,aMod[nPos][2],"") +" Menu: " + ZU2->ZU2_MENU
				nLin++
				@nLin,0 PSAY "Funciona: F, Não Funciona: NF, Não Utilizo: NU"
				nLin++
				While !FT_FEOF()
					
					If nLin > 60 // Salto de Página. Neste caso o formulario tem 55 linhas...
						Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
						nLin := 9
					Endif
					
					If AT('<Menu Status="Enable">',FT_FREADLN()) > 0 //MENU
						
						fT_fSkip()
						If AT('<Title lang="pt">&',FT_FREADLN()) > 0
							@nLin,0 PSAY Substr(FT_FREADLN(), AT('>',FT_FREADLN())+1, AT('</Title>',FT_FREADLN()) - AT('>',FT_FREADLN())-1)
							nLin++
						ElseIf AT('<Title lang="pt">',FT_FREADLN()) > 0	 //SUBMENU
							@nLin,2 PSAY Substr(FT_FREADLN(), AT('>',FT_FREADLN())+1, AT('</Title>',FT_FREADLN()) - AT('>',FT_FREADLN())-1)
							nLin++
						EndIf
						
					ElseIf AT('<MenuItem Status="Enable">',FT_FREADLN()) > 0   //ROTINA
						
						fT_fSkip()
						If AT('<Title lang="pt">',FT_FREADLN()) > 0
							@nLin,4 PSAY Padr(Substr(FT_FREADLN(), AT('>',FT_FREADLN())+1, AT('</Title>',FT_FREADLN()) - AT('>',FT_FREADLN())-1),21)
													
							While !FT_FEOF()
								fT_fSkip()
								If AT('<Function>',FT_FREADLN()) > 0

									@nLin,25 PSAY " - " + Substr(FT_FREADLN(), AT('>',FT_FREADLN())+1, AT('</Function>',FT_FREADLN()) - AT('>',FT_FREADLN())-1)
									@nLin,39 PSAY "_____________________________________(    )" 
									nLin++
									If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
										Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
										nLin := 9
									Endif
									
									Exit
								EndIf
							EndDo
						
						EndIf
									
					Else
						fT_fSkip()
					EndIf
					
				EndDo
				
				nLin++
				@nLin,20 PSAY "____________________________________" 
				nLin++
				@nLin,20 PSAY "Login: " + TMP->LOGIN
				nLin++  
				@nLin,20 PSAY "Nome:  " + TMP->NOME
				nLin++
				
				Cabec1 := "Codigo: " + TMP->COD + " - Login:  " + TMP->LOGIN
				Cabec2 := "Nome:   " + TMP->NOME + Space(37) + "Status"
				Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
				nLin := 9
			
			EndIf
			nLin++
			ZU2->(DbSkip())
		EndDo
		
	EndIf
	
	TMP->(dbSkip()) // Avanca o ponteiro do registro no arquivo     
	IncRegua()
	
	If TMP->(!EoF())	  				
		Cabec1 := "Codigo: " + TMP->COD + " - Login:  " + TMP->LOGIN
		Cabec2 := "Nome:   " + TMP->NOME + Space(37) + "Status"
	
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin := 9
	EndIf
	
EndDo

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Finaliza a execucao do relatorio...                                 ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

SET DEVICE TO SCREEN

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Se impressao em disco, chama o gerenciador de impressao...          ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

If aReturn[5]==1
	dbCommitAll()
	SET PRINTER TO
	OurSpool(wnrel)
Endif

MS_FLUSH()

TMP->(DbCloseArea())

Return

Static Function ValidPerg(cPerg)

Local aHelpPor := {}   

PutSx1( cPerg,'01','Codigo de :','','','mv_ch1','C',6,0,0,'G','','','','','MV_PAR01','','','','','','','','','','','','','','','','',aHelpPor,{},{} )
PutSx1( cPerg,'02','Codigo Ate:','','','mv_ch2','C',6,0,0,'G','','','','','MV_PAR02','','','','','','','','','','','','','','','','',aHelpPor,{},{} )

Return

