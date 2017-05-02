#include 'totvs.ch'
#include "rwmake.ch"
#include "tbiconn.ch"
#include "tbicode.ch"
#include "colors.ch"
//#include "color.ch"
#include "font.ch"
#include "ap5mail.ch"  

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³TIR003MAPF³ Autor ³ Jorez A. Nasato       ³ Data ³04/09/2012³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Relatorio das rotinas do sistemas e usuarios que tem acesso³±±
±±³          ³                                                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ TIR003MAPF()                                               ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso       ³Snap-On Do Brasil                                           ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³           Atualiza‡oes sofridas desde a constru‡ao inicial            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Programador ³Data      ³Motivo da Altera‡ao                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³            ³  /  /    ³                                               ³±±
±±³            ³          ³                                               ³±±
±±³            ³          ³                                               ³±±
±±³            ³          ³                                               ³±±
±±³            ³          ³                                               ³±±
±±³            ³          ³                                               ³±±
±±³            ³          ³                                               ³±±
±±³            ³          ³                                               ³±±
±±³            ³          ³                                               ³±±
±±³            ³          ³                                               ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/

User Function TIR003MAPF()
Private nUser		:= 0
Private cPerg		:= 'TIR002ACE'
Private aUsers		:= AllUsers()
Private aMenuAll	:= {}
Private cNomeArq	:= 'MENUS.DTC'
Private aMod		:= {}

	VldPerg()
	if ( !Pergunte(cPerg,.T.) )
		return .F.
	endIf

	if ( !Empty(mv_par01) )
		nUser	:= ASCAN(aUsers,{|x| AllTrim(x[1,1]) == AllTrim(mv_par01)})
	endif

	// Se foi definido um usuario o agrupamento sera por usuario
	if ( nUser <> 0 )
		mv_par03 == 2
	endif

	// Definindo os Modulos utilizados na Snap-On
	AADD(aMod,{'01', 'Ativo Fixo'})
	AADD(aMod,{'02', 'Compras'})
	AADD(aMod,{'03', 'Contabilidade'})
	AADD(aMod,{'04', 'Estoque/Custos'})
	AADD(aMod,{'05', 'Faturamento'})
	AADD(aMod,{'06', 'Financeiro'})
	AADD(aMod,{'07', 'Gestao de Pessoal'})
	AADD(aMod,{'09', 'Livros Fiscais'})
	AADD(aMod,{'10', 'Planej. Cont. Prod.'})
	AADD(aMod,{'11', 'Veículo'})
	AADD(aMod,{'12', 'Controle de Lojas'})
	AADD(aMod,{'13', 'Telemarketing'})
	AADD(aMod,{'16', 'Ponto Eletronico'})
	AADD(aMod,{'17', 'Cont. de Importacao'})
	AADD(aMod,{'18', 'Terminal de Condulta Funcionário'})
	AADD(aMod,{'19', 'Manutenção de Ativos'})
	AADD(aMod,{'20', 'Recrutamento de Seleção'})
	AADD(aMod,{'21', 'Inspeção de Entrada'})
	AADD(aMod,{'23', 'Front Loja'})
	AADD(aMod,{'34', 'Contab. Gerencial'})
	AADD(aMod,{'35', 'Medicina e Segurança do Trabalho'})
	AADD(aMod,{'39', 'OMS - Gestão de Distribuição'})
	AADD(aMod,{'42', 'WMS - Gestão de Armazenagem'})
	AADD(aMod,{'43', 'TMS - Gestão de Transporte'})
	AADD(aMod,{'64', 'Processos Trabalhistas'})
	

	Processa({|| CargaMenu()},'Analisando menus dos usuários...')

	if ( Len(aMenuAll) > 0 )
		Processa({|| CargaFun()},'Analisando as rotinas do sistema...')

		if ( mv_par02 == 1 )
			RptStatus({|| RunReport()}, 'Imprimindo as informações...')
		else
			Processa({|| RunArqCSV()}, 'Gerando as informações...')
		endif
	endif
	
TRB->(DbCloseArea())

return

Static Function CargaMenu()
Local i
Local j
Local cMenu
Local aUsuario	
Local aAcessos	
Local aMenus
Local aGrpMenu

	ProcRegua(Len(aUsers))

	for i := if(nUser == 0, 1, nUser) To if(nUser == 0, Len(aUsers), nUser)
		IncProc("Analisando usuario "+AllTrim(aUsers[i,1,4]))

		if ( aUsers[i,1,1] == '000000' ) // Não imprime o Administrador
			Loop
		endif

		if ( aUsers[i,1,17] ) // Se estiver bloqueado nao imprime 
			Loop
		endif

		aUsuario	:= aUsers[i,1]
		aAcessos	:= aUsers[i,2]
		aMenus		:= {}

		// Carrega somente os menus que o usuario tem acesso
		AEVAL(aUsers[i,3], {|f| if(Substr(f,3,1) <> 'X', AADD(aMenus, f), nil)}, 1) 

		// Verifica se o usuario pertence a algum grupo e carrega os menus do grupo
		if ( Len(aUsuario[10]) > 0 )
			for j := 1 to Len(aUsuario[10])
				AEval(FWGrpMenu(aUsuario[10][j]),{|e| if(ASCAN(aMenus, e) == 0 .and. Substr(e,3,1) <> 'X', AADD(aMenus, e), nil)}, 1)
			next
		endif

		cLogin	:= aUsuario[2]
		cNome	:= AllTrim(aUsuario[4])
		cDepto	:= AllTrim(aUsuario[12])
		cCargo	:= AllTrim(aUsuario[13])
		cEMail	:= AllTrim(aUsuario[14])

		for j := 1 To Len(aMenus)
			if ( ASCAN(aMod,{|x| x[1] == Substr(aMenus[j],1,2)}) > 0 )
				cMenu	:= AllTrim(Capital(AllTrim(Substr(aMenus[j],4,Len(aMenus[j])))))
				aAdd( aMenuAll,{ cMenu, cLogin, cNome, cDepto, cCargo, cEMail } )
			endif
		next
	next
return

Static Function ADMenu(pMenus, pVetor, pInd)

	AEVAL(pMenus, {|g,h| iif(Upper(AllTrim(pMenus[h])) == Upper(AllTrim(pVetor[pInd])),nil, AADD(pMenus, pVetor[pInd]))}, 1)
return

Static Function CargaFun()
Local i
Local j
Local k
Local l
Local aCampos
Local cRelTrb
Local aEstrutura
Local aItens
Local aFuncoes
Local cArquivo

	aCampos	:= {	{"MENU",	"C",35,0},;
					{"GRUPO",	"C",35,0},;
					{"DESCRI",	"C",35,0},;
					{"MODO",	"C",1,0},;
					{"TIPO",	"C",12,0},;
					{"FUNCAO",	"C",12,0},;
					{"ACS",		"C",10,0},;
					{"MODULO",	"C",02,0},;
					{"ARQS",	"C",255,0},;
					{"A1",		"C",1,0},;
					{"A2",		"C",1,0},;
					{"A3",		"C",1,0},;
					{"A4",		"C",1,0},;
					{"A5",		"C",1,0},;
					{"A6",		"C",1,0},;
					{"A7",		"C",1,0},;
					{"A8",		"C",1,0},;
					{"A9",		"C",1,0},;
					{"A0",		"C",1,0},;
					{"LOGIN",	"C",20,0},;
					{"NOME",	"C",35,0},;
					{"DEPTO",	"C",30,0},;
					{"CARGO",	"C",30,0},;
					{"EMAIL",	"C",35,0}}

	cRelTrb	:= CriaTrab(aCampos,.T.)

	DbUseArea(.T.,"DTCCDX",cRelTrb,"TRB",.T.,.F.)

	// Verifica se o arquivo ja existe e apaga
	if ( File(cNomeArq) )
		Delete File &(cNomeArq)
	endif

	ProcRegua(Len(aMenuAll))

	for i := 1 To Len(aMenuAll)
		IncProc( "Lendo " + Lower(AllTrim(aMenuAll[i,1])) )

		aEstrutura := XNULoad(aMenuAll[i,1])
	
		for j := 1 To Len(aEstrutura)
			cTipo := {'Atualizacoes', 'Consultas', 'Relatorios', 'Miscelaneas'}[j]

			aItens	:= aEstrutura[j,3]
			for k := 1 To Len(aItens)
				if ( ValType(aItens[k,3]) == "A" )
					aFuncoes	:= aItens[k,3]

					for l := 1 To Len(aFuncoes)
						if ( aFuncoes[l,2] <> 'E' ) // Verifica se o item do menu esta habilitado - E = Enable, D = Disable e H = Hide
							Loop
						endif

						if ( Len(aFuncoes[l]) >= 7 )
							cArquivo	:= ""
							aEval( aFuncoes[l,4],{ |h| cArquivo += h + " " },,)

							TRB->(RecLock("TRB",.T.))
							TRB->MENU		:= aMenuAll[i,1]
							TRB->DESCRI	:= RTrim(aFuncoes[l,1,1])
							TRB->MODO		:= aFuncoes[l,2]
							TRB->TIPO		:= cTipo
							TRB->FUNCAO	:= Upper(RTrim(aFuncoes[l,3]))
							TRB->ACS		:= Upper(aFuncoes[l,5])
							TRB->MODULO	:= RTrim(aFuncoes[l,6])
							TRB->ARQS		:= RTrim(cArquivo)
							TRB->GRUPO		:= aItens[k,1,1]
							TRB->A1			:= Substr(Upper(aFuncoes[l,5]) ,1,1)
							TRB->A2			:= Substr(Upper(aFuncoes[l,5]) ,2,1)						
							TRB->A3			:= Substr(Upper(aFuncoes[l,5]) ,3,1)						
							TRB->A4			:= Substr(Upper(aFuncoes[l,5]) ,4,1)
							TRB->A5			:= Substr(Upper(aFuncoes[l,5]) ,5,1)
							TRB->A6			:= Substr(Upper(aFuncoes[l,5]) ,6,1)												
							TRB->A7			:= Substr(Upper(aFuncoes[l,5]) ,7,1)
							TRB->A8			:= Substr(Upper(aFuncoes[l,5]) ,8,1)
							TRB->A9			:= Substr(Upper(aFuncoes[l,5]) ,9,1)
							TRB->A0			:= Substr(Upper(aFuncoes[l,5]) ,10,1)						
							TRB->LOGIN		:= RTrim(aMenuAll[i,2])
							TRB->NOME		:= RTrim(Capital(aMenuAll[i,3]))
							TRB->DEPTO		:= RTrim(Capital(aMenuAll[i,4]))
							TRB->CARGO		:= RTrim(Capital(aMenuAll[i,5]))
							TRB->EMAIL		:= RTrim(lower(aMenuAll[i,6]))
							TRB->(MsUnLock())								
						endif
					next
				else
					if ( AllTrim(aItens[k,2]) <> 'E' ) // Verifica se o item do menu esta habilitado - E = Enable, D = Disable e H = Hide
						Loop
					endif

					TRB->(RecLock("TRB",.T.))
					TRB->MENU		:= aMenuAll[i,1]
					TRB->DESCRI	:= RTrim(aItens[k,1,1])
					TRB->MODO		:= RTrim(aItens[k,2])
					TRB->TIPO		:= cTipo					
					TRB->FUNCAO	:= RTrim(Upper(RTrim(aItens[k,3])))
					TRB->ACS		:= RTrim(Upper(aItens[k,5]))
					TRB->MODULO	:= RTrim(aItens[k,6])
					TRB->ARQS		:= AllTrim(cArquivo)
					TRB->GRUPO		:= aItens[k,1,1]
					TRB->A1			:= Substr(Upper(aItens[k,5]) ,1,1)
					TRB->A2			:= Substr(Upper(aItens[k,5]) ,2,1)						
					TRB->A3			:= Substr(Upper(aItens[k,5]) ,3,1)						
					TRB->A4			:= Substr(Upper(aItens[k,5]) ,4,1)
					TRB->A5			:= Substr(Upper(aItens[k,5]) ,5,1)
					TRB->A6			:= Substr(Upper(aItens[k,5]) ,6,1)												
					TRB->A7			:= Substr(Upper(aItens[k,5]) ,7,1)
					TRB->A8			:= Substr(Upper(aItens[k,5]) ,8,1)
					TRB->A9			:= Substr(Upper(aItens[k,5]) ,9,1)
					TRB->A0			:= Substr(Upper(aItens[k,5]) ,10,1)
					TRB->LOGIN		:= RTrim(aMenuAll[i,2])
					TRB->NOME		:= RTrim(Capital(aMenuAll[i,3]))
					TRB->DEPTO		:= RTrim(Capital(aMenuAll[i,4]))
					TRB->CARGO		:= RTrim(Capital(aMenuAll[i,5]))
					TRB->EMAIL		:= RTrim(lower(aMenuAll[i,6]))
					TRB->(MsUnLock())
				endif    
			next	
		next
	next

	TRB->(DbGoTop())
	Copy To &cNomeArq
return

Static Function RunReport()
Local i
Local j	
Local cOpcoes
Local cAgrupa
Local cModulo
Local cInd 	:= CriaTrab(nil,.F.)
Local cDescri	:= ''
Local nLin		:= 100
Local cTitulo	:= 'Mapa de acessos por Item de Menu'
Private Cabec1
Private Cabec2
Private nTipo		:= 15
Private m_pag		:= 01
Private nLastKey	:= 0
Private aReturn	:= {"Zebrado",1,"Administracao",1,2,1,"",1}
Private Tamanho	:= "G"
Private nomeprog	:= "TIR003MAPF"	// Nome do programa para impressao no cabecalho
Private wnrel		:= "TIR003MAPF"	// Nome do arquivo usado para impressao em disco

	wnrel := SetPrint('',wnrel,cPerg ,cTitulo,,,,.F.,"",.T.,Tamanho,,.T.)

	if ( nLastKey == 27 )
		return
	endif

	SetDefault(aReturn,'')

	if ( nLastKey == 27 )
		return
	endif

	DbUseArea(.T.,"DTCCDX",'MENUS',"MENUS",.T.,.F.) //DBFCDX

	if ( mv_par03 == 1 ) // Agrupa por Item
		IndRegua('MENUS', cInd, "MODULO+TIPO+GRUPO+DESCRI+NOME")
	elseif ( mv_par03 == 2 ) // Agrupa por Usuario
		IndRegua('MENUS', cInd, "NOME+MODULO+TIPO+GRUPO+DESCRI")
	endif

	Menus->(DbGoBottom())
	SetRegua(Menus->(Recno()))
	Menus->(DbGoTop())

	Cabec1	:= '.'

	if ( mv_par03 == 1 ) // Agrupa por Item
//					"0         1         2         3         4         5         6         7         8         9       100         1         2         3         4         5         6         7         8         9       200         1         2
//	 				"01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
		Cabec2 :=	"Usuario                             Menu                                Módulo / Menu / SubMenu                                Opções"
//					"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXX / ATUALIZACOES / XXXXXXXXXXXXXXXX PESQUISAR, VISUALIZAR, INCLUIR, ALTERAR, EXCLUIR, OPCAO 6, OPCAO 7, OPCAO 8, OPCAO 9, OPCAO 10
	elseif ( mv_par03 == 2 ) // Agrupa por Usuario
//					"0         1         2         3         4         5         6         7         8         9       100         1         2         3         4         5         6         7         8         9       200         1         2
// 					"01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
		Cabec2 :=	"Módulo / Menu / SubMenu / Item                                                               Menu                                Opções"
//					"XXXXXXXXXXXXXXXXXXXX/ATUALIZACOES / XXXXXXXXXXXXXXXX / XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX PESQUISAR, VISUALIZAR, INCLUIR, ALTERAR, EXCLUIR, OPCAO 6, OPCAO 7, OPCAO 8, OPCAO 9, OPCAO 10
	endif

	do while ( !Menus->(Eof()) )
		IncRegua()

		if ( nLin > 60 )
			Cabec(cTitulo,Cabec1,Cabec2,,Tamanho,nTipo)
			nLin	:= 9
		endif

		cAgrupa	:= {Menus->Descri, Menus->Nome}[mv_par03]
		j		:= ASCAN(aMod,{|x| x[1] == Menus->Modulo})
		cModulo	:= Left(if(j > 0,aMod[j,2],''),20)

		if ( cDescri <> cAgrupa )
			if ( !Empty(cDescri) )
				nLin	+= 2
				@ nLin,000 PSay Replicate('-',220)
				nLin	+= 2

				if ( nLin > 60 )
					Cabec(cTitulo,Cabec1,Cabec2,,Tamanho,nTipo)
					nLin	:= 9
				endif
			endif

			if ( mv_par03 == 1 ) // Agrupa por Item
				@ nLin,10 PSay 'Item: '+AllTrim(cModulo) + ' / '+Menus->Descri
			elseif ( mv_par03 == 2 ) // Agrupo por Usuario
				@ nLin,10 PSay 'Usuario: '+Menus->Nome
			endif
			nLin	+= 2
			cDescri	:= cAgrupa
		endif

		cOpcoes	:= ''
		if ( Upper(AllTrim(Menus->Tipo)) == 'ATUALIZACOES' )
			for i:= 1 to 10
				if ( ( i < 10 ) .and. ( Upper(Menus->&('A'+cValToChar(i))) == 'X' ) ) .or. ( ( i == 10 ) .and. ( Upper(Menus->A0) == 'X' ) )
					if ( !Empty(cOpcoes) )
						cOpcoes	+= ', '
					endif
					cOpcoes += {'PESQUISAR', 'VISUALIZAR', 'INCLUIR', 'ALTERAR', 'EXCLUIR', 'OPCAO 6', 'OPCAO 7', 'OPCAO 8', 'OPCAO 9', 'OPCAO 10'}[i]
				endif
			next
		endif

		if ( mv_par03 == 1 ) // Agrupa por Item
			@ nLin,000 PSay Menus->Nome
			@ nLin,036 PSay Menus->Menu
			@ nLin,072 PSay AllTrim(cModulo) + ' / ' + Menus->Tipo + ' / ' + Left(Menus->Grupo,16)
		elseif ( mv_par03 == 2 ) // Agrupa por Usuario
			@ nLin,000 PSay Left(cModulo,20) + '/' + Menus->Tipo + ' / ' + Left(Menus->Grupo,16) + ' / ' + Menus->Descri
			@ nLin,091 PSay Menus->Menu
		endif

		@ nLin,127 PSay cOpcoes
		nLin++

		Menus->(DbSkip())
	enddo

	Menus->(DbCloseArea())

	Set Device To Screen

	if ( aReturn[5]==1 )
	   dbCommitAll()
	   SET PRINTER TO
	   OurSpool(wnrel)
	endif
	MS_FLUSH()
return

Static Function RunArqCSV()
Local i
Local j
Local cCaminho	:= "C:\RELATO\"
Local cArquivo	:= 'Mapa de acessos.csv'
Local cInd 	:= CriaTrab(nil,.F.)
Local nArq
Local cLinha
Local cOpcoes
Local cModulo
Local oExcelApp

	nArq := FCreate(cCaminho + cArquivo)
	if ( nArq == -1 )
		Aviso(FunDesc(),"O arquivo " + cCaminho + cArquivo + " não pode ser criado.",{"OK"})
		return
	endif

	DbUseArea(.T.,"DBFCDX",'MENUS',"MENUS",.T.,.F.)

	IndRegua('MENUS', cInd, "MODULO+TIPO+GRUPO+DESCRI+NOME")

	Menus->(DbGoBottom())
	ProcRegua(Menus->(Recno()))
	Menus->(DbGoTop())

	if ( mv_par03 == 1 ) // Agrupa por Item
//		cLinha :=	"Item;Usuario;Menu;Módulo / Menu / SubMenu;Opções"
		cLinha :=	"Módulo;Menu;SubMenu;Item;Usuario;Menu;Opções"
	elseif ( mv_par03 == 2 ) // Agrupa por Usuario
		cLinha :=	"Usuario;Módulo;Menu;SubMenu;Item;Menu;Opções"
	endif

	Fwrite(nArq,cLinha + CRLF)

	do while ( !Menus->(Eof()) )
		IncProc('Item: '+AllTrim(Menus->Descri))

		j		:= ASCAN(aMod,{|x| x[1] == Menus->Modulo})
		cModulo	:= if(j > 0,aMod[j,2],'')

		cOpcoes	:= ''
		if ( Upper(AllTrim(Menus->Tipo)) == 'ATUALIZACOES' )
			for i:= 1 to 10
				if ( ( i < 10 ) .and. ( Upper(Menus->&('A'+cValToChar(i))) == 'X' ) ) .or. ( ( i == 10 ) .and. ( Upper(Menus->A0) == 'X' ) )
					if ( !Empty(cOpcoes) )
						cOpcoes	+= ', '
					endif
					cOpcoes += {'PESQUISAR', 'VISUALIZAR', 'INCLUIR', 'ALTERAR', 'EXCLUIR', 'OPCAO 6', 'OPCAO 7', 'OPCAO 8', 'OPCAO 9', 'OPCAO 10'}[i]
				endif
			next
		endif

		if ( mv_par03 == 1 ) // Agrupa por Item
			cLinha	:= cModulo+';'
			cLinha	+= Menus->Tipo+';'
			cLinha	+= Menus->Grupo+';'
			cLinha	+= Menus->Descri+';'
			cLinha	+= Menus->Nome+';'
			cLinha	+= Menus->Menu+';'
			cLinha	+= cOpcoes
		elseif ( mv_par03 == 2 ) // Agrupa por Usuario
			cLinha	:= Menus->Nome+';'
			cLinha	+= cModulo+';'
			cLinha	+= Menus->Tipo+';'
			cLinha	+= Menus->Grupo+';'
			cLinha	+= Menus->Descri+';'
			cLinha	+= Menus->Menu+';'
			cLinha	+= cOpcoes
		endif

		Fwrite(nArq,cLinha + CRLF)

		Menus->(DbSkip())
	enddo

	Menus->(DbCloseArea())

	FClose(nArq)
	oExcelApp := MsExcel():New()
	oExcelApp:WorkBooks:Open( cCaminho + cArquivo )
	oExcelApp:SetVisible(.T.)
return

Static Function VldPerg
Local aRegs   := {} 

	cPerg := PADR(cPerg,10)
	//		   {Grupo /Ord		/Pergunta		/PergEsp  		/PergIng		/Variavel	/Tipo	/Tamanho	/Decimal	/Presel	/GSC/Valid	/Var01		/Def01			/DefSpa1	/DefEng1	/Cnt01	/Var02	/Def02		/DefSpa2	/DefEng2	/Cnt02	/Var03	/Def03	/DefSpa3	/DefEng3	/Cnt03	/Var04	/Def04	/DefSpa4	/DefEng4	/Cnt04	/Var05	/Def05	/DefSpa5	/DefEng5	/Cnt05	/F3		/PYME	/GRPSX6	/HELP	}
	aadd(aRegs,{cPerg,"01",	"Usuário",		"Usuário",		"Usuário",		"mv_ch1",	"C",	6,			0,			0,		"G","",		"mv_par01","",				"",			"",			"",		"",		"",			"",			"",			"",		"",		"",		"",			"",			"",		"",		"",		"",			"",			"",		"",		"",		"",			"",			"",		"USR",	"",		"",		""})
	aAdd(aRegs,{cPerg,"02",	"Tipo Relat.",	"Tipo Relat.",	"Tipo Relat.",	"mv_ch2",	"N",	1,			0,			0,		"C","",		"mv_par02","Impressão",	"",			"",			"",		"",		"Excel",	"",			"",			"",		"",		"",		"",			"",			"",		"",		"",		"",	   		"",			"",		"",		"",		"",			"",			"",		"",		"",		"",		""})
	aAdd(aRegs,{cPerg,"03",	"Agrupar Por",	"Agrupar Por",	"Agrupar Por",	"mv_ch3",	"N",	1,			0,			0,		"C","",		"mv_par03","Item",		"",			"",			"",		"",		"Usuario",	"",			"",			"",		"",		"",		"",			"",			"",		"",		"",		"",	   		"",			"",		"",		"",		"",			"",			"",		"",		"",		"",		""})
	lValidPerg(aRegs)
Return 
