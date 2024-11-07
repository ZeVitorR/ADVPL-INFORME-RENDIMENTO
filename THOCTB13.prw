#INCLUDE "PROTHEUS.CH"
#INCLUDE "FWPRINTSETUP.CH"
#INCLUDE "RPTDEF.CH"
#INCLUDE "TBICONN.CH" 
#INCLUDE "TOPCONN.CH"
#INCLUDE "FILEIO.CH" 

#DEFINE VBOX       080
#DEFINE VSPACE     008
#DEFINE HSPACE     010
#DEFINE SAYVSPACE  008
#DEFINE SAYHSPACE  008
#DEFINE HMARGEM    030
#DEFINE VMARGEM    030
#DEFINE IMP_SPOOL    2

/*/{Protheus.doc} THOCTB14

Relatorio de Informe Rendimento

@type function
@author José Vitor Rodrigues
@since 10/05/2024
@version V1
/*/
User Function THOCTB13()
	Local   aArea      			:= GetArea()
	Private cLogoD				:= GetSrvProfString("Startpath","") + "THCMLOGO.BMP"
	Private isBlind 			:= isBlind()
	Private cTitlePerg          := "Informe Rendimento"
	Private aTitulos            := {}
	Private aLotes              := {}
	Private aLotXCli            := {}
	Private cPreNuTp			:= ""
	
	IF ! Perguntas()
        RETURN
    ENDIF

	cClienteDe	:= mv_par01
	cLojaDe		:= mv_par02
	cProdutoDe	:= mv_par03
	cAnoBase	:= mv_par04
	Imprimi()

	RestArea(aArea)
Return

/*
	Função para imprimir o Relatorio
*/
Static Function Imprimi()
	Local oSetup	:= Nil
	Local cDirImp   := "c:\spool\"

	Private lImpGel	:= .T.
	Private cFont  		:= "Arial"  

	Private oFont08   	:= TFont():New(cFont,0,-1*(08+2),.T.,.F.,5,.T.,5,.T.,.F.)
	Private oFont08N	:= TFont():New(cFont,0,-1*(08+2),.T.,.T.,5,.T.,5,.T.,.F.)
	Private	oFont09   	:= TFont():New(cFont,0,-1*(08+2),.T.,.F.,5,.T.,5,.T.,.F.)
	Private	oFont09N  	:= TFont():New(cFont,0,-1*(08+2),.T.,.T.,5,.T.,5,.T.,.F.)
	Private	oFont10		:= TFont():New(cFont,0,-1*(10+2),.T.,.F.,5,.T.,5,.T.,.F.)
	Private	oFont10N 	:= TFont():New(cFont,0,-1*(10+2),.T.,.T.,5,.T.,5,.T.,.F.)
	Private	oFont11		:= TFont():New(cFont,0,-1*(11+2),.T.,.F.,5,.T.,5,.T.,.F.)
	Private	oFont11N 	:= TFont():New(cFont,0,-1*(11+2),.T.,.T.,5,.T.,5,.T.,.F.)
	Private	oFont12  	:= TFont():New(cFont,0,-1*(12+2),.T.,.F.,5,.T.,5,.T.,.F.)
	Private	oFont12N 	:= TFont():New(cFont,0,-1*(12+2),.T.,.T.,5,.T.,5,.T.,.F.)
	Private	oFont14  	:= TFont():New(cFont,0,-1*(14+2),.T.,.F.,5,.T.,5,.T.,.F.)
	Private	oFont14N  	:= TFont():New(cFont,0,-1*(14+2),.T.,.T.,5,.T.,5,.T.,.F.)
	Private	oFont16  	:= TFont():New(cFont,0,-1*(16+2),.T.,.F.,5,.T.,5,.T.,.F.)
	Private	oFont16N  	:= TFont():New(cFont,0,-1*(16+2),.T.,.T.,5,.T.,5,.T.,.F.)
	Private	oFont18N  	:= TFont():New(cFont,0,-1*(18+2),.T.,.T.,5,.T.,5,.T.,.F.)

	Private oBrushC     := TBrush():New(,RGB(166,166,166))  

	cHora:= substr(time(),1,2)+substr(time(),4,2)+substr(time(),7,2)
	oPrinter:= FWMSPrinter():New("INFORME_"+dtos(dDataBase)+"_"+cHora,IMP_PDF,.T.,cDirImp,.T.,.F.,,,.F.,.F.,,.T.)
	oPrinter:cPathPDF := cDirImp
	
	
	nFlags := PD_ISTOTVSPRINTER + PD_DISABLEPAPERSIZE + PD_DISABLEPREVIEW + PD_DISABLEMARGIN  + PD_DISABLEDESTINATION
	If (!oPrinter:lInJob)
		oSetup := FWPrintSetup():New(nFlags, "INFORME")
		// ----------------------------------------------
		// Define saida
		// ----------------------------------------------
		oSetup:SetPropert(PD_PRINTTYPE   , 6 ) //PDF
		oSetup:SetPropert(PD_ORIENTATION , 1 ) //PORTRAIT
		oSetup:SetPropert(PD_DESTINATION , 2 ) //Local
		oSetup:SetPropert(PD_MARGIN      , {60,60,60,60})
		oSetup:SetPropert(PD_PAPERSIZE   , 2) //A4
	EndIf

	If !isBlind
		If !(oSetup:Activate() == PD_OK) 
			Return (.F.) 
		Endif	
	EndIf
				
	oPrinter:SetPortrait()
	oPrinter:SetPaperSize(DMPAPER_A4)
	oPrinter:SetMargin(60,60,60,60)
	PixelX := oPrinter:nLogPixelX()
	PixelY := oPrinter:nLogPixelY()

	nHPage := oPrinter:nHorzRes()
	nHPage *= (300/PixelX)
	nHPage -= HMARGEM
	nVPage := oPrinter:nVertRes()
	nVPage *= (300/PixelY)
	nVPage -= VBOX  

	monta_vetores()
		
	ImprimiDados()
	
	If !isBlind
		oPrinter:Preview()
		FreeObj(oPrinter)
		oPrinter := Nil 
	Else
		oPrinter:Print()
		Return openFile(cDirImp,oPrinter:cFileName) 
	EndIf

Return

/*
	Função para impressão do relatorio
*/
Static Function ImprimiDados()
	Local nLote      := 0
	Local nTitulo    := 0
	Private nLinha   := 180

	For nLote:= 1 to Len(aLotes)
		nLinha   		:= 180
		cEmpreendimento	:= aLotes[nLote,2] 
		lImprimi		:= .F.
		For nTitulo:= 1 to Len(aTitulos)
			If alltrim(aTitulos[nTitulo,12]) == alltrim(cEmpreendimento)
				If substr(dtos(aTitulos[nTitulo,16]),1,4) == cAnoBase .AND. !empty(aTitulos[nTitulo,16]) .AND. aTitulos[nTitulo][21] == "NOR"
					lImprimi:= .T.
				Endif 
			Endif 
		Next nTitulo 
		
		If lImprimi
			oPrinter:StartPage()
			ImprimiCabecalho(nLote)
			ImprimiItens(nLote)
			
			oPrinter:EndPage()
		Endif
	Next nLote
	

Return

/*
	Função para impressão do cabeçalho
*/
Static Function ImprimiCabecalho(nLote)
	Local nNomFil := ''
	//010703
	//124001000000826

	cFilAtual   := xFilial("SE2")
	
	//definindo dados da empresa
	If(cFilAtual == '124001')
		cQuery := " SELECT M0_NOMECOM, M0_CGC "
		cQuery += " FROM "+ retsqlname('SM0')
		cQuery += " WHERE M0_CODFIL = '118001' "
		cQuery += " AND M0_CODIGO = '02' "
		TCQuery cQuery NEW ALIAS "FILIAL"
		DbSelectArea("FILIAL")
			cNomeEmp := ("FILIAL")->M0_NOMECOM
			cCnpj := TRANSFORM(AllTrim(("FILIAL")->M0_CGC),If(Len(AllTrim(("FILIAL")->M0_CGC))==14,"@R 99.999.999/9999-99","@R 999.999.999-99"))
		("FILIAL")->(dbCloseArea())
	Else
		cNomeEmp := SM0->M0_NOMECOM
		cCnpj := TRANSFORM(AllTrim(SM0->M0_CGC),If(Len(AllTrim(SM0->M0_CGC))==14,"@R 99.999.999/9999-99","@R 999.999.999-99"))
	EndIf	

	if(len(Alltrim(cNomeEmp)) >= 48)
		nPos = 340
	else
		nPos = 260
	endif

	oPrinter:SayBitmap(126,055,cLogoD,565,290)  
	nLinha+=20
	nNomFil := Alltrim(cNomeEmp)
	oPrinter:SayAlign(nLinha,0650,nNomFil,oFont18N,1580,200,CLR_BLACK,1,0)
	nLinha+=60
	oPrinter:SayAlign(nPos,0930,"CNPJ: "+ cCnpj,oFont12,1300,200,CLR_BLACK,1,0)
	nLinha+=60
	oPrinter:SayAlign(nLinha,0930,"ANO BASE: "+cAnoBase,oFont12,1300,100,CLR_BLACK,1,0)
	oPrinter:Line( 450, 50, 450, 2210, 0,"-2")

	dbSelectArea("SA1")
	dbSetOrder(1)
	dbSeek(FWxFilial("SA1")+aLotes[nLote,10]+aLotes[nLote,11])

	nPos:= aScan(aTitulos, {|x| x[12] == aLotes[nLote,2]})
	If nPos > 0
		dDataContrato:= dtoc(stod(aTitulos[nPos,6]))
	Else
		dDataContrato:= ""
	Endif
		
	nLinha+=200
	oPrinter:SayAlign(nLinha,055,"Informe Anual - "+cAnoBase,oFont18N,2150,300,CLR_BLACK,2,0)
	oPrinter:Line( 650, 50, 650, 2210, 0,"-2")

	nLinha+=220
	oPrinter:Say(nLinha,100,"Dados do Cliente:",oFont16N)
	nLinha+=60
	oPrinter:Say(nLinha,100,"Nome:",oFont12N)
	oPrinter:Say(nLinha,450,SA1->A1_NOME,oFont12)
	nLinha+=50
	oPrinter:Say(nLinha,100,"CPF/CNPJ:",oFont12N)
	oPrinter:Say(nLinha,450,TRANSFORM(SA1->A1_CGC,If(Len(Alltrim(SA1->A1_CGC))==14,"@R 99.999.999/9999-99","@R 999.999.999-99")),oFont12)
	nLinha+=50
	oPrinter:Say(nLinha,100,"Empreendimento:",oFont12N)
	oPrinter:Say(nLinha,450,aLotes[nLote,2],oFont12)
	nLinha+=50
	oPrinter:Say(nLinha,100,"Data do Contrato:",oFont12N)
	oPrinter:Say(nLinha,450,dDataContrato,oFont12)

	oPrinter:Line( 1000, 50, 1000, 2210, 0,"-2")
	

Return

/*
	Função para impressão do rodape
*/
Static Function ImprimiItens(nLote)
	Local nTitulo:= 0
	Local xNum   := ""
	local nPos   := 1325
	local nLinha2 := 0
	
	nLinha+=150
	oPrinter:Say(nLinha,100,"Resumo do Pagamento:",oFont16N)

	nLinha+=85

	oPrinter:FillRect( {1160,0100,1245,2210},oBrushC)   
	oPrinter:SayAlign(nLinha,0100,"Descrição",oFont14N,1000,0,CLR_WHITE,2,0)
	oPrinter:SayAlign(nLinha,1300,"Valor",oFont14N,1000,0,CLR_WHITE,2,0)
	// oPrinter:Say(nLinha,0100,"MES/ANO",oFont12N,,CLR_WHITE)
	// oPrinter:Say(nLinha,0400,"VALOR PRINCIPAL",oFont12N,,CLR_WHITE)
	// oPrinter:Say(nLinha,0800,"DESCONTO/NEGOCIACAO",oFont12N,,CLR_WHITE)
	// oPrinter:Say(nLinha,1350,"IPCA/IGPM/OUTROS",oFont12N,,CLR_WHITE)
	// oPrinter:Say(nLinha,1770,"JUROS MORA",oFont12N,,CLR_WHITE)
	// oPrinter:Say(nLinha,2050,"TOTAL",oFont12N,,CLR_WHITE)
	nLinha+=70

	aSort( aTitulos,,, { |x,y| x[32] < y[32] } )
	
	nValJuros  := 0
	nTotal     := 0
	nTotalGeral:= 0
	nValorTotal:= 0
	nDSCAno    := 0
	nAcrsAno   := 0
 	
	cEmpreendimento	:= aLotes[nLote,2] 
	cCliente		:= aLotes[nLote,10] 
	cLoja 			:= aLotes[nLote,11] 
	cProduto		:= aLotes[nLote,1] 
	For nTitulo:= 1 to Len(aTitulos)
		If alltrim(aTitulos[nTitulo,12]) == alltrim(cEmpreendimento)
			cDtBaixa        := SUBSTR(DTOS(aTitulos[nTitulo][16]),1,4)
			If substr(dtos(aTitulos[nTitulo,16]),1,4) == cAnoBase .AND. !empty(aTitulos[nTitulo,16]) .AND. aTitulos[nTitulo][21] == "NOR"
				cMes  	 	:= MesExtenso(substr(dtos(aTitulos[nTitulo,7]),5,2))+"/"+substr(dtos(aTitulos[nTitulo,7]),1,4)
				xNum  	 	:= aTitulos[nTitulo,1]
				nValJuros	+= aTitulos[nTitulo,18]
				cLote    	:= aTitulos[nTitulo,11]
				nAcrsAno    += aTitulos[nTitulo,23]
				nTotalItem  := (aTitulos[nTitulo,22]+aTitulos[nTitulo,23]+aTitulos[nTitulo,18])-aTitulos[nTitulo,29]
				// oPrinter:Say(nLinha,0100,cMes,oFont12)
				// oPrinter:Say(nLinha,0450,transform(aTitulos[nTitulo,22],"@E 999,999,999.99"),oFont12,,,,1)
				// oPrinter:Say(nLinha,0900,transform(aTitulos[nTitulo,29],"@E 999,999,999.99"),oFont12,,,,1)
				// oPrinter:Say(nLinha,1380,transform(aTitulos[nTitulo,23],"@E 999,999,999.99"),oFont12,,,,1)
				// oPrinter:Say(nLinha,1750,transform(aTitulos[nTitulo,18],"@E 999,999,999.99"),oFont12,,,,1)
				// oPrinter:Say(nLinha,2000,transform(nTotalItem,"@E 999,999,999.99"),oFont12,,,,1)
				// oPrinter:Line(nPos , 50, nPos, 2210, 1,"-1")
				// nPos += 80
				// nLinha+=80
				If nValorTotal == 0
					nValorTotal		:= GetValorContrato(cCliente,cProduto)
				Endif
			Endif
			
			If aTitulos[nTitulo][21] $ "NOR" .AND. cDtBaixa == cAnoBase .AND. !empty(aTitulos[nTitulo,16])
				nDSCAno     += ((aTitulos[nTitulo][22] + aTitulos[nTitulo][23] + aTitulos[nTitulo][18]) - aTitulos[nTitulo][10])  
				nTotal	 	:= aTitulos[nTitulo,10]-aTitulos[nTitulo,18]

				nTotalGeral	+= nTotal
			Endif
		Endif
	Next nTitulo

	If !empty(xNum)
		// Calcula Correção do período
		// cQuery  := " SELECT SUM (ZZN_DIFER) DIFER FROM " +retsqlname("ZZN")+ " "  // ZZN010
		// cQuery  += "   WHERE ZZN_CLIENT = '" + cClienteDe + "' "
		// cQuery  += "     AND ZZN_LOJA   = '" + cLojaDe + "' "
		
		// cQuery  += "  AND ZZN_FILIAL = '" + xFilial("ZZN") + "' "

		// cQuery  += "  AND ZZN_NUM    = '" + xNum + "' "   
		// cQuery  += "  AND ZZN_PREFIX    = 'CVD' "   
		// cQuery  += "  AND SUBSTRING(ZZN_DATA,1,4)  >= '" + cAnoBase + "' AND SUBSTRING(ZZN_DATA,1,4) <= '" + cAnoBase + "' "
		// cQuery  += "  AND D_E_L_E_T_ = '' "
		// TCQuery cQuery NEW ALIAS "TCACR"
		// DbSelectArea("TCACR")
		// //nAcrsAno := DIFER
		// TCACR->(DbCloseArea())

		// Calcula Saldo em Aberto período
		DbSelectArea("ZZN")
		DbSetOrder(2)
		cQuery := " SELECT * FROM " +retsqlname("SE1")   + " 						 " // SE1010
		cQuery += "   WHERE E1_CLIENTE  = '" + cClienteDe  + "' 					 "
		cQuery += "      AND E1_LOJA     = '" + cLojaDe + "' 						 "
		cQuery += "      AND CONCAT(E1_PREFIXO,E1_NUM,E1_TIPO)  = " + cPreNuTp
		cQuery += "      AND E1_PREFIXO <> 'ZZZ' 									 "
		cQuery += "      AND (E1_BAIXA   = '' OR E1_BAIXA > '"+cAnoBase+"1231') 	 "
		cQuery += "      AND D_E_L_E_T_  = '' ORDER BY E1_PARCELA 					 "
		TCQuery cQuery NEW ALIAS "TCSLD"
		DbSelectArea("TCSLD")
		DbGoTop()
		nSldAber := 0
		cDtIgpm  := "12/"+cAnoBase
		xFil     := E1_FILIAL
		xNum     := E1_NUM

		While !EOF()	
			cQuery := " SELECT * FROM " +retsqlname("ZZN")   + "  				"
			cQuery += "   WHERE ZZN_CLIENT    = '" + cClienteDe  + "' 			"
			cQuery += "      AND ZZN_LOJA     = '" + cLojaDe + "' 				"
			cQuery += "      AND ZZN_NUM      = '" + TCSLD->(E1_NUM) + "' 		"
			cQuery += "      AND ZZN_PARC     = '" + TCSLD->(E1_PARCELA) + "' 	"
			cQuery += "      AND ZZN_PREFIX   = '" + TCSLD->(E1_PREFIXO)+"' 	"
			cQuery += "      AND ZZN_FILIAL   = '" + TCSLD->(E1_FILIAL)+"' 		"
			cQuery += "      AND SUBSTRING(ZZN_MESREF,4,4)   = '" + cAnoBase+"' "
			cQuery += "      AND D_E_L_E_T_   = '' 								"
			cQuery += "      ORDER BY ZZN_MESREF DESC 							"
			TCQuery cQuery NEW ALIAS "TCZZN"
			DbSelectArea("TCZZN")
			DbGoTop()
			
			If ("TCZZN")->(!EOF())
				nSldAber := nSldAber + TCSLD->(E1_VALOR)+TCZZN->(Round(ZZN_ACRESC,2))
			Else
				nSldAber := nSldAber + TCSLD->(E1_VALOR+Round(E1_ACRESC,2))
			EndIf
			DbSelectArea("TCSLD")
			DbSkip()

			("TCZZN")->(dbCloseArea())
		EndDo
		If nSldAber == 0
			DbSelectArea("TCSLD")
			DbGoTop()
			While !EOF()
				nSldAber := nSldAber + TCSLD->E1_VALOR
				DbSkip()
			EndDo
		EndIf
		TCSLD->(DbCloseArea())


		nSaldo:= aLotes[1,3]+aLotes[1,4]+aLotes[1,5]
		
		nLinha2 := nLinha + 50
		// oPrinter:Box(nLinha-50,050,nLinha+300,2250)
		//SayAlign(nLinha,0100,"ANO BASE: "+cAnoBase,oFont12,1300,100,CLR_BLACK,1,0)
		oPrinter:SayAlign(nLinha,0100,"VALOR CONTRATO",oFont12N,1400,100,CLR_BLACK,0,0)
		oPrinter:SayAlign(nLinha,1800,"R$",oFont12,200,100,CLR_BLACK,0,0)
		oPrinter:SayAlign(nLinha,1900,transform(nValorTotal,"@E 99,999,999.99"),oFont12,300,100,CLR_BLACK,1,0)
		oPrinter:Line(nPos , 0100, nPos, 2210, 0,"-1")
		nPos += 80
		nLinha += 80
		nLinha2 += 80
		oPrinter:SayAlign(nLinha,0100,"VALOR PAGO EM "+cAnoBase,oFont12N,1400,100,CLR_BLACK,0,0)
		oPrinter:SayAlign(nLinha,1800,"R$",oFont12,200,100,CLR_BLACK,0,0)
		oPrinter:SayAlign(nLinha,1900,transform(nTotalGeral,"@E 99,999,999.99"),oFont12,300,100,CLR_BLACK,1,0)
		oPrinter:Line(nPos , 0100, nPos, 2210, 0,"-1")
		nPos += 80
		nLinha += 80
		nLinha2 += 80
		oPrinter:SayAlign(nLinha,0100,"SALDO EM "+"31/12/"+cAnoBase,oFont12N,1400,100,CLR_BLACK,0,0)
		oPrinter:SayAlign(nLinha,1800,"R$",oFont12,200,100,CLR_BLACK,0,0)
		oPrinter:SayAlign(nLinha,1900,transform(nSldAber,"@E 99,999,999.99"),oFont12,300,100,CLR_BLACK,1,0)
		oPrinter:Line(nPos , 0100, nPos, 2210, 0,"-1")
		// nPos += 80
		// nLinha += 80
		// nLinha2 += 80
		// oPrinter:SayAlign(nLinha,0100,"IPCA/IGPM/OUTROS",oFont12N,1400,100,CLR_BLACK,0,0)
		// oPrinter:SayAlign(nLinha,1800,"R$",oFont12,200,100,CLR_BLACK,0,0)
		// oPrinter:SayAlign(nLinha,1900,transform(nAcrsAno,"@E 99,999,999.99"),oFont12,300,100,CLR_BLACK,1,0)
		// oPrinter:Line(nPos , 0100, nPos, 2210, 0,"-1")
		// nPos += 80
		// nLinha += 80
		// nLinha2 += 80
		// If !Empty(nDSCAno)
		// 	oPrinter:SayAlign(nLinha,0100,"DESCONTO/NEGOCIACAO",oFont12N,1400,100,CLR_BLACK,0,0)
		// 	oPrinter:SayAlign(nLinha,1800,"R$",oFont12,200,100,CLR_BLACK,0,0)
		// 	oPrinter:SayAlign(nLinha,1900,transform(nDSCAno,"@E 99,999,999.99"),oFont12,300,100,CLR_BLACK,1,0)
		// 	oPrinter:Line(nPos , 0100, nPos, 2210, 0,"-1")
		// 	nPos += 80
		// 	nLinha += 80
		// 	nLinha2 += 80
		// EndIf
		// oPrinter:SayAlign(nLinha,0100,"JUROS MORA",oFont12N,1400,0,CLR_BLACK,0,0)
		// oPrinter:SayAlign(nLinha,1800,"R$",oFont12,200,100,CLR_BLACK,0,0)
		// oPrinter:SayAlign(nLinha,1900,transform(nValJuros,"@E 99,999,999.99"),oFont12,300,100,CLR_BLACK,1,0)
		// oPrinter:Line(nPos , 0100, nPos, 2210, 0,"-1")
		// nPos += 80
		// nLinha += 350

		// oPrinter:Say(nLinha,0700,"SANTA RITA DO PASSA QUATRO, "+substr(dtos(dDataBase),7,2)+" DE "+UPPER(MesExtenso(substr(dtos(dDataBase),5,2)))+" DE "+substr(dtos(dDataBase),1,4),oFont12)
		// nLinha+=200
		// oPrinter:Say(nLinha,0800,"____________________________________",oFont12)
		// nLinha+=50
		// oPrinter:Say(nLinha,0800,"ANGELICA FERRONATO SANTOS",oFont12)
		// nLinha+=50
		// oPrinter:Say(nLinha,0800,"CONTADORA",oFont12)
		// nLinha+=50
		// oPrinter:Say(nLinha,0800,"CRC 1 SP 250098/O-3",oFont12)
		// nLinha+=50
	Endif

Return

/* 
	Função para montar os vetores
*/
Static Function monta_vetores()
	Local Alias
	Local nPos
	Local cLote, cDescriLote, cZZSECUR
	Local cCliente, cLoja
	Local nDocAberto
	Local nDocLiquid
	Local nSaldo	
	Local nDesconto 
	Local nAcresc	
	Local nValLiq 	
	Local dVencRea	
	Local nAvencer	
	Local nVencido	
	Local nVencHoj  
	Local nNCCLiq
	Local nNCCSld
	Local cTitulo, cPrefix, cVetor, cTipo, aTit, cObsDesp
	Local xFilEnc   := Space(Len(SA1->A1_FILIAL))
	Local tFilial   := TRIM(SM0->M0_CODFIL)
  Local cMVZZFIENC:= GetMV("MV_ZZFIENC")

	If !Empty(cMVZZFIENC)
		If tFilial $ Substr(Trim(cMVZZFIENC),1+TAMSX3('A1_FILIAL')[1],TAMSX3('A1_FILIAL')[1])  
			Alert("Essa é uma Filial ENCERRADA, esse Relatório deve ser acessado pela Filial RESPONSÁVEL!")
			Return(.F.)
		EndIf
	EndIf

	
	If tFilial == Substr(Trim(cMVZZFIENC),1,TAMSX3('A1_FILIAL')[1])   
		xFilEnc := Substr(Trim(cMVZZFIENC),1+TAMSX3('A1_FILIAL')[1],TAMSX3('A1_FILIAL')[1])   
	EndIf

	cQuery := " SELECT SE1.E1_VENCTO VENCTO, SE1.E1_EMISSAO EMISSAO, SE1.E1_NUM NUM, SE1.E1_VALLIQ VALLIQ, SE1.E1_ZZSECUR ZZSECUR, SE1.E1_DTIGPM DTIGPM, "
	cQuery += " SE1.E1_PREFIXO PREFIXO, SE1.E1_PARCELA PARCELA,	SE1.E1_TIPO TIPO,  SE1.E1_JUROS JUROS,  SE1.E1_MULTA MULTA, SE1.E1_ZZINDIC ZZINDIC, "
	cQuery += " SE1.E1_VENCREA VENCREA, SE1.E1_VALOR VALOR, SE1.E1_DESCONT DESCONTO, SE1.E1_SALDO SALDO, SE1.E1_ACRESC ACRESC, SE1.E1_SDACRES SDACRES, "
	cQuery += " SE1.E1_CODOBSD CODOBSD,	SE1.E1_CLIENTE CLIENTE, SE1.E1_LOJA LOJA, SE1.E1_PRODUTO PRODUTO,   "
	cQuery += " SE1.E1_PORTADO BCO, SE1.E1_FIXVAR FV, SE1.E1_TIPIGPM TPIGPM, SB1.B1_COD CODPROD, SB1.B1_DESC DESCPROD,	"
	cQuery += " SE5.E5_TIPODOC TIPODOC, SE5.E5_MOTBX MOTBX , SE5.E5_DATA DATAMOV, SE5.E5_VALOR VALORMOV,	"
	cQuery += " SE5.E5_RECPAG RECPAG, "
	cQuery += " SA1.A1_NREDUZ NOMECLI,SA1.A1_CGC	"
	cQuery += " FROM " + retsqlname('SE1') + " SE1 "
	cQuery += " INNER JOIN " + retsqlname('SB1') + " SB1 ON SB1.B1_COD = SE1.E1_PRODUTO AND SB1.B1_FILIAL = '" + xFilial('SB1') + "' AND SB1.D_E_L_E_T_ = '' "
	cQuery += " INNER JOIN " + retsqlname('SA1') + " SA1 ON SA1.A1_COD = SE1.E1_CLIENTE AND SA1.A1_LOJA = SE1.E1_LOJA AND SA1.D_E_L_E_T_ = '' "
	cQuery += " LEFT OUTER JOIN " + retsqlname('SE5') + " SE5 " 

	If Empty(xFilEnc)
		cQuery += " ON SE5.E5_FILIAL = '" + xFilial('SE5') + "' "
	Else   
		cQuery += " ON SE5.E5_FILIAL IN ('" +tFilial+"','"+xFilEnc+"') "   //('44','67')
	EndIf

	cQuery += " 	AND SE5.E5_CLIFOR	= SE1.E1_CLIENTE   	"
	cQuery += " 	AND SE5.E5_PREFIXO	= SE1.E1_PREFIXO  	"
	cQuery += " 	AND SE5.E5_NUMERO	= SE1.E1_NUM 		"
	cQuery += " 	AND SE5.E5_PARCELA	= SE1.E1_PARCELA 	"
	cQuery += " 	AND SE5.E5_TIPO		= SE1.E1_TIPO 		"
	cQuery += " 	AND SE5.E5_TIPODOC	<> 'DB'		 		"   //DEBITO
	cQuery += " 	AND SE5.E5_TIPODOC	<> 'DC'	 		    "   //Desconto
	cQuery += " 	AND SE5.E5_TIPODOC	<> 'MT'	 		    "   //Multa
	cQuery += " 	AND SE5.E5_TIPODOC	<> 'JR'	 			"	//Juros
	cQuery += " 	AND SE5.E5_TIPODOC	<> 'ES'	 			"	//Estornos
	cQuery += " 	AND SE5.E5_VALOR <> 0	 		        "	//BAIXAS DE FORMAS DIVERSAS: CARTÓRIOS...ETC
	cQuery += " 	AND SE5.D_E_L_E_T_ = '' 				"

	If Empty(xFilEnc)
	cQuery += " WHERE SE1.E1_FILIAL = '" + xFilial('SE1') + "' "
	Else   
	cQuery += " WHERE SE1.E1_FILIAL IN ('" +tFilial+"','"+xFilEnc+"') "   
	EndIf

	cQuery += " AND SE1.E1_CLIENTE	= '" + cClienteDe  +  "' "
	cQuery += " AND SE1.E1_LOJA	    = '" + cLojaDe     +  "' "
	If !Empty(cProdutoDe)
	cQuery += " AND SE1.E1_PRODUTO	= '" + cProdutoDe +  "' "
	EndIf
	cQuery += " AND SE1.E1_PREFIXO	<> '" + 'ZZZ'	+  "' "
	cQuery += " AND SE1.D_E_L_E_T_ = '' ORDER BY PREFIXO, NUM,VENCTO, CLIENTE, DATAMOV "



	_aStru:={}//SE1SQL->(DbStruct())
	aadd( _aStru , {"PRF"       , "C" , 03 , 00 } )	
	aadd( _aStru , {"DOCUM"     , "C" , 06 , 00 } )	
	aadd( _aStru , {"PARCELA"   , "C" , 03 , 00 } )	
	aadd( _aStru , {"SITUA"     , "C" , 01 , 00 } )	
	aadd( _aStru , {"CODIGO"    , "C" , 06 , 00 } )
	aadd( _aStru , {"ORIGINAL"  , "N" , 12 , 02 } )
	aadd( _aStru , {"ACRESC"    , "N" , 12 , 02 } )
	aadd( _aStru , {"DESCONTO"  , "N" , 12 , 02 } )
	aadd( _aStru , {"VALOR"     , "N" , 12 , 02 } )
	aadd( _aStru , {"RECEBIDO"  , "N" , 12 , 02 } )

	
	AliasTMP:=GetNextAlias()    
	oTempTable := FWTemporaryTable():New(AliasTMP)	
	oTemptable:SetFields(_aStru)	
	oTempTable:Create()

	Alias 	:= ""
	cQuery 	:= ChangeQuery(cQuery)
	TCQuery cQuery new alias(Alias:=GetNextAlias())
	dbSelectArea(Alias)
	(Alias)->(dbGoTop())
	If Eof()
		MsgAlert("ARQUIVO VAZIO, VERIFIQUE OS PARAMETROS !!!")
		Return(.F.)
	EndIf
	While (Alias)->(!Eof())
		cZZSECUR    := (Alias)->ZZSECUR
		cLote 		:= (Alias)->CODPROD

		If cCliente == (Alias)->CLIENTE .and. cTitulo == (Alias)->NUM + (Alias)->PREFIXO + (Alias)->PARCELA
			nDesconto   := nDesconto + (Alias)->DESCONTO   
		Else
			nDesconto   := (Alias)->DESCONTO                
		EndIf

		cCliente	:= (Alias)->CLIENTE
		cLoja		:= (Alias)->LOJA
		cBco		:= (Alias)->BCO
		cMotbx      := (Alias)->MOTBX
		cDescriLote	:= (Alias)->DESCPROD
		cPrefix 	:= (Alias)->PREFIXO
		cTipo		:= (Alias)->TIPO
		nSaldo		:= ROUND((Alias)->SALDO,2)			
		nAcresc		:= ROUND((Alias)->ACRESC,2)			
		nValLiq 	:= ROUND((Alias)->VALORMOV,2)  		
		if ROUND((Alias)->JUROS,2) >= ROUND((Alias)->ACRESC,2) 
			nJuros := ROUND(((Alias)->MULTA+(Alias)->JUROS)-(Alias)->ACRESC,2)
		Else 
			// nJuros := ROUND((Alias)->MULTA+(Alias)->JUROS,2)
			nJuros := 0
		Endif   
		dVencRea	:= StoD((Alias)->VENCREA)	
		cTitulo     := (Alias)->NUM + (Alias)->PREFIXO	+ (Alias)->PARCELA
		nAvencer	:= 0					
		nVencido	:= 0   					
		nVencHoj	:= 0  					
		nNCCLiq     := 0                    
		nNCCSld     := 0                    

		If cTipo == "NCC"
			nDocAberto := 0
			nDocLiquid := 0
			nNCCSld := ROUND(nSaldo,2)   
			nSaldo 	:= 0	   	
		Else
			//Verifica se documento está em aberto
			If nSaldo > 0
				nDocAberto := 1
				nDocLiquid := 0
			Else
				nDocAberto := 0
				nDocLiquid := 1
			EndIf
		EndIf
		Do Case
			Case dVencRea > dDataBase	
				nAvencer := nSaldo + nAcresc
			Case dVencRea = dDataBase	
				nVencHoj := nSaldo + nAcresc
			Case dVencRea < dDataBase	
				nVencido := nSaldo + nAcresc
		EndCase
		
		nPos := aScan(aTitulos, {|x| x[1]+x[2]+x[3]+x[13]+x[14] == cTitulo + cCliente + cLoja})
		If nPos <> 0
			nAvencer 	:= 0
			nVencHoj 	:= 0
			nVencido 	:= 0
			nDocAberto 	:= 0
			nDocLiquid  := 0
			nNCCSld     := 0            
		EndIf
		If (Alias)->RECPAG == "P"
			nValLiq 	:= -nValLiq
		EndIf
		If cTipo == "NCC"
			nNCCLiq := nValLiq                    
			nValLiq := 0
		EndIf

		nPos := aScan(aLotes, {|x| x[1] == cLote})
		If nPos == 0
			AADD(aLotes, {cLote, cDescriLote, nAvencer, nVencHoj, nVencido, nDocAberto, (Alias)->ACRESC, nValLiq, nDocLiquid,"","",nNCCLiq, nNCCSld, cPrefix})
			nPos := Len(aLotes)
		Else
			aLotes[nPos][03] += nAvencer
			aLotes[nPos][04] += nVencHoj
			aLotes[nPos][05] += nVencido
			aLotes[nPos][06] += nDocAberto
			aLotes[nPos][08] += nValLiq
			aLotes[nPos][09] += nDocLiquid
			aLotes[nPos][12] += nNCCLiq
			aLotes[nPos][13] += nNCCSld
		EndIf
		aLotes[nPos][10] := cCliente
		aLotes[nPos][11] := cLoja

		
		cVetor := "aLotXCli"
		nPos := aScan(&cVetor, {|x| x[1] + x[2] + x[3] == cCliente + cLoja + cLote})
		If nPos == 0
			AADD(&cVetor, {cCliente, cLoja, cLote, nAvencer, nVencHoj, nVencido, nDocAberto, nDocLiquid, nValLiq, nNCCLiq, nNCCSld})
		Else
			&cVetor[nPos][04] += nAvencer
			&cVetor[nPos][05] += nVencHoj
			&cVetor[nPos][06] += nVencido
			&cVetor[nPos][07] += nDocAberto
			&cVetor[nPos][08] += nDocLiquid
			&cVetor[nPos][09] += nValLiq
			&cVetor[nPos][10] += nNCCLiq
			&cVetor[nPos][11] += nNCCSld
		EndIf

		cObsDesp := ""
		//Observação de Despesa
		If !Vazio(AllTrim((Alias)->CODOBSD))                 
			cObsDesp := MemoLine(MSMM((Alias)->CODOBSD,10))
			nPosPonto:= At(".",cObsDesp)
			If nPosPonto > 0
				//Imprime até encontrar o caractere "."
				cObsDesp := SubStr(cObsDesp,1,nPosPonto - 1)
			EndIf
		EndIf
		nValAcr := round((Alias)->VALOR,2)+round((Alias)->ACRESC,2)      
		aTit := {;
			(Alias)->NUM					,;      //01 - Número do Título
			(Alias)->PREFIXO				,;      //02 - Prefixo do Título
			(Alias)->PARCELA				,;      //03 - Parcela do Título
			(Alias)->TIPO					,;      //04 - Tipo do Título
			(Alias)->VENCTO					,;      //05 - Data de Vencimento
			(Alias)->EMISSAO				,;      //06 - Emissao do Título
			dVencRea						,;      //07 - Vencimento Real
			nValAcr							,;      //08 - Valor do Título
			round((Alias)->SALDO,2)			,;      //09 - Saldo em aberto do título
			round(nValLiq,2)				,;      //10 - Valor Liquidado do Título
			(Alias)->CODPROD   				,;  	//11 - Código Produto
			(Alias)->DESCPROD				,;  	//12 - Descrição Produto
			(Alias)->CLIENTE				,;		//13 - Código Cliente
			(Alias)->LOJA					,;  	//14 - Loja Cliente
			(Alias)->NOMECLI				,;  	//15 - Nome Cliente
			StoD((Alias)->DATAMOV)			,;		//16 - Data Movimentação
			(Alias)->TIPODOC				,;		//17 - Tipo Doc. -> Movimentacao Bancária
			round(nJuros,2)					,;		//18 - Valor Juros
			cObsDesp						,;		//19 - Código da observação de despesa
			cBco							,;		//20 - Banco
			Motbx                           ,;      //21 - Motivo Baixa
			round((Alias)->VALOR,2)			,;      //22 - Valor orig
			round((Alias)->ACRESC,2)		,;      //23 - Acrescimo
			(Alias)->FV						,;      //24 - Fixa / Variável
			(Alias)->TPIGPM					,;      //25 - Tipo de IGPM
			round((Alias)->SALDO,2)+round((Alias)->ACRESC,2),;  //26 - Saldo em aberto do título
			(Alias)->ZZSECUR			    ,;      //27 - Securitizadora
			(Alias)->ZZINDIC			    ,;      //28 - Indice: 1-IPCA/2-IGPM/3-INCC/6-Fixa
			(Alias)->DESCONTO			    ,;      //29 - Valor Desconto
			(Alias)->DTIGPM 			    ,;      //30 - Data da Correção do ìndice (IGPM/IPCA/INCC)
			(Alias)->A1_CGC			    	,;      //31 - Data da Correção do ìndice (IGPM/IPCA/INCC)
			substr(dtos(dVencRea),5,2)			;		//32 - Data Movimentação
		}

		
		AADD(aTitulos, aClone(aTit))                                                                      
		xValor := (Alias)->SALDO 

		dbSelectArea(AliasTMP)
		Locate for trim((AliasTMP)->DOCUM) == TRIM((Alias)->NUM) .AND. TRIM((AliasTMP)->CODIGO)==TRIM((Alias)->CLIENTE) .AND. TRIM((AliasTMP)->PARCELA)==TRIM((Alias)->PARCELA) .AND. (AliasTMP)->ORIGINAL==(Alias)->VALOR
		if !found()
		Reclock(AliasTMP,.T.)
		(AliasTMP)->PRF     := (Alias)->PREFIXO
		(AliasTMP)->DOCUM   := (Alias)->NUM
		(AliasTMP)->PARCELA := (Alias)->PARCELA
		(AliasTMP)->CODIGO  := (Alias)->CLIENTE
		(AliasTMP)->ORIGINAL:= (Alias)->VALOR
		(AliasTMP)->ACRESC  := (Alias)->ACRESC
		(AliasTMP)->DESCONTO:= (Alias)->DESCONTO
		(AliasTMP)->VALOR   := (Alias)->SALDO+(Alias)->ACRESC
		(AliasTMP)->RECEBIDO:= nValLiq
		else
		Reclock(AliasTMP,.F.)
		(AliasTMP)->RECEBIDO:= ((AliasTMP)->RECEBIDO + nValLiq)
		(AliasTMP)->DESCONTO:= (Alias)->DESCONTO
		Endif
		MsUnlock()
		(Alias)->(DbSkip())
	EndDo                   
	(Alias)->(dbCloseArea())
	(AliasTMP)->(DbCloseArea())
	oTempTable:Delete()
Return

/*
	Função para buscar valor Contrato
*/
Static Function GetValorContrato(cCliente,cProduto)
	Local nValor:= 0
	Local nTotal:= 0
	Local nVal  := 1

	cConsul := " SELECT DISTINCT E1_PREFIXO, E1_NUM, E1_TIPO	"
	cConsul += " FROM "+RetSqlTab("SE1")
	cConsul += " WHERE 1=1								"
	cConsul += " AND "+RetSqlFil("SE1")
	cConsul += " AND "+RetSqlDel("SE1")
	cConsul += " AND E1_CLIENTE = '"+cCliente+"' 		"
	cConsul += " AND E1_PRODUTO = '"+cProduto+"' 		"
	cConsul += " AND E1_BAIXA   <= '"+cAnoBase+"1231' 	"
	cConsul += " AND E1_EMISSAO <= '"+cAnoBase+"1231' 	"
	cConsul += " AND SE1.E1_TIPO <> 'RC' 				"
	cConsul += " AND SE1.E1_PREFIXO <> 'ZZZ' 			"

	TCQuery cConsul NEW ALIAS "TCONSUL"
	DbSelectArea("TCONSUL")

	If ("TCONSUL")->(!EOF())
		Count To nTotal
		("TCONSUL")->(DbGoTop())

		while ("TCONSUL")->(!EOF())
			If(nVal < nTotal)
				cSQL:= " SELECT SUM(E1_VALOR + E1_ACRESC) AS E1_VALOR FROM "+RetSqlTab("SE1")
				cSQL+= " LEFT OUTER JOIN " +RetSqlTab("SE5")
				cSQL+= " 		ON "+ RetSqlFil("SE5")
				cSQL+= " 			AND SE5.E5_CLIFOR	= SE1.E1_CLIENTE 						" 	
				cSQL+= " 			AND SE5.E5_PREFIXO	= SE1.E1_PREFIXO 						" 	
				cSQL+= " 			AND SE5.E5_NUMERO	= SE1.E1_NUM 							"	
				cSQL+= " 			AND SE5.E5_PARCELA	= SE1.E1_PARCELA 						" 	
				cSQL+= " 			AND SE5.E5_TIPO		= SE1.E1_TIPO 							"
				cSQL+= " 			AND SE5.E5_TIPODOC	NOT IN ('DB','DC','MT','JR','ES') 		"		        
				cSQL+= " 			AND SE5.D_E_L_E_T_ = '' 									"
				cSQL+= " WHERE 1=1 																"
				cSQL+= " AND "+RetSqlFil("SE1")
				cSQL+= " AND "+RetSqlDel("SE1")
				cSQL+= " AND SE1.E1_CLIENTE = '"+cCliente+"'									"
				cSQL+= " AND SE1.E1_PRODUTO = '"+cProduto+"'									"
				cSQL+= " AND SE1.E1_NUM = '"+("TCONSUL")->E1_NUM+"'								"
				cSQL+= " AND SE1.E1_PREFIXO = '"+("TCONSUL")->E1_PREFIXO+"'						"
				cSQL+= " AND SE1.E1_TIPO = '"+("TCONSUL")->E1_TIPO+"'							"
				cSQL+= " AND E5_MOTBX = 'NOR' 													"
				cSQL+= " AND SE1.E1_TIPO <> 'RC'												"
				
				TCQuery cSQL NEW ALIAS "TVALOR"
				DbSelectArea("TVALOR")

				If ("TVALOR")->(!EOF())
					While ("TVALOR")->(!EOF()) 
						nValor += ("TVALOR")->E1_VALOR
						("TVALOR")->(dbSkip())
					Enddo
				Endif

				("TVALOR")->(dbCloseArea())

				nVal ++
			Else
				cPreNuTp+= "'"+("TCONSUL")->E1_PREFIXO+("TCONSUL")->E1_NUM+("TCONSUL")->E1_TIPO+"'"
				cSQL:= " SELECT SUM(E1_VALOR) AS E1_VALOR FROM "+RetSqlTab("SE1")
				cSQL+= " WHERE 1=1 "
				cSQL+= " AND "+RetSqlFil("SE1")
				cSQL+= " AND "+RetSqlDel("SE1")
				cSQL+= " AND E1_CLIENTE = '"+cCliente+"'"
				cSQL+= " AND E1_PRODUTO = '"+cProduto+"'"
				cSQL+= " AND E1_NUM = '"+("TCONSUL")->E1_NUM+"'"
				cSQL+= " AND SE1.E1_PREFIXO = '"+("TCONSUL")->E1_PREFIXO+"'"
				cSQL+= " AND SE1.E1_TIPO = '"+("TCONSUL")->E1_TIPO+"'"
				cSQL+= " AND E1_TIPO <> 'RC'"
				
				TCQuery cSQL NEW ALIAS "TVALOR"
				DbSelectArea("TVALOR")

				If ("TVALOR")->(!EOF())
					While ("TVALOR")->(!EOF()) 
						nValor+= ("TVALOR")->E1_VALOR
						("TVALOR")->(dbSkip())
					Enddo
				Endif

				("TVALOR")->(dbCloseArea())
			Endif
			("TCONSUL")->(dbSkip())
		end		
	Endif 

	("TCONSUL")->(dbCloseArea())
Return nValor

/*
	Adiciona parâmetros da Pergunta Inicial
*/
Static Function Perguntas()
    Local aPergs    := {}
    Local cCli  := Space(Len(AvKey("","A1_COD")))
    Local cLoj  := Space(Len(AvKey("","A1_LOJA")))
    Local cPro  := Space(Len(AvKey("","B1_COD")))
    Local cAno  := Space(4)

     
    //Adiciona os parâmetros
    aadd(aPergs, {1, "Codigo do cliente" , cCli , "", ".T.", "SA1"   , ".T.", 30, .T.})
    aadd(aPergs, {1, "Loja"              , cLoj , "", ".T.", ""      , ".T.", 30, .T.})
    aadd(aPergs, {1, "Código do Produto" , cPro , "", ".T.", "SB1"   , ".T.", 70, .T.})
    aadd(aPergs, {1, "Ano base"       , cAno , "", ".T.", ""      , ".T.", 30, .T.})
     
    //Se a pergunta foi confirmada
    If ParamBox(aPergs, "Informe os parâmetros", /*aRet*/, /*bOK*/, /*aButtons*/, /*lCentered*/, /*nPosX*/, /*nPosY*/, /*oDlgWizard*/, /*cLoad*/, .F., .F.)
        Return .T.
    EndIf

Return .F.

/*
	Função para gerar base64 para portal
*/
Static Function openFile(cDirOS, cFileName)
	Local oReturn := JsonObject():New()
	Local cBuffer := ""
	Local nHandle := 0
	Local aTam    := {}

	cFileName := StrTran(cFileName,".rel",".pdf")

	aDir(cDirOS+cFileName,,@aTam)
	
	nHandle := FOPEN(cDirOS+cFileName)
	FREAD(nHandle,cBuffer,aTam[1])
	FCLOSE(nHandle)

	oReturn['status'] := 'sucesso'
	oReturn['data'] := Encode64(cBuffer)

Return oReturn
