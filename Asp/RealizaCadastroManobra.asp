<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/EnviarManobraAprov.asp"-->
<%
'•ACCENTURE
'	- Sistema			: CLA
'	- Arquivo			: RealizaCadastroManobra.asp
'	- Responsável		: Gustavo S. Reynaldo
'	- Descrição			: Chama a SP de cadastro de manobra
%>
<%
	dim objXmlDoc, strCid,strRet , strRetorno, objRSPro, retorno
	dim ndCCTOTIPO,ndTronco, ndPar
	dim ndFac,ndPed,ndAcf,ndRepresentacao,ndRec,ndFila,ndBastidor,ndRegua,ndPosicao,ndTimeslot,ndNumAcessoPtaEbt,ndNumAcessoCLI,ndCodProv,ndCCTOProvedor,ndCNLPontaA,ndCNLPontaB,ndSisID,ndPropModem,ndQtdeModem,ndFatura,ndAreaObs,ndProID
	
	'Criação dos objetos
	set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
		
	'Atribuição de valores para as variáveis 	
	objXmlDoc.load(Request)
	
	set ndFac 	=  objXmlDoc.selectSingleNode("//FacID")
	set ndPed  	=  objXmlDoc.selectSingleNode("//PedID")
	set ndAcf   =  objXmlDoc.selectSingleNode("//AcfID")	
	set ndRepresentacao  =  objXmlDoc.selectSingleNode("//Representacao")	
	set ndRec   =  objXmlDoc.selectSingleNode("//Rec")
	set ndFila   =  objXmlDoc.selectSingleNode("//Fila")
	set ndBastidor   =  objXmlDoc.selectSingleNode("//Bastidor")
	set ndRegua   =  objXmlDoc.selectSingleNode("//Regua")
	set ndPosicao   =  objXmlDoc.selectSingleNode("//Posicao")
	set ndTimeslot   =  objXmlDoc.selectSingleNode("//Timeslot")
	set ndNumAcessoPtaEbt   =  objXmlDoc.selectSingleNode("//NumAcessoPtaEbt")
	set ndNumAcessoCLI   =  objXmlDoc.selectSingleNode("//NumAcessoCLI")
	set ndCodProv   =  objXmlDoc.selectSingleNode("//CodProv")
	set ndCCTOProvedor   =  objXmlDoc.selectSingleNode("//CCTOProvedor")
	set ndCNLPontaA   =  objXmlDoc.selectSingleNode("//CNLPontaA")
	set ndCNLPontaB   =  objXmlDoc.selectSingleNode("//CNLPontaB")
	set ndSisID   =  objXmlDoc.selectSingleNode("//SisID")
	set ndPropModem   =  objXmlDoc.selectSingleNode("//PropModem")	
	set ndQtdeModem   =  objXmlDoc.selectSingleNode("//QtdeModem")		
	set ndFatura   =  objXmlDoc.selectSingleNode("//Fatura")	
	set ndAreaObs   =  objXmlDoc.selectSingleNode("//AreaObs")
	set ndPortadora   =  objXmlDoc.selectSingleNode("//Portadora")
	set ndTronco   =  objXmlDoc.selectSingleNode("//Tronco")
	set ndPar   =  objXmlDoc.selectSingleNode("//Par")
	set ndCabo   =  objXmlDoc.selectSingleNode("//Cabo")
	set ndTipoCabo   =  objXmlDoc.selectSingleNode("//TipoCabo")
	set ndCxEmenda   =  objXmlDoc.selectSingleNode("//CxEmenda")
	set ndCCTOTIPO  = objXmlDoc.selectSingleNode("//CCTOTipo")
	set ndCabo  = objXmlDoc.selectSingleNode("//Cabo")
	set ndDerivacao  = objXmlDoc.selectSingleNode("//Derivacao")
	set ndPade  = objXmlDoc.selectSingleNode("//Pade")
	set ndProID  = objXmlDoc.selectSingleNode("//ProID")
	
	if ndFac is nothing then
		Vetor_Campos(1)="adInteger,2,adParamInput,null"
	else	
		Vetor_Campos(1)="adInteger,2,adParamInput," & ndFac.Text	
	end if
	if ndPed is nothing then
		Vetor_Campos(2)="adInteger,2,adParamInput,null"
	else	
		Vetor_Campos(2)="adInteger,2,adParamInput," & ndPed.Text	 
	end if
	if ndAcf is nothing then
		Vetor_Campos(3)="adInteger,2,adParamInput,null"
	else	
		Vetor_Campos(3)="adInteger,2,adParamInput," & ndAcf.Text	
	end if
	if ndRepresentacao is nothing then
		Vetor_Campos(4)="adWChar,9,adParamInput,null"
	else	
		Vetor_Campos(4)="adWChar,9,adParamInput," & ndRepresentacao.Text
	end if
	if ndFila is nothing then
		Vetor_Campos(5)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(5)="adWChar,20,adParamInput," & ndFila.Text
	end if
	if ndBastidor is nothing then
		Vetor_Campos(6)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(6)="adWChar,20,adParamInput," & ndBastidor.Text
	end if
	if ndRegua is nothing then
		Vetor_Campos(7)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(7)="adWChar,20,adParamInput," & ndRegua.Text
	end if
	if ndPosicao is nothing then
		Vetor_Campos(8)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(8)="adWChar,20,adParamInput," & ndPosicao.Text
	end if
	if ndTimeslot is nothing then
		Vetor_Campos(9)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(9)="adWChar,20,adParamInput," & ndTimeslot.Text
	end if
	if ndNumAcessoPtaEbt is nothing then
		Vetor_Campos(10)="adWChar,25,adParamInput,null"
	else	
		Vetor_Campos(10)="adWChar,25,adParamInput," & ndNumAcessoPtaEbt.Text
	end if
	if ndNumAcessoCLI is nothing then
		Vetor_Campos(11)="adWChar,25,adParamInput,null"
	else	
		Vetor_Campos(11)="adWChar,25,adParamInput," & ndNumAcessoCLI.Text	
	end if
	if ndCodProv is nothing then
		Vetor_Campos(12)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(12)="adWChar,20,adParamInput," & ndCodProv.Text
	end if
	if ndCCTOProvedor is nothing then
		Vetor_Campos(13)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(13)="adWChar,20,adParamInput," & ndCCTOProvedor.Text
	end if
	if ndCNLPontaA is nothing then
		Vetor_Campos(14)="adWChar,4,adParamInput,null"
	else	
		Vetor_Campos(14)="adWChar,4,adParamInput," & ndCNLPontaA.Text	
	end if
	if ndCNLPontaB is nothing then
		Vetor_Campos(15)="adWChar,4,adParamInput,null"
	else	
		Vetor_Campos(15)="adWChar,4,adParamInput," & ndCNLPontaB.Text
	end if
	if ndPortadora is nothing then
		Vetor_Campos(16)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(16)="adWChar,20,adParamInput," & ndPortadora.Text
	end if
	if ndCCTOTIPO is nothing then
		Vetor_Campos(17)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(17)="adWChar,20,adParamInput," & ndCCTOTIPO.Text
	end if
	Vetor_Campos(18)="adInteger,4,adParamInput," & ndSisID.Text
	if ndTronco is nothing then
		Vetor_Campos(19)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(19)="adWChar,20,adParamInput," & ndTronco.Text
	end if
	if ndPar is nothing then
		Vetor_Campos(20)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(20)="adWChar,20,adParamInput," & ndPar.Text
	end if
	if ndCabo is nothing then
		Vetor_Campos(21)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(21)="adWChar,20,adParamInput," & ndCabo.Text
	end if
	if ndDerivacao is nothing then
		Vetor_Campos(22)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(22)="adWChar,20,adParamInput," & ndDerivacao.Text
	end if
	if ndTipoCabo is nothing then
		Vetor_Campos(23)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(23)="adWChar,20,adParamInput," & ndTipoCabo.Text
	end if
	if ndPade is nothing then
		Vetor_Campos(24)="adWChar,20,adParamInput,null"
	else	
		Vetor_Campos(24)="adWChar,20,adParamInput," & ndPade.Text
	end if
	if ndPropModem is nothing then
		Vetor_Campos(25)="adWChar,3,adParamInput,null"
	else	
		Vetor_Campos(25)="adWChar,3,adParamInput," & ndPropModem.Text 
	end if
	if ndQtdeModem is nothing then
		Vetor_Campos(26)="adInteger,2,adParamInput,null"
	else	
		Vetor_Campos(26)="adInteger,2,adParamInput," & ndQtdeModem.Text 
	end if
	if ndFatura is nothing then
		Vetor_Campos(27)="adWChar,1,adParamInput,null"
	else	
		Vetor_Campos(27)="adWChar,1,adParamInput," & ndFatura.Text 
	end if
	if ndAreaObs is nothing then
		Vetor_Campos(28)="adWChar,300,adParamInput,null"
	else	
		Vetor_Campos(28)="adWChar,300,adParamInput," & ndAreaObs.Text
	end if
	if ndProID is nothing then
		Vetor_Campos(29)="adInteger,4,adParamInput,null"
	else	
		Vetor_Campos(29)="adInteger,4,adParamInput," & ndProID.Text
	end if
	Vetor_Campos(30)="adInteger,4,adParamInput," & dblUsuId
	Vetor_Campos(31)="adInteger,4,adParamInput," & ndRec.Text 
	Vetor_Campos(32)="adInteger,4,adParamOutput,0"
					
	strSql = APENDA_PARAMSTRSQL("Cla_sp_upd_FacilidadeManobra",32,Vetor_Campos)

	'Call APENDA_PARAM("Cla_sp_upd_FacilidadeManobra",31,Vetor_Campos)
	
	Set objRSRec = db.Execute(strSql)
	
	'Set DBAction = objRSRec("ret")
	'DBAction = 0
	
	'DBAction = ObjCmd.Parameters("RET").value
	
		
	If objRSRec("ret") = 0 or objRSRec("ret") = 159 then
		strRet = "Sucesso"
		Vetor_Campos(1)="adWChar,15,adParamInput, null " 
		Vetor_Campos(2)="adInteger,2,adParamInput," & ndAcf.Text
		strSqlRet = APENDA_PARAMSTR("CLA_sp_view_solicitacaoAprov",2,Vetor_Campos)
		
		Set objRSDadosCla = db.Execute(strSqlRet)			
		If not objRSDadosCla.Eof and  not objRSDadosCla.Bof Then
			EnviarManobraAprov(ndAcf.Text)
		end if
	elseif objRSRec("ret") = 60 then
		strRet = "Uma das Facilidades digitadas não existe, favor alterar"
	elseif objRSRec("ret") = 161 then
		strRet = "Uma das Facilidades digitadas já está alocada outro pedido, favor alterar"
	else
		strRet = "Erro"
	End if
		
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	
	Response.Write(strRet)	
%>