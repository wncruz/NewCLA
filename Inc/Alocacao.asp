<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: AlocacaoFac.asp
'	- Descrição			: Funções para alocação de facilidades

'Monta Xml da Facilidade
Function MontarXmlFacilidade(objXml,objRS,strStatus,intTipoProcesso,strTipoConsulta)

	Dim objNodeAcesso
	Dim objElemento
	Dim intIndice 
	Dim objDic
	Dim strChave
	Dim strOrderBy
	
	Set objDic = Server.CreateObject("Scripting.Dictionary") 
	
	'Cria os elementos filhos para xmlDados
	intIndice = 1000
	if Not objRS.Eof and Not objRS.Bof then
		While not  objRS.EOF

			if isNull(objRS("Fac_Representacao")) then
				Select Case objRS("Sis_Id")
					Case 1
						strChave = objRS("Fac_TimeSlot")
						strOrderBy = "0000-" & objRS("Fac_TimeSlot")
					Case Else	
						strChave = objRS("Fac_Par")
						strOrderBy = "0000-" & objRS("Fac_Par")
				End Select		
			Else
				strChave = objRS("Fac_Representacao")
				strOrderBy = strChave
			End if

			if  Not objDic.Exists(strChave) then

				Call objDic.Add(strChave,objRS("Fac_Id")) 
				'Cria o nível fluxo
				Set objNodeAcesso = objXML.createNode("element", "Facilidade", "")
				objXML.documentElement.appendChild (objNodeAcesso)
				if isNull(objRS("Int_id")) then
					strStatus = "PENDENTE"
				End if	
				Call AddElemento(objXML,objNodeAcesso,"intIndice",intIndice)
				Call AddElemento(objXML,objNodeAcesso,"Fac_Id",objRS("Fac_Id"))
				Call AddElemento(objXML,objNodeAcesso,"Int_Id",objRS("Int_Id"))
				Call AddElemento(objXML,objNodeAcesso,"Int_CorOrigem",objRS("Int_CorOrigem"))
				Call AddElemento(objXML,objNodeAcesso,"Int_CorOrigemAntes",objRS("Int_CorOrigem"))
				Call AddElemento(objXML,objNodeAcesso,"Int_CorDestino",objRS("Int_CorDestino"))
				Call AddElemento(objXML,objNodeAcesso,"Ped_Id",objRS("Ped_ID"))
				Call AddElemento(objXML,objNodeAcesso,"Acf_Id",objRS("Acf_Id"))
				Call AddElemento(objXML,objNodeAcesso,"cboCodProv",objRS("Acf_CodProvedor"))
				Set ObjRSProv = db.Execute("CLA_sp_sel_provedor " & objRS("Acf_CodProvedor"))
				if Not ObjRSProv.Eof and Not ObjRSProv.Bof then
					Call AddElemento(objXML,objNodeAcesso,"cboCodProvText",ObjRSProv("Pro_Cod"))
				End if	
				Call AddElemento(objXML,objNodeAcesso,"txtNroAcessoEbt",objRS("Acf_NroAcessoPtaEbt"))
				Call AddElemento(objXML,objNodeAcesso,"txtFila",objRS("Fac_Fila"))
				Call AddElemento(objXML,objNodeAcesso,"txtBastidor",objRS("Fac_Bastidor"))
				Call AddElemento(objXML,objNodeAcesso,"txtRegua",objRS("Fac_Regua"))
				Call AddElemento(objXML,objNodeAcesso,"txtPosicao",objRS("Fac_Posicao"))
				Call AddElemento(objXML,objNodeAcesso,"txtNroAcessoCli",objRS("Acf_NroAcessoPtaCli"))
				Call AddElemento(objXML,objNodeAcesso,"cboCodProv",objRS("Acf_CodProvedor"))
				Call AddElemento(objXML,objNodeAcesso,"txtCCTOPro",objRS("Acf_NroAcessoCCTOProvedor"))
				Call AddElemento(objXML,objNodeAcesso,"txtCNLPtaA",objRS("Acf_CnlPta"))
				Call AddElemento(objXML,objNodeAcesso,"txtCNLPtaB",objRS("Acf_CnlPtB"))
				Call AddElemento(objXML,objNodeAcesso,"rdoUrbano",objRS("Acf_CCTOTipo"))
				
				Select Case objRS("Acf_CCTOTipo")
					Case "I"
						Call AddElemento(objXML,objNodeAcesso,"rdoUrbanoIndex","0")
					Case "U"
						Call AddElemento(objXML,objNodeAcesso,"rdoUrbanoIndex","1")
				End Select	
				Call AddElemento(objXML,objNodeAcesso,"cboRede",objRS("Sis_Id"))
				Call AddElemento(objXML,objNodeAcesso,"txtTronco",objRS("Fac_Tronco"))
				Call AddElemento(objXML,objNodeAcesso,"txtCabo",objRS("Fac_Tronco"))
				Call AddElemento(objXML,objNodeAcesso,"txtLateral",objRS("Fac_Lateral"))
				Call AddElemento(objXML,objNodeAcesso,"cboTipoCabo",objRS("Fac_TipoCabo"))
				Call AddElemento(objXML,objNodeAcesso,"txtCaixaEmenda",objRS("Fac_CxEmenda"))
				Call AddElemento(objXML,objNodeAcesso,"TipoAcao","A")
				Call AddElemento(objXML,objNodeAcesso,"OrderBy",strOrderBy)

				if isNull(objRS("Fac_Representacao")) then
					Call AddElemento(objXML,objNodeAcesso,"txtPar",objRS("Fac_Par"))
					Call AddElemento(objXML,objNodeAcesso,"txtTimeslot",objRS("Fac_TimeSlot"))
					Call AddElemento(objXML,objNodeAcesso,"Fac_Representacao",objRS("Fac_Representacao"))
				Else
					Call AddElemento(objXML,objNodeAcesso,"txtPar",objRS("Fac_Representacao"))
					Call AddElemento(objXML,objNodeAcesso,"txtTimeslot",objRS("Fac_Representacao"))
					Call AddElemento(objXML,objNodeAcesso,"Fac_Representacao",objRS("Fac_Representacao"))
				End if

				Call AddElemento(objXML,objNodeAcesso,"Pro_ID",objRS("Pro_ID"))
				Call AddElemento(objXML,objNodeAcesso,"Prm_ID",objRS("Prm_ID"))
				Call AddElemento(objXML,objNodeAcesso,"Reg_ID",objRS("Reg_ID"))
				Call AddElemento(objXML,objNodeAcesso,"Dst_ID",objRS("Dst_ID"))
				Call AddElemento(objXML,objNodeAcesso,"Sis_ID",objRS("Sis_ID"))
				Call AddElemento(objXML,objNodeAcesso,"Esc_ID",objRS("Esc_Id"))

				Call AddElemento(objXML,objNodeAcesso,"cboPropModem",objRS("Acf_ProprietarioEquip"))
				Call AddElemento(objXML,objNodeAcesso,"txtQtdeModem",objRS("Acf_QtdEquip"))
				Select Case objRS("Acf_CCTOFatura")
					Case "S"
						Call AddElemento(objXML,objNodeAcesso,"rdoFaturaIndex",0)
					Case "N" 
						Call AddElemento(objXML,objNodeAcesso,"rdoFaturaIndex",1)
				End Select
				Call AddElemento(objXML,objNodeAcesso,"rdoFatura",objRS("Acf_CCTOFatura"))
				Call AddElemento(objXML,objNodeAcesso,"txtObsFac",TratarAspasXml(objRS("Acf_Obs")))
				Call AddElemento(objXML,objNodeAcesso,"strTipoConsulta",Cstr(strTipoConsulta))

				if not isNull(objRS("Int_CorOrigem")) then
					Set objNodeInter = objXML.createNode("element", "Interligacao", "")
					objNodeAcesso.appendChild (objNodeInter)
					'Adiciona as interligações
					Call AddElemento(objXML,objNodeInter,"Int_Id",objRS("Int_Id"))
					Call AddElemento(objXML,objNodeInter,"Int_CorOrigem",objRS("Int_CorOrigem"))
					Call AddElemento(objXML,objNodeInter,"Int_CorOrigemAntes",objRS("Int_CorOrigem"))
					Call AddElemento(objXML,objNodeInter,"Int_CorDestino",objRS("Int_CorDestino"))
					Call AddElemento(objXML,objNodeInter,"Fac_Id",objRS("Fac_Id"))
					Call AddElemento(objXML,objNodeInter,"Ped_Id",objRS("Ped_ID"))
					Call AddElemento(objXML,objNodeInter,"Acf_Id",objRS("Acf_Id"))
					Call AddElemento(objXML,objNodeInter,"txtNroAcessoEbt",objRS("Acf_NroAcessoPtaEbt"))
					Call AddElemento(objXML,objNodeInter,"cboRede",objRS("Sis_Id"))
				End if
				
				Call AddElemento(objXML,objNodeAcesso,"txtVlan",objRS("Fac_Vlan"))
				Call AddElemento(objXML,objNodeAcesso,"txtPE",objRS("Fac_PE"))
				Call AddElemento(objXML,objNodeAcesso,"txtPorta",objRS("Fac_Porta"))
				Call AddElemento(objXML,objNodeAcesso,"txtLink",objRS("Fac_Link"))


				intIndice = intIndice + 1
			Else
				'Adiciona as interligações
				if not isNull(objRS("Int_CorOrigem")) then
					Set objNodeInter = objXML.createNode("element", "Interligacao", "")
					objNodeAcesso.appendChild (objNodeInter)
					Call AddElemento(objXML,objNodeInter,"Int_Id",objRS("Int_Id"))
					Call AddElemento(objXML,objNodeInter,"Int_CorOrigem",objRS("Int_CorOrigem"))
					Call AddElemento(objXML,objNodeInter,"Int_CorOrigemAntes",objRS("Int_CorOrigem"))
					Call AddElemento(objXML,objNodeInter,"Int_CorDestino",objRS("Int_CorDestino"))
					Call AddElemento(objXML,objNodeInter,"Ped_Id",objRS("Ped_ID"))
					Call AddElemento(objXML,objNodeInter,"Acf_Id",objRS("Acf_Id"))
					Call AddElemento(objXML,objNodeInter,"Fac_Id",objRS("Fac_Id"))
					Call AddElemento(objXML,objNodeInter,"txtNroAcessoEbt",objRS("Acf_NroAcessoPtaEbt"))
					Call AddElemento(objXML,objNodeInter,"cboRede",objRS("Sis_Id"))
				End if	
			End if

			objRS.MoveNext
			if strStatus = "" then strStatus = "EXECUTADO"
		Wend
	End if

	Set objNodeAcesso = Nothing
	Set objElemento = Nothing
	
	Set MontarXmlFacilidade = objXML

End Function
%>