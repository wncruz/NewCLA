<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: xmlAcessos.asp
'	- Descrição			: Cria xml com pontos de uma solicitação

Function MontarXmlAcesso(objXml,objRS,dlbIdAcessoLog)

	Dim objNodeAcesso
	Dim objElemento
	Dim objDicAcf
	Dim objDictionary

	Set objDicAcf = Server.CreateObject("Scripting.Dictionary")
	Set objDictionary = Server.CreateObject("Scripting.Dictionary")

	
	'Cria os elementos filhos para xmlDados
	intIndice = 1
	While not objRS.EOF
				
			'Cria o nível fluxo
			Set objNodeAcesso = objXML.createNode("element", "Acesso", "")
			objXML.documentElement.appendChild (objNodeAcesso)

			Call AddElemento(objXML,objNodeAcesso,"intIndice",intIndice)
			Call AddElemento(objXML,objNodeAcesso,"rdoPropAcessoFisicoIndex",indexPropAcesso(objRS("Acf_Proprietario")))
			Call AddElemento(objXML,objNodeAcesso,"rdoPropAcessoFisico",objRS("Acf_Proprietario"))
			Call AddElemento(objXML,objNodeAcesso,"cboTecnologia",objRS("newtec_id"))
			
			Call AddElemento(objXML,objNodeAcesso,"txtFacilidade",objRS("newfac_nome"))
			Call AddElemento(objXML,objNodeAcesso,"newfac_id",objRS("newfac_id"))
			
			Call AddElemento(objXML,objNodeAcesso,"cboTipoRadio",objRS("Trd_Id"))
			Call AddElemento(objXML,objNodeAcesso,"cboVersaoRadio",objRS("Acf_VersaoRadio"))
			Call AddElemento(objXML,objNodeAcesso,"cboVelAcesso",objRS("Vel_IdAcessoFis"))
			Call AddElemento(objXML,objNodeAcesso,"cboTipoVel",objRS("Acf_TipoVel"))
			Call AddElemento(objXML,objNodeAcesso,"cboTipoVelText",TipoVel(objRS("Acf_TipoVel")))
			Call AddElemento(objXML,objNodeAcesso,"txtSenha",objRS("Aprov_Senha"))

			'GPON
			Call AddElemento(objXML,objNodeAcesso,"cboFabricanteONT",objRS("Font_ID"))
			Call AddElemento(objXML,objNodeAcesso,"cboTipoONT",objRS("Tont_ID"))

			Call AddElemento(objXML,objNodeAcesso,"cboProvedor",objRS("Pro_Id"))
			if objRS("Pro_Id") <> "" then
				Set objRSAux = db.Execute("CLA_Sp_Sel_Provedor " & objRS("Pro_Id"))
				if not objRSAux.Eof and Not objRSAux.Bof then
					Call AddElemento(objXML,objNodeAcesso,"cboProvedorText",objRSAux("pro_nome"))
				End if
			End if

			if objRS("Vel_IdAcessoFis") <> "" then
				Set objRSAux = db.Execute("CLA_Sp_Sel_Velocidade " & objRS("Vel_IdAcessoFis"))
				if not objRSAux.Eof and Not objRSAux.Bof then
					Call AddElemento(objXML,objNodeAcesso,"cboVelAcessoText",Trim(objRSAux("Vel_Desc")))
				End if
			End if

			Call AddElemento(objXML,objNodeAcesso,"cboRegimeCntr",objRS("Reg_Id"))
			Call AddElemento(objXML,objNodeAcesso,"cboPromocao",objRS("Prm_Id"))
			Call AddElemento(objXML,objNodeAcesso,"cboUFEnd",objRS("Est_Sigla"))
			Call AddElemento(objXML,objNodeAcesso,"cboUFEndText",objRS("Est_Sigla"))
			Call AddElemento(objXML,objNodeAcesso,"txtEndCid",objRS("Cid_Sigla"))
			Call AddElemento(objXML,objNodeAcesso,"txtEndCidDesc",objRS("Cid_Desc"))
			Call AddElemento(objXML,objNodeAcesso,"cboLogrEnd",Trim(objRS("Tpl_Sigla")))
			Call AddElemento(objXML,objNodeAcesso,"cboLogrEndText",objRS("Tpl_Sigla"))
			Call AddElemento(objXML,objNodeAcesso,"txtEnd",objRS("End_NomeLogr"))
			Call AddElemento(objXML,objNodeAcesso,"txtNroEnd",objRS("End_NroLogr"))
			Call AddElemento(objXML,objNodeAcesso,"txtBairroEnd",objRS("End_Bairro"))
			Call AddElemento(objXML,objNodeAcesso,"txtCepEnd",objRS("End_CEP"))
			Call AddElemento(objXML,objNodeAcesso,"txtContatoEnd",objRS("Aec_Contato"))
			if not isNull(objRS("Aec_Telefone")) then
				strTel = objRS("Aec_Telefone")
				strTelLocal = Right(objRS("Aec_Telefone"),len(strTel)-2)
				strTelArea  = Left(objRS("Aec_Telefone"),2)
				Call AddElemento(objXML,objNodeAcesso,"txtTelEnd",strTelLocal)
				Call AddElemento(objXML,objNodeAcesso,"txtTelEndArea",strTelArea)
			End if

			Call AddElemento(objXML,objNodeAcesso,"txtCNLSiglaCentroCliDest",objRS("Acf_SiglaEstEntregaFisico"))
			Call AddElemento(objXML,objNodeAcesso,"txtComplSiglaCentroCliDest",objRS("Acf_ComplSiglaEstEntregaFisico"))
			'Call AddElemento(objXML,objNodeAcesso,"cboInterFaceEndFis",objRS("Acf_InterfaceEstEntregaFisico"))
			strOrigem = Request.Form("hdnOEOrigem")		
			dblAprovisiId = Trim(Request.Form("hdnAprovisiId"))
			dblSolId = ""
			
			If trim(strOrigem) = "Aprov" Then
				if Trim(Request.Form("hdnOrigem")) = "Recadastro" then
					Set objRSSolic = db.execute("CLA_sp_sel_AprovisionadorREC " & dblAprovisiId)					
				else					
					Set objRSSolic = db.execute("CLA_sp_sel_Aprovisionador " & dblAprovisiId)								
				end if
				
				If Not objRSSolic.eof or Not objRSSolic.bof Then
					'Povoamento:
					if Trim(objRSSolic("Acl_IDAcessoLogico")) <> "" then
						strIDLogico 		= Trim(Cstr(objRSSolic("Acl_IDAcessoLogico")))
					end if
					strIDSol 			= Trim(objRSSolic("Sol_ID"))
					strOriSol 			= Trim(objRSSolic("Orisol_ID"))	
					strInterFaceEnd = Trim(objRSSolic("Interface_Cliente"))		'Interface do logradouro(Cliente)
					strInterfaceEbt = Trim(objRSSolic("Interface_Embratel"))	'Interface Embratel					
					Call AddElemento(objXML,objNodeAcesso,"cboInterFaceEnd",strInterFaceEnd)
					if  isNull(strInterfaceEbt) then
					  Call AddElemento(objXML,objNodeAcesso,"cboInterFaceEndFis",strInterFaceEnd)
					else
					  Call AddElemento(objXML,objNodeAcesso,"cboInterFaceEndFis",strInterFaceEnd)
					end if  
				end if							
			end if	
			  
			if Not isNull(objRS("Acf_SiglaEstEntregaFisico")) and not isNull(objRS("Acf_ComplSiglaEstEntregaFisico")) then
				Set objRSEscEntrega = db.execute("CLA_sp_sel_estacao null,'" & objRS("Acf_SiglaEstEntregaFisico") & "','" & objRS("Acf_ComplSiglaEstEntregaFisico") & "'")
				if Not objRSEscEntrega.Eof And Not objRSEscEntrega.Bof then
					strEndEscEntrega = TratarAspasJS(Trim(objRSEscEntrega("Tpl_Sigla"))) & " " & _
					TratarAspasJS(Trim(objRSEscEntrega("Esc_NomeLogr"))) & ", " & _
					TratarAspasJS(Trim(objRSEscEntrega("Esc_NroLogr"))) & " " & _
					TratarAspasJS(Trim(objRSEscEntrega("Esc_Complemento"))) & " " & _
					TratarAspasJS(Trim(objRSEscEntrega("Esc_Bairro"))) & " " & _
					TratarAspasJS(Trim(objRSEscEntrega("Esc_Cod_Cep")))

					Call AddElemento(objXML,objNodeAcesso,"txtEndEstacaoEntrega",strEndEscEntrega)
				End if
				Set objRSEscEntrega = Nothing
			End if

			Call AddElemento(objXML,objNodeAcesso,"txtCNPJ",objRS("Aec_CNPJ"))
			Call AddElemento(objXML,objNodeAcesso,"txtIE",objRS("Aec_IE"))
			Call AddElemento(objXML,objNodeAcesso,"txtIM",objRS("Aec_IM"))
			Call AddElemento(objXML,objNodeAcesso,"txtComplEnd",objRS("Aec_Complemento"))
			Call AddElemento(objXML,objNodeAcesso,"txtPropEnd",objRS("Aec_PropEnd"))
			Call AddElemento(objXML,objNodeAcesso,"txtCodSAP",objRS("Aec_CodSAP"))
			Call AddElemento(objXML,objNodeAcesso,"txtNroPI",objRS("Acf_NroPI"))
			Call AddElemento(objXML,objNodeAcesso,"hdnCompartilhamento",0)
			Call AddElemento(objXML,objNodeAcesso,"hdnNodeCompartilhado",0)
			Call AddElemento(objXML,objNodeAcesso,"hdnCompartilhamento1",0)

			if objRS("Aec_SiglaCentroCliente") <> "" then
				strSiglaCentroCli = objRS("Aec_SiglaCentroCliente")
				Call AddElemento(objXML,objNodeAcesso,"txtCNLSiglaCentroCli",Left(strSiglaCentroCli,4))
				Call AddElemento(objXML,objNodeAcesso,"txtComplSiglaCentroCli",Mid(strSiglaCentroCli,5,len(strSiglaCentroCli)))
			End if
			Call AddElemento(objXML,objNodeAcesso,"cboTipoPonto",objRS("Alf_TipoCCTO"))
			Call AddElemento(objXML,objNodeAcesso,"cboInterFaceEnd",objRS("Acf_Interface"))
			Call AddElemento(objXML,objNodeAcesso,"intOrdem",intIndice)
			Call AddElemento(objXML,objNodeAcesso,"Acl_IdAcessoLogicoPrincipal",dlbIdAcessoLog)
			Call AddElemento(objXML,objNodeAcesso,"TipoAcao","A")
			Call AddElemento(objXML,objNodeAcesso,"Acf_Id",objRS("Acf_Id"))
			Call AddElemento(objXML,objNodeAcesso,"Alf_Id",objRS("Alf_Id"))
			Call AddElemento(objXML,objNodeAcesso,"Aec_Id",objRS("Aec_Id"))
			
			if not objDictionary.Exists(Trim(Ucase(objRS("Acf_Proprietario")))) then
				Call objDictionary.Add(Trim(Ucase(objRS("Acf_Proprietario"))),Trim(Ucase(objRS("Acf_Proprietario"))))
			End if						
			
			if Trim(objRS("Acf_Proprietario")) = "TER" or Trim(objRS("Acf_Proprietario")) = "CLI" then
				strPropAcessoFisico = Trim(objRS("Acf_Proprietario"))
			End if

			if intIndice = 1 then   'Seta variáveis para o primeiro ponto
				'Primeiro ponto
				dblTecId			= Trim(objRS("Tec_Id"))
				strTipoVel			= objRS("Acf_TipoVel")
			End if

			Call AddElemento(objXML,objNodeAcesso,"hdnIdAcessoFisico",objRS("Acf_IdAcessoFisico"))
			Call AddElemento(objXML,objNodeAcesso,"hdnAcfID",objRS("Acf_Id"))

			Set objNodeIdFis = objXML.createNode("element", "IdFisico", "")
			objNodeAcesso.appendChild (objNodeIdFis)
			Call AddElemento(objXML,objNodeIdFis,"Acf_Id",objRS("Acf_Id"))
			Call AddElemento(objXML,objNodeIdFis,"Alf_Id",objRS("Alf_Id"))
			Call AddElemento(objXML,objNodeIdFis,"Aec_Id",objRS("Aec_Id"))
			Call AddElemento(objXML,objNodeIdFis,"Acf_IdAcessoFisico",objRS("Acf_IdAcessoFisico"))
			Call AddElemento(objXML,objNodeIdFis,"Acf_NroAcessoPtaEbt",objRS("Acf_NroAcessoPtaEbt"))

			dblAcfIdXml = objRS("Acf_Id")
			if not objDicAcf.Exists(Cstr(dblAcfIdXml)) then
				Call objDicAcf.Add(Cstr(dblAcfIdXml),Cstr(dblAcfIdXml))

				if dblAcfIdXml = "" or isNull(dblAcfIdXml) then dblAcfIdXml = "null"
				Set objRSPed =	db.execute("CLA_sp_view_pedido null,null,null,null,null,null,null,null,null,'T',"& dblAcfIdXml)
				if not objRSPed.Eof and not objRSPed.Bof then
					While not objRSPed.Eof
						strDM	= Ucase(Trim(objRSPed("Ped_Prefixo"))) & "-" & Trim(right("00000" & objRSPed("Ped_Numero"),5)) & "/" & Trim(objRSPed("Ped_Ano"))
						Set objNodePed = objXML.createNode("element", "Pedido", "")
						objNodeAcesso.appendChild (objNodePed)
						Call AddElemento(objXML,objNodePed,"Ped_Id",objRSPed("Ped_Id"))
						Call AddElemento(objXML,objNodePed,"DM",strDM)
						Call AddElemento(objXML,objNodePed,"Aec_Id",objRS("Aec_Id"))
						Call AddElemento(objXML,objNodePed,"Acf_Id",objRS("Acf_Id"))
						Call AddElemento(objXML,objNodePed,"Alf_Id",objRS("Alf_Id"))
						objRSPed.MoveNext
					Wend
				Else
					Set objNodePed = objXML.createNode("element", "Pedido", "")
					objNodeAcesso.appendChild (objNodePed)
					Call AddElemento(objXML,objNodePed,"Cod_Cla",Request.Form("hdnSolId"))
					Call AddElemento(objXML,objNodePed,"Aec_Id",objRS("Aec_Id"))
					Call AddElemento(objXML,objNodePed,"Acf_Id",objRS("Acf_Id"))
					Call AddElemento(objXML,objNodePed,"Alf_Id",objRS("Alf_Id"))
				End if
			End if			
		intIndice = intIndice + 1
		objRS.MoveNext
	Wend
	set objDictionary = Nothing
	Set objNodeAcesso = Nothing
	Set objElemento = Nothing

	Set MontarXmlAcesso = objXML

End Function
%>