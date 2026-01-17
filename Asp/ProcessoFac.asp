<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/AlocacaoFac.asp"-->
<!--#include file="../inc/EnviarRetornoSolic_Apg.asp"-->
<!--#include file="../inc/EnviarEntregarAprov.asp"-->
<!--#include file="../inc/EnviarEntregarAprovASMS.asp"-->

<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoFac.asp
'	- Descrição			: Funções para alocação de facilidade

Function AlocaFacilidade(dblRecId)

	Dim objXmlDados
	Dim strStatus

	dblPedId = Trim(Request.Form("hdnPedId"))

	Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
	objXmlDados.loadXml(Request.Form("hdnXML"))

    strXml = "<?xml version='1.0' encoding='ISO-8859-1'?>" & objXmlDados.xml
	
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblRecId
	Vetor_Campos(2)="adInteger,2,adParamInput," & dblUsuId
	Vetor_Campos(3)="adlongvarchar," & len(strXml)& ",adParamInput," & strXml
	strSqlRet = APENDA_PARAMSTR("CLA_sp_AlocacaoFac",3,Vetor_Campos)

	strXml =  ForXMLAutoQuery(strSqlRet)

	objXmlDados.loadXml(strXml)
	
	' ///// implementacao eduardo.araujo  inicio de bloco \\\\\
    ' // bloco necessario para o envio da solicita return para o apg 
	' // tipo de processo cancelamento 
	
	OEorigem =  Request.Form("hdnOEOrigem")
	TipoProcesso = Request.Form("hdnTipoProcesso")
	dblSolid = Request.Form("hdnSolId")
		
	'OEorigem = 4 APG
	'TipoProcesso = 4 cancelamento
	if OEorigem <> "" and TipoProcesso <> "" then
	  if Cint(OEorigem) = 4 and  Cint(TipoProcesso) = 4  then
		
		Vetor_Campos(1)="adWChar,50,adParamInput,null "
		Vetor_Campos(2)="adInteger,1,adParamInput,null "
		Vetor_Campos(3)="adWChar,50,adParamInput,null "
		Vetor_Campos(4)="adInteger,1,adParamInput,null" 
		Vetor_Campos(5)="adInteger,10,adParamInput," & dblSolid
	    
		strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",5,Vetor_Campos)
			
		Set objRSDadosInterf = db.Execute(strSql)		
		If Not objRSDadosInterf.eof and  not objRSDadosInterf.Bof Then
				IdInterfaceAPG = objRSDadosInterf("ID_Tarefa_APG")
				strIdLog = objRSDadosInterf("id_logico")	
		End If
		
		msg_retorno = EnviarRetornoSolic_Apg_CAN_DES(strIdLog,IdInterfaceAPG,dblSolid)
	  End if
	End if
	' //// Fim de Bloco \\\\
	
	Set objNode = objXmlDados.selectNodes("//CLA_RetornoTmp[@Msg_ID=159]") 'Faciilidade alocada com sucesso

	''Response.Write "<script language=javascript>alert('teste ..');</script>"
	if objNode.length > 0 then
		'Repõe o XMl no Cliente para atualizar o Fac_Id	
		objXmlDados.loadXml("<xDados/>")
		Set objRS = db.Execute("CLA_SP_Sel_Facilidade " & dblPedId)
		Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,strAcfObs)
		strXmlFac = FormatarXml(objXmlDados)
		Response.Write "<script language=javascript>parent.objXmlGeral.loadXML('" & strXmlFac &"');parent.AtualizarListaFac();parent.document.forms[0].txtStatus.value = '"& strStatus &"';</script>"
		objXmlDados.loadXml(strXmlFac)
		Set objNode = objXmlDados.selectNodes("//Facilidade")
		if objNode.length = 0 then
			Response.Write "<script language=javascript>parent.document.forms[0].cboRede.disabled=false;//parent.document.forms[0].cboPlataforma.disabled=false</script>"
		Else	
			Response.Write "<script language=javascript>parent.document.forms[0].cboRede.disabled=true;//parent.document.forms[0].cboPlataforma.disabled=true</script>"
		End if
	End if
	Response.Write "<script language=javascript>var objXmlRet = new ActiveXObject(""Microsoft.XMLDOM"");objXmlRet.loadXML('" & strXml &"');parent.JanelaConfirmacaoFac(objXmlRet);</script>"

End Function

Select Case Trim(Request.Form("hdnAcao"))

	Case "GravarFacilidade"
		'Seta provedor Embratel
		dim strPlataforma 
		If Request.Form("hdnRede") = "3" then 
			dblProvedor = 11 'EMBRATEL
		Else
			dblProvedor = Request.Form("cboProvedor")
		End if
		
		IF  Request.Form("cboPlataforma") = "" THEN 
			strPlataforma = Request.Form("hdnPlataforma")
		ELSE
			strPlataforma = Request.Form("cboPlataforma")
		END IF 
		
		Vetor_Campos(1)="adInteger,2,adParamInput," & Request.Form("cboLocalInstala")
		Vetor_Campos(2)="adInteger,2,adParamInput," & Request.Form("cboDistLocalInstala")
		Vetor_Campos(3)="adInteger,2,adParamInput," & dblProvedor
		Vetor_Campos(4)="adInteger,2,adParamInput," & Request.Form("hdnRede")
		Vetor_Campos(5)="adInteger,2,adParamInput," & strPlataforma
		Vetor_Campos(6)="adInteger,2,adParamOutput,0"

		strSqlRet = APENDA_PARAMSTR("CLA_sp_check_recurso2",6,Vetor_Campos)
		Set objRSRec = db.execute(strSqlRet)
		Set DBAction = objRSRec("ret")
		
		dblRecId = ""
		If DBAction = 0 then
			dblRecId = objRSRec("Rec_ID")
		End if
		if Request.Form("hdnTipoProcesso") <> "4" and dblRecId = "" then
			Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
			Response.End 
		End if

		Call AlocaFacilidade(dblRecId)

		Case "AlocarFacConsRedeDetManobra"
		
				set objRS = db.execute("CLA_sp_sel_FacilidadeUnica " & Request.Form("hdnFacId") )
				
				if not objRS.eof then 					
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetBastidor.value = '" & objRS("Fac_Bastidor") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetRegua.value = '" & objRS("Fac_Regua") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetPosicao.value = '" & objRS("Fac_posicao") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetTimeslot.value = '" & objRS("Fac_TimeSlot") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetFila.value = '" & objRS("Fac_Fila") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetEstacao.value = '" & objRS("Esc_Id") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetDistribuidor.value = '" & objRS("Dst_Id") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetPlataforma.value = '" & objRS("Pla_ID") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].hdnFacDetid.value = '" & objRS("Fac_Detid") & "'</script>"
				end if 
					
				set objRS = nothing 
		
			
				dblPedId = Request.Form("hdnPedId") 
				if dblPedId = "" then dblPedId = "null"
				
					Set rs = db.execute("CLA_sp_sel_statusPedido null,null,null,null," & dblPedId)
				if Not rs.Eof and Not rs.bof then
					'if not isNull(rs("Ped_DtConclusao")) then
					'	Response.Write "<script language=javascript>alert('Pedido concluído.');parent.LimparFacilidade();</script>"
					'	Response.End 
					'End if
					'if not isNull(rs("Ped_DtCancelamento")) then
					'	Response.Write "<script language=javascript>alert('Pedido cancelado.');parent.LimparFacilidade();</script>"
					'	Response.End 
					'End if
					if rs("Sis_ID") <> 1 and not isNull(rs("Sis_ID"))  then
						'Veirifica se existem facilidades alocadas
						set objRSFac = db.Execute("CLA_sp_sel_facilidade " & dblPedId)
						if not objRSFac.Eof and Not objRSFac.Bof then
							Response.Write "<script language=javascript>alert('Pedido alocado para outro tipo de rede.');parent.LimparFacilidade();</script>"
							Response.End 
						End if	
					End if
					dblSolId = rs("Sol_Id")
					dblPedId = rs("Ped_Id")
					Response.Write "<script language=javascript>parent.document.forms[0].hdnSolId.value = '" & dblSolId & "';parent.document.forms[0].hdnPedId.value = '" & dblPedId & "';parent.window.close();</script>"
				Else
					Response.Write "<script language=javascript>alert('Pedido não encontrado.');</script>"
					Response.End 
				End if	

	Case "AlocarFacConsRedeDet"
		
				set objRS = db.execute("CLA_sp_sel_FacilidadeUnica " & Request.Form("hdnFacId") )
				
				if not objRS.eof then 					
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetBastidor.value = '" & objRS("Fac_Bastidor") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetRegua.value = '" & objRS("Fac_Regua") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetPosicao.value = '" & objRS("Fac_posicao") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetTimeslot.value = '" & objRS("Fac_TimeSlot") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetFila.value = '" & objRS("Fac_Fila") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetEstacao.value = '" & objRS("Esc_Id") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetDistribuidor.value = '" & objRS("Dst_Id") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].txtRedDetPlataforma.value = '" & objRS("Pla_ID") & "'</script>"
					Response.Write "<script language=javascript>parent.document.forms[0].hdnFacDetid.value = '" & objRS("Fac_Detid") & "'</script>"
				end if 
					
				set objRS = nothing 
		
			
				dblPedId = Request.Form("hdnPedId") 
				if dblPedId = "" then dblPedId = "null"
				
				Set rs = db.execute("CLA_sp_sel_statusPedido null,null,null,null," & dblPedId)
				if Not rs.Eof and Not rs.bof then
					if not isNull(rs("Ped_DtConclusao")) then
						Response.Write "<script language=javascript>alert('Pedido concluído.');parent.LimparFacilidade();</script>"
						Response.End 
					End if
					if not isNull(rs("Ped_DtCancelamento")) then
						Response.Write "<script language=javascript>alert('Pedido cancelado.');parent.LimparFacilidade();</script>"
						Response.End 
					End if
					if rs("Sis_ID") <> 1 and not isNull(rs("Sis_ID"))  then
						'Veirifica se existem facilidades alocadas
						set objRSFac = db.Execute("CLA_sp_sel_facilidade " & dblPedId)
						if not objRSFac.Eof and Not objRSFac.Bof then
							Response.Write "<script language=javascript>alert('Pedido alocado para outro tipo de rede.');parent.LimparFacilidade();</script>"
							Response.End 
						End if	
					End if
					dblSolId = rs("Sol_Id")
					dblPedId = rs("Ped_Id")
					Response.Write "<script language=javascript>parent.document.forms[0].hdnSolId.value = '" & dblSolId & "';parent.document.forms[0].hdnPedId.value = '" & dblPedId & "';parent.window.close();</script>"
				Else
					Response.Write "<script language=javascript>alert('Pedido não encontrado.');</script>"
					Response.End 
				End if	
				
		Case "AlocacaoGLA"

			dblPedId = Request.Form("hdnPedId") 
			dblSolId = Request.Form("hdnSolId") 

			Vetor_Campos(1)="adInteger,4,adParamInput," & dblPedId
			Vetor_Campos(2)="adInteger,4,adParamInput," & dblSolId
			Vetor_Campos(3)="adWChar,30,adParamInput," & strUserName
			Vetor_Campos(4)="adInteger,4,adParamOutput,0"
	
			Call APENDA_PARAM("CLA_sp_AlocacaoCarteira",4,Vetor_Campos)
			ObjCmd.Execute'pega dbaction
			'Response.Write "<script language=javascript>alert('Pedido sendo testado.');window.location.replace('facilidade_main.asp');</script>"
			DBAction = ObjCmd.Parameters("RET").value

			Response.Write "<script language=javascript>parent.resposta(" & DBAction & ",'');parent.ContinuaAlocacao(" & dblSolId & "," & dblPedId &");</script>"

		Case "LiberaServico"

			dblPedId		= Request.Form("hdnPedId") 
			dblIdAcessoLog	= Request.Form("hdnIdLog") 
			dblSolId		= Request.Form("hdnSolId") 

			Vetor_Campos(1)="adInteger,4,adParamInput," & dblPedId
			Vetor_Campos(2)="adDouble,8,adParamInput,"	& dblIdAcessoLog
			Vetor_Campos(3)="adInteger,4,adParamInput," & dblSolId
			Vetor_Campos(4)="adInteger,4,adParamInput," & dblUsuId
			Vetor_Campos(5)="adInteger,4,adParamOutput,0"
	
			Call APENDA_PARAM("CLA_sp_LiberaAcessoFisico",5,Vetor_Campos)
			ObjCmd.Execute'pega dbaction
			DBAction = ObjCmd.Parameters("RET").value
			
			if DBAction = 180 then 
			
			strSQL = "select max(sol_id) as sol_id from cla_solicitacao where acl_idacessologico = " & dblIdAcessoLog
			SET objRSSol =  db.execute(strSQL)
			
			If Not objRSSol.eof and not objRSSol.Bof Then
				dblSolId 		= trim(objRSSol("Sol_id"))
			end if 
			
			strSQL = "select OriSol_id , Acl_IDAcessoLogico , tprc_id from cla_solicitacao where sol_id = " & dblSolId
			SET objRS =  db.execute(strSQL)
			
			If Not objRS.eof and not objRS.Bof Then
				origem 		= trim(objRS("OriSol_id"))
				dblIdLogico = trim(objRS("Acl_IDAcessoLogico"))
				tprc_id = trim(objRS("tprc_id"))
				
				if tprc_id <> 2 then 
					if origem = "6" or origem = "7" or origem = "9" then
					    Set objRSMisto = db.Execute("select top 1 sts.sol_id from cla_solicitacao sol inner join cla_statussolicitacao sts on sol.stssol_id = sts.stssol_id where sts.sts_id = 38 and Acl_IDAcessoLogico = " & dblIdLogico & " and sol.sol_id = (select max(sol_id) from cla_solicitacao where Acl_IDAcessoLogico = " & dblIdLogico & ")" )
						  'Set objRSMisto = db.Execute("select top 1 sts.sol_id from cla_solicitacao sol inner join cla_statussolicitacao sts on sol.stssol_id = sts.stssol_id where sts.sts_id = 38 and Acl_IDAcessoLogico = " & dblIdLogico )
					     If Not objRSMisto.eof and  not objRSMisto.Bof Then
						   dblsol_id = objRSMisto("sol_id")
						 End if
						 
						'Só enviar o construir quando todos estiverem construídos.
						 if ( (  isnull(dblsol_id) or dblsol_id = "" )and (origem = "6" or origem = "7" ) ) then
							EnviarEntregarAprov dblIdLogico
								Response.Write "<script language=javascript>parent.resposta(" & DBAction & ",'');parent.window.close();</script>"
					     end if
						 
						  if (  ( isnull(dblsol_id) or dblsol_id = "" ) and (origem = "9" ) ) then
							EnviarEntregarAprovASMS dblIdLogico
								Response.Write "<script language=javascript>parent.resposta(" & DBAction & ",'');parent.window.close();</script>"
					     end if
					end if
				end if
				
				if origem = "4" then
					InterfaceAPG()
				end if 
			end if				
				
			end if 
			
			Response.Write "<script language=javascript>parent.resposta(" & DBAction & ",'');</script>"

		Case "LiberarFacilidade"		

			dblPedId		= Request.Form("hdnPedId") 

			Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
			objXmlDados.loadXml(Request.Form("hdnXmlFacLibera"))
			
			strXml = "<?xml version='1.0' encoding='ISO-8859-1'?>" & objXmlDados.xml

			Vetor_Campos(1)="adlongvarchar," & len(strXml)& ",adParamInput," & strXml
			Vetor_Campos(2)="adInteger,2,adParamOutput,0"

			Call APENDA_PARAM("CLA_sp_AlocacaoFacAux",2,Vetor_Campos)
			ObjCmd.Execute'pega dbaction
			DBAction = ObjCmd.Parameters("RET").value
			
			if DBAction = 159 then

				Response.Write "<script language=javascript>parent.resposta(" & DBAction & ",'');</script>"

				Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")

				'Repõe o XMl no Cliente para atualizar o Fac_Id	
				objXmlDados.loadXml("<xDados/>")
				Set objRS = db.Execute("CLA_SP_Sel_Facilidade " & dblPedId)
				Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,strAcfObs)
				strXmlFac = FormatarXml(objXmlDados)
				Response.Write "<script language=javascript>parent.objXmlGeral.loadXML('" & strXmlFac &"');parent.AtualizarListaFac();parent.document.forms[0].txtStatus.value = '"& strStatus &"';</script>"
				objXmlDados.loadXml(strXmlFac)
				Set objNode = objXmlDados.selectNodes("//Facilidade")
				if objNode.length = 0 then
					Response.Write "<script language=javascript>parent.document.forms[0].cboRede.disabled=false; parent.document.forms[0].cboPlataforma.disabled=false</script>"
				Else	
					Response.Write "<script language=javascript>parent.document.forms[0].cboRede.disabled=true;parent.document.forms[0].cboPlataforma.disabled=true</script>"
				End if
			Else
				Response.Write "<script language=javascript>parent.resposta(" & DBAction & ",'');</script>"
			End if	

End Select
%>
