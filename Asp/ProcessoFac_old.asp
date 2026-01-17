<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/AlocacaoFac.asp"-->
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
		
	'Salva no disco ...Temporário
	objXmlDados.save(Server.MapPath("xmlAlocacao.xml"))
    strXml = "<?xml version='1.0' encoding='ISO-8859-1'?>" & objXmlDados.xml

	Vetor_Campos(1)="adInteger,2,adParamInput," & dblRecId
	Vetor_Campos(2)="adInteger,2,adParamInput," & dblUsuId
	Vetor_Campos(3)="adlongvarchar," & len(strXml)& ",adParamInput," & strXml
	strSqlRet = APENDA_PARAMSTR("CLA_sp_AlocacaoFac",3,Vetor_Campos)

	''Response.Write "<script language=javascript>alert('" & strSqlRet & "');</script>"
	'on Error Resume Next
	strXml =  ForXMLAutoQuery(strSqlRet)

	''Response.Write "<script language=javascript>alert('teste ..');</script>"

	'if err.number <> 0 then
	'	Response.Write "<script language=javascript>alert('" & TratarAspasJS(Err.Description) & "')</script>"
	'	On Error Goto 0
	'	Response.End 
	'End if
	objXmlDados.loadXml(strXml)
	
	''Response.Write "<script language=javascript>alert('teste ...');</script>"
	
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
			Response.Write "<script language=javascript>parent.document.forms[0].cboRede.disabled=false;parent.document.forms[0].cboPlataforma.disabled=false</script>"
		Else	
			Response.Write "<script language=javascript>parent.document.forms[0].cboRede.disabled=true;parent.document.forms[0].cboPlataforma.disabled=true</script>"
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

		
		''Response.Write "<script language=javascript>alert('teste');</script>"
		
		strSqlRet = APENDA_PARAMSTR("CLA_sp_check_recurso2",6,Vetor_Campos)
		Set objRSRec = db.execute(strSqlRet)
		Set DBAction = objRSRec("ret")
		
		''Response.Write "<script language=javascript>alert('" & DBAction & "teste');</script>"
		
		dblRecId = ""
		If DBAction = 0 then
			dblRecId = objRSRec("Rec_ID")
		End if
		if Request.Form("hdnTipoProcesso") <> "4" and dblRecId = "" then
			Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
			Response.End 
		End if

		Call AlocaFacilidade(dblRecId)

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

			Response.Write "<script language=javascript>parent.resposta(" & DBAction & ",'');</script>"

		Case "LiberarFacilidade"		

			dblPedId		= Request.Form("hdnPedId") 

			Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
			objXmlDados.loadXml(Request.Form("hdnXmlFacLibera"))
				
			'Salva no disco ...Temporário
			'objXmlDados.save(Server.MapPath("FacilidadeLiberada.xml"))
			'Response.End 
			strXml = "<?xml version='1.0' encoding='ISO-8859-1'?>" & objXmlDados.xml
			
			'Response.Write strXml

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
