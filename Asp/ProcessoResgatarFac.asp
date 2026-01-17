<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/AlocacaoFac.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoResgatarFac.asp
'	- Responsável		: Vital
'	- Descrição			: Retorna facilidades de estoque e e compartilahmento na alocação

Select Case Trim(Request.Form("hdnAcao"))

	Case "ResgatarIdFisicoComp"
	
		Call ResgatarIdFisicoComp(Request.Form("hdnIdAcessoFisico"),"IDFis")
		
	Case "ResgatarEstoque"	

		Call ResgatarIdFisicoComp(Request.Form("hdnIdAcessoFisico1"),"Estoque")
		

End Select

Function ResgatarIdFisicoComp(strIdFisico,strTipo)

	Dim strNDet
	Dim strDet
	Dim strAde

	Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
	objXmlDados.loadXml("<xDados/>")

	Select Case strTipo 
		Case "IDFis"
			Set objRS = db.Execute("CLA_SP_Sel_Facilidade null,null,'" & strIdFisico & "'")
			Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,"C")

		Case "Estoque" 
			Response.Write ("CLA_SP_Sel_Facilidade null,null,'" & strIdFisico & "',null,null,'E'")
			Set objRS = db.Execute("CLA_SP_Sel_Facilidade null,null,'" & strIdFisico & "',null,null,'E'")
			Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,"E")
	End Select		
	
	Set objNode = objXmlDados.selectNodes("//Facilidade")
	if objNode.length > 0 then

		set objNodeAux = objXmlDados.getElementsByTagName("Pro_ID")
		if objNodeAux.length > 0 then dblProId = objNodeAux(0).childNodes(0).text
		set objNodeAux = objXmlDados.getElementsByTagName("Prm_ID")
		if objNodeAux.length > 0 then dblPrmId = objNodeAux(0).childNodes(0).text
		set objNodeAux = objXmlDados.getElementsByTagName("Reg_ID")		
		if objNodeAux.length > 0 then  dblRegId = objNodeAux(0).childNodes(0).text
		set objNodeAux = objXmlDados.getElementsByTagName("Dst_ID")
		if objNodeAux.length > 0 then  dblDstId = objNodeAux(0).childNodes(0).text
		set objNodeAux = objXmlDados.getElementsByTagName("Sis_ID")
		if objNodeAux.length > 0 then  dblSisId = objNodeAux(0).childNodes(0).text
		set objNodeAux = objXmlDados.getElementsByTagName("Esc_ID")
		if objNodeAux.length > 0 then  dblEscId = objNodeAux(0).childNodes(0).text 'Estação de entrega
		set objNodeAux = objXmlDados.getElementsByTagName("Ped_Id")
		if objNodeAux.length > 0 then  dblPedId = objNodeAux(0).childNodes(0).text 'Pedido

		'Resgatar distribuidores para a estação atual
		if dblEscId <> "" then
			set objRS = db.execute("CLA_sp_view_recursodistribuicao " & dblEscId)
			strCboRet = "<Select name=cboDistLocalInstala style=""width:200px"">"
			strCboRet = strCboRet & "<Option value=""""></Option>"
			While Not objRS.Eof
				if Trim(objRS("Dst_ID")) = Trim(dblDstId) then strSel = " selected "
				strCboRet = strCboRet & "<Option value=""" & Trim(objRS("Dst_ID")) & """" & strSel & ">" & objRS("Dst_Desc") & "</Option>"
				strSel = ""
				objRS.MoveNext
			Wend
			strCboRet = strCboRet & "</Select>"
			'Distribuidor da estação de entrega setado
			Response.Write "<script language=javascript>parent.spnDistLocalInstala.innerHTML = '" & TratarAspasJS(strCboRet) & "'</script>"
		Else
			strCboRet = "<Select name=cboDistLocalInstala style=""width:200px"">"
			strCboRet = strCboRet & "<Option value=""""></Option>"
			strCboRet = strCboRet & "</Select>"
			'Distribuidor da estação de entrega setado
			Response.Write "<script language=javascript>parent.spnDistLocalInstala.innerHTML = '" & TratarAspasJS(strCboRet) & "'</script>"
		End if

		if dblProID <> "" then

			Set objRS = db.execute("CLA_sp_sel_promocaoprovedor 0," & dblProID)
			strCboRet = "<Select name=""cboPromocao"" style=""width:170px"">"
			strCboRet = strCboRet & "<Option value=""""></Option>"
			While Not objRS.Eof
				if Trim(objRS("Prm_ID")) = Trim(dblPrmId) then strSel = " selected "
				strCboRet = strCboRet & "<Option value=""" & Trim(objRS("Prm_ID")) & """" & strSel & ">" & objRS("Prm_Desc") & "</Option>"
				strSel = ""
				objRS.MoveNext
			Wend
			strCboRet = strCboRet & "</Select>"
			'Promoção
			Response.Write "<script language=javascript>parent.spnPromocao.innerHTML = '" & TratarAspasJS(strCboRet)  & "'</script>"

			Set objRS = db.execute("CLA_sp_sel_regimecontrato 0," & dblProID)

			strCboRet = "<Select name=""cboRegimeCntr"" style=""width:200px"">"
			strCboRet = strCboRet & "<Option value=""""></Option>"
			While Not objRS.Eof
				if Trim(objRS("Reg_ID")) = Trim(dblRegId) then strSel = " selected "
				strCboRet = strCboRet & "<Option value=""" & Trim(objRS("Reg_ID")) & """" & strSel & ">" & TratarAspasJS(Trim(objRS("Pro_Nome"))) & " - " & TratarAspasJS(Trim(objRS("Tct_Desc"))) & "</Option>"
				strSel = ""
				objRS.MoveNext
			Wend
			strCboRet = strCboRet & "</Select>"

			Response.Write "<script language=javascript>parent.spnRegimeCntr.innerHTML = '" & TratarAspasJS(strCboRet) & "'</script>"
		Else	

			strCboRet = "<Select name=""cboPromocao"" style=""width:170px"">"
			strCboRet = strCboRet & "<Option value=""""></Option>"
			strCboRet = strCboRet & "</Select>"
			'Promoção
			Response.Write "<script language=javascript>parent.spnPromocao.innerHTML = '" & TratarAspasJS(strCboRet)  & "'</script>"

			strCboRet = "<Select name=""cboRegimeCntr"" style=""width:200px"">"
			strCboRet = strCboRet & "<Option value=""""></Option>"
			strCboRet = strCboRet & "</Select>"

			Response.Write "<script language=javascript>parent.spnRegimeCntr.innerHTML = '" & TratarAspasJS(strCboRet) & "'</script>"

		End if

		Response.Write "<script language=javascript>parent.ResgatarInfoRedeEstoque("& dblSisId  &")</script>"
		strRet = "<script language=javascript>with(parent.document.forms[0]){"
		strRet = strRet & "cboProvedor.value = '" & dblProId  & "';"
		strRet = strRet & "cboRede.value = '" & dblSisId & "';"
		strRet = strRet & "hdnRede.value = '" & dblSisId & "';"
		strRet = strRet & "cboLocalInstala.value = '" & dblEscId & "';"
		strRet = strRet & "cboRede.disabled = true;"
		strRet = strRet & "}</script>"
		Response.Write strRet

		'Para as Rede Deterministica e Tronco/Par 
		if dblSisId = 1 or dblSisId= 2 then
			if dblProId <> "" then
				set objRS = db.execute("CLA_sp_sel_Provedor null," & dblProId)
				strCboRet = "<Select name=""cboCodProv"" onChange=""ResgatarPadraoProvedor(this,0)"" >"
				strCboRet = strCboRet & "<Option value=""""></Option>"
				While Not objRS.Eof
					strCboRet = strCboRet & "<Option value=" & Trim(objRS("Pro_ID")) & ">" & TratarAspasJS(objRS("Pro_Cod")) & "</Option>"
					objRS.MoveNext
				Wend
				strCboRet = strCboRet & "</Select>"

				Response.Write "<script language=javascript>parent.spnCodProv.innerHTML = '" & TratarAspasJS(strCboRet)  & "';</script>"

			End if	
		End if

		'Verifica se houve alteracao no cliente,velocidade físico e endereço e complemento
		Vetor_Campos(1)="adInteger,2,adParamInput," & Request.Form("hdnPedId")
		Vetor_Campos(2)="adWChar,15,adParamInput," & strIdFisico
		Vetor_Campos(3)="adInteger,2,adParamOutput,0"

		Call APENDA_PARAM("CLA_sp_check_pendenciamanobra",3,Vetor_Campos)
		ObjCmd.Execute'pega dbaction
		DBAction = ObjCmd.Parameters("RET").value
		Response.Write "<script language=javascript>parent.document.forms[0].hdnAlteracao.value = '" & DBAction & "'</script>"
		Response.Write "<script language=javascript>parent.document.forms[0].hdnPodeAlterar.value = 'N'</script>"

		strXmlFac = FormatarXml(objXmlDados)
		Response.Write "<script language=javascript>parent.objXmlGeral.loadXML('" & strXmlFac &"');parent.AtualizarListaFac();parent.document.forms[0].txtStatus.value = '"& strStatus &"';</script>"

	Else	
		Response.Write "<script language=javascript>alert('Facilidade(s) não encontrada(s)!');parent.limparIDFisico(1);</script>"
		Response.End		
	End If

End Function
%>