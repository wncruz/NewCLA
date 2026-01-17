<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/AlocacaoFac.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Manobra.asp
'	- Descrição			: Alteração de Facilidade

Function TrocarInterLigacao(dblEscId,dblDstId,dblSisId,dblProId,strCoordenadaDE,strCoordenadaPARA,strNumeroacessoDE,dblRecId,dblfacId)

	Vetor_Campos(1)="adWChar,20,adParamInput,"& ucase(strCoordenadaDE)
	Vetor_Campos(2)="adWChar,20,adParamInput,"& ucase(strCoordenadaPARA)
	Vetor_Campos(3)="adInteger,10,adParamInput,"& dblRecId
	Vetor_Campos(4)="adWChar,25,adParamInput,"& strNumeroacessoDE
	Vetor_Campos(5)="adInteger,2,adParamInput,"& dblfacId
	Vetor_Campos(6)="adInteger,3,adParamOutput,0"
	Call APENDA_PARAM("CLA_sp_troca_interligacao",6,Vetor_Campos)
	ObjCmd.Execute
	DBAction = ObjCmd.Parameters("RET").value

	TrocarInterLigacao = DBAction

End Function

Select Case Request.Form("hdnAcao")
		
	Case "AlterarFacilidade"

		DBAction = TrocaFacilidade
		Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'')</script>"

	Case "TrocarInterLigacao"
		
		if request("estacao") <> "" and request("distribuicao") <> "" and request("sistema") <> "" and request("provedor") <> "" then
			if (cint(request("sistema")) = cint(TroncoPar) or cint(request("sistema")) = cint(strRedeAde)) and request("tipo") = "I" then
				dblEscId = request("estacao")
				dblDstId = request("distribuicao")
				dblSisId = request("sistema")
				dblProId = request("provedor")
				dblRecId = request("hdnRecId") 
				For intIndex=1 to Request.Form("hdnCount")
					strNumeroacessoDE	= Request.Form("numeroacessode" & intIndex)
					strCoordenadaDE	= Request.Form("coordenadade" & intIndex)
					strCoordenadaPARA   = Request.Form("coordenadapara" & intIndex)
					dblfacId = Request.Form("fac" & intIndex)
					
					DBAction = TrocarInterLigacao(dblEscId,dblDstId,dblSisId,dblProId,strCoordenadaDE,strCoordenadaPARA,strNumeroacessoDE,dblRecId,dblfacId)
					if DBAction <> 2 then 
						Exit For 
					End if	
				Next
				Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
				if DBAction = 2 then
					Response.Write "<script language=javascript>top.location.replace('PendenciaManobra.asp')</script>"
				End if
			Else
				Response.Write "<script language=javascript>alert('Sistema dever ser ADE ou Não Deterministico!')</script>"
			End if
		Else
			Response.Write "<script language=javascript>alert('Recurso não encontrado!')</script>"
		End if		

	Case "RetirarPendenciaInterligacao" 'Página PedenciaInterligacaoRet.asp
	
		For Each item in Request.Form("chkRetPen") 
			
			Vetor_Campos(1)="adInteger,4,adParamInput,"& item
			Vetor_Campos(2)="adInteger,4,adParamInput,"& dblUsuId

			Vetor_Campos(3)="adInteger,2,adParamOutput,0"
			
			Call APENDA_PARAM("CLA_sp_upd_Interligacaolib",3,Vetor_Campos)
			
			ObjCmd.Execute
			DBAction = ObjCmd.Parameters("RET").value

		Next

		Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');parent.Procurar();</script>"

	Case "ProcurarNroAcesso"

		Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
		objXmlDados.loadXml("<xDados/>")
	
		'Repõe o XMl no Cliente para atualizar o Fac_Id	
		strNroAcesso = Request.Form("txtNroAcesso") 
		strNroPedido = Request.Form("txtNroPedido")

		Vetor_Campos(1)="adInteger,2,adParamInput,"
		Vetor_Campos(2)="AdWChar,25,adParamInput," & strNroAcesso
		Vetor_Campos(3)="AdWChar,15,adParamInput,"
		Vetor_Campos(4)="adInteger,2,adParamInput,"
		Vetor_Campos(5)="AdWChar,3,adParamInput,"
		Vetor_Campos(6)="AdWChar,1,adParamInput,"
		Vetor_Campos(7)="AdWChar,25,adParamInput," & strNroPedido

		strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_Sel_Facilidade",7,Vetor_Campos)
		Set objRS = db.Execute(strSqlRet)
		
		if objRS.Eof then
			Response.Write "<script language=javascript>alert('Registro(s) não encontrado(s).')</script>"
			Response.End 
		End if

		dblSisId = objRS("Sis_Id")
		dblAcfId = objRS("Acf_Id")
		dblProId = objRS("Pro_Id")
		dblRecId = objRS("Rec_Id")
		dblPedId = objRS("Ped_Id")
		dblEscId = objRS("Esc_Id")
		dblDstId = objRS("Dst_Id")
		dblSolId = objRS("Sol_Id")		
		dblPlaId = objRS("Pla_Id")		

		Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,strAcfObs)

		strXmlFac = FormatarXml(objXmlDados)
		Response.Write "<script language=javascript>parent.objXmlGeral.loadXML('" & strXmlFac &"');with(parent.document.forms[0]){cboRede.value='" & dblSisId & "';cboPlataforma.value = '"& dblPlaId &"';Acf_Id.value='"& dblAcfId &"';hdnRecId.value = '"& dblRecId & "';Ped_Id.value = '"& dblPedId & "';cboLocalInstala.value = '"& dblEscId & "';cboDistLocalInstala.value = '"& dblDstId & "';cboProvedor.value = '"& dblProId & "'; hdnSolId.value = '"& dblSolId & "'; }parent.ResgatarInfoRede();parent.AtualizarListaFac();parent.ResgatarProvedoresAssociados("&dblProId&");</script>"

	Case "RealocarFacilidade"
		Call RealocarFacilidade()

	Case "RealocarInterligacao"
		RealocarInterligacao()

End Select

Function RealocarInterligacao()

	Dim objXmlDados
	Dim strStatus

	Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
	objXmlDados.loadXml(Request.Form("hdnXML"))

    strXml = "<?xml version='1.0' encoding='ISO-8859-1'?>" & objXmlDados.xml

	dblRecId = Request.Form("hdnRecId") 
		
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblRecId
	Vetor_Campos(2)="adInteger,2,adParamInput," & dblUsuId
	Vetor_Campos(3)="adlongvarchar," & len(strXml)& ",adParamInput," & strXml
	strSqlRet = APENDA_PARAMSTR("CLA_sp_ManobraInterligacao",3,Vetor_Campos)
	strXml =  ForXMLAutoQuery(strSqlRet)

	objXmlDados.loadXml(strXml)
	
	Response.Write "<script language=javascript>var objXmlRet = new ActiveXObject(""Microsoft.XMLDOM"");objXmlRet.loadXML('" & strXml &"');parent.JanelaConfirmacaoFac(objXmlRet);</script>"

End Function


Function RealocarFacilidade()

	Dim objXmlDados
	Dim strStatus

	Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
	objXmlDados.loadXml(Request.Form("hdnXML"))
		
    strXml = "<?xml version='1.0' encoding='ISO-8859-1'?>" & objXmlDados.xml

	dblRecId = Request.Form("hdnRecId") 
	
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblRecId
	Vetor_Campos(2)="adInteger,2,adParamInput," & dblUsuId
	Vetor_Campos(3)="adlongvarchar," & len(strXml)& ",adParamInput," & strXml
	strSqlRet = APENDA_PARAMSTR("CLA_sp_Manobra",3,Vetor_Campos)
	strXml =  ForXMLAutoQuery(strSqlRet)

	objXmlDados.loadXml(strXml)
	
	Set objNode = objXmlDados.selectNodes("//CLA_RetornoTmp[@Msg_ID=159]") 'Faciilidade alocada com sucesso

	if objNode.length > 0 then
		'Repõe o XMl no Cliente para atualizar o Fac_Id	
		objXmlDados.loadXml("<xDados/>")

		strNroAcesso = Request.Form("txtNroAcesso") 
		strNroPedido = Request.Form("txtNroPedido")

		Vetor_Campos(1)="adInteger,2,adParamInput,"
		Vetor_Campos(2)="AdWChar,25,adParamInput," & strNroAcesso
		Vetor_Campos(3)="AdWChar,15,adParamInput,"
		Vetor_Campos(4)="adInteger,2,adParamInput,"
		Vetor_Campos(5)="AdWChar,3,adParamInput,"
		Vetor_Campos(6)="AdWChar,1,adParamInput,"
		Vetor_Campos(7)="AdWChar,25,adParamInput," & strNroPedido

		strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_Sel_Facilidade",7,Vetor_Campos)
		Set objRS = db.Execute(strSqlRet)

		Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,strAcfObs)
		strXmlFac = FormatarXml(objXmlDados)
		Response.Write "<script language=javascript>parent.objXmlGeral.loadXML('" & strXmlFac &"');parent.AtualizarListaFac();</script>"
	End if

	Response.Write "<script language=javascript>var objXmlRet = new ActiveXObject(""Microsoft.XMLDOM"");objXmlRet.loadXML('" & strXml &"');parent.JanelaConfirmacaoFac(objXmlRet);</script>"

End Function
%>