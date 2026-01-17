<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/xmlAcessos.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoSolic.asp
'	- Descrição			: Funções que atende a Gravação/Alteração da solicitação

'Retira Chr(10) e Chr(13) do objeto xml para enviar p/ o JS/SQL
Function FormatarXml(objXml)

	Dim strXmlDadosAux
	'Retira a quebra de linha que tem no final XML e passa para a variável que vai para o HTML
	strXmlDadosAux = Replace(objXml.xml,Chr(13),"")
	strXmlDadosAux = Replace(strXmlDadosAux,Chr(10),"")

	FormatarXml = strXmlDadosAux
End Function

Function indexPropAcesso(strProp)
	Select Case strProp
		Case "TER"
			indexPropAcesso = 0
		Case "EBT"
			indexPropAcesso = 1
		Case "CLI"
			indexPropAcesso = 2
	End Select
End Function

'Adiciona um Node a um objeto XML se existir atualiza
Function AdicionarNode(strNomeNode,objXML,varValorNode)

	Dim objNodeFilho
	Dim objNodeList

    If objXML.xml = "" Then
	   objXML.loadXML "<xmlDados></xmlDados>"
	End If

	'Verifica se já existe
	Set objNodeList = objXml.selectNodes("*/" & strNomeNode)

	if objNodeList.Length = 0 then
		'Cria
		Set objNodeFilho = objXML.createNode("element", strNomeNode, "")
		objNodeFilho.text = varValorNode
		objXML.documentElement.appendChild (objNodeFilho)
	Else
		'Atualiza
		objNodeList.Item(0).Text = varValorNode
	End If

	Set AdicionarNode = objXML

End Function

Dim objXmlDadosForm
Set objXmlDadosForm = Server.CreateObject("Microsoft.XMLDOM")

'Gravar solicitação de acesso
Function GravarSolicitacao(intTipoProc,intTipoProcAlt)

	Dim objXmlDadosDB
	Dim objAcesso
	Dim strIdLogico
	Dim dblSolId
	Dim pi

	Set objXmlDadosDB = Server.CreateObject("Microsoft.XMLDOM")

	if intTipoProc = 1 then
		'Recupera dados do xml da página anterior
		If Trim(Request.Form("hdnXml")) <> "" Then
			objXmlDadosForm.preserveWhiteSpace = True
			objXmlDadosForm.loadXML(Request.Form("hdnXml"))
		Else
			Response.Write "<script language=javascript>alert('Informações do Acesso são Obrigatórias');</script>"
			Response.End
		End If
	End if

	Set objXmlDadosForm  = AdicionarNode("intTipoProc",objXmlDadosForm,intTipoProc)
	if Trim(intTipoProcAlt) <> "" then
		Set objXmlDadosForm  = AdicionarNode("intTipoProcAlt",objXmlDadosForm,intTipoProcAlt)
	End if

	dblSolId	= Request.Form("hdnSolId")
	dblIdLogico	= Request.Form("hdnIdAcessoLogico")

	objXmlDadosDB.loadXml("<xDados/>")

	if dblIdLogico <> "" then
		Vetor_Campos(1)="adInteger,4,adParamInput,"
		Vetor_Campos(2)="adInteger,4,adParamInput,"
		Vetor_Campos(3)="adInteger,4,adParamInput," & dblIdLogico
		strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_PONTO",3,Vetor_Campos)

		Set objDicProp = Server.CreateObject("Scripting.Dictionary")

		'Response.Write strSqlRet
		Set objRSFis = db.Execute(strSqlRet)
		if Not objRSFis.EOF and not objRSFis.BOF then
			Set objXmlDadosDB = MontarXmlAcesso(objXmlDadosDB,objRSFis,"")
		End if
	End if

	'Salva no disco ...Temporário
	''objXmlDadosForm.save(Server.MapPath("Acessos.xml"))

    strXml = "<?xml version='1.0' encoding='ISO-8859-1'?>" & objXmlDadosForm.xml
	strXmlDataBase = "<?xml version='1.0' encoding='ISO-8859-1'?>" & objXmlDadosDB.xml

	Vetor_Campos(1)="adlongvarchar," & len(strXml)& ",adParamInput," & strXml
	Vetor_Campos(2)="adlongvarchar," & len(strXmlDataBase)& ",adParamInput," & strXmlDataBase

	strSqlRet = APENDA_PARAMSTR("CLA_sp_ins_solicitacaoAtivacao",2,Vetor_Campos)

	strXml =  ForXMLAutoQuery(strSqlRet)

	objXmlDadosForm.loadXml(strXml)
	Set objNode = objXmlDadosForm.selectNodes("//CLA_RetornoTmp[@Msg_Id=155]")
	if objNode.length > 0 then
		dblSolId = objNode(0).attributes(2).value
	End if

	Response.Write "<script language=javascript>var objXmlRet = new ActiveXObject(""Microsoft.XMLDOM"");objXmlRet.loadXML('" & strXml &"');parent.Message(objXmlRet);</script>"

	if intTipoProc <> 3 then
		if dblSolId = "" then dblSolId = Request.Form("hdnSolId")
		if dblSolId <> "" then	Call EnviarEmailAlteracaoStatus(dblSolId,0,"")
	Else
		if intTipoProcAlt = 1 then
			if dblSolId = "" then dblSolId = Request.Form("hdnSolId")
			if dblSolId <> "" then	Call EnviarEmailAlteracaoStatus(dblSolId,0,"")
		End if
	End if

	Set objXmlDadosForm = Nothing

End Function

'Gravar solicitação de acesso
Function AlterarInfoAcesso(intTipoProc,intTipoProcAlt)

	Dim objXmlDados
	Dim objAcesso
	Dim strIdLogico
	Dim dblSolId
	Dim pi

	Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")

	'Recupera dados do xml da página anterior
	If Trim(Request.Form("hdnXml")) <> "" Then
		objXmlDados.loadXML(Request.Form("hdnXml"))
	Else
		Response.Write "<script language=javascript>alert('Informações do Acesso são Obrigatórias');</script>"
		Response.End
	End If
	Set objXmlDados  = AdicionarNode("intTipoProc",objXmlDados,intTipoProc)
	if Trim(intTipoProcAlt) <> "" then
		Set objXmlDados  = AdicionarNode("intTipoProcAlt",objXmlDados,intTipoProcAlt)
	End if

	'Salva no disco ...Temporário
	'objXmlDados.save(Server.MapPath("xmlAcessos.xml"))

    strXml = "<?xml version='1.0' encoding='ISO-8859-1'?>" & objXmlDados.xml
	Vetor_Campos(1)="adlongvarchar," & len(strXml)& ",adParamInput," & strXml
	strSqlRet = APENDA_PARAMSTR("CLA_sp_upd_solicitacaoAlteracao",1,Vetor_Campos)

	on error resume next
	strXml =  ForXMLAutoQuery(strSqlRet)
	if err.number <> 0 then
		Response.Write "<script language=javascript>alert('" & err.Description & "');</script>"
	end if

	Response.Write "<script language=javascript>var objXmlRet = new ActiveXObject(""Microsoft.XMLDOM"");objXmlRet.loadXML('" & strXml &"');parent.Message(objXmlRet);</script>"

	Set objXmlDados = Nothing

End Function

'Procura o GLA resposável pelo Cliente(Razão Social)
Function ResgatarGLA()

	Dim strHtmlGla

	Vetor_Campos(1)="adWChar,1,adParamInput," & Left(Trim(Request.Form("hdnRazaoSocial")),1) 'Letra
	Vetor_Campos(2)="adInteger,4,adParamInput," & Request.Form("hdnCtfcId") 'Ctfc_Id

	Call APENDA_PARAM("CLA_sp_check_usuario_redirsolicitacao",2,Vetor_Campos)

	Set objRS = ObjCmd.Execute()

	if not objRS.Eof and not objRS.Bof then

		strHtmlGla	= "<table cellspacing=1 cellpadding=0 width=760px border=0><tr class=clsSilver >"
		strHtmlGla	= strHtmlGla & "<td width=170px ><font class=clsObrig>:: </font>UserName GLA</td>"
		strHtmlGla	= strHtmlGla & "<td colspan=5>"
		strHtmlGla	= strHtmlGla & "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=2 bordercolorlight=#003388 bordercolordark=#ffffff width=355 >"
		strHtmlGla	= strHtmlGla & "<tr><td class=lightblue>&nbsp;"
		strHtmlGla	= strHtmlGla & "<span id=spnUserNameGLA>" & Trim(objRS("Usu_UserName")) &  "</span>"
		strHtmlGla	= strHtmlGla & "</td></tr>"
		strHtmlGla	= strHtmlGla & "</table>"
		strHtmlGla	= strHtmlGla & "</td>"
		strHtmlGla	= strHtmlGla & "</tr>"
		strHtmlGla	= strHtmlGla & "<tr class=clsSilver>"
		strHtmlGla	= strHtmlGla & "<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;GLA</td>"
		strHtmlGla	= strHtmlGla & "<td width=355px>"
		strHtmlGla	= strHtmlGla & "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=2 bordercolorlight=#003388 bordercolordark=#ffffff width=100% >"
		strHtmlGla	= strHtmlGla & "<tr><td class=lightblue>&nbsp;"
		strHtmlGla	= strHtmlGla & "<span id=spnNomeGLA>" & Trim(objRS("Usu_Nome")) &  "</span>"
		strHtmlGla	= strHtmlGla & "</td></tr>"
		strHtmlGla	= strHtmlGla & "</table>"
		strHtmlGla	= strHtmlGla & "</td>"
		strHtmlGla	= strHtmlGla & "<td align=right >Ramal&nbsp;</td>"
		strHtmlGla	= strHtmlGla & "<td colspan=3 align=left >"
		strHtmlGla	= strHtmlGla & "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=2 bordercolorlight=#003388 bordercolordark=#ffffff width=100px >"
		strHtmlGla	= strHtmlGla & "<tr><td class=lightblue>&nbsp;"
		strHtmlGla	= strHtmlGla & "<span id=spnRamalGLA>" & Trim(objRS("Usu_Ramal")) & "</span>"
		strHtmlGla	= strHtmlGla & "</td></tr>"
		strHtmlGla	= strHtmlGla & "</table>"
		strHtmlGla	= strHtmlGla & "</td>"
		strHtmlGla	= strHtmlGla & "</tr></table>"

		Response.Write "<script language=javascript>parent.spnGLA.innerHTML = """ & strHtmlGla  & """;" & _
						"parent.document.forms[1].hdntxtGLA.value = """ & objRS("Usu_UserName") & """;</script>"
		Set objRS = Nothing


	Else

		strHtmlGla	= "<table cellspacing=1 cellpadding=0 width=760px ><tr class=clsSilver >"
		strHtmlGla	= strHtmlGla & "<td width=170px ><font class=clsObrig>:: </font>UserName GLA</td>"
		strHtmlGla	= strHtmlGla & "<td colspan=5>"
		strHtmlGla	= strHtmlGla & "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=2 bordercolorlight=#003388 bordercolordark=#ffffff width=20% >"
		strHtmlGla	= strHtmlGla & "<tr><td class=lightblue>&nbsp;"
		strHtmlGla	= strHtmlGla & "<span id=spnUserNameGLA><font color=red>Não Encontrado teste </font></span>"
		strHtmlGla	= strHtmlGla & "</td></tr>"
		strHtmlGla	= strHtmlGla & "</table>"
		strHtmlGla	= strHtmlGla & "</td>"
		strHtmlGla	= strHtmlGla & "</tr>"
		strHtmlGla	= strHtmlGla & "<tr class=clsSilver>"
		strHtmlGla	= strHtmlGla & "<td width=170px>&nbsp;&nbsp;&nbsp;&nbsp;GLA</td>"
		strHtmlGla	= strHtmlGla & "<td width=355px>"
		strHtmlGla	= strHtmlGla & "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=2 bordercolorlight=#003388 bordercolordark=#ffffff width=100% >"
		strHtmlGla	= strHtmlGla & "<tr><td class=lightblue>&nbsp;"
		strHtmlGla	= strHtmlGla & "<span id=spnNomeGLA><font color=red>**********************</font></span>"
		strHtmlGla	= strHtmlGla & "</td></tr>"
		strHtmlGla	= strHtmlGla & "</table>"
		strHtmlGla	= strHtmlGla & "</td>"
		strHtmlGla	= strHtmlGla & "<td align=right >Ramal&nbsp;</td>"
		strHtmlGla	= strHtmlGla & "<td colspan=3 align=left >"
		strHtmlGla	= strHtmlGla & "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=2 bordercolorlight=#003388 bordercolordark=#ffffff width=100px >"
		strHtmlGla	= strHtmlGla & "<tr><td class=lightblue>&nbsp;"
		strHtmlGla	= strHtmlGla & "<span id=spnRamalGLA><font color=red>*******</font></span>"
		strHtmlGla	= strHtmlGla & "</td></tr>"
		strHtmlGla	= strHtmlGla & "</table>"
		strHtmlGla	= strHtmlGla & "</td>"
		strHtmlGla	= strHtmlGla & "</tr></table>"



		Response.Write "<script language=javascript>parent.spnGLA.innerHTML = """ & strHtmlGla  & """;" & _
						"parent.document.forms[1].hdntxtGLA.value = '';</script>"
	End if

End Function

'Resgata a descrição da cidade pela UF,CNL,Perfil do usuário
Function ResgatarCidadeCNL(strUF,strCNL,strUserName,strNomeCnl)

	Dim strCboCid

	'Response.Write ("CLA_sp_sel_cidadesubcefFull '" & strUF & "','" & strUserName & "',null,'" & TratarAspasSQL(strCNL) & "',0")
	Set objRS = db.execute("CLA_sp_sel_cidadesubcefFull '" & strUF & "','" & strUserName & "',null,'" & TratarAspasSQL(strCNL) & "',0")

	if Not objRS.Eof then
		ResgatarCidadeCNL = TratarAspasSQL(Trim(objRS("Cid_Desc")))
	Else
		Response.Write "<script language=javascript>alert('Cidade não pertence ao seu centro funcional.');" & _
						"parent.document.forms[1]." & strNomeCnl & ".value = '';" & _
						"parent.document.forms[1]." & strNomeCnl & ".focus();" & _
						"</script>"
		ResgatarCidadeCNL = ""
	End if

End Function

'Resgata os dados de uma SEV no sistema SSA
Function ResgatarSev(dblNroSev)
	'consulta no banco de dados SSA
	Dim StrConn

	Dim ObjCmdSSA
	Dim ObjParam
	Dim objRSCli
	Dim strCep
	Dim strCidDescRet
	Dim strSolSel
	Dim blnAchou

	Vetor_Campos(1)="adInteger,4,adParamInput," & dblNroSev
	Vetor_Campos(2)="adInteger,4,adParamOutput,0"

	Call APENDA_PARAM("CLA_sp_check_sev",2,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

	if DBAction <> 0 then
		Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
		Response.End
	End if

	'Localiza informações da solução SSA
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblNroSev
	Vetor_Campos(2)="adInteger,2,adParamOutput,0"
	Call APENDA_PARAM("CLA_sp_sel_solucao_ssa",2,Vetor_Campos)
	Set objRSCli = ObjCmd.Execute
	DBAction = ObjCmd.Parameters("RET").value

	if DBAction <> 0 then
		Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
	Else
		If Not objRSCli.eof and  Not objRSCli.bof then

			strCep = TratarAspasJS(Trim(objRSCli("Pre_Cod_Cep")))
			if Trim(strCep) <> "" and len(strCep) = 8  then
				strCep = Mid(strCep,1,5) & "-" & Mid(strCep,6,8)
			End if

			Response.Write "<script language=javascript>with (parent.document.forms[0]){" & _
			"txtRazaoSocial.value = '" & Left(TratarAspasJS(Trim(objRSCli("Cli_des"))),55) & "';" & _
			"txtContaSev.value = '" & TratarAspasJS(Trim(objRSCli("Cli_CC"))) & "';" & _
			"txtNomeFantasia.value = '" & TratarAspasJS(Trim(objRSCli("Cli_NomeFantasia"))) & "';" & _
			"}</script>"

			Response.Write "<script language=javascript>with (parent.document.forms[1]){" & _
			"cboLogrEnd.value = '" & TratarAspasJS(Trim(objRSCli("Tpl_Sigla"))) & "';" & _
			"txtEnd.value = '" & TratarAspasJS(Trim(objRSCli("Pre_NomeLogr"))) & "';" & _
			"txtNroEnd.value = '" & TratarAspasJS(Trim(objRSCli("Pre_NroLogr"))) & "';" & _
			"cboUFEnd.value = '" & TratarAspasJS(Trim(objRSCli("Est_Sigla"))) & "';" & _
			"txtEndCid.value = '" & TratarAspasJS(Trim(objRSCli("Cid_Sigla"))) & "';" & _
			"txtCepEnd.value = '" & Trim(strCep) & "';" & _
			"txtBairroEnd.value = '" & TratarAspasJS(Trim(objRSCli("Pre_Bairro"))) & "';" & _
			"}</script>"

			if isNumeric(objRSCli("Cli_CNPJ")) then Response.Write "<script language=javascript>parent.document.forms[1].txtCNPJ.value = '" & TratarAspasJS(Trim(objRSCli("Cli_CNPJ"))) & "';</script>"
			if isNumeric(objRSCli("Cli_IE")) then	Response.Write "<script language=javascript>parent.document.forms[1].txtIE.value = '" & TratarAspasJS(Trim(objRSCli("Cli_IE"))) & "';</script>"
			if isNumeric(objRSCli("Cli_IM")) then	Response.Write "<script language=javascript>parent.document.forms[1].txtIM.value = '" & TratarAspasJS(Trim(objRSCli("Cli_IM"))) & "';</script>"

			'Seta o combo de cidade
			strCidDescRet = ResgatarCidadeCNL(Trim(objRSCli("Est_Sigla")),Trim(objRSCli("Cid_Sigla")),Trim(Request.Form("hdnUserGICL")),"txtEndCid")
			Response.Write "<script language=javascript>parent.document.forms[1].txtEndCidDesc.value='" & strCidDescRet  & "';try{parent.ResgatarGLA();}catch(e){};</script>"


			strSolSel = "<table border=0 cellspacing=1 cellpadding=0 width=350 ><tr class=clsSilver2><td>Provedor</td><td>Facilidade</td><td>Prazo</td></tr>"
			'Soluções indicadas pelo SSA
			blnAchou = false
			While Not objRSCli.eof
				if Trim(objRSCli("Sol_Selecionada")) = 1 and Trim(objRSCli("For_Des")) <> "" then
 					strSolSel = strSolSel & "<tr class=clsSilver2><td>" & Trim(objRSCli("For_Des")) & "</td><td>" & Trim(objRSCli("Fac_Des")) & "</td><td>" & Trim(objRSCli("Sol_PrazoCompleto")) & "</td></tr>"
 					blnAchou = true
				End if
				objRSCli.MoveNext
			Wend
			strSolSel = strSolSel & "</table>"
			if blnAchou then
				Response.Write "<script language=javascript>parent.strProvedorSelSev.innerHTML = '" & strSolSel & "';</script>"
			Else
				Response.Write "<script language=javascript>parent.strProvedorSelSev.innerHTML = '<table border=0><tr class=clsSilver2><td>Resposta não encontrada</td></tr></table>';</script>"
			End if

		Else
			Response.Write "<script language=javascript>alert('SEV não encontrada.');</script>"
			Response.Write "<script language=javascript>with (parent.document.forms[0]){" & _
			"txtRazaoSocial.value = '';" & _
			"txtContaSev.value = '';" & _
			"txtNomeFantasia.value = '';" & _
			"txtNroSev.value = '';" & _
			"}</script>"

			Response.Write "<script language=javascript>with (parent.document.forms[1]){" & _
			"cboLogrEnd.value = '';" & _
			"txtEnd.value = '';" & _
			"txtNroEnd.value = '';" & _
			"cboUFEnd.value = '';" & _
			"txtEndCid.value = '';" & _
			"txtCepEnd.value = '';" & _
			"txtBairroEnd.value = '';" & _
			"txtCNPJ.value = '';" & _
			"txtIE.value = '';" & _
			"txtIM.value = '';" & _
			"txtEndCid.value = '';" & _
			"}</script>"
			Response.Write "<script language=javascript>parent.strProvedorSelSev.innerHTML = '';</script>"
			Response.Write "<script language=javascript>parent.spnEndCidInstala.innerHTML = '';try{parent.ResgatarGLA();}catch(e){};</script>"

		End if
	End if
End Function

'Procura um CEP e popula os campos da solicitação
Function ProcurarCEP(strCep,intTipo,dblId)

	Dim strCboRet
	Dim intCount
	Dim strNomeCboCid
	Dim strNomespnCid
	Dim strCidSel

	if Cdbl("0" & dblId) <> 0 then
		Set objRS = db.execute("CLA_SP_VIEW_CEP null," & dblId)
	Else
		Set objRS = db.execute("CLA_SP_VIEW_CEP '" & strCep & "'")
	End if

	If Not objRS.eof and  Not objRS.bof then
		intCount = 0
		strCboRet = "<Select name=cboCEPS onChange=""ProcurarCEP(" & intTipo & ",2)"">"
		strCboRet = strCboRet & "<Option value="""">SELECIONE UM CEP</Option>"
		While Not objRS.Eof
			strCboRet = strCboRet & "<Option value=" & Trim(objRS("Cep_ID")) & ">" & TratarAspasJS(Trim(objRS("RuaCompleta"))) & " - " & TratarAspasJS(Trim(objRS("Cep"))) & "</Option>"
			objRS.MoveNext
			intCount = intCount + 1
		Wend
		strCboRet = strCboRet & "</Select>"

		if intCount > 1 then 'Retorna um combo com os CEPS encontrados

			Select Case intTipo
				Case 1
					Response.Write "<script language=javascript>parent.spnCEPSInstala.innerHTML = '" & strCboRet & "'</script>"
				Case 2
					Response.Write "<script language=javascript>parent.spnCEPSInstalaDest.innerHTML = '" & strCboRet & "'</script>"
				Case else
					Response.Write "<script language=javascript>parent.spnCEPS.innerHTML = '" & strCboRet & "'</script>"
			End Select

		Else
			objRS.MoveFirst
			Select Case intTipo
				Case 1
						Response.Write "<script language=javascript>with (parent.document.forms[1]){" & _
						"cboLogrEnd.value ='" & TratarAspasJS(Trim(objRS("Logradouro"))) & "';" & _
						"txtEnd.value ='" & Trim(TratarAspasJS(Trim(objRS("Titulo"))) & " " & Trim(TratarAspasJS(Trim(objRS("Preposicao"))) & " " & TratarAspasJS(Trim(objRS("Rua"))))) & "';" & _
						"cboUFEnd.value ='" & TratarAspasJS(Trim(objRS("Est_Sigla"))) & "';" & _
						"txtBairroEnd.value ='" & TratarAspasJS(Trim(objRS("BairroInicial"))) & "';" & _
						"txtEndCid.value ='" & TratarAspasJS(Trim(objRS("Cid_Sigla"))) & "';" & _
						"txtCepEnd.value ='" & TratarAspasJS(Trim(objRS("Cep"))) & "';" & _
						"}</script>"

						Response.Write "<script language=javascript>parent.spnCEPSInstala.innerHTML = ''</script>"

						'Seta a desc de cidade
						strCidDescRet = ResgatarCidadeCNL(Trim(objRS("Est_Sigla")),Trim(objRS("Cid_Sigla")),Trim(Request.Form("hdnUserGICL")),"txtEndCid")
						Response.Write "<script language=javascript>parent.document.forms[1].txtEndCidDesc.value = '" & strCidDescRet  & "';</script>"

				Case 2
						Response.Write "<script language=javascript>with (parent.document.forms[1]){" & _
						"cboLogrEndDest.value ='" & TratarAspasJS(Trim(objRS("Logradouro"))) & "';" & _
						"txtEndDest.value ='" & Trim(TratarAspasJS(Trim(objRS("Titulo"))) & " " & Trim(TratarAspasJS(Trim(objRS("Preposicao"))) & " " & TratarAspasJS(Trim(objRS("Rua"))))) & "';" & _
						"cboUFEndDest.value ='" & TratarAspasJS(Trim(objRS("Est_Sigla"))) & "';" & _
						"txtBairroEndDest.value ='" & TratarAspasJS(Trim(objRS("BairroInicial"))) & "';" & _
						"txtEndCidDest.value ='" & TratarAspasJS(Trim(objRS("Cid_Sigla"))) & "';" & _
						"txtCepEndDest.value ='" & TratarAspasJS(Trim(objRS("Cep"))) & "';" & _
						"}</script>"

						Response.Write "<script language=javascript>parent.spnCEPSInstalaDest.innerHTML = ''</script>"

						'Seta a desc de cidade
						strCidDescRet = ResgatarCidadeCNL(Trim(objRS("Est_Sigla")),Trim(objRS("Cid_Sigla")),Trim(Request.Form("hdnUserGICL")),"txtEndCidDest")
						Response.Write "<script language=javascript>parent.document.forms[1].txtEndCidDescDest.value = '" & strCidDescRet  & "';</script>"

				Case else

						Response.Write "<script language=javascript>with (parent.document.forms[0]){" & _
						"cboLogr.value ='" & TratarAspasJS(Trim(objRS("Logradouro"))) & "';" & _
						"txtEnd.value ='" & TratarAspasJS(Trim(objRS("Rua"))) & "';" & _
						"cboUF.value ='" & TratarAspasJS(Trim(objRS("Est_Sigla"))) & "';" & _
						"txtBairro.value ='" & TratarAspasJS(Trim(objRS("BairroInicial"))) & "';" & _
						"}</script>"

						Response.Write "<script language=javascript>parent.spnCEPS.innerHTML = ''</script>"

						'Seta o combo de cidade
						strNomeCboCid = "cboCid"
						strNomespnCid = "spnCid"
						strCidSel = Trim(objRS("Cid_Sigla"))

						Response.Write "<script language=javascript>parent." & strNomespnCid & ".innerHTML = '" & ResgatarCidade(Trim(objRSCli("Est_Sigla")),strNomeCboCid,strCidSel,"")  & "';</script>"

			End Select

		End If

	Else
		Response.Write "<script language=javascript>alert('CEP não encontrado.')</script>"
	End if

End Function


'Resgata o usuário a partir do perfil
Function ResgatarUserCoordenacao(strNomeObj,strTipoUser)

	Vetor_Campos(1)="adInteger,4,adParamInput," 'Id do usuário
	Vetor_Campos(2)="adWChar,30,adParamInput," & Request.Form(strNomeObj) 'UserName
	Vetor_Campos(3)="adWChar,10,adParamInput," & strTipoUser 'Tipo do Usuário

	Vetor_Campos(4)="adInteger,4,adParamOutput,0"

	Call APENDA_PARAM("CLA_sp_check_usuario",4,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

	if DBAction <> 0 then
		Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
		Response.Write "<script language=javascript>parent.spnNome"& Right(strNomeObj,len(strNomeObj)-3) & ".innerHTML = """";</script>"
		Response.Write "<script language=javascript>parent.spnRamal"& Right(strNomeObj,len(strNomeObj)-3) & ".innerHTML = """";</script>"
		Response.Write "<script language=javascript>parent.document.forms[2].txt"& Right(strNomeObj,len(strNomeObj)-3) & ".value = """";</script>"
		Response.Write "<script language=javascript>parent.document.forms[2].txt"& Right(strNomeObj,len(strNomeObj)-3) & ".focus();</script>"
	Else
		Set objRS = ObjCmd.Execute()
		Response.Write "<script language=javascript>parent.spnNome"& Right(strNomeObj,len(strNomeObj)-3) & ".innerHTML = """ & objRS("Usu_Nome") & """</script>"
		Response.Write "<script language=javascript>parent.spnRamal"& Right(strNomeObj,len(strNomeObj)-3) & ".innerHTML = """ & objRS("Usu_Ramal") & """</script>"
	End if

End Function

'Procura um cliente no CLA
Function ProcurarCliente()

	Dim strCboRet
	Dim strCliente
	Dim strRetDados
	Dim objRSCompl

	if Trim(Request.Form("txtRazaoSocial")) <> "" and Trim(Request.Form("cboClienteSel")) = "" and Trim(Request.Form("cboEndCliSel")) = "" then

		Vetor_Campos(1)="adInteger,4,adParamInput,"
		Vetor_Campos(2)="adWChar,60,adParamInput," & Trim(Request.Form("txtRazaoSocial")) 'Razao social

		Call APENDA_PARAM("CLA_sp_sel_cliente",2,Vetor_Campos)

				Set objRS = ObjCmd.Execute() 'pega dbaction
		'DBAction = ObjCmd.Parameters("RET").value
		if Not objRS.Eof then

			'Monta combo spnLabelCliente
			if objRS("Tot_Cli") > 3000 then
				Response.Write "<script language=javascript>parent.alert('Detalhe a razão social para localização do cliente.')</script>"
				Response.end
			end if


			strCboRet = "<Select name=cboClienteSel onChange=""ProcurarCliente()"" >"
			strCboRet = strCboRet & "<Option value=""""></Option>"
			While Not objRS.Eof
				'strCliente = objRS("Cli_Nome") & " • " & objRS("End_NomeLogr") & ", " & objRS("End_NroLogr") & " • " & objRS("End_CEP") & " • " & objRS("Cid_Desc") & " • " & objRS("Est_Desc")
				strCliente = objRS("Cli_Nome") & " • Conta: " & objRS("Cli_CC") & " • Sub Conta: " & objRS("Cli_SubCC") '& " • " & objRS("End_CEP") & " • " & objRS("Cid_Desc") & " • " & objRS("Est_Desc")
				strCboRet = strCboRet & "<Option value=""" & Trim(objRS("Cli_ID")) & """>" & TratarAspasJS(strCliente) & "</Option>"
				objRS.MoveNext
			Wend
			strCboRet = strCboRet & "</Select>"

			Response.Write "<script language=javascript>parent.spnLabelCliente.innerHTML = '&nbsp;&nbsp;&nbsp;&nbsp;Cliente(s)'</script>"
			Response.Write "<script language=javascript>parent.spnCliente.innerHTML = '" & strCboRet & "'</script>"
			Response.End

		Else
			Response.Write "<script language=javascript>alert('Cliente não encontrado.');</script>"
			Response.Write "<script language=javascript>parent.spnLabelCliente.innerHTML = ''</script>"
			Response.Write "<script language=javascript>parent.spnCliente.innerHTML = ''</script>"
			Response.End

		End if

	End if

	if Trim(Request.Form("cboClienteSel")) <> "" then  'Monta combo com endereços para o cliente atual


		Vetor_Campos(1)="adInteger,4,adParamInput," & Request.Form("cboClienteSel") 'Cli_id

		Call APENDA_PARAM("CLA_sp_sel_cliente",1,Vetor_Campos)
		Set objRSCliente = ObjCmd.Execute() 'pega dbaction



		Vetor_Campos(1)="adInteger,4,adParamInput,"
		Vetor_Campos(2)="adInteger,4,adParamInput," & Request.Form("cboClienteSel") 'Cli_id
		Call APENDA_PARAM("CLA_sp_sel_endereco",2,Vetor_Campos)

		Set objRS = ObjCmd.Execute() 'pega dbaction

		if Not objRS.Eof then
			'Monta combo spnLabelCliente
			strCboRet = "<Select name=cboEndCliSel onChange=""ProcurarCliente()"" >"
			strCboRet = strCboRet & "<Option value=""""></Option>"


			While Not objRS.Eof
				'Complemento do endereco
				Vetor_Campos(1)="adInteger,4,adParamInput,"	& Trim(objRS("End_ID")) 'End_id
				Vetor_Campos(2)="adInteger,4,adParamInput," & Request.Form("cboClienteSel")

				Call APENDA_PARAM("CLA_sp_sel_endComplemento",2,Vetor_Campos)
				Set objRSCompl = ObjCmd.Execute() 'pega dbaction

				'Montagem do combo
				strCliente = objRS("Tpl_Sigla") & " " & objRS("End_NomeLogr") & ", "& objRS("End_NroLogr") &  " " & objRSCompl("Aec_Complemento") & " • Cep.: " & objRS("End_Cep") & " • Cid Sigla.: " & objRS("Cid_Sigla") & " • Uf: " & objRS("Est_Sigla")
				strCboRet = strCboRet & "<Option value=""" & Trim(Request.Form("cboClienteSel")) & "," & Trim(objRS("End_ID")) & """>" & TratarAspasJS(strCliente) & "</Option>"
				objRS.MoveNext
			Wend
			strCboRet = strCboRet & "</Select>"
			Response.Write strCboRet

			Response.Write "<script language=javascript>parent.spnLabelCliente.innerHTML = '&nbsp;&nbsp;&nbsp;&nbsp;Endereços(s)'</script>"
			Response.Write "<script language=javascript>parent.spnCliente.innerHTML = '" & strCboRet & "'</script>"
			Response.End

		Else
				'Exibe informacoes quando so existe o Cliente e nao existe Endereco
				strRetDados = strRetDados & "<script language=javascript>with (parent.document.forms[0]){"
				strRetDados = strRetDados & "txtRazaoSocial.value = '" & TratarAspasJS(Trim(objRSCliente("Cli_Nome"))) & "';"
				strRetDados = strRetDados & "txtContaSev.value = '" & TratarAspasJS(Trim(objRSCliente("Cli_CC"))) & "';"
				strRetDados = strRetDados & "txtSubContaSev.value = '" & TratarAspasJS(Trim(objRSCliente("Cli_SubCC"))) & "';"
				strRetDados = strRetDados & "txtNomeFantasia.value = '" & TratarAspasJS(Trim(objRSCliente("Cli_NomeFantasia"))) & "';"
				strRetDados = strRetDados & ";txtRazaoSocial.focus();}</script>"

				Response.Write strRetDados

				'Retira o Label e o Combo
				Response.Write "<script language=javascript>parent.spnLabelCliente.innerHTML = ''</script>"
				Response.Write "<script language=javascript>parent.spnCliente.innerHTML = ''</script>"
				Response.End
		End if

	Else
		'Seta o informações do Cliente e do Endereço esse se não tivermos SEV selecionada
		if Trim(Request.Form("cboEndCliSel"))  <> "" then

			Vetor_Campos(1)="adInteger,4,adParamInput," & split(Request.Form("cboEndCliSel"),",")(1) 'End_id
			Vetor_Campos(2)="adInteger,4,adParamInput," & split(Request.Form("cboEndCliSel"),",")(0) 'Cli_id

			Call APENDA_PARAM("CLA_sp_sel_endereco",2,Vetor_Campos)
			Set objRS = ObjCmd.Execute() 'pega dbaction

			Vetor_Campos(1)="adInteger,4,adParamInput,"	& split(Request.Form("cboEndCliSel"),",")(1) 'End_id
			Vetor_Campos(2)="adInteger,4,adParamInput," & split(Request.Form("cboEndCliSel"),",")(0) 'Cli_id

			Call APENDA_PARAM("CLA_sp_sel_endComplemento",2,Vetor_Campos)
			Set objRSCompl = ObjCmd.Execute() 'pega dbaction

			if Not objRS.Eof then

				strRetDados = strRetDados & "<script language=javascript>with (parent.document.forms[0]){"
				strRetDados = strRetDados & "txtRazaoSocial.value = '" & TratarAspasJS(Trim(objRS("Cli_Nome"))) & "';"
				strRetDados = strRetDados & "txtContaSev.value = '" & TratarAspasJS(Trim(objRS("Cli_CC"))) & "';"
				strRetDados = strRetDados & "txtSubContaSev.value = '" & TratarAspasJS(Trim(objRS("Cli_SubCC"))) & "';"
				strRetDados = strRetDados & "txtNomeFantasia.value = '" & TratarAspasJS(Trim(objRS("Cli_NomeFantasia"))) & "';"
				strRetDados = strRetDados & ";txtRazaoSocial.focus();}</script>"

				if Request.Form("txtNroSev") = "" then
					strRetDados = strRetDados & "<script language=javascript>with (parent.document.forms[1]){"
					strRetDados = strRetDados & "txtCNPJ.value = '" & TratarAspasJS(Trim(objRSCompl("Aec_CNPJ"))) & "';"
					strRetDados = strRetDados & "txtIE.value = '" & TratarAspasJS(Trim(objRSCompl("Aec_IE"))) & "';"
					strRetDados = strRetDados & "txtIM.value = '" & TratarAspasJS(Trim(objRSCompl("Aec_IM"))) & "';"
					strRetDados = strRetDados & "}</script>"
				End if

				'Response.Write strRetDados
				if  Trim(Request.Form("txtNroSev")) = "" then
					strRetDados = strRetDados & "<script language=javascript>with (parent.document.forms[1]){"
					strRetDados = strRetDados & "txtComplEnd.value = '" & TratarAspasJS(Trim(objRSCompl("Aec_Complemento"))) & "';"
					strRetDados = strRetDados & "txtContatoEnd.value = '" & TratarAspasJS(Trim(objRSCompl("Aec_Contato"))) & "';"
					strRetDados = strRetDados & "txtTelEnd.value = '" & TratarAspasJS(Trim(objRSCompl("Aec_Telefone"))) & "';"
					strRetDados = strRetDados & "cboLogrEnd.value = '" & TratarAspasJS(Trim(objRS("Tpl_Sigla"))) & "';"
					strRetDados = strRetDados & "txtEnd.value = '" & TratarAspasJS(Trim(objRS("End_NomeLogr"))) & "';"
					strRetDados = strRetDados & "txtNroEnd.value = '" & TratarAspasJS(Trim(objRS("End_NroLogr"))) & "';"
					strRetDados = strRetDados & "cboUFEnd.value = '" & TratarAspasJS(Trim(objRS("Est_Sigla"))) & "';"
					strRetDados = strRetDados & "txtCepEnd.value = '" & TratarAspasJS(Trim(objRS("End_CEP"))) & "';"
					strRetDados = strRetDados & "txtBairroEnd.value = '" & TratarAspasJS(Trim(objRS("End_Bairro"))) & "';"
					strRetDados = strRetDados & "txtEndCid.value = '" & TratarAspasJS(Trim(objRS("Cid_Sigla"))) & "';"
					strRetDados = strRetDados & "}</script>"

					'Seta a desc de cidade
					strCidDescRet = ResgatarCidadeCNL(Trim(objRS("Est_Sigla")),Trim(objRS("Cid_Sigla")),Trim(Request.Form("hdnUserGICL")),"txtEndCid")
					Response.Write "<script language=javascript>parent.document.forms[1].txtEndCidDesc.value='" & strCidDescRet  & "';</script>"
				End if

				Response.Write strRetDados

				'Retira o Label e o Combo
				Response.Write "<script language=javascript>parent.spnLabelCliente.innerHTML = ''</script>"
				Response.Write "<script language=javascript>parent.spnCliente.innerHTML = ''</script>"

			End if

		Else
			Response.Write "<script language=javascript>alert('Cliente não encontrado.');</script>"
			Response.Write "<script language=javascript>parent.spnLabelCliente.innerHTML = ''</script>"
			Response.Write "<script language=javascript>parent.spnCliente.innerHTML = ''</script>"
		End if

	End if

	Set objRSCompl = Nothing

End Function

'Lista RS
Function ListarRS(RS)

	Dim str, fld, cor

	str = str & "<TABLE cellspacing=0 rules=all bordercolorlight=#ffffff bordercolordark=#003399 width=500>" & vbCrLf

	str = str & "<TR>" & vbCrLf
	For Each fld In Rs.Fields
		str = str & "<TH>&nbsp;" & fld.Name & "</TH>" & vbCrLf
	Next
	str = str & "</TR>" & vbCrLf

	cor = "#dddddd"
	Do Until RS.EOF

		if cor = "#dddddd" then
			cor = "#eeeeee"
		else
			cor = "#dddddd"
		end if

		str = str & "<TR>" & vbCrLf
		For Each fld In Rs.Fields
			str = str & "<TD bgcolor=" & cor & ">&nbsp;" & fld.Value & "</TD>" & vbCrLf
		Next
		str = str & "</TR>" & vbCrLf

		RS.MoveNext
	Loop

	str = str & "</TABLE>" & vbCrLf

	str = str & "<table width=415><tr><td align=right>" & RS.RecordCount & " rows</td></tr></table>"

	ListarRS = str

End Function

Dim strCidDescRet
Dim strEndereco
dim objRSEsc

Select Case Trim(Request.Form("hdnAcao"))

	Case "GravarSolicitacao"

		 Call GravarSolicitacao(1,"")

	Case "AlterarInfoAcesso"

		 Call AlterarInfoAcesso(1,"")

	Case "Alteracao"

		Dim objXmlDadosDB
		Dim objXmlDadosRet

		DBAction = ValidarProcesso()
		if DBAction <> 0 then
			Response.Write "<script language=javascript>parent.resposta(" & DBAction & ",'')</script>"
			Response.End
		End if

		Set objXmlDadosForm = Server.CreateObject("Microsoft.XMLDOM")
		Set objXmlDadosDB = Server.CreateObject("Microsoft.XMLDOM")
		Set objXmlDadosRet = Server.CreateObject("Microsoft.XMLDOM")

		'Recupera dados do xml da página anterior
		If Trim(Request.Form("hdnXml")) <> "" Then
			objXmlDadosForm.loadXML(Request.Form("hdnXml"))
		Else
			Response.Write "<script language=javascript>alert('Informações do Acesso são Obrigatórias');</script>"
			Response.End
		End If

		Set objXmlDadosForm  = AdicionarNode("intTipoProc",objXmlDadosForm,3) 'Tipo de Processo Alteração/Ativação/Cancelamento
		'Set objXmlDadosForm  = AdicionarNode("intTipoProcAlt",objXmlDadosForm,2)
		Set objXmlDadosForm  = AdicionarNode("hdnDesigAcessoPriDB",objXmlDadosForm,Request.Form("hdnDesigAcessoPriDB"))

		dblSolId = Request.Form("hdnSolId")
		Set objRSSolic =	db.execute("CLA_sp_view_solicitacaomin " & dblSolId)
		dblIdLogico			= Request.Form("hdnIdAcessoLogico")

		objXmlDadosDB.loadXml("<xDados/>")

		Vetor_Campos(1)="adInteger,4,adParamInput,"
		Vetor_Campos(2)="adInteger,4,adParamInput,"
		Vetor_Campos(3)="adInteger,4,adParamInput," & dblIdLogico
		strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_PONTO",3,Vetor_Campos)

		Set objDicProp = Server.CreateObject("Scripting.Dictionary")

		Set objRSFis = db.Execute(strSqlRet)
		if Not objRSFis.EOF and not objRSFis.BOF then
			Set objXmlDadosDB = MontarXmlAcesso(objXmlDadosDB,objRSFis,"")
		End if


		objXmlDadosDB.save(Server.MapPath("objXmlDadosDB.xml"))
		objXmlDadosForm.save(Server.MapPath("objXmlDadosForm.xml"))

		strXmlFormulario = "<?xml version='1.0' encoding='ISO-8859-1'?>" & objXmlDadosForm.xml
		strXmlDataBase = "<?xml version='1.0' encoding='ISO-8859-1'?>" & objXmlDadosDB.xml

		Vetor_Campos(1)="adlongvarchar," & len(strXmlDataBase)& ",adParamInput," & strXmlDataBase
		Vetor_Campos(2)="adlongvarchar," & len(strXmlFormulario)& ",adParamInput," & strXmlFormulario
		Vetor_Campos(3)="adInteger,2,adParamOutput,0"

		strSqlRet = APENDA_PARAMSTR("CLA_sp_check_tipoalteracao",3,Vetor_Campos)
		strXml =  ForXMLAutoQuery(strSqlRet)

		objXmlDadosRet.loadXml(strXml)
		Set objNode = objXmlDadosRet.selectNodes("//TableControle")
		intTipoProcAlt = 0
		if objNode.length > 0 then
			for intIndex = 0 to objNode.length - 1
				intOrdem = objNode(intIndex).attributes(0).value
				DBAction = objNode(intIndex).attributes(1).value

				if DBAction = 1 then intTipoProcAlt = 1
				if DBAction <> 0 and DBAction <> 1 then
					Response.Write "<script language=javascript>parent.resposta(" & DBAction & ",'')</script>"
					Response.End
				End if

				Set objNodeAcesso = objXmlDadosForm.selectSingleNode("//xDados/Acesso[intOrdem="& intOrdem &"]")
				Set objNodeListAtual = objNodeAcesso.getElementsByTagName("hdnNovoPedido")
				if (objNodeListAtual.length = 0) then 'Cria
					Set objNodeFilho = objXmlDadosForm.createNode("element", "hdnNovoPedido", "")
					objNodeFilho.text = DBAction
					objNodeAcesso.appendChild (objNodeFilho)
				Else 'Atualiza
					objNodeListAtual.item(0).text = DBAction
				End if
			Next
		End if

		Call GravarSolicitacao(3,intTipoProcAlt)

	Case "ResgatarGLA"
		Call ResgatarGLA()

	Case "ResgatarGLA&Gravar"
		Call ResgatarGLA()
		Response.Write "<script language=javascript>parent.ContinuarGravacao();</script>"

	Case "ResgatarCidadeCNL"

		'Response.Write "<script language=javascript>parent.alert('teste')</script>"
		strCidDescRet = ResgatarCidadeCNL(Trim(Request.Form("hdnUFAtual")),Trim(Request.Form("hdnCNLAtual")),Trim(Request.Form("hdnUserGICL")),Trim(Request.Form("hdnCNLNome")))
		Response.Write "<script language=javascript>parent.document.forms[1]." & Request.Form("hdnNomeTxtCidDesc") & ".value = '" & strCidDescRet  & "'</script>"

	Case "ResgatarSev"

		Call ResgatarSev(Trim(Request.Form("txtNroSev")))

	Case "ResgatarUserCoordenacao"

		Call ResgatarUserCoordenacao(Trim(Request.Form("hdnCoordenacaoAtual")),Right(Trim(Request.Form("hdnCoordenacaoAtual")),len(Trim(Request.Form("hdnCoordenacaoAtual")))-3))

	Case "ProcurarCEP"

		Call ProcurarCEP(Trim(Request.Form("hdnCEP")),Trim(Request.Form("hdnTipoCEP")),Trim(Request.Form("cboCEPS")))

	Case "ResgatarEnderecoEstacao"
		'Endereço do local de instalação
		Set objRS = db.execute("CLA_sp_sel_estacao " & Trim(Request.Form("hdnEstacaoAtual")))

		if Not objRS.Eof And Not objRS.Bof then
			strEndereco = Trim(Cstr("" & objRS("Tpl_Sigla"))) & " " &  Trim(Cstr("" & objRS("Esc_NomeLogr"))) & " nº " & Trim(Cstr("" & objRS("Esc_NroLogr"))) & " - " & Trim(Cstr("" & objRS("Est_Sigla"))) & " - " & Trim(Cstr("" & objRS("Cid_Sigla")))
			Response.Write "<script language=javascript>parent.spnContEndLocalInstala.innerHTML = '"	& Replace(Trim(Cstr("" & objRS("Esc_Contato"))),"'","´")	& "'</script>"
			Response.Write "<script language=javascript>parent.spnTelEndLocalInstala.innerHTML = '"		& Replace(Trim(Cstr("" & objRS("Esc_Telefone"))),"'","´")	& "'</script>"
		End if

	Case "ResgatarEstacaoOrigem"

		'Endereço do local de instalação
		Set objRS = db.execute("CLA_sp_sel_estacao null,'" & Trim(Request.Form("txtCNLSiglaCentroCli")) & "','" & Trim(Request.Form("txtComplSiglaCentroCli")) & "'")
		'Response.Write ("CLA_sp_sel_estacao null,'" & Trim(Request.Form("txtCNLSiglaCentroCli")) & "','" & Trim(Request.Form("txtComplSiglaCentroCli")) & "'")
		if Not objRS.Eof And Not objRS.Bof then

			Response.Write "<script language=javascript>with (parent.document.forms[1]){" & _
			"cboLogrEnd.value = '" & TratarAspasJS(Trim(objRS("Tpl_Sigla"))) & "';" & _
			"txtEnd.value = '" & TratarAspasJS(Trim(objRS("Esc_NomeLogr"))) & "';" & _
			"txtNroEnd.value = '" & TratarAspasJS(Trim(objRS("Esc_NroLogr"))) & "';" & _
			"txtComplEnd.value = '" & TratarAspasJS(Trim(objRS("Esc_Complemento"))) & "';" & _
			"cboUFEnd.value = '" & TratarAspasJS(Trim(objRS("Est_Sigla"))) & "';" & _
			"txtEndCid.value = '" & TratarAspasJS(Trim(objRS("Cid_Sigla"))) & "';" & _
			"txtCepEnd.value = '" &	TratarAspasJS(Trim(objRS("Esc_Cod_Cep")))	& "';" & _
			"txtBairroEnd.value = '" & TratarAspasJS(Trim(objRS("Esc_Bairro"))) & "';" & _
			"}</script>"

			'Seta o combo de cidade
			strCidDescRet = ResgatarCidadeCNL(Trim(objRS("Est_Sigla")),Trim(objRS("Cid_Sigla")),Trim(Request.Form("hdnUserGICL")),"txtEndCid")
			Response.Write "<script language=javascript>parent.document.forms[1].txtEndCidDesc.value='" & strCidDescRet  & "';</script>"

		Else
			Response.Write "<script language=javascript>alert('Estação não encontrada.');</script>"
			Response.Write "<script language=javascript>with (parent.document.forms[1]){" & _
			"cboLogrEnd.value = '';" & _
			"txtEnd.value = '';" & _
			"txtNroEnd.value = '';" & _
			"txtComplEnd.value = '';" & _
			"cboUFEnd.value = '';" & _
			"txtEndCid.value = '';" & _
			"txtEndCidDesc.value = '';" & _
			"txtCepEnd.value = '';" & _
			"txtBairroEnd.value = '';" & _
			"}</script>"
		End if

	Case "ResgatarEstacaoDestino"


		'Endereço do local de instalação
		Set objRS = db.execute("CLA_sp_sel_estacao null,'" & Trim(Request.Form("txtCNLSiglaCentroCliDest")) & "','" & Trim(Request.Form("txtComplSiglaCentroCliDest")) & "'")
		'Response.Write ("CLA_sp_sel_estacao null,'" & Trim(Request.Form("txtCNLSiglaCentroCliDest")) & "','" & Trim(Request.Form("txtComplSiglaCentroCliDest")) & "'")
		if Not objRS.Eof And Not objRS.Bof then

			set objRSEsc  =  db.execute("CLA_sp_sel_UsuarioEscCtf " & Request.Form("hdnUsuID") & ",'"  & Trim(Request.Form("txtCNLSiglaCentroCliDest")) & "','"& Trim(Request.Form("txtComplSiglaCentroCliDest")) & "', 1")

			if Not objRSEsc.Eof And Not objRSEsc.Bof then
				Response.Write "<script language=javascript>with (parent.document.forms[1]){" & _
				"txtEndEstacaoEntrega.value = '" & TratarAspasJS(Trim(objRS("Tpl_Sigla"))) & " " & _
				TratarAspasJS(Trim(objRS("Esc_NomeLogr"))) & ", " & _
				TratarAspasJS(Trim(objRS("Esc_NroLogr"))) & " " & _
				TratarAspasJS(Trim(objRS("Esc_Complemento"))) & " " & _
				TratarAspasJS(Trim(objRS("Esc_Bairro"))) & " " & _
				TratarAspasJS(Trim(objRS("Esc_Cod_Cep"))) & "';" & _
				"}</script>"
			else
				Response.Write "<script language=javascript>alert('Usuário não possui acesso a esta estação.');</script>"
				Response.Write "<script language=javascript>with (parent.document.forms[1]){" & _
					"txtEndEstacaoEntrega.value = '' " & _
					"}</script>"
			end if

		Else
			Response.Write "<script language=javascript>alert('Estação não encontrada.');</script>"
			Response.Write "<script language=javascript>with (parent.document.forms[1]){" & _
			"txtEndEstacaoEntrega.value = '' " & _
			"}</script>"
		End if

	Case "ProcurarCliente"
		Call ProcurarCliente()

	Case "ResgatarAcessoFisComp"

		Set objXmlAcessoFisComp = Server.CreateObject("Microsoft.XMLDOM")

		objXmlAcessoFisComp.loadXml("<xDados/>")
		dblAecId = Request.Form("hdnAecIdFis")

		Vetor_Campos(1)="adInteger,4,adParamInput,"	& dblAecId
		Vetor_Campos(2)="adInteger,4,adParamInput,"
		Vetor_Campos(3)="adInteger,4,adParamInput,"
		strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_PONTO",3,Vetor_Campos)

		Set objDicProp = Server.CreateObject("Scripting.Dictionary")

		Set objRSFis = db.Execute(strSqlRet)
		if Not objRSFis.EOF and not objRSFis.BOF then
			Set objXmlAcessoFisComp = MontarXmlAcesso(objXmlAcessoFisComp,objRSFis,"")
		End if

		'Itens que foram populados no compartilhamento do acesso físico e dever estar com seus valores atualizados
		Set objXmlAcessoFisComp  = UpdNodeAcesso("hdnIdAcessoFisico",objXmlAcessoFisComp,Request.Form("hdnIdAcessoFisico"))
		Set objXmlAcessoFisComp  = UpdNodeAcesso("hdnPropIdFisico",objXmlAcessoFisComp,Request.Form("hdnPropIdFisico"))
		Set objXmlAcessoFisComp  = UpdNodeAcesso("hdnCompartilhamento",objXmlAcessoFisComp,Request.Form("hdnCompartilhamento"))
		Set objXmlAcessoFisComp  = UpdNodeAcesso("hdnAecIdFis",objXmlAcessoFisComp,Request.Form("hdnAecIdFis"))

		strXml = FormatarXml(objXmlAcessoFisComp)

		Response.Write "<script language=javascript>parent.objXmlAcessoFisComp.loadXML('" & strXml &"');parent.ResgatarAcessoFisComp(1,parent.objXmlAcessoFisComp);</script>"

End Select

'Adiciona um Node a um objeto XML se existir atualiza
Function UpdNodeAcesso(strNomeNode,objXML,varValorNode)

	Dim objNodeFilho
	Dim objNodeList

    If objXML.xml = "" Then
	   objXML.loadXML "<xmlDados></xmlDados>"
	End If

	'Verifica se já existe
	Set objNodeList = objXml.selectNodes("//Acesso/" & strNomeNode)

	if objNodeList.Length = 0 then
		'Cria
		Set objNodeList = objXml.selectNodes("//Acesso")
		Set objNodeFilho = objXML.createNode("element", strNomeNode, "")
		objNodeFilho.text = varValorNode
		objNodeList(0).appendChild (objNodeFilho)
	Else
		'Atualiza
		objNodeList.Item(0).Text = varValorNode
	End If

	Set UpdNodeAcesso = objXML

End Function

%>
