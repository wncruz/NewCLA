<!--#include file="../inc/data.asp"-->
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoHistoricoFac.asp
'	- Responsável		: Vital
'	- Descrição			: Consulta de Histórico de Facilidades
Dim strLinkXls
strLinkXls =	"<table border=0 width=100% ><tr><td colspan=2 align=right>" & _
				"<span id=spnXls style=""cursor:hand;color:#003388;"" onclick=""javascript:AbrirXls()"" onmouseover=""showtip(this,event,\'Consulta em formato Excel...\')""><img src=\'../imagens/excel.gif\' border=0></span>&nbsp;" & _
				"<span id=spnIpr style=""cursor:hand;color:#003388;"" onclick=""javascript:TelaImpressao(800,600,\'Consulta de Histórico de Facilidades - " & date() & " " & Time() & " \')"" onmouseover=""showtip(this,event,\'Tela de Impressão...\')""><img src=\'../imagens/impressora.gif\' border=0></span></td></tr>" & _ 
				"</table>"

Function ConsultarHistoricoFac(dblRecId,intRede,intWidthTable,blnLink)
	
	intEstId		= Request.Form("cboLocalInstala")
	intDstId		= Request.Form("cboDistLocalInstala")
	strTronco		= request.Form("txtTronco")
	strPar			= request.Form("txtPar")
	strBastidor		= request.Form("txtBastidor")
	strRegua		= request.Form("txtRegua")
	strPosicao		= request.Form("txtPosicao")
	strTimeSlot		= request.Form("txtTimeslot")
	strDominio		= request.Form("txtDominio")
	strNo			= request.Form("txtNO")
	strSlot			= request.Form("txtSlot")
	strPorta		= request.Form("txtPorta")
	strTipoCabo		= Request.Form("cboTipoCabo")
	strLateral		= Request.Form("txtLateral")
	strCaixaEmenda	= Request.Form("txtCaixaEmenda")
	intStatus		= Request.Form("rdoStatusFac")
	intQtdeReg		= Request.Form("txtQtdeRegistros")

	strXlsDet = ""
	strXmlDet =  "<?xml version=""1.0"" encoding=""ISO-8859-1""?><root>"

	'Não é PADE/PAC
	'Vetor_Campos(18)="adInteger,2,adParamInput," & intQtdeReg
	'Vetor_Campos(19)="adInteger,2,adParamInput," & intStatus

	'strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_sel_facilidade_entrada",19,Vetor_Campos)

	strPedido = Trim(Request.Form("txtPedido"))
	if len(strPedido) <= 3 then strPedido = ""
	
	Vetor_Campos(1)="adInteger,2,adParamInput," 'Ped_Id
	Vetor_Campos(2)="adWChar,15,adParamInput," & strPedido
	Vetor_Campos(3)="adWChar,25,adParamInput," & request("txtNroAcesso")
	Vetor_Campos(4)="adInteger,2,adParamInput," & request("txtSolId")
	Vetor_Campos(5)="adInteger,2,adParamInput," & request("cboLocalInstala")
	Vetor_Campos(6)="adInteger,2,adParamInput," & request("cboDistLocalInstala")
	Vetor_Campos(7)="adInteger,2,adParamInput," & request("cboProvedor")
	Vetor_Campos(8)="adInteger,2,adParamInput," & request("cboSistema")
	Vetor_Campos(9)="adWChar,20,adParamInput," & strTronco  
	Vetor_Campos(10)="adWChar,20,adParamInput," & strPar     
	Vetor_Campos(11)="adWChar,20,adParamInput," & strBastidor
	Vetor_Campos(12)="adWChar,20,adParamInput," & strRegua   
	Vetor_Campos(13)="adWChar,20,adParamInput," & strPosicao 
	Vetor_Campos(14)="adWChar,20,adParamInput," & strTimeslot
	Vetor_Campos(15)="adWChar,20,adParamInput," & strDominio 
	Vetor_Campos(16)="adWChar,20,adParamInput," & strNo      
	Vetor_Campos(17)="adWChar,2,adParamInput," & strSlot    
	Vetor_Campos(18)="adWChar,1,adParamInput," & strPorta
	Vetor_Campos(19)="adWChar,20,adParamInput," & strTipoCabo
	Vetor_Campos(20)="adWChar,20,adParamInput," & strLateral 
	Vetor_Campos(21)="adWChar,20,adParamInput," & strCxEmenda

	strSqlRet = APENDA_PARAMSTRSQL("CLA_sp_sel_historicofacilidade",21,Vetor_Campos)

	Set rs = db.Execute(strSqlRet)

		if not rs.Eof and Not rs.Bof then
			intRede = rs("Sis_Id")
			Select Case intRede

				Case 1	'DET

					strRet = strRet  & "<table border=0 cellspacing=1 cellpadding=0 width=""" & intWidthTable & "px"">"
					strRet = strRet  & "<tr>"
					strRet = strRet  & "<th>&nbsp;+</th>"
					strRet = strRet  & "<th>&nbsp;Fila</th>"
					strRet = strRet  & "<th>&nbsp;Bastidor</th>"
					strRet = strRet  & "<th>&nbsp;Régua</th>"
					strRet = strRet  & "<th>&nbsp;Posição</th>"
					strRet = strRet  & "<th>&nbsp;Domínio</th>"
					strRet = strRet  & "<th>&nbsp;Nó</th>"
					strRet = strRet  & "<th>&nbsp;Slot</th>"
					strRet = strRet  & "<th>&nbsp;Porta</th>"
					strRet = strRet  & "<th>&nbsp;TimeSlot</th>"
					strRet = strRet  & "<th>&nbsp;Estação</th>"
					strRet = strRet  & "<th>&nbsp;Distribuidor</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Sol</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Nº Acso EBT</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Nº Acso CLI</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Nº Acso CCTO Prov</th>"
					strRet = strRet  & "<th>&nbsp;Pedido</th>"
					strRet = strRet  & "<th>&nbsp;Cliente</th>"
					'strRet = strRet  & "<th>&nbsp;Sts</th>"
					strRet = strRet  & "</tr>"

					strXlsDet = strXlsDet  & "<table border=0 cellspacing=1 cellpadding=0 width=""" & intWidthTable & "px"">"
					strXlsDet = strXlsDet  & "<tr height=20>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Fila</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Bastidor</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Régua</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Posição</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Domínio</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Nó</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Slot</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Porta</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;TimeSlot</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Estação</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Distrib</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Bastidor Interno</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Régua Interna</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Posição Interna</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;OTS</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Designação do Tronco</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Link</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Obs</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Sol</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Nº Acso EBT</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Nº Acso CLI</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Nº Acso CCTO Prov</th>"
					strXlsDet = strXlsDet  & "<th nowrap>&nbsp;Pedido</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Cliente</th>"
					strXlsDet = strXlsDet  & "<th>&nbsp;Sts</th>"
					strXlsDet = strXlsDet  & "</tr>"

					strClass = "clsSilver"
					blnDet = false	
					While Not rs.eof
							
						if not isNull(rs("Fac_Bastidor")) then
							if strClass = "clsSilver" then strClass = "clsSilver2"	else strClass = "clsSilver"	end if

							strRet = strRet  & "<tr class="& strClass &">"
							strRet = strRet  & "<td nowrap align=center><span id=spnFac style=""cursor:hand;color:#003388;"" onclick=""javascript:DetalharFacilidade(" & rs("fac_Id") & ")"" onmouseover=""showtip(this,event,\'Detalhes da facilidade...\')"">...</span></td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Fila")) & "</td>"
							'strRet = strRet  & "<td nowrap>&nbsp;</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Bastidor")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Regua")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Posicao")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Dominio")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_No")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Slot")) & "</td>"
							'strRet = strRet  & "<td nowrap>&nbsp;</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Porta")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Timeslot")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Cid_Sigla")) & " " & TratarAspasJS(rs("Esc_Sigla")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Dst_Desc")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Sol_Id")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Acf_NroAcessoPtaEBT")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Acf_NroAcessoPtaCLI")) & "</td>"
							'strRet = strRet  & "<td nowrap>&nbsp;</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Acf_NroAcessoCCTOProvedor")) & "</td>"
							'strRet = strRet  & "<td nowrap>&nbsp;</td>"
							if blnLink then
								if not isNull(rs("Ped_Prefixo")) then
									strRet = strRet  & "<td nowrap>&nbsp;<span id=spnCopy style=""cursor:hand;color:#003388;"" onclick=\'javascript:DetalharSolicitacao(" & RS("Sol_ID") & ")\' >"  & ucase(rs("Ped_Prefixo")) & "-" & right("00000" & rs("Ped_Numero"), 5) & "/" & rs("Ped_Ano") & "</span></td>"
								Else	
									strRet = strRet  & "<td nowrap>&nbsp;</td>"
								End if	
							Else
								if not isNull(rs("Ped_Prefixo")) then
									strRet = strRet  & "<td nowrap>&nbsp;"  & ucase(rs("Ped_Prefixo")) & "-" & right("00000" & rs("Ped_Numero"), 5) & "/" & rs("Ped_Ano") & "</td>"
								Else
									strRet = strRet  & "<td nowrap>&nbsp;</td>"
								End if	
							End if	
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Cli_Nome")) & "</td>"
							'strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Status")) & "</td>"
							strRet = strRet  & "</tr>"

							strXlsDet = strXlsDet  & "<tr class=" & strClass & ">"
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_Fila")) & "</td>"
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_Bastidor")) & "</td>"
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_Regua")) & "</td>"
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_Posicao")) & "</td>"
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_Dominio")) & "</td>"
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_No")) & "</td>"
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_Slot")) & "</td>"
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_Porta")) & "</td>"
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_Timeslot")) & "</td>"
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Cid_Sigla")) & " " & TratarAspasJS(rs("Esc_Sigla")) & "</td>"
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Dst_Desc")) & "</td>"
							strXlsDet = strXlsDet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_BasInter")) & "</td>"
							strXlsDet = strXlsDet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_RegInter")) & "</td>"
							strXlsDet = strXlsDet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_PosInter")) & "</td>"
							strXlsDet = strXlsDet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_OTS")) & "</td>"
							strXlsDet = strXlsDet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_DesigTronco")) & "</td>"
							strXlsDet = strXlsDet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Link")) & "</td>"
							strXlsDet = strXlsDet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Obs")) & "</td>"
							strXlsDet = strXlsDet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Sol_Id")) & "</td>"
							strXlsDet = strXlsDet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Acf_NroAcessoPtaEBT")) & "</td>"
							strXlsDet = strXlsDet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Acf_NroAcessoPtaCLI")) & "</td>"
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Acf_NroAcessoCCTOProvedor")) & "</td>"
							if Not isNull(rs("Ped_Prefixo")) then
								strXlsDet = strXlsDet  & "<td nowrap>&nbsp;"  & ucase(rs("Ped_Prefixo")) & "-" & right("00000" & rs("Ped_Numero"), 5) & "/" & rs("Ped_Ano") & "</td>"
							Else
								strXlsDet = strXlsDet  & "<td nowrap>&nbsp;</td>"
							End if	
							strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Cli_Nome")) & "</td>"
							'strXlsDet = strXlsDet  & "<td >&nbsp;" & TratarAspasJS(rs("Status")) & "</td>"
							strXlsDet = strXlsDet  & "<td nowrap>&nbsp;</td>"
							strXlsDet = strXlsDet  & "</tr>"

							strXmlDet = strXmlDet & "<Facilidade Fac_Id=""" & rs("Fac_Id") & """ Fac_BasInter="""& TratarAspasXML(rs("Fac_BasInter")) &""" Fac_RegInter="""& TratarAspasXML(rs("Fac_RegInter")) &""" Fac_PosInter="""& TratarAspasXML(rs("Fac_PosInter")) &""" Fac_Fila="""& TratarAspasXML(rs("Fac_Fila")) &""" Fac_OTS="""& TratarAspasXML(rs("Fac_OTS")) &""" Fac_DesigTronco="""& TratarAspasXML(rs("Fac_DesigTronco")) &"""  Fac_Link="""& TratarAspasXML(rs("Fac_Link")) &"""  Fac_Obs="""& TratarAspasXML(rs("Fac_Obs")) &"""  Fac_Slot="""& TratarAspasXML(rs("Fac_Slot")) &"""  Fac_Porta="""& TratarAspasXML(rs("Fac_Porta")) &"""/>"
							
							blnDet = true	
						End if	

						rs.movenext

					Wend

					strRet = strRet  & "</table>"
					strXlsDet = strXlsDet  & "</table>"

					if blnDet then
						strXmlDet = strXmlDet & "</root>"
						Response.Write	"<script language=javascript>parent.document.forms[0].hdnXls[0].value = '" & strXlsDet & "';</script>"
						Response.Write	"<script language=javascript>parent.spnPosicoes.innerHTML = '" & strLinkXls & strRet & "'</script>"
						Response.Write "<script language=javascript>parent.objXmlGeral.loadXML('" & strXmlDet & "');</script>"
					Else
						Response.Write "<script language=javascript>alert('Facilidade(s) não encontrada(s).')</script>"
					End if	
					
				Case 2	'NDET

					strRet = strRet  & "<table border=0 cellspacing=1 cellpadding=0 width=""" & intWidthTable & "px"">"
					strRet = strRet  & "<tr>"
					strRet = strRet  & "<th nowrap>&nbsp;Tronco</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Par</th>"
					strRet = strRet  & "<th nowrap>&nbsp;PADE/PAC</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Estação</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Distribuidor</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Obs</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Sol</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Nº Acso EBT</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Nº Acso CLI</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Nº Acso CCTO Prov</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Pedido</th>"
					strRet = strRet  & "<th nowrap>&nbsp;Cliente</th>"
					strRet = strRet  & "</tr>"

					blnNDet = false
					While Not rs.eof

						if not isNull(rs("Fac_Tronco")) then
							if strClass = "clsSilver" then strClass = "clsSilver2"	else strClass = "clsSilver"	end if
							strRet = strRet  & "<tr class="&strClass&">"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Tronco")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Par")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;"
							if rs("Int_ID") <> "" and not isnull(rs("Int_ID")) then
								strRet = strRet  & TratarAspasJS(rs("Int_CorOrigem")) & "&nbsp;>&nbsp;" & TratarAspasJS(rs("Int_CorDestino"))
							End if
							strRet = strRet  & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Cid_Sigla")) & " " & TratarAspasJS(rs("Esc_Sigla")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Dst_Desc")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Obs")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Sol_Id")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Acf_NroAcessoPtaEBT")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Acf_NroAcessoPtaCLI")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Acf_NroAcessoCCTOProvedor")) & "</td>"
							if blnLink then
								if not isNull(rs("Ped_Prefixo")) then
									strRet = strRet  & "<td nowrap>&nbsp;<span id=spnCopy style=""cursor:hand;color:#003388;""  onclick=\'javascript:DetalharSolicitacao(" & RS("Sol_ID") & ")\' >"  & TratarAspasJS(ucase(rs("Ped_Prefixo"))) & "-" & TratarAspasJS(right("00000" & rs("Ped_Numero"), 5)) & "/" & TratarAspasJS(rs("Ped_Ano")) & "</span></td>"
								Else	
									strRet = strRet  & "<td nowrap>&nbsp;</td>"
								End if	
							Else
								if not isNull(rs("Ped_Prefixo")) then
									strRet = strRet  & "<td nowrap>&nbsp;"  & TratarAspasJS(ucase(rs("Ped_Prefixo"))) & "-" & TratarAspasJS(right("00000" & rs("Ped_Numero"), 5)) & "/" & TratarAspasJS(rs("Ped_Ano")) & "</td>"
								Else	
									strRet = strRet  & "<td nowrap>&nbsp;</td>"
								End if	
							End if	
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Cli_Nome")) & "</td>"
							strRet = strRet  & "</tr>"

							blnNDet = true
						End if	

						rs.MoveNext
					Wend	

					strRet = strRet  & "</table>"
					if blnNDet then
						Response.Write	"<script language=javascript>parent.document.forms[0].hdnXls[0].value = '" & strRet & "';</script>"
						Response.Write "<script language=javascript>parent.spnPosicoes.innerHTML = '" & strLinkXls & strRet &"'</script>"
					Else
						Response.Write "<script language=javascript>alert('Facilidade(s) não encontrada(s).')</script>"
					End if	

				Case 3 'ADE

					strRet = strRet  & "<table border=0 cellspacing=1 cellpadding=0 width=""" & intWidthTable & "px"">"
					strRet = strRet  & "<tr>"
					strRet = strRet  & "<th>&nbsp;Cabo</th>"
					strRet = strRet  & "<th>&nbsp;Par</th>"
					strRet = strRet  & "<th>&nbsp;PADE/PAC</th>"
					strRet = strRet  & "<th>&nbsp;Derivação</th>"
					strRet = strRet  & "<th>&nbsp;Tipo Cabo</th>"
					strRet = strRet  & "<th>&nbsp;PADE</th>"
					strRet = strRet  & "<th>&nbsp;Estação</th>"
					strRet = strRet  & "<th>&nbsp;Distribuidor</th>"
					strRet = strRet  & "<th>&nbsp;Obs</th>"
					strRet = strRet  & "<th>&nbsp;Sol</th>"
					strRet = strRet  & "<th>&nbsp;Nº Acesso</th>"
					strRet = strRet  & "<th>&nbsp;Pedido</th>"
					strRet = strRet  & "<th>&nbsp;Cliente</th>"

					strRet = strRet  & "</tr>"

					blnAde = false						
					While Not rs.eof
							
						if not isNull(rs("Fac_Tronco"))then
							if strClass = "clsSilver" then strClass = "clsSilver2"	else strClass = "clsSilver"	end if

							strRet = strRet  & "<tr class="&strClass&">"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Tronco")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Fac_Par")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;"
							if rs("Int_ID") <> "" and not isnull(rs("Int_ID")) then
								strRet = strRet  & TratarAspasJS(rs("Int_CorOrigem")) & "&nbsp;><br>&nbsp;" & TratarAspasJS(rs("Int_CorDestino"))
							End if
							strRet = strRet  & "</td>"
							strRet = strRet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_Lateral")) & "</td>"
							strRet = strRet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_TipoCabo")) & "</td>"
							strRet = strRet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_CxEmenda")) & "</td>"
							strRet = strRet  & "<td >&nbsp;" & TratarAspasJS(rs("Cid_Sigla")) & " " & TratarAspasJS(rs("Esc_Sigla")) & "</td>"
							strRet = strRet  & "<td >&nbsp;" & TratarAspasJS(rs("Dst_Desc")) & "</td>"
							strRet = strRet  & "<td >&nbsp;" & TratarAspasJS(rs("Fac_Obs")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Sol_Id")) & "</td>"
							strRet = strRet  & "<td nowrap>&nbsp;" & TratarAspasJS(rs("Acf_NroAcessoPtaEBT")) & "</td>"
							if blnLink then
								if not isNull(rs("Ped_Prefixo")) then
									strRet = strRet  & "<td nowrap>&nbsp;<span id=spnCopy style=""cursor:hand;color:#003388;""  onclick=\'javascript:DetalharSolicitacao(" & RS("Sol_ID") & ")\' >"  & TratarAspasJS(ucase(rs("Ped_Prefixo"))) & "-" & right("00000" & TratarAspasJS(rs("Ped_Numero")), 5) & "/" & TratarAspasJS(rs("Ped_Ano")) & "</span></td>"
								Else
									strRet = strRet  & "<td>&nbsp;</td>"
								End if	
							Else
								if not isNull(rs("Ped_Prefixo")) then
									strRet = strRet  & "<td nowrap>&nbsp;"  & TratarAspasJS(ucase(rs("Ped_Prefixo"))) & "-" & TratarAspasJS(right("00000" & rs("Ped_Numero"), 5)) & "/" & TratarAspasJS(rs("Ped_Ano")) & "</td>"
								Else
									strRet = strRet  & "<td>&nbsp;</td>"
								End if	
							End if							
							strRet = strRet  & "<td>&nbsp;" & TratarAspasJS(rs("Cli_Nome")) & "</td>"

							strRet = strRet  & "</tr>"

							blnAde = true
						End if	
						rs.movenext

					Wend
						
					strRet = strRet  & "</table>"
					if blnAde then
						Response.Write	"<script language=javascript>parent.document.forms[0].hdnXls[0].value = '" & strRet & "';</script>"
						Response.Write "<script language=javascript>parent.spnPosicoes.innerHTML = '" & strLinkXls & strRet &"'</script>"
					Else	
						Response.Write "<script language=javascript>alert('Facilidade(s) não encontrada(s).')</script>"
					End if	

			End Select
		Else
			Response.Write "<script language=javascript>alert('Facilidade(s) não encontrada(s).');parent.spnPosicoes.innerHTML = '';</script>"
		End if
End Function

Select Case Request.Form("hdnAcao")

	Case "ConsultarHistoricoFac"
		dblRecId = 0
		intRede = Request.Form("cboSistema") 
		Call ConsultarHistoricoFac(dblRecId,intRede,760,true)

End Select
%>
