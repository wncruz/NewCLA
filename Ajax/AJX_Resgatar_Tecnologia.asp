<!--#include file="../inc/data.asp"-->
<%
Response.Expiresabsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

Set objXmlDadosForm = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlRetorno = Server.CreateObject("Microsoft.XMLDOM")
objXmlDadosForm.load(Request)

paramFac = objXmlDadosForm.selectSingleNode("//param").text
paramTec = objXmlDadosForm.selectSingleNode("//param2").text
paramSolid = objXmlDadosForm.selectSingleNode("//param3").text
paramAcfId = objXmlDadosForm.selectSingleNode("//param4").text
paramACL = objXmlDadosForm.selectSingleNode("//param5").text

Function TratarAspasHtml2(strTxt)
	if not isNull(strTxt) then
		TratarAspasHtml = Server.HTMLEncode(strTxt)
	Else
		TratarAspasHtml = ""
	End if
End Function

'response.write "<script>alert('"&paramFac&"')</script>"

'set objRS = db.execute("CLA_sp_sel_TipoONT null," & paramID)

cbo = ""
cbo = cbo & "  <script language=javascript src=../javascript/cla.js></script> "
		
		'Vetor_Campos(1)="adInteger,8,adParamInput," & paramFac
		'Vetor_Campos(2)="adInteger,8,adParamInput,"& paramTec
		'Vetor_Campos(3)="adInteger,8,adParamInput," & paramSolid
	
		'strSqlRet = APENDA_PARAMSTR("CLA_sp_alterTecnologia",3,Vetor_Campos)
		'Set objRS2 = db.Execute(strSqlRet)
		'objRS2.Close
		'objRS2.CursorLocation = adUseClient
		'objRS2.Open

		set objRS = db.execute("CLA_sp_alterTecnologia " & paramFac & " , " & paramTec & " , " &  paramSolid & " , " & paramAcfId ) 

	strFacilidadeServico = "0"
	

	set objRS = db.execute("CLA_sp_sel_newfacilidadeServico " & paramSolid)

	intCount=1
	if not objRS.Eof and not objRS.Bof then
	
	
		if Trim(objRS("orisol_id"))= "10" then 

			strVlan =  "VLAN "
			strSVlan = "SVLAN "
			strPorta = "Porta "
			strPe =    "Eqpto Agregador "


		else

			strVlan =  "VLAN"
			strSVlan = "SVLAN"
			strPorta = "Porta"
			strPe =    "PE"


		end if

	
	
		strFacilidadeServico = "1"

		cbo = cbo & " <table cellspacing=0 cellpadding=0 width=760 border=0> "
		cbo = cbo & "	<tr><th colspan=10> Facilidade do Servico</th></tr> "
		
		cbo = cbo & "	<tr class=clsSilver2> "
		cbo = cbo & "		<td width=170px nowrap><font class=clsObrig>:: </font>" & strPE  & "</td> "
		cbo = cbo & "		<td colspan=9 > "
		cbo = cbo & "			<input type=text class=text name=ser_PE  size=40 maxlength=15 value=" & Trim(objRS("newfacservico_pe"))& " > "
		cbo = cbo & "		</td> "
		cbo = cbo & "	</tr> "
	
		cbo = cbo & "	<tr class=clsSilver> "
		cbo = cbo & "		<td width=170px nowrap><font class=clsObrig>:: </font>" & strPorta & " </td> "
		cbo = cbo & "		<td colspan=9 > "
		cbo = cbo & "			<input type=text class=text name=ser_portaOLt  size=40 maxlength=30 value=" & Trim(objRS("newfacservico_porta"))& "> "
		cbo = cbo & "		</td> "
		cbo = cbo & "	</tr> "

		cbo = cbo & "	<tr class=clsSilver2> "
		cbo = cbo & "		<td width=170px nowrap><font class=clsObrig>:: </font>" & strVlan & " </td> "
		cbo = cbo & "		<td colspan=9> "
		cbo = cbo & "		<input type=text class=text name=ser_Vlan  size=5 maxlength=4  onKeyUp=ValidarTipo(this,0) value=" & objRS("newfacservico_vlan")& " > "
		cbo = cbo & "		</td> "
		cbo = cbo & "	</tr> "
		
		cbo = cbo & "	<tr class=clsSilver> "
		cbo = cbo & "		<td width=170px nowrap> &nbsp;&nbsp; " & strSVlan & " </td> "
		cbo = cbo & "		<td colspan=9 > "
		cbo = cbo & "			<input type=text class=text name=ser_SVLAN  size=5 maxlength=4 onKeyUp=ValidarTipo(this,0) value=" & Trim(objRS("newfacservico_svlan"))& "> "
		cbo = cbo & "		</td> "
		cbo = cbo & "	</tr> "	

		cbo = cbo & " </table> "

		

	End if

		cbo = cbo & " <table border=0 cellspacing=1 cellpadding=0  width=760> "
		cbo = cbo & "	<tr><th colspan=10> Facilidades do Acesso fisico</th></tr> "
		cbo = cbo & "	<tr> "
		cbo = cbo & "		<th> Estrutura</th> "
		cbo = cbo & "		<th colspan=3> Conteudo</th> "
		cbo = cbo & "		<th colspan=3> Solicitacao</th> "
		cbo = cbo & "	</tr> "


			Vetor_Campos(1)="adInteger,2,adParamInput," & paramSolid
			Vetor_Campos(2)="adInteger,2,adParamInput," & paramAcfId
			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_ConsultanewAlocacaoAcesso",2,Vetor_Campos)
			'Response.Write strSqlRet
			Set objRS = db.Execute(strSqlRet)
			objRS.Close
			objRS.CursorLocation = adUseClient
			objRS.Open
			intCount=1
			if not objRS.Eof and not objRS.Bof then  
				'For intIndex = 1 to objRS.PageSize
				While Not objRS.Eof
					if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
					
					cbo = cbo & " <tr class=" & strClass & " > "
					cbo = cbo & "	<td width=170 >&nbsp;TratarAspasHtml(" & objRS("label")& " ) </td> "
						
					cbo = cbo & "	<td colspan=3 >" & trim(objRS("conteudo"))& " </td> "
					cbo = cbo & "	<td colspan=3 >" & trim(objRS("sol_id"))& " </td> "
						
					cbo = cbo & " </tr> "
					
					intCount = intCount+1
					objRS.MoveNext
				Wend
					
			End if


		cbo = cbo & "	</td> "
		cbo = cbo & "	</tr> "
		cbo = cbo & " </table> "

		cbo = cbo & " <table cellspacing=0 cellpadding=0 width=760 border=0> "
		cbo = cbo & "	<tr><th colspan=10> Facilidade do Acesso</th></tr> "
		

Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim strTipo 

'strSql = "CLA_sp_sel_AssocTecnologiaFacilidade"

strFacilidadeAcesso = "false"

'Call PaginarRS(0,strSql)
			Vetor_Campos(1)="adInteger,2,adParamInput," & paramSolid
			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_newAlocacaoAcesso",1,Vetor_Campos)
			'Response.Write strSqlRet
			Set objRS = db.Execute(strSqlRet)
			objRS.Close
			objRS.CursorLocation = adUseClient
			objRS.Open
intCount=1
'response.write "<script>alert('"&objRS.PageSize&"')</script>"
if not objRS.Eof and not objRS.Bof then
	strFacilidadeAcesso = "true"
	'For intIndex = 1 to objRS.PageSize
	While Not objRS.Eof
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

			strobri = ""

			if trim(objRS("formato")) = "TEXTO" 	then strTipo = "7" end if
			if trim(objRS("formato")) = "NUMERICO"  then strTipo = "0" end if 
			if trim(objRS("obrigatorio")) = "S" 	then strobri = "<font class=clsObrig>:: </font>" end if
			'response.write "<script>alert('"&trim(objRS("formato"))&"')</script>"
			'response.write "<script>alert('"&trim(strTipo)&"')</script>"
			'strLabel = ""
			'strLabel = Server.HTMLEncode(objRS("label"))
		
		cbo = cbo & " <tr class=" & strClass & " > "
		cbo = cbo & "	<td width=170px nowrap>  " & strobri 
		cbo = cbo & " " & Server.HTMLEncode(objRS("label"))

		cbo = cbo & " </td> "


		'& "  if trim(objRS("obrigatorio")) = "S" 	then  & " <font class=clsObrig>:: </font>  " &	 end if   TratarAspasHtml2( objRS("label") )  & " </td> "
		'cbo = cbo & "	& <td width="170px" nowrap>	 if trim(objRS("obrigatorio")) = "S" 	then  & " <font class=clsObrig>:: </font>  " &	 end if   TratarAspasHtml2( objRS("label") )  & " </td> "
		'cbo = cbo & "			<font class=clsObrig>:: </font>  "
		'cbo = cbo & "	&	 end if </td> "

		'cbo = cbo & "  strLabel   "  	</td> "		
		'cbo = cbo & " & TratarAspasHtml2( objRS("label") )  & " </td> "
	    
		
		
		 if trim(objRS("formato")) = "COMBO"  then 
			
				
					set objRS2 = db.execute("CLA_sp_sel_EstruturaCombo null,  " & trim(objRS("newcombo_id")) )
					estrutura_combo_id = Trim(objRS2("estrutura_combo_id")) 
				
				

				cbo = cbo & "	<td colspan=9 > "
						
				cbo = cbo & "		<select name=campo_"&intCount&" style=width:250px > "
				cbo = cbo & "				<option value=></option> "
								
									While not objRS2.Eof
										strItemSel = ""
										if Trim(objRS("conteudo")) = Trim(objRS2("label")) then strItemSel = " Selected " End if
										cbo = cbo & " <Option value=" & objRS2("estrutura_combo_id") & strItemSel & ">" & objRS2("label") & " </Option> "
										objRS2.MoveNext
									Wend
									strItemSel = ""
								
						
				cbo = cbo & "		</select> "
			    'Williams (Z518145) - 27/05/2025
				'Habilitado a obrigatoriadade quando a seleção do combo estiver vazia. Estava perdendo na alteração a tecnologia
			    cbo = cbo & "		<input type=Hidden name=" & intCount & " value=" & trim(objRS("obrigatorio")) & " > "
				cbo = cbo & "	</td> "
				
		else 			
			cbo = cbo & "	<td colspan=9 > "
			cbo = cbo & "		<input type=text class=text name=campo_"&intCount&" onKeyUp=ValidarTipo(this,"&trim(strTipo)&") size="& trim(objRS("tamanho"))& " maxlength="&trim(objRS("tamanho"))&" value="& trim(objRS("conteudo"))& " > "
			'Williams (Z518145) - 20/05/2025
			'Ajustado a sintaxe da validação para somente numerico ou texto. Estava perdendo na alteração a tecnologia
			cbo = cbo & "		<input type=Hidden name=" & intCount & " value=" & trim(objRS("obrigatorio")) & " > "
			cbo = cbo & "	</td> "		
	
		end if 
		
			cbo = cbo & " </tr> "
		intCount = intCount+1
		objRS.MoveNext
	Wend
		'objRS.MoveNext
		'if objRS.EOF then Exit For
	'Next
End if

		
cbo = cbo & " </table> "

cbo = cbo & " <input type=Hidden name=facilidadeAcesso value="& strFacilidadeAcesso & "> "
cbo = cbo & " <input type=Hidden name=hdnfacilidadeServico value="& strFacilidadeServico & "> "
'cbo = cbo & " <input type=Hidden name=hdnSolId value=" & paramSolid & "> "

cbo = cbo & " <input type=Hidden name=hdnAcf_ID2 value="& paramAcfId & "> "
cbo = cbo & " <input type=Hidden name=hdnAcl_IDAcessoLogico value="& paramACL & " > " 

%>
<%=cbo%>