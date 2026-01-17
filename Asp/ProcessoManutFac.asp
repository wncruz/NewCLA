<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoManutFac.asp
'	- Descrição			: Altera/Remove Facilidades e Interligações
%>
<!--#include file="../inc/data.asp"-->
<%
Dim vRecurso
'NDet /Ade    
Dim vTronco
Dim vPar
'Det          
Dim vBastidor
Dim vRegua		 
Dim vPosicao		 
Dim vTimeSlot	 
Dim vDominio	 
Dim vNo
Dim vSlot
Dim vPorta
Dim vBastidorInter
Dim vReguaInter
Dim vPosicaoInter		 
Dim vFila
Dim vOTS
Dim vDesigTronco
Dim vLink
Dim vObsFac
'ADE
Dim vTipoCabo
Dim vLateral
Dim vCxEmenda
Dim vIdentProvedor

Dim vTronco_Novo	
Dim vPar_Novo		
Dim vBastidor_Novo
Dim vRegua_Novo		
Dim vPosicao_Novo		
Dim vTimeSlot_Novo
Dim vDominio_Novo	
Dim vNo_Novo
Dim vSlot_Novo
Dim vPorta_Novo
Dim vBastidorInter_Novo
Dim vReguaInter_Novo
Dim vPosicaoInter_Novo
Dim vFila_Novo
Dim vOTS_Novo
Dim vDesigTronco_Novo
Dim vLink_Novo
Dim vObsFac_Novo
'ADE
Dim vTipoCabo_Novo
Dim vLateral_Novo
Dim vCxEmenda_Novo
Dim vIdentProvedor_Novo

'Interligacao
Dim vParInter1
Dim vCoodInter1
Dim vEscInter1
Dim vDistInter1

Dim vParInter1_Novo
Dim vCoodInter1_Novo
Dim vEscInter1_Novo
Dim vDistInter1_Novo

Dim vParInter2
Dim vCoodInter2
Dim vEscInter2
Dim vDistInter2

Dim vParInter2_Novo
Dim vCoodInter2_Novo
Dim vEscInter2_Novo
Dim vDistInter2_Novo
Dim vQtdePares

Dim intervalo		
Dim vRecurso1
Dim	vRecurso2

Dim vRecurso1_Novo
Dim vRecurso2_Novo
Dim vFacObs_Novo

Dim vtxtSWITCHa	
Dim vTont_id	
Dim vFont_ID	
Dim vtxtPortaPEa	
Dim vtxtPEa	
				
		
						
Function ResgatarCamposInterligacao(strRede , strDistInstala , strDistConfig)

	Dim strRet
	Dim strItemSelInstala
	Dim strItemSelConfig
	Dim strEscInstala
	Dim strEscConfig
	
	strEscInstala = "	<select name=""cboLocalInstalaIntera"" onChange=""ResgatarDistribuicao(this,5)"">"
	strEscInstala = strEscInstala & "	<option value=""""></option>"
	strEscConfig =  "	<select name=""cboLocalConfiga"" onChange=""ResgatarDistribuicao(this,4)"">"
	strEscConfig =	strEscConfig & "	<option value=""""></option>"
	
	Set objRS = db.execute("CLA_sp_sel_usuarioesc " & dblUsuId)
	While not objRS.Eof 
		strItemSelInstala = ""
		strItemSelConfig = ""
		if Trim(vEscInter1_Novo) = Trim(objRS("Esc_ID")) then strItemSelInstala = " Selected " End if
		if Trim(vEscInter2_Novo) = Trim(objRS("Esc_ID")) then strItemSelConfig = " Selected " End if
		strEscInstala = strEscInstala & "<Option value=""" & objRS("Esc_ID") & """" & strItemSelInstala & ">" & TratarAspasJS(objRS("Cid_Sigla")) & "  " & TratarAspasJS(objRS("Esc_Sigla")) & "</Option>"
		strEscConfig = strEscConfig & "<Option value=""" & objRS("Esc_ID") & """" & strItemSelConfig & ">" & TratarAspasJS(objRS("Cid_Sigla")) & "  " & TratarAspasJS(objRS("Esc_Sigla")) & "</Option>"
		objRS.MoveNext
	Wend
	
	'Não deterministico
	strRet = ""
	strRet = strRet & "<table border=0 cellspacing=1 cellpadding=0 width=760>"
	strRet = strRet & "	<tr>"
	strRet = strRet & "<th  colspan=4>&nbsp;•&nbsp;PARA</th>"
	strRet = strRet & "	</tr>"

	strRet = strRet & "	<tr class=clsSilver>"
	strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Par Inicial</th>"
	strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Coordenada</th>"
	strRet = strRet & "		<th></th>"
	strRet = strRet & "		<th></th>"
	strRet = strRet & "	</tr>"
	strRet = strRet & "<tr class=clsSilver>"
	strRet = strRet & "	<td>1&nbsp;•&nbsp;<input type=text class=text name=txtParInicial1a  value=""" & TratarAspasJS(vParInter1_Novo) & """ maxlength=4 onKeyUp=""ValidarTipo(this,0)"" size=4 onblur=""CompletarCampo(this)"" TIPO=N >(N4)</td>"
	strRet = strRet & "	<td><input type=text class=text name=txtCoordenada1a  value=""" & TratarAspasJS(vCoodInter1_Novo) & """  maxlength=20 size=10>(A20)</td>"

	strRet = strRet & "	<td nowrap><font class=clsObrig>::</font>&nbsp;Local de Instalação&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	
	strRet = strRet & strEscInstala
	
	strRet = strRet & "</select>"
	strRet = strRet & "</td>"
	strRet = strRet & "<td nowrap><font class=clsObrig>::</font>&nbsp;Distribuidor"
	strRet = strRet & "	<span id=""spnDistLocalInstalaIntera"">"
	
	strRet = strRet &  replace(strDistInstala,"cboDistLocalInstalaInter","cboDistLocalInstalaIntera") '"<Option value=""" & objRS("Dst_ID") & """" & strItemSel & ">" & TratarAspasJS(objRS("Dst_Desc")) & "</Option>"
	
	strRet = strRet & "	</select>"
	strRet = strRet & "	</td>"

	strRet = strRet & "</tr>"
	strRet = strRet & "<tr class=clsSilver>"
	strRet = strRet & "	<td>2&nbsp;•&nbsp;<input type=text class=text name=txtParInicial2a  value=""" & TratarAspasJS(vParInter2_Novo) & """  onKeyUp=""ValidarTipo(this,0)"" maxlength=4 size=4 onblur=""CompletarCampo(this)"" TIPO=N >(N4)</td>"
	strRet = strRet & "	<td><input type=text class=text name=txtCoordenada2a  value=""" & TratarAspasJS(vCoodInter2_Novo) & """  maxlength=20 size=10>(A20)</td>"
	strRet = strRet & "	<td nowrap><font class=clsObrig>::</font>&nbsp;Local de Configuração"
	
	strRet = strRet & strEscConfig  

	strRet = strRet & "</select>"
	strRet = strRet & "</td>"
	strRet = strRet & "<td nowrap><font class=clsObrig>::</font>&nbsp;Distribuidor"
	strRet = strRet & "	<span id=""spnDistLocalConfiga"">"
	
	strRet = strRet &  replace(strDistConfig,"cboDistLocalConfig","cboDistLocalConfiga") '"<Option value=""" & objRS("Dst_ID") & """" & strItemSel & ">" & TratarAspasJS(objRS("Dst_Desc")) & "</Option>"
	
	strRet = strRet & "	</select>"
	strRet = strRet & "	</td>"
	strRet = strRet & "</tr>"
				
	strRet = strRet & "<tr>"
	strRet = strRet & "<th colspan=4>&nbsp;Obs</th>"
	strRet = strRet & "	</tr>"

	strRet = strRet & "<tr class=clsSilver>"
	strRet = strRet & "	<td colspan=4><textarea name=txtObsFaca onkeydown=""MaxLength(this,300);""  cols=40 rows=2>" & vFacObs_Novo & "</textarea>(A300)</td>"
	strRet = strRet & "</tr>"

	strRet = strRet & "</table>" 

	ResgatarCamposInterligacao = strRet

End Function

Function ResgatarCamposFacilidade(strRede)

	Dim strRet
	Dim strLocalInstalaInter
	Dim strDistrib
	Dim strItemSel
	
	strRet = ""
	Select Case int("0" & strRede)

		Case 1

			strRet = strRet & "<table border=0 cellspacing=1 cellpadding=0 width=760>"
			strRet = strRet & "<tr>"
			strRet = strRet & "<th  colspan=8>&nbsp;•&nbsp;DE</th>"
			strRet = strRet & "</tr>"
			strRet = strRet & "<tr>"
			strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Fila</th>"
			strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Bastidor</th>"
			strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Regua</th>"
			strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Posicao</th>"
			strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Domínio</th>"
			strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Nó</th>"
			strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Timeslot</th>"
			strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Slot</th>"

			strRet = strRet & "	</tr>"
			strRet = strRet & "<tr class=clsSilver>"
			strRet = strRet & "	<td><input type=text class=text name=txtFila		maxlength=5   size=6 onKeyUp=""ValidarTipo(this,2)"">(A5)</td>"
			strRet = strRet & "	<td><input type=text class=text name=txtBastidor	maxlength=3  size=4 onKeyUp=""ValidarTipo(this,0)"">(N3)</td>"
			strRet = strRet & "	<td><input type=text class=text name=txtRegua		maxlength=2  size=3 onKeyUp=""ValidarTipo(this,2)"">(A2)</td>"
			strRet = strRet & "	<td><input type=text class=text name=txtPosicao		maxlength=3  size=4 onKeyUp=""ValidarTipo(this,2)"">(A3)</td>"
			strRet = strRet & "	<td><input type=text class=text name=txtDominio		maxlength=3  size=4 onKeyUp=""ValidarTipo(this,0)"">(N3)</td>"
			strRet = strRet & "	<td><input type=text class=text name=txtNO			maxlength=3  size=4 onKeyUp=""ValidarTipo(this,0)"">(N3)</td>"
			strRet = strRet & "	<td><input type=text class=text name=txtTimeSlot	maxlength=9 onKeyUp=""ValidarNTipo(this,0,4,4,1,0,4)"" size=10 value=""0001-0031"">(N4-N4)</td>"
			strRet = strRet & "	<td><input type=text class=text name=txtSlot		maxlength=2   size=2 onBlur=""SearchDom(this,\'A1,A2,A3,A4,A5,A6,A7,A8,B1,B2,B3,B4,B5,B6,B7,B8\')"">(A2)</td>"
			strRet = strRet & "</tr>"

			strRet = strRet & "<tr>"
			strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Porta</th>"
			strRet = strRet & "<th>&nbsp;Bastidor Interno</th>"
			strRet = strRet & "<th>&nbsp;Régua Interna</th>"
			strRet = strRet & "<th>&nbsp;Porta Interna</th>"
			strRet = strRet & "<th colspan=2>&nbsp;OTS</th>"
			strRet = strRet & "<th colspan=2>&nbsp;Link</th>"
			
			strRet = strRet & "	</tr>"

			strRet = strRet & "<tr class=clsSilver>"
			strRet = strRet & "	<td><input type=text class=text name=txtPorta		maxlength=1   size=1 onBlur=""SearchDom(this,\'A,B\')"">(A1)</td>"
			strRet = strRet & "	<td><input type=text class=text name=txtBastidorInter	maxlength=20  size=10>(A20)</td>"
			strRet = strRet & "	<td><input type=text class=text name=txtReguaInter		maxlength=20  size=10>(A20)</td>"
			strRet = strRet & "	<td><input type=text class=text name=txtPosicaoInter		maxlength=20  size=10>(A20)</td>"
			strRet = strRet & "	<td colspan=2><input type=text class=text name=txtOTS	maxlength=10  size=10>(A10)</td>"
			strRet = strRet & "	<td colspan=2><input type=text class=text name=txtLink			 maxlength=25 size=25>(A25)</td>"
			strRet = strRet & "</tr>"

			strRet = strRet & "<tr>"
			strRet = strRet & "<th colspan=2>&nbsp;Designação do Tronco</th>"
			strRet = strRet & "<th colspan=2>&nbsp;Identificação Provedor</th>"
			strRet = strRet & "<th colspan=4>&nbsp;Obs</th>"
			strRet = strRet & "	</tr>"

			strRet = strRet & "<tr class=clsSilver>"
			strRet = strRet & "	<td colspan=2><input type=text class=text name=txtDesigTronco		maxlength=25 size=25>(A25)</td>"
			strRet = strRet & "	<td colspan=2><input type=text class=text name=txtIdentProvedor	maxlength=20  size=20>(A20)</td>"
			strRet = strRet & "	<td colspan=4><textarea name=txtObsFac onkeydown=""MaxLength(this,300);""  cols=40 rows=2></textarea>(A300)</td>"
			strRet = strRet & "</tr>"


			strRet = strRet & "</table>"

		Case 2
		
			'Não deterministico
			strRet = ""
			strRet = strRet & "<table border=0 cellspacing=1 cellpadding=0 width=760>"
			strRet = strRet & "	<tr class=clsSilver>"
			strRet = strRet & "		<td><font class=clsObrig>::</font>&nbsp;<input type=radio value=E name=rdoInter "
			if Trim(Request.Form("rdoInter")) = "E" or Trim(Request.Form("rdoInter")) = "" then
				strRet = strRet & "checked  "
			End if
			strRet = strRet & " onClick=""ResgatarCamposFac(this)"">&nbsp;Entrada</td>"
			strRet = strRet & "		<td colspan=3><input type=radio value=I name=rdoInter "
			if Trim(Request.Form("rdoInter")) = "I" then
				strRet = strRet & "checked  "
			End if
			strRet = strRet & " onClick=""ResgatarCamposFac(this)"">&nbsp;PADE/PAC</td>"
			strRet = strRet & "	</tr>"
			strRet = strRet & "	<tr>"
			strRet = strRet & "<th  colspan=4>&nbsp;•&nbsp;DE</th>"
			strRet = strRet & "	</tr>"

			if Trim(Request.Form("rdoInter")) = "E" or Trim(Request.Form("rdoInter")) = "" then
				strRet = strRet & "	<tr class=clsSilver>"
				strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Tronco</th>"
				strRet = strRet & "		<th colspan=3><font class=clsObrig>::</font>&nbsp;Par</th>"
				strRet = strRet & "	</tr>"
				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "	<td><input type=text class=text name=txtTronco  maxlength=20 size=10>(A20)</td>"
				strRet = strRet & "	<td colspan=3><input type=text class=text name=txtPar  maxlength=9 size=10 onKeyUp=""if(this.value.lenght > 1) ValidarNTipo(this,0,4,4,1,0,4)"" TIPO=N >(N4-N4) ou (A1)</td>"
				strRet = strRet & "</tr>"
			Else
				strRet = strRet & "	<tr class=clsSilver>"
				strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Par Inicial</th>"
				strRet = strRet & "		<th colspan=3><font class=clsObrig>::</font>&nbsp;PADE/PAC</th>"
				strRet = strRet & "	</tr>"
				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "	<td>1&nbsp;•&nbsp;<input type=text class=text name=txtParInicial1  maxlength=4 onKeyUp=""ValidarTipo(this,0)"" size=4 onblur=""CompletarCampo(this)"" TIPO=N >(N4)</td>"
				strRet = strRet & "	<td><input type=text class=text name=txtCoordenada1  maxlength=20 size=10>(A20)</td>"

				strRet = strRet & "	<td nowrap><font class=clsObrig>::</font>&nbsp;Local de Instalação&nbsp;&nbsp;&nbsp;&nbsp;"
				strRet = strRet & "	<select name=""cboLocalInstalaInter"" onChange=""ResgatarDistribuicao(this,3)"">"
				strRet = strRet & "	<option value=""""></option>"
				set objRS = db.execute("CLA_sp_sel_usuarioesc " & dblUsuId)
				strLocalInstalaInter = Request.Form("cboLocalInstala")
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strLocalInstalaInter) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
					strRet = strRet & "<Option value=""" & objRS("Esc_ID") & """" & strItemSel & ">" & TratarAspasJS(objRS("Cid_Sigla")) & "  " & TratarAspasJS(objRS("Esc_Sigla")) & "</Option>"
					objRS.MoveNext
				Wend
				strRet = strRet & "</select>"
				strRet = strRet & "</td>"
				strRet = strRet & "<td nowrap><font class=clsObrig>::</font>&nbsp;Distribuidor"
				strRet = strRet & "	<span id=""spnDistLocalInstalaInter"">"
				strRet = strRet & "	<select name=""cboDistLocalInstalaInter"" >"
				strRet = strRet & "	<option value=""""></option>"
				strDistrib = Request.Form("cboDistLocalInstala") 
				if Trim(strLocalInstalaInter) <> "" then 
					set objRS = db.execute("CLA_sp_view_recursodistribuicao " & strLocalInstalaInter)
					While not objRS.Eof 
						strItemSel = ""
						if Trim(strDistrib) = Trim(objRS("Dst_ID")) then strItemSel = " Selected " End if
						strRet = strRet &  "<Option value=""" & objRS("Dst_ID") & """" & strItemSel & ">" & TratarAspasJS(objRS("Dst_Desc")) & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				End if	
				strRet = strRet & "	</select>"
				strRet = strRet & "	</td>"

				strRet = strRet & "</tr>"
				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "	<td>2&nbsp;•&nbsp;<input type=text class=text name=txtParInicial2 onKeyUp=""ValidarTipo(this,0)"" maxlength=4 size=4 onblur=""CompletarCampo(this)"" TIPO=N >(N4)</td>"
				strRet = strRet & "	<td><input type=text class=text name=txtCoordenada2  maxlength=20 size=10>(A20)</td>"
				strRet = strRet & "	<td nowrap><font class=clsObrig>::</font>&nbsp;Local de Configuração"
				strRet = strRet & "	<select name=""cboLocalConfig"" onChange=""ResgatarDistribuicao(this,2)"">"
				strRet = strRet & "	<option value=""""></option>"
				set objRS = db.execute("CLA_sp_sel_estacao  null")
				While not objRS.Eof 
					strItemSel = ""
					'if Trim(strLocalConfig) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
					strRet = strRet & "<Option value=""" & objRS("Esc_ID") & """" & strItemSel & ">" & TratarAspasJS(objRS("Cid_Sigla")) & "  " & TratarAspasJS(objRS("Esc_Sigla")) & "</Option>"
					objRS.MoveNext
				Wend
				strRet = strRet & "</select>"
				strRet = strRet & "</td>"
				strRet = strRet & "<td nowrap><font class=clsObrig>::</font>&nbsp;Distribuidor"
				strRet = strRet & "	<span id=""spnDistLocalConfig"">"
				strRet = strRet & "	<select name=""cboDistLocalConfig"" >"
				strRet = strRet & "	<option value=""""></option>"
				strRet = strRet & "	</select>"
				strRet = strRet & "	</td>"
				strRet = strRet & "</tr>"
	
				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "		<td><font class=clsObrig>::</font>&nbsp;Quantidade de Pares</td>"
				strRet = strRet & "		<td><input type=text name=txtQtdePares size=5 onKeyUp=""ValidarTipo(this,0)"" class=text></td>"
				strRet = strRet & "		<td></td>"
				strRet = strRet & "		<td></td>"
				strRet = strRet & "</tr>"

				strRet = strRet & "<tr>"
				strRet = strRet & "<th colspan=4>&nbsp;Obs</th>"
				strRet = strRet & "	</tr>"

				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "	<td colspan=4><textarea name=txtObsFac onkeydown=""MaxLength(this,300);""  cols=40 rows=2></textarea>(A300)</td>"
				strRet = strRet & "</tr>"


			End if
				
			strRet = strRet & "</table>" 

		Case 3

			strRet = ""
			strRet = strRet & "<table border=0 cellspacing=1 cellpadding=0 width=760>"

			strRet = strRet & "	<tr class=clsSilver>"
			strRet = strRet & "		<td><font class=clsObrig>::</font>&nbsp;<input type=radio value=E name=rdoInter "
			if Trim(Request.Form("rdoInter")) = "E" or Trim(Request.Form("rdoInter")) = "" then
				strRet = strRet & "checked  "
			End if
			strRet = strRet & " onClick=""ResgatarCamposFac(this)"">&nbsp;Entrada</td>"
			strRet = strRet & "		<td colspan=4><input type=radio value=I name=rdoInter "
			if Trim(Request.Form("rdoInter")) = "I" then
				strRet = strRet & "checked  "
			End if
			strRet = strRet & " onClick=""ResgatarCamposFac(this)"">&nbsp;PADE/PAC</td>"
			strRet = strRet & "	</tr>"

			strRet = strRet & "	<tr>"
			strRet = strRet & "<th  colspan=5>&nbsp;•&nbsp;DE</th>"
			strRet = strRet & "	</tr>"

			if Trim(Request.Form("rdoInter")) = "E" or Trim(Request.Form("rdoInter")) = "" then
				strRet = strRet & "	<tr class=clsSilver >"
				strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Cabo</th>"
				strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Par</th>"
				strRet = strRet & "		<th nowrap><font class=clsObrig>::</font>&nbsp;Tipo do Cabo</th>"
				strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Derivação</th>"
				strRet = strRet & "		<th nowrap><font class=clsObrig>::</font>&nbsp;PADE</th>"
				strRet = strRet & "	</tr>"
				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "<td><input type=text class=text name=txtTronco	maxlength=2 size=2 onKeyUp=""parent.ValidarTipo(this,0)"" onblur=""CompletarCampo(this)"" TIPO=N >(N2)</td>"
				strRet = strRet & "<td><input type=text class=text name=txtPar 	maxlength=9 size=9 onKeyUp=""parent.ValidarNTipo(this,0,4,4,1,0,4)"">(N4-N4)</td>"
				strRet = strRet & "<td><select name=cboTipoCabo >"
				strRet = strRet & "			<option value=""""></option>"
				strRet = strRet & "			<option value=H >H</option>"
				strRet = strRet & "			<option value=X >X</option>"
				'strRet = strRet & "			<option value=XDSL >XDSL</option>"
				strRet = strRet & "			<option value=C >C</option>"
				strRet = strRet & "		</select></td>"
				strRet = strRet & "<td><input type=text class=text name=txtLateral  onKeyUp=""parent.ValidarTipo(this,2)""	maxlength=4 size=5>(A4)</td>"
				strRet = strRet & "<td><input type=text class=text name=txtCaixaEmenda onKeyUp=""parent.ValidarTipo(this,2)""  maxlength=3 size=5 >(A3)</td>"
				strRet = strRet & "</tr>"
			Else

				strRet = strRet & "	<tr class=clsSilver>"
				strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Par Inicial</th>"
				strRet = strRet & "		<th colspan=3><font class=clsObrig>::</font>&nbsp;PADE/PAC</th>"
				strRet = strRet & "	</tr>"
				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "	<td>1&nbsp;•&nbsp;<input type=text class=text name=txtParInicial1  maxlength=4 onKeyUp=""ValidarTipo(this,0)"" size=4 onblur=""CompletarCampo(this)"" TIPO=N >(N4)</td>"
				strRet = strRet & "	<td><input type=text class=text name=txtCoordenada1  maxlength=20 size=10>(A20)</td>"

				strRet = strRet & "	<td nowrap><font class=clsObrig>::</font>&nbsp;Local de Instalação&nbsp;&nbsp;&nbsp;&nbsp;"
				strRet = strRet & "	<select name=""cboLocalInstalaInter"" onChange=""ResgatarDistribuicao(this,3)"">"
				strRet = strRet & "	<option value=""""></option>"
				set objRS = db.execute("CLA_sp_sel_usuarioesc " & dblUsuId)
				strLocalInstalaInter = Request.Form("cboLocalInstala")
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strLocalInstalaInter) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
					strRet = strRet & "<Option value=""" & objRS("Esc_ID") & """" & strItemSel & ">" & TratarAspasJS(objRS("Cid_Sigla")) & "  " & TratarAspasJS(objRS("Esc_Sigla")) & "</Option>"
					objRS.MoveNext
				Wend
				strRet = strRet & "</select>"
				strRet = strRet & "</td>"
				strRet = strRet & "<td nowrap><font class=clsObrig>::</font>&nbsp;Distribuidor"
				strRet = strRet & "	<span id=""spnDistLocalInstalaInter"">"
				strRet = strRet & "	<select name=""cboDistLocalInstalaInter"" >"
				strRet = strRet & "	<option value=""""></option>"

				strDistrib = Request.Form("cboDistLocalInstala") 
				if Trim(strLocalInstalaInter) <> "" then 
					set objRS = db.execute("CLA_sp_view_recursodistribuicao " & strLocalInstalaInter)
					While not objRS.Eof 
						strItemSel = ""
						if Trim(strDistrib) = Trim(objRS("Dst_ID")) then strItemSel = " Selected " End if
						strRet = strRet &  "<Option value=""" & objRS("Dst_ID") & """" & strItemSel & ">" & TratarAspasJS(objRS("Dst_Desc")) & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				End if	
				strRet = strRet & "	</select>"
				strRet = strRet & "	</td>"


				strRet = strRet & "</tr>"
				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "	<td>2&nbsp;•&nbsp;<input type=text class=text name=txtParInicial2 onKeyUp=""ValidarTipo(this,0)"" maxlength=4 size=4 onblur=""CompletarCampo(this)"" TIPO=N >(N4)</td>"
				strRet = strRet & "	<td><input type=text class=text name=txtCoordenada2  maxlength=20 size=10>(A20)</td>"
				strRet = strRet & "	<td nowrap><font class=clsObrig>::</font>&nbsp;Local de Configuração"
				strRet = strRet & "	<select name=""cboLocalConfig"" onChange=""ResgatarDistribuicao(this,2)"">"
				strRet = strRet & "	<option value=""""></option>"
				set objRS = db.execute("CLA_sp_sel_estacao  null")
				While not objRS.Eof 
					strItemSel = ""
					'if Trim(strLocalConfig) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
					strRet = strRet & "<Option value=""" & objRS("Esc_ID") & """" & strItemSel & ">" & TratarAspasJS(objRS("Cid_Sigla")) & "  " & TratarAspasJS(objRS("Esc_Sigla")) & "</Option>"
					objRS.MoveNext
				Wend
				strRet = strRet & "</select>"
				strRet = strRet & "</td>"
				strRet = strRet & "<td nowrap><font class=clsObrig>::</font>&nbsp;Distribuidor"
				strRet = strRet & "	<span id=""spnDistLocalConfig"">"
				strRet = strRet & "	<select name=""cboDistLocalConfig"" >"
				strRet = strRet & "	<option value=""""></option>"
				strRet = strRet & "	</select>"
				strRet = strRet & "	</td>"
				strRet = strRet & "</tr>"
				
				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "		<td><font class=clsObrig>::</font>&nbsp;Quantidade de Pares</td>"
				strRet = strRet & "		<td><input type=text name=txtQtdePares size=5 onKeyUp=""ValidarTipo(this,0)"" class=text></td>"
				strRet = strRet & "		<td></td>"
				strRet = strRet & "		<td></td>"
				strRet = strRet & "</tr>"

				strRet = strRet & "<tr>"
				strRet = strRet & "<th colspan=4>&nbsp;Obs</th>"
				strRet = strRet & "	</tr>"

				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "	<td colspan=4><textarea name=txtObsFac onkeydown=""MaxLength(this,300);""  cols=40 rows=2></textarea>(A300)</td>"
				strRet = strRet & "</tr>"

			End if	
			strRet = strRet & "</table>"
			
	End Select
	
	ResgatarCamposFacilidade = strRet

End Function


Function ProcurarFacilidade(strRede)

	Dim strRet
	
	Call ValidarFacilidade(strRede) 
	strRet = ""

	if Trim(Request.Form("rdoInter")) = "I" then strRede = 2'Nao Determ.

	if Trim(Request.Form("rdoInter")) = "E" then

		Select Case int("0" & strRede)

			Case 1

				strRet = strRet & "<table border=0 cellspacing=1 cellpadding=0 width=760>"
				strRet = strRet & "<tr>"
				strRet = strRet & "<th  colspan=8>&nbsp;•&nbsp;PARA</th>"
				strRet = strRet & "</tr>"

				strRet = strRet & "<tr>"
				strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Fila</th>"
				strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Bastidor</th>"
				strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Regua</th>"
				strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Posicao</th>"
				strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Domínio</th>"
				strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Nó</th>"
				strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Slot</th>"
				strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Porta</th>"

				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "	<td><input type=text class=text name=txtFilaa		value=""" & TratarAspasJS(vFila_Novo) & """				maxlength=5   size=6 onKeyUp=""ValidarTipo(this,2)"">(A5)</td>"
				strRet = strRet & "	<td><input type=text class=text name=txtBastidora	value=""" & TratarAspasJS(vBastidor_Novo) & """	maxlength=3  size=4 onKeyUp=""ValidarTipo(this,0)"">(N3)</td>"
				strRet = strRet & "	<td><input type=text class=text name=txtReguaa		value=""" & TratarAspasJS(vRegua_Novo) & """		maxlength=2  size=3 onKeyUp=""ValidarTipo(this,2)"">(A2)</td>"
				strRet = strRet & "	<td><input type=text class=text name=txtPosicaoa	value=""" & TratarAspasJS(vPosicao_Novo) & """		maxlength=3  size=4 onKeyUp=""ValidarTipo(this,2)"">(A3)</td>"


				if Request.Form("hdnPlataforma") <> "LUCENT" then  
					strRet = strRet & "	<td><input type=text class=text name=txtDominioa	value=""" & TratarAspasJS(vDominio_Novo) & """		maxlength=3  size=4 onKeyUp=""ValidarTipo(this,0)"">(N3)</td>"
					strRet = strRet & "	<td><input type=text class=text name=txtNOa			value=""" & TratarAspasJS(vNo_Novo) & """			maxlength=3  size=4 onKeyUp=""ValidarTipo(this,0)"">(N3)</td>"
					strRet = strRet & "	<td><input type=text class=text name=txtSlota		value=""" & TratarAspasJS(vSlot_Novo) & """		maxlength=2   size=2 onBlur=""SearchDom(this,\'A1,A2,A3,A4,A5,A6,A7,A8,B1,B2,B3,B4,B5,B6,B7,B8\')"">(A2)</td>"
					strRet = strRet & "	<td><input type=text class=text name=txtPortaa		value=""" & TratarAspasJS(vPorta_Novo) & """		maxlength=1   size=1 onBlur=""SearchDom(this,\'A,B\')"">(A1)</td>"
				else
					strRet = strRet & "	<td><input type=text class=text name=txtDominioa	value=""" & TratarAspasJS(vDominio_Novo) & """		maxlength=10  size=8 onKeyUp=""ValidarTipo(this,2)"">(A10)</td>"
					strRet = strRet & "	<td><input type=text class=text name=txtNOa			value=""" & TratarAspasJS(vNo_Novo) & """			maxlength=2  size=2 onKeyUp=""ValidarTipo(this,0)"">(N2)</td>"
					strRet = strRet & "	<td><input type=text class=text name=txtSlota		value=""" & TratarAspasJS(vSlot_Novo) & """		maxlength=2   size=2 onBlur=""SearchDom(this,\'01,02,03,04,05,06,07,08,09,10,11,12\')"">(N2)</td>"
					strRet = strRet & "	<td><input type=text class=text name=txtPortaa		value=""" & TratarAspasJS(vPorta_Novo) & """		maxlength=2   size=2 onBlur=""SearchDom(this,\'01,02,03,04,05,06,07,08,09,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63\')"">(N2)</td>"
				end if 
				
				
				
				
				strRet = strRet & "</tr>"

				strRet = strRet & "<tr>"
				strRet = strRet & "<th><font class=clsObrig>::</font>&nbsp;Timeslot</th>"
				strRet = strRet & "<th>&nbsp;Bastidor Interno</th>"
				strRet = strRet & "<th>&nbsp;Régua Interna</th>"
				strRet = strRet & "<th>&nbsp;Porta Interna</th>"
				strRet = strRet & "<th colspan=2>&nbsp;OTS</th>"
				strRet = strRet & "<th colspan=2>&nbsp;Link</th>"
				
				strRet = strRet & "	</tr>"

				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "	<td><input type=text class=text name=txtTimeSlota	value=""" & TratarAspasJS(vTimeSlot_Novo) & """	maxlength=9 onKeyUp=""ValidarNTipo(this,0,4,4,1,0,4)"" size=10 readonly style=""background-color:#dcdcdc"">(N4-N4)</td>"
				strRet = strRet & "	<td><input type=text class=text name=txtBastidorIntera	value=""" & TratarAspasJS(vBastidorInter_Novo) & """	maxlength=20  size=10>(A20)</td>"
				strRet = strRet & "	<td><input type=text class=text name=txtReguaIntera		value=""" & TratarAspasJS(vReguaInter_Novo) & """		maxlength=20  size=10>(A20)</td>"
				strRet = strRet & "	<td><input type=text class=text name=txtPosicaoIntera	value=""" & TratarAspasJS(vPosicaoInter_Novo) & """		maxlength=20  size=10>(A20)</td>"
				strRet = strRet & "	<td colspan=2><input type=text class=text name=txtOTSa	value=""" & TratarAspasJS(vOTS_Novo) & """				maxlength=10  size=10>(A10)</td>"
				strRet = strRet & "	<td colspan=2><input type=text class=text name=txtLinka			value=""" & TratarAspasJS(vLink_Novo) & """	 maxlength=25 size=25>(A25)</td>"
				
				strRet = strRet & "</tr>"

				strRet = strRet & "<tr>"
				strRet = strRet & "<th colspan=2>&nbsp;Designação do Tronco</th>"
				strRet = strRet & "<th colspan=2>&nbsp;Identificação Provedor</th>"
				strRet = strRet & "<th colspan=4>&nbsp;Obs</th>"
				strRet = strRet & "	</tr>"

				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "	<td colspan=2><input type=text class=text name=txtDesigTroncoa	value=""" & TratarAspasJS(vDesigTronco_Novo) & """		maxlength=25 size=25>(A25)</td>"

				strRet = strRet & "	<td colspan=2><input type=text class=text name=txtIdentProvedora	value=""" & TratarAspasJS(vIdentProvedor_Novo) & """	 maxlength=20 size=20>(A20)</td>"

				strRet = strRet & "	<td colspan=4><textarea name=txtObsFaca onkeydown=""MaxLength(this,300);""  cols=40 rows=2>" & TratarAspasJS(vObsFac_Novo) & "</textarea>(A300)</td>"
				strRet = strRet & "</tr>"

				strRet = strRet & "</table>"
				
				'Atualiza campos não obrigatórios
				Response.Write "<script laguange=javascript> with (parent.document.forms[0]) {" & _
								 "txtBastidorInter.value=""" & TratarAspasJS(vBastidorInter_Novo) & """; " & _
								 "txtReguaInter.value=""" & TratarAspasJS(vReguaInter_Novo) & """; " & _
								 "txtPosicaoInter.value=""" & TratarAspasJS(vPosicaoInter_Novo) & """; " & _
								 "txtFila.value=""" & TratarAspasJS(vFila_Novo) & """; " & _
								 "txtOTS.value=""" & TratarAspasJS(vOTS_Novo) & """; " & _
								 "txtDesigTronco.value=""" & TratarAspasJS(vDesigTronco_Novo) & """; " & _
								 "txtLink.value=""" & TratarAspasJS(vLink_Novo) & """; " & _

								 "txtIdentProvedor.value=""" & TratarAspasJS(vIdentProvedor_Novo) & """; " & _

								 "txtObsFac.value=""" & TratarAspasJS(vObsFac_Novo) & """; " & _
								 "}</script>"
				if Trim(vTimeSlot_Novo) <> Trim(vTimeSlot) then 'Atualiza o procurado quando os timeslots são diferentes
					Response.Write "<script laguange=javascript>parent.document.forms[0].txtTimeSlot.value=""" & TratarAspasJS(vTimeSlot_Novo) & """</script>"
				End if
				
			Case 2
			
				'Não deterministico
				strRet = ""
				strRet = strRet & "<table border=0 cellspacing=1 cellpadding=0 width=760>"
				strRet = strRet & "	<tr class=clsSilver>"
				strRet = strRet & "		<th width =50% ><font class=clsObrig>::</font>&nbsp;Tronco</th>"
				strRet = strRet & "		<th width =50% ><font class=clsObrig>::</font>&nbsp;Par</th>"
				strRet = strRet & "	</tr>"
				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "	<td><input type=text class=text name=txtTroncoa	value=""" & TratarAspasJS(vTronco_Novo) & """	maxlength=20 size=10>(A20)</td>"
				strRet = strRet & "	<td><input type=text class=text name=txtPara	value=""" & TratarAspasJS(vPar_Novo) & """		maxlength=9 size=10 onKeyUp=""if(this.value.length > 1)ValidarNTipo(this,0,4,4,1,0,4)"" TIPO=N >(N4-N4) ou (A1)</td>"
				strRet = strRet & "</tr>"
				if Trim(vPar_Novo) <> Trim(vPar) then 'Atualiza o procurado quando o time slot são diferentes
					Response.Write "<script laguange=javascript>parent.document.forms[0].txtPar.value=""" & TratarAspasJS(vPar_Novo) & """</script>"
				End if
				strRet = strRet & "</table>" 

			Case 3

				strRet = ""
				strRet = strRet & "<table border=0 cellspacing=1 cellpadding=0 width=760>"
				strRet = strRet & "	<tr>"
				strRet = strRet & "<th  colspan=5>&nbsp;•&nbsp;PARA</th>"
				strRet = strRet & "	</tr>"

				strRet = strRet & "	<tr class=clsSilver >"
				strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Cabo</th>"
				strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Par</th>"
				strRet = strRet & "		<th nowrap><font class=clsObrig>::</font>&nbsp;Tipo do Cabo</th>"
				strRet = strRet & "		<th><font class=clsObrig>::</font>&nbsp;Derivação</th>"
				strRet = strRet & "		<th nowrap><font class=clsObrig>::</font>&nbsp;PADE</th>"
				strRet = strRet & "	</tr>"
				strRet = strRet & "<tr class=clsSilver>"
				strRet = strRet & "<td width=129><input type=text class=text		name=txtTroncoa	value=""" & TratarAspasJS(vTronco_Novo) & """	maxlength=2 size=2 onKeyUp=""parent.ValidarTipo(this,0)"" onblur=""CompletarCampo(this)"" TIPO=N >(N2)</td>"
				strRet = strRet & "<td><input type=text class=text name=txtPara 	maxlength=9		value=""" & TratarAspasJS(vPar_Novo) & """		size=9 onKeyUp=""parent.ValidarNTipo(this,0,4,4,1,0,4)"">(N4-N4)</td>"
				strRet = strRet & "<td><select name=cboTipoCaboa >"
											strH = ""
											strX = ""
											strC = ""
											Select Case vTipoCabo_Novo
												Case "H"
													strH = " selected "
												Case "X"
													strX = " selected "
												Case "C"
													strC = " selected "	
											End Select
				strRet = strRet & "			<option value=""""></option>"
				strRet = strRet & "			<option value=H" & strH & " >H</option>"
				strRet = strRet & "			<option value=X" & strX & " >X</option>"
				strRet = strRet & "			<option value=C" & strC & " >C</option>"
				strRet = strRet & "		</select></td>"
				strRet = strRet & "<td><input type=text class=text name=txtLaterala  value=""" & TratarAspasJS(vLateral_Novo) & """ onKeyUp=""parent.ValidarTipo(this,2)""	maxlength=4 size=5>(A4)</td>"
				strRet = strRet & "<td><input type=text class=text name=txtCaixaEmendaa  value=""" & TratarAspasJS(vCxEmenda_Novo) & """  onKeyUp=""parent.ValidarTipo(this,2)""  maxlength=3 size=5 >(A3)</td>"
				strRet = strRet & "</tr>"

				if Trim(vPar_Novo) <> Trim(vPar) then 'Atualiza o procurado quando o time slot são diferentes
					Response.Write "<script laguange=javascript>parent.document.forms[0].txtPar.value=""" & TratarAspasJS(vPar_Novo) & """</script>"
				End if

				strRet = strRet & "</table>"

				
			Case 12 'PADE/PAC

				strRet = ""
				strRet = strRet & " <table  border=0 cellspacing=1 cellpadding=1 width=760> "
				strRet = strRet & "	<tr>"
				strRet = strRet & " <th  colspan=5><font class=clsObrig>::</font>&nbsp; PARA SWITCH Concentrador</th> "
				strRet = strRet & "	</tr> "
				
				strRet = strRet & "	<tr class=clsSilver> "
				strRet = strRet & "		<td colspan=5><font class=clsObrig>::</font>&nbsp;Designação do SWITCH Concentrador&nbsp;<input type=text name=txtSWITCHa size=20  class=text value=""" & TratarAspasJS(vtxtSWITCHa) & """ ></td> "
				strRet = strRet & "	</tr> "
	
				
				strRet = strRet & "	 <tr class=clsSilver> "
				strRet = strRet & "		<th></th> "
				strRet = strRet & "		<th></th> "
				strRet = strRet & "		<th></th> "
				strRet = strRet & "	</tr> "
				strRet = strRet & " <tr class=clsSilver> "
								
	
				strRet = strRet & "	 <td nowrap><font class=clsObrig>::</font>&nbsp;Fabricante&nbsp;&nbsp;&nbsp;&nbsp; "
				strRet = strRet & "	<select name=cboFabricanteONTa onchange=ResgatarTipoONTa() > "
				
				strRet = strRet & "	<option value=""""></option> "
				
				
				set objRS = db.execute("CLA_sp_sel_FabricanteONT 0 , null , null , null , " & int("0" & strRede) )
				
					
				While Not objRS.eof
				  strItemSel = ""
				  if Trim(vFont_ID) = Trim(objRS("Font_ID")) then strItemSel = " Selected " End if
				  strRet = strRet & " <Option value=""" &  objRS("Font_ID") & """" & strItemSel & "> " & Trim(objRS("Font_Nome")) &  " </Option> "
				  objRS.MoveNext
				Wend
				strItemSel = ""
				
				strRet = strRet & " </select> "
				strRet = strRet & " </td> "
				
				'strRet = strRet & " <td nowrap colspan=2> <span ID=spncboTipoONTa></span> "
				
				'strRet = strRet & "<tr class=""clsSilver"">"
				strRet = strRet & "	 <td nowrap><font class=clsObrig>::</font>&nbsp;Modelo&nbsp;&nbsp;&nbsp;&nbsp; "
				'strRet = strRet & "	<td nowrap colspan=2 >&nbsp;&nbsp;Modelo</td><td>&nbsp;&nbsp; <span ID=spncboTipoONTa> "
				strRet = strRet & " <span ID=spncboTipoONTa> "
				
				set objRS = db.execute("CLA_sp_sel_TipoONT null," & vFont_ID)
		
					
				strRet = strRet & " <select name=cboTipoONTa > "
				strRet = strRet & "	<option value=""""></option> "
				
				While Not objRS.eof
				  strItemSel = ""
				  if Trim(vTont_ID) = Trim(objRS("Tont_ID")) then strItemSel = " Selected " End if
				  strRet = strRet & " <Option value=""" &  objRS("Tont_ID") & """" & strItemSel & "> " & Trim(objRS("Tont_Modelo")) &  " </Option> "
				  objRS.MoveNext
				Wend
				strItemSel = ""
					
				''While Not objRScbo.eof
				'	'strItemSel = ""
				'	'if Trim(vTont_ID) = Trim(objRScbo("Tont_ID")) then strItemSel = " Selected " End if
				''		strRet = strRet & " <Option value='" & objRScbo("Tont_ID") & "'"
						
				''		if objRScbo("Tont_ID") = vTont_ID then 
				''  			strRet = strRet & " selected "
				''		end if
						
				''		strRet = strRet & " > " & Trim(objRScbo("Tont_Modelo")) & " </Option> "
				''	objRScbo.MoveNext
				''Wend
				''strItemSel = ""
				
				strRet = strRet & " </select> "
					
				strRet = strRet & " </span> "
				
				
				
				strRet = strRet & "	</td>"
				strRet = strRet & " </tr> "
	
				strRet = strRet & " <tr class=clsSilver> "
				strRet = strRet & "	<td nowrap><font class=clsObrig>::</font>&nbsp;Porta PE "
				strRet = strRet & "	<input type=text class=text name=txtPortaPEa maxlength=20 size=20 value=""" & TratarAspasJS(vtxtPortaPEa) & """ >(A20)</td> "
				strRet = strRet & " <td nowrap colspan=2><font class=clsObrig>::</font>&nbsp;PE&nbsp; "
				strRet = strRet & "	<input type=text class=text name=txtPEa  maxlength=13 size=13 value=""" & TratarAspasJS(vtxtPEa) & """ >(A13)</td> "
				strRet = strRet & " </tr> "
	
				
				strRet = strRet & " </table> "


		End Select

	Else
		strDistrib = vDistInter1_Novo
		
		call RetornaDistribuidor (strDistInstala , strDistConfig )
		
		Response.Write "<script laguange=javascript> with (parent.document.forms[0]) {" & _
						"txtParInicial1.value = """ & TratarAspasJS(vParInter1_Novo) & """;" & _
						"txtCoordenada1.value = """ & TratarAspasJS(vCoodInter1_Novo) & """;" & _
						"txtParInicial2.value = """ & TratarAspasJS(vParInter2_Novo) & """;" & _
						"txtCoordenada2.value = """ & TratarAspasJS(vCoodInter2_Novo) & """;" & _
						"cboLocalInstalaInter.value = """ & TratarAspasJS(vEscInter1_Novo) & """;" & _
						"cboLocalConfig.value = """ & TratarAspasJS(vEscInter2_Novo) & """;" & _
						"parent.spnDistLocalInstalaInter.innerHTML = '" & strDistInstala & "';" & _
						"parent.spnDistLocalConfig.innerHTML = '" & strDistConfig & "';" & _
						"txtObsFac.value = """ & TratarAspasJS(vFacObs_Novo) & """;" & _
						"}</script>"

		strRet = ""
		strRet = strRet & "<table border=0 cellspacing=1 cellpadding=0 width=760>"
		strRet = strRet & "	<tr>"
		strRet = strRet & "<th  colspan=5>&nbsp;•&nbsp;PARA</th>"
		strRet = strRet & "	</tr>"

		strRet = ResgatarCamposInterligacao(strRede , strDistInstala , strDistConfig) 
	end if 
	
		ProcurarFacilidade = strRet
End Function

sub RetornaDistribuidor (strDistInstala , strDistConfig )

		Dim xstrDistInstala
		Dim xstrDistConfig
			'Atualiza campos não obrigatórios
		xstrDistInstala = ""
		xstrDistInstala = xstrDistInstala & "	<select name=""cboDistLocalInstalaInter"" >"
		xstrDistInstala = xstrDistInstala & "	<option value=""""></option>"
		
		xstrDistConfig = xstrDistConfig & "	<select name=""cboDistLocalConfig"" >"
		xstrDistConfig = xstrDistConfig & "	<option value=""""></option>"

		if Trim(vEscInter1_Novo) <> "" then 
		
			set objRS = db.execute("CLA_sp_sel_distribuicao ")
			
			While not objRS.Eof 
				strItemSel = ""
				if Trim(vDistInter1_Novo) = Trim(objRS("Dst_ID")) then strItemSel = " Selected " End if
				xstrDistInstala = xstrDistInstala &  "<Option value=""" & objRS("Dst_ID") & """" & strItemSel & ">" & TratarAspasJS(objRS("Dst_Desc")) & "</Option>"
				
				strItemSel = ""
				if Trim(vDistInter2_Novo) = Trim(objRS("Dst_ID")) then strItemSel = " Selected " End if
				xstrDistConfig = xstrDistConfig &  "<Option value=""" & objRS("Dst_ID") & """" & strItemSel & ">" & TratarAspasJS(objRS("Dst_Desc")) & "</Option>"
				
				objRS.MoveNext
			Wend
			strItemSel = ""
			
		End if	
		strDistInstala = xstrDistInstala & "	</select>"
		strDistConfig = xstrDistConfig & "	</select>"
end sub  


Sub ValidarFacilidade(strRede)
	
	Dim objRSRec
	Dim confirma

	if Trim(Request.Form("rdoInter")) = "E" or Trim(Request.Form("rdoInter")) = "" then

		'Não é Coordenada
		strPlataforma  = request("cboPlataforma")
		if strPlataforma  = "" then  strPlataforma = "null"
		
		Set objRSRec = db.execute("CLA_sp_check_recurso2 " & request("cboLocalInstala") & ", " & request("cboDistLocalInstala") & ", " & request("cboProvedor") & ", " & request("cboSistema") & "," & strPlataforma & ",0")
		Set DBAction = objRSRec("ret")

		If DBAction = 0 then

			vRecurso	= objRSRec("Rec_ID")
			'NDet /Ade
			vTronco		= Request.Form("txtTronco")
			vPar		= Request.Form("txtPar")
			'Det
			vBastidor	= Request.Form("txtBastidor")
			vRegua		= Request.Form("txtRegua")
			vPosicao	= Request.Form("txtPosicao")
			vTimeSlot	= Request.Form("txtTimeSlot")
			vDominio	= Request.Form("txtDominio")
			vNo			= Request.Form("txtNO")
			vSlot		= Request.Form("txtSlot")
			vPorta		= Request.Form("txtPorta")
			vBastidorInter	= Request.Form("txtBastidorInter")
			vReguaInter		= Request.Form("txtReguaInter")
			vPosicaoInter		= Request.Form("txtPosicaoInter")
			vFila			= Request.Form("txtFila")
			vOTS			= Request.Form("txtOTS")
			vDesigTronco	= Request.Form("txtDesigTronco")
			vLink			= Request.Form("txtLink")
			vObsFac			= Request.Form("txtObsFac")
			vIdentProvedor	= Request.Form("txtIdentProvedor")

			vTipoCabo	= Request.Form("cboTipoCabo")
			vLateral	= Request.Form("txtLateral")	
			vCxEmenda	= Request.Form("txtCaixaEmenda")
			
			vtxtSWITCH	= Request.Form("txtSWITCH")
			
			Vetor_Campos(1)="adInterger,2,adParamInput," & vRecurso
			Vetor_Campos(2)="adWChar,20,adParamInput," & vTronco  
			Vetor_Campos(3)="adWChar,20,adParamInput," & vPar     
			Vetor_Campos(4)="adWChar,20,adParamInput," & vBastidor
			Vetor_Campos(5)="adWChar,20,adParamInput," & vRegua   
			Vetor_Campos(6)="adWChar,20,adParamInput," & vPosicao 
			Vetor_Campos(7)="adWChar,20,adParamInput," & vTimeslot
			Vetor_Campos(8)="adWChar,20,adParamInput," & vDominio 
			Vetor_Campos(9)="adWChar,20,adParamInput," & vNo      
			Vetor_Campos(10)="adWChar,2,adParamInput," & vSlot    
			Vetor_Campos(11)="adWChar,2,adParamInput," & vPorta   
			Vetor_Campos(12)="adWChar,20,adParamInput," & vTipoCabo
			Vetor_Campos(13)="adWChar,20,adParamInput," & vLateral 
			Vetor_Campos(14)="adWChar,20,adParamInput," & vCxEmenda
			Vetor_Campos(15)="adInteger,2,adParamInput," & strRede
			Vetor_Campos(16)="adWChar,30,adParamInput," & vtxtSWITCH

			Dim strSqlRet
			strSqlRet = APENDA_PARAMSTR("CLA_sp_cons_Facilidade",16,Vetor_Campos)
			Set objRS = db.Execute(strSqlRet)			
			
			'dim objFSO, objFile, strCaminho 
			'Set objFSO = CreateObject("Scripting.FileSystemObject")
			'strCaminho = server.MapPath("..\")
			'Set objFile = objFSO.CreateTextFile(strCaminho & "\CartasProvedor\consutageral.txt",  true)
			'objFile.WriteLine(strSqlRet)
			'objFile.Close
		
		
			If objRS.EOF and objRS.BOF then
				Response.Write "<script language=javascript>alert('Facilidade não encontrada.')</script>"
				Response.Write "<script language=javascript>parent.spnCamposFacRet.innerHTML = ''</script>"
				Response.Write "<script language=javascript>parent.spnConsulta.innerHTML = ''</script>"
				Response.End 
			Else
				objRS.Close
				objRS.CursorLocation = AdUseClient
				objRS.open
				Select Case strRede
					Case 1
						vBastidor_Novo	= objRS("Fac_Bastidor")
						vRegua_Novo		= objRS("Fac_Regua")
						vPosicao_Novo	= objRS("Fac_Posicao")
						vDominio_Novo	= objRS("Fac_Dominio")
						vNo_Novo		= objRS("Fac_No")
						vSlot_Novo		= objRS("Fac_Slot")
						vPorta_Novo		= objRS("Fac_Porta")
						vBastidorInter_Novo	= objRS("Fac_BasInter")
						vReguaInter_Novo	= objRS("Fac_RegInter")
						vPosicaoInter_Novo	= objRS("Fac_PosInter")
						vFila_Novo			= objRS("Fac_Fila")
						vOTS_Novo			= objRS("Fac_OTS")
						vDesigTronco_Novo	= objRS("Fac_DesigTronco")
						vLink_Novo			= objRS("Fac_Link")
						vObsFac_Novo		= objRS("Fac_Obs")
						vIdentProvedor_Novo	= objRS("Fac_IdentProvedor")

						intInicio		= objRS("Fac_TimeSlot")
						objRS.MoveLast
						intFim			= objRS("Fac_TimeSlot")
						if intInicio <> intFim then
							vTimeSlot_Novo	= intInicio & "-" & intFim
						Else
							vTimeSlot_Novo	= intInicio
						End if	
					Case 2
						vTronco_Novo	= objRS("Fac_Tronco")
						intInicio		= objRS("Fac_Par")
						objRS.MoveLast
						intFim			= objRS("Fac_Par")
						if intInicio <> intFim then
							vPar_Novo		= intInicio & "-" & intFim
						Else
							vPar_Novo		= intInicio
						End if	
					Case 3
						vTronco_Novo	= objRS("Fac_Tronco")
						intInicio		= objRS("Fac_Par")
						objRS.MoveLast
						intFim			= objRS("Fac_Par")
						if intInicio <> intFim then
							vPar_Novo		= intInicio & "-" & intFim
						Else
							vPar_Novo		= intInicio
						End if	
						vTipoCabo_Novo	= objRS("Fac_TipoCabo")
						vLateral_Novo	= objRS("Fac_Lateral")
						vCxEmenda_Novo	= objRS("Fac_CxEmenda")
						
					Case 12
						vtxtSWITCHa	= objRS("switch")
						vTont_id	= objRS("tont_id")
						vFont_ID	= objRS("font_id")
						vtxtPortaPEa	= objRS("OntVlan_PortaOLT")
						vtxtPEa	= objRS("ontVlan_PE")
						
				End Select		
				'objRS.movefirst
				'Call ListaFacilidades(strRede,"100%",false,objRS)
			End if
		Else
			Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
			Response.End 
		End if
	Else
		'PADE/PAC
		vParInter1	= Request.Form("txtParInicial1")
		vCoodInter1 = Request.Form("txtCoordenada1")
		vEscInter1	= Request.Form("cboLocalInstalaInter")
		vDistInter1 = Request.Form("cboDistLocalInstalaInter")

		vParInter2	= Request.Form("txtParInicial2")
		vCoodInter2 = Request.Form("txtCoordenada2")
		vEscInter2	= Request.Form("cboLocalConfig")
		vDistInter2 = Request.Form("cboDistLocalConfig")
		vQtdePares	= Request.Form("txtQtdePares")

		Vetor_Campos(1)="adInterger,2,adParamInput," & vEscInter1
		Vetor_Campos(2)="adWChar,20,adParamInput," & vEscInter2
		Vetor_Campos(3)="adWChar,20,adParamInput," & vDistInter1
		Vetor_Campos(4)="adWChar,20,adParamInput," & vDistInter2
		Vetor_Campos(5)="adWChar,20,adParamInput," & vParInter1
		Vetor_Campos(6)="adWChar,20,adParamInput," & vCoodInter1
		Vetor_Campos(7)="adWChar,20,adParamInput," & vParInter2
		Vetor_Campos(8)="adWChar,20,adParamInput," & vCoodInter2
		Vetor_Campos(9)="adWChar,20,adParamInput," & vQtdePares
		
		strSqlRet = APENDA_PARAMSTR("CLA_sp_cons_interligacao",9,Vetor_Campos)
		Set confirma = db.Execute(strSqlRet)

		'Set confirma = db.execute("'" & vParInter1 & "','" & vCoodInter1 & "','" & vParInter2 & "','" & vCoodInter2 & "'," & vQtdePares)

		If confirma.EOF and confirma.BOF then
			Response.Write "<script language=javascript>alert('Facilidade não encontrada.')</script>"
			Response.Write "<script language=javascript>parent.spnCamposFacRet.innerHTML = ''</script>"
			Response.Write "<script language=javascript>parent.spnConsulta.innerHTML = ''</script>"
			Response.End 
		Else

			'ListaInterligacao(confirma)

			objAryAux = split(confirma("Int_CorOrigem"),"-")
			vParInter1_Novo	= objAryAux(UBound(objAryAux))
			vCoodInter1_Novo= Left(confirma("Int_CorOrigem"),len(confirma("Int_CorOrigem"))-5)
			vEscInter1_Novo	= confirma("Esc_IdOrigem")
			vDistInter1_Novo= confirma("Dst_IdOrigem")

			objAryAux = split(confirma("Int_CorDestino"),"-")
			vParInter2_Novo	= objAryAux(UBound(objAryAux))
			vCoodInter2_Novo= Left(confirma("Int_CorDestino"),len(confirma("Int_CorDestino"))-5)
			vEscInter2_Novo	= confirma("Esc_IdDestino")
			vDistInter2_Novo= confirma("Dst_IdDestino")

			vFacObs_Novo	= confirma("Int_Obs")

		End if			
	End if

End Sub

Function ListaInterligacao(objRSIn)

	strInterHtml = strInterHtml & "<table border=0 cellspacing=1 cellpadding=0 width=100% >"
	strInterHtml = strInterHtml & "	<tr>"
	strInterHtml = strInterHtml & "<th  colspan=2>&nbsp;•&nbsp;Interligações Encontradas</th>"
	strInterHtml = strInterHtml & "	</tr>"
	strInterHtml = strInterHtml & "<tr>"
	strInterHtml = strInterHtml & "<th>Origem</th>"
	strInterHtml = strInterHtml & "<th>Destino</th>"
	strInterHtml = strInterHtml & "</tr>"

	Dim vQtdeParesTemp
	vQtdeParesTemp = 0
	While not objRSIn.Eof
		strInterHtml = strInterHtml & "<tr class=clsSilver>"
		strInterHtml = strInterHtml & "<td>" & objRSIn("Int_CorOrigem") & "</td>"
		strInterHtml = strInterHtml & "<td>" & objRSIn("Int_CorDestino") & "</td>"
		strInterHtml = strInterHtml & "</tr>"
		objRSIn.movenext
		vQtdeParesTemp = cint(vQtdeParesTemp) + 1

	Wend
	set objRSIn = nothing

	strInterHtml = strInterHtml & "<tr>"
	strInterHtml = strInterHtml & "<th>Total de Pares</th>"
	strInterHtml = strInterHtml & "<th>" & vQtdeParesTemp & "</th>"
	strInterHtml = strInterHtml & "</tr>"

	Response.Write "<script language=javascript>parent.spnConsulta.innerHTML = '" & strInterHtml & "';</script>"
					
End function

Function AlterarFacilidade(strRede)

	Dim objRSRec1
	Dim objRSRec2
	Dim strInter
	Dim blnProcurar
	Dim Update
	Dim strRet

	if Trim(Request.Form("rdoInter")) = "E" or Trim(Request.Form("rdoInter")) = "" then
		
		vTronco		= Request.Form("txtTronco")
		vPar		= Request.Form("txtPar")
		'Det
		vBastidor	= Request.Form("txtBastidor")
		vRegua		= Request.Form("txtRegua")
		vPosicao	= Request.Form("txtPosicao")
		vTimeSlot	= Request.Form("txtTimeSlot")
		vDominio	= Request.Form("txtDominio")
		vNo			= Request.Form("txtNO")
		vSlot		= Request.Form("txtSlot")
		vPorta		= Request.Form("txtPorta")
		vBastidorInter	= Request.Form("txtBastidorInter")
		vReguaInter		= Request.Form("txtReguaInter")
		vPosicaoInter		= Request.Form("txtPosicaoInter")
		vFila			= Request.Form("txtFila")
		vOTS			= Request.Form("txtOTS")
		vDesigTronco	= Request.Form("txtDesigTronco")
		vLink			= Request.Form("txtLink")
		vIdentProvedor	= Request.Form("txtIdentProvedor")
		vTipoCabo	= Request.Form("cboTipoCabo")
		vLateral	= Request.Form("txtLateral")	
		vCxEmenda	= Request.Form("txtCaixaEmenda")

		'NDet /Ade
		vTronco_Novo	= Request.Form("txtTroncoa")
		vPar_Novo		= Request.Form("txtPara")
		'Det
		vBastidor_Novo	= Request.Form("txtBastidora")
		vRegua_Novo		= Request.Form("txtReguaa")
		vPosicao_Novo	= Request.Form("txtPosicaoa")
		vTimeSlot_Novo	= Request.Form("txtTimeSlota")
		vDominio_Novo	= Request.Form("txtDominioa")
		vNo_Novo		= Request.Form("txtNOa")
		vSlot_Novo		= Request.Form("txtSlota")
		vPorta_Novo		= Request.Form("txtPortaa")
		vBastidorInter_Novo	= Request.Form("txtBastidorIntera")
		vReguaInter_Novo	= Request.Form("txtReguaIntera")
		vPosicaoInter_Novo	= Request.Form("txtPosicaoIntera")
		vFila_Novo			= Request.Form("txtFilaa")
		vOTS_Novo			= Request.Form("txtOTSa")
		vDesigTronco_Novo	= Request.Form("txtDesigTroncoa")
		vLink_Novo			= Request.Form("txtLinka")
		vObsFac_Novo		= Request.Form("txtObsFaca")
		vIdentProvedor_Novo	= Request.Form("txtIdentProvedora")

		vTipoCabo_Novo	= Request.Form("cboTipoCaboa")
		vLateral_Novo	= Request.Form("txtLaterala")
		vCxEmenda_Novo	= Request.Form("txtCaixaEmendaa")
		vObsFacNovo		= Request.Form("txtObsFaca")
		vObsFac			= Request.Form("txtObsFac")
		
		vtxtSWITCHa	= Request.Form("txtSWITCHa")
		vFont_ID	= Request.Form("cboFabricanteONTa")
		vTont_id	= Request.Form("cboTipoONTa")
		vtxtPortaPEa		= Request.Form("txtPortaPEa")
		vtxtPEa			= Request.Form("txtPEa")		
		
		
		
		'Resgata o Recurso
		
		strPlataforma  = request("cboPlataforma")
		if strPlataforma  = "" then  strPlataforma = "null"
		
		Set objRSRec = db.execute("CLA_sp_check_recurso2 " & request("cboLocalInstala") & ", " & request("cboDistLocalInstala") & ", " & request("cboProvedor") & ", " & request("cboSistema") & "," & strPlataforma & ",0")
		Set DBAction = objRSRec("ret")

		If DBAction <> 0 then
			Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
			Response.End 
		End if

			vRecurso	= objRSRec("Rec_ID")

			Vetor_Campos(1)="adInteger,2,adParamInput," & vRecurso
			Vetor_Campos(2)="adWChar,20,adParamInput," & vTronco
			Vetor_Campos(3)="adWChar,20,adParamInput," & vPar  
			Vetor_Campos(4)="adWChar,20,adParamInput," & vBastidor  
			Vetor_Campos(5)="adWChar,20,adParamInput," & vRegua  
			Vetor_Campos(6)="adWChar,20,adParamInput," & vPosicao  
			Vetor_Campos(7)="adWChar,20,adParamInput," & vTimeslot  
			Vetor_Campos(8)="adWChar,20,adParamInput," & vDominio  
			Vetor_Campos(9)="adWChar,20,adParamInput," & vNo  
			Vetor_Campos(10)="adWChar,2,adParamInput," & vSlot
			Vetor_Campos(11)="adWChar,2,adParamInput," & vPorta
			Vetor_Campos(12)="adWChar,20,adParamInput," & vBastidorInter  
			Vetor_Campos(13)="adWChar,20,adParamInput," & vReguaInter  
			Vetor_Campos(14)="adWChar,20,adParamInput," & vPosicaoInter  
			Vetor_Campos(15)="adWChar,5,adParamInput," & vFila  
			Vetor_Campos(16)="adWChar,10,adParamInput," & vOTS  
			Vetor_Campos(17)="adWChar,25,adParamInput," & vDesigTronco  
			Vetor_Campos(18)="adWChar,25,adParamInput," & vLink  
			Vetor_Campos(19)="adWChar,20,adParamInput," & vTipoCabo  
			Vetor_Campos(20)="adWChar,20,adParamInput," & vLateral  
			Vetor_Campos(21)="adWChar,20,adParamInput," & vCxEmenda  
			Vetor_Campos(22)="adWChar,20,adParamInput," & vTronco_Novo  
			Vetor_Campos(23)="adWChar,20,adParamInput," & vPar_Novo
			Vetor_Campos(24)="adWChar,20,adParamInput," & vBastidor_Novo
			Vetor_Campos(25)="adWChar,20,adParamInput," & vRegua_Novo  
			Vetor_Campos(26)="adWChar,20,adParamInput," & vPosicao_Novo  
			Vetor_Campos(27)="adWChar,20,adParamInput," & vTimeSlot_Novo
			Vetor_Campos(28)="adWChar,20,adParamInput," & vDominio_Novo  
			Vetor_Campos(29)="adWChar,20,adParamInput," & vNo_Novo
			Vetor_Campos(30)="adWChar,2,adParamInput," & vSlot_Novo
			Vetor_Campos(31)="adWChar,2,adParamInput," & vPorta_Novo
			Vetor_Campos(32)="adWChar,20,adParamInput," & vBastidorInter_Novo
			Vetor_Campos(33)="adWChar,20,adParamInput," & vReguaInter_Novo
			Vetor_Campos(34)="adWChar,20,adParamInput," & vPosicaoInter_Novo
			Vetor_Campos(35)="adWChar,5,adParamInput," & vFila_Novo
			Vetor_Campos(36)="adWChar,10,adParamInput," & vOTS_Novo
			Vetor_Campos(37)="adWChar,25,adParamInput," & vDesigTronco_Novo
			Vetor_Campos(38)="adWChar,25,adParamInput," & vLink_Novo
			Vetor_Campos(39)="adWChar,20,adParamInput," & vTipoCabo_Novo
			Vetor_Campos(40)="adWChar,20,adParamInput," & vLateral_Novo
			Vetor_Campos(41)="adWChar,20,adParamInput," & vCxEmenda_Novo
			Vetor_Campos(42)="adlongvarchar,300,adParamInput," & vObsFac
			Vetor_Campos(43)="adlongvarchar,300,adParamInput," & vObsFacNovo
			Vetor_Campos(44)="adInteger,2,adParamInput," & strRede
			Vetor_Campos(45)="adWChar,2,adParamInput," & trim(vIdentProvedor_Novo)
			
			
			Vetor_Campos(46)="adWChar,30,adParamInput," & vtxtSWITCHa
			Vetor_Campos(47)="adInteger,2,adParamInput," & vFont_ID
			Vetor_Campos(48)="adInteger,2,adParamInput," & vTont_id
			Vetor_Campos(49)="adWChar,20,adParamInput," & vtxtPortaPEa
			Vetor_Campos(50)="adWChar,13,adParamInput," & vtxtPEa
			
			''Vetor_Campos(45)="adInteger,2,adParamInput," & trim(vIdentProvedor)
			
			
			
			Dim strSqlRet
			strSqlRet = APENDA_PARAMSTR("CLA_sp_upd_Facilidade",50,Vetor_Campos)
			
			'Response.Write "<script language=javascript>alert('" & vIdentProvedor_Novo & "' );</script>"
			

			Set objRS = db.Execute(strSqlRet)

			
			Select Case strRede

				Case 1
					strRet = strRet & "parent.document.forms[0].txtBastidor.value=""" & TratarAspasJS(vBastidor_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtRegua.value=""" & TratarAspasJS(vRegua_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtPosicao.value=""" & TratarAspasJS(vPosicao_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtTimeSlot.value=""" & TratarAspasJS(vTimeSlot_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtDominio.value=""" & TratarAspasJS(vDominio_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtNO.value=""" & TratarAspasJS(vNo_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtSlot.value=""" & TratarAspasJS(vSlot_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtPorta.value=""" & TratarAspasJS(vPorta_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtBastidorInter.value=""" & TratarAspasJS(vBastidorInter_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtReguaInter.value=""" & TratarAspasJS(vReguaInter_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtPosicaoInter.value=""" & TratarAspasJS(vPosicaoInter_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtFila.value=""" & TratarAspasJS(vFila_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtOTS.value=""" & TratarAspasJS(vOTS_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtDesigTronco.value=""" & TratarAspasJS(vDesigTronco_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtLink.value=""" & TratarAspasJS(vLink_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtObsFac.value=""" & TratarAspasJS(vObsFac_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtIdentProvedor.value=""" & TratarAspasJS(vIdentProvedor_Novo) & """;"

				Case 2
					strRet = strRet & "parent.document.forms[0].txtTronco.value=""" & TratarAspasJS(vTronco_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtPar.value=""" & TratarAspasJS(vPar_Novo) & """;"

				Case 3
					strRet = strRet & "parent.document.forms[0].txtTronco.value=""" & TratarAspasJS(vTronco_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtPar.value=""" & TratarAspasJS(vPar_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].cboTipoCabo.value=""" & TratarAspasJS(vTipoCabo_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtLateral.value=""" & TratarAspasJS(vLateral_Novo) & """;"
					strRet = strRet & "parent.document.forms[0].txtCaixaEmenda.value=""" & TratarAspasJS(vCxEmenda_Novo) & """;"

			End Select		

			Response.Write "<script language=javascript>alert('Registro(s) atualizado(s) com sucesso.');" & strRet & "</script>"

	Else

		vParInter1	= Request.Form("txtParInicial1")
		vCoodInter1 = Request.Form("txtCoordenada1")
		vEscInter1	= Request.Form("cboLocalInstalaInter")
		vDistInter1 = Request.Form("cboDistLocalInstalaInter")

		vParInter2	= Request.Form("txtParInicial2")
		vCoodInter2 = Request.Form("txtCoordenada2")
		vEscInter2	= Request.Form("cboLocalConfig")
		vDistInter2 = Request.Form("cboDistLocalConfig")

		vParInter1_Novo	= Request.Form("txtParInicial1a")
		vCoodInter1_Novo= Request.Form("txtCoordenada1a")
		vEscInter1_Novo	= Request.Form("cboLocalInstalaIntera")
		vDistInter1_Novo= Request.Form("cboDistLocalInstalaIntera")

		vParInter2_Novo	= Request.Form("txtParInicial2a")
		vCoodInter2_Novo= Request.Form("txtCoordenada2a")
		vEscInter2_Novo	= Request.Form("cboLocalConfiga")
		vDistInter2_Novo= Request.Form("cboDistLocalConfiga")
		vFacObs_Novo	= Request.Form("txtObsFaca")

		vQtdePares	= Request.Form("txtQtdePares")

		Vetor_Campos(1)="adInteger,4,adParamInput," & vEscInter1
		Vetor_Campos(2)="adInteger,4,adParamInput," & vEscInter2
		Vetor_Campos(3)="adInteger,4,adParamInput," & vDistInter1
		Vetor_Campos(4)="adInteger,4,adParamInput," & vDistInter2

		Vetor_Campos(5)="adInteger,4,adParamInput," & vParInter1
		Vetor_Campos(6)="adWChar,20,adParamInput," & vCoodInter1
		Vetor_Campos(7)="adInteger,4,adParamInput," & vParInter2
		Vetor_Campos(8)="adWChar,20,adParamInput," & vCoodInter2

		Vetor_Campos(9)="adInteger,4,adParamInput," & vEscInter1_Novo
		Vetor_Campos(10)="adInteger,4,adParamInput," & vEscInter2_Novo
		Vetor_Campos(11)="adInteger,4,adParamInput," & vDistInter1_Novo
		Vetor_Campos(12)="adInteger,4,adParamInput," & vDistInter2_Novo

		Vetor_Campos(13)="adInteger,4,adParamInput," & vParInter1_Novo
		Vetor_Campos(14)="adWChar,20,adParamInput," & vCoodInter1_Novo
		Vetor_Campos(15)="adInteger,4,adParamInput," & vParInter2_Novo
		Vetor_Campos(16)="adWChar,20,adParamInput," & vCoodInter2_Novo
		Vetor_Campos(17)="adWChar,300,adParamInput," & vFacObs_Novo

		Vetor_Campos(18)="adInteger,4,adParamInput," & vQtdePares

		strSqlRet = APENDA_PARAMSTR("CLA_sp_upd_Interligacao",18,Vetor_Campos)

		Call db.Execute(strSqlRet)


		'Atualiza campos não obrigatórios
		call RetornaDistribuidor (strDistInstala , strDistConfig )
		
		Response.Write "<script laguange=javascript> with (parent.document.forms[0]) {" & _
						"txtParInicial1.value = """ & TratarAspasJS(vParInter1_Novo) & """;" & _
						"txtCoordenada1.value = """ & TratarAspasJS(vCoodInter1_Novo) & """;" & _
						"txtParInicial2.value = """ & TratarAspasJS(vParInter2_Novo) & """;" & _
						"txtCoordenada2.value = """ & TratarAspasJS(vCoodInter2_Novo) & """;" & _
						"cboLocalInstalaInter.value = """ & TratarAspasJS(vEscInter1_Novo) & """;" & _
						"cboLocalConfig.value = """ & TratarAspasJS(vEscInter2_Novo) & """;" & _
						"parent.spnDistLocalInstalaInter.innerHTML = '" & strDistInstala & "';" & _
						"parent.spnDistLocalConfig.innerHTML = '" & strDistConfig & "';" & _
						"txtObsFac.value = """ & TratarAspasJS(vFacObs_Novo) & """;" & _
						"}</script>"

		Response.Write "<script language=javascript>alert('Registro(s) atualizado(s) com sucesso.');</script>"

	End if
End Function

Function RemoverFacilidade(strRede)

	Dim objRSRec1
	Dim objRSRec2
	Dim strInter
	Dim blnProcurar
	Dim Update
	Dim strRet

	if Trim(Request.Form("rdoInter")) = "E" or Trim(Request.Form("rdoInter")) = "" then
		
		vTronco		= Request.Form("txtTronco")
		vPar		= Request.Form("txtPar")
		'Det
		vBastidor	= Request.Form("txtBastidor")
		vRegua		= Request.Form("txtRegua")
		vPosicao	= Request.Form("txtPosicao")
		vTimeSlot	= Request.Form("txtTimeSlot")
		vDominio	= Request.Form("txtDominio")
		vNo			= Request.Form("txtNO")
		vSlot		= Request.Form("txtSlot")
		vPorta		= Request.Form("txtPorta")
		vBastidorInter	= Request.Form("txtBastidorInter")
		vReguaInter		= Request.Form("txtReguaInter")
		vPosicaoInter		= Request.Form("txtPosicaoInter")
		vFila			= Request.Form("txtFila")
		vOTS			= Request.Form("txtOTS")
		vDesigTronco	= Request.Form("txtDesigTronco")
		vLink			= Request.Form("txtLink")
		vIdentProvedor  = Request.Form("txtIDentProvedor")

		vTipoCabo	= Request.Form("cboTipoCabo")
		vLateral	= Request.Form("txtLateral")	
		vCxEmenda	= Request.Form("txtCaixaEmenda")
		vObsFac			= Request.Form("txtObsFac")
		
		'Resgata o Recurso
		strPlataforma  = request("cboPlataforma")
		if strPlataforma  = "" then  strPlataforma = "null"
		
		Set objRSRec = db.execute("CLA_sp_check_recurso2 " & request("cboLocalInstala") & ", " & request("cboDistLocalInstala") & ", " & request("cboProvedor") & ", " & request("cboSistema") & "," & strPlataforma & ",0")
		Set DBAction = objRSRec("ret")

		If DBAction <> 0 then
			Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
			Response.End 
		End if

			vRecurso	= objRSRec("Rec_ID")

			Vetor_Campos(1)="adInteger,2,adParamInput," & vRecurso
			Vetor_Campos(2)="adWChar,20,adParamInput," & vTronco
			Vetor_Campos(3)="adWChar,20,adParamInput," & vPar  
			Vetor_Campos(4)="adWChar,20,adParamInput," & vBastidor  
			Vetor_Campos(5)="adWChar,20,adParamInput," & vRegua  
			Vetor_Campos(6)="adWChar,20,adParamInput," & vPosicao  
			Vetor_Campos(7)="adWChar,20,adParamInput," & vTimeslot  
			Vetor_Campos(8)="adWChar,20,adParamInput," & vDominio  
			Vetor_Campos(9)="adWChar,20,adParamInput," & vNo  
			Vetor_Campos(10)="adWChar,2,adParamInput," & vSlot
			Vetor_Campos(11)="adWChar,2,adParamInput," & vPorta
			Vetor_Campos(12)="adWChar,20,adParamInput," & vBastidorInter  
			Vetor_Campos(13)="adWChar,20,adParamInput," & vReguaInter  
			Vetor_Campos(14)="adWChar,20,adParamInput," & vPosicaoInter  
			Vetor_Campos(15)="adWChar,5,adParamInput," & vFila  
			Vetor_Campos(16)="adWChar,10,adParamInput," & vOTS  
			Vetor_Campos(17)="adWChar,25,adParamInput," & vDesigTronco  
			Vetor_Campos(18)="adWChar,25,adParamInput," & vLink  
			Vetor_Campos(19)="adWChar,20,adParamInput," & vTipoCabo  
			Vetor_Campos(20)="adWChar,20,adParamInput," & vLateral  
			Vetor_Campos(21)="adWChar,20,adParamInput," & vCxEmenda  
			Vetor_Campos(22)="adlongvarchar,300,adParamInput," & vObsFac
			Vetor_Campos(23)="adInteger,2,adParamInput," & strRede
			
			Dim strSqlRet
			strSqlRet = APENDA_PARAMSTR("CLA_sp_del_Facilidade",23,Vetor_Campos)
			Set objRS = db.Execute(strSqlRet)

			Response.Write "<script language=javascript>alert('Registro(s) removido(s) com sucesso.');parent.spnCamposFacRet.innerHTML = ''</script>"

	Else

		vParInter1	= Request.Form("txtParInicial1")
		vCoodInter1 = Request.Form("txtCoordenada1")
		vEscInter1	= Request.Form("cboLocalInstalaInter")
		vDistInter1 = Request.Form("cboDistLocalInstalaInter")

		vParInter2	= Request.Form("txtParInicial2")
		vCoodInter2 = Request.Form("txtCoordenada2")
		vEscInter2	= Request.Form("cboLocalConfig")
		vDistInter2 = Request.Form("cboDistLocalConfig")

		vQtdePares	= Request.Form("txtQtdePares")

		Vetor_Campos(1)="adInteger,2,adParamInput," & vEscInter1
		Vetor_Campos(2)="adInteger,2,adParamInput," & vEscInter2
		Vetor_Campos(3)="adInteger,2,adParamInput," & vDistInter1
		Vetor_Campos(4)="adInteger,2,adParamInput," & vDistInter2
		Vetor_Campos(5)="adInteger,4,adParamInput," & vParInter1
		Vetor_Campos(6)="adWChar,20,adParamInput," & vCoodInter1
		Vetor_Campos(7)="adInteger,4,adParamInput," & vParInter2
		Vetor_Campos(8)="adWChar,20,adParamInput," & vCoodInter2
		Vetor_Campos(9)="adInteger,4,adParamInput," & vQtdePares

		strSqlRet = APENDA_PARAMSTR("CLA_sp_del_Interligacao",9,Vetor_Campos)
		Call db.Execute(strSqlRet)
		Response.Write "<script language=javascript>alert('Registro(s) removido(s) com sucesso.');parent.spnCamposFacRet.innerHTML = ''</script>"

	End if
End Function


Function ConfirmacaoAlteracaoFac(strRede)

	Dim objRSRec1
	Dim objRSRec2
	Dim strInter
	Dim blnProcurar
	Dim Update
	Dim strRet

	if Trim(Request.Form("rdoInter")) = "E" or Trim(Request.Form("rdoInter")) = "" then
		
		vTronco		= Request.Form("txtTronco")
		vPar		= Request.Form("txtPar")
		'Det
		vBastidor	= Request.Form("txtBastidor")
		vRegua		= Request.Form("txtRegua")
		vPosicao	= Request.Form("txtPosicao")
		vTimeSlot	= Request.Form("txtTimeSlot")
		vDominio	= Request.Form("txtDominio")
		vNo			= Request.Form("txtNO")
		vSlot		= Request.Form("txtSlot")
		vPorta		= Request.Form("txtPorta")

		vBastidorInter	= Request.Form("txtBastidorInter")
		vReguaInter		= Request.Form("txtReguaInter")
		vPosicaoInter	= Request.Form("txtPosicaoInter")
		vFila			= Request.Form("txtFila")
		vOTS			= Request.Form("txtOTS")
		vDesigTronco	= Request.Form("txtDesigTronco")
		vLink			= Request.Form("txtLink")

		vTipoCabo		= Request.Form("cboTipoCabo")
		vLateral		= Request.Form("txtLateral")	
		vCxEmenda		= Request.Form("txtCaixaEmenda")

		'NDet /Ade
		vTronco_Novo	= Request.Form("txtTroncoa")
		vPar_Novo		= Request.Form("txtPara")
		'Det
		vBastidor_Novo	= Request.Form("txtBastidora")
		vRegua_Novo		= Request.Form("txtReguaa")
		vPosicao_Novo	= Request.Form("txtPosicaoa")
		vTimeSlot_Novo	= Request.Form("txtTimeSlota")
		vDominio_Novo	= Request.Form("txtDominioa")
		vNo_Novo		= Request.Form("txtNOa")
		vSlot_Novo		= Request.Form("txtSlota")
		vPorta_Novo		= Request.Form("txtPortaa")

		vBastidorInter_Novo	= Request.Form("txtBastidorIntera")
		vReguaInter_Novo	= Request.Form("txtReguaIntera")
		vPosicaoInter_Novo	= Request.Form("txtPosicaoIntera")
		vFila_Novo			= Request.Form("txtFilaa")
		vOTS_Novo			= Request.Form("txtOTSa")
		vDesigTronco_Novo	= Request.Form("txtDesigTroncoa")
		vLink_Novo			= Request.Form("txtLinka")
		vObsFac_Novo		= Request.Form("txtObsFaca")

		vTipoCabo_Novo	= Request.Form("cboTipoCaboa")
		vLateral_Novo	= Request.Form("txtLaterala")
		vCxEmenda_Novo	= Request.Form("txtCaixaEmendaa")
		vObsFacNovo		= Request.Form("txtObsFaca")
		vObsFac			= Request.Form("txtObsFac")

		'Resgata o Recurso
		strPlataforma  = request("cboPlataforma")
		if strPlataforma  = "" then  strPlataforma = "null"
		
		Set objRSRec = db.execute("CLA_sp_check_recurso2 " & request("cboLocalInstala") & ", " & request("cboDistLocalInstala") & ", " & request("cboProvedor") & ", " & request("cboSistema") & "," & strPlataforma & ",0")
		Set DBAction = objRSRec("ret")

		If DBAction <> 0 then
			Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
			Response.End 
		End if

		vRecurso	= objRSRec("Rec_ID")

		Vetor_Campos(1)="adInteger,2,adParamInput," & vRecurso
		Vetor_Campos(2)="adWChar,20,adParamInput," & vTronco
		Vetor_Campos(3)="adWChar,20,adParamInput," & vPar  
		Vetor_Campos(4)="adWChar,20,adParamInput," & vBastidor  
		Vetor_Campos(5)="adWChar,20,adParamInput," & vRegua  
		Vetor_Campos(6)="adWChar,20,adParamInput," & vPosicao  
		Vetor_Campos(7)="adWChar,20,adParamInput," & vTimeslot  
		Vetor_Campos(8)="adWChar,20,adParamInput," & vDominio  
		Vetor_Campos(9)="adWChar,20,adParamInput," & vNo  
		Vetor_Campos(10)="adWChar,2,adParamInput," & vSlot
		Vetor_Campos(11)="adWChar,2,adParamInput," & vPorta
		Vetor_Campos(12)="adWChar,20,adParamInput," & vBastidorInter  
		Vetor_Campos(13)="adWChar,20,adParamInput," & vReguaInter  
		Vetor_Campos(14)="adWChar,20,adParamInput," & vPosicaoInter  
		Vetor_Campos(15)="adWChar,5,adParamInput," & vFila  
		Vetor_Campos(16)="adWChar,10,adParamInput," & vOTS  
		Vetor_Campos(17)="adWChar,25,adParamInput," & vDesigTronco  
		Vetor_Campos(18)="adWChar,25,adParamInput," & vLink  
		Vetor_Campos(19)="adWChar,20,adParamInput," & vTipoCabo  
		Vetor_Campos(20)="adWChar,20,adParamInput," & vLateral  
		Vetor_Campos(21)="adWChar,20,adParamInput," & vCxEmenda  
		Vetor_Campos(22)="adWChar,20,adParamInput," & vTronco_Novo  
		Vetor_Campos(23)="adWChar,20,adParamInput," & vPar_Novo
		Vetor_Campos(24)="adWChar,20,adParamInput," & vBastidor_Novo
		Vetor_Campos(25)="adWChar,20,adParamInput," & vRegua_Novo  
		Vetor_Campos(26)="adWChar,20,adParamInput," & vPosicao_Novo  
		Vetor_Campos(27)="adWChar,20,adParamInput," & vTimeSlot_Novo
		Vetor_Campos(28)="adWChar,20,adParamInput," & vDominio_Novo  
		Vetor_Campos(29)="adWChar,20,adParamInput," & vNo_Novo
		Vetor_Campos(30)="adWChar,2,adParamInput," & vSlot_Novo
		Vetor_Campos(31)="adWChar,2,adParamInput," & vPorta_Novo
		Vetor_Campos(32)="adWChar,20,adParamInput," & vBastidorInter_Novo
		Vetor_Campos(33)="adWChar,20,adParamInput," & vReguaInter_Novo
		Vetor_Campos(34)="adWChar,20,adParamInput," & vPosicaoInter_Novo
		Vetor_Campos(35)="adWChar,5,adParamInput," & vFila_Novo
		Vetor_Campos(36)="adWChar,10,adParamInput," & vOTS_Novo
		Vetor_Campos(37)="adWChar,25,adParamInput," & vDesigTronco_Novo
		Vetor_Campos(38)="adWChar,25,adParamInput," & vLink_Novo
		Vetor_Campos(39)="adWChar,20,adParamInput," & vTipoCabo_Novo
		Vetor_Campos(40)="adWChar,20,adParamInput," & vLateral_Novo
		Vetor_Campos(41)="adWChar,20,adParamInput," & vCxEmenda_Novo
		Vetor_Campos(42)="adlongvarchar,300,adParamInput," & vObsFac
		Vetor_Campos(43)="adlongvarchar,300,adParamInput," & vObsFacNovo
		Vetor_Campos(44)="adInteger,2,adParamInput," & strRede
			
		strSqlRet = APENDA_PARAMSTR("CLA_sp_ValidarFacilidadeUpd",44,Vetor_Campos)
		strXml =  ForXMLAutoQuery(strSqlRet)

		Response.Write "<script language=javascript>var objXmlGeral = new ActiveXObject(""Microsoft.XMLDOM"");objXmlGeral.loadXML('" & strXml &"');parent.JanelaConfirmacao(objXmlGeral);</script>"

	Else
		'PADE/PAC

		vParInter1	= Request.Form("txtParInicial1")
		vCoodInter1 = Request.Form("txtCoordenada1")
		vEscInter1	= Request.Form("cboLocalInstalaInter")
		vDistInter1 = Request.Form("cboDistLocalInstalaInter")

		vParInter2	= Request.Form("txtParInicial2")
		vCoodInter2 = Request.Form("txtCoordenada2")
		vEscInter2	= Request.Form("cboLocalConfig")
		vDistInter2 = Request.Form("cboDistLocalConfig")

		vParInter1_Novo	= Request.Form("txtParInicial1a")
		vCoodInter1_Novo= Request.Form("txtCoordenada1a")
		vEscInter1_Novo	= Request.Form("cboLocalInstalaIntera")
		vDistInter1_Novo= Request.Form("cboDistLocalInstalaIntera")

		vParInter2_Novo	= Request.Form("txtParInicial2a")
		vCoodInter2_Novo= Request.Form("txtCoordenada2a")
		vEscInter2_Novo	= Request.Form("cboLocalConfiga")
		vDistInter2_Novo= Request.Form("cboDistLocalConfiga")

		vQtdePares	= Request.Form("txtQtdePares")

		Vetor_Campos(1)="adInteger,2,adParamInput," & vEscInter1
		Vetor_Campos(2)="adInteger,2,adParamInput," & vEscInter2
		Vetor_Campos(3)="adInteger,2,adParamInput," & vDistInter1
		Vetor_Campos(4)="adInteger,2,adParamInput," & vDistInter2

		
		Vetor_Campos(5)="adInteger,4,adParamInput," & vParInter1
		Vetor_Campos(6)="adWChar,20,adParamInput," & vCoodInter1
		Vetor_Campos(7)="adInteger,4,adParamInput," & vParInter2
		Vetor_Campos(8)="adWChar,20,adParamInput," & vCoodInter2

		Vetor_Campos(9)="adInteger,2,adParamInput," & vEscInter1_Novo
		Vetor_Campos(10)="adInteger,2,adParamInput," & vEscInter2_Novo
		Vetor_Campos(11)="adInteger,2,adParamInput," & vDistInter1_Novo
		Vetor_Campos(12)="adInteger,2,adParamInput," & vDistInter2_Novo
		
		Vetor_Campos(13)="adInteger,4,adParamInput," & vParInter1_Novo
		Vetor_Campos(14)="adWChar,20,adParamInput," & vCoodInter1_Novo
		Vetor_Campos(15)="adInteger,4,adParamInput," & vParInter2_Novo
		Vetor_Campos(16)="adWChar,20,adParamInput," & vCoodInter2_Novo

		Vetor_Campos(17)="adInteger,4,adParamInput," & vQtdePares

		strSqlRet = APENDA_PARAMSTR("CLA_sp_ValidarInterligacaoUpd",17,Vetor_Campos)
		strXml =  ForXMLAutoQuery(strSqlRet)

		Response.Write "<script language=javascript>var objXmlGeral = new ActiveXObject(""Microsoft.XMLDOM"");objXmlGeral.loadXML('" & strXml &"');parent.JanelaConfirmacao(objXmlGeral);</script>"
	End if
End Function

Function ConfirmacaoRemomcaoFac(strRede)

	Dim objRSRec1
	Dim objRSRec2
	Dim strInter
	Dim blnProcurar
	Dim Update
	Dim strRet

	if Trim(Request.Form("rdoInter")) = "E" or Trim(Request.Form("rdoInter")) = "" then
		
		vTronco		= Request.Form("txtTronco")
		vPar		= Request.Form("txtPar")
		'Det
		vBastidor	= Request.Form("txtBastidor")
		vRegua		= Request.Form("txtRegua")
		vPosicao	= Request.Form("txtPosicao")
		vTimeSlot	= Request.Form("txtTimeSlot")
		vDominio	= Request.Form("txtDominio")
		vNo			= Request.Form("txtNO")
		vSlot		= Request.Form("txtSlot")
		vPorta		= Request.Form("txtPorta")

		vBastidorInter	= Request.Form("txtBastidorInter")
		vReguaInter		= Request.Form("txtReguaInter")
		vPosicaoInter	= Request.Form("txtPosicaoInter")
		vFila			= Request.Form("txtFila")
		vOTS			= Request.Form("txtOTS")
		vDesigTronco	= Request.Form("txtDesigTronco")
		vLink			= Request.Form("txtLink")

		vTipoCabo		= Request.Form("cboTipoCabo")
		vLateral		= Request.Form("txtLateral")	
		vCxEmenda		= Request.Form("txtCaixaEmenda")
		vObsFac			= Request.Form("txtObsFac")

		'Resgata o Recurso
		
		strPlataforma  = request("cboPlataforma")
		if strPlataforma  = "" then  strPlataforma = "null"
	
		Set objRSRec = db.execute("CLA_sp_check_recurso2 " & request("cboLocalInstala") & ", " & request("cboDistLocalInstala") & ", " & request("cboProvedor") & ", " & request("cboSistema") & "," & strPlataforma & ",0")
		Set DBAction = objRSRec("ret")

		If DBAction <> 0 then
			Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
			Response.End 
		End if

		vRecurso	= objRSRec("Rec_ID")

		Vetor_Campos(1)="adInteger,2,adParamInput," & vRecurso
		Vetor_Campos(2)="adWChar,20,adParamInput," & vTronco
		Vetor_Campos(3)="adWChar,20,adParamInput," & vPar  
		Vetor_Campos(4)="adWChar,20,adParamInput," & vBastidor  
		Vetor_Campos(5)="adWChar,20,adParamInput," & vRegua  
		Vetor_Campos(6)="adWChar,20,adParamInput," & vPosicao  
		Vetor_Campos(7)="adWChar,20,adParamInput," & vTimeslot  
		Vetor_Campos(8)="adWChar,20,adParamInput," & vDominio  
		Vetor_Campos(9)="adWChar,20,adParamInput," & vNo  
		Vetor_Campos(10)="adWChar,2,adParamInput," & vSlot
		Vetor_Campos(11)="adWChar,2,adParamInput," & vPorta
		Vetor_Campos(12)="adWChar,20,adParamInput," & vBastidorInter  
		Vetor_Campos(13)="adWChar,20,adParamInput," & vReguaInter  
		Vetor_Campos(14)="adWChar,20,adParamInput," & vPosicaoInter  
		Vetor_Campos(15)="adWChar,5,adParamInput," & vFila  
		Vetor_Campos(16)="adWChar,10,adParamInput," & vOTS  
		Vetor_Campos(17)="adWChar,25,adParamInput," & vDesigTronco  
		Vetor_Campos(18)="adWChar,25,adParamInput," & vLink  
		Vetor_Campos(19)="adWChar,20,adParamInput," & vTipoCabo  
		Vetor_Campos(20)="adWChar,20,adParamInput," & vLateral  
		Vetor_Campos(21)="adWChar,20,adParamInput," & vCxEmenda  
		Vetor_Campos(22)="adlongvarchar,300,adParamInput," & vObsFac
		Vetor_Campos(23)="adInteger,2,adParamInput," & strRede
			
		strSqlRet = APENDA_PARAMSTR("CLA_sp_ValidarFacilidadeRem",23,Vetor_Campos)
		
		strXml =  ForXMLAutoQuery(strSqlRet)

		Response.Write "<script language=javascript>var objXmlGeral = new ActiveXObject(""Microsoft.XMLDOM"");objXmlGeral.loadXML('" & strXml &"');parent.JanelaConfirmacaoRem(objXmlGeral);</script>"

	Else
		'PADE/PAC

		vParInter1	= Request.Form("txtParInicial1")
		vCoodInter1 = Request.Form("txtCoordenada1")
		vEscInter1	= Request.Form("cboLocalInstalaInter")
		vDistInter1 = Request.Form("cboDistLocalInstalaInter")

		vParInter2	= Request.Form("txtParInicial2")
		vCoodInter2 = Request.Form("txtCoordenada2")
		vEscInter2	= Request.Form("cboLocalConfig")
		vDistInter2 = Request.Form("cboDistLocalConfig")

		vQtdePares	= Request.Form("txtQtdePares")

		Vetor_Campos(1)="adInteger,2,adParamInput," & vEscInter1
		Vetor_Campos(2)="adInteger,2,adParamInput," & vEscInter2
		Vetor_Campos(3)="adInteger,2,adParamInput," & vDistInter1 
		Vetor_Campos(4)="adInteger,2,adParamInput," & vDistInter2
		Vetor_Campos(5)="adInteger,4,adParamInput," & vParInter1
		Vetor_Campos(6)="adWChar,20,adParamInput," & vCoodInter1
		Vetor_Campos(7)="adInteger,4,adParamInput," & vParInter2
		Vetor_Campos(8)="adWChar,20,adParamInput," & vCoodInter2

		Vetor_Campos(9)="adInteger,4,adParamInput," & vQtdePares

		strSqlRet = APENDA_PARAMSTR("CLA_sp_ValidarInterligacaoRem",9,Vetor_Campos)
		strXml =  ForXMLAutoQuery(strSqlRet)

		Response.Write "<script language=javascript>var objXmlGeral = new ActiveXObject(""Microsoft.XMLDOM"");objXmlGeral.loadXML('" & strXml &"');parent.JanelaConfirmacaoRem(objXmlGeral);</script>"

	End if
End Function

Function ListaFacilidades(intRede,intWidthTable,blnLink,rs)

	Dim blnDet
	Dim intCount

	if not rs.Eof and Not rs.Bof then

		Select Case intRede

			Case 1	'DET

				strRet = strRet  & "<table border=0 cellspacing=1 cellpadding=0 width=" & intWidthTable & ">"
				strRet = strRet  & "<tr>"
				strRet = strRet  & "<th>&nbsp;Bastidor</th>"
				strRet = strRet  & "<th>&nbsp;Régua</th>"
				strRet = strRet  & "<th>&nbsp;Posicao</th>"
				strRet = strRet  & "<th>&nbsp;TimeSlot</th>"
				strRet = strRet  & "<th>&nbsp;Domínio</th>"
				strRet = strRet  & "<th>&nbsp;Nó</th>"
				strRet = strRet  & "</tr>"

				blnDet = false	
				intCount = 1
				While Not rs.eof
						
					if not isNull(rs("Fac_Bastidor")) then
						if (intCount mod 2) = 0 then strClass = "clsSilver"	else strClass = "clsSilver2" end if
						valores = "\'" & rs("Fac_Bastidor") & "\'"
						valores = valores & ",\'" & rs("Fac_Regua") & "\'"
						valores = valores & ",\'" & rs("Fac_Posicao") & "\'"
						valores = valores & ",\'" & rs("Fac_Timeslot") & "\'"
						valores = valores & ",\'" & rs("Fac_Dominio") & "\'"
						valores = valores & ",\'" & rs("Fac_No") & "\'"

						strRet = strRet  & "<tr class=" & strClass & ">"
						if blnLink then
							strRet = strRet  & "<td>&nbsp;<a href=\'#\' onClick=""javascript:copyPaste(" & valores & ")"">" & rs("Fac_Bastidor") & "</a></td>"
						Else
							strRet = strRet  & "<td>&nbsp;" & rs("Fac_Bastidor") & "</td>"
						End if	
						strRet = strRet  & "<td>&nbsp;" & rs("Fac_Regua") & "</td>"
						strRet = strRet  & "<td>&nbsp;" & rs("Fac_Posicao") & "</td>"
						strRet = strRet  & "<td>&nbsp;" & rs("Fac_Timeslot") & "</td>"
						strRet = strRet  & "<td>&nbsp;" & rs("Fac_Dominio") & "</td>"
						strRet = strRet  & "<td>&nbsp;" & rs("Fac_No") & "</td>"
						strRet = strRet  & "</tr>"
						blnDet = true	
					End if	

					rs.movenext
					intCount = intCount + 1

				Wend

				strRet = strRet  & "</table>"
				if blnDet then
					Response.Write "<script language=javascript>parent.spnConsulta.innerHTML = '" & strRet &"'</script>"
				Else
					Response.Write "<script language=javascript>alert('FACILIDADE NÃO ENCONTRADA!')</script>"
				End if	
				
			Case 2	'NDET

				strRet = strRet  & "<table border=0 cellspacing=1 cellpadding=0 width=" & intWidthTable & ">"
				strRet = strRet  & "<tr>"
				strRet = strRet  & "<th>&nbsp;Tronco</th>"
				strRet = strRet  & "<th>&nbsp;Par</th>"
				strRet = strRet  & "</tr>"

				blnNDet = false
				intCount = 1
				While Not rs.eof

					if not isNull(rs("Fac_Tronco")) then
						if (intCount mod 2) = 0 then strClass = "clsSilver"	else strClass = "clsSilver2" end if
						valores = "\'" & rs("Fac_Tronco") & "\'"
						valores = valores & ",\'" & rs("Fac_Par") & "\'"
						strRet = strRet  & "<tr class=" & strClass & ">"
						if blnLink then
							strRet = strRet  & "<td>&nbsp;<a href=\'#\' onClick=""javascript:copyPaste(" & valores & ");"">" & rs("Fac_Tronco") & "</a></td>"
						Else	
							strRet = strRet  & "<td>&nbsp;" & rs("Fac_Tronco") & "</td>"
						End if	
						strRet = strRet  & "<td>&nbsp;" & rs("Fac_Par") & "</td>"
						strRet = strRet  & "</tr>"
						blnNDet = true
					End if	

					rs.MoveNext
					intCount = intCount + 1
				Wend	

				strRet = strRet  & "</table>"
				if blnNDet then
					Response.Write "<script language=javascript>parent.spnConsulta.innerHTML = '" & strRet &"'</script>"
				Else
					Response.Write "<script language=javascript>alert('FACILIDADE NÃO ENCONTRADA!')</script>"
				End if	

			Case 3 'ADE

				strRet = strRet  & "<table border=0 cellspacing=1 cellpadding=0 width=" & intWidthTable & ">"
				strRet = strRet  & "<tr>"
				strRet = strRet  & "<th>&nbsp;Cabo</th>"
				strRet = strRet  & "<th>&nbsp;Par</th>"
				strRet = strRet  & "<th>&nbsp;Derivação</th>"
				strRet = strRet  & "<th>&nbsp;Tipo Cabo</th>"
				strRet = strRet  & "<th>&nbsp;PADE</th>"
				strRet = strRet  & "</tr>"

				blnAde = false
				intCount = 1
				While Not rs.eof
						
					if not isNull(rs("Fac_Tronco"))then
						if (intCount mod 2) = 0 then strClass = "clsSilver"	else strClass = "clsSilver2" end if
						valores = ""
						valores = valores & "\'" & rs("Fac_Tronco") & "\'"
						valores = valores & ",\'" & rs("Fac_Par") & "\'"
						valores = valores & ",\'" & rs("Fac_Lateral") & "\'"
						valores = valores & ",\'" & rs("Fac_TipoCabo") & "\'"
						valores = valores & ",\'" & rs("Fac_CxEmenda") & "\'"

						strRet = strRet  & "<tr class=" & strClass & ">"
						if blnLink then
							strRet = strRet  & "<td>&nbsp;<a href=\'#\' onClick=""javascript:copyPaste(" & valores & ")"">" & rs("Fac_Tronco") & "</a></td>"
						Else	
							strRet = strRet  & "<td>&nbsp;" & rs("Fac_Tronco") & "</td>"
						End if	
						strRet = strRet  & "<td>&nbsp;" & rs("Fac_Par") & "</td>"
						strRet = strRet  & "<td>&nbsp;" & rs("Fac_Lateral") & "</td>"
						strRet = strRet  & "<td>&nbsp;" & rs("Fac_TipoCabo") & "</td>"
						strRet = strRet  & "<td>&nbsp;" & rs("Fac_CxEmenda") & "</td>"
						strRet = strRet  & "</tr>"
						blnAde = true
					End if	
					rs.movenext
					intCount = intCount + 1
				Wend
					
				strRet = strRet  & "</table>"
				if blnAde then
					Response.Write "<script language=javascript>parent.spnConsulta.innerHTML = '" & strRet &"'</script>"
				Else	
					Response.Write "<script language=javascript>alert('FACILIDADE NÃO ENCONTRADA!')</script>"
				End if	

		End Select
	Else
		Response.Write "<script language=javascript>alert('FACILIDADE NÃO ENCONTRADA!')</script>"
	End if

End Function


Select Case Request.Form("hdnAcao")
	
	Case "ResgatarCamposFacilidade"
		Response.Write "AAA"
		Response.Write "<script language=javascript>parent.spnCamposFac.innerHTML = '" & ResgatarCamposFacilidade(Request.Form("cboSistema")) & "'</script>"
		Response.Write "BBB"
		Response.Write "<script language=javascript>parent.spnCamposFacRet.innerHTML = ''</script>"

	Case "ProcurarFacilidade"
		Response.Write "<script language=javascript>parent.spnCamposFacRet.innerHTML = '" & ProcurarFacilidade(Request.Form("cboSistema")) & "'</script>"
		
	Case "AlterarFacilidade"	
		Call AlterarFacilidade(Request.Form("cboSistema"))

	Case "ConfirmacaoAlteracaoFac"	
		Call ConfirmacaoAlteracaoFac(Request.Form("cboSistema"))

	Case "ConfirmacaoRemomcaoFac"	
		ConfirmacaoRemomcaoFac(Request.Form("cboSistema"))

	Case "RemoverFacilidade"	
		RemoverFacilidade(Request.Form("cboSistema"))
	
End Select
%>
