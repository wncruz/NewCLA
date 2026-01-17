
<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr>
		<td>
			<iframe	id			= "IFrmMotivoPend"
				    name        = "IFrmMotivoPend"
				    width       = "100%"
				    height      = "120px"
				    src			= "../inc/MotivoPendencia.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>"
				    frameborder = "0"
				    scrolling   = "no"
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>
<% if strEild = "S"  then%>
<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr>
		<td>
			<iframe	id			= "IFrmListaStatusSNOA"
				    name        = "IFrmListaStatusSNOA"
				    width       = "100%"
				    height      = "160px"
				    src			= "../inc/ListaStatusSNOA.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>&telaaceitar=0"
				    frameborder = "0"
				    scrolling   = "no"
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>
<%end if %>

<!-- CH-56418KBG - Inicio -->
<%

	//dblSolId = 1535099
	Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_designacaoFornecedora ",1,Vetor_Campos)
	Set objRS_DSF = db.Execute(strSqlRet)

	if not objRS_DSF.eof then %>

		<table border=0 cellspacing="1" cellpadding="0" width="760">
			<tr>
				<td>
					<iframe	id			= "IFrmListaDesigFornecSNOA"
						    name        = "IFrmListaDesigFornecSNOA"
						    width       = "100%"
						    height      = "60px"
						    src			= "../inc/ListaDesigFornecSNOA.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>&telaaceitar=<%=telaaceitar%>"
						    frameborder = "0"
						    scrolling   = "no"
						    align       = "left">
					</iFrame>
				</td>
			</tr>
		</table>

	<% end if %>	
<!-- CH-56418KBG - Fim -->


<script>
	function Trim(str){return str.replace(/^\s+|\s+$/g,"");}

	function GravarSwitch ()
	{
		with (document.forms[0])
		{
			if (txtSwitch.value==null || Trim(txtSwitch.value)==""){
				alert("Favor informar o Switch Concentrador");
				return;
			}
			if (txtPortaSwitch.value==null || Trim(txtPortaSwitch.value)==""){
				alert("Favor informar a Porta do Switch");
				return;
			}
			
						
			target = "IFrmProcesso"
			action = "AlocarSwitch.asp"
			submit()
		}
	}
	
	function GravarEthernet ()
	{
		with (document.forms[0])
		{
			if (txtNroAcessoEbtEthernet.value==null || Trim(txtNroAcessoEbtEthernet.value)==""){
				alert("Favor informar a Designação do Acesso");
				return;
			}
					
						
			target = "IFrmProcesso"
			action = "AlocarEthernet.asp"
			submit()
		}
	}
</script>

<% if strTecnologia = "FO EDD" then %>
		<table rules="groups"  border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
		<tr>
			<th colspan=10 class="clsSilver">&nbsp;•&nbsp; Controle do Switch</th>
		</tr>
		
    	<tr class="clsSilver">
			<td>&nbsp;&nbsp;Switch Concentrador</td><td>&nbsp;&nbsp; <input type="text" class="text" name="txtSwitch" value="<%=strSwitch%>" size="25" maxlength="20" ></td><td>&nbsp;&nbsp;</td>
		</tr>
		<tr class="clsSilver">
			<td>&nbsp;&nbsp;Porta do Switch</td><td>&nbsp;&nbsp; <input type="text" class="text" name="txtPortaSwitch" value="<%=strSwitchPorta%>" size="25" maxlength="10"></td><td>&nbsp;&nbsp;</td>
		</tr>
	</table>
	<table width="760" border=0>
		<tr>
			<td>
				<table width=50% border=0 align=center cellspacing=1 cellpadding=1>
					<tr class=clsSilver2>
						<td colspan=4 align=center>
							<input type="button" class="button" name="btnOK" style="width:150px;height:22px" value="Alocar" onclick="return GravarSwitch()" accesskey="I" onmouseover="showtip(this,event,'Alocar Facilidade(s)(Alt+I)');">&nbsp;
							<input type="button" class="button" name="btnSair" value="Sair" style="width:150px;height:22px"  onClick="sair()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
						</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		
<% end if %>		
	
<% if strEild <>"S"  then%>
<% '@@Davif  %>
<table border=0 cellspacing="0" cellpadding="0" width="760">
	<tr>
		<td>
			<iframe	id			= "IFrmEntregaProv"
				    name        = "IFrmEntregaProv"
				    width       = "100%"
				    height      = "83"
				    src			= "../inc/PrevisaoProvedor.asp?dblAcfId=<%=DblAcf_ID%>&dblSolId=<%=dblSolId%>&dblEild=<%=strEild%>&dblPonta=<%=strPonta%>"
					frameborder = "0"
				    scrolling   = "no"
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>

<table border=0 cellspacing="0" cellpadding="0" width="760" height="190px">
	<tr>
		<td>
			<iframe	id			= "IFrmDadosTecnicoOperadora"
				    name        = "IFrmDadosTecnicoOperadora"
				    width       = "100%"
				    height      = "100%"
				    src			= "../inc/DadosTecnicoOperadora.asp?dblAcfId=<%=DblAcf_ID%>&dblSolId=<%=dblSolId%>&dblEild=<%=strEild%>&dblPonta=<%=strPonta%>"
					frameborder = "0"
				    scrolling   = "no"
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>
<%end if %>
<% if strTecnologia <> "FO EDD" then %>
<%'@@Davif  %>
<table rules="groups"  border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
	<tr>
		<th colspan=10>&nbsp;•&nbsp;Recurso</th>
	</tr>

<!--JCARTUS-->
	<tr class="clsSilver">
		<td width="170px" nowrap><font class="clsObrig">:: </font>Local de Entrega</td>
		<td colspan="8">
	    	<%
			set objRS = db.execute("CLA_sp_sel_estacao " & Trim(strLocalInstala))
			%> 
			<input type="Hidden" name="cboLocalInstala" value="<%=strLocalInstala%>">
		  <input type="text" class="text" name="txtCNLLocalEntrega" value=<%=objRS("Cid_Sigla")%> maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		  &nbsp;<input type="text" class="text" name="txtComplLocalEntrega" value=<%=objRS("Esc_Sigla")%> maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsuFac(document.Form2.txtCNLLocalEntrega,document.Form2.txtComplLocalEntrega,<%=dblUsuId%>,1);" TIPO="A">
		</td> 
		<td colspan="1">&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px" nowrap><font class="clsObrig">:: </font>Local de Configuração</td>
		<td colspan="8">
			<%
			set objRS = db.execute("CLA_sp_sel_estacao " & Trim(strLocalConfig))
			%>
			<input type="Hidden" name="cboLocalConfig" value="<%=strLocalConfig%>">
		  <input type="text" class="text" name="txtCNLLocalConfig" value=<%=objRS("Cid_Sigla")%> maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		  &nbsp;<input type="text" class="text" name="txtComplLocalConfig" value=<%=objRS("Esc_Sigla")%> maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsuFac(document.Form2.txtCNLLocalConfig,document.Form2.txtComplLocalConfig,<%=dblUsuId%>,2);" TIPO="A">
		</td>
		<td colspan="1">&nbsp;</td>
	</tr>
<!--JCARTUS-->

	<tr class=clsSilver>
		<td width=170><font class="clsObrig">:: </font>Distribuidor</td>
		<td colspan="9">
			<span id="spnDistLocalInstala">
			<select name="cboDistLocalInstala" style="width:200px">
			<option value=""></option>
			<%	if Trim(strLocalInstala) <> "" then
					set objRS = db.execute("CLA_sp_view_recursodistribuicao " & strLocalInstala )
					While not objRS.Eof
						strItemSel = ""
						if Trim(strDistrib) = Trim(objRS("Dst_ID")) then strItemSel = " Selected " End if
						Response.Write "<Option value=" & objRS("Dst_ID") & strItemSel & ">" & objRS("Dst_Desc") & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				End if
				%>
			</select>
			</span>
		</td>
	</tr>

	<tr class=clsSilver>
		<td width=170><font class="clsObrig">:: </font>Rede</td>
		<td colspan="9" >
			<select name="cboRede" onChange="RetornaCboPlataforma(cboRede.value,<%=strPla%>)" onBlur="ControleRede(this);"  >
				<option value=""></option>
				<%
					Set objRS = db.execute("CLA_sp_sel_sistema 0")
					While not objRS.Eof
						strItemSel = ""
						if Trim(strRede) = Trim(objRS("Sis_ID")) then strItemSel = " Selected " End if
						Response.Write "<Option value='" & objRS("Sis_ID")& "'" & strItemSel & ">" & Ucase(objRS("Sis_Desc")) & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""

				%>
			</select>
			<span ID = spnPlataforma>
			</span>
		</td>
	</tr>

	<tr class=clsSilver>
		<td width=170 ><font class="clsObrig">:: </font>Provedor</td>
		<td colspan="9" >
			<%
			set objRS = db.execute("CLA_sp_sel_provedor " & Trim(strProId))
			%>
			<select name="cboProvedor" style="width:250px" readonly>
				<%Response.Write "<Option value='" & Trim(objRS("Pro_ID")) & "' tag_provedor=" & strCartaProv & strItemSel & ">" & objRS("Pro_Nome") & "</Option>"%>
			</select>
		</td>
	</tr>
	<tr class=clsSilver>
		<td width=170><font class="clsObrig">:: </font>Prazos de Contratação de Acesso</td>
		<td colspan="9" >
			<span id=spnRegimeCntr>
				<select name="cboRegimeCntr">
					<option value=""></option>
					<%
						if Trim(strProID) <> "" then
							set objRS = db.execute("CLA_sp_sel_regimecontrato 0," & strProID)
							While not objRS.Eof
								strItemSel = ""
								if Trim(strRegId) = Trim(objRS("Reg_ID")) then strItemSel = " Selected " End if
								Response.Write "<Option value='" & Trim(objRS("Reg_ID")) & "'" & strItemSel & ">" & LimparStr(Trim(objRS("Pro_Nome"))) & " - " & LimparStr(Trim(objRS("Tct_Desc"))) & "</Option>"
								objRS.MoveNext
							Wend
							strItemSel = ""
						End if
					%>
				</select>
			</span>
		</td>
	</tr>
	<tr class=clsSilver>
		<td width=170>&nbsp;&nbsp;&nbsp;Promoção</td>
		<td colspan="9" >
			<span id=spnPromocao>
			<select name="cboPromocao" style="width:200px">
				<option value=""></option>
				<%
					if Trim(strProId) <> "" then
						set objRS = db.execute("CLA_sp_sel_promocaoprovedor 0," & strProId)
						While not objRS.Eof
							strItemSel = ""
							if Trim(strPrmId) = Trim(objRS("Prm_ID")) then strItemSel = " Selected " End if
							Response.Write "<Option value='" & Trim(objRS("Prm_ID")) & "'" & strItemSel & ">" & objRS("Prm_Desc") & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					End if
				%>
			</select>
			</span>
		</td>
	</tr>
	<tr class=clsSilver>
		<td width=170>&nbsp;&nbsp;&nbsp;Status</td>
		<td colspan="9">
			<input type=text size=25 readonly style='BACKGROUND-COLOR:#eeeeee' class=text name="txtStatus" tabIndex=-1 VALUE="<%=strStatus%>" onClick="AbrirInterligacoes()" style="cursor:hand" onmouseover="showtip(this,event,'Interligações Alocadas');">
		</td>
	</tr>
	<%
	if intTipoProcesso = 3 and intTipoAcao = 3 then

			Set objDic = Server.CreateObject("Scripting.Dictionary")

			Vetor_Campos(1)="adInteger,2,adParamInput," & dblPedId
			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_HistoricoFac",1,Vetor_Campos)
			'Response.Write strSqlRet
			Set objRS = db.Execute(strSqlRet)
			objRS.Close
			objRS.CursorLocation = adUseClient
			objRS.Open

			strXmlFacOrigem = FormatarXml(MontaXmlFacOrigem(objRS))

			if not objRS.Eof and not objRS.Bof then
				'objRS.MoveFirst

				strHtml = "<tr class=clsSilver>"
				strHtml = strHtml &  "<td colspan=2>"
				strHtml = "<table cellpadding=0 cellspacing=1 width=760>"

				dblSisIdFacOrig = objRS("Sis_ID")

				Select Case dblSisIdFacOrig

					Case 1
						strHtml = strHtml &  "<tr>"
						strHtml = strHtml &  "<th colspan=7>&nbsp;•&nbsp;Facilidades de Origem</th>"
						strHtml = strHtml &  "</tr>"

						strHtml = strHtml &  "<tr>"
						strHtml = strHtml &  "<th nowrap >&nbsp;Nº Acesso Pta EBT</th>"
						strHtml = strHtml &  "<th nowrap >&nbsp;Fila</th>"
						strHtml = strHtml &  "<th >&nbsp;Bastidor</th>"
						strHtml = strHtml &  "<th >&nbsp;Régua</th>"
						strHtml = strHtml &  "<th >&nbsp;Posição</th>"
						strHtml = strHtml &  "<th >&nbsp;Timeslot</th>"
						strHtml = strHtml &  "<th></th>"
						strHtml = strHtml &  "</tr>"
					Case 2
						strHtml = strHtml &  "<tr>"
						strHtml = strHtml &  "<th colspan=4>&nbsp;•&nbsp;Facilidades de Origem</th>"
						strHtml = strHtml &  "</tr>"

						strHtml = strHtml &  "<tr>"
						strHtml = strHtml &  "<th nowrap >&nbsp;Nº Acesso Pta EBT</th>"
						strHtml = strHtml &  "<th >&nbsp;Tronco</th>"
						strHtml = strHtml &  "<th >&nbsp;Par</th>"
						strHtml = strHtml &  "<th></th>"
						strHtml = strHtml &  "</tr>"
					Case 3
						strHtml = strHtml &  "<tr>"
						strHtml = strHtml &  "<th colspan=7>&nbsp;•&nbsp;Facilidades de Origem</th>"
						strHtml = strHtml &  "</tr>"

						strHtml = strHtml &  "<tr>"
						strHtml = strHtml &  "<th nowrap >&nbsp;Nº Acesso Pta EBT</th>"
						strHtml = strHtml &  "<th >&nbsp;Cabo</th>"
						strHtml = strHtml &  "<th >&nbsp;Par</th>"
						strHtml = strHtml &  "<th >&nbsp;PADE</th>"
						strHtml = strHtml &  "<th >&nbsp;Derivação</th>"
						strHtml = strHtml &  "<th >&nbsp;Tipo Cabo</th>"
						strHtml = strHtml &  "<th></th>"
						strHtml = strHtml &  "</tr>"

				End Select

				blnCor = true
				While Not objRS.Eof
					Select Case dblSisIdFacOrig
						Case 1
							strRepresentacao = objRS("Fac_TimeSlot")
						Case Else
							strRepresentacao = objRS("Fac_Par")

					End Select

						str_Acf_NroAcessoPtaEbt = objRS("Acf_NroAcessoPtaEbt")
					if not isNull(objRS("Fac_Representacao")) then
						strRepresentacao = objRS("Fac_Representacao")
					End if

					Select Case dblSisIdFacOrig
						Case 1
							if  Not objDic.Exists(strRepresentacao) then
								Call objDic.Add(strRepresentacao,objRS("Fac_Id"))
								if blnCor then
									strHtml = strHtml &  "<tr class=clsSilver >"
									blnCor = false
								Else
									strHtml = strHtml &  "<tr class=clsSilver2>"
									blnCor = true
								End if
								strHtml = strHtml &  "<td nowrap >&nbsp;"	& objRS("Acf_NroAcessoPtaEbt")	&"</td>"
								strHtml = strHtml &  "<td nowrap >&nbsp;"	& objRS("Fac_Fila")	&"</td>"
								strHtml = strHtml &  "<td nowrap >&nbsp;"	& objRS("Fac_Bastidor")	&"</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;"	& objRS("Fac_Regua")		& "</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;"	& objRS("Fac_Posicao")	& "</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;"	& strRepresentacao	& "</td>"
								strHtml = strHtml &  "<td >&nbsp;<input type=button class=button name=btnLiberarFac value=Alocar onClick='LiberarFacilidade(" & objRS("Fac_Id") & ")'></td>"
								strHtml = strHtml &  "</tr>"
							End if
						Case 2
							if  Not objDic.Exists(strRepresentacao) then
								Call objDic.Add(strRepresentacao,objRS("Fac_Id"))
								if blnCor then
									strHtml = strHtml &  "<tr class=clsSilver >"
									blnCor = false
								Else
									strHtml = strHtml &  "<tr class=clsSilver2>"
									blnCor = true
								End if
								strHtml = strHtml &  "<td nowrap >&nbsp;"	& objRS("Acf_NroAcessoPtaEbt")	&"</td>"
								strHtml = strHtml &  "<td nowrap >&nbsp;"& objRS("Fac_Tronco")&"</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;" & strRepresentacao	& "</td>"
								strHtml = strHtml &  "<td >&nbsp;<input type=button class=button name=btnLiberarFac value=Alocar onClick='LiberarFacilidade(" & objRS("Fac_Id") & ")'></td>"
								strHtml = strHtml &  "</tr>"
							End if
						Case 3
							if  Not objDic.Exists(strRepresentacao) then
								Call objDic.Add(strRepresentacao,objRS("Fac_Id"))
								if blnCor then
									strHtml = strHtml &  "<tr class=clsSilver >"
									blnCor = false
								Else
									strHtml = strHtml &  "<tr class=clsSilver2>"
									blnCor = true
								End if
								strHtml = strHtml &  "<td nowrap >&nbsp;"& objRS("Acf_NroAcessoPtaEbt")	&"</td>"
								strHtml = strHtml &  "<td nowrap >&nbsp;"& objRS("Fac_Tronco")	& "</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;" & objRS("Fac_Par")		& "</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;" & objRS("Fac_Lateral")	& "</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;" & objRS("Fac_CxEmenda")	& "</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;" & strRepresentacao	& "</td>"
								strHtml = strHtml &  "<td >&nbsp;<input type=button class=button name=btnLiberarFac value=Alocar onClick='LiberarFacilidade(" & objRS("Fac_Id") & ")'></td>"
								strHtml = strHtml &  "</tr>"
							End if
					End Select
					objRS.MoveNext
				Wend
				strHtml = strHtml &  "</table></td></tr>"
				Response.Write strHtml
			End if
	End if%>
	<tr class=clsSilver>
		<td colspan="10">
			<div id=divIDFis1 style="DISPLAY: 'none'">
				<table width=100%>
					<tr>
						<td >
							<iframe	id			= "IFrmIDFis1"
								    name        = "IFrmIDFis1"
								    width       = "100%"
								    height      = "45px"
								    frameborder = "0"
								    scrolling   = "auto"
								    align       = "left">
							</iFrame>
						</td>
					</tr>
				</table>
			</div>
		</td>
	</tr>
</table>
<table cellspacing=0 cellpadding=0 width=760 border=0>
	<tr><th colspan=6>&nbsp;•&nbsp;Facilidades do Pedido</th></tr>
</table>
<span id=spnCampos></span>

<span id=spnCamposEthernetInterconexao></span>
<span id=spnEthernetInterconexao></span>
<span id=spnCamposEthernetMetro></span>
<span id=spnEthernetMetro></span>
<span id=spnCamposEthernetEDD></span>
<span id=spnEthernetEDD></span>
<span id=spnCamposEthernetPE></span>
<span id=spnEthernetPE></span>
<span id=spnCamposEthernetVlan></span>

<table cellspacing=1 cellpadding=0 width=760 border=0>
	<tr class=clsSilver>
		<td >
			<iframe id=IFrmFacilidade
					name=IFrmFacilidade
					align=left
					src="ListaFacilidades.asp"
					frameBorder=0
					width="100%"
					BORDER=0
					height=45>
			</iframe>
		</td>
	</tr>
</table>

<span id=spnDet></span>
<span id=spnNDet></span>
<span id=spnAde></span>
<span id=spnBsodNet></span>
<span id=spnBsodVia></span>
<span id=spnBsod></span>
<span id=spnBsodLight></span>
<span id=spnFoEtherNet></span>
<span id=spnSwitchRadioIP></span>


<table border=0 cellspacing="0" cellpadding="0" width=760 >
	<tr>
		<td >
			<iframe	id			= "IFrmProcesso1"
				    name        = "IFrmProcesso1"
				    width       = "760"
				    height      = "18px"
				    frameborder = "0"
				    scrolling   = "no"
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>



<table width="760" border=0>
	<tr><td>
	<table width=50% border=0 align=center cellspacing=1 cellpadding=1>
		<tr class=clsSilver2>
			<%if cint("0" & intTipoProcesso) = 4 then%>
				<td colspan=4 align=center><input type="button" class="button" name="btnOK" style="width:150px;height:22px" value="Concluir Cancelamento" onclick="return GravarFacilidade()" accesskey="I" onmouseover="showtip(this,event,'Concluir cancelamento (Alt+I)');"></td>
			<%Else%>
				<%if (  intTipoProcesso = 3 and intTipoAcao = 3 ) then%>
					<td colspan=4 align=center>
						<input type="button" class="button" name="btnOK" style="width:150px;height:22px" value="Alocar Facilidade(s)" onclick="return GravarFacilidade()" accesskey="I" onmouseover="showtip(this,event,'Alocar Facilidade(s)(Alt+I)');">&nbsp;
						<% if ( str_Acf_NroAcessoPtaEbt <> "" ) then%>
						<input type="button" class="button" name="btnOK" style="width:150px;height:22px" value="Liberar para Serviço" onclick="return liberarServico()" accesskey="U" onmouseover="showtip(this,event,'Liberar para Serviço(Alt+U)');">
						<% end if %>
					</td>
				<%elseif (  intTipoProcesso = 2 ) then%>
					<td colspan=4 align=center>
						<input type="button" class="button" name="btnOK" style="width:150px;height:22px" value="Liberar para Serviço" onclick="return liberarServico()" accesskey="U" onmouseover="showtip(this,event,'Liberar para Serviço(Alt+U)');">
					</td>
				<%Else%>
					<td colspan=4 align=center><input type="button" class="button" name="btnOK" style="width:150px;height:22px" value="Alocar Facilidade(s)" onclick="return GravarFacilidade()" accesskey="I" onmouseover="showtip(this,event,'Alocar Facilidade(s)(Alt+I)');"></td>
				<%End if%>
			<%End if%>
		</tr>
		
		<tr>
			<td><input type="button" class="button" style="width:150px" name="consultar" value="Posições Disponíveis" onClick="objAryFac[0][0]='';PosicoesLivre('L')" accesskey="L" onmouseover="showtip(this,event,'Consultar posições disponíveis (Alt+L)');" 
			
			<% 
				if ( strRede = 1 or strRede = "" or strRede = 4 or strRede = 5 or  strRede = 10 or  strRede = 11 or strRede = 13) then
				Response.Write " disabled=true "
			end if 
			%>
			></td>
			<td><input type=button name=btnConsEstoque style="width:150px" class=button value="Consultar Estoque" onClick="ProcurarIDFis(1)" <%if blnGravado or Trim(Ucase(strTecSigla)) = "ADE" then Response.Write " disabled " End if %> accesskey="S" onmouseover="showtip(this,event,'Consultar estoque (Alt+S)');"></td>
			
			
			<td><input type="button" class="button"
				name="btnEmailPro" style="width:150px"
				value="Enviar e-mail para provedor"
				onclick="EnviarEmail(0,document.forms[0].cboProvedor(document.forms[0].cboProvedor.selectedIndex).tag_provedor)"
				accesskey="M" onmouseover="showtip(this,event,'Enviar email para provedor(Alt+M)');"  
				<% 'response.write "<script>alert('"&strEild&"')</script>"
					if ( ( orisolid = 3 ) or ( strEild  = "S" ) ) then 
						Response.Write " disabled "
					end if %>
>
			</td>
			<td><input type="button" class="button" name="btnRedeDet" style="width:150px" value="Rede Determinística" onclick="JanelaRedeDet()" accesskey="Z" onmouseover="showtip(this,event,'Rede Determinística(Alt+Z)');"
				
				
				<%	
					 if intTipoProcesso = 4 or strRede <> 1 or strRede = "" then 
					Response.Write " disabled=true "
				end if %>>
			</td>
		</tr>
		<tr>
			<td><input type="button" class="button" style="width:150px" name="ocupados" value="Posições Ocupadas" onClick="PosicoesLivre('O')" accesskey="O" onmouseover="showtip(this,event,'Consultar posições ocupadas (Alt+O)');"
			<%	
				if ( strRede = 1 or strRede = "" or strRede = 4 or strRede = 5 or  strRede = 10 or  strRede = 11 or strRede = 13)  then
					Response.Write " disabled=true "
				end if 
			%>
			></td>
			<td><input type=button name=btnConsIdFis class=button style="width:150px" value="Consultar ID's Físicos" onClick="ProcurarIDFis(2)" <%if blnGravado then Response.Write " disabled " End if %> accesskey="J" onmouseover="showtip(this,event,'Consultar ID\'s Físicos (Alt+J)');"></td>
			<td><input type="button" class="button" name="btnLimparEstoque" style="width:150px" value="Limpar Estoque/ID Fis Compart." onClick="LimparCompartilhamento()" accesskey="&" onmouseover="showtip(this,event,'Limpar dados de estoque / Compartilhamento  (Shift+Alt+&)');"></td>
			<td><input type="button" class="button" name="btnSair" value="Sair" style="width:150px"  onClick="sair()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');"></td>
			<!--<td ><input type="button" class="button" name="btnSair" value="Voltar" style="width:150px" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');" onClick="VoltarOrigem()"></td>-->
		</tr>
		
		
		
	</table>
	</td>
	</tr>
</table>
<% end if %>
</Form>
<table width="760">
	<tr>
		<td>
			<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
		</td>
	</tr>
	<tr>
		<td>
			&nbsp;&nbsp;&nbsp;&nbsp;Legenda: A - Alfanumérico;  N - Numérico;  L - Letra
		</td>
	</tr>

</table>
<SCRIPT LANGUAGE=javascript>
<!--
var strDet = new String('<%=strDet%>')
var strNDet = new String('<%=strNDet%>')
var strAde = new String('<%=strAde%>')
var strBsodNet = new String('<%=strBsodNet%>')
var strBsodVia = new String('<%=strBsodVia%>')
var strBsod = new String('<%=strBsod%>')
var strBsodLight = new String('<%=strBsodLight%>')
var strFoEtherNet = new String('<%=strFoEtherNet%>')

var strCamposDet = new String('<%=strCamposDet%>')
var strCamposNDet = new String('<%=strCamposNDet%>')
var strCamposAde = new String('<%=strCamposAde%>')
var strCamposBsodNet = new String('<%=strCamposBsodNet%>')
var strCamposBsodVia = new String('<%=strCamposBsodVia%>')
var strCamposBsod = new String('<%=strCamposBsod%>')
var strCamposBsodLight = new String('<%=strCamposBsodLight%>')
var strCamposFoEtherNet = new String('<%=strCamposFoEtherNet%>')

var strCamposBsodRadioIP = new String('<%=strCamposBsodRadioIP%>')

var strCamposEthernetInterconexao = new String('<%=strCamposEthernetInterconexao%>')
var strCamposEthernetMetro = new String('<%=strCamposEthernetMetro%>')
var strCamposEthernetEDD = new String('<%=strCamposEthernetEDD%>')
var strCamposEthernetPE = new String('<%=strCamposEthernetPE%>')
var strCamposEthernetVlan = new String('<%=strCamposEthernetVlan%>')

function ResgatarInfoRede(){
<%
Select Case Cint("0" & strRede)
	Case 1 'Det
		Response.write "spnDet.innerHTML = '" & strDet &"';"
		Response.write "spnCampos.innerHTML = '" & strCamposDet &"';"
	Case 2 'NDet
		Response.Write "spnNDet.innerHTML = '" & strNDet &"';"
		Response.write "spnCampos.innerHTML = '" & strCamposNDet &"';"
	Case 3 'ADE
		Response.Write "spnAde.innerHTML = '" & strAde &"';"
		Response.write "spnCampos.innerHTML = '" & strCamposAde &"';"
	Case 4 'BSOD Net
		Response.Write "spnBsodNet.innerHTML = '" & strBsodNet &"';"
		Response.write "spnCampos.innerHTML = '" & strCamposBsodNet &"';"
	Case 5 'BSOD Via
		Response.Write "spnBsodVia.innerHTML = '" & strBsodVia &"';"
		Response.write "spnCampos.innerHTML = '" & strCamposBsodVia &"';"
		'Response.write "alert('aaaaaaaaaaaaaaaaaaaaa');"
	Case 13 'HFC BSOD
		Response.Write "spnBsod.innerHTML = '" & strBsod &"';"
		Response.write "spnCampos.innerHTML = '" & strCamposBsod &"';"				
		'Response.write "alert('zzzzzzzzzzzzzzzzzzz');"
	Case 10 'BSOD Light
		Response.Write "spnBsodLight.innerHTML = '" & strBsodLight &"';"
		Response.write "spnCampos.innerHTML = '" & strCamposBsodLight &"';"
	Case 11 'FO EtherNet
		Response.Write "spnFoEtherNet.innerHTML = '" & strFoEtherNet &"';"
		Response.write "spnCampos.innerHTML = '" & strCamposFoEtherNet &"';"
	Case 12 'FO EtherNet
		'Response.Write "spnFoEtherNet.innerHTML = '" & strFoEtherNet &"';"
		Response.write "spnCampos.innerHTML = '" & strCamposBsodRadioIP &"';"
		Response.write "ResgatarSwitchRadioIp()"
	Case 14 'EtherNet
		'Response.Write "spnFoEtherNet.innerHTML = '" & strFoEtherNet &"';"
		Response.write " spnCamposEthernetInterconexao.innerHTML = '" & strCamposEthernetInterconexao &"';"
		
		Response.write "spnCamposEthernetMetro.innerHTML = '" & strCamposEthernetMetro &"'; "
				
		Response.write "spnCamposEthernetEDD.innerHTML = '" & strCamposEthernetEDD &"';"
				
		Response.write "spnCamposEthernetPE.innerHTML = '" & strCamposEthernetPE &"';"
				
		Response.write "spnCamposEthernetVlan.innerHTML = '" & strCamposEthernetVlan &"';"		
		
		Response.write " ResgatarEthernetInterconexao() ; "
		Response.write " ResgatarEthernetMetro(); "
		Response.write " ResgatarEthernetEDD(); "
		Response.write " ResgatarEthernetPE(); "
		
		
		'
		'
		
End Select
%>
}

function ResgatarInfoRedeEstoque(intRede)
//alert('intRede='+intRede)
{
	switch (parseInt(intRede))
	{
		case 1: //Det
			spnDet.innerHTML = strDet
			spnCampos.innerHTML = strCamposDet
			spnNDet.innerHTML = ''
			spnAde.innerHTML = ''
			spnBsodNet.innerHTML = ''
			spnBsodVia.innerHTML = ''
			spnBsod.innerHTML = ''
			spnBsodLight.innerHTML = ''
			spnFoEtherNet.innerHTML = ''
			spnCamposEthernetInterconexao.innerHTML = ''
			spnCamposEthernetMetro.innerHTML = ''
			spnCamposEthernetEDD.innerHTML = ''
			spnCamposEthernetPE.innerHTML = ''
			spnCamposEthernetVlan.innerHTML = ''
			break
		case 2: //NDet
			spnNDet.innerHTML = strNDet
			spnCampos.innerHTML = strCamposNDet
			spnDet.innerHTML = ''
			spnAde.innerHTML = ''
			spnBsodNet.innerHTML = ''
			spnBsodVia.innerHTML = ''
			spnBsod.innerHTML = ''
			spnBsodLight.innerHTML = ''
			spnFoEtherNet.innerHTML = ''
			spnCamposEthernetInterconexao.innerHTML = ''
			spnCamposEthernetMetro.innerHTML = ''
			spnCamposEthernetEDD.innerHTML = ''
			spnCamposEthernetPE.innerHTML = ''
			spnCamposEthernetVlan.innerHTML = ''
			break
		case 3: //ADE
			spnAde.innerHTML = strAde
			spnCampos.innerHTML = strCamposAde
			spnNDet.innerHTML = ''
			spnDet.innerHTML = ''
			spnBsodNet.innerHTML = ''
			spnBsodVia.innerHTML = ''
			spnBsod.innerHTML = ''
			spnBsodLight.innerHTML = ''
			spnFoEtherNet.innerHTML = ''
			spnCamposEthernetInterconexao.innerHTML = ''
			spnCamposEthernetMetro.innerHTML = ''
			spnCamposEthernetEDD.innerHTML = ''
			spnCamposEthernetPE.innerHTML = ''
			spnCamposEthernetVlan.innerHTML = ''
			break
		case 4: //BsodNet
			spnBsodNet.innerHTML = strBsodNet
			spnCampos.innerHTML = strCamposBsodNet
			spnNDet.innerHTML = ''
			spnDet.innerHTML = ''
			spnAde.innerHTML = ''
			spnBsodVia.innerHTML = ''
			spnBsod.innerHTML = ''
			spnBsodLight.innerHTML = ''
			spnFoEtherNet.innerHTML = ''
			spnCamposEthernetInterconexao.innerHTML = ''
			spnCamposEthernetMetro.innerHTML = ''
			spnCamposEthernetEDD.innerHTML = ''
			spnCamposEthernetPE.innerHTML = ''
			spnCamposEthernetVlan.innerHTML = ''
			break
		case 5: //BsodVia
			spnBsodVia.innerHTML = strBsodVia
			spnCampos.innerHTML = strCamposBsodVia
			spnNDet.innerHTML = ''
			spnDet.innerHTML = ''
			spnAde.innerHTML = ''
			spnBsod.innerHTML = ''
			spnBsodNet.innerHTML = ''
			spnBsodLight.innerHTML = ''
			spnFoEtherNet.innerHTML = ''
			spnCamposEthernetInterconexao.innerHTML = ''
			spnCamposEthernetMetro.innerHTML = ''
			spnCamposEthernetEDD.innerHTML = ''
			spnCamposEthernetPE.innerHTML = ''
			spnCamposEthernetVlan.innerHTML = ''
			break
		case 13: //HFC BSOD
			spnBsodVia.innerHTML = ''
			spnCampos.innerHTML = strCamposBsod
			spnNDet.innerHTML = ''
			spnDet.innerHTML = ''
			spnAde.innerHTML = ''
			spnBsod.innerHTML = strBsod
			spnBsodNet.innerHTML = ''
			spnBsodLight.innerHTML = ''
			spnFoEtherNet.innerHTML = ''
			spnCamposEthernetInterconexao.innerHTML = ''
			spnCamposEthernetMetro.innerHTML = ''
			spnCamposEthernetEDD.innerHTML = ''
			spnCamposEthernetPE.innerHTML = ''
			spnCamposEthernetVlan.innerHTML = ''
			break			
		case 10: //BsodLight
			spnBsodLight.innerHTML = strBsodLight
			spnCampos.innerHTML = strCamposBsodLight
			spnBsodVia.innerHTML = ''
			spnNDet.innerHTML = ''
			spnDet.innerHTML = ''
			spnAde.innerHTML = ''
			spnBsod.innerHTML = ''
			spnBsodNet.innerHTML = ''
			spnFoEtherNet.innerHTML = ''
			spnCamposEthernetInterconexao.innerHTML = ''
			spnCamposEthernetMetro.innerHTML = ''
			spnCamposEthernetEDD.innerHTML = ''
			spnCamposEthernetPE.innerHTML = ''
			spnCamposEthernetVlan.innerHTML = ''
			break
		case 11: //FoEtherNet
			spnFoEtherNet.innerHTML = strFoEtherNet
			spnCampos.innerHTML = strCamposFoEtherNet
			spnBsodVia.innerHTML = ''
			spnNDet.innerHTML = ''
			spnDet.innerHTML = ''
			spnAde.innerHTML = ''
			spnBsod.innerHTML = ''
			spnBsodNet.innerHTML = ''
			spnBsodLight.innerHTML = ''
			spnCamposEthernetInterconexao.innerHTML = ''
			spnCamposEthernetMetro.innerHTML = ''
			spnCamposEthernetEDD.innerHTML = ''
			spnCamposEthernetPE.innerHTML = ''
			spnCamposEthernetVlan.innerHTML = ''
			break
		case 12: //FoEtherNet
			spnDet.innerHTML = ''
			spnCampos.innerHTML = strCamposBsodRadioIP
			spnBsodVia.innerHTML = ''
			spnNDet.innerHTML = ''
			spnDet.innerHTML = ''
			spnAde.innerHTML = ''
			spnBsod.innerHTML = ''
			spnBsodNet.innerHTML = ''
			spnBsodLight.innerHTML = ''
			spnCamposEthernetInterconexao.innerHTML = ''
			spnCamposEthernetMetro.innerHTML = ''
			spnCamposEthernetEDD.innerHTML = ''
			spnCamposEthernetPE.innerHTML = ''
			spnCamposEthernetVlan.innerHTML = ''
			break		
		case 14: //EtherNet
			spnDet.innerHTML = ''
			spnCampos.innerHTML = ''
			spnBsodVia.innerHTML = ''
			spnNDet.innerHTML = ''
			spnDet.innerHTML = ''
			spnAde.innerHTML = ''
			spnBsod.innerHTML = ''
			spnBsodNet.innerHTML = ''
			spnBsodLight.innerHTML = ''
			spnCamposEthernetInterconexao.innerHTML = strCamposEthernetInterconexao
			spnCamposEthernetMetro.innerHTML = strCamposEthernetMetro
			spnCamposEthernetEDD.innerHTML = strCamposEthernetEDD
			spnCamposEthernetPE.innerHTML = strCamposEthernetPE
			spnCamposEthernetVlan.innerHTML = strCamposEthernetVlan
			break
			
	}
}

function PopularCamposFac(){
	<%if Trim(strProId) <> "" then %>
		if ('<%=strRede%>'=='3') ResgatarPadraoProvedor(document.forms[0].cboProvedor,1)
		try{
			document.forms[0].hdnProvedor.value = '<%=strProId%>'
			ResgatarProvedoresAssociados(document.forms[0].hdnProvedor)
		}catch(e){}
	<%End if%>
}

function LimparCompartilhamento()
{
	with (document.forms[0])
	{
		var objNodeE = objXmlGeral.selectNodes("//Facilidade[strTipoConsulta='C' || strTipoConsulta='E']")
		if (objNodeE.length > 0)
		{
			hdnIdAcessoFisico.value = "0"
			hdnIdAcessoFisico1.value = "0"
			hdnPropIdFisico.value = "0"
			hdnPropIdFisico1.value = "0"
			hdnCompartilhamento.value = "0"
			hdnCompartilhamento1.value = "0"
			hdnAlteracao.value = "0"
			hdnPodeAlterar.value = "S"
			divIDFis1.style.display = 'none'
			objXmlGeral.loadXML("<xDados/>")
			cboRede.disabled = false;
			AtualizarListaFac()
		}else{
			alert("Item de estoque/compartilhamento não selecionado.")
			return
		}
	}
}

function LiberarFacilidade()
{
	with (document.forms[0])
	{
		var objXmlFacOrigemAux = new ActiveXObject("Microsoft.XMLDOM")
		objXmlFacOrigemAux.loadXML(objXmlFacOrigem.xml)

		var objNode = objXmlFacOrigemAux.selectNodes("//xDados/Facilidade[Fac_ID="+parseInt(arguments[0])+"]")

		if 	(objNode.length > 0)
		{
			var objNodeRequest = objNode[0].getElementsByTagName("Fac_Representacao")
			if (objNodeRequest.length != 0) //Tem representação
			{
				var stRepresentacao = objNodeRequest[0].childNodes[0].text

				for (var intIndex=0;intIndex<objNode.length;intIndex++)
				{
					var objNodeRepresentacao = objNode[intIndex].getElementsByTagName("Fac_Representacao")
					if (objNodeRepresentacao.length > 0)
					{
						if (stRepresentacao == objNodeRepresentacao[0].childNodes[0].text)
						{
							var objNodeListAtual = objNode[intIndex].getElementsByTagName("intFacSel")
							if (objNodeListAtual.length == 0)
							{
								objNodeFilho = objXmlFacOrigemAux.createNode("element", "intFacSel", "")
								objNodeFilho.text = 1
								objNode[intIndex].appendChild (objNodeFilho)
							}
							else
							{
								objNodeListAtual.item(0).text = 1
							}
						}
					}
				}
			}
			else
			{
				for (var intIndex=0;intIndex<objNode.length;intIndex++)
				{
					var objNodeListAtual = objNode[intIndex].getElementsByTagName("intFacSel")
					if (objNodeListAtual.length == 0)
					{
						objNodeFilho = objXmlFacOrigemAux.createNode("element", "intFacSel", "")
						objNodeFilho.text = 1
						objNode[intIndex].appendChild (objNodeFilho)
					}
					else
					{
						objNodeListAtual.item(0).text = 1
					}
				}
			}
		}else
		{
			alert("Facilidade não encontrada.")
			return
		}
		hdnAcao.value = "LiberarFacilidade"
		hdnXmlFacLibera.value = objXmlFacOrigemAux.xml
		target = "IFrmProcesso"
		action = "ProcessoFac.asp"
		submit()
	}
}
<%
if Trim(Request.Form("cboProvedor")) <> "" and  strRede = "" then
		Response.Write	"spnDet.innerHTML = strDet;"
		Response.Write	"spnCampos.innerHTML = strCamposDet;"
		Response.Write	"spnNDet.innerHTML = '';"
		Response.Write	"spnAde.innerHTML = '';"
		Response.Write	"spnBsodNet.innerHTML = '';"
		Response.Write	"spnBsodVia.innerHTML = '';"
		Response.Write	"spnBsod.innerHTML = '';"
		Response.Write	"spnBsodLight.innerHTML = '';"
	    Response.Write	"spnFoEtherNet.innerHTML = '';"
		
		Response.Write	"spnCamposEthernetInterconexao.innerHTML = '';"
		Response.Write	"spnCamposEthernetMetro.innerHTML = '';"
		Response.Write	"spnCamposEthernetEDD.innerHTML = '';"
		Response.Write	"spnCamposEthernetPE.innerHTML = '';"
		Response.Write	"spnCamposEthernetVlan.innerHTML = '';"
End if
%>
//-->
function FileStatus(NameFile)
{
   var blnret ;
   try{
   var fso  = new ActiveXObject("Scripting.FileSystemObject");
	}
	catch(e)
	{
		alert(e.description)
	}

   if (fso.FileExists(NameFile))
      blnret  = true ;
   else
      blnret  = false ;

   return(blnret);
}
</SCRIPT>
