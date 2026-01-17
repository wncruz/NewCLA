
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

<script>
	function Trim(str){return str.replace(/^\s+|\s+$/g,"");}

	function Desalocar ()
	{
		with (document.forms[0])
		{
			
			target = "IFrmProcesso"
			action = "Desalocar.asp"
			submit()
		}
	}
</script>
<% if strTecnologia = "FO EDD" then %>
		<table rules="groups"  border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
		<tr>
			<th colspan=10 class="clsSilver">&nbsp;?&nbsp; Controle do Switch</th>
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
		
		
	
<%else %>
<% '@@Davif  %>
<!--
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
-->
<%'@@Davif  %>
<table rules="groups"  border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
	<tr>
		<th colspan=10>&nbsp;?&nbsp;Recurso</th>
	</tr>

<!--JCARTUS-->
	<tr class="clsSilver">
		<td width="170px" nowrap><font class="clsObrig">:: </font>Local de Entrega</td>
		<td colspan="8">
	    	<%
			set objRS = db.execute("CLA_sp_sel_estacao " & Trim(strLocalInstala))
			%> 
			<input type="Hidden" name="cboLocalInstala" value="<%=strLocalInstala%>">
		  <input type="text" readonly class="text" name="txtCNLLocalEntrega" value=<%=objRS("Cid_Sigla")%> maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		  &nbsp;<input type="text" readonly class="text" name="txtComplLocalEntrega" value=<%=objRS("Esc_Sigla")%> maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsuFac(document.Form2.txtCNLLocalEntrega,document.Form2.txtComplLocalEntrega,<%=dblUsuId%>,1);" TIPO="A">
		</td> 
		<td colspan="1">&nbsp;</td>
	</tr>
	<tr class="clsSilver">
		<td width="170px" nowrap><font class="clsObrig">:: </font>Local de Configura??o</td>
		<td colspan="8">
			<%
			set objRS = db.execute("CLA_sp_sel_estacao " & Trim(strLocalConfig))
			%>
			<input type="Hidden" name="cboLocalConfig" value="<%=strLocalConfig%>">
		  <input type="text" readonly class="text" name="txtCNLLocalConfig" value=<%=objRS("Cid_Sigla")%> maxlength="4" size="6" onKeyUp="ValidarTipo(this,1)"	onblur="CompletarCampo(this)" TIPO="A">&nbsp;
		  &nbsp;<input type="text" readonly class="text" name="txtComplLocalConfig" value=<%=objRS("Esc_Sigla")%> maxlength="3" size="6" onKeyUp="ValidarTipo(this,7)" onblur="CompletarCampo(this);CheckEstacaoUsuFac(document.Form2.txtCNLLocalConfig,document.Form2.txtComplLocalConfig,<%=dblUsuId%>,2);" TIPO="A">
		</td>
		<td colspan="1">&nbsp;</td>
	</tr>
<!--JCARTUS-->

	

	

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
	<%
		set objRS = db.execute("select top 1 Sol_Referencia from cla_solicitacao where sol_id =  " & dblSolId )

		dblSol_Referencia = objRS("Sol_Referencia")
		
			'response.write dblSol_Referencia
			'response.write "-"
			'response.write DblAcf_ID
			'response.end 
			
		
		if dblSol_Referencia = "" or IsNull(dblSol_Referencia) then 
			
			set objRS = db.execute("select top 1 sol_id  from cla_newAcessoFisico  where acf_id =  " & DblAcf_ID )

				dblSol_Referencia = objRS("sol_id")
				
			'response.write dblSol_Referencia
			'response.write "-"
			'response.write DblAcf_ID
			
		end if 
		
			'response.end 


		set objRS = db.execute("CLA_sp_sel_newconsultaTecnologiaFacilidade " & dblSol_Referencia  )
	%>
	
	<tr class=clsSilver>
		<td width=170 ><font class="clsObrig">:: </font>Facilidade</td>
		<td colspan="9" >
			
			<select name="cboNewFacilidade" style="width:250px" readonly>
				<%Response.Write "<Option value='" & Trim(objRS("newfac_id")) & "' >" & objRS("newfac_Nome") & "</Option>"%>
			</select>
		</td>
	</tr>
	<tr class=clsSilver>
		<td width=170 ><font class="clsObrig">:: </font>Tecnologia</td>
		<td colspan="9" >
			
			<select name="cboNewTecnologia" style="width:250px" readonly>
				<%Response.Write "<Option value='" & Trim(objRS("newtec_id")) & "' >" & objRS("newtec_Nome") & "</Option>"%>
			</select>
		</td>
	</tr>
	
	
	
	
</table>
<%
			
			Vetor_Campos(1)="adInteger,2,adParamInput," & dblSol_Referencia 
			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_newfacilidadeServico",1,Vetor_Campos)
			'Response.Write strSqlRet
			Set objRS = db.Execute(strSqlRet)
			objRS.Close
			objRS.CursorLocation = adUseClient
			objRS.Open

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
	


%>					
	<table cellspacing=0 cellpadding=0 width=760 border=0>
		<tr><th colspan=10>&nbsp;?&nbsp;Facilidade do Servi?o</th></tr>
		<tr class=clsSilver>
			<td width="170px" nowrap><font class="clsObrig">:: </font><%=strVlan %> </td>
			<td colspan="9" >
			<input type=text readonly class=text name='ser_Vlan'  size='5' maxlength='5'  onKeyUp="ValidarTipo(this,0)" value="<%=objRS("newfacservico_vlan")%>">
			</td>
		</tr>
		<tr class=clsSilver2> 
			<td width="170px" nowrap><font class="clsObrig">:: </font><%=strPorta %></td>
			<td colspan="9" >
				<input type=text readonly class=text name='ser_portaOLt'  size='16' maxlength='16' value="<%=Trim(objRS("newfacservico_porta"))%>"> 
			</td>
		</tr>
		<tr class=clsSilver>
			<td width="170px" nowrap><font class="clsObrig">:: </font><%=strSVlan %></td>
			<td colspan="9" >
				<input type=text readonly class=text name='ser_SVLAN'  size='5' maxlength='5' onKeyUp="ValidarTipo(this,0)" value="<%=Trim(objRS("newfacservico_svlan"))%>">
			</td>
		</tr>
		<tr class=clsSilver2>
			<td width="170px" nowrap><font class="clsObrig">:: </font><%=strPE %></td>
			<td colspan="9" >
				<input type=text readonly class=text name='ser_PE'  size='40' maxlength='40' value="<%=Trim(objRS("newfacservico_pe"))%>" >
			</td>
		</tr>

	</table>
<%

End if

%>

<table border=0 cellspacing=1 cellpadding=0  width=760>
	<tr><th colspan=10>&nbsp;?&nbsp;Facilidades do Acesso f?sico</th></tr>
	<tr>
		<th>&nbsp;Estrutura</th>
		<th colspan=3>&nbsp;Conte?do</th>
		<th colspan=3>&nbsp;Solicita??o</th>
	</tr>

			<%
			'response.write dblSol_Referencia
			'response.write "-"
			'response.write DblAcf_ID
			'response.end 

			Vetor_Campos(1)="adInteger,2,adParamInput," & dblSol_Referencia 
			Vetor_Campos(2)="adInteger,2,adParamInput," & DblAcf_ID
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
		%>
		<tr class=<%=strClass%>>
			<td width=170 >&nbsp;<%=TratarAspasHtml(objRS("label"))%></td>
			
			<td colspan=3 readonly ><%=trim(objRS("conteudo"))%></td>
			<td colspan=3 readonly ><%=trim(objRS("sol_id"))%></td>
			
		</tr>
		<%
		intCount = intCount+1
		objRS.MoveNext
	Wend
		
End if
%>
		</td>
	</tr>
</table>


<table cellspacing=0 cellpadding=0 width=760 border=0>
	<tr><th colspan=9>&nbsp;?&nbsp;Facilidade do Acesso</th></tr>
		<!--</table>
		<span id=spnCampos></span> -->

		<!--<table cellspacing=1 cellpadding=0 width=760 border=0>
			<tr class=clsSilver>
			<td > -->
			
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass

'strSql = "CLA_sp_sel_AssocTecnologiaFacilidade"

'Call PaginarRS(0,strSql)
			Vetor_Campos(1)="adInteger,2,adParamInput," & dblSol_Referencia 
			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_newAlocacaoAcesso",1,Vetor_Campos)
			'Response.Write strSqlRet
			Set objRS = db.Execute(strSqlRet)
			objRS.Close
			objRS.CursorLocation = adUseClient
			objRS.Open
intCount=1
'response.write "<script>alert('"&objRS.PageSize&"')</script>"
if not objRS.Eof and not objRS.Bof then
	'For intIndex = 1 to objRS.PageSize
	While Not objRS.Eof
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class=<%=strClass%>>
			<td width="170px" nowrap><font class="clsObrig">:: </font><%=TratarAspasHtml(objRS("label"))%></td>
			
			<td colspan="9" >
				<input type=text readonly class=text name='campo_<%=intCount%>' size="<%=trim(objRS("tamanho"))%>" maxlength="<%=trim(objRS("tamanho"))%>" value="<%=trim(objRS("conteudo"))%>">
			</td>
		</tr>
		<%
		intCount = intCount+1
		objRS.MoveNext
	Wend
		'objRS.MoveNext
		'if objRS.EOF then Exit For
	'Next
End if
%>
		<!--</td>
	</tr>-->
</table>

<input type="Hidden" name="hdnAcl_IDAcessoLogico" value="<%=strIdLogico%>">	
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
			
			<td colspan=4 align=center><input type="button" class="button" name="btnOK" style="width:150px;height:22px" value="Concluir Desalocação" onclick="return Desalocar()" accesskey="I" onmouseover="showtip(this,event,'Desaloca??o (Alt+I)');"></td>
			
			<td><input type="button" class="button" name="btnSair" value="Sair" style="width:150px;height:22px"  onClick="sair()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');"></td>
			
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
			<font class="clsObrig">:: </font> Campos de preenchimento obrigatorio.
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
			alert("Item de estoque/compartilhamento n?o selecionado.")
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
			if (objNodeRequest.length != 0) //Tem representa??o
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
			alert("Facilidade n?o encontrada.")
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
