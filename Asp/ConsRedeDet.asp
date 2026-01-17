<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsRedeDet.asp
'	- Descrição			: Controle de rede deterministica

Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXmlReturn") <> "" then
	objXmlDados.loadXml(Request.Form("hdnXmlReturn"))
Else
	objXmlDados.loadXml("<xDados/>")
End if
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function ValidarPar(obj,strCampo)
{
	if (obj.value.length < 4)
	{
		alert("Verifique o padrão para o campo " + strCampo + " (min.:N4, max.:N4-N4)!")
		obj.focus()
		return false
	}

	if (obj.value.length > 4 && obj.value.length != 9 )
	{
		alert("Verifique o padrão para o campo " + strCampo + " (min.:N4, max.:N4-N4)!")
		obj.focus()
		return false
	}

	return true

}

function Procurar()
{
	with (document.forms[0])
	{
		if (cboProvedor.value != "")
		{
			if (!ValidarCampos(cboProvedor,"Provedor")) return
			if (!ValidarCampos(cboLocalInstala,"Estação")) return
			if (!ValidarCampos(cboPlataforma,"Plataforma")) return

			hdnAcao.value = "ResgatarTimeSlots"
			target = "IFrmProcesso"
			//action = "ProcessoConsRedeDet_Eric.asp"
			action = "ProcessoConsRedeDet.asp"
			submit()
		}	
	}
}

function ResgatarDominioNO(obj)
{
	with (document.forms[0])
	{
		if (obj.value != "")
		{
			if (cboProvedor.value == "" || cboLocalInstala.value == "" || cboPlataforma.value == "")
			{
				alert("Provedor/Estação/Plataforma são obrigatórios para resgatar Domínio-NO.")
				cboLocalInstala.value = ""
				return
			}
			hdnPlataforma.value  = cboPlataforma.value
			hdnAcao.value = "ResgatarDominioNO"
			target = "IFrmProcesso"
			action = "ProcessoConsRedeDet.asp"
			submit()
		}	
	}
}

function DetalharItem(dblSolId)
{
	with (document.forms[0])
	{
		hdnSolId.value = dblSolId
		var strNome = "Facilidade" + dblSolId 
		var objJanela = window.open("about:blank",null,"status=no,toolbar=no,enubar=no,location=no,scrollbars = Yes,resizable=Yes")
		objJanela.name = strNome
		target = strNome
		action = "facilidadeDet.asp"
		submit()
		//
		//null, null, "status=no,toolbar=no,menubar=no,location=no,resizable=Yes,scrollbars = Yes"
		//var intRet = window.showModalDialog('facilidadeDet.asp', dblSolId,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
		//DetalharFac()
	}	
}

function DetalharFacilidade(intFacId){
	if (intFacId != ""){
		var objNode = objXmlGeral.selectNodes("//Facilidade[@Fac_Id="+intFacId+"]")
	}else{
		var objNode = objXmlGeral.selectNodes("//Facilidade")
	}	
	if (objNode.length>0){
		var intRet = window.showModalDialog('MessageConsFac.asp', objNode,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	}
}
//-->
</SCRIPT>
<form method="post" name="Form1">
<input type="hidden" name="hdnAcao">
<input type="hidden" name="hdnSolId">
<input type="hidden" name="hdnPedId">
<input type="hidden" name="timeslot">
<input type="hidden" name="bastidor">
<input type="hidden" name="regua">
<input type="hidden" name="posicao">
<input type="hidden" name="fila">
<input type="hidden" name="hdnEstacaoAtual" >
<input type="hidden" name="hdnNomeLocal" >
<input type="hidden" name="hdnRede" >
<input type="hidden" name="estacao">
<input type="hidden" name="distribuidor">
<input type="hidden" name="hdnXmlReturn">
<input type="hidden" name ="hdnPlataforma">
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<tr>
<td >
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2 ><p align="center">Consulta de Rede Determinística</p></td>
</tr>

<tr class=clsSilver>
	<td width=150px ><font class="clsObrig">:: </font>Provedor</td>
	<td >
		<select name="cboProvedor" >
			<option value=""></option>
			<%	
				strProId = Request.Form("cboProvedor")
				if strProId = "" then
					set objNode = objXmlDados.getElementsByTagName("cboProvedor")
					if objNode.length > 0 then
						strProId = objNode(0).childNodes(0).text
					End if
				End if	

				Set objRS = db.execute("CLA_sp_sel_provedor 0")
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strProId) = Trim(objRS("Pro_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value='" & Trim(objRS("Pro_ID")) & "'" & strItemSel & ">" & objRS("Pro_Nome") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td width=150px nowrap><font class="clsObrig">:: </font>Estação</td>
	<td  >
		<select name="cboLocalInstala" >
			<option value=""></option>
			<%	strEstacao = Request.Form("cboLocalInstala")
				if strEstacao = "" then
					set objNode = objXmlDados.getElementsByTagName("cboLocalInstala")
					if objNode.length > 0 then
						strEstacao = objNode(0).childNodes(0).text
					End if
				End if	

				set objRS = db.execute("CLA_sp_sel_usuarioesc " & dblUsuId)
				'set objRS = db.execute("CLA_sp_sel_estacao 0")
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strEstacao) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value=" & objRS("Esc_ID") & strItemSel & ">" & objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td width=150px nowrap><font class="clsObrig">:: </font>Plataforma</td>
	<td  >
		<select name="cboPlataforma" onChange="ResgatarDominioNO(this)">
			<option value=""></option>
			<%	strPlataforma = Request.Form("cboPlataforma")
				if strPlataforma = "" then
					set objNode = objXmlDados.getElementsByTagName("cboPlataforma")
					if objNode.length > 0 then
						strPlataforma = objNode(0).childNodes(0).text
					End if
				End if	

				set objRS = db.execute("CLA_sp_sel_plataforma " )
				'set objRS = db.execute("CLA_sp_sel_estacao 0")
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strPlataforma) = Trim(objRS("Pla_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value=" & objRS("Pla_ID") & strItemSel & ">" & objRS("Pla_Tipoplataforma") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Status da Facilidade</td>
	<td>
		<input type=radio value=0 name=rdoStatusFac checked>Livres&nbsp;
		<input type=radio value=1 name=rdoStatusFac>Ocupadas&nbsp;
		<input type=radio value=2 name=rdoStatusFac>Todos
	</td>
</tr>
<tr class=clsSilver>
	<td width=200px nowrap>&nbsp;&nbsp;&nbsp;&nbsp;Domínio - NO - Slot - Porta</td>
	<td>
		<span id=spnDominioNO>
			<select name="cboDominioNO">
				<option value=""></option>
			<%	Response.Write Request.Form("cboDominioNO")
				if Trim(strProId) <> "" and  Trim(strEstacao) <> "" then
					strDominioNO = Request.Form("cboDominioNO")
					if strDominioNO = "" then
						set objNode = objXmlDados.getElementsByTagName("cboDominioNO")
						if objNode.length > 0 then
							strDominioNO = objNode(0).childNodes(0).text
						End if
					End if	

					set objRS = db.execute("CLA_sp_sel_facilidade_entrada_Agrupado " & strProId & "," & strEstacao)
					While not objRS.Eof 
						strItemSel = ""
						if Trim(strDominioNO) = Trim(objRS("Fac_Dominio")) & "•" & Trim(objRS("Fac_NO")) then strItemSel = " Selected " End if
						Response.Write "<Option value=""" & Trim(objRS("Fac_Dominio")) & "•" & Trim(objRS("Fac_NO")) & """ " & strItemSel & ">" & Trim(objRS("Fac_Dominio")) & " - " & Trim(objRS("Fac_NO")) & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				End if	
			%>
			</select>
		</span>	
	</td>
</tr>
</table>
<table width="760">
	<tr>
		<td colspan=2 align="center">
		<input type="button" class="button" name="btnGravar" value="Procurar" onclick="Procurar()" >&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparForm()">&nbsp;
		<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.location.replace('main.asp')">
		</td>
	</tr>
</table>
<table width="760" cellspacing="0" cellpadding="0">
<tr>
	<td align="center" width=100%>
		<iframe	id			= "IFrmProcesso"
			    name        = "IFrmProcesso" 
			    width       = "100%" 
			    height      = "175px"
			    frameborder = "0"
			    scrolling   = "overflow" 
			    align       = "left">
		</iFrame>
	</td>
</tr>
</table>
<table cellspacing=1 width=760 cellpadding=0 border=0>
	<tr>
		<td colspan=2 class=clsSilver2>&nbsp;•&nbsp;Legenda
		</td>
	</tr>
	<tr class=clsSilver>
		<td width=5px bgcolor=blue>&nbsp;&nbsp;</td>
		<td width=755px nowrap>&nbsp;&nbsp;Status do pedido "Aceito/Instalado" (time-slot reservado - acesso entregue)</td>
	</tr>	
	<tr class=clsSilver>
		<td width=5px bgcolor=red>&nbsp;&nbsp;</td>
		<td width=755px nowrap>&nbsp;&nbsp;Status do pedido "Pendente" (time-slot reservado - acesso não entregue)</td>
	</tr>	
	<tr class=clsSilver>
		<td width=5px bgcolor=#33CC33>&nbsp;&nbsp;</td>
		<td width=755px nowrap>&nbsp;&nbsp;Em estoque</td>
	</tr>	
	<tr class=clsSilver>
		<td width=5px bgcolor=white>&nbsp;&nbsp;</td>
		<td nowrap width=755px>&nbsp;&nbsp;Vago</td>
	</tr>	
	<tr class=clsSilver>
		<td>
			<font class="clsObrig" align=center>&nbsp;::&nbsp;</font>
		</td>
		<td>
			&nbsp;&nbsp;Campos de preenchimento obrigatório.
		</td>
	</tr>
</table>
</td>
</tr>
</table>
<SCRIPT LANGUAGE=javascript>
<!--
function CarregarDoc()
{
	//document.onreadystatechange = CheckStateDoc;
	//document.resolveExternals = false;
}

function CheckStateDoc()
{
  var state = document.readyState;
  
  if (state == "complete")
  {
	Procurar()  
  }
}
CarregarDoc()
//-->
</SCRIPT>
</body>
