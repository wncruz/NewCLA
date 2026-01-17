<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Cons_HistoricoFacilidade.asp
'	- Responsável		: Vital
'	- Descrição			: Consulta de Histórico de Facilidades
Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXmlReturn") <> "" then
	objXmlDados.loadXml(Request.Form("hdnXmlReturn"))
Else
	objXmlDados.loadXml("<xDados/>")
End if

set objNode = objXmlDados.getElementsByTagName("cboSistema")
if objNode.length > 0 then
	strRede = objNode(0).childNodes(0).text
End if
set objNode = objXmlDados.getElementsByTagName("cboProvedor")
if objNode.length > 0 then
	strProId = objNode(0).childNodes(0).text
End if
set objNode = objXmlDados.getElementsByTagName("cboLocalInstala")
if objNode.length > 0 then
	strLocalInstala = objNode(0).childNodes(0).text
End if
set objNode = objXmlDados.getElementsByTagName("cboDistLocalInstala")
if objNode.length > 0 then
	strDistrib = objNode(0).childNodes(0).text
End if
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function ResgatarDistribuicao(obj,intDistrib)
{
	with (document.forms[0])
	{
		if (obj.value != "")
		{
			hdnAcao.value = "ResgatarDistribuicao"
			hdnEstacaoAtual.value = obj.value

			switch (intDistrib)
			{
				case 1:
					hdnNomeLocal.value = "DistLocalInstala"
					break
				case 2:
					hdnNomeLocal.value = "DistLocalConfig"
					break
				case 3:
					hdnNomeLocal.value = "DistLocalInstalaInter"
					break
			}
			target = "IFrmProcesso"
			action = "ProcessoCla.asp"
			submit()
		}	
	}
}

function ResgatarCamposFac(obj)
{
	with (document.forms[0])
	{
		spnPosicoes.innerHTML = "" 
		hdnAcao.value = "ResgatarCamposFacilidade"
		target = "IFrmProcesso2"
		action = "ProcessoConsFac.asp"
		submit()
	}
}

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
		if (!ValidarDM(txtPedido)) return
		hdnAcao.value = "ConsultarHistoricoFac"
		target = "IFrmProcesso"
		action = "ProcessoHistoricoFac.asp"
		submit()
		return
		
		if (!ValidarCampos(cboProvedor,"Provedor")) return
		if (!ValidarCampos(cboSistema,"Rede")) return

		if (cboSistema.value == 2 || cboSistema.value == 3)
		{
			try //pode não existir ainda
			{
				if (rdoInter[0].checked)
				{
					if (!ValidarCampos(cboLocalInstala,"Local de Instalação")) return
					if (!ValidarCampos(cboDistLocalInstala,"Distribuidor do Local de Instalação")) return
				}
			}
			catch(e){}	
		}
		else
		{
			if (!ValidarCampos(cboLocalInstala,"Local de Instalação")) return
			if (!ValidarCampos(cboDistLocalInstala,"Distribuidor do Local de Instalação")) return
		}

		if (!ValidarRede())
		{ 
			return 
		}
		else
		{
			hdnRede.value = cboSistema.value
			target = "IFrmProcesso"
			action = "ProcessoPosicoes.asp"
			//submit()
		}
	}
}
function DetalharSolicitacao(dblSolId)
{
	with (document.forms[0])
	{
		//Enviar para pagina de Detalhamento do pedido
		PopularXml(objXmlReturn)
		hdnSolId.value = dblSolId
		DetalharFac()
	}	
}

function CarregarDocLog()
{
	document.onreadystatechange = CheckStateDocLog;
	document.resolveExternals = false;
}

function CheckStateDocLog()
{
  var state = document.readyState;
  
  if (state == "complete")
  {
	CarregarLista()
  }
}

var objXmlReturn = new ActiveXObject("Microsoft.XMLDOM")
function CarregarLista()
{
	objXmlReturn.onreadystatechange = CheckStateXml;
	objXmlReturn.resolveExternals = false;
	<%if Request.Form("hdnXmlReturn") = "" then%>
		objXmlReturn.loadXML("<xDados/>")
	<%Else%>
		objXmlReturn.loadXML("<%=FormatarStrXML(Request.Form("hdnXmlReturn"))%>") 
	<%End if%>	
}
//Verifica se o Xml já esta carregado
function CheckStateXml()
{
  var state = objXmlReturn.readyState;
  
  if (state == 4)
  {
    var err = objXmlReturn.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    } 
    else 
    {
		PopularForm(objXmlReturn)
		<%if Request.ServerVariables("CONTENT_LENGTH") > 0 then%>
			ResgatarCamposFac(document.forms[0].cboSistema)
			Procurar()
		<%End if%>	
	}
  }
}

function DetalharFacilidade(intFacId){
	var objNode = objXmlGeral.selectNodes("//Facilidade[@Fac_Id="+intFacId+"]")
	if (objNode.length>0){
		var intRet = window.showModalDialog('MessageConsFac.asp',objNode,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	}
}

CarregarDocLog()

//-->
</SCRIPT>

<form method="post" name="Form1">
<input type="hidden" name="hdnAcao">
<input type="hidden" name="hdnEstacaoAtual" >
<input type="hidden" name="hdnNomeLocal" >
<input type="hidden" name="hdnRede" >
<input type="hidden" name="txtNroReg" value=50 >
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnXmlReturn>
<input type=hidden name=hdnNomeCons value="Consulta-Historico-Fac">
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">

<tr>
<td >
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2 ><p align="center">Hitórico de Facilidade</p></td>
</tr>
<tr class=clsSilver>
<td>
	Pedido de Acesso
</td>
<td>
	<input type="text" class="text" name="txtPedido" value="DM-" maxlength="13" size="15"></td>
</tr>
<tr class=clsSilver>
	<td width=150px>Nº Solicitação</td>
	<td><input type="text" name="txtSolId" size=10 class=text onKeyUp="ValidarTipo(this,0)" maxlength=10></td>
</tr>
<tr class=clsSilver>
	<td width=150px>Nº do Acesso</td>
	<td>
		<input type=text class=text name=txtNroAcesso size=20>
	</td>
</tr>
<tr>
	<th colspan=2 >&nbsp;•&nbsp;Recurso</td>
</tr>
<tr class=clsSilver>
	<td width=150px >Provedor</td>
	<td >
		<select name="cboProvedor" >
			<option value=""></option>
			<%	set objRS = db.execute("CLA_sp_sel_provedor 0")
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
	<td>Rede</td>
	<td >
		<select name="cboSistema" onChange="ResgatarCamposFac(this)" >
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
	</td>
</tr>
</table>
<div id=divInstala style="DISPLAY: ''">
<table rules="groups" border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
<tr class=clsSilver>
	<td width=150px nowrap>Estação</td>
	<td  >
		<select name="cboLocalInstala" onChange="ResgatarDistribuicao(this,1)">
			<option value=""></option>
			<%set objRS = db.execute("CLA_sp_sel_usuarioesc " & dblUsuId)
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strLocalInstala) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value=" & objRS("Esc_ID") & strItemSel & ">" & objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
		</select>
	</td>
</tr>

<tr class=clsSilver>
	<td>Distribuidor</td>
	<td >
		<span id="spnDistLocalInstala">
		<select name="cboDistLocalInstala" >
			<option value=""></option>
			<%	if Trim(strLocalInstala) <> "" then 
					set objRS = db.execute("CLA_sp_view_recursodistribuicao " & strLocalInstala)
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
</table>
</div>

<table width="760">
<tr>
	<td align="center" >
		<span id=spnCamposFac></span>
	</td>
</tr>
<table width="760">
	<tr>
		<td colspan=2 align="center">
		<input type="button" class="button" name="btnGravar" value="Procurar" onclick="Procurar()" >&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparForm()">&nbsp;
		<input type="button" class="button" name="btnSair" value="Sair" onClick="javascript:window.location.replace('main.asp')">
		</td>
	</tr>
</table>
<tr>
	<td align="center" >
		<span id=spnPosicoes></span>
	</td>
</tr>
<table width="760">
<tr>
	<td>
	<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
	</td>
</tr>
<tr>
	<td>
		<font class="clsObrig">:: </font>Legenda: A - Alfanumérico;  N - Numérico;  L - Letra
	</td>
</tr>

</table>
</td>
</tr>
</table>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "100%"
	    height      = "100%"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
<iframe	id			= "IFrmProcesso2"
	    name        = "IFrmProcesso2" 
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</body>
</html>
