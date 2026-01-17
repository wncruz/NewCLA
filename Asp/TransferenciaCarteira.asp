<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: TransferenciaCarteira.ASP
'	- Descrição			: Transferência de Carteira
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%

Dim strDtPedido
strDtPedido = right("0" & day(now),2) & "/" & right("0" & month(now),2) & "/" & year(now)
Function FormatarXmlLog(strXml)

	Dim strXmlDadosAux
	'Retira a quebra de linha que tem no final XML e passa para a variável que vai para o HTML
	strXmlDadosAux = Replace(strXml,Chr(13),"") 
	strXmlDadosAux = Replace(strXmlDadosAux,Chr(10),"")

	FormatarXmlLog = strXmlDadosAux
        
End Function

%>

<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
var objXmlCarteira = new ActiveXObject("Microsoft.XMLDOM")

function ProcurarCarteira(obj)
{
	with (document.forms[0])
	{
		if (obj =='[object]') strValue = obj.value
		else strValue = obj
		if (strValue != "")
		{
			hdnAcao.value = "ConsultarCarteira"
			hdnUsuarioAtual.value = strValue
			target = "IFrmProcesso2"
			action = "ProcessoTraferirCarteira.asp"
			submit()
		}else{
			alert("Selecione o usuário");
			return
		}
		
	}
}
function DetalharItem(dblSolId)
{
	with (document.forms[0])
	{
		PopularXml()
		hdnSolId.value = dblSolId
		DetalharFac()
	}	
}

function CarregarLista()
{
	objXmlGeral.onreadystatechange = CheckStateXml;
	objXmlGeral.resolveExternals = false;
	<%if Request.Form("hdnXmlReturn") = "" then%>
		objXmlGeral.loadXML("<xDados/>")
	<%Else%>
		objXmlGeral.loadXML("<%=FormatarXMLLog(Request.Form("hdnXmlReturn"))%>") 
	<%End if%>	
}
//Verifica se o Xml já esta carregado
function CheckStateXml()
{
  var state = objXmlGeral.readyState;
  
  if (state == 4)
  {
    var err = objXmlGeral.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    } 
    else 
    {
		PopularForm()
		ProcurarCarteira(document.forms[0].hdnUsuarioAtual.value)
	}
  }
}

function CarregarDoc()
{
	document.onreadystatechange = CheckStateDoc;
	document.resolveExternals = false;
}

function CheckStateDoc()
{
  var state = document.readyState;
  
  if (state == "complete")
  {
	CarregarLista()
  }
}

function Transferir(obj)
{
	with (document.forms[0])
	{
		if (!ValidarCampos(cboUsuarioDe,"Usuário \"DE\" ")) return
		if (!ValidarCampos(cboUsuarioPara,"Usuário \"Para\" ")) return
		if (obj.checked)
		{
			hdnAcao.value = "Transferir"
			hdnSolId.value = obj.value
		
			target = "IFrmProcesso"
			action = "ProcessoTraferirCarteira.asp"
			submit()
		}
	}
}

function SelecionarCarteira(obj)
{
	if (objXmlCarteira.xml == ""){objXmlCarteira.loadXML("<xDados></xDados>")}
	var objNode = objXmlCarteira.selectNodes("//xDados[Sol_Id="+obj.value+"]")
	if (objNode.length > 0 && !obj.checked)
	{
		 objNode.item(0).parentNode.removeChild(objNode.item(0))
	}
	else
	{
		if (objNode.length == 0 && obj.checked)
		{
			objNodeFilho = objXmlCarteira.createNode("element", "Sol_Id", "")
			objNodeFilho.text = obj.value
			objXmlCarteira.documentElement.appendChild(objNodeFilho)
		}	
	}
}

<%if Request.ServerVariables("CONTENT_LENGTH") > 0 then%>
	CarregarDoc()
<%End if%>	
//-->
</SCRIPT>
<form method="post" name="Form1">
<input type=hidden name="hdnAcao">
<input type=hidden name="hdnUsuarioAtual">
<input type=hidden name=hdnPedId>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnXmlReturn>
<input type=hidden name=hdnXml>
<tr>
<td >
<table border=0 cellspacing="1" cellpadding = 0 width="760" >
<tr>
	<th colspan=2><p align=center>Transferência de Carteira</p></th>
</tr>
<tr class=clsSilver>
	<td>Do Usuário</td>
	<td><select name="cboUsuarioDe" title="(I) = INATIVO">
			<option value=""></option>
			<%
			Vetor_Campos(1)="adInteger,4,adParamInput," & dblUsuId
			Vetor_Campos(2)="adWChar,3,adParamInput,"
			Vetor_Campos(3)="adInteger,4,adParamOutput,0"  
	
			Call APENDA_PARAM("CLA_sp_sel_UsuarioCtfcAge_ativ_inativ",3,Vetor_Campos)
			ObjCmd.Execute'pega dbaction
			DBAction = ObjCmd.Parameters("RET").value
			Set objRS = ObjCmd.Execute()

			if DBAction = 0 then
				While not objRS.Eof 
					strItemSel = ""
					if Trim(dblUsuId) = Trim(objRS("Usu_ID")) then strItemSel = " Selected " End if
					  var_option = "<Option value=" & objRS("Usu_ID") & strItemSel & ">"
					  if objRS("Usu_Inativo") = "S" then
					    var_option = var_option & "(I) "
					  end if
					  var_option = var_option & objRS("Usu_Nome") & "</Option>"
					Response.Write var_option
					objRS.MoveNext
				Wend
				strItemSel = ""
			End if
			%>
		</select>&nbsp;<input type="button" class="button" name="btnVerCarteira1" value="Ver Carteira" style="width:100px" onclick="ProcurarCarteira(document.forms[0].cboUsuarioDe)" accesskey="1" onmouseover="showtip(this,event,'Ver Carteira(Alt+1)');">&nbsp;
	</td>	
</tr>
<tr class=clsSilver>
	<td>Para o Usuário</td>
	<td><select name="cboUsuarioPara">
			<option value=""></option>
			<%
			Vetor_Campos(1)="adInteger,4,adParamInput," & dblUsuId
			Vetor_Campos(2)="adWChar,3,adParamInput,"
			Vetor_Campos(3)="adInteger,4,adParamOutput,0"  
	
			Call APENDA_PARAM("CLA_sp_sel_usuarioctfcAge",3,Vetor_Campos)
			ObjCmd.Execute'pega dbaction
			DBAction = ObjCmd.Parameters("RET").value
			Set objRS = ObjCmd.Execute()

			if DBAction = 0 then
				While not objRS.Eof 
					strItemSel = ""
					if Trim(dblUsuIdMonint) = Trim(objRS("Usu_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value=" & objRS("Usu_ID") & strItemSel & ">" & objRS("Usu_Nome") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			End if
			%>
		</select>
	</td>	
</tr>
<tr class=clsSilver>
	<td width=200px >Pedido de Acesso</td>
	<td>
		<input type="text" class="text" name="txtPedNum" value="<%if request("txtPedNum") <> "" then response.write ucase(request("txtPedNum")) else response.write "DM-" end if%>" maxlength="25" size="20">
	</td>
</tr>
<tr class=clsSilver style="position:absolute;visibility:hidden">
	<td>CNL</td>
	<td>
		<select name="cboCNL">
			<option value="">-- TODAS --</option>
			<%
				dblCNL = Request.Form("cboCNL")
				if dblCNL = "" then
					set objNode = objXmlDados.getElementsByTagName("cboCNL")
					if objNode.length > 0 then
						dblCNL = objNode(0).childNodes(0).text
					End if
				End if	

				Set objRS = db.execute("CLA_sp_sel_usuarioesc '" & dblUsuId & "',5")
				While not objRS.Eof 
					strSel = ""
					if Trim(dblCNL) = Trim(objRS("Cid_Sigla")) then strSel = " Selected " End if
					Response.Write "<Option value='" & objRS("Cid_Sigla") & "'" & strSel & ">" & objRS("Cid_Sigla") & "</Option>"
					objRS.MoveNext
				Wend
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td width=25% >Nº Solicitação</td>
	<td><input type="text" name="txtSolId" size=10 class=text value="<%=request("txtSolId")%>" onKeyUp="ValidarTipo(this,0)" maxlength=9></td>
</tr>
<tr class=clsSilver>
	<td width=25% >Numero do Acesso</td>
	<td><input type="text" name="txtFacID" size=10  value="<%=request("txtFacID")%>" class=text onKeyUp="ValidarTipo(this,0)" maxlength=10></td>
</tr>
<tr class=clsSilver>
	<td >Cliente</td>
	<td ><input type="text" class="text" name="txtCliente" value="<%=request("txtCliente")%>"  maxlength="60" size="50"></td>
</tr>
<tr class=clsSilver>
	<td nowrap>Endereço</td>
	<td nowrap>
		<input type="text" class="text" name="txtEndereco" value="<%=request("txtEndereco")%>" maxlength="60" size="50">&nbsp;Nº&nbsp;
		<input type="text" class="text" name="txtNroEnd" value="<%=request("txtNroEnd")%>" maxlength="10" size="10">&nbsp;
		Compl&nbsp;<input type="text" class="text" name="txtComplemento" value="<%=request("txtComplemento")%>" maxlength="30" size="20">
	</td>
</tr>
<tr>
	<td colspan=2 align="center" height=30px >
		<!--<input type="button" class="button" name="btnTransferir" value="Transferir Carteira" style="width:100px" onclick="Transferir()" accesskey="I" onmouseover="showtip(this,event,'Enviar (Alt+I)');">&nbsp;-->
		<input type="button" class="button" name="btnVerCarteira1" value="Ver Carteira" style="width:100px" onclick="ProcurarCarteira(document.forms[0].cboUsuarioDe)" accesskey="1" onmouseover="showtip(this,event,'Ver Carteira(Alt+1)');">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparForm();" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>
<span id=spnLinks></span>
<table border=0 width=758 cellspacing=1 cellpadding=1 >
<table border=0 cellspacing=1 width=760>
<tr><b>:: Para transferir basta selecionar o item</b></tr>

	<tr>
		<th width=22>&nbsp;</th>
		<th width=100>&nbsp;Pedido</th>
		<th width=40>&nbsp;Sol</th>
		<th width=240>&nbsp;Cliente</th>
		<th width=100>&nbsp;Ação</th>
		<th width=100 >&nbsp;Nº do Contrato</th>
		<th width=140>&nbsp;Status Atual</th>
		<th width=20>&nbsp;Perfil</th>
	</tr>
</table>
<table border=0 width=774 cellspacing=0 cellpadding=0 >
<tr>
	<td width=774>
		<iframe	id			= "IFrmProcesso2"
			    name        = "IFrmProcesso2" 
			    width       = "776"
			    height      = "220"
			    frameborder = "0"
			    scrolling   = "overflow" 
			    align       = "left">
		</iFrame>
	</td>
</tr>	
</table>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0"
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</td>
</tr>
</table>
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>?acao=<%=Trim(Request("acao"))%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnNomeCons value="Carteira">
</form>
</body>
</html>
<%DesconectarCla()%>